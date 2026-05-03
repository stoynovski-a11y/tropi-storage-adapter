"""Dropbox backend for the StorageAdapter."""
from __future__ import annotations

import os
import uuid
from typing import Any

import dropbox
from dropbox.exceptions import ApiError, AuthError as DbxAuthError, RateLimitError
from dropbox.files import (
    CommitInfo,
    DeletedMetadata,
    FileMetadata,
    FolderMetadata,
    UploadSessionCursor,
    WriteMode,
)

from ..adapter import StorageAdapter
from ..exceptions import (
    AuthError,
    BackendError,
    ConflictError,
    LockError,
    NotFoundError,
    ThrottledError,
)
from ..logging_config import log_operation
from ..path_utils import normalize_path
from ..retry import retry_on_transient

# Dropbox /upload accepts up to 150 MB per call. Above that, use upload sessions.
_UPLOAD_SINGLE_LIMIT = 140 * 1024 * 1024
_UPLOAD_CHUNK_SIZE = 8 * 1024 * 1024


def _to_storage_error(exc: Exception) -> Exception:
    """Translate Dropbox SDK exceptions to our exception hierarchy."""
    if isinstance(exc, RateLimitError):
        backoff = getattr(exc, "backoff", None)
        return ThrottledError(str(exc), retry_after=float(backoff) if backoff else None)
    if isinstance(exc, DbxAuthError):
        return AuthError(str(exc))
    if isinstance(exc, ApiError):
        # Dropbox API errors carry a structured `error` field with `is_*()` checks.
        err = exc.error
        # path lookups return a LookupError with .is_not_found()
        for attr in ("get_path", "get_path_lookup", "get_from_lookup"):
            if hasattr(err, attr):
                try:
                    sub = getattr(err, attr)()
                    if hasattr(sub, "is_not_found") and sub.is_not_found():
                        return NotFoundError(str(exc))
                except Exception:
                    pass
        # Some errors are themselves not-found (e.g. delete on missing)
        if hasattr(err, "is_not_found") and err.is_not_found():
            return NotFoundError(str(exc))
        return BackendError(str(exc))
    return exc


class DropboxBackend(StorageAdapter):
    """Backed by the official Dropbox Python SDK using a refresh token."""

    backend_name = "dropbox"

    def __init__(
        self,
        *,
        app_key: str | None = None,
        app_secret: str | None = None,
        refresh_token: str | None = None,
        client: dropbox.Dropbox | None = None,
    ):
        self._instance_id = str(uuid.uuid4())
        self._held_locks: set[str] = set()

        if client is not None:
            self._dbx = client
            return

        app_key = app_key or os.getenv("DROPBOX_APP_KEY")
        app_secret = app_secret or os.getenv("DROPBOX_APP_SECRET")
        refresh_token = refresh_token or os.getenv("DROPBOX_REFRESH_TOKEN")

        if not (app_key and app_secret and refresh_token):
            raise AuthError(
                "Dropbox backend requires DROPBOX_APP_KEY, DROPBOX_APP_SECRET, "
                "and DROPBOX_REFRESH_TOKEN env vars."
            )

        self._dbx = dropbox.Dropbox(
            oauth2_refresh_token=refresh_token,
            app_key=app_key,
            app_secret=app_secret,
        )

    # --- helpers -----------------------------------------------------------
    @staticmethod
    def _entry_to_dict(entry: Any) -> dict:
        if isinstance(entry, FileMetadata):
            return {
                "name": entry.name,
                "path": entry.path_display or entry.path_lower,
                "type": "file",
                "size": entry.size,
                "modified": entry.server_modified.isoformat() if entry.server_modified else None,
                "content_hash": entry.content_hash,
                "id": entry.id,
                "etag": entry.content_hash,  # Dropbox uses content_hash as our etag
            }
        if isinstance(entry, FolderMetadata):
            return {
                "name": entry.name,
                "path": entry.path_display or entry.path_lower,
                "type": "folder",
                "size": 0,
                "modified": None,
                "content_hash": None,
                "id": entry.id,
                "etag": None,
            }
        if isinstance(entry, DeletedMetadata):
            return {
                "name": entry.name,
                "path": entry.path_display or entry.path_lower,
                "type": "deleted",
                "size": 0,
                "modified": None,
                "content_hash": None,
                "id": None,
                "etag": None,
            }
        return {"name": getattr(entry, "name", ""), "path": "", "type": "unknown"}

    def _lock_path(self, path: str) -> str:
        # Sibling .lock file: /a/b/file.xlsx -> /a/b/file.xlsx.lock
        return normalize_path(path) + ".lock"

    # --- core ops ----------------------------------------------------------
    @retry_on_transient()
    def read(self, path: str) -> bytes:
        p = normalize_path(path)
        with log_operation(self.backend_name, "read", p):
            try:
                _, response = self._dbx.files_download(p)
                return response.content
            except Exception as e:
                raise _to_storage_error(e) from e

    @retry_on_transient()
    def write(self, path: str, data: bytes, overwrite: bool = True) -> dict:
        p = normalize_path(path)
        with log_operation(self.backend_name, "write", p):
            mode = WriteMode("overwrite") if overwrite else WriteMode("add")
            try:
                if len(data) <= _UPLOAD_SINGLE_LIMIT:
                    meta = self._dbx.files_upload(data, p, mode=mode, mute=True)
                else:
                    meta = self._chunked_upload(p, data, mode)
                return self._entry_to_dict(meta)
            except Exception as e:
                raise _to_storage_error(e) from e

    def _chunked_upload(self, path: str, data: bytes, mode: WriteMode) -> FileMetadata:
        # Start session with first chunk.
        first = data[:_UPLOAD_CHUNK_SIZE]
        session = self._dbx.files_upload_session_start(first)
        cursor = UploadSessionCursor(session_id=session.session_id, offset=len(first))
        offset = len(first)

        while offset < len(data) - _UPLOAD_CHUNK_SIZE:
            chunk = data[offset : offset + _UPLOAD_CHUNK_SIZE]
            self._dbx.files_upload_session_append_v2(chunk, cursor)
            offset += len(chunk)
            cursor = UploadSessionCursor(session_id=session.session_id, offset=offset)

        # Finish with the remainder.
        commit = CommitInfo(path=path, mode=mode, mute=True)
        return self._dbx.files_upload_session_finish(data[offset:], cursor, commit)

    @retry_on_transient()
    def list(self, path: str, recursive: bool = False) -> list[dict]:
        p = normalize_path(path)
        # Dropbox uses '' for root, not '/'.
        api_path = "" if p == "/" else p
        with log_operation(self.backend_name, "list", p):
            try:
                results: list[dict] = []
                resp = self._dbx.files_list_folder(api_path, recursive=recursive)
                results.extend(self._entry_to_dict(e) for e in resp.entries)
                while resp.has_more:
                    resp = self._dbx.files_list_folder_continue(resp.cursor)
                    results.extend(self._entry_to_dict(e) for e in resp.entries)
                return results
            except Exception as e:
                raise _to_storage_error(e) from e

    @retry_on_transient()
    def delete(self, path: str) -> None:
        p = normalize_path(path)
        with log_operation(self.backend_name, "delete", p):
            try:
                self._dbx.files_delete_v2(p)
            except Exception as e:
                err = _to_storage_error(e)
                if isinstance(err, NotFoundError):
                    return  # idempotent
                raise err from e

    @retry_on_transient()
    def move(self, src: str, dst: str) -> None:
        s = normalize_path(src)
        d = normalize_path(dst)
        with log_operation(self.backend_name, "move", f"{s} -> {d}"):
            try:
                self._dbx.files_move_v2(s, d, autorename=False, allow_ownership_transfer=False)
            except Exception as e:
                raise _to_storage_error(e) from e

    @retry_on_transient()
    def copy(self, src: str, dst: str) -> None:
        s = normalize_path(src)
        d = normalize_path(dst)
        with log_operation(self.backend_name, "copy", f"{s} -> {d}"):
            try:
                self._dbx.files_copy_v2(s, d, autorename=False, allow_ownership_transfer=False)
            except Exception as e:
                raise _to_storage_error(e) from e

    @retry_on_transient()
    def ensure_folder(self, path: str) -> None:
        p = normalize_path(path)
        if p == "/":
            return
        with log_operation(self.backend_name, "ensure_folder", p):
            try:
                self._dbx.files_create_folder_v2(p, autorename=False)
            except ApiError as e:
                # Already exists is fine.
                if "conflict" in str(e).lower() or "exists" in str(e).lower():
                    return
                raise _to_storage_error(e) from e
            except Exception as e:
                raise _to_storage_error(e) from e

    @retry_on_transient()
    def get_metadata(self, path: str) -> dict:
        p = normalize_path(path)
        with log_operation(self.backend_name, "get_metadata", p):
            try:
                meta = self._dbx.files_get_metadata(p)
                d = self._entry_to_dict(meta)
                d["exists"] = True
                return d
            except Exception as e:
                err = _to_storage_error(e)
                if isinstance(err, NotFoundError):
                    return {"exists": False, "path": p, "name": p.rsplit("/", 1)[-1]}
                raise err from e

    # --- locking (simulated via .lock sibling file) ------------------------
    def checkout(self, path: str) -> None:
        p = normalize_path(path)
        lock = self._lock_path(p)
        with log_operation(self.backend_name, "checkout", p):
            existing = self.get_metadata(lock)
            if existing.get("exists"):
                # Read it and check if WE hold it (re-entrant).
                try:
                    holder = self.read(lock).decode("utf-8").strip()
                except Exception:
                    holder = ""
                if holder != self._instance_id:
                    raise LockError(f"{p} is locked by another holder")
                return
            self.write(lock, self._instance_id.encode("utf-8"), overwrite=False)
            self._held_locks.add(p)

    def checkin(self, path: str) -> None:
        p = normalize_path(path)
        lock = self._lock_path(p)
        with log_operation(self.backend_name, "checkin", p):
            existing = self.get_metadata(lock)
            if not existing.get("exists"):
                self._held_locks.discard(p)
                return
            try:
                holder = self.read(lock).decode("utf-8").strip()
            except Exception:
                holder = ""
            if holder and holder != self._instance_id:
                raise LockError(f"{p} held by another holder; refusing to release")
            self.delete(lock)
            self._held_locks.discard(p)

    # --- conditional write -------------------------------------------------
    def write_with_etag(self, path: str, data: bytes, etag: str) -> dict:
        p = normalize_path(path)
        with log_operation(self.backend_name, "write_with_etag", p):
            current = self.get_metadata(p)
            # If file doesn't exist and etag is empty, treat as create.
            if current.get("exists"):
                if current.get("etag") != etag:
                    raise ConflictError(
                        f"{p} changed since etag was read "
                        f"(have {current.get('etag')!r}, expected {etag!r})"
                    )
            elif etag:
                raise ConflictError(f"{p} does not exist but caller passed etag {etag!r}")
            return self.write(p, data, overwrite=True)

    # --- health ------------------------------------------------------------
    def healthcheck(self) -> dict[str, Any]:
        import time
        start = time.perf_counter()
        authenticated = False
        can_list_root = False
        try:
            self._dbx.users_get_current_account()
            authenticated = True
        except Exception:
            pass
        try:
            self._dbx.files_list_folder("", limit=1)
            can_list_root = True
        except Exception:
            pass
        return {
            "backend": self.backend_name,
            "authenticated": authenticated,
            "can_list_root": can_list_root,
            "latency_ms": round((time.perf_counter() - start) * 1000, 2),
        }
