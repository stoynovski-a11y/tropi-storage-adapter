"""StorageAdapter abstract interface and factory."""
from __future__ import annotations

import os
from abc import ABC, abstractmethod
from typing import Any

from .logging_config import configure_logging, init_sentry_if_configured

# Item dict shape returned by `list()` and `get_metadata()`:
#   {
#     "name": str,
#     "path": str,
#     "type": "file" | "folder",
#     "size": int,
#     "modified": str (ISO 8601),
#     "content_hash": str | None,   # backend-specific (Dropbox sha256 / Graph cTag-derived)
#     "id": str,                    # backend-specific item id
#     "etag": str,                  # used for conditional writes
#     "exists": bool,               # only on get_metadata
#   }


class StorageAdapter(ABC):
    """Backend-agnostic file-storage interface.

    All paths are POSIX-style, absolute, with a leading slash, e.g. `/Co/foo.xlsx`.
    Backends translate to their native path format internally.
    """

    backend_name: str = "abstract"

    # --- core ops ----------------------------------------------------------
    @abstractmethod
    def read(self, path: str) -> bytes:
        """Download file as bytes. Raises NotFoundError if missing."""

    @abstractmethod
    def write(self, path: str, data: bytes, overwrite: bool = True) -> dict:
        """Upload file. Returns metadata dict for the written file."""

    @abstractmethod
    def list(self, path: str, recursive: bool = False) -> list[dict]:
        """List folder contents. Returns a list of item dicts."""

    @abstractmethod
    def delete(self, path: str) -> None:
        """Delete file or folder. No-op if missing (idempotent)."""

    @abstractmethod
    def move(self, src: str, dst: str) -> None:
        """Move/rename file or folder."""

    @abstractmethod
    def copy(self, src: str, dst: str) -> None:
        """Copy file."""

    @abstractmethod
    def ensure_folder(self, path: str) -> None:
        """Create folder if missing, including parents. Idempotent."""

    @abstractmethod
    def get_metadata(self, path: str) -> dict:
        """Return metadata dict. `exists` key is False if missing (no exception)."""

    # --- locking -----------------------------------------------------------
    @abstractmethod
    def checkout(self, path: str) -> None:
        """Take exclusive lock on file. Implementation differs per backend."""

    @abstractmethod
    def checkin(self, path: str) -> None:
        """Release exclusive lock."""

    # --- conditional write -------------------------------------------------
    @abstractmethod
    def write_with_etag(self, path: str, data: bytes, etag: str) -> dict:
        """Conditional write — raises ConflictError if file changed since etag."""

    # --- health ------------------------------------------------------------
    def healthcheck(self) -> dict[str, Any]:
        """Return connectivity status. Override in backends for richer checks."""
        import time
        start = time.perf_counter()
        try:
            self.list("/")
            ok = True
        except Exception:
            ok = False
        latency_ms = round((time.perf_counter() - start) * 1000, 2)
        return {
            "backend": self.backend_name,
            "authenticated": ok,
            "can_list_root": ok,
            "latency_ms": latency_ms,
        }


def get_adapter() -> StorageAdapter:
    """Return the concrete adapter selected by the STORAGE_BACKEND env var.

    Initializes JSON logging and (optionally) Sentry as a side effect.
    """
    configure_logging()
    init_sentry_if_configured()

    backend = os.getenv("STORAGE_BACKEND", "dropbox").strip().lower()
    if backend == "dropbox":
        # Local imports to avoid pulling SDK dependencies when unused.
        from .backends.dropbox_backend import DropboxBackend
        return DropboxBackend()
    if backend == "m365":
        from .backends.graph_backend import GraphBackend
        return GraphBackend()
    raise ValueError(
        f"Unknown STORAGE_BACKEND: {backend!r} (expected 'dropbox' or 'm365')"
    )
