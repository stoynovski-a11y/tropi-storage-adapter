"""Microsoft Graph (SharePoint / OneDrive) backend.

Uses `httpx` against the Graph REST API and `msal` for client-credentials OAuth.
This is a deliberate divergence from the spec's `msgraph-sdk` recommendation —
the official SDK is async-first (Kiota-generated) and would force every adapter
call through `asyncio.run()`, which is awkward for the existing sync services.
The functionality is identical; only the transport differs.
"""
from __future__ import annotations

import os
import time
import urllib.parse
from typing import Any

import httpx
import msal

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
from ..path_utils import normalize_path, split_parent
from ..retry import retry_on_transient

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPE = ["https://graph.microsoft.com/.default"]

# Files larger than this go through createUploadSession (chunked).
_UPLOAD_SINGLE_LIMIT = 4 * 1024 * 1024
_UPLOAD_CHUNK_SIZE = 5 * 1024 * 1024  # must be a multiple of 320 KiB; 5 MiB is safe


def _http_to_storage_error(resp: httpx.Response) -> Exception:
    """Translate an HTTP error response into our exception hierarchy."""
    status = resp.status_code
    try:
        body = resp.json()
    except Exception:
        body = {"error": {"message": resp.text[:500]}}
    msg = body.get("error", {}).get("message", "") or resp.text[:200]

    if status == 401:
        return AuthError(f"Graph 401: {msg}")
    if status == 403:
        return AuthError(f"Graph 403: {msg}")
    if status == 404:
        return NotFoundError(f"Graph 404: {msg}")
    if status == 409:
        return ConflictError(f"Graph 409: {msg}")
    if status == 412:
        return ConflictError(f"Graph 412 precondition failed: {msg}")
    if status == 423:
        return LockError(f"Graph 423 locked: {msg}")
    if status == 429:
        retry_after = resp.headers.get("Retry-After")
        return ThrottledError(
            f"Graph 429 throttled: {msg}",
            retry_after=float(retry_after) if retry_after else None,
        )
    if status in (500, 502, 503, 504):
        retry_after = resp.headers.get("Retry-After")
        return ThrottledError(
            f"Graph {status} transient: {msg}",
            retry_after=float(retry_after) if retry_after else None,
        )
    return BackendError(f"Graph {status}: {msg}", status_code=status)


# Custom transient set combining ThrottledError + connection-level errors.
_TRANSIENT = (
    ThrottledError,
    httpx.ConnectError,
    httpx.ReadTimeout,
    httpx.WriteTimeout,
    httpx.PoolTimeout,
    httpx.RemoteProtocolError,
)


class GraphBackend(StorageAdapter):
    """SharePoint document library accessed via Microsoft Graph."""

    backend_name = "m365"

    def __init__(
        self,
        *,
        tenant_id: str | None = None,
        client_id: str | None = None,
        client_secret: str | None = None,
        site_hostname: str | None = None,
        site_path: str | None = None,
        http_client: httpx.Client | None = None,
    ):
        self._tenant_id = tenant_id or os.getenv("M365_TENANT_ID", "")
        self._client_id = client_id or os.getenv("M365_CLIENT_ID", "")
        self._client_secret = client_secret or os.getenv("M365_CLIENT_SECRET", "")
        self._site_hostname = site_hostname or os.getenv("M365_SITE_HOSTNAME", "")
        self._site_path = site_path or os.getenv("M365_SITE_PATH", "")

        if not all([self._tenant_id, self._client_id, self._client_secret,
                    self._site_hostname, self._site_path]):
            raise AuthError(
                "Graph backend requires M365_TENANT_ID, M365_CLIENT_ID, "
                "M365_CLIENT_SECRET, M365_SITE_HOSTNAME, M365_SITE_PATH env vars."
            )

        self._http: httpx.Client = http_client or httpx.Client(timeout=60.0)
        self._msal_app: msal.ConfidentialClientApplication | None = None
        self._token: str | None = None
        self._token_expires_at: float = 0.0

        # Cached on first use.
        self._site_id: str | None = None
        self._drive_id: str | None = None
        self._held_locks: set[str] = set()

    # --- auth --------------------------------------------------------------
    def _get_msal(self) -> msal.ConfidentialClientApplication:
        if self._msal_app is None:
            self._msal_app = msal.ConfidentialClientApplication(
                client_id=self._client_id,
                authority=f"https://login.microsoftonline.com/{self._tenant_id}",
                client_credential=self._client_secret,
            )
        return self._msal_app

    def _get_token(self) -> str:
        # Refresh ~60s before expiry.
        now = time.time()
        if self._token and now < self._token_expires_at - 60:
            return self._token
        result = self._get_msal().acquire_token_for_client(scopes=SCOPE)
        if "access_token" not in result:
            raise AuthError(
                f"MSAL token acquisition failed: "
                f"{result.get('error_description', result.get('error', 'unknown'))}"
            )
        self._token = result["access_token"]
        self._token_expires_at = now + int(result.get("expires_in", 3600))
        return self._token

    def _headers(self, extra: dict[str, str] | None = None) -> dict[str, str]:
        h = {"Authorization": f"Bearer {self._get_token()}"}
        if extra:
            h.update(extra)
        return h

    # --- HTTP plumbing -----------------------------------------------------
    def _request(self, method: str, url: str, **kwargs) -> httpx.Response:
        if not url.startswith("http"):
            url = GRAPH_BASE + url
        headers = kwargs.pop("headers", {})
        kwargs["headers"] = self._headers(headers)
        # httpx connection errors propagate up to the @retry_on_transient decorator.
        resp = self._http.request(method, url, **kwargs)
        if resp.status_code >= 400:
            raise _http_to_storage_error(resp)
        return resp

    # --- site/drive resolution --------------------------------------------
    def _resolve_site_and_drive(self) -> tuple[str, str]:
        if self._site_id and self._drive_id:
            return self._site_id, self._drive_id
        # GET /sites/{hostname}:{site-path}
        site_path = self._site_path if self._site_path.startswith("/") else "/" + self._site_path
        url = f"/sites/{self._site_hostname}:{site_path}"
        site = self._request("GET", url).json()
        self._site_id = site["id"]
        drive = self._request("GET", f"/sites/{self._site_id}/drive").json()
        self._drive_id = drive["id"]
        return self._site_id, self._drive_id

    @staticmethod
    def _encode_path(path: str) -> str:
        """Encode a folder/file path for use after `root:`."""
        # Path segments are URL-encoded; '/' separators are preserved.
        return urllib.parse.quote(path.lstrip("/"), safe="/")

    def _item_url(self, path: str, suffix: str = "") -> str:
        """Build /drives/{id}/root:/path/to/file:[suffix]."""
        _, drive_id = self._resolve_site_and_drive()
        p = normalize_path(path)
        if p == "/":
            return f"/drives/{drive_id}/root{suffix}" if not suffix else f"/drives/{drive_id}/root{suffix}"
        encoded = self._encode_path(p)
        return f"/drives/{drive_id}/root:/{encoded}:{suffix}" if suffix else f"/drives/{drive_id}/root:/{encoded}"

    def _item_id_url(self, item_id: str, suffix: str = "") -> str:
        _, drive_id = self._resolve_site_and_drive()
        return f"/drives/{drive_id}/items/{item_id}{suffix}"

    @staticmethod
    def _entry_to_dict(item: dict) -> dict:
        is_folder = "folder" in item
        full_path = ""
        ref = item.get("parentReference") or {}
        parent = ref.get("path", "")
        if parent.startswith("/drive/root:") or "/root:" in parent:
            parent = parent.split("root:", 1)[-1]
        name = item.get("name", "")
        if parent and name:
            full_path = parent.rstrip("/") + "/" + name
        elif name:
            full_path = "/" + name
        return {
            "name": name,
            "path": full_path or "/" + name,
            "type": "folder" if is_folder else "file",
            "size": item.get("size", 0),
            "modified": item.get("lastModifiedDateTime"),
            "content_hash": (item.get("file") or {}).get("hashes", {}).get("quickXorHash")
                            or (item.get("file") or {}).get("hashes", {}).get("sha256Hash"),
            "id": item.get("id"),
            "etag": item.get("cTag") or item.get("eTag"),
        }

    # --- core ops ----------------------------------------------------------
    @retry_on_transient(transient_exceptions=_TRANSIENT)
    def read(self, path: str) -> bytes:
        p = normalize_path(path)
        with log_operation(self.backend_name, "read", p):
            url = self._item_url(p, "/content")
            resp = self._request("GET", url, follow_redirects=True)
            return resp.content

    @retry_on_transient(transient_exceptions=_TRANSIENT)
    def write(self, path: str, data: bytes, overwrite: bool = True) -> dict:
        p = normalize_path(path)
        with log_operation(self.backend_name, "write", p):
            self._ensure_parent(p)
            if len(data) <= _UPLOAD_SINGLE_LIMIT:
                return self._simple_upload(p, data, overwrite)
            return self._chunked_upload(p, data, overwrite)

    def _simple_upload(self, path: str, data: bytes, overwrite: bool) -> dict:
        url = self._item_url(path, "/content")
        if not overwrite:
            url += "?@microsoft.graph.conflictBehavior=fail"
        else:
            url += "?@microsoft.graph.conflictBehavior=replace"
        headers = {"Content-Type": "application/octet-stream"}
        resp = self._request("PUT", url, content=data, headers=headers)
        return self._entry_to_dict(resp.json())

    def _chunked_upload(self, path: str, data: bytes, overwrite: bool) -> dict:
        # 1. Create upload session.
        url = self._item_url(path, "/createUploadSession")
        body = {
            "item": {
                "@microsoft.graph.conflictBehavior": "replace" if overwrite else "fail"
            }
        }
        session = self._request("POST", url, json=body).json()
        upload_url = session["uploadUrl"]

        # 2. Upload chunks. The upload URL is pre-authenticated.
        total = len(data)
        offset = 0
        last_resp: httpx.Response | None = None
        while offset < total:
            chunk = data[offset : offset + _UPLOAD_CHUNK_SIZE]
            end = offset + len(chunk) - 1
            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {offset}-{end}/{total}",
            }
            last_resp = self._http.put(upload_url, content=chunk, headers=headers)
            if last_resp.status_code >= 400:
                raise _http_to_storage_error(last_resp)
            offset += len(chunk)

        assert last_resp is not None
        return self._entry_to_dict(last_resp.json())

    @retry_on_transient(transient_exceptions=_TRANSIENT)
    def list(self, path: str, recursive: bool = False) -> list[dict]:
        p = normalize_path(path)
        with log_operation(self.backend_name, "list", p):
            return self._list_inner(p, recursive)

    def _list_inner(self, path: str, recursive: bool) -> list[dict]:
        url = self._item_url(path, "/children")
        results: list[dict] = []
        while url:
            resp = self._request("GET", url).json()
            for item in resp.get("value", []):
                d = self._entry_to_dict(item)
                results.append(d)
                if recursive and d["type"] == "folder":
                    child_path = path.rstrip("/") + "/" + d["name"] if path != "/" else "/" + d["name"]
                    results.extend(self._list_inner(child_path, recursive=True))
            url = resp.get("@odata.nextLink", "")
            # nextLink is a full URL; bypass the GRAPH_BASE prefix.
            if url and url.startswith(GRAPH_BASE):
                url = url[len(GRAPH_BASE):]
        return results

    @retry_on_transient(transient_exceptions=_TRANSIENT)
    def delete(self, path: str) -> None:
        p = normalize_path(path)
        with log_operation(self.backend_name, "delete", p):
            try:
                self._request("DELETE", self._item_url(p))
            except NotFoundError:
                return  # idempotent

    @retry_on_transient(transient_exceptions=_TRANSIENT)
    def move(self, src: str, dst: str) -> None:
        s = normalize_path(src)
        d = normalize_path(dst)
        with log_operation(self.backend_name, "move", f"{s} -> {d}"):
            self._ensure_parent(d)
            dst_parent, dst_name = split_parent(d)
            parent_meta = self._get_item(dst_parent)
            body = {
                "parentReference": {"id": parent_meta["id"]},
                "name": dst_name,
            }
            self._request("PATCH", self._item_url(s), json=body)

    @retry_on_transient(transient_exceptions=_TRANSIENT)
    def copy(self, src: str, dst: str) -> None:
        s = normalize_path(src)
        d = normalize_path(dst)
        with log_operation(self.backend_name, "copy", f"{s} -> {d}"):
            self._ensure_parent(d)
            dst_parent, dst_name = split_parent(d)
            parent_meta = self._get_item(dst_parent)
            body = {
                "parentReference": {"driveId": parent_meta["parentReference"]["driveId"],
                                    "id": parent_meta["id"]},
                "name": dst_name,
            }
            # Graph copy is async — returns 202 with Location header to poll.
            resp = self._request("POST", self._item_url(s, "/copy"), json=body)
            monitor = resp.headers.get("Location")
            if monitor:
                self._await_async_op(monitor)

    def _await_async_op(self, monitor_url: str, timeout_s: float = 120.0) -> None:
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            resp = self._http.get(monitor_url)
            if resp.status_code == 200:
                status = resp.json().get("status")
                if status == "completed":
                    return
                if status in ("failed", "cancelled"):
                    raise BackendError(f"Async op {status}: {resp.text}")
            elif resp.status_code in (201, 303):
                return
            time.sleep(1.0)
        raise BackendError(f"Async op did not complete within {timeout_s}s")

    @retry_on_transient(transient_exceptions=_TRANSIENT)
    def ensure_folder(self, path: str) -> None:
        p = normalize_path(path)
        if p == "/":
            return
        with log_operation(self.backend_name, "ensure_folder", p):
            self._ensure_folder_recursive(p)

    def _ensure_folder_recursive(self, path: str) -> None:
        # Walk from root creating each missing segment.
        segments = [s for s in path.split("/") if s]
        current = ""
        _, drive_id = self._resolve_site_and_drive()
        for seg in segments:
            parent = current or "/"
            current = current + "/" + seg
            existing = self._get_item_or_none(current)
            if existing and "folder" in existing:
                continue
            if existing and "file" in existing:
                raise BackendError(f"{current} exists as a file, cannot create folder")
            if parent == "/":
                children_url = f"/drives/{drive_id}/root/children"
            else:
                children_url = self._item_url(parent, "/children")
            body = {
                "name": seg,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "replace",  # idempotent
            }
            self._request("POST", children_url, json=body)

    def _ensure_parent(self, path: str) -> None:
        parent, _ = split_parent(path)
        if parent and parent != "/":
            self._ensure_folder_recursive(parent)

    def _get_item(self, path: str) -> dict:
        p = normalize_path(path)
        if p == "/":
            _, drive_id = self._resolve_site_and_drive()
            return self._request("GET", f"/drives/{drive_id}/root").json()
        return self._request("GET", self._item_url(p)).json()

    def _get_item_or_none(self, path: str) -> dict | None:
        try:
            return self._get_item(path)
        except NotFoundError:
            return None

    @retry_on_transient(transient_exceptions=_TRANSIENT)
    def get_metadata(self, path: str) -> dict:
        p = normalize_path(path)
        with log_operation(self.backend_name, "get_metadata", p):
            try:
                item = self._get_item(p)
                d = self._entry_to_dict(item)
                d["exists"] = True
                return d
            except NotFoundError:
                return {"exists": False, "path": p, "name": p.rsplit("/", 1)[-1]}

    # --- locking (native checkout/checkin) --------------------------------
    @retry_on_transient(transient_exceptions=_TRANSIENT)
    def checkout(self, path: str) -> None:
        p = normalize_path(path)
        with log_operation(self.backend_name, "checkout", p):
            item = self._get_item(p)
            self._request("POST", self._item_id_url(item["id"], "/checkout"))
            self._held_locks.add(p)

    @retry_on_transient(transient_exceptions=_TRANSIENT)
    def checkin(self, path: str) -> None:
        p = normalize_path(path)
        with log_operation(self.backend_name, "checkin", p):
            item = self._get_item(p)
            self._request(
                "POST",
                self._item_id_url(item["id"], "/checkin"),
                json={"comment": "tropi-storage-adapter checkin"},
            )
            self._held_locks.discard(p)

    # --- conditional write -------------------------------------------------
    def write_with_etag(self, path: str, data: bytes, etag: str) -> dict:
        p = normalize_path(path)
        with log_operation(self.backend_name, "write_with_etag", p):
            self._ensure_parent(p)
            url = self._item_url(p, "/content?@microsoft.graph.conflictBehavior=replace")
            headers = {
                "Content-Type": "application/octet-stream",
                "If-Match": etag,
            }
            try:
                resp = self._request("PUT", url, content=data, headers=headers)
            except ConflictError as e:
                # 412 is mapped to ConflictError already; re-raise with cleaner msg.
                raise ConflictError(f"{p} changed since etag {etag!r} was read") from e
            return self._entry_to_dict(resp.json())

    # --- health ------------------------------------------------------------
    def healthcheck(self) -> dict[str, Any]:
        start = time.perf_counter()
        authenticated = False
        can_list_root = False
        try:
            self._get_token()
            authenticated = True
        except Exception:
            pass
        try:
            self._resolve_site_and_drive()
            self._request("GET", self._item_url("/", "/children") + "?$top=1")
            can_list_root = True
        except Exception:
            pass
        return {
            "backend": self.backend_name,
            "authenticated": authenticated,
            "can_list_root": can_list_root,
            "latency_ms": round((time.perf_counter() - start) * 1000, 2),
        }
