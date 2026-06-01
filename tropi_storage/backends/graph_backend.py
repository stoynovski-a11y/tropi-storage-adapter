"""Microsoft Graph (SharePoint / OneDrive) backend.

Uses `httpx` against the Graph REST API and `msal` for client-credentials OAuth.
This is a deliberate divergence from the spec's `msgraph-sdk` recommendation —
the official SDK is async-first (Kiota-generated) and would force every adapter
call through `asyncio.run()`, which is awkward for the existing sync services.
The functionality is identical; only the transport differs.

Multi-site routing
------------------
The first path segment selects a (site_path, library_name) pair via
`tropi_storage.routing`.  Paths whose first segment is not in the route table
fall back to the default site/drive configured via `M365_SITE_PATH` and
`M365_DEFAULT_LIBRARY`.  Services that were written before routing was added
therefore continue to work without changes.
"""
from __future__ import annotations

import os
import threading
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
from ..routing import load_routes, load_strip_prefix, resolve_route

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
    """SharePoint document library accessed via Microsoft Graph.

    ``site_path`` is now **optional**.  When omitted, every logical path
    *must* match a route in ``self._routes`` (populated from
    ``DEFAULT_ROUTES`` + optional ``M365_ROUTES`` env var); unmatched paths
    raise ``BackendError``.  If a default site is configured it acts as a
    catch-all for unmatched paths, preserving backward-compatibility.
    """

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

        # site_path is optional — if absent, all paths must match a route.
        _sp = site_path or os.getenv("M365_SITE_PATH", "") or None
        self._default_site_path: str | None = _sp

        # Named library to use for the default site (optional).
        self._default_drive_name: str | None = os.getenv("M365_DEFAULT_LIBRARY") or None

        if not all([self._tenant_id, self._client_id, self._client_secret,
                    self._site_hostname]):
            raise AuthError(
                "Graph backend requires M365_TENANT_ID, M365_CLIENT_ID, "
                "M365_CLIENT_SECRET, M365_SITE_HOSTNAME env vars."
            )

        self._http: httpx.Client = http_client or httpx.Client(timeout=60.0)
        self._msal_app: msal.ConfidentialClientApplication | None = None
        self._token: str | None = None
        self._token_expires_at: float = 0.0

        # Per-(site_path) site-id cache and per-(site_path, drive_name) drive-id cache.
        # Both are guarded by a lock because APScheduler threads may call concurrently.
        self._site_id_cache: dict[str, str] = {}
        self._drive_id_cache: dict[tuple[str, str | None], str] = {}
        self._cache_lock = threading.Lock()

        # Lock-tracking (used by checkout/checkin).
        self._held_locks: set[str] = set()

        # Route table: first-segment → (site_path, library_name).
        self._routes = load_routes()

        # Optional legacy-prefix to strip from every path before routing.
        self._strip_prefix = load_strip_prefix()

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
        # Refresh ~60 s before expiry.
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

    # --- site/drive resolution (multi-site, cached) -----------------------

    def _resolve_site_id(self, site_path: str) -> str:
        """Return the Graph site id for *site_path*, using the cache."""
        with self._cache_lock:
            if site_path in self._site_id_cache:
                return self._site_id_cache[site_path]

        sp = site_path if site_path.startswith("/") else "/" + site_path
        url = f"/sites/{self._site_hostname}:{sp}"
        site = self._request("GET", url).json()
        site_id: str = site["id"]

        with self._cache_lock:
            self._site_id_cache[site_path] = site_id
        return site_id

    def _resolve_drive_id(self, site_path: str, drive_name: str | None) -> str:
        """Return the Graph drive id for the given site + library name.

        If *drive_name* is falsy the site's default drive is used.
        Raises BackendError if a named library cannot be found.
        """
        key = (site_path, drive_name)
        with self._cache_lock:
            if key in self._drive_id_cache:
                return self._drive_id_cache[key]

        site_id = self._resolve_site_id(site_path)

        if not drive_name:
            # Default drive.
            drive = self._request("GET", f"/sites/{site_id}/drive").json()
            drive_id: str = drive["id"]
        else:
            # Walk (paginated) drives list to find matching library by name.
            drive_id = self._find_drive_by_name(site_id, site_path, drive_name)

        with self._cache_lock:
            self._drive_id_cache[key] = drive_id
        return drive_id

    def _find_drive_by_name(self, site_id: str, site_path: str, drive_name: str) -> str:
        """Paginate /sites/{site_id}/drives until ``drive_name`` is found."""
        url: str = f"/sites/{site_id}/drives"
        while url:
            resp = self._request("GET", url).json()
            for entry in resp.get("value", []):
                if entry.get("name") == drive_name:
                    return entry["id"]
            next_link: str = resp.get("@odata.nextLink", "")
            if next_link and next_link.startswith(GRAPH_BASE):
                next_link = next_link[len(GRAPH_BASE):]
            url = next_link
        raise BackendError(
            f"Library {drive_name!r} not found in site {site_path!r}."
        )

    def _resolve_default_site_and_drive(self) -> tuple[str, str]:
        """Resolve the default site + drive ids.  Raises BackendError if no default configured."""
        if not self._default_site_path:
            raise BackendError(
                "No default site configured (M365_SITE_PATH) and no route matched."
            )
        drive_id = self._resolve_drive_id(self._default_site_path, self._default_drive_name)
        site_id = self._resolve_site_id(self._default_site_path)
        return site_id, drive_id

    # Thin backward-compat alias used only by the old TestSiteResolution test.
    def _resolve_site_and_drive(self) -> tuple[str, str]:
        """Backward-compatible: resolve the *default* site+drive ids."""
        return self._resolve_default_site_and_drive()

    # --- routing helpers ---------------------------------------------------

    def _route(self, path: str) -> tuple[str, str]:
        """Resolve *path* to (drive_id, item_path) via the route table."""
        site_path, drive_name, item_path = resolve_route(
            path, self._routes, self._default_site_path, self._default_drive_name,
            strip_prefix=self._strip_prefix,
        )
        drive_id = self._resolve_drive_id(site_path, drive_name)
        return drive_id, item_path

    # --- URL builders, drive-aware ----------------------------------------

    @staticmethod
    def _encode_path(path: str) -> str:
        """Encode a folder/file path for use after ``root:``."""
        return urllib.parse.quote(path.lstrip("/"), safe="/")

    def _build_item_url(self, drive_id: str, item_path: str, suffix: str = "") -> str:
        """Return a /drives/{drive_id}/root[:/path][:suffix] URL fragment."""
        p = normalize_path(item_path)
        if p == "/":
            return f"/drives/{drive_id}/root{suffix}"
        enc = self._encode_path(p)
        if suffix:
            return f"/drives/{drive_id}/root:/{enc}:{suffix}"
        return f"/drives/{drive_id}/root:/{enc}"

    def _item_url(self, path: str, suffix: str = "") -> str:
        """Route *path* then build /drives/{id}/root:/path:[suffix]."""
        drive_id, ip = self._route(path)
        return self._build_item_url(drive_id, ip, suffix)

    def _item_id_url(self, drive_id: str, item_id: str, suffix: str = "") -> str:
        """Return /drives/{drive_id}/items/{item_id}{suffix}."""
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

    # --- item helpers -------------------------------------------------------

    def _get_item(self, path: str) -> dict:
        """Fetch the Graph item metadata for *path*."""
        drive_id, ip = self._route(path)
        if ip == "/":
            return self._request("GET", f"/drives/{drive_id}/root").json()
        return self._request("GET", self._build_item_url(drive_id, ip)).json()

    def _get_item_or_none(self, drive_id: str, item_path: str) -> dict | None:
        """Fetch item metadata; return None on 404."""
        try:
            p = normalize_path(item_path)
            if p == "/":
                return self._request("GET", f"/drives/{drive_id}/root").json()
            return self._request("GET", self._build_item_url(drive_id, p)).json()
        except NotFoundError:
            return None

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
        base = normalize_path(path)
        while url:
            resp = self._request("GET", url).json()
            for item in resp.get("value", []):
                d = self._entry_to_dict(item)
                # Return the *logical* path (caller namespace + name), not the
                # drive-relative path that Graph's parentReference yields. Under
                # multi-site routing the leading route segment is consumed during
                # resolution, so a drive-relative path would NOT round-trip back
                # through read()/move()/copy() — it would fall through to the
                # default site and 404 ("Requested site could not be found").
                # Rebuilding from the caller's input path keeps list() output
                # usable as input to every other op. (No-op for single-site,
                # where base already equals the drive-relative parent.)
                d["path"] = (
                    base.rstrip("/") + "/" + d["name"] if base != "/" else "/" + d["name"]
                )
                results.append(d)
                if recursive and d["type"] == "folder":
                    # Recurse on the logical child path so routing re-applies.
                    results.extend(self._list_inner(d["path"], recursive=True))
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
            # Cross-library moves are not supported by a single PATCH.
            src_drive, _ = self._route(s)
            dst_drive, _ = self._route(d)
            if src_drive != dst_drive:
                raise BackendError(
                    "Cross-library move is not supported; use copy + delete instead."
                )
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
                "parentReference": {
                    "driveId": parent_meta.get("parentReference", {}).get("driveId")
                               or parent_meta.get("parentReference", {}).get("id"),
                    "id": parent_meta["id"],
                },
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
        """Create each missing folder segment under the correct library drive."""
        drive_id, item_path = self._route(path)
        if item_path == "/":
            return  # Library root always exists.

        segments = [s for s in item_path.split("/") if s]
        current_item_path = ""
        for seg in segments:
            parent_item_path = current_item_path or "/"
            current_item_path = current_item_path + "/" + seg
            existing = self._get_item_or_none(drive_id, current_item_path)
            if existing and "folder" in existing:
                continue
            if existing and "file" in existing:
                raise BackendError(
                    f"{current_item_path} exists as a file, cannot create folder"
                )
            if parent_item_path == "/":
                children_url = f"/drives/{drive_id}/root/children"
            else:
                children_url = self._build_item_url(drive_id, parent_item_path, "/children")
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
            drive_id, _ = self._route(p)
            self._request("POST", self._item_id_url(drive_id, item["id"], "/checkout"))
            self._held_locks.add(p)

    @retry_on_transient(transient_exceptions=_TRANSIENT)
    def checkin(self, path: str) -> None:
        p = normalize_path(path)
        with log_operation(self.backend_name, "checkin", p):
            item = self._get_item(p)
            drive_id, _ = self._route(p)
            self._request(
                "POST",
                self._item_id_url(drive_id, item["id"], "/checkin"),
                json={"comment": "storage adapter checkin"},
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
            # Pick an appropriate probe path.
            probe_env = os.getenv("M365_HEALTHCHECK_PATH", "").strip()
            if probe_env:
                probe_path = probe_env
            elif self._routes:
                probe_path = "/" + next(iter(self._routes))
            else:
                probe_path = "/"
            self.list(probe_path)
            can_list_root = True
        except Exception:
            pass
        return {
            "backend": self.backend_name,
            "authenticated": authenticated,
            "can_list_root": can_list_root,
            "latency_ms": round((time.perf_counter() - start) * 1000, 2),
        }
