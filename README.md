# tropi-storage-adapter

Unified storage adapter. Pick a backend at runtime via an env var; the rest of your code is identical.

```python
from tropi_storage import get_adapter

storage = get_adapter()  # Dropbox or M365 (SharePoint/Graph) — picked by STORAGE_BACKEND
data = storage.read("/Documents/example.xlsx")
storage.write("/Documents/example.xlsx", new_bytes)
```

Supports gradual migration from Dropbox to Microsoft 365 (SharePoint) with no code changes in consuming services.

---

## Install

From the local checkout:

```bash
pip install -e .
```

From GitHub (when published):

```bash
pip install git+https://github.com/your-org/tropi-storage-adapter.git
```

---

## Configure

Copy `env.example` to `.env` (the `env.example` filename has no leading dot to keep it safely editable) and fill in values.

| Variable | When | Notes |
|---|---|---|
| `STORAGE_BACKEND` | always | `dropbox` (default) or `m365` |
| `DROPBOX_APP_KEY` / `_APP_SECRET` / `_REFRESH_TOKEN` | `dropbox` | Refresh-token OAuth flow |
| `M365_TENANT_ID` / `_CLIENT_ID` / `_CLIENT_SECRET` | `m365` | App registration for client-credentials flow |
| `M365_SITE_HOSTNAME` | `m365` | e.g. `yourcompany.sharepoint.com` |
| `M365_SITE_PATH` | `m365` | e.g. `/sites/YourSite` |
| `SENTRY_DSN` | optional | If set, captures unhandled exceptions automatically |
| `LOG_LEVEL` | optional | `INFO` (default), `DEBUG`, etc. |
| `STORAGE_MAX_RETRIES` | optional | Default `5` for transient errors |
| `INTEGRATION_TESTS` | tests only | Set to `1` to run live tests against a real backend |

Secrets must never be committed. `.env` is gitignored.

---

## Interface

`StorageAdapter` exposes the same surface for both backends:

| Method | Purpose |
|---|---|
| `read(path) -> bytes` | Download file |
| `write(path, data, overwrite=True) -> dict` | Upload file (chunked > 4 MiB on Graph, > 140 MiB on Dropbox) |
| `list(path, recursive=False) -> list[dict]` | List folder, paginates internally |
| `delete(path)` | Delete file/folder; idempotent |
| `move(src, dst)` | Move/rename |
| `copy(src, dst)` | Copy file (Graph copy is async; we poll) |
| `ensure_folder(path)` | Create folder + parents; idempotent |
| `get_metadata(path) -> dict` | Includes `exists: bool`; never raises NotFoundError |
| `checkout(path)` | Exclusive lock — Graph native; Dropbox simulated via sibling `.lock` file with adapter-instance UUID |
| `checkin(path)` | Release lock |
| `write_with_etag(path, data, etag) -> dict` | Conditional write; raises `ConflictError` if changed |
| `healthcheck() -> dict` | `{backend, authenticated, can_list_root, latency_ms}` |

Path templating helper:

```python
from tropi_storage import expand_path
expand_path("/2026/{year}/{ww}/file.xlsx")  # -> /2026/2026/18/file.xlsx
```

Tokens: `{year}`, `{month}`, `{day}` (zero-padded), `{ww}` (ISO week, zero-padded).

---

## Exceptions

All backend errors are translated to:

- `StorageError` (base)
- `NotFoundError` — path missing
- `ConflictError` — etag mismatch / 412 / Graph 409
- `AuthError` — 401 / 403 / token failure
- `ThrottledError` — 429 / Dropbox rate-limit (carries `retry_after`)
- `LockError` — already locked / not held
- `BackendError` — anything else (carries `status_code`)

---

## Observability

Every operation emits one structured JSON log line:

```json
{"timestamp":"2026-05-03T08:42:11","level":"INFO","logger":"tropi_storage",
 "message":"read ok","backend":"m365","operation":"read","path":"/Documents/example.xlsx",
 "duration_ms":124.3,"result":"success"}
```

Set `LOG_LEVEL=DEBUG` for verbose request/response details. If `SENTRY_DSN` is set and `sentry-sdk` is installed (`pip install tropi-storage-adapter[sentry]`), uncaught exceptions are reported automatically.

---

## Backend notes

**Dropbox.** Uses the official `dropbox==12.0.2` SDK pinned to match existing services. `etag` ↔ Dropbox `content_hash` (sha256). `checkout`/`checkin` write a sibling `.lock` file containing the adapter-instance UUID; another process taking the same lock raises `LockError`.

**M365 / Graph.** Uses `httpx` + `msal` directly rather than the async-first `msgraph-sdk` — keeps the API synchronous, which matches every consuming service. Site and drive IDs are resolved once and cached. `checkout`/`checkin` use the native Graph endpoints. `etag` ← `cTag` (content-only changes, not metadata-only). 429s honor `Retry-After`.

---

## Testing

```bash
pip install -e ".[dev]"
pytest -v
```

Unit tests use mocked SDKs / mocked `httpx.MockTransport`. Set `INTEGRATION_TESTS=1` and provide real M365 credentials to exercise live calls against a real SharePoint site.

---

## Migration sequence

Switch `STORAGE_BACKEND` per service in your deployment platform environment. Start with read-only services to validate before enabling writes.
