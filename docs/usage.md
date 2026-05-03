# Usage examples

Drop-in replacement for ad-hoc Dropbox SDK calls. Same code, two backends — choose at runtime via `STORAGE_BACKEND`.

## Setup

```python
from tropi_storage import get_adapter

storage = get_adapter()
```

The factory reads `STORAGE_BACKEND` env var (`dropbox` or `m365`), initializes JSON logging, and (if `SENTRY_DSN` is set) Sentry.

---

## Read a file

```python
data = storage.read("/Co/A_23 TropiCommodity Acc Мере/DOCS/foo.xlsx")
print(len(data), "bytes")
```

Raises `NotFoundError` if missing.

---

## Write a file

```python
new_bytes = open("/tmp/report.xlsx", "rb").read()
meta = storage.write("/Co/foo/report.xlsx", new_bytes, overwrite=True)
print(meta["etag"], meta["size"])
```

Returns `{name, path, type, size, modified, content_hash, id, etag}`. Parent folders are created automatically on the Graph backend; Dropbox auto-creates them too.

---

## List a folder (paginated)

```python
items = storage.list("/Co/foo", recursive=False)
for item in items:
    print(item["type"], item["name"], item["size"])
```

Pagination is handled internally — for huge folders, all pages are fetched. Use `recursive=True` to walk subfolders.

---

## Delete (idempotent)

```python
storage.delete("/Co/foo/old.xlsx")  # safe to call even if the file is gone
```

---

## Move and copy

```python
storage.move("/Inbox/file.pdf", "/Processed/file.pdf")
storage.copy("/Templates/master.xlsx", "/Drafts/master.xlsx")
```

On Graph, copy is asynchronous; the adapter polls the operation until completion (default timeout 120s).

---

## Ensure a folder exists (idempotent)

```python
storage.ensure_folder("/Reports/2026/05/03")
```

Creates every missing segment, ignores already-existing folders.

---

## Check a path without raising

```python
meta = storage.get_metadata("/Co/maybe.xlsx")
if meta["exists"]:
    print(meta["etag"], meta["size"])
else:
    print("not there")
```

---

## Year/date templating in paths

Stop hardcoding `2025` / `2026` in every service:

```python
from tropi_storage import expand_path

template = "/2021 Имена Кодове/03 Продукция/{year} Ф1 Сведения Производство/{year}_{ww}_report.xlsx"
path = expand_path(template)
# -> "/2021 Имена Кодове/03 Продукция/2026 Ф1 Сведения Производство/2026_18_report.xlsx"

storage.write(path, data)
```

Tokens: `{year}`, `{month}` (01-12), `{day}` (01-31), `{ww}` (ISO week 01-53).

---

## Conditional write (etag)

Optimistic concurrency — fail loudly if the file changed since you read it:

```python
from tropi_storage import ConflictError

meta = storage.get_metadata("/Co/register.xlsx")
data = storage.read("/Co/register.xlsx")

# ... mutate data ...

try:
    storage.write_with_etag("/Co/register.xlsx", new_data, etag=meta["etag"])
except ConflictError:
    print("Someone else wrote the file. Re-read and retry.")
```

---

## Locking (checkout / checkin)

Prevents two services writing the same file at the same time.

```python
from tropi_storage import LockError

try:
    storage.checkout("/Каси/Каса Бургас/Каса_Бургас_Draft.xlsx")
except LockError:
    print("Another instance is editing this file; try later.")
    return

try:
    data = storage.read("/Каси/Каса Бургас/Каса_Бургас_Draft.xlsx")
    # ... mutate ...
    storage.write("/Каси/Каса Бургас/Каса_Бургас_Draft.xlsx", new_data)
finally:
    storage.checkin("/Каси/Каса Бургас/Каса_Бургас_Draft.xlsx")
```

- **Graph backend**: native `POST /checkout` / `POST /checkin`. The file is hidden from other users until checkin.
- **Dropbox backend**: simulated via a sibling `.lock` file containing the adapter-instance UUID. Re-entrant within the same `StorageAdapter` instance only.

---

## Health check (for `/health` endpoints)

```python
from fastapi import FastAPI
from tropi_storage import get_adapter

app = FastAPI()
storage = get_adapter()

@app.get("/health")
def health():
    return storage.healthcheck()
    # -> {"backend": "m365", "authenticated": true, "can_list_root": true, "latency_ms": 87.4}
```

---

## Errors

```python
from tropi_storage import (
    NotFoundError, ConflictError, AuthError,
    ThrottledError, LockError, BackendError, StorageError,
)
```

`ThrottledError` and connection-level errors are retried automatically with exponential backoff (defaults: 5 retries, 1s/2s/4s/8s/16s, honoring `Retry-After`). Tune via `STORAGE_MAX_RETRIES`.

---

## Switching backends

To migrate a service from Dropbox to M365, set on Railway/Vercel:

```
STORAGE_BACKEND=m365
M365_TENANT_ID=...
M365_CLIENT_ID=...
M365_CLIENT_SECRET=...
M365_SITE_HOSTNAME=tropicommodity.sharepoint.com
M365_SITE_PATH=/sites/Multipack
```

No code changes. Same paths (`/Co/...`) — the Graph backend translates them to `sites/{siteId}/drive/root:/Co/...` internally.

For the testing pilot, point at `/sites/Playgroud` instead of `/sites/Multipack`.
