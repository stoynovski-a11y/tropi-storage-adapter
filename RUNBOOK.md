# RUNBOOK — tropi-storage-adapter (shared library)

**Repo:** `~/dev/tropi-storage-adapter` · GitHub `stoynovski-a11y/tropi-storage-adapter` — **PUBLIC repo, never commit secrets or real site/route names.** This RUNBOOK is local-only (untracked) and contains real fleet values — do not commit it as-is.
**Type:** Python shared library (pip-installed from a pinned git commit). NOT a deployed service — no Railway project, no Procfile, no Dockerfile, no CI workflows.
**Package:** `tropi-storage-adapter` v0.1.0, import name `tropi_storage` (Python ≥3.10).

## Purpose

One file-storage interface for the whole fleet. Every service that reads/writes files (invoices, bank statements, registers, templates) calls this library instead of talking to SharePoint directly. The library translates a logical path like `/Top Segment/sub/x.pdf` into Microsoft Graph API calls against the right SharePoint site and document library. It was built to migrate the fleet off Dropbox: services switched backends by flipping one env var, with zero code changes. The Dropbox backend is still in the code **only as a rollback escape hatch** — Dropbox itself is decommissioned.

In plain language: it's the "post office" of the fleet. Services hand it an address (path) and a parcel (bytes); it knows which SharePoint building and floor that address maps to.

## Public API surface

Everything consumers use is exported from `tropi_storage/__init__.py`:

| Symbol | What it does |
|---|---|
| `get_adapter()` | Factory — returns the backend chosen by `STORAGE_BACKEND` (default **m365**). Also configures JSON logging + Sentry. `tropi_storage/adapter.py:99` |
| `StorageAdapter` methods | `read(path)→bytes`, `write(path, data, overwrite=True)→dict`, `list(path, recursive=False)→list[dict]`, `delete(path)` (idempotent), `move(src, dst)`, `copy(src, dst)`, `ensure_folder(path)`, `get_metadata(path)→dict` (`exists: False` instead of raising), `checkout/checkin(path)` (SharePoint file lock), `write_with_etag(path, data, etag)` (conditional write), `healthcheck()→dict` |
| `expand_path(template)` | Expands `{year}/{month}/{day}/{ww}` (ISO week) in path templates. `path_utils.py` |
| `normalize_path`, `split_parent` | Path helpers (leading slash, no trailing slash) |
| Exceptions | `StorageError` base; `NotFoundError`, `ConflictError`, `AuthError`, `ThrottledError` (has `.retry_after`), `LockError`, `BackendError` (has `.status_code`). `exceptions.py` |

All paths are POSIX-style absolute (`/Top Segment/sub/file.xlsx`). Item dicts contain `name, path, type, size, modified, content_hash, id, etag` (shape documented in `adapter.py:10-21`).

## How a path becomes a Graph call (data flow)

```
logical path "/Legacy Prefix/Top Segment/sub/x.pdf"
  → strip M365_STRIP_PREFIX ("/Legacy Prefix")               routing.py:load_strip_prefix
  → first segment "Top Segment" looked up in M365_ROUTES     routing.py:resolve_route
      = ["/sites/alpha", "Library A"]  (site, library)
  → resolve site-id + drive-id via Graph (cached per process) graph_backend.py:_resolve_drive_id
  → Graph REST call /drives/{drive_id}/root:/sub/x.pdf        httpx, MSAL client-credentials token
```

(Values above are illustrative — the real prefix and route map live only in each service's Railway env vars, never in this repo.)

- Routes come **only** from the `M365_ROUTES` env var (JSON); `DEFAULT_ROUTES` in `routing.py` is deliberately empty because the repo is public.
- Unmatched first segments fall through to the default site (`M365_SITE_PATH` + `M365_DEFAULT_LIBRARY`) if set; otherwise `BackendError`.
- Uploads >4 MiB go through a chunked Graph upload session (`_chunked_upload`). `copy()` is async on Graph — the library polls the monitor URL with progressive backoff 0.25→2.0 s, 120 s timeout (`_await_async_op`).
- Auth: MSAL `ConfidentialClientApplication`, app-only token for `https://graph.microsoft.com/.default`, refreshed 60 s before expiry (`graph_backend.py:_get_token`).

## Key files

| File | What |
|---|---|
| `tropi_storage/adapter.py` | Abstract interface + `get_adapter()` factory (backend selection, m365 default) |
| `tropi_storage/backends/graph_backend.py` | The real backend — Graph/SharePoint, 717 lines, all gotcha fixes live here |
| `tropi_storage/backends/dropbox_backend.py` | Legacy rollback backend (Dropbox SDK) |
| `tropi_storage/routing.py` | Multi-site router: `M365_ROUTES` parsing, `M365_STRIP_PREFIX`, `resolve_route()` |
| `tropi_storage/path_utils.py` | `expand_path` date templating, normalization |
| `tropi_storage/retry.py` | `@retry_on_transient` — exp backoff 1→30 s, jitter, honors `Retry-After`, `STORAGE_MAX_RETRIES` |
| `tropi_storage/exceptions.py` | Exception hierarchy |
| `tropi_storage/logging_config.py` | One JSON log line per operation (logger `tropi_storage`); secrets redacted; Sentry init |
| `scripts/smoke_test_m365.py` | Manual live smoke test |
| `env.example` | Env template (no leading dot on purpose) |

## Env vars (read by the library, set on each consuming service)

| Var | What it does | Secret | Notes |
|---|---|---|---|
| `STORAGE_BACKEND` | `m365` (default) or `dropbox` | n | **Rollback flag.** `dropbox` flips the whole service back to the Dropbox backend — pointless now (Dropbox decommissioned, creds deleted). Default changed dropbox→m365 in commit `b17946e` so an unset/typo'd var can never silently hit dead Dropbox. |
| `M365_TENANT_ID` / `M365_CLIENT_ID` / `M365_CLIENT_SECRET` | Entra app for client-credentials Graph auth | **yes** (secret) | All four incl. hostname required or `AuthError` at construction (`graph_backend.py:143`) |
| `M365_SITE_HOSTNAME` | e.g. `vaklin.sharepoint.com` | n | required |
| `M365_ROUTES` | JSON: `{"First Segment": ["/sites/x", "Library Name"], …}` | n | The whole multi-site map. Bad JSON / wrong shape → `BackendError` at first call |
| `M365_STRIP_PREFIX` | Legacy prefix stripped before routing (per ops notes the fleet sets `/1 Onedrive Transfer`; actual value lives in each service's Railway env, not in this repo) | n | Idempotent — non-matching paths untouched |
| `M365_SITE_PATH` / `M365_DEFAULT_LIBRARY` | Catch-all default site/library for unrouted segments | n | Optional. Without it, every path MUST match a route |
| `M365_HEALTHCHECK_PATH` | Path probed by `healthcheck()` | n | Optional; falls back to first route segment |
| `STORAGE_MAX_RETRIES` | Transient-retry count (default 5) | n | `retry.py` |
| `SENTRY_DSN` / `LOG_LEVEL` | Observability | DSN semi | Sentry init is automatic in `get_adapter()` if DSN set + sdk installed |
| `DROPBOX_APP_KEY/_APP_SECRET/_REFRESH_TOKEN` | Legacy backend only (`dropbox_backend.py` requires all 3) | yes | Per ops notes deleted fleet-wide 2026-06-05 (not verifiable from this repo) |
| `INTEGRATION_TESTS` | **Dead switch** — documented in README/env.example but read by no code or test | n | Live end-to-end check is `scripts/smoke_test_m365.py`, not an env flag |

## Consumers and pins (verified in `~/dev/*/requirements.txt`, 2026-06-11)

Consumers install a **frozen commit**: `tropi-storage-adapter @ git+https://github.com/stoynovski-a11y/tropi-storage-adapter.git@<sha>`. Upgrading a service = edit its requirements.txt pin + redeploy that service. Pushing to this repo deploys **nothing** by itself.

| Pin | Services | Missing vs HEAD (`943c6c8`) |
|---|---|---|
| `943c6c8` (HEAD) | svedenie-generator | — |
| `d5aa66d` | railway-bank-parser, railway-invoice-renamer, railway-keyaccounts, railway-metro, railway-warehouse-receipts, railway-zajavki-aggregator, sales-autofill, metro-order-parser, warehouse-transfer, kasa-automation, microinvest-export | only the copy-poll backoff perf tweak — fine |
| `85a85e5` | invoice-generator | move-with-rename fix (`dbdf758`), transient-404 move retry (`7635760`), **m365 default** (`b17946e`), folder cache — must keep `STORAGE_BACKEND=m365` set explicitly |
| `94ccf1c` | digital-delivery-api (parked per owner notes — not in use) | predates list() round-trip fix (`85a85e5`), move-with-rename fix (`dbdf758`), m365 default — **re-pin before reviving** |

Also: `tropi-excel-adapter` lists it as an *optional* dependency (uses `resolve_route` for logical-path resolution).

## Deploy / release

There is no deploy. Workflow: branch → PR → merge to `main` → note the new sha → bump pins in consuming services one at a time → redeploy each (Railway auto-deploy or `railway up` per that service's runbook). Tests: `pytest` locally (unit tests fully mock HTTP via `httpx.MockTransport`). For a live end-to-end check, fill `.env` and run `scripts/smoke_test_m365.py` — the `INTEGRATION_TESTS` flag mentioned in the README is not wired to anything. No GitHub Actions in this repo.

## Failure modes & debugging

| Symptom | Likely cause | Where to look / fix |
|---|---|---|
| `AuthError: Graph backend requires M365_TENANT_ID…` at startup | Missing env var on the consuming service | Railway → service → Variables. All 4 of tenant/client/secret/hostname needed |
| `AuthError: MSAL token acquisition failed` or Graph 401/403 | Expired/rotated client secret, or app lacks site permission | Entra app registration; check secret expiry. 403 on mailboxes = RAOP policy (see MEMORY.md), not this lib |
| `BackendError: No M365 route for top-level segment '…'` | Path's first segment not in `M365_ROUTES` and no default site | Fix the service's `M365_ROUTES` JSON or the path; remember `M365_STRIP_PREFIX` must match the legacy prefix exactly |
| `BackendError: M365_ROUTES is not valid JSON` | Quoting mangled when setting the Railway var | Re-set as single-line JSON, single-quoted in shell |
| `NotFoundError` right after a write/move elsewhere | SharePoint path index is eventually consistent | `move()` already retries lookups ~3.5 s (`_get_item_settling`); `read()/list()` deliberately fail fast — caller should retry |
| `ThrottledError` storms / slow ops | Graph 429/5xx | Auto-retried up to `STORAGE_MAX_RETRIES` (5) with backoff honoring `Retry-After`. Persistent → check Railway logs of the consumer for the JSON `tropi_storage` log lines (`"result": "error"`) |
| `ConflictError` on `write_with_etag` | File changed since etag read | By design — caller must re-read and retry |
| Dropbox-flavored errors anywhere | A service fell back to `STORAGE_BACKEND=dropbox` | Dropbox is dead; set the var to `m365` (or unset it on pins ≥ `b17946e`) |

Cost note: this library calls no paid AI APIs. The money guards (Gemini attempt caps, Redis ledgers) live in the consuming services — but adapter behavior matters to them: a swallowed/wrapped error here once caused the bank-parser Gemini loop (see Gotchas).

Observability: every operation emits one JSON line on the `tropi_storage` logger (success at INFO, errors at ERROR, `NotFoundError` at WARNING so probe-misses don't spam Sentry). Read them in the consuming service's Railway logs.

## Gotchas (all verified in current code)

1. **Public repo.** Routes (`DEFAULT_ROUTES`) and fixtures are sanitized. Real site names live only in env vars. `.env` and `.env.bak*` are gitignored (a tenant-secret backup nearly leaked once — commit `a84cafa`). Never add real paths to code, tests, or docs here.
2. **`list()` returns logical paths, not drive-relative ones** (fix `85a85e5`). Under multi-site routing the route segment is consumed; the backend rebuilds `path` from the caller's input so list output round-trips into `read()/move()/copy()`. Services pinned before `85a85e5` get drive-relative paths back → 404 "Requested site could not be found".
3. **`move()` addresses the source by item id** (fix `dbdf758`). Graph rejects a rename-while-moving when the item is addressed by path — broke autorename-on-conflict (`foo (1).pdf`).
4. **Cross-library `move()` raises `BackendError`** ("use copy + delete instead") — a single Graph PATCH can't cross drives (`graph_backend.py:498`).
5. **Error wrapping crosses the adapter boundary.** A Dropbox/Graph conflict can surface as a generic `BackendError`, not `ConflictError` — bank-parser's type-only `_is_conflict()` check missed it and caused the 2026-05 Gemini cost loop. Consumers should match the message/payload too, not just the type.
6. **Folder cache** (`d5aa66d`): folders confirmed/created are cached per process (`_known_folders`); invalidated by this process's `delete()/move()` but NOT by out-of-band deletion. Don't delete-and-expect-recreate from another process within the same consumer's lifetime.
7. **`README.md` is stale** — it says `STORAGE_BACKEND` defaults to `dropbox`; code defaults to `m365` since `b17946e`. Trust the code.
8. **The stale item-id cache bug (`321594d`) is in tropi-EXCEL-adapter, not this repo** — don't hunt for it here.
