"""Tests for GraphBackend with httpx and msal fully mocked.

We never make real Graph or Azure AD calls. The backend takes an
`http_client=` parameter so we inject an httpx.MockTransport-backed client.
The MSAL token call is patched to return a static fake token.
"""
from __future__ import annotations

from unittest.mock import patch

import httpx
import pytest

from tropi_storage import (
    AuthError,
    ConflictError,
    NotFoundError,
    ThrottledError,
)
from tropi_storage.backends.graph_backend import GraphBackend

# Constants matching what the fake routes return.
FAKE_SITE_ID = "fake-site-id"
FAKE_DRIVE_ID = "fake-drive-id"


def make_handler(routes: dict | None = None, default_status: int = 404):
    """Build an httpx.MockTransport handler from a {(method, url_substring): response} dict.

    Each value is either an httpx.Response or a callable(request) -> httpx.Response.
    Order matters: first match wins (so put more-specific keys first).
    """
    routes = routes or {}

    def handler(request: httpx.Request) -> httpx.Response:
        url = str(request.url)
        for (method, needle), out in routes.items():
            if request.method == method and needle in url:
                if callable(out):
                    return out(request)
                return out
        return httpx.Response(default_status, json={"error": {"message": f"no route for {request.method} {url}"}})

    return handler


@pytest.fixture
def env(monkeypatch):
    monkeypatch.setenv("M365_TENANT_ID", "fake-tenant")
    monkeypatch.setenv("M365_CLIENT_ID", "fake-client")
    monkeypatch.setenv("M365_CLIENT_SECRET", "fake-secret")
    monkeypatch.setenv("M365_SITE_HOSTNAME", "x.sharepoint.com")
    monkeypatch.setenv("M365_SITE_PATH", "/sites/X")


@pytest.fixture
def fake_token():
    """Patch MSAL so token acquisition returns a static fake token."""
    with patch("tropi_storage.backends.graph_backend.msal.ConfidentialClientApplication") as MockApp:
        instance = MockApp.return_value
        instance.acquire_token_for_client.return_value = {
            "access_token": "fake-token",
            "expires_in": 3600,
        }
        yield instance


def make_backend(routes: dict, env, fake_token, **kwargs) -> GraphBackend:
    """Build a GraphBackend wired to a MockTransport that returns `routes`."""
    transport = httpx.MockTransport(make_handler(routes))
    client = httpx.Client(transport=transport)
    return GraphBackend(http_client=client, **kwargs)


# Common pre-baked routes for site/drive resolution.
def site_drive_routes():
    return {
        ("GET", "/sites/x.sharepoint.com:/sites/X"): httpx.Response(
            200, json={"id": FAKE_SITE_ID, "name": "X"}),
        ("GET", f"/sites/{FAKE_SITE_ID}/drive"): httpx.Response(
            200, json={"id": FAKE_DRIVE_ID, "name": "Documents"}),
    }


class TestInit:
    def test_missing_creds_raises(self, monkeypatch):
        for var in ["M365_TENANT_ID", "M365_CLIENT_ID", "M365_CLIENT_SECRET",
                    "M365_SITE_HOSTNAME", "M365_SITE_PATH"]:
            monkeypatch.delenv(var, raising=False)
        with pytest.raises(AuthError):
            GraphBackend()

    def test_token_failure_raises(self, env, monkeypatch):
        with patch("tropi_storage.backends.graph_backend.msal.ConfidentialClientApplication") as MockApp:
            MockApp.return_value.acquire_token_for_client.return_value = {
                "error": "invalid_client", "error_description": "bad secret"}
            transport = httpx.MockTransport(make_handler({}))
            backend = GraphBackend(http_client=httpx.Client(transport=transport))
            with pytest.raises(AuthError, match="bad secret"):
                backend._get_token()


class TestSiteResolution:
    def test_caches_site_and_drive(self, env, fake_token):
        calls = {"count": 0}

        def site_handler(req):
            calls["count"] += 1
            return httpx.Response(200, json={"id": FAKE_SITE_ID})

        routes = {
            ("GET", "/sites/x.sharepoint.com:"): site_handler,
            ("GET", f"/sites/{FAKE_SITE_ID}/drive"): httpx.Response(
                200, json={"id": FAKE_DRIVE_ID}),
        }
        backend = make_backend(routes, env, fake_token)
        backend._resolve_site_and_drive()
        backend._resolve_site_and_drive()  # second call must be cached
        assert calls["count"] == 1


class TestRead:
    def test_returns_bytes(self, env, fake_token):
        routes = {
            **site_drive_routes(),
            ("GET", "/root:/foo.xlsx:/content"): httpx.Response(200, content=b"hello-bytes"),
        }
        backend = make_backend(routes, env, fake_token)
        assert backend.read("/foo.xlsx") == b"hello-bytes"

    def test_404_raises_not_found(self, env, fake_token):
        routes = {
            **site_drive_routes(),
            ("GET", "/content"): httpx.Response(
                404, json={"error": {"message": "Not found"}}),
        }
        backend = make_backend(routes, env, fake_token)
        with pytest.raises(NotFoundError):
            backend.read("/missing.xlsx")


class TestWrite:
    def test_small_upload(self, env, fake_token):
        captured = {}

        def put_handler(req):
            captured["body"] = req.content
            captured["url"] = str(req.url)
            return httpx.Response(200, json={
                "id": "item1", "name": "foo.xlsx", "size": 5,
                "cTag": "ctag-v1", "file": {"hashes": {}},
                "lastModifiedDateTime": "2026-05-03T12:00:00Z",
                "parentReference": {"path": "/drive/root:"},
            })

        routes = {
            **site_drive_routes(),
            ("PUT", "/root:/foo.xlsx:/content"): put_handler,
        }
        backend = make_backend(routes, env, fake_token)
        meta = backend.write("/foo.xlsx", b"hello")
        assert captured["body"] == b"hello"
        assert "conflictBehavior=replace" in captured["url"]
        assert meta["etag"] == "ctag-v1"
        assert meta["type"] == "file"

    def test_no_overwrite_uses_fail_behavior(self, env, fake_token):
        captured = {}

        def put_handler(req):
            captured["url"] = str(req.url)
            return httpx.Response(200, json={
                "id": "i", "name": "x", "size": 1, "cTag": "c",
                "lastModifiedDateTime": "2026-05-03T12:00:00Z",
                "file": {"hashes": {}},
                "parentReference": {"path": "/drive/root:"},
            })

        routes = {**site_drive_routes(), ("PUT", "/content"): put_handler}
        backend = make_backend(routes, env, fake_token)
        backend.write("/x.txt", b"x", overwrite=False)
        assert "conflictBehavior=fail" in captured["url"]


class TestList:
    def test_single_page(self, env, fake_token):
        routes = {
            **site_drive_routes(),
            ("GET", "/root:/folder:/children"): httpx.Response(200, json={
                "value": [
                    {"id": "1", "name": "a.xlsx", "size": 10, "cTag": "c1",
                     "lastModifiedDateTime": "2026-05-03T12:00:00Z",
                     "file": {"hashes": {}},
                     "parentReference": {"path": "/drive/root:/folder"}},
                    {"id": "2", "name": "sub", "folder": {},
                     "lastModifiedDateTime": "2026-05-03T12:00:00Z",
                     "parentReference": {"path": "/drive/root:/folder"}},
                ]
            }),
        }
        backend = make_backend(routes, env, fake_token)
        items = backend.list("/folder")
        assert len(items) == 2
        assert items[0]["type"] == "file"
        assert items[1]["type"] == "folder"
        assert items[0]["etag"] == "c1"

    def test_paginates_via_nextlink(self, env, fake_token):
        routes = {
            ("GET", "/sites/x.sharepoint.com:/sites/X"): httpx.Response(
                200, json={"id": FAKE_SITE_ID}),
            ("GET", f"/sites/{FAKE_SITE_ID}/drive"): httpx.Response(
                200, json={"id": FAKE_DRIVE_ID}),
            ("GET", "skiptoken=AAA"): httpx.Response(200, json={
                "value": [{"id": "2", "name": "b", "size": 1, "cTag": "c",
                           "lastModifiedDateTime": "2026-05-03T12:00:00Z",
                           "file": {"hashes": {}},
                           "parentReference": {"path": "/drive/root:/x"}}],
            }),
            ("GET", "/root:/x:/children"): httpx.Response(200, json={
                "value": [{"id": "1", "name": "a", "size": 1, "cTag": "c",
                           "lastModifiedDateTime": "2026-05-03T12:00:00Z",
                           "file": {"hashes": {}},
                           "parentReference": {"path": "/drive/root:/x"}}],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/drives/x/items/y/children?skiptoken=AAA",
            }),
        }
        backend = make_backend(routes, env, fake_token)
        items = backend.list("/x")
        assert [i["name"] for i in items] == ["a", "b"]

    def test_multisite_path_round_trips(self, env, fake_token, monkeypatch):
        """Listed paths under a routed site must include the route segment so
        they round-trip back through read()/move(). Regression: drive-relative
        paths (without the segment) fell through to the default site and 404'd.
        """
        monkeypatch.setenv("M365_ROUTES", '{"TopA": ["/sites/alpha", "LibA"]}')
        ALPHA_SITE_ID, ALPHA_DRIVE_ID = "alpha-site-id", "alpha-drive-id"
        routes = {
            ("GET", "/sites/x.sharepoint.com:/sites/alpha"): httpx.Response(
                200, json={"id": ALPHA_SITE_ID, "name": "alpha"}),
            ("GET", f"/sites/{ALPHA_SITE_ID}/drives"): httpx.Response(200, json={
                "value": [{"id": ALPHA_DRIVE_ID, "name": "LibA"}]}),
            ("GET", f"/drives/{ALPHA_DRIVE_ID}/root:/sub:/children"): httpx.Response(
                200, json={"value": [
                    {"id": "1", "name": "file.xlsx", "size": 10, "cTag": "c1",
                     "lastModifiedDateTime": "2026-05-03T12:00:00Z",
                     "file": {"hashes": {}},
                     # Graph reports the *drive-relative* parent (no route segment).
                     "parentReference": {"path": "/drive/root:/sub"}},
                ]}),
        }
        backend = make_backend(routes, env, fake_token)
        items = backend.list("/TopA/sub")
        # Logical path keeps the "TopA" route segment, not the bare "/sub/...".
        assert items[0]["path"] == "/TopA/sub/file.xlsx"


class TestDelete:
    def test_basic(self, env, fake_token):
        called = {"deleted": False}

        def del_handler(req):
            called["deleted"] = True
            return httpx.Response(204)

        routes = {**site_drive_routes(), ("DELETE", "/root:/foo.xlsx"): del_handler}
        backend = make_backend(routes, env, fake_token)
        backend.delete("/foo.xlsx")
        assert called["deleted"]

    def test_idempotent_on_404(self, env, fake_token):
        routes = {
            **site_drive_routes(),
            ("DELETE", "/root:/missing"): httpx.Response(
                404, json={"error": {"message": "not found"}}),
        }
        backend = make_backend(routes, env, fake_token)
        backend.delete("/missing")  # no exception


class TestMove:
    def test_move_with_rename_uses_item_id(self, env, fake_token):
        """A move that also renames (autorename-on-conflict → 'src (1).pdf')
        must PATCH the source by item id, never by path — Graph rejects a
        name-change on a path-addressed item."""
        import json

        patched = {}

        def patch_handler(req):
            patched["url"] = str(req.url)
            patched["body"] = json.loads(req.content)
            return httpx.Response(200, json={"id": "srcid", "name": "src (1).pdf"})

        routes = {
            **site_drive_routes(),
            # destination parent folder exists (used by ensure-parent + _get_item)
            ("GET", "/root:/dst-folder"): httpx.Response(
                200, json={"id": "dstfolder", "name": "dst-folder", "folder": {}}),
            # source item → resolve its id
            ("GET", "/root:/src.pdf"): httpx.Response(
                200, json={"id": "srcid", "name": "src.pdf", "file": {}}),
            # the move PATCH must hit the item-id URL, not the path URL
            ("PATCH", f"/drives/{FAKE_DRIVE_ID}/items/srcid"): patch_handler,
        }
        backend = make_backend(routes, env, fake_token)
        backend.move("/src.pdf", "/dst-folder/src (1).pdf")

        assert "/items/srcid" in patched["url"]
        assert "/root:/src.pdf" not in patched["url"]  # NOT path-addressed
        assert patched["body"]["name"] == "src (1).pdf"
        assert patched["body"]["parentReference"]["id"] == "dstfolder"

    def test_move_retries_transient_404_on_source(self, env, fake_token, monkeypatch):
        """SharePoint can 404 the source path right after the file lands in the
        folder (index lag). move() should retry the lookup and succeed once the
        path settles, instead of failing the archive."""
        monkeypatch.setattr(
            "tropi_storage.backends.graph_backend.time.sleep", lambda *_: None)
        src_calls = {"n": 0}

        def src_handler(req):
            src_calls["n"] += 1
            if src_calls["n"] == 1:  # first lookup 404s, then the path settles
                return httpx.Response(404, json={"error": {"message": "not found"}})
            return httpx.Response(200, json={"id": "srcid", "name": "src.pdf", "file": {}})

        moved = {}

        def patch_handler(req):
            moved["ok"] = True
            return httpx.Response(200, json={"id": "srcid", "name": "src.pdf"})

        routes = {
            **site_drive_routes(),
            ("GET", "/root:/dst-folder"): httpx.Response(
                200, json={"id": "dstfolder", "name": "dst-folder", "folder": {}}),
            ("GET", "/root:/src.pdf"): src_handler,
            ("PATCH", f"/drives/{FAKE_DRIVE_ID}/items/srcid"): patch_handler,
        }
        backend = make_backend(routes, env, fake_token)
        backend.move("/src.pdf", "/dst-folder/src.pdf")

        assert src_calls["n"] == 2  # retried once
        assert moved.get("ok") is True

    def test_move_raises_when_source_truly_missing(self, env, fake_token, monkeypatch):
        """A persistent 404 (genuinely missing source) still fails — after the
        bounded settle retries, not forever."""
        monkeypatch.setattr(
            "tropi_storage.backends.graph_backend.time.sleep", lambda *_: None)
        routes = {
            **site_drive_routes(),
            ("GET", "/root:/dst-folder"): httpx.Response(
                200, json={"id": "dstfolder", "name": "dst-folder", "folder": {}}),
            ("GET", "/root:/gone.pdf"): httpx.Response(
                404, json={"error": {"message": "not found"}}),
        }
        backend = make_backend(routes, env, fake_token)
        with pytest.raises(NotFoundError):
            backend.move("/gone.pdf", "/dst-folder/gone.pdf")


class TestEnsureFolder:
    def test_creates_when_missing(self, env, fake_token):
        creates = []

        def get_handler(req):
            return httpx.Response(404, json={"error": {"message": "not found"}})

        def post_handler(req):
            import json
            creates.append(json.loads(req.content))
            return httpx.Response(201, json={"id": "new", "folder": {}, "name": "newfolder"})

        routes = {
            **site_drive_routes(),
            ("GET", "/root:/newfolder"): get_handler,
            ("POST", "/root/children"): post_handler,
        }
        backend = make_backend(routes, env, fake_token)
        backend.ensure_folder("/newfolder")
        assert creates and creates[0]["name"] == "newfolder"
        assert creates[0]["folder"] == {}

    def test_root_is_noop(self, env, fake_token):
        backend = make_backend(site_drive_routes(), env, fake_token)
        backend.ensure_folder("/")  # no calls beyond what's already cached

    def test_skips_existing_folder(self, env, fake_token):
        post_called = {"n": 0}

        def post_handler(req):
            post_called["n"] += 1
            return httpx.Response(201, json={"id": "x", "folder": {}, "name": "x"})

        routes = {
            **site_drive_routes(),
            ("GET", "/root:/existing"): httpx.Response(200, json={
                "id": "ex", "name": "existing", "folder": {}}),
            ("POST", "/children"): post_handler,
        }
        backend = make_backend(routes, env, fake_token)
        backend.ensure_folder("/existing")
        assert post_called["n"] == 0


class TestGetMetadata:
    def test_existing(self, env, fake_token):
        routes = {
            **site_drive_routes(),
            ("GET", "/root:/foo.xlsx"): httpx.Response(200, json={
                "id": "1", "name": "foo.xlsx", "size": 99, "cTag": "ctag1",
                "lastModifiedDateTime": "2026-05-03T12:00:00Z",
                "file": {"hashes": {"quickXorHash": "QXH=="}},
                "parentReference": {"path": "/drive/root:"},
            }),
        }
        backend = make_backend(routes, env, fake_token)
        meta = backend.get_metadata("/foo.xlsx")
        assert meta["exists"] is True
        assert meta["etag"] == "ctag1"
        assert meta["content_hash"] == "QXH=="
        assert meta["size"] == 99

    def test_missing_returns_exists_false(self, env, fake_token):
        routes = {
            **site_drive_routes(),
            ("GET", "/root:/gone"): httpx.Response(404, json={"error": {"message": "not found"}}),
        }
        backend = make_backend(routes, env, fake_token)
        meta = backend.get_metadata("/gone")
        assert meta["exists"] is False
        assert meta["name"] == "gone"


class TestCheckoutCheckin:
    def test_checkout_calls_native_endpoint(self, env, fake_token):
        called = {"n": 0}

        def post_handler(req):
            called["n"] += 1
            return httpx.Response(204)

        routes = {
            **site_drive_routes(),
            ("GET", "/root:/foo.xlsx"): httpx.Response(200, json={
                "id": "fid", "name": "foo.xlsx", "size": 1, "cTag": "c",
                "lastModifiedDateTime": "2026-05-03T12:00:00Z",
                "file": {"hashes": {}},
                "parentReference": {"path": "/drive/root:"},
            }),
            ("POST", "/items/fid/checkout"): post_handler,
        }
        backend = make_backend(routes, env, fake_token)
        backend.checkout("/foo.xlsx")
        assert called["n"] == 1

    def test_checkin_calls_native_endpoint(self, env, fake_token):
        called = {"n": 0}

        def post_handler(req):
            called["n"] += 1
            return httpx.Response(204)

        routes = {
            **site_drive_routes(),
            ("GET", "/root:/foo.xlsx"): httpx.Response(200, json={
                "id": "fid", "name": "foo.xlsx", "size": 1, "cTag": "c",
                "lastModifiedDateTime": "2026-05-03T12:00:00Z",
                "file": {"hashes": {}},
                "parentReference": {"path": "/drive/root:"},
            }),
            ("POST", "/items/fid/checkin"): post_handler,
        }
        backend = make_backend(routes, env, fake_token)
        backend.checkin("/foo.xlsx")
        assert called["n"] == 1


class TestWriteWithEtag:
    def test_sends_if_match_header(self, env, fake_token):
        captured = {}

        def put_handler(req):
            captured["if_match"] = req.headers.get("If-Match")
            return httpx.Response(200, json={
                "id": "i", "name": "foo.xlsx", "size": 5, "cTag": "new",
                "lastModifiedDateTime": "2026-05-03T12:00:00Z",
                "file": {"hashes": {}},
                "parentReference": {"path": "/drive/root:"},
            })

        routes = {**site_drive_routes(), ("PUT", "/root:/foo.xlsx:/content"): put_handler}
        backend = make_backend(routes, env, fake_token)
        backend.write_with_etag("/foo.xlsx", b"x", etag="my-etag")
        assert captured["if_match"] == "my-etag"

    def test_412_translates_to_conflict(self, env, fake_token):
        routes = {
            **site_drive_routes(),
            ("PUT", "/root:/foo.xlsx:/content"): httpx.Response(
                412, json={"error": {"message": "precondition failed"}}),
        }
        backend = make_backend(routes, env, fake_token)
        with pytest.raises(ConflictError):
            backend.write_with_etag("/foo.xlsx", b"x", etag="stale")


class TestRetry:
    def test_throttled_then_succeeds(self, env, fake_token, monkeypatch):
        monkeypatch.setattr("tropi_storage.retry.time.sleep", lambda s: None)
        attempts = {"n": 0}

        def get_handler(req):
            attempts["n"] += 1
            if attempts["n"] < 3:
                return httpx.Response(429, headers={"Retry-After": "1"},
                                       json={"error": {"message": "throttled"}})
            return httpx.Response(200, content=b"finally")

        routes = {**site_drive_routes(), ("GET", "/root:/foo:/content"): get_handler}
        backend = make_backend(routes, env, fake_token)
        assert backend.read("/foo") == b"finally"
        assert attempts["n"] == 3

    def test_gives_up_after_max(self, env, fake_token, monkeypatch):
        monkeypatch.setenv("STORAGE_MAX_RETRIES", "2")
        monkeypatch.setattr("tropi_storage.retry.time.sleep", lambda s: None)
        routes = {
            **site_drive_routes(),
            ("GET", "/root:/foo:/content"): httpx.Response(429, json={"error": {"message": "x"}}),
        }
        backend = make_backend(routes, env, fake_token)
        with pytest.raises(ThrottledError):
            backend.read("/foo")


class TestCopyPollDelay:
    def test_progressive_delay_sequence(self, env, fake_token, monkeypatch):
        """_await_async_op must use a progressive delay: start at _COPY_POLL_START,
        double each iteration, cap at _COPY_POLL_CAP (0.25, 0.5, 1.0, 2.0, 2.0, …).
        """
        from tropi_storage.backends.graph_backend import (
            _COPY_POLL_START,
            _COPY_POLL_CAP,
        )

        sleep_calls = []
        monkeypatch.setattr(
            "tropi_storage.backends.graph_backend.time.sleep",
            lambda s: sleep_calls.append(s),
        )

        poll_count = {"n": 0}

        def monitor_handler(req):
            poll_count["n"] += 1
            # Respond "inProgress" for the first 5 polls, then "completed".
            if poll_count["n"] < 6:
                return httpx.Response(200, json={"status": "inProgress"})
            return httpx.Response(200, json={"status": "completed"})

        # Wire the monitor URL directly — copy() would normally hand it off.
        routes = {
            **site_drive_routes(),
            ("GET", "monitor-op-123"): monitor_handler,
        }
        backend = make_backend(routes, env, fake_token)
        backend._await_async_op("https://graph.microsoft.com/v1.0/monitor-op-123")

        # Six polls total; the last one returns "completed" → no trailing sleep.
        assert poll_count["n"] == 6
        assert len(sleep_calls) == 5  # five "still in progress" polls each slept

        # Progressive: 0.25, 0.5, 1.0, 2.0, 2.0 (capped)
        assert sleep_calls[0] == _COPY_POLL_START
        assert sleep_calls[1] == min(_COPY_POLL_START * 2, _COPY_POLL_CAP)
        assert sleep_calls[2] == min(_COPY_POLL_START * 4, _COPY_POLL_CAP)
        assert sleep_calls[3] == _COPY_POLL_CAP
        assert sleep_calls[4] == _COPY_POLL_CAP  # stays capped

    def test_copy_poll_immediate_complete(self, env, fake_token, monkeypatch):
        """When Graph returns 'completed' on the first poll, no sleep at all."""
        sleep_calls = []
        monkeypatch.setattr(
            "tropi_storage.backends.graph_backend.time.sleep",
            lambda s: sleep_calls.append(s),
        )

        routes = {
            **site_drive_routes(),
            ("GET", "monitor-fast"): httpx.Response(200, json={"status": "completed"}),
        }
        backend = make_backend(routes, env, fake_token)
        backend._await_async_op("https://graph.microsoft.com/v1.0/monitor-fast")

        assert sleep_calls == []

    def test_copy_poll_303_no_sleep(self, env, fake_token, monkeypatch):
        """A 303 redirect response is treated as immediate completion."""
        sleep_calls = []
        monkeypatch.setattr(
            "tropi_storage.backends.graph_backend.time.sleep",
            lambda s: sleep_calls.append(s),
        )

        routes = {
            **site_drive_routes(),
            ("GET", "monitor-303"): httpx.Response(303),
        }
        backend = make_backend(routes, env, fake_token)
        backend._await_async_op("https://graph.microsoft.com/v1.0/monitor-303")

        assert sleep_calls == []


class TestKnownFolderCache:
    def test_second_ensure_skips_existence_gets(self, env, fake_token):
        """Once a deep path is walked, a second ensure of the same path makes
        zero existence GETs — the per-segment checks are served from cache."""
        get_calls = {"n": 0}

        def folder_get(req):
            get_calls["n"] += 1
            return httpx.Response(200, json={"id": "f", "name": "seg", "folder": {}})

        routes = {
            **site_drive_routes(),
            ("GET", "/root:/a/b/c"): folder_get,
            ("GET", "/root:/a/b"): folder_get,
            ("GET", "/root:/a"): folder_get,
        }
        backend = make_backend(routes, env, fake_token)
        backend.ensure_folder("/a/b/c")
        assert get_calls["n"] == 3  # walked a, b, c once
        backend.ensure_folder("/a/b/c")
        assert get_calls["n"] == 3  # fully cached — no new GETs

    def test_delete_invalidates_cache(self, env, fake_token):
        """Deleting a cached folder forces it (but not its surviving ancestors)
        to be re-verified on the next ensure."""
        get_calls = {"n": 0}

        def folder_get(req):
            get_calls["n"] += 1
            return httpx.Response(200, json={"id": "f", "name": "seg", "folder": {}})

        routes = {
            **site_drive_routes(),
            ("DELETE", "/root:/a/b"): httpx.Response(204),
            ("GET", "/root:/a/b"): folder_get,
            ("GET", "/root:/a"): folder_get,
        }
        backend = make_backend(routes, env, fake_token)
        backend.ensure_folder("/a/b")   # GETs a, b → cached
        assert get_calls["n"] == 2
        backend.delete("/a/b")          # forgets /a/b, keeps /a
        backend.ensure_folder("/a/b")   # /a cached, /b re-GET → +1
        assert get_calls["n"] == 3
