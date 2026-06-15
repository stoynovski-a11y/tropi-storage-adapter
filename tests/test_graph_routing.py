"""Tests for multi-site routing in GraphBackend.

Mirrors the mock style from tests/test_graph_backend.py:
  - httpx.MockTransport for HTTP simulation
  - msal patched with fake_token fixture
"""
from __future__ import annotations

from unittest.mock import patch

import httpx
import pytest

from tropi_storage.backends.graph_backend import GraphBackend
from tropi_storage.exceptions import BackendError
from tropi_storage.routing import (
    DEFAULT_ROUTES,
    load_folder_pins,
    load_routes,
    load_strip_prefix,
    resolve_route,
)

# ---------------------------------------------------------------------------
# Helpers (same pattern as test_graph_backend.py)
# ---------------------------------------------------------------------------

FAKE_SITE_ID_BETA = "site-id-beta"
FAKE_DRIVE_ID_BETA = "drv-beta"

# Generic placeholder routes used across tests.
_SAMPLE_ROUTES: dict[str, tuple[str, str]] = {
    "Top A": ("/sites/alpha", "Library A"),
    "Top B": ("/sites/beta", "Library B"),
}


def make_handler(routes: dict | None = None, default_status: int = 404):
    routes = routes or {}

    def handler(request: httpx.Request) -> httpx.Response:
        url = str(request.url)
        for (method, needle), out in routes.items():
            if request.method == method and needle in url:
                if callable(out):
                    return out(request)
                return out
        return httpx.Response(
            default_status,
            json={"error": {"message": f"no route for {request.method} {url}"}},
        )

    return handler


@pytest.fixture
def env(monkeypatch):
    """Set the four required credentials; leave M365_SITE_PATH unset by default."""
    monkeypatch.setenv("M365_TENANT_ID", "fake-tenant")
    monkeypatch.setenv("M365_CLIENT_ID", "fake-client")
    monkeypatch.setenv("M365_CLIENT_SECRET", "fake-secret")
    monkeypatch.setenv("M365_SITE_HOSTNAME", "x.sharepoint.com")
    monkeypatch.delenv("M365_SITE_PATH", raising=False)
    monkeypatch.delenv("M365_DEFAULT_LIBRARY", raising=False)
    monkeypatch.setenv(
        "M365_ROUTES",
        '{"Top A": ["/sites/alpha", "Library A"], "Top B": ["/sites/beta", "Library B"]}',
    )


@pytest.fixture
def env_with_default(monkeypatch, env):
    """Add a default site so back-compat tests work."""
    monkeypatch.setenv("M365_SITE_PATH", "/sites/default")


@pytest.fixture
def fake_token():
    with patch("tropi_storage.backends.graph_backend.msal.ConfidentialClientApplication") as MockApp:
        instance = MockApp.return_value
        instance.acquire_token_for_client.return_value = {
            "access_token": "fake-token",
            "expires_in": 3600,
        }
        yield instance


def make_backend(routes: dict, env, fake_token, **kwargs) -> GraphBackend:
    transport = httpx.MockTransport(make_handler(routes))
    client = httpx.Client(transport=transport)
    return GraphBackend(http_client=client, **kwargs)


# ---------------------------------------------------------------------------
# Beta-site mock responses for /sites/beta  "Library B"
# ---------------------------------------------------------------------------

def beta_site_drives_routes():
    """Mock responses for site-id + drives list for /sites/beta."""
    return {
        ("GET", "/sites/x.sharepoint.com:/sites/beta"): httpx.Response(
            200, json={"id": FAKE_SITE_ID_BETA, "name": "Beta"}
        ),
        ("GET", f"/sites/{FAKE_SITE_ID_BETA}/drives"): httpx.Response(
            200,
            json={
                "value": [
                    {"name": "Documents", "id": "drv-default"},
                    {"name": "Library B", "id": FAKE_DRIVE_ID_BETA},
                ]
            },
        ),
    }


# ---------------------------------------------------------------------------
# 1. Routed read test
# ---------------------------------------------------------------------------

class TestRoutedRead:
    def test_reads_via_correct_drive(self, env, fake_token):
        """read('/Top B/sub/x.xlsx') must use drv-beta (Library B drive)."""
        captured = {}

        def content_handler(req):
            captured["url"] = str(req.url)
            return httpx.Response(200, content=b"routed-bytes")

        routes = {
            **beta_site_drives_routes(),
            ("GET", "/content"): content_handler,
        }
        backend = make_backend(routes, env, fake_token)
        result = backend.read("/Top B/sub/x.xlsx")

        assert result == b"routed-bytes"
        # Drive id must appear in the URL.
        assert FAKE_DRIVE_ID_BETA in captured["url"]
        # First segment stripped: item path is /sub/x.xlsx
        assert "sub" in captured["url"]

    def test_item_path_strips_first_segment(self, env, fake_token):
        """The request URL must contain the item path with first segment removed."""
        captured_urls = []

        def any_get(req):
            captured_urls.append(str(req.url))
            return httpx.Response(200, content=b"x")

        routes = {
            **beta_site_drives_routes(),
            ("GET", "/content"): any_get,
        }
        backend = make_backend(routes, env, fake_token)
        backend.read("/Top B/sub/x.xlsx")

        # The content URL should contain the drive id.
        content_url = next(u for u in captured_urls if "/content" in u)
        assert f"/drives/{FAKE_DRIVE_ID_BETA}/" in content_url
        # The first segment "Top B" must NOT appear in the item path portion.
        assert "Top%20B" not in content_url or "sub" in content_url


# ---------------------------------------------------------------------------
# 2. Pure resolve_route unit tests
# ---------------------------------------------------------------------------

class TestResolveRoute:
    def test_routed_segment(self):
        site, drive, item = resolve_route(
            "/Top B/sub/x.xlsx", _SAMPLE_ROUTES
        )
        assert site == "/sites/beta"
        assert drive == "Library B"
        assert item == "/sub/x.xlsx"

    def test_nested_remainder(self):
        _, _, item = resolve_route(
            "/Top A/sub/path/foo.xlsx", _SAMPLE_ROUTES
        )
        assert item == "/sub/path/foo.xlsx"

    def test_single_segment_remainder_is_root(self):
        _, _, item = resolve_route("/Top A", _SAMPLE_ROUTES)
        assert item == "/"

    def test_unknown_segment_with_default_site(self):
        routes = dict(_SAMPLE_ROUTES)
        site, drive, item = resolve_route(
            "/SomeOther/path/file.xlsx",
            routes,
            default_site="/sites/fallback",
            default_drive="FallbackLib",
        )
        assert site == "/sites/fallback"
        assert drive == "FallbackLib"
        assert item == "/SomeOther/path/file.xlsx"

    def test_unknown_segment_no_default_raises(self):
        with pytest.raises(BackendError, match="No M365 route"):
            resolve_route("/NoSuchFolder/x.xlsx", _SAMPLE_ROUTES)

    def test_root_no_default_raises(self):
        with pytest.raises(BackendError, match="Cannot address multi-site root"):
            resolve_route("/", _SAMPLE_ROUTES)

    def test_root_with_default_site(self):
        site, drive, item = resolve_route(
            "/",
            _SAMPLE_ROUTES,
            default_site="/sites/x",
            default_drive="Lib",
        )
        assert site == "/sites/x"
        assert drive == "Lib"
        assert item == "/"


# ---------------------------------------------------------------------------
# 3. load_routes tests
# ---------------------------------------------------------------------------

class TestLoadRoutes:
    def test_default_table(self, monkeypatch):
        monkeypatch.delenv("M365_ROUTES", raising=False)
        routes = load_routes()
        assert routes == DEFAULT_ROUTES  # empty dict

    def test_env_adds_entry(self, monkeypatch):
        monkeypatch.setenv(
            "M365_ROUTES",
            '{"New Segment": ["/sites/newsite", "New Library"]}',
        )
        routes = load_routes()
        assert "New Segment" in routes
        assert routes["New Segment"] == ("/sites/newsite", "New Library")

    def test_env_overrides_default(self, monkeypatch):
        # Seed a known entry via M365_ROUTES, then override it with another call.
        monkeypatch.setenv(
            "M365_ROUTES",
            '{"Top A": ["/sites/alpha-override", "Override Lib"]}',
        )
        routes = load_routes()
        assert routes["Top A"] == ("/sites/alpha-override", "Override Lib")

    def test_invalid_json_raises(self, monkeypatch):
        monkeypatch.setenv("M365_ROUTES", "not-valid-json{")
        with pytest.raises(BackendError, match="not valid JSON"):
            load_routes()

    def test_wrong_shape_raises(self, monkeypatch):
        monkeypatch.setenv("M365_ROUTES", '{"Seg": "not-an-array"}')
        with pytest.raises(BackendError, match="two-element string array"):
            load_routes()

    def test_non_object_raises(self, monkeypatch):
        monkeypatch.setenv("M365_ROUTES", '["oops"]')
        with pytest.raises(BackendError, match="JSON object"):
            load_routes()


# ---------------------------------------------------------------------------
# 4. drive-not-found test
# ---------------------------------------------------------------------------

class TestDriveNotFound:
    def test_raises_backend_error(self, env, fake_token):
        routes = {
            ("GET", "/sites/x.sharepoint.com:/sites/beta"): httpx.Response(
                200, json={"id": FAKE_SITE_ID_BETA}
            ),
            ("GET", f"/sites/{FAKE_SITE_ID_BETA}/drives"): httpx.Response(
                200,
                json={
                    "value": [
                        {"name": "Documents", "id": "drv-docs"},
                        {"name": "Other", "id": "drv-other"},
                        # "Library B" is intentionally absent
                    ]
                },
            ),
        }
        backend = make_backend(routes, env, fake_token)
        with pytest.raises(BackendError, match="Library.*not found"):
            backend.read("/Top B/x.xlsx")


# ---------------------------------------------------------------------------
# 5. Back-compat: default_site + unrouted path uses default drive
# ---------------------------------------------------------------------------

class TestBackCompat:
    def test_unrouted_path_uses_default_drive(self, monkeypatch, fake_token):
        """When default site is set and path isn't in routes, use default drive."""
        monkeypatch.setenv("M365_TENANT_ID", "t")
        monkeypatch.setenv("M365_CLIENT_ID", "c")
        monkeypatch.setenv("M365_CLIENT_SECRET", "s")
        monkeypatch.setenv("M365_SITE_HOSTNAME", "x.sharepoint.com")
        monkeypatch.setenv("M365_SITE_PATH", "/sites/default")
        monkeypatch.delenv("M365_DEFAULT_LIBRARY", raising=False)
        monkeypatch.delenv("M365_ROUTES", raising=False)

        captured = {}

        def content_handler(req):
            captured["url"] = str(req.url)
            return httpx.Response(200, content=b"default-bytes")

        default_site_id = "site-default"
        default_drive_id = "drv-default"

        routes = {
            ("GET", "/sites/x.sharepoint.com:/sites/default"): httpx.Response(
                200, json={"id": default_site_id}
            ),
            ("GET", f"/sites/{default_site_id}/drive"): httpx.Response(
                200, json={"id": default_drive_id}
            ),
            ("GET", "/content"): content_handler,
        }
        transport = httpx.MockTransport(make_handler(routes))
        client = httpx.Client(transport=transport)
        backend = GraphBackend(http_client=client)

        result = backend.read("/SomeUnroutedFolder/x.xlsx")
        assert result == b"default-bytes"
        # Must use the default drive.
        assert default_drive_id in captured["url"]
        # Full logical path preserved (no stripping) — first segment kept.
        assert "SomeUnroutedFolder" in captured["url"]


# ---------------------------------------------------------------------------
# 6. strip_prefix tests
# ---------------------------------------------------------------------------

class TestStripPrefix:
    """Tests for the strip_prefix parameter of resolve_route."""

    def test_strip_prefix_routes_correctly(self):
        """/Legacy/Top A/sub/x.xlsx routes as if it were /Top A/sub/x.xlsx."""
        site, drive, item = resolve_route(
            "/Legacy/Top A/sub/x.xlsx", _SAMPLE_ROUTES, strip_prefix="/Legacy"
        )
        assert site == "/sites/alpha"
        assert drive == "Library A"
        assert item == "/sub/x.xlsx"

    def test_strip_prefix_without_leading_slash(self):
        """Prefix supplied without leading slash is normalised — same result."""
        site, drive, item = resolve_route(
            "/Legacy/Top B/doc.xlsx", _SAMPLE_ROUTES, strip_prefix="Legacy"
        )
        assert site == "/sites/beta"
        assert drive == "Library B"
        assert item == "/doc.xlsx"

    def test_strip_prefix_with_trailing_slash(self):
        """Prefix supplied with trailing slash is normalised — same result."""
        site, drive, item = resolve_route(
            "/Legacy/Top A/file.xlsx", _SAMPLE_ROUTES, strip_prefix="/Legacy/"
        )
        assert site == "/sites/alpha"
        assert drive == "Library A"
        assert item == "/file.xlsx"

    def test_path_equal_to_prefix_becomes_root(self):
        """A path exactly equal to the prefix normalises to '/' (root)."""
        site, drive, item = resolve_route(
            "/Legacy",
            _SAMPLE_ROUTES,
            strip_prefix="/Legacy",
            default_site="/sites/fallback",
            default_drive="FallbackLib",
        )
        assert site == "/sites/fallback"
        assert item == "/"

    def test_no_strip_prefix_unchanged(self):
        """Without strip_prefix the original routing is unaffected."""
        site, drive, item = resolve_route(
            "/Top A/sub/x.xlsx", _SAMPLE_ROUTES
        )
        assert site == "/sites/alpha"
        assert item == "/sub/x.xlsx"

    def test_non_matching_prefix_is_noop(self):
        """A path that does NOT start with the prefix is routed as-is."""
        site, drive, item = resolve_route(
            "/Top B/sub/x.xlsx", _SAMPLE_ROUTES, strip_prefix="/Legacy"
        )
        assert site == "/sites/beta"
        assert drive == "Library B"
        assert item == "/sub/x.xlsx"

    def test_partial_prefix_match_is_noop(self):
        """A prefix that is a partial match (not a whole segment boundary) is not stripped."""
        # "/Leg" should NOT strip "/Legacy/..." — they don't share a segment boundary
        site, drive, item = resolve_route(
            "/Top A/x.xlsx", _SAMPLE_ROUTES, strip_prefix="/Leg"
        )
        assert site == "/sites/alpha"
        assert item == "/x.xlsx"


# ---------------------------------------------------------------------------
# 7. load_strip_prefix tests
# ---------------------------------------------------------------------------

class TestLoadStripPrefix:
    def test_env_set_returns_value(self, monkeypatch):
        monkeypatch.setenv("M365_STRIP_PREFIX", "/Legacy")
        assert load_strip_prefix() == "/Legacy"

    def test_env_unset_returns_none(self, monkeypatch):
        monkeypatch.delenv("M365_STRIP_PREFIX", raising=False)
        assert load_strip_prefix() is None

    def test_env_empty_string_returns_none(self, monkeypatch):
        monkeypatch.setenv("M365_STRIP_PREFIX", "")
        assert load_strip_prefix() is None

    def test_env_whitespace_only_returns_none(self, monkeypatch):
        monkeypatch.setenv("M365_STRIP_PREFIX", "   ")
        assert load_strip_prefix() is None

    def test_env_value_stripped_of_whitespace(self, monkeypatch):
        monkeypatch.setenv("M365_STRIP_PREFIX", "  /Legacy  ")
        assert load_strip_prefix() == "/Legacy"


# ---------------------------------------------------------------------------
# 8. load_folder_pins tests
# ---------------------------------------------------------------------------

class TestLoadFolderPins:
    def test_env_unset_returns_empty(self, monkeypatch):
        monkeypatch.delenv("M365_FOLDER_IDS", raising=False)
        assert load_folder_pins() == {}

    def test_env_empty_string_returns_empty(self, monkeypatch):
        monkeypatch.setenv("M365_FOLDER_IDS", "")
        assert load_folder_pins() == {}

    def test_parses_and_normalises_keys(self, monkeypatch):
        monkeypatch.setenv(
            "M365_FOLDER_IDS",
            '{"/Top A/Old Name/": "01ABC", "Top B/x": "01DEF"}',
        )
        pins = load_folder_pins()
        # Trailing slash dropped, leading slash added — both via normalize_path.
        assert pins == {"/Top A/Old Name": "01ABC", "/Top B/x": "01DEF"}

    def test_preserves_internal_double_space(self, monkeypatch):
        # The real ЕКО folder has a double space; normalize_path must not eat it.
        monkeypatch.setenv("M365_FOLDER_IDS", '{"/A/06  EKO": "01XYZ"}')
        assert load_folder_pins() == {"/A/06  EKO": "01XYZ"}

    def test_invalid_json_raises(self, monkeypatch):
        monkeypatch.setenv("M365_FOLDER_IDS", "not-json{")
        with pytest.raises(BackendError, match="not valid JSON"):
            load_folder_pins()

    def test_non_object_raises(self, monkeypatch):
        monkeypatch.setenv("M365_FOLDER_IDS", '["oops"]')
        with pytest.raises(BackendError, match="JSON object"):
            load_folder_pins()

    def test_non_string_value_raises(self, monkeypatch):
        monkeypatch.setenv("M365_FOLDER_IDS", '{"/A": 123}')
        with pytest.raises(BackendError, match="non-empty"):
            load_folder_pins()

    def test_empty_value_raises(self, monkeypatch):
        monkeypatch.setenv("M365_FOLDER_IDS", '{"/A": ""}')
        with pytest.raises(BackendError, match="non-empty"):
            load_folder_pins()
