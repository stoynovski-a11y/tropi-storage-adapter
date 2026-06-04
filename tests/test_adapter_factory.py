"""Tests for the get_adapter() factory and the abstract interface."""
from __future__ import annotations

import pytest

from tropi_storage import StorageAdapter, get_adapter


class TestGetAdapter:
    def test_defaults_to_m365_when_unset(self, monkeypatch):
        # An unset STORAGE_BACKEND must default to m365, never the legacy
        # Dropbox backend (the old default caused silent Dropbox fallbacks).
        monkeypatch.delenv("STORAGE_BACKEND", raising=False)
        monkeypatch.setenv("M365_TENANT_ID", "fake-tenant")
        monkeypatch.setenv("M365_CLIENT_ID", "fake-client")
        monkeypatch.setenv("M365_CLIENT_SECRET", "fake-secret")
        monkeypatch.setenv("M365_SITE_HOSTNAME", "x.sharepoint.com")
        monkeypatch.setenv("M365_SITE_PATH", "/sites/X")
        adapter = get_adapter()
        assert isinstance(adapter, StorageAdapter)
        assert adapter.backend_name == "m365"

    def test_dropbox_explicit(self, monkeypatch):
        # Dropbox backend is retained for explicit rollback only.
        monkeypatch.setenv("STORAGE_BACKEND", "dropbox")
        monkeypatch.setenv("DROPBOX_APP_KEY", "fake-key")
        monkeypatch.setenv("DROPBOX_APP_SECRET", "fake-secret")
        monkeypatch.setenv("DROPBOX_REFRESH_TOKEN", "fake-refresh")
        adapter = get_adapter()
        assert isinstance(adapter, StorageAdapter)
        assert adapter.backend_name == "dropbox"

    def test_m365_selects_graph(self, monkeypatch):
        monkeypatch.setenv("STORAGE_BACKEND", "m365")
        monkeypatch.setenv("M365_TENANT_ID", "fake-tenant")
        monkeypatch.setenv("M365_CLIENT_ID", "fake-client")
        monkeypatch.setenv("M365_CLIENT_SECRET", "fake-secret")
        monkeypatch.setenv("M365_SITE_HOSTNAME", "x.sharepoint.com")
        monkeypatch.setenv("M365_SITE_PATH", "/sites/X")
        adapter = get_adapter()
        assert isinstance(adapter, StorageAdapter)
        assert adapter.backend_name == "m365"

    def test_unknown_backend_raises(self, monkeypatch):
        monkeypatch.setenv("STORAGE_BACKEND", "googledrive")
        with pytest.raises(ValueError, match="Unknown STORAGE_BACKEND"):
            get_adapter()

    def test_case_insensitive(self, monkeypatch):
        monkeypatch.setenv("STORAGE_BACKEND", "DROPBOX")
        monkeypatch.setenv("DROPBOX_APP_KEY", "x")
        monkeypatch.setenv("DROPBOX_APP_SECRET", "x")
        monkeypatch.setenv("DROPBOX_REFRESH_TOKEN", "x")
        adapter = get_adapter()
        assert adapter.backend_name == "dropbox"


class TestStorageAdapterInterface:
    def test_cannot_instantiate_abstract_directly(self):
        with pytest.raises(TypeError):
            StorageAdapter()  # type: ignore

    def test_required_methods_exist(self):
        # Sanity check the abstract interface still defines what the spec needs.
        required = {"read", "write", "list", "delete", "move", "copy",
                    "ensure_folder", "get_metadata", "checkout", "checkin",
                    "write_with_etag", "healthcheck"}
        assert required.issubset(set(dir(StorageAdapter)))
