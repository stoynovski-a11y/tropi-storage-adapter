"""Tests for DropboxBackend with the SDK fully mocked.

We never make real Dropbox API calls. The backend takes a `client=` parameter
for dependency injection — every test passes a MagicMock there.
"""
from __future__ import annotations

import datetime as dt
from unittest.mock import MagicMock

import pytest
from dropbox.exceptions import ApiError, RateLimitError
from dropbox.files import (
    DeletedMetadata,
    FileMetadata,
    FolderMetadata,
    ListFolderResult,
)

from tropi_storage import (
    AuthError,
    ConflictError,
    LockError,
    NotFoundError,
    ThrottledError,
)
from tropi_storage.backends.dropbox_backend import DropboxBackend


@pytest.fixture
def mock_dbx():
    return MagicMock()


@pytest.fixture
def backend(mock_dbx):
    return DropboxBackend(client=mock_dbx)


def _hash(label: str = "h") -> str:
    """Return a valid 64-char Dropbox-style content hash."""
    return (label * 64)[:64]


def make_file(path="/foo.xlsx", name="foo.xlsx", size=42, content_hash=None):
    if content_hash is None:
        content_hash = _hash("a")
    return FileMetadata(
        name=name,
        path_lower=path.lower(),
        path_display=path,
        id="id:1",
        size=size,
        content_hash=content_hash,
        server_modified=dt.datetime(2026, 5, 3, 12, 0, 0),
    )


def make_folder(path="/foo", name="foo"):
    return FolderMetadata(name=name, path_lower=path.lower(), path_display=path, id="id:f1")


class TestInit:
    def test_missing_creds_raises(self, monkeypatch):
        monkeypatch.delenv("DROPBOX_APP_KEY", raising=False)
        monkeypatch.delenv("DROPBOX_APP_SECRET", raising=False)
        monkeypatch.delenv("DROPBOX_REFRESH_TOKEN", raising=False)
        with pytest.raises(AuthError):
            DropboxBackend()

    def test_accepts_injected_client(self):
        m = MagicMock()
        b = DropboxBackend(client=m)
        assert b._dbx is m
        assert b.backend_name == "dropbox"


class TestRead:
    def test_returns_bytes(self, backend, mock_dbx):
        resp = MagicMock()
        resp.content = b"hello"
        mock_dbx.files_download.return_value = (make_file(), resp)
        assert backend.read("/foo.xlsx") == b"hello"
        mock_dbx.files_download.assert_called_once_with("/foo.xlsx")

    def test_normalizes_path(self, backend, mock_dbx):
        resp = MagicMock(content=b"x")
        mock_dbx.files_download.return_value = (make_file(), resp)
        backend.read("foo.xlsx")
        mock_dbx.files_download.assert_called_once_with("/foo.xlsx")

    def test_translates_not_found(self, backend, mock_dbx):
        err = ApiError("rid", MagicMock(is_path=lambda: True,
                                         get_path=lambda: MagicMock(is_not_found=lambda: True)),
                       "user", "request")
        mock_dbx.files_download.side_effect = err
        with pytest.raises(NotFoundError):
            backend.read("/missing.xlsx")


class TestWrite:
    def test_small_upload(self, backend, mock_dbx):
        mock_dbx.files_upload.return_value = make_file(path="/foo.xlsx", size=5)
        meta = backend.write("/foo.xlsx", b"hello")
        assert meta["size"] == 5
        assert meta["type"] == "file"
        mock_dbx.files_upload.assert_called_once()
        args, kwargs = mock_dbx.files_upload.call_args
        assert args[0] == b"hello"
        assert args[1] == "/foo.xlsx"

    def test_overwrite_false_uses_add_mode(self, backend, mock_dbx):
        mock_dbx.files_upload.return_value = make_file()
        backend.write("/foo.xlsx", b"hi", overwrite=False)
        kwargs = mock_dbx.files_upload.call_args.kwargs
        assert kwargs["mode"].is_add()


class TestList:
    def test_single_page(self, backend, mock_dbx):
        mock_dbx.files_list_folder.return_value = ListFolderResult(
            entries=[make_file(), make_folder()],
            cursor="c",
            has_more=False,
        )
        items = backend.list("/")
        assert len(items) == 2
        assert items[0]["type"] == "file"
        assert items[1]["type"] == "folder"
        # Root '/' must translate to '' for the Dropbox API.
        mock_dbx.files_list_folder.assert_called_once_with("", recursive=False)

    def test_paginates(self, backend, mock_dbx):
        page1 = ListFolderResult(entries=[make_file(name="a")], cursor="c1", has_more=True)
        page2 = ListFolderResult(entries=[make_file(name="b")], cursor="c2", has_more=False)
        mock_dbx.files_list_folder.return_value = page1
        mock_dbx.files_list_folder_continue.return_value = page2
        items = backend.list("/x")
        assert [i["name"] for i in items] == ["a", "b"]
        mock_dbx.files_list_folder_continue.assert_called_once_with("c1")

    def test_skips_deleted_marker_translation(self, backend, mock_dbx):
        deleted = DeletedMetadata(name="gone", path_lower="/gone", path_display="/gone")
        mock_dbx.files_list_folder.return_value = ListFolderResult(
            entries=[deleted], cursor="c", has_more=False)
        items = backend.list("/")
        assert items[0]["type"] == "deleted"


class TestDelete:
    def test_basic(self, backend, mock_dbx):
        mock_dbx.files_delete_v2.return_value = MagicMock()
        backend.delete("/foo.xlsx")
        mock_dbx.files_delete_v2.assert_called_once_with("/foo.xlsx")

    def test_idempotent_on_not_found(self, backend, mock_dbx):
        err = ApiError("rid", MagicMock(is_path_lookup=lambda: True,
                                         get_path_lookup=lambda: MagicMock(is_not_found=lambda: True)),
                       "user", "request")
        mock_dbx.files_delete_v2.side_effect = err
        backend.delete("/missing.xlsx")  # should not raise


class TestMoveCopy:
    def test_move(self, backend, mock_dbx):
        backend.move("/a", "/b")
        mock_dbx.files_move_v2.assert_called_once()
        args = mock_dbx.files_move_v2.call_args[0]
        assert args[0] == "/a" and args[1] == "/b"

    def test_copy(self, backend, mock_dbx):
        backend.copy("/a", "/b")
        mock_dbx.files_copy_v2.assert_called_once()


class TestEnsureFolder:
    def test_creates(self, backend, mock_dbx):
        backend.ensure_folder("/new/folder")
        mock_dbx.files_create_folder_v2.assert_called_once_with("/new/folder", autorename=False)

    def test_root_is_noop(self, backend, mock_dbx):
        backend.ensure_folder("/")
        mock_dbx.files_create_folder_v2.assert_not_called()

    def test_already_exists_is_ok(self, backend, mock_dbx):
        err = ApiError("rid", "path/conflict/folder/.", "user", "request")
        mock_dbx.files_create_folder_v2.side_effect = err
        backend.ensure_folder("/foo")  # no exception


class TestGetMetadata:
    def test_existing_file(self, backend, mock_dbx):
        h = _hash("a")
        mock_dbx.files_get_metadata.return_value = make_file(content_hash=h)
        meta = backend.get_metadata("/foo.xlsx")
        assert meta["exists"] is True
        assert meta["etag"] == h
        assert meta["content_hash"] == h

    def test_missing_returns_exists_false(self, backend, mock_dbx):
        err = ApiError("rid", MagicMock(is_path=lambda: True,
                                         get_path=lambda: MagicMock(is_not_found=lambda: True)),
                       "user", "request")
        mock_dbx.files_get_metadata.side_effect = err
        meta = backend.get_metadata("/missing.xlsx")
        assert meta["exists"] is False
        assert meta["path"] == "/missing.xlsx"


class TestCheckoutCheckin:
    def test_acquire_lock(self, backend, mock_dbx):
        # No existing lock file → write our UUID.
        not_found = ApiError("rid", MagicMock(is_path=lambda: True,
                                               get_path=lambda: MagicMock(is_not_found=lambda: True)),
                              "user", "request")
        mock_dbx.files_get_metadata.side_effect = not_found
        mock_dbx.files_upload.return_value = make_file(path="/foo.xlsx.lock")
        backend.checkout("/foo.xlsx")
        # Lock file written with our instance UUID.
        upload_args = mock_dbx.files_upload.call_args
        assert upload_args[0][1] == "/foo.xlsx.lock"
        assert upload_args[0][0].decode() == backend._instance_id

    def test_acquire_lock_held_by_other_raises(self, backend, mock_dbx):
        # get_metadata returns a real file, read returns someone else's UUID.
        mock_dbx.files_get_metadata.return_value = make_file(path="/foo.xlsx.lock")
        resp = MagicMock(content=b"someone-elses-uuid")
        mock_dbx.files_download.return_value = (make_file(path="/foo.xlsx.lock"), resp)
        with pytest.raises(LockError):
            backend.checkout("/foo.xlsx")

    def test_release_lock(self, backend, mock_dbx):
        # Lock exists, holds our UUID, gets deleted.
        mock_dbx.files_get_metadata.return_value = make_file(path="/foo.xlsx.lock")
        resp = MagicMock(content=backend._instance_id.encode())
        mock_dbx.files_download.return_value = (make_file(), resp)
        backend.checkin("/foo.xlsx")
        mock_dbx.files_delete_v2.assert_called_once_with("/foo.xlsx.lock")


class TestWriteWithEtag:
    def test_matching_etag_writes(self, backend, mock_dbx):
        h_old, h_new = _hash("a"), _hash("b")
        mock_dbx.files_get_metadata.return_value = make_file(content_hash=h_old)
        mock_dbx.files_upload.return_value = make_file(content_hash=h_new)
        meta = backend.write_with_etag("/foo.xlsx", b"x", etag=h_old)
        assert meta["content_hash"] == h_new

    def test_mismatched_etag_raises_conflict(self, backend, mock_dbx):
        mock_dbx.files_get_metadata.return_value = make_file(content_hash=_hash("a"))
        with pytest.raises(ConflictError):
            backend.write_with_etag("/foo.xlsx", b"x", etag=_hash("z"))

    def test_create_when_missing_with_empty_etag(self, backend, mock_dbx):
        not_found = ApiError("rid", MagicMock(is_path=lambda: True,
                                               get_path=lambda: MagicMock(is_not_found=lambda: True)),
                              "user", "request")
        mock_dbx.files_get_metadata.side_effect = not_found
        new_hash = _hash("n")
        mock_dbx.files_upload.return_value = make_file(content_hash=new_hash)
        meta = backend.write_with_etag("/new.xlsx", b"x", etag="")
        assert meta["content_hash"] == new_hash


class TestRetry:
    def test_retries_on_throttled(self, backend, mock_dbx, monkeypatch):
        # Fast retries.
        monkeypatch.setattr("tropi_storage.retry.time.sleep", lambda s: None)
        rate_err = RateLimitError("rid", MagicMock(retry_after=1), backoff=0.0)
        mock_dbx.files_download.side_effect = [
            rate_err,
            rate_err,
            (make_file(), MagicMock(content=b"ok")),
        ]
        assert backend.read("/foo.xlsx") == b"ok"
        assert mock_dbx.files_download.call_count == 3

    def test_gives_up_after_max(self, backend, mock_dbx, monkeypatch):
        monkeypatch.setenv("STORAGE_MAX_RETRIES", "2")
        monkeypatch.setattr("tropi_storage.retry.time.sleep", lambda s: None)
        rate_err = RateLimitError("rid", MagicMock(retry_after=1), backoff=0.0)
        mock_dbx.files_download.side_effect = rate_err
        with pytest.raises(ThrottledError):
            backend.read("/foo.xlsx")
        assert mock_dbx.files_download.call_count == 3  # 1 + 2 retries
