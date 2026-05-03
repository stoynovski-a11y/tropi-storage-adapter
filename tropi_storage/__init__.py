"""Tropi Storage Adapter — unified Dropbox / Microsoft Graph file storage."""
from .adapter import StorageAdapter, get_adapter
from .exceptions import (
    AuthError,
    BackendError,
    ConflictError,
    LockError,
    NotFoundError,
    StorageError,
    ThrottledError,
)
from .path_utils import expand_path, normalize_path, split_parent

__version__ = "0.1.0"

__all__ = [
    "StorageAdapter",
    "get_adapter",
    "expand_path",
    "normalize_path",
    "split_parent",
    "StorageError",
    "NotFoundError",
    "ConflictError",
    "AuthError",
    "ThrottledError",
    "LockError",
    "BackendError",
    "__version__",
]
