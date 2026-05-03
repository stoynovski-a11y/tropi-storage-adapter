"""Exception hierarchy for the storage adapter.

Backend-specific errors are translated to these so callers can write
backend-agnostic error handling.
"""


class StorageError(Exception):
    """Base class for all storage adapter errors."""


class NotFoundError(StorageError):
    """Path does not exist on the backend."""


class ConflictError(StorageError):
    """Conditional write failed — the file changed since the etag was read."""


class AuthError(StorageError):
    """Authentication or authorization failed."""


class ThrottledError(StorageError):
    """Backend returned a rate-limit response (HTTP 429 or Dropbox 'too_many_requests')."""

    def __init__(self, message: str, retry_after: float | None = None):
        super().__init__(message)
        self.retry_after = retry_after


class LockError(StorageError):
    """Lock-related failure (already locked by someone else, lock not held, etc.)."""


class BackendError(StorageError):
    """Generic backend failure that doesn't fit a more specific subclass."""

    def __init__(self, message: str, status_code: int | None = None):
        super().__init__(message)
        self.status_code = status_code
