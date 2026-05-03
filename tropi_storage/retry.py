"""Exponential-backoff retry decorator for transient errors."""
from __future__ import annotations

import functools
import os
import random
import time
from typing import Callable, TypeVar

from .exceptions import ThrottledError
from .logging_config import get_logger

T = TypeVar("T")

DEFAULT_MAX_RETRIES = 5
BACKOFF_BASE_SECONDS = 1.0
BACKOFF_CAP_SECONDS = 30.0


def _max_retries() -> int:
    try:
        return int(os.getenv("STORAGE_MAX_RETRIES", str(DEFAULT_MAX_RETRIES)))
    except ValueError:
        return DEFAULT_MAX_RETRIES


def _compute_delay(attempt: int, server_hint: float | None = None) -> float:
    """Exponential backoff (1s, 2s, 4s, 8s, 16s) with jitter, honoring server hints."""
    if server_hint is not None and server_hint > 0:
        return min(server_hint, BACKOFF_CAP_SECONDS)
    delay = BACKOFF_BASE_SECONDS * (2 ** attempt)
    delay = min(delay, BACKOFF_CAP_SECONDS)
    # Small jitter to avoid thundering-herd.
    return delay + random.uniform(0, 0.25)


def retry_on_transient(
    *,
    transient_exceptions: tuple[type[BaseException], ...] = (ThrottledError,),
    sleep: Callable[[float], None] = time.sleep,
) -> Callable[[Callable[..., T]], Callable[..., T]]:
    """Retry a callable on transient errors with exponential backoff.

    `transient_exceptions` defines what to catch and retry. Backends extend this
    with their HTTP-layer transient errors (timeouts, connection resets, 5xx).
    """

    def decorator(fn: Callable[..., T]) -> Callable[..., T]:
        @functools.wraps(fn)
        def wrapper(*args, **kwargs) -> T:
            max_retries = _max_retries()
            last_exc: BaseException | None = None
            for attempt in range(max_retries + 1):
                try:
                    return fn(*args, **kwargs)
                except transient_exceptions as exc:
                    last_exc = exc
                    if attempt >= max_retries:
                        break
                    server_hint = getattr(exc, "retry_after", None)
                    delay = _compute_delay(attempt, server_hint)
                    get_logger().warning(
                        f"transient error on {fn.__name__} "
                        f"(attempt {attempt + 1}/{max_retries + 1}): "
                        f"{type(exc).__name__} — sleeping {delay:.1f}s",
                        extra={"operation": fn.__name__, "attempt": attempt + 1,
                               "error_type": type(exc).__name__},
                    )
                    sleep(delay)
            assert last_exc is not None
            raise last_exc

        return wrapper

    return decorator
