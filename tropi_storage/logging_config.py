"""Structured JSON logging for adapter operations.

Each operation emits one log line as JSON. Caller can attach a stdlib logging
handler to the `tropi_storage` logger to capture or forward these.
"""
from __future__ import annotations

import json
import logging
import os
import time
from contextlib import contextmanager
from typing import Any

LOGGER_NAME = "tropi_storage"

_SECRET_KEYS = {"authorization", "access_token", "refresh_token", "client_secret",
                "x-vercel-protection-bypass", "cookie", "set-cookie"}


def _redact(obj: Any) -> Any:
    """Strip secret-looking values from dicts before logging."""
    if isinstance(obj, dict):
        return {
            k: ("***" if k.lower() in _SECRET_KEYS else _redact(v))
            for k, v in obj.items()
        }
    if isinstance(obj, list):
        return [_redact(v) for v in obj]
    return obj


class JsonFormatter(logging.Formatter):
    """Render LogRecords as a single-line JSON object."""

    def format(self, record: logging.LogRecord) -> str:
        payload: dict[str, Any] = {
            "timestamp": self.formatTime(record, "%Y-%m-%dT%H:%M:%S"),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
        }
        # Pick up any extra={} fields the caller attached.
        for key in ("backend", "operation", "path", "duration_ms",
                    "result", "error_type", "status_code", "attempt"):
            if hasattr(record, key):
                payload[key] = getattr(record, key)
        if record.exc_info:
            payload["exception"] = self.formatException(record.exc_info)
        return json.dumps(payload, ensure_ascii=False, default=str)


def configure_logging(level: str | None = None) -> logging.Logger:
    """Idempotently install a JSON handler on the tropi_storage logger.

    Returns the configured logger. Safe to call repeatedly — only the first call
    attaches a handler.
    """
    logger = logging.getLogger(LOGGER_NAME)
    if not getattr(logger, "_tropi_configured", False):
        handler = logging.StreamHandler()
        handler.setFormatter(JsonFormatter())
        logger.addHandler(handler)
        logger.propagate = False
        logger._tropi_configured = True  # type: ignore[attr-defined]

    resolved = (level or os.getenv("LOG_LEVEL") or "INFO").upper()
    logger.setLevel(getattr(logging, resolved, logging.INFO))
    return logger


def get_logger() -> logging.Logger:
    return logging.getLogger(LOGGER_NAME)


@contextmanager
def log_operation(backend: str, operation: str, path: str):
    """Time and log a single adapter operation.

    Emits one INFO record on success and one ERROR (or WARNING) record on
    exception. NotFoundError is logged at WARNING level — it's a normal
    control-flow signal (caller probes a path that may or may not exist)
    and shouldn't pollute Sentry with thousands of "missing file" events.
    All other exceptions log at ERROR.
    """
    from .exceptions import NotFoundError  # local import to avoid cycle

    logger = get_logger()
    start = time.perf_counter()
    try:
        yield
    except Exception as exc:
        duration_ms = round((time.perf_counter() - start) * 1000, 2)
        log_level = logging.WARNING if isinstance(exc, NotFoundError) else logging.ERROR
        logger.log(
            log_level,
            f"{operation} failed",
            extra={
                "backend": backend,
                "operation": operation,
                "path": path,
                "duration_ms": duration_ms,
                "result": "error",
                "error_type": type(exc).__name__,
            },
        )
        raise
    else:
        duration_ms = round((time.perf_counter() - start) * 1000, 2)
        logger.info(
            f"{operation} ok",
            extra={
                "backend": backend,
                "operation": operation,
                "path": path,
                "duration_ms": duration_ms,
                "result": "success",
            },
        )


def init_sentry_if_configured() -> bool:
    """Initialize Sentry SDK if SENTRY_DSN is set. Returns True if initialized."""
    dsn = os.getenv("SENTRY_DSN", "").strip()
    if not dsn:
        return False
    try:
        import sentry_sdk  # type: ignore
    except ImportError:
        get_logger().warning("SENTRY_DSN set but sentry-sdk not installed; skipping")
        return False
    # Init exactly once: get_adapter() calls this on every invocation, and concurrent
    # re-inits race inside sentry_sdk's auto-enabling integration imports (a failed
    # integration import is retried on every init and is not thread-safe — seen live
    # as "cannot import name 'LangchainIntegration'" when a poll thread collided with
    # an APScheduler job). auto_enabling_integrations=False: the adapter only needs
    # plain error reporting, and the host service's own init decides integrations.
    if sentry_sdk.is_initialized():
        return True
    sentry_sdk.init(
        dsn=dsn,
        traces_sample_rate=0.05,
        send_default_pii=False,
        auto_enabling_integrations=False,
    )
    return True
