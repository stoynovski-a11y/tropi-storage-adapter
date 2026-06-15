"""Route logical path first-segments to specific SharePoint sites and named libraries.

A *route* maps a top-level folder name (the first path component) to a
(site_path, library_name) tuple, allowing one adapter instance to serve
multiple SharePoint document libraries transparently.

Configuration
-------------
``DEFAULT_ROUTES`` is intentionally empty in this public package.  All real
routes must be supplied at runtime via the ``M365_ROUTES`` environment
variable as a JSON object whose values are two-element arrays::

    M365_ROUTES='{"Top A": ["/sites/alpha", "Library A"], "Top B": ["/sites/beta", "Library B"]}'

Without ``M365_ROUTES`` (and no ``default_site`` argument) ``resolve_route``
will raise ``BackendError`` for any path — this is expected; configure the env
var for every deployment.

``M365_DEFAULT_LIBRARY`` is honoured only by the backend; routing.py itself
does not consume it.
"""
from __future__ import annotations

import json
import os
from typing import Optional

from .exceptions import BackendError
from .path_utils import normalize_path


# ---------------------------------------------------------------------------
# Static default table — empty; real routes are provided via M365_ROUTES.
# ---------------------------------------------------------------------------

DEFAULT_ROUTES: dict[str, tuple[str, str]] = {}


# ---------------------------------------------------------------------------
# load_strip_prefix
# ---------------------------------------------------------------------------

def load_strip_prefix() -> str | None:
    """Return the legacy path prefix to strip, or None if not configured.

    Reads ``M365_STRIP_PREFIX`` from the environment.  The value is stripped of
    surrounding whitespace.  An empty (or whitespace-only) value is treated as
    unset and returns None.

    Example: if ``M365_STRIP_PREFIX=/Legacy`` then every path passed to
    ``resolve_route`` that starts with ``/Legacy/`` is normalised to the
    portion after ``/Legacy`` before routing decisions are made — so
    ``/Legacy/Top A/x`` routes as if it were ``/Top A/x``.
    """
    val = os.getenv("M365_STRIP_PREFIX", "").strip()
    return val if val else None


# ---------------------------------------------------------------------------
# load_routes
# ---------------------------------------------------------------------------

def load_routes() -> dict[str, tuple[str, str]]:
    """Return the merged route table (defaults + optional M365_ROUTES override).

    Raises BackendError if ``M365_ROUTES`` is set but cannot be parsed or has
    the wrong shape.
    """
    routes: dict[str, tuple[str, str]] = dict(DEFAULT_ROUTES)

    env_val = os.getenv("M365_ROUTES", "").strip()
    if not env_val:
        return routes

    try:
        extra = json.loads(env_val)
    except json.JSONDecodeError as exc:
        raise BackendError(
            f"M365_ROUTES is not valid JSON: {exc}"
        ) from exc

    if not isinstance(extra, dict):
        raise BackendError(
            "M365_ROUTES must be a JSON object mapping segment names to "
            "[site_path, library_name] arrays; got a non-object value."
        )

    for key, val in extra.items():
        if (
            not isinstance(val, (list, tuple))
            or len(val) != 2
            or not all(isinstance(v, str) for v in val)
        ):
            raise BackendError(
                f"M365_ROUTES entry {key!r} must be a two-element string array "
                f'[site_path, library_name]; got {val!r}.'
            )
        routes[key] = (val[0], val[1])

    return routes


# ---------------------------------------------------------------------------
# load_folder_pins
# ---------------------------------------------------------------------------

def load_folder_pins() -> dict[str, str]:
    """Return the folder-ID pin table from the optional ``M365_FOLDER_IDS`` env var.

    Maps a logical folder path (the location callers still reference) to a Graph
    driveItem id::

        M365_FOLDER_IDS='{"/Legacy/Top A/Old Name": "01ABC...", ...}'

    The backend addresses a pinned folder — and everything beneath it — by that
    stable id instead of by name, so the folder can be renamed or moved within
    its drive without breaking any caller. Empty/unset → no pins (normal
    name-based addressing).

    Raises BackendError if set but not a JSON object of string→non-empty-string.
    Keys are normalised; matching against item paths happens in the backend.
    """
    env_val = os.getenv("M365_FOLDER_IDS", "").strip()
    if not env_val:
        return {}

    try:
        data = json.loads(env_val)
    except json.JSONDecodeError as exc:
        raise BackendError(f"M365_FOLDER_IDS is not valid JSON: {exc}") from exc

    if not isinstance(data, dict):
        raise BackendError(
            "M365_FOLDER_IDS must be a JSON object mapping a logical folder "
            "path to a driveItem id string; got a non-object value."
        )

    pins: dict[str, str] = {}
    for key, val in data.items():
        if not isinstance(val, str) or not val:
            raise BackendError(
                f"M365_FOLDER_IDS entry {key!r} must map to a non-empty "
                f"driveItem id string; got {val!r}."
            )
        pins[normalize_path(key)] = val
    return pins


# ---------------------------------------------------------------------------
# resolve_route
# ---------------------------------------------------------------------------

def resolve_route(
    path: str,
    routes: dict[str, tuple[str, str]],
    default_site: Optional[str] = None,
    default_drive: Optional[str] = None,
    strip_prefix: Optional[str] = None,
) -> tuple[str, Optional[str], str]:
    """Map *path* to (site_path, drive_name | None, item_path).

    Resolution rules
    ----------------
    1. If *strip_prefix* is given, the leading prefix is removed from *path*
       before any routing decision is made.  Paths that do not start with the
       prefix are left unchanged (idempotent).  Example: prefix ``/Legacy``
       turns ``/Legacy/Top A/sub/x.xlsx`` into routing on ``/Top A/sub/x.xlsx``.
       A path equal to the prefix itself (e.g. ``/Legacy``) normalises to
       ``"/"``.
    2. ``path == "/"``  — if *default_site* is given return
       ``(default_site, default_drive, "/")``, otherwise raise BackendError.
    3. First path segment present in *routes* → strip the segment and return
       the configured (site_path, drive_name) with the remainder as item_path.
    4. First path segment NOT in *routes* but *default_site* given → return
       the whole path under the default site/drive (backward-compat).
    5. Otherwise raise BackendError.

    Examples
    --------
    Given ``routes = {"Top A": ("/sites/alpha", "Library A")}``:

    >>> resolve_route("/Top A/sub/x.xlsx", routes)
    ("/sites/alpha", "Library A", "/sub/x.xlsx")

    >>> resolve_route("/Top A", routes)
    ("/sites/alpha", "Library A", "/")
    """
    p = normalize_path(path)

    if strip_prefix:
        pfx = "/" + strip_prefix.strip("/")
        if p == pfx:
            p = "/"
        elif p.startswith(pfx + "/"):
            p = p[len(pfx):]  # leaves a leading "/..."

    if p == "/":
        if default_site:
            return (default_site, default_drive, "/")
        raise BackendError(
            "Cannot address multi-site root '/'; route a specific path or "
            "set a default site (M365_SITE_PATH)."
        )

    # Split off the first segment.
    # p is guaranteed to start with "/" and have at least one more char.
    without_leading = p[1:]  # strip the leading "/"
    if "/" in without_leading:
        seg, rest = without_leading.split("/", 1)
        remainder = "/" + rest
    else:
        seg = without_leading
        remainder = "/"

    if seg in routes:
        site_path, drive_name = routes[seg]
        return (site_path, drive_name, remainder)

    if default_site:
        # Full path under the default site — preserves old single-site behaviour.
        return (default_site, default_drive, p)

    raise BackendError(
        f"No M365 route for top-level segment {seg!r} (path={path!r}); "
        "add it to DEFAULT_ROUTES or set the M365_ROUTES env var."
    )
