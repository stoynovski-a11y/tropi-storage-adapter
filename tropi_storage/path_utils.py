"""Path helpers — date/year templating and cross-backend normalization."""
from __future__ import annotations

import datetime as _dt
import re

_TEMPLATE_RE = re.compile(r"\{(year|month|day|ww)\}")


def expand_path(template: str, *, when: _dt.date | None = None) -> str:
    """Expand `{year}`, `{month}`, `{day}`, `{ww}` (ISO week) in a path template.

    Examples:
        expand_path('/x/{year}/foo.xlsx')          -> '/x/2026/foo.xlsx'
        expand_path('/x/{year}_{ww}_doc.xlsx')     -> '/x/2026_18_doc.xlsx'
        expand_path('/x/{year}/{month}/{day}/y')   -> '/x/2026/05/03/y'

    `when` defaults to today (local date). Pass a date for deterministic tests.
    """
    today = when or _dt.date.today()
    iso_year, iso_week, _ = today.isocalendar()
    values = {
        "year": f"{today.year}",
        "month": f"{today.month:02d}",
        "day": f"{today.day:02d}",
        # ISO week — use ISO calendar (matches Python isocalendar, matches Excel WEEKNUM type=21).
        "ww": f"{iso_week:02d}",
    }
    return _TEMPLATE_RE.sub(lambda m: values[m.group(1)], template)


def normalize_path(path: str) -> str:
    """Ensure a leading slash and no trailing slash (except for root '/').

    Both backends use this canonical form internally.
    """
    if not path:
        return "/"
    p = path.strip()
    if not p.startswith("/"):
        p = "/" + p
    if len(p) > 1 and p.endswith("/"):
        p = p.rstrip("/")
    return p


def split_parent(path: str) -> tuple[str, str]:
    """Return (parent_path, basename). Root parent of '/foo' is '/'."""
    p = normalize_path(path)
    if p == "/":
        return "/", ""
    parent, _, name = p.rpartition("/")
    return (parent or "/"), name
