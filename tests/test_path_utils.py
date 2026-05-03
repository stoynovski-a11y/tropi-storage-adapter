"""Tests for path templating helpers."""
from __future__ import annotations

import datetime as dt

import pytest

from tropi_storage.path_utils import expand_path, normalize_path, split_parent


class TestExpandPath:
    def test_year_token(self):
        d = dt.date(2026, 5, 3)
        assert expand_path("/x/{year}/foo.xlsx", when=d) == "/x/2026/foo.xlsx"

    def test_all_tokens(self):
        d = dt.date(2026, 5, 3)
        # 2026-05-03 is a Sunday — ISO week 18.
        assert expand_path("{year}_{month}_{day}_{ww}", when=d) == "2026_05_03_18"

    def test_month_and_day_zero_padded(self):
        d = dt.date(2026, 1, 9)
        assert expand_path("{year}/{month}/{day}", when=d) == "2026/01/09"

    def test_iso_week_padded(self):
        # ISO week 1 of 2025 starts 2024-12-30.
        d = dt.date(2024, 12, 30)
        assert "01" in expand_path("{ww}", when=d)

    def test_no_tokens_passthrough(self):
        assert expand_path("/Documents/foo.xlsx") == "/Documents/foo.xlsx"

    def test_unknown_token_left_alone(self):
        # {x} is not a recognized token; must remain literal.
        assert expand_path("/x/{x}/y") == "/x/{x}/y"

    def test_default_uses_today(self):
        out = expand_path("{year}")
        assert int(out) == dt.date.today().year


class TestNormalizePath:
    @pytest.mark.parametrize("inp,want", [
        ("foo", "/foo"),
        ("/foo", "/foo"),
        ("/foo/", "/foo"),
        ("/foo/bar/", "/foo/bar"),
        ("/", "/"),
        ("", "/"),
        ("  /a  ", "/a"),
    ])
    def test_normalize(self, inp, want):
        assert normalize_path(inp) == want


class TestSplitParent:
    @pytest.mark.parametrize("inp,parent,name", [
        ("/foo.xlsx", "/", "foo.xlsx"),
        ("/a/b/c.xlsx", "/a/b", "c.xlsx"),
        ("/a/b/", "/a", "b"),  # trailing slash stripped first → /a/b, then split → ('/a', 'b')
        ("/", "/", ""),
    ])
    def test_split(self, inp, parent, name):
        p, n = split_parent(inp)
        assert p == parent
        assert n == name
