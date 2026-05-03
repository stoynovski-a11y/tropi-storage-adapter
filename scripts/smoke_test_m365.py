"""Live smoke test against a SharePoint site via the M365 backend.

Loads .env, then forces STORAGE_BACKEND=m365 (the .env may have it set to
dropbox; we don't edit .env from here because of the safety hook). Runs the
end-to-end write/read/delete cycle against the configured site.

Usage:
    python scripts/smoke_test_m365.py

Configure M365_SITE_HOSTNAME, M365_SITE_PATH, and optionally SMOKE_TEST_FOLDER
in your .env before running.
"""
from __future__ import annotations

import datetime as dt
import os
import sys
import traceback
from pathlib import Path

# 1. Load .env (project root) before importing the adapter.
try:
    from dotenv import load_dotenv
except ImportError:
    print("python-dotenv is not installed; install with: pip install python-dotenv")
    sys.exit(2)

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")

# 2. Force m365 backend regardless of what .env says.
os.environ["STORAGE_BACKEND"] = "m365"

from tropi_storage import get_adapter  # noqa: E402


EXPECTED_ROOT_ENTRIES = set()  # populate with folder names you expect at the site root
TEST_FOLDER = os.environ.get("SMOKE_TEST_FOLDER", "/Documents")
TEST_NAME = "smoke-test.txt"
TEST_PATH = f"{TEST_FOLDER}/{TEST_NAME}"


def step(label: str, fn):
    print(f"\n=== {label} ===")
    try:
        result = fn()
        print(f"OK: {label}")
        return True, result
    except Exception as e:  # noqa: BLE001
        print(f"FAIL: {label}: {type(e).__name__}: {e}")
        traceback.print_exc()
        return False, None


def main() -> int:
    results: dict[str, bool] = {}

    # --- get adapter ---
    ok, adapter = step("Build adapter", get_adapter)
    results["build_adapter"] = ok
    if not ok:
        return 1
    print(f"adapter.backend_name = {adapter.backend_name}")

    # --- healthcheck ---
    def _hc():
        hc = adapter.healthcheck()
        print(f"healthcheck = {hc}")
        assert hc["backend"] == "m365"
        assert hc["authenticated"] is True, "authenticated=False"
        assert hc["can_list_root"] is True, "can_list_root=False"
        return hc
    ok, _ = step("healthcheck()", _hc)
    results["healthcheck"] = ok

    # --- list root ---
    def _list_root():
        items = adapter.list("/")
        names = sorted(i["name"] for i in items)
        print(f"root contains {len(items)} item(s):")
        for i in items:
            print(f"  - {i['type']:6} {i['name']}")
        missing = EXPECTED_ROOT_ENTRIES - set(names)
        if missing:
            print(f"WARN: expected entries not present: {missing}")
        return items
    ok, _ = step("list('/')", _list_root)
    results["list_root"] = ok

    # --- write test file ---
    timestamp = dt.datetime.now(dt.timezone.utc).isoformat()
    body = f"storage adapter smoke test\nWritten at: {timestamp}\n".encode("utf-8")

    def _write():
        meta = adapter.write(TEST_PATH, body, overwrite=True)
        print(f"write() returned: name={meta.get('name')!r} size={meta.get('size')} id={meta.get('id')}")
        return meta
    ok, write_meta = step(f"write('{TEST_PATH}')", _write)
    results["write"] = ok

    # --- read test file ---
    def _read():
        got = adapter.read(TEST_PATH)
        print(f"read() returned {len(got)} bytes")
        if got != body:
            raise AssertionError(
                f"content mismatch:\n  expected {body!r}\n  got      {got!r}"
            )
        return got
    ok, _ = step(f"read('{TEST_PATH}') and verify content matches", _read)
    results["read_match"] = ok

    # --- delete test file ---
    def _delete():
        adapter.delete(TEST_PATH)
        # confirm gone
        meta = adapter.get_metadata(TEST_PATH)
        if meta.get("exists"):
            raise AssertionError("file still exists after delete()")
        print(f"deleted; get_metadata reports exists={meta.get('exists')}")
    ok, _ = step(f"delete('{TEST_PATH}')", _delete)
    results["delete"] = ok

    # --- summary ---
    print("\n=== SUMMARY ===")
    for name, passed in results.items():
        print(f"  [{ 'PASS' if passed else 'FAIL' }] {name}")

    all_pass = all(results.values())
    print(f"\nOverall: {'PASS' if all_pass else 'FAIL'}")
    return 0 if all_pass else 1


if __name__ == "__main__":
    sys.exit(main())
