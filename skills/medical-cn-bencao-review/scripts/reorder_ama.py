#!/usr/bin/env python3
"""Compatibility gate for the retired fixed-AMA citation workflow."""

from pathlib import Path
import runpy
import sys


CONFIRMATION_FLAG = "--journal-numeric-confirmed"


if CONFIRMATION_FLAG not in sys.argv:
    raise SystemExit(
        "reorder_ama.py is retired. Verify the target journal's current official "
        "instructions, then use audit_numeric_citations.py. If an older workflow "
        "must call this path and the journal is confirmed to use numeric sequential "
        f"citations, add {CONFIRMATION_FLAG}."
    )

sys.argv.remove(CONFIRMATION_FLAG)
target = Path(__file__).with_name("audit_numeric_citations.py")
runpy.run_path(target, run_name="__main__")
