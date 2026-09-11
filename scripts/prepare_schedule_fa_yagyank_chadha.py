#!/usr/bin/env python3
"""Backward-compatible entry: generates both Yagyank Chadha account workbooks. """
from __future__ import annotations
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from prepare_schedule_fa_yagyank_u15172057 import run as run_main
from prepare_schedule_fa_yagyank_u20291582 import run as run_satellite

if __name__ == "__main__":
    run_main()
    run_satellite()
