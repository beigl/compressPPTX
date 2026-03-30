#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# -----------------------------------------------------------------------------
# File:        compress_pptx.py
# Project:     compressPPTX
# Purpose:     Backward-compatible wrapper for the packaged CLI entry point.
# Author:      Michael Beigl
# Copyright:   CC-BY 4.0
# License:     CC-BY 4.0
# Created:     2026-03-30
# Python:      3.10+
# Usage:       python compress_pptx.py --input-dir input --output-dir output
# -----------------------------------------------------------------------------

from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from compresspptx.core import main

__author__ = "Michael Beigl"
__copyright__ = "CC-BY 4.0"
__license__ = "CC-BY 4.0"


if __name__ == "__main__":
    main()
