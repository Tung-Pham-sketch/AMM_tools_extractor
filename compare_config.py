"""
compare_config.py
=================
Configuration for the preload comparator (compare_preload.py).
Edit the paths below before running.
"""

import os
from pathlib import Path
# ── Tool report ────────────────────────────────────────────────────────────────
# Folder where main.py writes its *Tools Report*.xlsx files.
TOOL_REPORT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

# Leave as "" to auto-pick the most recently modified *Tools Report*.xlsx.
# Set to a full path to use a specific file.
TOOL_REPORT_FILE = r"D:\D\AMM__tool_extractor\amm_tool_extractor\amm_tool_extractor\output\A861_LDG_Replacment_tool_report.xlsx"

# ── Preload file ───────────────────────────────────────────────────────────────
# Folder where preload .xlsx files are stored.
PRELOAD_DIR = r"D:\D\AMM__tool_extractor\preload"

# Leave as "" to auto-pick the most recently modified .xlsx in PRELOAD_DIR.
# Set to a full path to use a specific file.
PRELOAD_FILE = r"D:\D\AMM__tool_extractor\amm_tool_extractor\Preload\A861-030426-CHK-A12_LDG.xlsx"

# ── Output ─────────────────────────────────────────────────────────────────────
# Comparison report is saved here (defaults to same output folder as tool report).
COMPARE_OUTPUT_DIR = TOOL_REPORT_DIR