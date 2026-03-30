"""
AMM Tool Extractor - Configuration
====================================
Edit the settings below before running.
"""

import os

# ── AMM Zip files ──────────────────────────────────────────────────────────────
AMM_ZIP_FILES = [
    r"D:\787 Toolbox\Offline_VIE_B787TBR-02MAR2026_20260302011557_F_Part_2_of_4.zip",
    r"D:\787 Toolbox\Offline_VIE_B787TBR-02MAR2026_20260302011557_F_Part_3_of_4.zip",
    r"D:\787 Toolbox\Offline_VIE_B787TBR-02MAR2026_20260302011557_F_Part_4_of_4.zip",
]
AMM_ZIP_INTERNAL_PREFIX = "deploy/Application/v2/vie/787/amm/data"

# ── IPD Zip files ──────────────────────────────────────────────────────────────
IPD_ZIP_FILES = [
    r"D:\787 Toolbox\Offline_VIE_B787TBR-02MAR2026_20260302011557_F_Part_2_of_4.zip",
    r"D:\787 Toolbox\Offline_VIE_B787TBR-02MAR2026_20260302011557_F_Part_3_of_4.zip",
    r"D:\787 Toolbox\Offline_VIE_B787TBR-02MAR2026_20260302011557_F_Part_4_of_4.zip",
]
IPD_ZIP_INTERNAL_PREFIX = "deploy/Application/v2/vie/787/ipd/data"


# ── Parent task DMC codes ──────────────────────────────────────────────────────
# Add as many DMC codes as needed for this workorder.
PARENT_DMCS = [
    "B787-A-32-11-01-00B-520A-A",   # Main Landing Gear - Removal
    "B787-A-32-11-01-00B-720A-A",   # Main Landing Gear - Installation
    "B787-A-32-21-01-00B-520A-A",   # Nose Landing Gear - Removal
    "B787-A-32-21-01-00B-720A-A",   # Nose Landing Gear - Installation
]

# ── Tool List DMC ──────────────────────────────────────────────────────────────
TOOL_LIST_DMC        = "B787-A-00-40-01-00A-00NA-A"
CONSUMABLE_LIST_DMC  = "B787-A-00-40-01-00A-00LB-A"

# ── Storage files folder ───────────────────────────────────────────────────────
STORAGE_DIR = r"D:\D\AMM__tool_extractor\amm_tool_extractor\amm_tool_extractor\tool_storage"

# ── Ignore List ────────────────────────────────────────────────────────────────
# Excel file containing part numbers to exclude from output
# Must have 'prq2.partno' column
# Panels/Panel Assy are auto-ignored (hardcoded) - no need to add them here
IGNORE_LIST_FILE = r"D:\D\AMM__tool_extractor\amm_tool_extractor\amm_tool_extractor\ignored_items\Ignore_item.xlsx"
# ── Output ─────────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# ── Traversal ──────────────────────────────────────────────────────────────────
MAX_DEPTH          = 3
USE_LATEST_REVISION = True