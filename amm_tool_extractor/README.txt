AMM Tool Extractor
==================

SETUP
-----
1. Install Python 3.10+
2. pip install -r requirements.txt

CONFIGURATION (config.py)
--------------------------
Edit AMM_ZIP_FILES to point to your Boeing AMM zip files, e.g.:
  AMM_ZIP_FILES = [
      r"D:\AMM\B787_Part1.zip",
      r"D:\AMM\B787_Part2.zip",
      ...
  ]

MAX_DEPTH controls reference traversal depth:
  None = unlimited (recommended)
  0    = parent task only
  1    = parent + direct references only

USAGE
-----
1. Drop the parent task HTM file into the  input/  folder
2. Run:  python main.py
3. Pick up the report from  output/amm_tools_report.xlsx

OUTPUT — amm_tools_report.xlsx
-------------------------------
Sheet "Summary"
  - One row per unique tool/equipment Reference ID
  - Deduplicated across all tasks
  - "Found In Tasks" columns show which tasks need each tool

Sheet "Detail"
  - Every tool occurrence in BFS traversal order
  - Parent task rows highlighted in blue
  - Tasks with no Support Equipment show a placeholder row
  - Depth column: Parent / Ref (depth 1) / Ref (depth 2) / ...

NOTES
-----
- The program never follows the same DMC twice (loop-safe).
- References are taken from the formal References table only,
  not from inline cross-references within procedure steps.
- Missing references (not found in any zip) are logged as warnings
  and traversal continues without them.
- When multiple revisions of a DMC exist, the latest is used.
