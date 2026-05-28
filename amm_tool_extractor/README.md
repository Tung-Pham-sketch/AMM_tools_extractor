# AMM Tool Extractor

Extracts **Support Equipment / Special Tools** from Boeing 787 AMM tasks (S1000D HTML format),
recursively following all referenced tasks to build a complete tool list.

---

## Setup

```
pip install beautifulsoup4 openpyxl lxml
```

---

## Folder layout

```
amm_tool_extractor/
├── input/          ← Drop your parent task .HTM file here (one at a time)
├── output/         ← Excel report is saved here
├── config.py       ← Edit this: zip file paths, depth limit, etc.
├── main.py         ← Run this
├── extractor/
│   ├── html_parser.py      Parses one HTM file
│   ├── zip_resolver.py     Finds HTM files inside zip archives by DMC code
│   └── crawler.py          BFS traversal of task references
└── writer/
    └── excel_writer.py     Writes the .xlsx report
```

---

## Quick start

1. **Edit `config.py`** — add your zip file paths:

   ```python
   AMM_ZIP_PATHS = [
       r"C:\AMM_Data\B787_AMM_Vol1.zip",
       r"C:\AMM_Data\B787_AMM_Vol2.zip",
       r"C:\AMM_Data\B787_AMM_Vol3.zip",
       r"C:\AMM_Data\B787_AMM_Vol4.zip",
   ]
   ```

2. **Drop your parent task HTM** into the `input/` folder.

3. **Run:**

   ```
   python main.py
   ```
   Or specify paths explicitly:
   ```
   python main.py --input path/to/task.HTM --output path/to/report.xlsx
   ```

4. **Open the Excel report** in `output/`.

---

## Excel output

| Sheet | Contents |
|---|---|
| **Tool Summary** | Deduplicated tool list — one row per unique Reference ID (SPL-XXXX / STD-XXXX), with all part numbers and all tasks that require it |
| **Task Detail** | One row per tool per task, showing the full reference chain and which task introduced each tool |

---

## Key config options (`config.py`)

| Setting | Default | Description |
|---|---|---|
| `AMM_ZIP_PATHS` | `[]` | List of paths to your Boeing AMM zip files |
| `MAX_DEPTH` | `None` | Recursion depth. `None` = unlimited. `1` = parent + direct refs only |
| `REVISION_STRATEGY` | `"latest"` | When multiple revisions of a DMC exist: `"latest"` or `"first"` |
| `REFERENCE_FILTER_PREFIX` | `None` | Only follow refs matching this DMC prefix (e.g. `"B787-A-32"`) |
| `OUTPUT_FILENAME` | `AMM_Tools_Report.xlsx` | Base name for the output file |

---

## How it works

```
input/parent_task.HTM
        │
        ▼
    html_parser          ← extracts tools + reference DMC list
        │
        ▼
    crawler (BFS)        ← for each ref DMC not yet visited:
        │                     zip_resolver → find & read HTM from zip
        │                     html_parser  → extract tools + more refs
        │                     repeat until no new refs
        ▼
    excel_writer         ← Tool Summary + Task Detail sheets
        │
        ▼
output/AMM_Tools_Report_YYYYMMDD_HHMMSS.xlsx
```

### Loop protection
The crawler maintains a **visited set** of DMC codes. Circular references
(Task A → Task B → Task A) are automatically broken regardless of depth setting.

### Missing references
If a referenced DMC is not found in any zip file, it is listed in the
Task Detail sheet with status `NOT FOUND IN ZIPS` so you can investigate manually.
