"""
ipd_resolver.py
===============
Indexes IPD JSON files from zip archives by DMC code.
"""

import re
import json
import zipfile
import logging
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

_FNAME_RE = re.compile(
    r"(B787-A-[\w\-]+?)_(\d{3}-\d{2})\.json$", re.IGNORECASE
)


class IpdResolver:
    def __init__(self, zip_paths: List[str], use_latest: bool = True,
                 internal_prefix: str = ""):
        self._use_latest = use_latest
        self._prefix     = internal_prefix.replace("\\", "/").rstrip("/")
        self._index: Dict[str, Tuple[str, str, str]] = {}
        self._build_index(zip_paths)

    def _build_index(self, zip_paths: List[str]) -> None:
        for zip_path in zip_paths:
            try:
                with zipfile.ZipFile(zip_path, "r") as zf:
                    all_entries = zf.namelist()

                    # ── Diagnostic: log first 10 JSON entries so you can
                    #    verify IPD_ZIP_INTERNAL_PREFIX is correct ──────────
                    json_entries = [e for e in all_entries if e.endswith('.json')]
                    if json_entries:
                        logger.info("IPD zip sample paths (first 5 .json entries):")
                        for e in json_entries[:5]:
                            logger.info("  %s", e)
                    else:
                        logger.warning("No .json files found at all in %s", zip_path)

                    for entry in all_entries:
                        entry_norm = entry.replace("\\", "/")
                        if self._prefix:
                            if not entry_norm.startswith(self._prefix + "/"):
                                continue
                        basename = entry_norm.split("/")[-1]
                        m = _FNAME_RE.search(basename)
                        if not m:
                            continue
                        dmc   = m.group(1).upper()
                        issue = m.group(2)
                        existing = self._index.get(dmc)
                        if existing is None:
                            self._index[dmc] = (issue, zip_path, entry)
                        elif self._use_latest and issue > existing[0]:
                            self._index[dmc] = (issue, zip_path, entry)
            except (zipfile.BadZipFile, FileNotFoundError) as exc:
                logger.warning("Could not open IPD zip %s: %s", zip_path, exc)

        logger.info("IPD index built: %d unique DMC codes.", len(self._index))
        if len(self._index) == 0:
            logger.error(
                "IPD index is EMPTY. Check that IPD_ZIP_INTERNAL_PREFIX matches "
                "the actual folder path shown in the diagnostic log above."
            )

    def resolve(self, dmc: str) -> Optional[dict]:
        key = dmc.upper()
        entry_info = self._index.get(key)
        if entry_info is None:
            logger.debug("IPD DMC not found: %s", dmc)
            return None
        _, zip_path, entry_name = entry_info
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                raw = zf.read(entry_name).decode("utf-8", errors="replace")
                return json.loads(raw)
        except Exception as exc:
            logger.error("Failed to read IPD %s: %s", entry_name, exc)
            return None