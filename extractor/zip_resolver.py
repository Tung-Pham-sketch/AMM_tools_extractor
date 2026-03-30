"""
zip_resolver.py
===============
Builds an index of DMC → filename by scanning zip files once,
then resolves DMC codes to HTML content on demand.

Filename pattern inside zips:
  {prefix}/DME-81205-A9301-{DMC}_{issue}-{revision}.HTM
  e.g. deploy/Application/v2/vie/787/amm/data/DME-81205-A9301-B787-A-32-00-30-00A-720A-A_032-00.HTM

When USE_LATEST_REVISION=True (default), if multiple issues exist
for the same DMC, the one with the highest _NNN-NN suffix is used.
"""

import re
import zipfile
import logging
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# Match filenames with or without .HTM extension:
#   DME-81205-A9301-B787-A-32-00-30-00A-720A-A_032-00.HTM   (with extension)
#   DME-81205-A9301-B787-A-32-00-30-00A-720A-A_032-00       (without extension)
_FNAME_RE = re.compile(
    r"(B787-[A-Za-z0-9\-]+?)_(\d{3}-\d{2})(\.HTM)?$", re.IGNORECASE
)


class ZipResolver:
    """
    Scans a list of zip files once and builds a DMC → (zip_path, entry_name)
    index.  Call resolve(dmc) to get the HTML content for a task.
    """

    def __init__(self, zip_paths: List[str], use_latest: bool = True,
                 internal_prefix: str = ""):
        self._use_latest = use_latest
        # Normalise prefix: strip trailing slash, use forward slashes
        self._prefix = internal_prefix.replace("\\", "/").rstrip("/")
        # dmc → (issue_str, zip_path, entry_name)
        self._index: Dict[str, Tuple[str, str, str]] = {}
        self._build_index(zip_paths)

    def _build_index(self, zip_paths: List[str]) -> None:
        for zip_path in zip_paths:
            try:
                with zipfile.ZipFile(zip_path, "r") as zf:
                    for entry in zf.namelist():
                        # Normalise to forward slashes
                        entry_norm = entry.replace("\\", "/")

                        # If a prefix is set, only consider entries inside it
                        if self._prefix:
                            if not entry_norm.startswith(self._prefix + "/"):
                                continue

                        basename = entry_norm.split("/")[-1]
                        m = _FNAME_RE.search(basename)
                        if not m:
                            continue
                        dmc = m.group(1)
                        issue = m.group(2)
                        existing = self._index.get(dmc)
                        if existing is None:
                            self._index[dmc] = (issue, zip_path, entry)
                        elif self._use_latest and issue > existing[0]:
                            self._index[dmc] = (issue, zip_path, entry)
            except (zipfile.BadZipFile, FileNotFoundError) as exc:
                logger.warning("Could not open zip %s: %s", zip_path, exc)

        logger.info("Zip index built: %d unique DMC codes found.", len(self._index))

    def resolve(self, dmc: str) -> Optional[str]:
        """Return the HTML content for a DMC code, or None if not found."""
        entry_info = self._index.get(dmc)
        if entry_info is None:
            logger.warning("DMC not found in any zip: %s", dmc)
            return None
        _, zip_path, entry_name = entry_info
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                return zf.read(entry_name).decode("utf-8", errors="replace")
        except Exception as exc:
            logger.error("Failed to read %s from %s: %s", entry_name, zip_path, exc)
            return None

    def known_dmcs(self) -> List[str]:
        return list(self._index.keys())