'''
Loads all storage Excel files from STORAGE_DIR and builds a lookup:
    part_number (uppercase) → list of StockRecord

Each StockRecord is one (Station, Store, Location) group with summed Qty.
Multiple rows with the same part number and same location are summed.'''


import os
import re
import glob
import logging
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple
import pandas as pd

from .html_parser import ToolEntry

logger = logging.getLogger(__name__)

# Regex to parse one part_number string from ToolEntry.part_numbers
# Matches both:
#   (Part #: K32020-1, Supplier: ...)
#   (Opt Part #: K32012-1, Supplier: ...)
_PN_RE = re.compile(r'\((Opt )?Part #:\s*([^,]+),', re.IGNORECASE)


@dataclass
class StockRecord:
    part_number: str      # as found in storage file
    station: str
    store: str
    location: str
    qty: int
    is_opt: bool          # was this an "Opt Part #" in the AMM?


@dataclass
class StockResult:
    """All stock information for one ToolEntry."""
    records: List[StockRecord] = field(default_factory=list)

    @property
    def found(self) -> bool:
        return len(self.records) > 0

    @property
    def total_qty(self) -> int:
        return sum(r.qty for r in self.records)


def _parse_part_numbers(part_numbers: List[str]) -> List[Tuple[str, bool]]:
    """
    Extract (part_number, is_opt) tuples from a ToolEntry.part_numbers list.
    Returns empty list for STD tools (no part numbers).
    """
    results = []
    for raw in part_numbers:
        m = _PN_RE.match(raw.strip())
        if m:
            is_opt = bool(m.group(1))
            pn = m.group(2).strip()
            results.append((pn, is_opt))
    return results


class StockChecker:
    """
    Loads all .xlsx files from storage_dir once and builds an index.
    Call check(tool) to get StockResult for a ToolEntry.
    """

    def __init__(self, storage_dir: str):
        # index: normalised_part_number → list of (station, store, location, qty)
        self._index: Dict[str, List[Tuple[str, str, str, int]]] = {}
        self._load(storage_dir)

    def _load(self, storage_dir: str) -> None:
        files = glob.glob(os.path.join(storage_dir, "*.xlsx"))
        if not files:
            logger.warning("No .xlsx files found in STORAGE_DIR: %s", storage_dir)
            return

        frames = []
        for path in files:
            try:
                df = pd.read_excel(path, dtype={"Part Number": str})
                df["_source"] = os.path.basename(path)
                frames.append(df)
                logger.info("Loaded storage file: %s (%d rows)", os.path.basename(path), len(df))
            except Exception as exc:
                logger.warning("Could not read %s: %s", path, exc)

        if not frames:
            return

        combined = pd.concat(frames, ignore_index=True)

        # Required columns check
        required = {"Part Number", "Station", "Store", "Location", "Qty"}
        missing = required - set(combined.columns)
        if missing:
            logger.error("Storage files missing columns: %s", missing)
            return

        # Group by (Part Number, Station, Store, Location) and sum Qty
        combined["Part Number"] = combined["Part Number"].fillna("").str.strip()
        combined["Station"]     = combined["Station"].fillna("").astype(str).str.strip()
        combined["Store"]       = combined["Store"].fillna("").astype(str).str.strip()
        combined["Location"]    = combined["Location"].fillna("").astype(str).str.strip()
        combined["Qty"]         = pd.to_numeric(combined["Qty"], errors="coerce").fillna(0).astype(int)

        grouped = (
            combined
            .groupby(["Part Number", "Station", "Store", "Location"], as_index=False)["Qty"]
            .sum()
        )

        for _, row in grouped.iterrows():
            pn_key = row["Part Number"].upper()
            if pn_key not in self._index:
                self._index[pn_key] = []
            self._index[pn_key].append((
                row["Station"], row["Store"], row["Location"], int(row["Qty"])
            ))

        logger.info("Stock index built: %d unique part numbers across %d file(s).",
                    len(self._index), len(frames))

    def check(self, tool: ToolEntry) -> StockResult:
        """
        Look up all part numbers for a tool and return combined StockResult.
        Each matched (Station, Store, Location) becomes one StockRecord.
        """
        result = StockResult()
        parsed = _parse_part_numbers(tool.part_numbers)

        if not parsed:
            # STD tools have no part numbers — cannot check stock
            return result

        for pn, is_opt in parsed:
            matches = self._index.get(pn.upper(), [])
            for station, store, location, qty in matches:
                result.records.append(StockRecord(
                    part_number=pn,
                    station=station,
                    store=store,
                    location=location,
                    qty=qty,
                    is_opt=is_opt,
                ))

        return result