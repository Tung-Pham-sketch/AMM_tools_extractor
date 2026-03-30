"""
consumable_list_parser.py
=========================
Parses the Boeing AMM Consumable Material List
(DMC: B787-A-00-40-01-00A-00LB-A) and builds a lookup:
    reference_id (e.g. G01048) → ConsumableListEntry

Two table types exist in the document:
  Standard (6 cols): Reference | Description | Specification |
                     Material | Supplier | Safety Data Sheet
  Engine   (7 cols): Reference | Engine Mfr Reference | Description |
                     Specification | Material | Supplier | Safety Data Sheet

Each entry may span multiple rows for alternate material/supplier options.
A new entry starts when column 0 matches the reference ID pattern [A-Z]ddddd.
"""

import re
import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

# Reference ID pattern: one uppercase letter followed by 5 digits
_REF_RE = re.compile(r'^[A-Z]\d{5}$')


@dataclass
class ConsumableListEntry:
    reference_id:  str
    description:   str
    specification: str
    material:      str        # trade name / product name
    supplier:      str        # supplier code
    safety_data_sheet: str


class ConsumableListParser:
    """
    Parse the consumable material list HTM and expose a lookup by Reference ID.
    """

    def __init__(self, html_content: str):
        self._index: Dict[str, ConsumableListEntry] = {}
        self._parse(html_content)
        logger.info("Consumable list parsed: %d entries indexed.", len(self._index))

    def _parse(self, html: str) -> None:
        soup = BeautifulSoup(html, "lxml")

        for tbl in soup.find_all("table"):
            rows = tbl.find_all("tr")
            if not rows:
                continue

            # Detect table type from header row
            header_cells = [c.get_text(strip=True).lower()
                            for c in rows[0].find_all(["th", "td"])]
            has_engine_col = "engine mfr reference" in header_cells
            # Column index offsets
            # Standard: ref=0, desc=1, spec=2, mat=3, sup=4, sds=5
            # Engine:   ref=0, eng=1, desc=2, spec=3, mat=4, sup=5, sds=6
            offset = 1 if has_engine_col else 0

            current: Optional[ConsumableListEntry] = None

            for row in rows[1:]:
                cells = [c.get_text(strip=True)
                         for c in row.find_all(["th", "td"])]
                if not cells:
                    continue

                col0 = cells[0]

                # ── New entry row ──────────────────────────────────────────
                if _REF_RE.match(col0):
                    desc  = cells[1 + offset] if len(cells) > 1 + offset else ""
                    spec  = cells[2 + offset] if len(cells) > 2 + offset else ""
                    mat   = cells[3 + offset] if len(cells) > 3 + offset else ""
                    sup   = cells[4 + offset] if len(cells) > 4 + offset else ""
                    sds   = cells[5 + offset] if len(cells) > 5 + offset else ""
                    current = ConsumableListEntry(
                        reference_id=col0,
                        description=desc,
                        specification=spec,
                        material=mat,
                        supplier=sup,
                        safety_data_sheet=sds,
                    )
                    self._index[col0] = current
                # ── Continuation row — skip (alternate materials not needed
                #    for enrichment; first entry is the primary one) ────────

    def lookup(self, reference_id: str) -> Optional[ConsumableListEntry]:
        return self._index.get(reference_id.strip())

    def __len__(self):
        return len(self._index)