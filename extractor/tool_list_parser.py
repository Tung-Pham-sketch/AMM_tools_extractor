"""
tool_list_parser.py
===================
Parses the Boeing AMM Tool List document (DMC: B787-A-00-40-01-00A-00NA-A)
and builds a lookup:
    reference_id (e.g. SPL-8455) → ToolListEntry

The document has 3 tables:
    Table 0: STD tools  (Reference, Description)
    Table 1: COM tools  (Reference, Description, Part Number, Supplier, A/P Effectivity)
    Table 2: SPL tools  (Reference, Description, Part Number, Supplier, A/P Effectivity)

Each entry may span multiple rows for alternate part numbers.
A new entry begins when column 0 contains a Reference ID.
Continuation rows have an empty column 0.
"""

import re
import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

_REF_RE = re.compile(r'^(SPL|STD|COM)-\d+$', re.IGNORECASE)


@dataclass
class ToolListEntry:
    reference_id: str
    description: str
    # Each item: {"part_number": str, "is_opt": bool,
    #             "supplier": str, "effectivity": str}
    part_numbers: List[dict] = field(default_factory=list)

    def all_part_number_strings(self) -> List[str]:
        """Return part numbers formatted the same way as ToolEntry.part_numbers."""
        lines = []
        for pn in self.part_numbers:
            opt = "Opt Part #" if pn["is_opt"] else "Part #"
            lines.append(
                f"({opt}: {pn['part_number']}, "
                f"Supplier: {pn['supplier']}, "
                f"A/P Effectivity: {pn['effectivity']})"
            )
        return lines


class ToolListParser:
    """
    Parse the tool list HTM and expose a lookup by Reference ID.
    """

    def __init__(self, html_content: str):
        self._index: Dict[str, ToolListEntry] = {}
        self._parse(html_content)
        logger.info("Tool list parsed: %d entries indexed.", len(self._index))

    def _parse(self, html: str) -> None:
        soup = BeautifulSoup(html, "lxml")
        tables = soup.find_all("table")

        if len(tables) < 3:
            logger.warning("Tool list: expected 3 tables, found %d.", len(tables))

        for tbl_idx, tbl in enumerate(tables[:3]):
            has_pn = tbl_idx > 0   # Table 0 (STD) has no part number columns
            current: Optional[ToolListEntry] = None

            for row in tbl.find_all("tr"):
                cells = [c.get_text(strip=True) for c in row.find_all(["th", "td"])]
                if not cells:
                    continue

                col0 = cells[0]

                # ── Header row ─────────────────────────────────────────────
                if col0.lower() == "reference":
                    continue

                # ── New entry row ───────────────────────────────────────────
                if _REF_RE.match(col0):
                    current = ToolListEntry(
                        reference_id=col0.upper(),
                        description=cells[1] if len(cells) > 1 else "",
                    )
                    self._index[current.reference_id] = current

                    if has_pn and len(cells) >= 5:
                        pn_raw = cells[2]
                        is_opt = pn_raw.lower().startswith("opt:")
                        pn = pn_raw[4:].strip() if is_opt else pn_raw.strip()
                        current.part_numbers.append({
                            "part_number": pn,
                            "is_opt": is_opt,
                            "supplier": cells[3],
                            "effectivity": cells[4],
                        })
                    continue

                # ── Continuation row (extra part numbers) ───────────────────
                if current and has_pn and col0 == "" and len(cells) >= 3:
                    pn_raw = cells[0] if col0 != "" else cells[0]
                    # In continuation rows all cells shift left:
                    # [part_number, supplier, effectivity]
                    pn_raw = cells[0]
                    is_opt = pn_raw.lower().startswith("opt:")
                    pn = pn_raw[4:].strip() if is_opt else pn_raw.strip()
                    if pn:
                        current.part_numbers.append({
                            "part_number": pn,
                            "is_opt": is_opt,
                            "supplier": cells[1] if len(cells) > 1 else "",
                            "effectivity": cells[2] if len(cells) > 2 else "",
                        })

    def lookup(self, reference_id: str) -> Optional[ToolListEntry]:
        return self._index.get(reference_id.upper())

    def __len__(self):
        return len(self._index)