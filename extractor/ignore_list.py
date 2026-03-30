"""
ignore_list.py
==============
Handles filtering of items that should be excluded from the output:
  1. Auto-ignore: Panels/Panel Assy (hardcoded)
  2. Manual-ignore: Items from the ignore list Excel file (by Part Number)
"""

import os
import logging
import pandas as pd
from typing import Set, List

logger = logging.getLogger(__name__)


class IgnoreList:
    """
    Loads an ignore list Excel file and provides filtering capabilities.
    Also includes hardcoded auto-ignore rules (e.g., panels).
    """

    def __init__(self, ignore_file_path: str = None):
        """
        Initialize the ignore list.

        Args:
            ignore_file_path: Path to the Excel file containing ignored part numbers.
                            If None or file doesn't exist, only auto-ignore rules apply.
        """
        self._ignored_part_numbers: Set[str] = set()
        self._load_ignore_file(ignore_file_path)

    def _load_ignore_file(self, file_path: str) -> None:
        """Load ignored part numbers from Excel file."""
        if not file_path:
            logger.info("No ignore file specified - using auto-ignore rules only.")
            return

        if not os.path.isfile(file_path):
            logger.warning("Ignore file not found: %s - using auto-ignore rules only.", file_path)
            return

        try:
            df = pd.read_excel(file_path)

            # Check for expected column
            if 'prq2.partno' not in df.columns:
                logger.error("Ignore file missing 'prq2.partno' column. Found: %s", list(df.columns))
                return

            # Load part numbers (strip whitespace, convert to uppercase, remove empty)
            part_numbers = df['prq2.partno'].dropna().astype(str).str.strip().str.upper()
            self._ignored_part_numbers = set(part_numbers[part_numbers != ''])

            logger.info("Ignore list loaded: %d part numbers from %s",
                        len(self._ignored_part_numbers), os.path.basename(file_path))

        except Exception as exc:
            logger.error("Failed to load ignore file %s: %s", file_path, exc)

    @staticmethod
    def is_panel(description: str) -> bool:
        """
        Check if an item is a panel (auto-ignore rule).

        Args:
            description: Item description to check

        Returns:
            True if the item is a panel (should be ignored)
        """
        if not description:
            return False

        desc_lower = description.lower()

        # Panel keywords - these are structural access components, not consumables
        panel_keywords = [
            'panel',
            'panel assy',
            'panel assembly',
            'access panel',
            'inspection panel',
            'cowl panel',
            'door panel',
        ]

        return any(keyword in desc_lower for keyword in panel_keywords)

    def is_ignored_by_part_number(self, part_number: str) -> bool:
        """
        Check if a part number is in the manual ignore list.

        Args:
            part_number: Part number to check (e.g., "STD-123", "G00034")

        Returns:
            True if the part number should be ignored
        """
        if not part_number or not self._ignored_part_numbers:
            return False

        return part_number.strip().upper() in self._ignored_part_numbers

    def should_ignore_tool(self, reference_id: str, description: str,
                           part_numbers: List[str]) -> bool:
        """
        Check if a tool should be ignored.

        Args:
            reference_id: Tool reference ID (e.g., "STD-123", "SPL-8455")
            description: Tool description
            part_numbers: List of part number strings (may be empty for STD tools)

        Returns:
            True if the tool should be ignored
        """
        # Check auto-ignore: panels
        if self.is_panel(description):
            return True

        # Check manual-ignore: reference ID
        if self.is_ignored_by_part_number(reference_id):
            return True

        # Check manual-ignore: any part number in the list
        for pn_string in part_numbers:
            # Extract part number from formats like:
            # "(Part #: K32020-1, Supplier: ...)"
            # "(Opt Part #: K32012-1, Supplier: ...)"
            import re
            match = re.search(r'Part #:\s*([^,]+)', pn_string)
            if match:
                pn = match.group(1).strip()
                if self.is_ignored_by_part_number(pn):
                    return True

        return False

    def should_ignore_consumable(self, reference_id: str, description: str,
                                 specification: str) -> bool:
        """
        Check if a consumable should be ignored.

        Args:
            reference_id: Consumable reference ID (e.g., "G00034")
            description: Consumable description
            specification: Specification (e.g., "BMS3-33")

        Returns:
            True if the consumable should be ignored
        """
        # Check auto-ignore: panels
        if self.is_panel(description):
            return True

        # Check manual-ignore: reference ID
        if self.is_ignored_by_part_number(reference_id):
            return True

        return False

    def should_ignore_expendable(self, amm_item: str, part_description: str,
                                 part_number: str, ipd_figure_title: str) -> bool:
        """
        Check if an expendable should be ignored.

        Args:
            amm_item: AMM item text (e.g., "cotter pin [106]")
            part_description: IPD part description
            part_number: IPD part number
            ipd_figure_title: IPD figure title

        Returns:
            True if the expendable should be ignored
        """
        # Check auto-ignore: panels (check multiple fields)
        if (self.is_panel(amm_item) or
                self.is_panel(part_description) or
                self.is_panel(ipd_figure_title)):
            return True

        # Check manual-ignore: part number
        if part_number and self.is_ignored_by_part_number(part_number):
            return True

        return False

    def get_ignore_stats(self) -> dict:
        """Return statistics about the ignore list."""
        return {
            "manual_ignore_count": len(self._ignored_part_numbers),
            "auto_ignore_rules": ["Panel/Panel Assy (all types)"]
        }