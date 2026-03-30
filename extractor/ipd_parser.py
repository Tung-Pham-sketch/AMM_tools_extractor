"""
ipd_parser.py
=============
Given an IPD JSON dict and an item number, returns part details.
Also provides enrich_expendables() to resolve all ExpendableEntry
objects in a task list using the IpdResolver.
"""

import logging
from typing import Optional, List
from .html_parser import ExpendableEntry, TaskData

logger = logging.getLogger(__name__)


def lookup_item(ipd_data: dict, item_number: str) -> Optional[dict]:
    """
    Find a part by itemNumber in the IPD parts array.
    item_number may be zero-padded (e.g. '078') or plain ('78').
    Returns dict with keys: part_number, description, quantity
    """
    target = str(item_number).lstrip("0") or "0"
    for part in ipd_data.get("parts", []):
        pn_item = str(part.get("itemNumber", "")).lstrip("0") or "0"
        if pn_item == target:
            return {
                "part_number":  part.get("pnr", ""),
                "description":  part.get("dfp", ""),
                "quantity":     str(part.get("quantityPerAssembly", "")),
            }
    return None


def enrich_expendables(all_tasks: List[TaskData], ipd_resolver) -> None:
    """
    Walk all ExpendableEntry objects across all tasks and fill in
    part_number, part_description, quantity via IPD lookup.
    Modifies entries in-place.
    """
    # Cache IPD JSON to avoid re-reading the same zip entry repeatedly
    ipd_cache: dict = {}

    for task in all_tasks:
        for exp in task.expendables:
            dmc = exp.ipd_dmc
            if dmc not in ipd_cache:
                ipd_cache[dmc] = ipd_resolver.resolve(dmc)

            ipd_data = ipd_cache[dmc]
            if ipd_data is None:
                logger.warning("IPD not found: %s (ref'd by %s)", dmc, task.dmc)
                continue

            result = lookup_item(ipd_data, exp.ipd_item_number)
            if result:
                exp.part_number      = result["part_number"]
                exp.part_description = result["description"]
                exp.quantity         = result["quantity"]
            else:
                logger.warning("Item %s not found in IPD %s",
                               exp.ipd_item_number, dmc)