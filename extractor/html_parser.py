"""
html_parser.py
==============
Parses a single Boeing AMM task HTM file and returns:
  - task metadata  (DMC code, title)
  - references     (list of DMC codes from the References table)
  - support tools  (list of ToolEntry)
  - consumables    (list of ConsumableEntry)
  - expendables    (list of ExpendableEntry — IPD lookup deferred)
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional
from urllib.parse import unquote
from bs4 import BeautifulSoup


# ── Data classes ───────────────────────────────────────────────────────────────

@dataclass
class ToolEntry:
    reference_id: str
    description: str
    part_numbers: List[str]
    source_dmc: str
    source_title: str


@dataclass
class ConsumableEntry:
    reference_id: str       # e.g. G01048
    description: str
    specification: str      # e.g. BMS3-33
    source_dmc: str
    source_title: str


@dataclass
class ExpendableEntry:
    amm_item: str           # e.g. "cotter pin [106]"
    ipd_figure_title: str   # e.g. "BUILDUP ASSY-MLG (...)"
    ipd_dmc: str            # derived DMC e.g. "B787-A-32-11-01-010-941A-D"
    ipd_item_number: str    # e.g. "150"
    # Filled in later by IPD lookup:
    part_number: str = ""
    part_description: str = ""
    quantity: str = ""
    source_dmc: str = ""
    source_title: str = ""


@dataclass
class TaskData:
    dmc: str
    title: str
    references: List[str]
    tools: List[ToolEntry]
    consumables: List[ConsumableEntry] = field(default_factory=list)
    expendables: List[ExpendableEntry] = field(default_factory=list)


# ── Regex helpers ──────────────────────────────────────────────────────────────

_URN_AMM_RE  = re.compile(r"linkToUrn\('URN:S1000D:DME-\d+-\w+-(B787-[^']+)'")
_URN_IPD_RE  = re.compile(r"linkToUrn\('URN:X-BOEING:DMC[^']*'")
_PARAM_RE    = re.compile(r'(\w+)=([^&\']+)')
_ITEM_RE     = re.compile(r'Item\s+(\d+)', re.IGNORECASE)


def _ipd_dmc_from_params(params: dict) -> str:
    """
    Build IPD DMC code from URN query parameters.
    e.g. CH=32,SE=11,SU=01,DC=01,DCV=0 -> B787-A-32-11-01-010-941A-D
    DC+DCV are concatenated and zero-padded to 3 digits.
    """
    ch  = params.get('CH', '')
    se  = params.get('SE', '')
    su  = params.get('SU', '')
    dc  = params.get('DC', '')
    dcv = params.get('DCV', '0')
    ic  = params.get('IC', '941')
    icv = params.get('ICV', 'A')
    loc = params.get('LOC', 'D')
    dc_full = f"{dc}{dcv}"   # e.g. "01"+"0" = "010"
    return f"B787-A-{ch}-{se}-{su}-{dc_full}-{ic}{icv}-{loc}"


# ── Main parser ────────────────────────────────────────────────────────────────

def parse_task(html_content: str, source_dmc: str) -> TaskData:
    soup = BeautifulSoup(html_content, "lxml")

    # Title
    title_tag = soup.find("title")
    title = title_tag.get_text(" ", strip=True) if title_tag else source_dmc

    # ── References ─────────────────────────────────────────────────────────────
    references: List[str] = []
    ref_heading = soup.find(lambda t: t.name in ("h3", "h4") and
                            "References" in t.get_text())
    if ref_heading:
        ref_table = ref_heading.find_next("table")
        if ref_table:
            for a_tag in ref_table.find_all("a", attrs={"ng-click": True}):
                m = _URN_AMM_RE.search(a_tag["ng-click"])
                if m:
                    dmc = m.group(1)
                    if dmc not in references:
                        references.append(dmc)

    # ── Support Equipment ──────────────────────────────────────────────────────
    tools: List[ToolEntry] = []
    supp_heading = soup.find(lambda t: t.name in ("h3", "h4") and
                             "Support Equipment" in t.get_text())
    if supp_heading:
        supp_table = supp_heading.find_next("table")
        if supp_table:
            for row in supp_table.find_all("tr"):
                row_id = row.get("id", "")
                if not re.match(r"^(SPL|STD)-", row_id):
                    continue
                cells = row.find_all("td")
                if len(cells) < 2:
                    continue
                ref_id = cells[0].get_text(strip=True)
                desc_cell = cells[1]
                span = desc_cell.find("span")
                description = span.get_text(strip=True) if span else desc_cell.get_text(strip=True)
                part_numbers = [d.get_text(strip=True)
                                for d in desc_cell.find_all("div", class_="line")
                                if d.get_text(strip=True)]
                tools.append(ToolEntry(
                    reference_id=ref_id,
                    description=description,
                    part_numbers=part_numbers,
                    source_dmc=source_dmc,
                    source_title=title,
                ))

    # ── Consumable Materials ───────────────────────────────────────────────────
    consumables: List[ConsumableEntry] = []
    cons_heading = soup.find(lambda t: t.name in ("h3", "h4") and
                             "Consumable" in t.get_text())
    if cons_heading:
        cons_table = cons_heading.find_next("table")
        if cons_table:
            for row in cons_table.find_all("tr"):
                cells = [c.get_text(strip=True) for c in row.find_all(["th", "td"])]
                if len(cells) < 2:
                    continue
                if cells[0].lower() == "reference":
                    continue
                if not cells[0]:
                    continue
                consumables.append(ConsumableEntry(
                    reference_id=cells[0],
                    description=cells[1] if len(cells) > 1 else "",
                    specification=cells[2] if len(cells) > 2 else "",
                    source_dmc=source_dmc,
                    source_title=title,
                ))

    # ── Expendables/Parts ──────────────────────────────────────────────────────
    expendables: List[ExpendableEntry] = []
    exp_heading = soup.find(lambda t: t.name in ("h3", "h4") and
                            "Expendable" in t.get_text())
    if exp_heading:
        exp_table = exp_heading.find_next("table")
        if exp_table:
            for row in exp_table.find_all("tr"):
                cells = row.find_all(["th", "td"])
                if len(cells) < 2:
                    continue
                amm_text = cells[0].get_text(strip=True)
                if not amm_text or amm_text.lower().startswith("amm"):
                    continue
                # Skip applicability-only rows (no link)
                link = cells[1].find("a", attrs={"ng-click": True})
                if not link:
                    continue
                ipd_full_text = cells[1].get_text(strip=True)
                # Decode URN and extract params
                decoded = unquote(link["ng-click"])
                params = dict(_PARAM_RE.findall(decoded))
                if not params:
                    continue
                ipd_dmc = _ipd_dmc_from_params(params)
                # Extract item number from display text
                m = _ITEM_RE.search(ipd_full_text)
                item_number = m.group(1) if m else ""
                # Extract figure title (text before the parenthesised IPD ref)
                fig_title = re.sub(r'\(IPD.*', '', ipd_full_text).strip()
                expendables.append(ExpendableEntry(
                    amm_item=amm_text,
                    ipd_figure_title=fig_title,
                    ipd_dmc=ipd_dmc,
                    ipd_item_number=item_number,
                    source_dmc=source_dmc,
                    source_title=title,
                ))

    return TaskData(
        dmc=source_dmc,
        title=title,
        references=references,
        tools=tools,
        consumables=consumables,
        expendables=expendables,
    )
