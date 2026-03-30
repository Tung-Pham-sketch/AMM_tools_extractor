"""
AMM Tool Extractor
==================
Usage:
    1. Set PARENT_DMCS in config.py.
    2. Set STORAGE_DIR, IPD_ZIP_FILES, IPD_ZIP_INTERNAL_PREFIX in config.py.
    3. Set IGNORE_LIST_FILE in config.py (optional).
    4. Run:  python main.py
"""

import os
import re
import sys
import logging

import config
from extractor.zip_resolver         import ZipResolver
from extractor.ipd_resolver         import IpdResolver
from extractor.crawler              import crawl_with_visited
from extractor.stock_checker        import StockChecker
from extractor.tool_list_parser     import ToolListParser
from extractor.consumable_list_parser import ConsumableListParser
from extractor.ipd_parser           import enrich_expendables
from extractor.ignore_list          import IgnoreList
from writer.excel_writer            import write_report
from writer.tree_writer             import write_tree_html

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


def filter_tasks(all_tasks, ignore_list):
    """
    Filter out ignored items from all tasks.
    Modifies tasks in-place by removing ignored tools, consumables, and expendables.

    Returns statistics about what was filtered.
    """
    stats = {
        "tools_removed": 0,
        "consumables_removed": 0,
        "expendables_removed": 0,
        "tools_auto_ignored": 0,     # Panels
        "tools_manual_ignored": 0,   # From ignore list
        "consumables_auto_ignored": 0,
        "consumables_manual_ignored": 0,
        "expendables_auto_ignored": 0,
        "expendables_manual_ignored": 0,
    }

    for task in all_tasks:
        # ── Filter Tools ───────────────────────────────────────────────────
        filtered_tools = []
        for tool in task.tools:
            if ignore_list.should_ignore_tool(
                tool.reference_id,
                tool.description,
                tool.part_numbers
            ):
                stats["tools_removed"] += 1
                # Determine if auto or manual ignore
                if ignore_list.is_panel(tool.description):
                    stats["tools_auto_ignored"] += 1
                else:
                    stats["tools_manual_ignored"] += 1
            else:
                filtered_tools.append(tool)
        task.tools = filtered_tools

        # ── Filter Consumables ─────────────────────────────────────────────
        filtered_consumables = []
        for cons in task.consumables:
            if ignore_list.should_ignore_consumable(
                cons.reference_id,
                cons.description,
                cons.specification
            ):
                stats["consumables_removed"] += 1
                if ignore_list.is_panel(cons.description):
                    stats["consumables_auto_ignored"] += 1
                else:
                    stats["consumables_manual_ignored"] += 1
            else:
                filtered_consumables.append(cons)
        task.consumables = filtered_consumables

        # ── Filter Expendables ─────────────────────────────────────────────
        filtered_expendables = []
        for exp in task.expendables:
            if ignore_list.should_ignore_expendable(
                exp.amm_item,
                exp.part_description,
                exp.part_number,
                exp.ipd_figure_title
            ):
                stats["expendables_removed"] += 1
                if (ignore_list.is_panel(exp.amm_item) or
                    ignore_list.is_panel(exp.part_description) or
                    ignore_list.is_panel(exp.ipd_figure_title)):
                    stats["expendables_auto_ignored"] += 1
                else:
                    stats["expendables_manual_ignored"] += 1
            else:
                filtered_expendables.append(exp)
        task.expendables = filtered_expendables

    return stats


def main():
    logger.info("=" * 60)
    logger.info("AMM Tool Extractor")
    logger.info("=" * 60)

    # ── 1. Validate config ─────────────────────────────────────────────────────
    parent_dmcs = getattr(config, "PARENT_DMCS", [])
    if isinstance(parent_dmcs, str):
        parent_dmcs = [parent_dmcs]
    parent_dmcs = [d.strip() for d in parent_dmcs if d.strip()]
    if not parent_dmcs:
        logger.error("PARENT_DMCS is empty in config.py.")
        sys.exit(1)

    # ── 2. Build AMM zip index ─────────────────────────────────────────────────
    logger.info("Building AMM zip index (%d zip file(s))...", len(config.AMM_ZIP_FILES))
    resolver = ZipResolver(
        config.AMM_ZIP_FILES,
        use_latest=config.USE_LATEST_REVISION,
        internal_prefix=config.AMM_ZIP_INTERNAL_PREFIX,
    )

    # ── 3. Build IPD zip index ─────────────────────────────────────────────────
    ipd_zip_files = getattr(config, "IPD_ZIP_FILES", [])
    ipd_prefix    = getattr(config, "IPD_ZIP_INTERNAL_PREFIX", "")
    if ipd_zip_files:
        logger.info("Building IPD zip index (%d zip file(s))...", len(ipd_zip_files))
        ipd_resolver = IpdResolver(
            ipd_zip_files,
            use_latest=config.USE_LATEST_REVISION,
            internal_prefix=ipd_prefix,
        )
    else:
        logger.warning("IPD_ZIP_FILES not set — expendable parts will not be resolved.")
        ipd_resolver = None

    # ── 4. Load ignore list ────────────────────────────────────────────────────
    ignore_list_file = getattr(config, "IGNORE_LIST_FILE", "")
    logger.info("Loading ignore list...")
    ignore_list = IgnoreList(ignore_list_file)
    ignore_stats = ignore_list.get_ignore_stats()
    logger.info("  Manual ignore list: %d part numbers", ignore_stats["manual_ignore_count"])
    logger.info("  Auto-ignore rules: %s", ", ".join(ignore_stats["auto_ignore_rules"]))

    # ── 5. Load tool list ──────────────────────────────────────────────────────
    tool_list     = None
    tool_list_dmc = getattr(config, "TOOL_LIST_DMC", "").strip()
    if tool_list_dmc:
        logger.info("Loading tool list: %s", tool_list_dmc)
        tool_list_html = resolver.resolve(tool_list_dmc)
        if tool_list_html:
            tool_list = ToolListParser(tool_list_html)
            logger.info("Tool list loaded: %d entries.", len(tool_list))
        else:
            logger.warning("Tool list DMC not found in zips.")

    # ── 6. Load consumable list ────────────────────────────────────────────────
    consumable_list     = None
    consumable_list_dmc = getattr(config, "CONSUMABLE_LIST_DMC", "").strip()
    if consumable_list_dmc:
        logger.info("Loading consumable list: %s", consumable_list_dmc)
        cons_list_html = resolver.resolve(consumable_list_dmc)
        if cons_list_html:
            consumable_list = ConsumableListParser(cons_list_html)
            logger.info("Consumable list loaded: %d entries.", len(consumable_list))
        else:
            logger.warning("Consumable list DMC not found in zips.")

    # ── 7. Load storage / stock data ───────────────────────────────────────────
    storage_dir = getattr(config, "STORAGE_DIR", "")
    if storage_dir and os.path.isdir(storage_dir):
        logger.info("Loading storage files from: %s", storage_dir)
        stock_checker = StockChecker(storage_dir)
    else:
        logger.warning("STORAGE_DIR not set or not found — stock check skipped.")
        stock_checker = None

    # ── 8. Crawl AMM tasks ─────────────────────────────────────────────────────
    globally_visited: set = set()
    all_tasks  = []
    task_groups = []

    for parent_dmc in parent_dmcs:
        logger.info("-" * 60)
        logger.info("Processing: %s", parent_dmc)
        parent_html = resolver.resolve(parent_dmc)
        if parent_html is None:
            logger.error("Not found in any zip: %s — skipping.", parent_dmc)
            continue

        tasks = crawl_with_visited(
            parent_html=parent_html,
            parent_dmc=parent_dmc,
            resolver=resolver.resolve,
            max_depth=config.MAX_DEPTH,
            globally_visited=globally_visited,
        )
        if not tasks:
            logger.warning("No tasks returned for %s.", parent_dmc)
            continue

        parent_title = tasks[0].title
        task_groups.append((parent_dmc, parent_title, tasks))
        all_tasks.extend(tasks)
        logger.info("  %d tasks | %d tools | %d consumables | %d expendables",
                    len(tasks),
                    sum(len(t.tools) for t in tasks),
                    sum(len(t.consumables) for t in tasks),
                    sum(len(t.expendables) for t in tasks))

    if not all_tasks:
        logger.error("No tasks processed.")
        sys.exit(1)

    # ── 8b. Build output name & write HTML reference tree ─────────────────────
    def sanitise(s):
        return re.sub(r'[\\/*?:"<>|]', "", s).replace("–", "-").strip()

    if len(task_groups) == 1:
        dmc, title, _ = task_groups[0]
        safe_name = f"{dmc} - {sanitise(title)}"
    else:
        safe_name = " & ".join(sanitise(t) for _, t, _ in task_groups)

    tree_path = os.path.join(config.OUTPUT_DIR, f"{safe_name} - Reference Tree.html")
    write_tree_html(task_groups, all_tasks, tree_path, max_depth=config.MAX_DEPTH)
    logger.info("Reference tree saved to: %s", tree_path)

    # ── 9. Resolve IPD expendables ─────────────────────────────────────────────
    if ipd_resolver:
        logger.info("Resolving expendable parts from IPD...")
        enrich_expendables(all_tasks, ipd_resolver)

    # ── 10. Apply ignore filters ───────────────────────────────────────────────
    logger.info("Applying ignore filters...")
    filter_stats = filter_tasks(all_tasks, ignore_list)

    logger.info("  Items removed:")
    logger.info("    Tools:       %d (Auto: %d, Manual: %d)",
               filter_stats["tools_removed"],
               filter_stats["tools_auto_ignored"],
               filter_stats["tools_manual_ignored"])
    logger.info("    Consumables: %d (Auto: %d, Manual: %d)",
               filter_stats["consumables_removed"],
               filter_stats["consumables_auto_ignored"],
               filter_stats["consumables_manual_ignored"])
    logger.info("    Expendables: %d (Auto: %d, Manual: %d)",
               filter_stats["expendables_removed"],
               filter_stats["expendables_auto_ignored"],
               filter_stats["expendables_manual_ignored"])

    # ── 11. Totals log ─────────────────────────────────────────────────────────
    logger.info("=" * 60)
    logger.info("Final counts after filtering:")
    logger.info("Tasks: %d | Tools: %d | Consumables: %d | Expendables: %d",
                len(all_tasks),
                sum(len(t.tools) for t in all_tasks),
                sum(len(t.consumables) for t in all_tasks),
                sum(len(t.expendables) for t in all_tasks))

    # ── 12. Output filename ────────────────────────────────────────────────────
    output_path = os.path.join(config.OUTPUT_DIR, f"{safe_name} - Tools Report.xlsx")

    # ── 13. Write Excel ────────────────────────────────────────────────────────
    write_report(
        all_tasks, output_path,
        task_groups=task_groups,
        stock_checker=stock_checker,
        tool_list=tool_list,
        consumable_list=consumable_list,
    )
    logger.info("Report saved to: %s", output_path)
    logger.info("Done.")


if __name__ == "__main__":
    main()