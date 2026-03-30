"""
crawler.py
==========
BFS crawler that:
  1. Parses the parent task
  2. Follows References recursively up to MAX_DEPTH
  3. Returns all collected TaskData objects (visited set prevents loops)
  4. Records missing references as placeholder TaskData so they appear
     in the Excel report with a clear "NOT FOUND" marker.
"""

from collections import deque
from typing import List, Optional, Callable, Set
import logging

from .html_parser import parse_task, TaskData, ToolEntry

logger = logging.getLogger(__name__)

NOT_FOUND_MARKER = "NOT FOUND IN ZIP"


def _make_missing_task(dmc: str, referenced_by: str) -> TaskData:
    placeholder_tool = ToolEntry(
        reference_id="—",
        description=f"Task not found in zip files (referenced by: {referenced_by})",
        part_numbers=[],
        source_dmc=dmc,
        source_title=NOT_FOUND_MARKER,
    )
    return TaskData(
        dmc=dmc,
        title=NOT_FOUND_MARKER,
        references=[],
        tools=[placeholder_tool],
    )


def crawl(
    parent_html: str,
    parent_dmc: str,
    resolver: Callable[[str], Optional[str]],
    max_depth: Optional[int] = None,
) -> List[TaskData]:
    """Convenience wrapper — uses its own fresh visited set."""
    return crawl_with_visited(
        parent_html=parent_html,
        parent_dmc=parent_dmc,
        resolver=resolver,
        max_depth=max_depth,
        globally_visited=set(),
    )


def crawl_with_visited(
    parent_html: str,
    parent_dmc: str,
    resolver: Callable[[str], Optional[str]],
    max_depth: Optional[int] = None,
    globally_visited: Optional[Set[str]] = None,
) -> List[TaskData]:
    """
    BFS crawl starting from parent_dmc.

    globally_visited is shared across multiple calls so that tasks already
    processed by a previous parent are not duplicated in the combined report.
    """
    if globally_visited is None:
        globally_visited = set()

    results: List[TaskData] = []
    # Local queue only — visited check uses the shared global set
    queue: deque = deque()
    queue.append((parent_html, parent_dmc, 0, "root"))

    while queue:
        html, dmc, depth, referenced_by = queue.popleft()

        if dmc in globally_visited:
            continue
        globally_visited.add(dmc)

        if html is None:
            logger.warning("  NOT FOUND in zips [depth=%d]: %s  (ref'd by: %s)",
                           depth, dmc, referenced_by)
            results.append(_make_missing_task(dmc, referenced_by))
            continue

        logger.info("Parsing [depth=%d] %s", depth, dmc)
        task = parse_task(html, dmc)
        results.append(task)

        if max_depth is not None and depth >= max_depth:
            continue

        for ref_dmc in task.references:
            if ref_dmc in globally_visited:
                continue
            ref_html = resolver(ref_dmc)
            queue.append((ref_html, ref_dmc, depth + 1, dmc))

    return results