"""
tree_writer.py
==============
Generates a self-contained HTML file visualising the AMM reference
hierarchy as a compact horizontal tree with a zoom slider and pan support.

Depth is capped at config.MAX_DEPTH — matching exactly what the crawler
collected — so the tree never shows more levels than were processed.
"""

import os
import json
from typing import List, Dict, Set, Optional
from amm_tool_extractor.extractor.html_parser import TaskData

NOT_FOUND = "NOT FOUND IN ZIP"


def _build_map(all_tasks: List[TaskData]) -> Dict[str, TaskData]:
    return {t.dmc: t for t in all_tasks}


def _build_tree_dict(
    dmc: str,
    task_map: Dict[str, TaskData],
    shown: Set[str],
    depth: int,
    max_depth: Optional[int],
) -> dict:
    task  = task_map.get(dmc)
    title = task.title if task else NOT_FOUND

    if dmc in shown:
        return {"dmc": dmc, "title": title, "already_shown": True,
                "depth": depth, "children": []}

    shown.add(dmc)

    # Stop expanding children once we hit the configured max depth
    children = []
    if task and (max_depth is None or depth < max_depth):
        for ref in task.references:
            children.append(
                _build_tree_dict(ref, task_map, shown, depth + 1, max_depth)
            )

    return {"dmc": dmc, "title": title, "already_shown": False,
            "depth": depth, "children": children}


def _trees_to_json(task_groups, all_tasks, max_depth):
    task_map = _build_map(all_tasks)
    roots = []
    for parent_dmc, parent_title, _ in task_groups:
        shown: Set[str] = set()
        roots.append(
            _build_tree_dict(parent_dmc, task_map, shown, 0, max_depth)
        )
    return json.dumps(roots, ensure_ascii=False)


_HTML = """\
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>AMM Reference Tree</title>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

  body {{
    font-family: 'Courier New', Courier, monospace;
    background: #d8dde3;
    min-height: 100vh;
    padding: 20px;
    color: #1a1a2e;
  }}

  header {{
    display: flex;
    align-items: center;
    gap: 24px;
    margin-bottom: 20px;
    padding-bottom: 12px;
    border-bottom: 2px solid #aaa;
  }}

  header h1 {{
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    color: #444;
    flex: 1;
  }}

  .zoom-ctrl {{
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 10px;
    color: #555;
    letter-spacing: 0.05em;
  }}

  .zoom-ctrl input[type=range] {{
    width: 100px;
    accent-color: #1a5fa8;
  }}

  .zoom-label {{
    min-width: 34px;
    font-size: 10px;
    color: #333;
    font-weight: 700;
  }}

  .tree-section {{
    margin-bottom: 28px;
    background: #e8ecf0;
    border: 1px solid #c0c6cd;
    border-radius: 3px;
    padding: 14px 14px 18px;
  }}

  .section-title {{
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #1a5fa8;
    margin-bottom: 14px;
  }}

  .tree-viewport {{
    overflow: auto;
    width: 100%;
    cursor: grab;
    /* height is set by JS after measuring the canvas */
  }}
  .tree-viewport:active {{ cursor: grabbing; }}

  /* Outer sizer shrinks with the scale so the viewport has correct scrollable area */
  .tree-sizer {{
    display: inline-block;
    transform-origin: top left;
  }}

  .tree-canvas {{
    display: inline-flex;
    padding: 16px 40px 24px 16px;
  }}

  .node {{
    display: flex;
    flex-direction: column;
    justify-content: center;
    width: 168px;
    min-height: 40px;
    padding: 5px 8px;
    border: 2px solid #1a5fa8;
    background: #fff;
    text-align: center;
    flex-shrink: 0;
    margin: 6px 0;
  }}

  .node .dmc {{
    font-size: 7.5px;
    letter-spacing: 0.04em;
    color: #999;
    margin-bottom: 2px;
    word-break: break-all;
    line-height: 1.2;
  }}

  .node .lbl {{
    font-size: 9px;
    font-weight: 700;
    line-height: 1.3;
    color: #1a1a2e;
  }}

  .d0 {{ border-color: #1a5fa8; }}
  .d1 {{ border-color: #e07b2a; }}
  .d2 {{ border-color: #27a86e; }}
  .d3 {{ border-color: #9b3db8; }}
  .d4 {{ border-color: #b5860d; }}
  .d5 {{ border-color: #1a5fa8; }}
  .d6 {{ border-color: #e07b2a; }}
  .d7 {{ border-color: #27a86e; }}
  .d8 {{ border-color: #9b3db8; }}
  .d9 {{ border-color: #b5860d; }}

  .node.nf {{ border-color: #c0392b; background: #fff5f5; }}
  .node.nf .lbl {{ color: #c0392b; font-size: 8px; }}
  .node.as {{ border-style: dashed; opacity: 0.6; }}
  .node.as .lbl {{ font-style: italic; }}

  .h-line {{ width: 22px; height: 2px; background: #555; flex-shrink: 0; }}
  .stub   {{ width: 22px; height: 2px; background: #555; flex-shrink: 0; }}

  .bracket {{
    display: flex;
    flex-direction: column;
    align-items: flex-start;
    position: relative;
    flex-shrink: 0;
  }}

  .child-row {{
    display: flex;
    flex-direction: row;
    align-items: center;
  }}
</style>
</head>
<body>

<header>
  <h1>AMM Reference Tree &nbsp;<span style="font-weight:400;color:#888;">(max depth: {max_depth})</span></h1>
  <div class="zoom-ctrl">
    ZOOM
    <input type="range" id="zoomSlider" min="20" max="150" value="65">
    <span class="zoom-label" id="zoomLabel">65%</span>
  </div>
</header>

<div id="app"></div>

<script>
const TREES = {trees_json};

function makeNode(item) {{
  const div = document.createElement('div');
  const depth = Math.min(item.depth, 9);
  let cls = `node d${{depth}}`;
  if (item.title === 'NOT FOUND IN ZIP') cls += ' nf';
  if (item.already_shown) cls += ' as';
  div.className = cls;

  const dmc = document.createElement('div');
  dmc.className = 'dmc';
  dmc.textContent = item.dmc;
  div.appendChild(dmc);

  const lbl = document.createElement('div');
  lbl.className = 'lbl';
  lbl.textContent = item.already_shown
    ? '\\u21a9 ' + item.title
    : item.title === 'NOT FOUND IN ZIP' ? 'NOT FOUND' : item.title;
  div.appendChild(lbl);
  return div;
}}

function buildTree(item) {{
  const row = document.createElement('div');
  row.className = 'child-row';
  row.appendChild(makeNode(item));

  if (item.children && item.children.length > 0) {{
    const hline = document.createElement('div');
    hline.className = 'h-line';
    row.appendChild(hline);

    const bracket = document.createElement('div');
    bracket.className = 'bracket';

    item.children.forEach(child => {{
      const childRow = document.createElement('div');
      childRow.className = 'child-row';
      const stub = document.createElement('div');
      stub.className = 'stub';
      childRow.appendChild(stub);
      childRow.appendChild(buildTree(child));
      bracket.appendChild(childRow);
    }});

    row.appendChild(bracket);

    requestAnimationFrame(() => {{
      const kids = Array.from(bracket.children);
      if (kids.length < 2) return;
      const bRect  = bracket.getBoundingClientRect();
      const top    = kids[0].getBoundingClientRect().top    + kids[0].getBoundingClientRect().height / 2 - bRect.top;
      const bottom = kids[kids.length-1].getBoundingClientRect().top + kids[kids.length-1].getBoundingClientRect().height / 2 - bRect.top;
      const bar = document.createElement('div');
      bar.style.cssText = `position:absolute;left:0;top:${{top}}px;width:2px;height:${{bottom-top}}px;background:#555;pointer-events:none;`;
      bracket.appendChild(bar);
    }});
  }}
  return row;
}}

const app = document.getElementById('app');

TREES.forEach(root => {{
  const section = document.createElement('div');
  section.className = 'tree-section';

  const title = document.createElement('div');
  title.className = 'section-title';
  title.textContent = root.dmc + '  \u2014  ' + root.title;
  section.appendChild(title);

  const viewport = document.createElement('div');
  viewport.className = 'tree-viewport';

  const sizer = document.createElement('div');
  sizer.className = 'tree-sizer';

  const canvas = document.createElement('div');
  canvas.className = 'tree-canvas';
  canvas.appendChild(buildTree(root));
  sizer.appendChild(canvas);
  viewport.appendChild(sizer);
  section.appendChild(viewport);
  app.appendChild(section);
}});

// ── Zoom ────────────────────────────────────────────────────────────────────
const slider = document.getElementById('zoomSlider');
const label  = document.getElementById('zoomLabel');

function applyZoom(z) {{
  label.textContent = z + '%';
  const scale = z / 100;
  // Apply scale to the sizer wrapper, then resize it so the viewport
  // scrollable area matches the scaled content exactly — no clipping.
  document.querySelectorAll('.tree-sizer').forEach(sizer => {{
    const canvas = sizer.querySelector('.tree-canvas');
    // Reset scale to measure natural size
    sizer.style.transform = '';
    sizer.style.width  = '';
    sizer.style.height = '';
    const w = canvas.scrollWidth;
    const h = canvas.scrollHeight;
    // Apply scale via transform, then shrink the sizer to scaled dims
    sizer.style.transform = `scale(${{scale}})`;
    sizer.style.transformOrigin = 'top left';
    sizer.style.width  = (w * scale) + 'px';
    sizer.style.height = (h * scale) + 'px';
  }});
}}

slider.addEventListener('input', () => applyZoom(parseInt(slider.value)));
requestAnimationFrame(() => requestAnimationFrame(() => applyZoom(parseInt(slider.value))));

// ── Pan to scroll ────────────────────────────────────────────────────────────
document.querySelectorAll('.tree-viewport').forEach(vp => {{
  let isDown = false, startX, startY, scrollLeft, scrollTop;
  vp.addEventListener('mousedown', e => {{
    isDown = true;
    startX = e.pageX - vp.offsetLeft;
    startY = e.pageY - vp.offsetTop;
    scrollLeft = vp.scrollLeft;
    scrollTop  = vp.scrollTop;
  }});
  vp.addEventListener('mouseleave', () => isDown = false);
  vp.addEventListener('mouseup',    () => isDown = false);
  vp.addEventListener('mousemove',  e => {{
    if (!isDown) return;
    e.preventDefault();
    vp.scrollLeft = scrollLeft - (e.pageX - vp.offsetLeft - startX);
    vp.scrollTop  = scrollTop  - (e.pageY - vp.offsetTop  - startY);
  }});
}});
</script>
</body>
</html>
"""


def write_tree_html(
    task_groups: list,
    all_tasks: List[TaskData],
    output_path: str,
    max_depth: Optional[int] = None,
) -> None:
    """
    Generate a self-contained HTML tree file.

    Args:
        task_groups: [(parent_dmc, parent_title, tasks), ...]
        all_tasks:   flat list of all TaskData objects
        output_path: full path for the .html output file
        max_depth:   from config.MAX_DEPTH — caps tree expansion
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    trees_json = _trees_to_json(task_groups, all_tasks, max_depth)
    depth_label = str(max_depth) if max_depth is not None else "unlimited"
    html = _HTML.format(trees_json=trees_json, max_depth=depth_label)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)