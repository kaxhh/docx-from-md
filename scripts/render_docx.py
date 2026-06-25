#!/usr/bin/env python3
"""
render_docx.py — JSON → DOCX renderer for the detailed design spec template.

Template: AI+工业操作系统研发项目详细设计规格说明书.docx
Input:    content_data.json (produced by md_to_json.py)

Uses python-docx to clone the template's page setup, headers, footers,
styles, and numbering, then appends generated content. Tables are rendered
with one function per table type. Formatting (fonts, shading, borders,
cell margins, row heights) is derived from the template's table_catalog.json
so it stays in sync when the template changes.

Usage:
    python3 render_docx.py \
        --template template.docx \
        --input content_data.json \
        --table-catalog analysis/table_catalog.json \
        --output generated.docx

Limitations:
    - Cover page and TOC from the template are preserved; generated content
      is appended after the existing body. The template's sample content
      (chapter 4 "示例" etc.) is not auto-removed — either strip it from
      the template beforehand or accept that the output contains both.
    - Figures are rendered as Word picture placeholders referencing the
      source path; actual image file embedding is supported only for PNG/JPG
      files present on disk. EMF / OLE / Visio objects are not embedded —
      the figure placeholder shows the caption and source path.
    - Flowchart markers (<!-- FLOWCHART: name -->) are rendered as
      "[流程图: name]" text placeholders.
"""

import argparse
import copy
import json
import re
import sys
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from docx import Document
    from docx.shared import Pt, Cm, Twips, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError as e:
    sys.exit(f"python-docx is required: {e}")

# ---------------------------------------------------------------------------
# Template-derived formatting constants.
#
# These are read from table_catalog.json at runtime, so the renderer
# stays correct if the template's widths or headers change.
# ---------------------------------------------------------------------------

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Style IDs used by the template (verified by inspection).
HEADING_STYLE_IDS = {1: "2", 2: "3", 3: "4", 4: "5", 5: "6"}
TABLE_CELL_STYLE_ID = "94"  # "表格内文字"
CAPTION_STYLE_ID = "14"  # "Caption"
CODE_STYLE_ID = "147"  # "代码"
BODY_STYLE_ID = None  # None = Normal

# Default header fill color (catalog shows D7D7D7 for simple tables and
# D9D9D9 for interface_spec_table — we use D9D9D9 as the common fallback,
# but the actual value is read from the catalog if available).
DEFAULT_HEADER_FILL = "D9D9D9"

# Flowchart marker regex.
FLOWCHART_RE = re.compile(r"^\s*<!--\s*FLOWCHART:\s*(.+?)\s*-->\s*$")


# ---------------------------------------------------------------------------
# Catalog loader.
# ---------------------------------------------------------------------------


def load_catalog(catalog_path):
    """Load table_catalog.json and extract per-family metadata."""
    if not catalog_path:
        return {}
    data = json.loads(Path(catalog_path).read_text(encoding="utf-8"))
    families = {}
    for fam in data.get("summary", []):
        ttype = fam["type"]
        # Find the matching table in the full tables array.
        ex = fam["examples"][0]
        full_table = next(
            (t for t in data.get("tables", []) if t["id"] == ex["id"]), {}
        )
        header_labels = full_table.get("header") or [
            cell["text"] for cell in full_table.get("rows", [{}])[0].get("cells", [])
        ]
        families[ttype] = {
            "col_widths": fam["canonical_columns"],
            "header_labels": header_labels,
            "has_merges": ex.get("has_merges", False),
        }
    # Read header fill from the first body cell with shading.
    header_fill = DEFAULT_HEADER_FILL
    for t in data.get("tables", []):
        for row in t.get("rows", []):
            for cell in row.get("cells", []):
                shd = cell.get("shading")
                if shd and shd != "auto":
                    header_fill = shd
                    break
            if header_fill != DEFAULT_HEADER_FILL:
                break
        if header_fill != DEFAULT_HEADER_FILL:
            break
    families["_header_fill"] = header_fill
    return families


# ---------------------------------------------------------------------------
# Table rendering — one function per table type.
# ---------------------------------------------------------------------------


def _set_cell_shading(cell, fill_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)


def _set_cell_border_width(cell, sz="4"):
    """Ensure the cell has explicit borders with the given sz."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for side in ("top", "left", "bottom", "right"):
        elem = borders.find(qn(f"w:{side}"))
        if elem is None:
            elem = OxmlElement(f"w:{side}")
            borders.append(elem)
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:color"), "auto")
        elem.set(qn("w:sz"), sz)
        elem.set(qn("w:space"), "0")


def _set_cell_width(cell, dxa):
    tc_pr = cell._tc.get_or_add_tcPr()
    tcW = tc_pr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tc_pr.append(tcW)
    tcW.set(qn("w:w"), str(dxa))
    tcW.set(qn("w:type"), "dxa")


def _set_cell_vertical_align(cell, align="center"):
    tc_pr = cell._tc.get_or_add_tcPr()
    vAlign = tc_pr.find(qn("w:vAlign"))
    if vAlign is None:
        vAlign = OxmlElement("w:vAlign")
        tc_pr.append(vAlign)
    vAlign.set(qn("w:val"), align)


def _set_grid_span(cell, span):
    if span <= 1:
        return
    tc_pr = cell._tc.get_or_add_tcPr()
    gs = tc_pr.find(qn("w:gridSpan"))
    if gs is None:
        gs = OxmlElement("w:gridSpan")
        tc_pr.append(gs)
    gs.set(qn("w:val"), str(span))


def _set_v_merge(cell, val):
    """Set vMerge on a cell. val is 'restart' or 'continue'."""
    tc_pr = cell._tc.get_or_add_tcPr()
    vm = tc_pr.find(qn("w:vMerge"))
    if vm is None:
        vm = OxmlElement("w:vMerge")
        tc_pr.append(vm)
    if val == "restart":
        vm.set(qn("w:val"), "restart")
    # 'continue' is represented by an empty w:vMerge element (no w:val).


def _remove_grid_span_cells(row, current_idx, span):
    """Remove the (span - 1) physical cells right after the current one.

    python-docx creates all physical cells upfront. Setting w:gridSpan makes
    one cell span multiple grid columns, but the covered <w:tc> elements
    still exist and create redundant empty cells.

    row.cells is a cached @property — after the first XML removal its
    indices are stale. Collect all elements to remove in one pass, then
    remove them.
    """
    if span <= 1:
        return
    tcs = list(row._tr.iterchildren(qn("w:tc")))
    to_remove = tcs[current_idx + 1 : current_idx + span]
    for tc in to_remove:
        row._tr.remove(tc)


def _set_table_grid(table, col_widths):
    """Replace the table's <w:tblGrid> with the given column widths."""
    tbl = table._tbl
    tblGrid = tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        tbl.remove(tblGrid)
    tblGrid = OxmlElement("w:tblGrid")
    for w in col_widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        tblGrid.append(gc)
    # Insert after tblPr.
    tblPr = tbl.find(qn("w:tblPr"))
    tblPr.addnext(tblGrid)


def _set_table_borders(table, sz="4"):
    """Apply uniform single borders of the given sz to the table."""
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        elem = borders.find(qn(f"w:{side}"))
        if elem is None:
            elem = OxmlElement(f"w:{side}")
            borders.append(elem)
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:color"), "auto")
        elem.set(qn("w:sz"), sz)
        elem.set(qn("w:space"), "0")


def _set_table_cell_margins(table, left=108, right=108, top=0, bottom=0):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    cm = tblPr.find(qn("w:tblCellMar"))
    if cm is None:
        cm = OxmlElement("w:tblCellMar")
        tblPr.append(cm)
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        elem = cm.find(qn(f"w:{name}"))
        if elem is None:
            elem = OxmlElement(f"w:{name}")
            cm.append(elem)
        elem.set(qn("w:w"), str(val))
        elem.set(qn("w:type"), "dxa")


def _set_table_layout_fixed(table):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")


def _write_cell_text(cell, text, style_id):
    """Replace cell text, splitting newlines into multiple paragraphs.

    Applies the given paragraph style to every paragraph in the cell.
    """
    lines = (text or "").split("\n")
    if not lines:
        lines = [""]
    # python-docx cells always start with one empty paragraph.
    # Reuse it for the first line, add new paragraphs for subsequent lines.
    first_p = cell.paragraphs[0]
    if style_id:
        first_p.style = style_id
    first_p.text = lines[0]
    for line in lines[1:]:
        p = cell.add_paragraph()
        if style_id:
            p.style = style_id
        p.text = line


def render_simple_table(doc, table_data, family_meta, header_fill):
    """Render macro_definition_table / struct_member_table / global_variable_table."""
    col_widths = family_meta["col_widths"]
    header_labels = family_meta["header_labels"]
    rows = table_data.get("rows", [])

    table = doc.add_table(rows=1 + len(rows), cols=len(col_widths))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_grid(table, col_widths)
    _set_table_borders(table, sz="4")
    _set_table_cell_margins(table)
    _set_table_layout_fixed(table)

    # Header row.
    hdr = table.rows[0]
    for i, label in enumerate(header_labels):
        cell = hdr.cells[i]
        _set_cell_width(cell, col_widths[i])
        _set_cell_shading(cell, header_fill)
        _set_cell_vertical_align(cell, "center")
        _set_cell_border_width(cell, "4")
        _write_cell_text(cell, label, TABLE_CELL_STYLE_ID)

    # Body rows — column keys come from the JSON row dicts (which are
    # produced by md_to_json.py in the same order as header_labels).
    keys = list(rows[0].keys()) if rows else []
    for r_idx, row in enumerate(rows):
        tr = table.rows[r_idx + 1]
        for c_idx in range(len(col_widths)):
            cell = tr.cells[c_idx]
            _set_cell_width(cell, col_widths[c_idx])
            _set_cell_shading(cell, "auto")
            _set_cell_vertical_align(cell, "center")
            _set_cell_border_width(cell, "4")
            key = keys[c_idx] if c_idx < len(keys) else None
            text = row.get(key, "") if key else ""
            _write_cell_text(cell, text, TABLE_CELL_STYLE_ID)


def render_interface_spec_table(doc, table_data, family_meta, header_fill):
    """Render interface_spec_table with vMerge label groups."""
    col_widths = family_meta["col_widths"]  # [1378, 1984, 1276, 4068]
    total_width = sum(col_widths)
    spec = table_data.get("rows", {})
    if not isinstance(spec, dict):
        raise SystemExit(
            f"ERROR: interface_spec_table rows must be an object, got {type(spec).__name__}"
        )

    # Build row plan: each entry is (cells, merge_info).
    # cells: list of (text, gridSpan, is_header).
    # merge_info: None | ("param_header", n) | "param_data"
    #           | ("return_header", n) | "return_data".
    row_plan = []
    row_plan.append(
        (
            [("接口原型", 1, True), (spec.get("prototype", ""), 3, False)],
            None,
        )
    )
    row_plan.append(
        (
            [("接口描述", 1, True), (spec.get("description", ""), 3, False)],
            None,
        )
    )

    # Parameters group.
    params = spec.get("parameters", [])
    if isinstance(params, str):
        # Scalar "无。" case.
        row_plan.append(
            ([("接口参数", 1, True), (params, 3, False)], None)
        )
    elif params:
        n = len(params)
        row_plan.append(
            (
                [
                    ("接口参数", 1, True),
                    ("数据类型", 1, True),
                    ("参数名称", 1, True),
                    ("参数说明", 1, True),
                ],
                ("param_header", n),
            )
        )
        for p in params:
            row_plan.append(
                (
                    [
                        ("", 1, True),
                        (p.get("type", ""), 1, False),
                        (p.get("name", ""), 1, False),
                        (p.get("desc", ""), 1, False),
                    ],
                    "param_data",
                )
            )
    else:
        row_plan.append(
            ([("接口参数", 1, True), ("无。", 3, False)], None)
        )

    # Returns group.
    returns = spec.get("returns", [])
    if isinstance(returns, str):
        row_plan.append(
            ([("返回值", 1, True), (returns, 3, False)], None)
        )
    elif returns:
        n = len(returns)
        row_plan.append(
            (
                [
                    ("返回值", 1, True),
                    ("数据类型", 1, True),
                    ("返回值说明", 2, True),
                ],
                ("return_header", n),
            )
        )
        for r in returns:
            row_plan.append(
                (
                    [
                        ("", 1, True),
                        (r.get("type", ""), 1, False),
                        (r.get("desc", ""), 2, False),
                    ],
                    "return_data",
                )
            )
    else:
        row_plan.append(
            ([("返回值", 1, True), ("无。", 3, False)], None)
        )

    # Constraints and notes.
    row_plan.append(
        ([("使用限制", 1, True), (spec.get("constraints", ""), 3, False)], None)
    )
    row_plan.append(
        ([("其它说明", 1, True), (spec.get("notes", ""), 3, False)], None)
    )

    table = doc.add_table(rows=len(row_plan), cols=len(col_widths))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_grid(table, col_widths)
    _set_table_borders(table, sz="4")
    _set_table_cell_margins(table)
    _set_table_layout_fixed(table)

    # Walk row_plan and build cells. We need vMerge for multi-row groups.
    # The first row of each group gets vMerge=restart; subsequent rows
    # get vMerge=continue on the label cell.
    group_cursor = None  # (label_row_idx, group_size) for active vMerge

    for r_idx, (cells, meta) in enumerate(row_plan):
        tr = table.rows[r_idx]
        c_idx = 0
        physical_col = 0
        grid_col = 0

        # vMerge tracking: is this row a group start, continuation, or standalone?
        is_group_header = isinstance(meta, tuple) and meta[0] in (
            "param_header",
            "return_header",
        )
        is_group_data = meta in ("param_data", "return_data")

        if is_group_header:
            group_cursor = (r_idx, meta[1])
        elif is_group_data:
            pass  # handled below
        else:
            group_cursor = None

        for cell_text, span, is_header in cells:
            cell = tr.cells[physical_col]
            # Compute cell width as sum of spanned grid columns.
            span_width = sum(
                col_widths[grid_col + s]
                for s in range(span)
                if grid_col + s < len(col_widths)
            )
            _set_cell_width(cell, span_width)
            _set_grid_span(cell, span)
            if is_header:
                _set_cell_shading(cell, header_fill)
                _set_cell_vertical_align(cell, "center")
            else:
                _set_cell_shading(cell, "auto")
            _set_cell_border_width(cell, "4")
            _write_cell_text(cell, cell_text, TABLE_CELL_STYLE_ID)
            # Remove the (span-1) redundant physical cells after this one,
            # then advance physical index by 1 and grid index by span.
            _remove_grid_span_cells(tr, physical_col, span)
            physical_col += 1
            grid_col += span

        # Apply vMerge to the label cell (column 0) if part of a group.
        if is_group_header:
            _set_v_merge(tr.cells[0], "restart")
        elif is_group_data:
            _set_v_merge(tr.cells[0], "continue")


# Map table type to renderer.
TABLE_RENDERERS = {
    "macro_definition_table": render_simple_table,
    "struct_member_table": render_simple_table,
    "global_variable_table": render_simple_table,
    "interface_spec_table": render_interface_spec_table,
}


# ---------------------------------------------------------------------------
# Content rendering.
# ---------------------------------------------------------------------------


def render_paragraph(doc, block):
    btype = block.get("type", "paragraph")
    text = block.get("text", "")

    if btype == "caption":
        p = doc.add_paragraph()
        p.style = CAPTION_STYLE_ID
        p.text = text
        return

    # Flowchart placeholder from <!-- FLOWCHART: name -->.
    m = FLOWCHART_RE.match(text)
    if m:
        p = doc.add_paragraph()
        p.style = BODY_STYLE_ID
        run = p.add_run(f"[流程图: {m.group(1)}]")
        run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        return

    p = doc.add_paragraph()
    if BODY_STYLE_ID:
        p.style = BODY_STYLE_ID
    p.text = text


def render_code_block(doc, cb):
    caption = cb.get("caption")
    if caption:
        p = doc.add_paragraph()
        p.style = CAPTION_STYLE_ID
        p.text = caption
    # Render code lines as consecutive paragraphs with the code style.
    code = cb.get("code", "")
    for line in code.split("\n"):
        p = doc.add_paragraph()
        p.style = CODE_STYLE_ID
        p.text = line


def render_figure(doc, fig):
    source = fig.get("source", "")
    caption = fig.get("caption", "")

    # Image insertion: only for files we can actually read.
    source_path = Path(source)
    if source_path.exists() and source_path.suffix.lower() in (".png", ".jpg", ".jpeg", ".gif", ".bmp"):
        try:
            doc.add_picture(str(source_path), width=Cm(14))
        except Exception as e:
            p = doc.add_paragraph()
            p.text = f"[图: {source}] ({e})"
    else:
        # Placeholder for EMF / OLE / missing files.
        p = doc.add_paragraph()
        p.style = BODY_STYLE_ID
        run = p.add_run(f"[图: {source}]")
        run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    if caption:
        p = doc.add_paragraph()
        p.style = CAPTION_STYLE_ID
        p.text = caption


def render_table(doc, table_data, catalog, header_fill):
    ttype = table_data["type"]
    renderer = TABLE_RENDERERS.get(ttype)
    if renderer is None:
        raise SystemExit(f"ERROR: No renderer for table type {ttype!r}")
    family_meta = catalog.get(ttype)
    if family_meta is None:
        raise SystemExit(f"ERROR: Table type {ttype!r} not in catalog")
    caption = table_data.get("caption")
    if caption:
        # Table caption already emitted as a caption block in paragraphs
        # by the converter — do NOT render it again here. The converter
        # places caption blocks before the table in document order.
        pass
    renderer(doc, table_data, family_meta, header_fill)


def render_section(doc, section, catalog, header_fill):
    level = section.get("level", 1)
    title = section.get("title", "")

    # Heading.
    style_id = HEADING_STYLE_IDS.get(level)
    if style_id:
        p = doc.add_heading(title, level=level)
        # python-docx add_heading sets the heading style; override style_id
        # if the template uses non-standard IDs.
        try:
            p.style = style_id
        except Exception:
            pass
    else:
        p = doc.add_paragraph()
        p.text = title

    # Paragraphs.
    for block in section.get("paragraphs", []):
        render_paragraph(doc, block)

    # Code blocks.
    for cb in section.get("code_blocks", []):
        render_code_block(doc, cb)

    # Figures.
    for fig in section.get("figures", []):
        render_figure(doc, fig)

    # Tables — caption already in paragraphs, render body only.
    for tbl in section.get("tables", []):
        render_table(doc, tbl, catalog, header_fill)

    # Recurse into children.
    for child in section.get("children", []):
        render_section(doc, child, catalog, header_fill)


# ---------------------------------------------------------------------------
# Main.
# ---------------------------------------------------------------------------


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--template", required=True, help="Template DOCX file")
    ap.add_argument("--input", required=True, help="Content JSON file")
    ap.add_argument("--table-catalog", required=True, help="table_catalog.json")
    ap.add_argument("--output", required=True, help="Output DOCX file")
    args = ap.parse_args()

    catalog = load_catalog(args.table_catalog)
    header_fill = catalog.pop("_header_fill", DEFAULT_HEADER_FILL)

    data = json.loads(Path(args.input).read_text(encoding="utf-8"))
    doc = Document(args.template)

    # Insert a page break so generated content is visually separated from
    # the template's existing sample content.
    doc.add_page_break()

    for section in data.get("sections", []):
        render_section(doc, section, catalog, header_fill)

    doc.save(args.output)
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
