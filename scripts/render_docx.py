#!/usr/bin/env python3
"""JSON → DOCX renderer for AI+工业设计文档模板.

Derived from template: skill-test/AI+template.docx
Template analysis: skill-test/docx-template-analysis/

Supported table families:
  - macro_definition_table (3 col, no merges)
  - struct_member_table (3 col, no merges)
  - global_variable_table (3 col, no merges)
  - interface_spec_table (4 col, vMerge + gridSpan)

Fidelity: High — clones template styles, uses fixed DXA widths,
preserves heading numbering, table cell style '表格内文字',
caption style 'Caption', header fill D9D9D9, border sz=4.

Usage:
    python3 render_docx.py \
        --template template.docx \
        --input content_data.json \
        --output generated.docx
"""

import argparse
import copy
import json
import os
import re
import shutil
import sys

from docx import Document
from docx.shared import Pt, Twips, RGBColor
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

# ---------------------------------------------------------------------------
# Table configuration — derived from table_catalog.json
# ---------------------------------------------------------------------------
TABLE_CONFIG = {
    "macro_definition_table": {
        "grid_columns": [2305, 1736, 4341],
        "header_labels": ["宏名称", "宏内容", "说明"],
        "row_keys": ["name", "value", "description"],
    },
    "struct_member_table": {
        "grid_columns": [2683, 1991, 3818],
        "header_labels": ["成员变量", "类型", "说明"],
        "row_keys": ["name", "type", "description"],
    },
    "global_variable_table": {
        "grid_columns": [5270, 1674, 1548],
        "header_labels": ["全局变量名称", "类型", "说明"],
        "row_keys": ["name", "type", "description"],
    },
    "interface_spec_table": {
        "grid_columns": [1378, 1984, 1276, 4068],
        "row_labels": {
            "prototype": "接口原型",
            "description": "接口描述",
            "parameters": "接口参数",
            "returns": "返回值",
            "constraints": "使用限制",
            "notes": "其它说明",
        },
        "param_sub_header": ["数据类型", "参数名称", "参数说明"],
        "return_sub_header": ["数据类型", "返回值说明"],
    },
}

# Common formatting constants (from template analysis)
HEADER_FILL = "D9D9D9"
CELL_PARA_STYLE = "表格内文字"
CAPTION_STYLE = "Caption"
BORDER_SZ = "4"
ROW_HEIGHT_VAL = "270"
CELL_MARGIN_LEFT = 108
CELL_MARGIN_RIGHT = 108


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------
def _make_element(tag, **attrs):
    """Create an OxmlElement with attributes."""
    elem = OxmlElement(tag)
    for k, v in attrs.items():
        elem.set(qn(f"w:{k}"), str(v))
    return elem


def _set_cell_shading(tc, fill):
    """Set cell background fill."""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    # Remove existing shd
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def _set_cell_width(tc, width_dxa):
    """Set cell width in DXA."""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(width_dxa))
    tcW.set(qn("w:type"), "dxa")


def _set_cell_borders(tc, sz=BORDER_SZ):
    """Set all four cell borders."""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    for side in ("top", "left", "bottom", "right"):
        elem = borders.find(qn(f"w:{side}"))
        if elem is None:
            elem = OxmlElement(f"w:{side}")
            borders.append(elem)
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:color"), "auto")
        elem.set(qn("w:sz"), sz)
        elem.set(qn("w:space"), "0")


def _set_cell_valign(tc, val="center"):
    """Set cell vertical alignment."""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    vAlign = tcPr.find(qn("w:vAlign"))
    if vAlign is None:
        vAlign = OxmlElement("w:vAlign")
        tcPr.append(vAlign)
    vAlign.set(qn("w:val"), val)


def _set_grid_span(tc, span):
    """Set w:gridSpan on a cell."""
    if span <= 1:
        return
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    gs = tcPr.find(qn("w:gridSpan"))
    if gs is None:
        gs = OxmlElement("w:gridSpan")
        tcPr.append(gs)
    gs.set(qn("w:val"), str(span))


def _set_vmerge(tc, mode):
    """Set vMerge on a cell: 'restart' or 'continue'."""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    vm = tcPr.find(qn("w:vMerge"))
    if vm is None:
        vm = OxmlElement("w:vMerge")
        tcPr.append(vm)
    if mode == "restart":
        vm.set(qn("w:val"), "restart")
    # 'continue' = no w:val attribute (Word convention)


def _remove_grid_span_cells(tr, current_idx, span):
    """Remove (span-1) physical cells after the current cell."""
    if span <= 1:
        return
    tcs = list(tr.iterchildren(qn("w:tc")))
    to_remove = tcs[current_idx + 1: current_idx + span]
    for tc in to_remove:
        tr.remove(tc)


def _set_row_height(tr, val=ROW_HEIGHT_VAL, rule="atLeast"):
    """Set row height."""
    trPr = tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr.insert(0, trPr)
    trHeight = trPr.find(qn("w:trHeight"))
    if trHeight is None:
        trHeight = OxmlElement("w:trHeight")
        trPr.append(trHeight)
    trHeight.set(qn("w:val"), val)
    trHeight.set(qn("w:hRule"), rule)


def _set_table_borders(tbl, outer_sz="6", inner_sz=BORDER_SZ):
    """Set table-level borders with outer/inner distinction."""
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        return
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)

    for side in ("top", "left", "bottom", "right"):
        elem = borders.find(qn(f"w:{side}"))
        if elem is None:
            elem = OxmlElement(f"w:{side}")
            borders.append(elem)
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:color"), "auto")
        elem.set(qn("w:sz"), outer_sz)
        elem.set(qn("w:space"), "0")

    for side in ("insideH", "insideV"):
        elem = borders.find(qn(f"w:{side}"))
        if elem is None:
            elem = OxmlElement(f"w:{side}")
            borders.append(elem)
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:color"), "auto")
        elem.set(qn("w:sz"), inner_sz)
        elem.set(qn("w:space"), "0")


def _set_table_cell_margins(tbl):
    """Set table cell margins."""
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        return
    mar = tblPr.find(qn("w:tblCellMar"))
    if mar is None:
        mar = OxmlElement("w:tblCellMar")
        tblPr.append(mar)
    for side, val in [("top", 0), ("left", CELL_MARGIN_LEFT), ("bottom", 0), ("right", CELL_MARGIN_RIGHT)]:
        elem = mar.find(qn(f"w:{side}"))
        if elem is None:
            elem = OxmlElement(f"w:{side}")
            mar.append(elem)
        elem.set(qn("w:w"), str(val))
        elem.set(qn("w:type"), "dxa")


def _write_cell_text(cell, text, style_name=CELL_PARA_STYLE):
    """Write text into a cell, splitting on newlines into separate paragraphs."""
    # Clear existing content
    for p in cell.paragraphs:
        p_elem = p._element
        for child in list(p_elem):
            if child.tag != qn("w:pPr"):
                p_elem.remove(child)

    lines = text.split("\n") if text else [""]
    for i, line in enumerate(lines):
        if i == 0:
            p = cell.paragraphs[0]
        else:
            p = cell.add_paragraph()
        p.style = style_name if style_name else p.style
        run = p.add_run(line)


# ---------------------------------------------------------------------------
# Table renderers
# ---------------------------------------------------------------------------
def render_simple_table(doc, table_data, config):
    """Render a 3-column flat table (macro/struct/global)."""
    grid_cols = config["grid_columns"]
    header_labels = config["header_labels"]
    row_keys = config["row_keys"]
    total_width = sum(grid_cols)

    num_rows = 1 + len(table_data.get("rows", []))
    table = doc.add_table(rows=num_rows, cols=len(grid_cols))

    # Table properties
    tbl_elem = table._tbl
    _set_table_borders(tbl_elem)
    _set_table_cell_margins(tbl_elem)

    # Set grid column widths
    tblGrid = tbl_elem.find(qn("w:tblGrid"))
    if tblGrid is None:
        tblGrid = OxmlElement("w:tblGrid")
        tbl_elem.insert(1, tblGrid)
    else:
        for gc in list(tblGrid):
            tblGrid.remove(gc)
    for w in grid_cols:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        tblGrid.append(gc)

    # Header row
    header_row = table.rows[0]
    _set_row_height(header_row._tr)
    for col_idx, (label, width) in enumerate(zip(header_labels, grid_cols)):
        cell = header_row.cells[col_idx]
        tc = cell._tc
        _set_cell_width(tc, width)
        _set_cell_shading(tc, HEADER_FILL)
        _set_cell_borders(tc)
        _set_cell_valign(tc)
        _write_cell_text(cell, label)

    # Body rows
    for row_idx, row_data in enumerate(table_data.get("rows", []), start=1):
        row = table.rows[row_idx]
        _set_row_height(row._tr)
        for col_idx, (key, width) in enumerate(zip(row_keys, grid_cols)):
            cell = row.cells[col_idx]
            tc = cell._tc
            _set_cell_width(tc, width)
            _set_cell_shading(tc, "auto")
            _set_cell_borders(tc)
            _set_cell_valign(tc)
            _write_cell_text(cell, row_data.get(key, ""))

    return table


def render_interface_spec_table(doc, table_data, config):
    """Render a merged interface specification table."""
    grid_cols = config["grid_columns"]
    row_labels = config["row_labels"]
    total_width = sum(grid_cols)
    rows_data = table_data.get("rows", {})

    # Build row plan: list of (cells, merge_info)
    # Each cell: (text, span, is_header)
    row_plan = []

    # 1. 接口原型
    proto_text = rows_data.get("prototype", "")
    row_plan.append([
        (row_labels["prototype"], 1, True),
        (proto_text, 3, False),
    ])

    # 2. 接口描述
    desc_text = rows_data.get("description", "")
    row_plan.append([
        (row_labels["description"], 1, True),
        (desc_text, 3, False),
    ])

    # 3. 接口参数
    params = rows_data.get("parameters", [])
    if params:
        row_plan.append([
            (row_labels["parameters"], 1, True, "restart"),
            (config["param_sub_header"][0], 1, True),
            (config["param_sub_header"][1], 1, True),
            (config["param_sub_header"][2], 1, True),
        ])
        for p in params:
            row_plan.append([
                ("", 1, True, "continue"),
                (p.get("type", ""), 1, False),
                (p.get("name", ""), 1, False),
                (p.get("description", ""), 1, False),
            ])
    else:
        row_plan.append([
            (row_labels["parameters"], 1, True),
            ("无", 3, False),
        ])

    # 4. 返回值
    returns = rows_data.get("returns", [])
    if returns:
        row_plan.append([
            (row_labels["returns"], 1, True, "restart"),
            (config["return_sub_header"][0], 1, True),
            (config["return_sub_header"][1], 2, True),
        ])
        for r in returns:
            row_plan.append([
                ("", 1, True, "continue"),
                (r.get("type", ""), 1, False),
                (r.get("description", ""), 2, False),
            ])
    else:
        row_plan.append([
            (row_labels["returns"], 1, True),
            ("无", 3, False),
        ])

    # 5. 使用限制
    row_plan.append([
        (row_labels["constraints"], 1, True),
        (rows_data.get("constraints", "无"), 3, False),
    ])

    # 6. 其它说明
    row_plan.append([
        (row_labels["notes"], 1, True),
        (rows_data.get("notes", "无"), 3, False),
    ])

    # Create table
    num_rows = len(row_plan)
    table = doc.add_table(rows=num_rows, cols=4)

    tbl_elem = table._tbl
    _set_table_borders(tbl_elem)
    _set_table_cell_margins(tbl_elem)

    # Set grid column widths
    tblGrid = tbl_elem.find(qn("w:tblGrid"))
    if tblGrid is None:
        tblGrid = OxmlElement("w:tblGrid")
        tbl_elem.insert(1, tblGrid)
    else:
        for gc in list(tblGrid):
            tblGrid.remove(gc)
    for w in grid_cols:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        tblGrid.append(gc)

    # Render rows
    for r_idx, cells_spec in enumerate(row_plan):
        tr = table.rows[r_idx]._tr
        _set_row_height(tr)
        physical_col = 0
        grid_col = 0

        for cell_spec in cells_spec:
            if len(cell_spec) == 4:
                text, span, is_header, vmerge_mode = cell_spec
            else:
                text, span, is_header = cell_spec
                vmerge_mode = None

            cell = table.rows[r_idx].cells[physical_col]
            tc = cell._tc

            # Width = sum of grid columns covered
            span_width = sum(grid_cols[grid_col + s] for s in range(span))
            _set_cell_width(tc, span_width)
            _set_cell_borders(tc)
            _set_cell_valign(tc)

            # Shading
            if is_header:
                _set_cell_shading(tc, HEADER_FILL)
            else:
                _set_cell_shading(tc, "auto")

            # GridSpan
            _set_grid_span(tc, span)

            # vMerge
            if vmerge_mode:
                _set_vmerge(tc, vmerge_mode)

            # Text
            _write_cell_text(cell, text)

            # Remove covered physical cells
            _remove_grid_span_cells(tr, physical_col, span)

            physical_col += 1
            grid_col += span

    return table


# ---------------------------------------------------------------------------
# Document renderer
# ---------------------------------------------------------------------------
class DocxRenderer:
    """Render JSON content into a DOCX document based on the template."""

    HEADING_STYLES = {
        1: "Heading 1",
        2: "Heading 2",
        3: "Heading 3",
        4: "Heading 4",
        5: "Heading 5",
        6: "Heading 6",
    }

    def __init__(self, template_path):
        self.template_path = template_path
        self.doc = Document(template_path)
        self._prepare_template()

    def _prepare_template(self):
        """Remove template body content, keeping styles/sections/headers/footers."""
        # Remove all body elements except sectPr (section properties)
        body = self.doc.element.body
        elements_to_remove = []
        for child in list(body):
            if child.tag == qn("w:sectPr"):
                continue
            elements_to_remove.append(child)
        for elem in elements_to_remove:
            body.remove(elem)

    def render(self, data, output_path):
        """Render JSON data to output .docx file."""
        # Render document title (if needed, could be in cover page)
        # Render sections recursively
        for section in data.get("sections", []):
            self._render_section(section)

        self.doc.save(output_path)

    def _render_section(self, section):
        """Render a section (heading + content + children)."""
        level = section.get("level", 1)
        title = section.get("title", "")

        # Add heading
        style_name = self.HEADING_STYLES.get(level, "Heading 1")
        self.doc.add_heading(title, level=level)

        # Add paragraphs
        for para in section.get("paragraphs", []):
            self._render_paragraph(para)

        # Add code blocks
        for block in section.get("code_blocks", []):
            self._render_code_block(block)

        # Add tables
        for table_data in section.get("tables", []):
            self._render_table(table_data)

        # Add figures
        for fig in section.get("figures", []):
            self._render_figure(fig)

        # Render children
        for child in section.get("children", []):
            self._render_section(child)

    def _render_paragraph(self, para):
        """Render a paragraph block."""
        text = para.get("text", "")
        para_type = para.get("type", "text")

        if para_type == "caption":
            p = self.doc.add_paragraph(text)
            p.style = CAPTION_STYLE
        elif para_type == "text":
            self.doc.add_paragraph(text)

    def _render_code_block(self, block):
        """Render a code block."""
        code = block.get("code", "")
        lang = block.get("language", "")

        # Use '代码' style if available, else Normal with monospace
        p = self.doc.add_paragraph()
        try:
            p.style = "代码"
        except KeyError:
            pass

        run = p.add_run(code)
        run.font.name = "Courier New"
        run.font.size = Pt(9)

    def _render_table(self, table_data):
        """Render a table based on its type."""
        table_type = table_data.get("type")
        config = TABLE_CONFIG.get(table_type)

        if config is None:
            print(f"WARNING: Unknown table type: {table_type}", file=sys.stderr)
            return

        # Caption is already rendered as a paragraph block — do NOT render
        # tables[].caption again (single source of truth rule)

        if table_type == "interface_spec_table":
            render_interface_spec_table(self.doc, table_data, config)
        else:
            render_simple_table(self.doc, table_data, config)

    def _render_figure(self, fig):
        """Render a figure placeholder + caption."""
        source = fig.get("source", "")
        caption = fig.get("caption", "")

        # Add figure placeholder paragraph
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"[Figure: {source}]")
        run.font.size = Pt(10)

        # Add caption after figure
        if caption:
            cp = self.doc.add_paragraph(caption)
            cp.style = CAPTION_STYLE


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="JSON → DOCX renderer for AI+设计文档模板")
    parser.add_argument("--template", required=True, help="Template .docx file")
    parser.add_argument("--input", required=True, help="Input JSON file")
    parser.add_argument("--output", required=True, help="Output .docx file")
    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    renderer = DocxRenderer(args.template)
    renderer.render(data, args.output)

    print(f"Generated: {args.output}")


if __name__ == "__main__":
    main()
