#!/usr/bin/env python3
"""
render_docx.py — JSON to DOCX renderer for AI+工业操作系统详细设计规格说明书

Template-specific renderer that clones the template shell and generates content.

Fidelity mode: Hybrid
- Cover page, headers, footers, styles: cloned from template
- Simple tables: built with python-docx, matching template formatting
- Interface spec tables: explicit vMerge/gridSpan with correct cell removal
- Paragraphs and headings: python-docx with template heading styles

All formatting constants (column widths, header text, cell shading, borders)
are derived from the template analysis and hardcoded in this script.
When the template changes, regenerate this script.

Usage:
    python3 render_docx.py \\
      --template template.docx \\
      --input data.json \\
      --output generated.docx
"""

import argparse
import copy
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

from lxml import etree

from docx import Document
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from docx.shared import Pt, Twips, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT


# ============================================================================
# Template Formatting Constants (from analysis)
# ============================================================================

CELL_STYLE_ID = '94'  # 表格内文字
CELL_STYLE_NAME = '表格内文字'
CAPTION_STYLE_ID = '14'  # caption
CAPTION_STYLE_NAME = 'caption'
HEADER_FILL = 'D9D9D9'
BORDER_SZ = '4'
BORDER_VAL = 'single'
BORDER_COLOR = 'auto'
CELL_MARGIN_TOP = '0'
CELL_MARGIN_LEFT = '108'
CELL_MARGIN_BOTTOM = '0'
CELL_MARGIN_RIGHT = '108'

# Heading style IDs from template: H1=2, H2=3, H3=4, H4=5, H5=6
HEADING_STYLE_MAP = {1: '2', 2: '3', 3: '4', 4: '5', 5: '6'}

# Table column widths (DXA, derived from template analysis)
TABLE_WIDTHS = {
    'macro_definition_table': [2305, 1736, 4341],
    'struct_member_table': [2683, 1991, 3818],
    'global_variable_table': [5270, 1674, 1548],
    'interface_spec_table': [1378, 1984, 1276, 4068],
}

TABLE_HEADERS = {
    'macro_definition_table': ['宏名称', '宏内容', '说明'],
    'struct_member_table': ['成员变量', '类型', '说明'],
    'global_variable_table': ['全局变量名称', '类型', '说明'],
}

# Interface spec table row labels
INTERFACE_ROW_LABELS = {
    'prototype': '接口原型',
    'description': '接口描述',
    'parameters': '接口参数',
    'returns': '返回值',
    'constraints': '使用限制',
    'notes': '其它说明',
}


# ============================================================================
# XML Helper Functions
# ============================================================================

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _get_style(doc, style_name, style_id=None):
    """Get a style by name with optional fallback to ID."""
    try:
        return doc.styles[style_name]
    except KeyError:
        if style_id:
            try:
                return doc.styles.get_by_id(style_id, 'paragraph')
            except:
                pass
        return None


def _make_tc_pr(width_dxa, shading=None, grid_span=1, v_merge=None,
                border_sz=BORDER_SZ):
    """Create a <w:tcPr> element."""
    tc_pr = etree.SubElement(etree.Element('dummy'), qn('w:tcPr'))

    # Width
    tc_w = etree.SubElement(tc_pr, qn('w:tcW'))
    tc_w.set(qn('w:w'), str(width_dxa))
    tc_w.set(qn('w:type'), 'dxa')

    # Borders
    borders = etree.SubElement(tc_pr, qn('w:tcBorders'))
    for side in ('top', 'left', 'bottom', 'right'):
        b = etree.SubElement(borders, qn(f'w:{side}'))
        b.set(qn('w:val'), BORDER_VAL)
        b.set(qn('w:color'), BORDER_COLOR)
        b.set(qn('w:sz'), border_sz)
        b.set(qn('w:space'), '0')

    # Shading
    shd = etree.SubElement(tc_pr, qn('w:shd'))
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), shading if shading else 'auto')

    # Grid span
    if grid_span > 1:
        gs = etree.SubElement(tc_pr, qn('w:gridSpan'))
        gs.set(qn('w:val'), str(grid_span))

    # Vertical merge
    if v_merge == 'restart':
        vm = etree.SubElement(tc_pr, qn('w:vMerge'))
        vm.set(qn('w:val'), 'restart')
    elif v_merge == 'continue':
        vm = etree.SubElement(tc_pr, qn('w:vMerge'))

    # Vertical align
    va = etree.SubElement(tc_pr, qn('w:vAlign'))
    va.set(qn('w:val'), 'center')

    return tc_pr


def _remove_grid_span_cells(row, current_idx, span):
    """Remove (span-1) physical cells right after the current cell."""
    if span <= 1:
        return
    tcs = list(row._tr.iterchildren(qn('w:tc')))
    to_remove = tcs[current_idx + 1: current_idx + span]
    for tc in to_remove:
        row._tr.remove(tc)


def _set_cell_style(cell, style_name=CELL_STYLE_NAME):
    """Set paragraph style for all paragraphs in a cell."""
    style = _get_style(cell.part.document, style_name)
    if style:
        for para in cell.paragraphs:
            para.style = style


def _set_cell_text(cell, text, style_name=CELL_STYLE_NAME, is_header=False):
    """Set cell text with proper paragraph style."""
    # Get style by name
    style = _get_style(cell.part.document, style_name)

    # Clear existing content
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ''

    # Handle multiline text
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if i == 0:
            para = cell.paragraphs[0]
            para.text = ''
            run = para.add_run(line)
        else:
            para = cell.add_paragraph()
            run = para.add_run(line)
        if style:
            para.style = style


def _add_table_to_doc(doc, n_rows, n_cols, col_widths):
    """Add a table with correct grid columns."""
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set tblGrid
    tbl = table._tbl
    tbl_grid = tbl.find(qn('w:tblGrid'))
    if tbl_grid is None:
        tbl_grid = etree.SubElement(tbl, qn('w:tblGrid'))
    else:
        for gc in tbl_grid.findall(qn('w:gridCol')):
            tbl_grid.remove(gc)

    for w in col_widths:
        gc = etree.SubElement(tbl_grid, qn('w:gridCol'))
        gc.set(qn('w:w'), str(w))

    # Set tblPr
    tbl_pr = tbl.find(qn('w:tblPr'))
    if tbl_pr is not None:
        # Cell margins
        cell_mar = tbl_pr.find(qn('w:tblCellMar'))
        if cell_mar is None:
            cell_mar = etree.SubElement(tbl_pr, qn('w:tblCellMar'))
        for side, val in [('top', CELL_MARGIN_TOP), ('left', CELL_MARGIN_LEFT),
                          ('bottom', CELL_MARGIN_BOTTOM), ('right', CELL_MARGIN_RIGHT)]:
            m = cell_mar.find(qn(f'w:{side}'))
            if m is None:
                m = etree.SubElement(cell_mar, qn(f'w:{side}'))
            m.set(qn('w:w'), val)
            m.set(qn('w:type'), 'dxa')

    return table


# ============================================================================
# Simple Table Renderers
# ============================================================================

def render_simple_table(doc, table_data, table_type):
    """Render a simple (non-merged) table."""
    col_widths = TABLE_WIDTHS[table_type]
    headers = TABLE_HEADERS[table_type]

    # Get column keys from data
    if table_type == 'macro_definition_table':
        keys = ['name', 'value', 'description']
    elif table_type == 'struct_member_table':
        keys = ['name', 'type', 'description']
    elif table_type == 'global_variable_table':
        keys = ['name', 'type', 'description']
    else:
        keys = ['name', 'type', 'description']

    rows = table_data.get('rows', [])
    n_rows = 1 + len(rows)  # header + data

    table = _add_table_to_doc(doc, n_rows, len(col_widths), col_widths)

    # Header row
    for j, (header_text, width) in enumerate(zip(headers, col_widths)):
        cell = table.rows[0].cells[j]
        _set_cell_text(cell, header_text, is_header=True)
        # Apply header shading via tcPr
        tc = cell._tc
        tc_pr = tc.find(qn('w:tcPr'))
        if tc_pr is None:
            tc_pr = etree.Element(qn('w:tcPr'))
            tc.insert(0, tc_pr)
        shd = tc_pr.find(qn('w:shd'))
        if shd is None:
            shd = etree.SubElement(tc_pr, qn('w:shd'))
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), HEADER_FILL)

    # Data rows
    for i, row_data in enumerate(rows):
        for j, key in enumerate(keys):
            cell = table.rows[i + 1].cells[j]
            text = row_data.get(key, '') if isinstance(row_data, dict) else str(row_data)
            _set_cell_text(cell, text)

    return table


# ============================================================================
# Interface Spec Table Renderer
# ============================================================================

def render_interface_spec_table(doc, table_data):
    """Render an interface_spec_table with vMerge and gridSpan."""
    col_widths = TABLE_WIDTHS['interface_spec_table']
    total_width = sum(col_widths)
    rows_data = table_data.get('rows', {})

    # Build row plan: list of (cells_text, span_info, v_merge_info)
    row_plan = []

    # Row 0: 接口原型 (label + content spanning 3 cols)
    row_plan.append({
        'cells': [
            {'text': INTERFACE_ROW_LABELS['prototype'], 'span': 1, 'shading': HEADER_FILL, 'vmerge': None},
            {'text': rows_data.get('prototype', ''), 'span': 3, 'shading': None, 'vmerge': None},
        ]
    })

    # Row 1: 接口描述 (label + content spanning 3 cols)
    row_plan.append({
        'cells': [
            {'text': INTERFACE_ROW_LABELS['description'], 'span': 1, 'shading': HEADER_FILL, 'vmerge': None},
            {'text': rows_data.get('description', ''), 'span': 3, 'shading': None, 'vmerge': None},
        ]
    })

    # Parameters section
    params = rows_data.get('parameters', [])
    if params:
        # Label row (vmerge restart) + sub-header
        row_plan.append({
            'cells': [
                {'text': INTERFACE_ROW_LABELS['parameters'], 'span': 1, 'shading': HEADER_FILL, 'vmerge': 'restart'},
                {'text': '数据类型', 'span': 1, 'shading': HEADER_FILL, 'vmerge': None},
                {'text': '参数名称', 'span': 1, 'shading': HEADER_FILL, 'vmerge': None},
                {'text': '参数说明', 'span': 1, 'shading': HEADER_FILL, 'vmerge': None},
            ]
        })
        for p in params:
            row_plan.append({
                'cells': [
                    {'text': '', 'span': 1, 'shading': None, 'vmerge': 'continue'},
                    {'text': p.get('type', ''), 'span': 1, 'shading': None, 'vmerge': None},
                    {'text': p.get('name', ''), 'span': 1, 'shading': None, 'vmerge': None},
                    {'text': p.get('description', ''), 'span': 1, 'shading': None, 'vmerge': None},
                ]
            })
    else:
        row_plan.append({
            'cells': [
                {'text': INTERFACE_ROW_LABELS['parameters'], 'span': 1, 'shading': HEADER_FILL, 'vmerge': None},
                {'text': '无', 'span': 3, 'shading': None, 'vmerge': None},
            ]
        })

    # Returns section
    returns = rows_data.get('returns', [])
    if returns:
        row_plan.append({
            'cells': [
                {'text': INTERFACE_ROW_LABELS['returns'], 'span': 1, 'shading': HEADER_FILL, 'vmerge': 'restart'},
                {'text': '数据类型', 'span': 1, 'shading': HEADER_FILL, 'vmerge': None},
                {'text': '返回值说明', 'span': 2, 'shading': HEADER_FILL, 'vmerge': None},
            ]
        })
        for r in returns:
            row_plan.append({
                'cells': [
                    {'text': '', 'span': 1, 'shading': None, 'vmerge': 'continue'},
                    {'text': r.get('type', ''), 'span': 1, 'shading': None, 'vmerge': None},
                    {'text': r.get('description', ''), 'span': 2, 'shading': None, 'vmerge': None},
                ]
            })
    else:
        row_plan.append({
            'cells': [
                {'text': INTERFACE_ROW_LABELS['returns'], 'span': 1, 'shading': HEADER_FILL, 'vmerge': None},
                {'text': '无', 'span': 3, 'shading': None, 'vmerge': None},
            ]
        })

    # Constraints
    row_plan.append({
        'cells': [
            {'text': INTERFACE_ROW_LABELS['constraints'], 'span': 1, 'shading': HEADER_FILL, 'vmerge': None},
            {'text': rows_data.get('constraints', '无'), 'span': 3, 'shading': None, 'vmerge': None},
        ]
    })

    # Notes
    row_plan.append({
        'cells': [
            {'text': INTERFACE_ROW_LABELS['notes'], 'span': 1, 'shading': HEADER_FILL, 'vmerge': None},
            {'text': rows_data.get('notes', '无'), 'span': 3, 'shading': None, 'vmerge': None},
        ]
    })

    n_rows = len(row_plan)
    n_cols = 4  # grid columns
    table = _add_table_to_doc(doc, n_rows, n_cols, col_widths)

    # Render each row
    for r_idx, row_info in enumerate(row_plan):
        tr = table.rows[r_idx]
        physical_col = 0
        grid_col = 0

        for cell_info in row_info['cells']:
            span = cell_info['span']
            vmerge = cell_info['vmerge']
            shading = cell_info['shading']
            text = cell_info['text']

            cell = tr.cells[physical_col]

            # Calculate width for this cell
            span_width = sum(col_widths[grid_col + s] for s in range(span))

            # Set cell properties
            tc = cell._tc
            tc_pr = tc.find(qn('w:tcPr'))
            if tc_pr is None:
                tc_pr = etree.Element(qn('w:tcPr'))
                tc.insert(0, tc_pr)

            # Width
            tc_w = tc_pr.find(qn('w:tcW'))
            if tc_w is None:
                tc_w = etree.SubElement(tc_pr, qn('w:tcW'))
            tc_w.set(qn('w:w'), str(span_width))
            tc_w.set(qn('w:type'), 'dxa')

            # Borders
            borders = tc_pr.find(qn('w:tcBorders'))
            if borders is None:
                borders = etree.SubElement(tc_pr, qn('w:tcBorders'))
            for side in ('top', 'left', 'bottom', 'right'):
                b = borders.find(qn(f'w:{side}'))
                if b is None:
                    b = etree.SubElement(borders, qn(f'w:{side}'))
                b.set(qn('w:val'), BORDER_VAL)
                b.set(qn('w:color'), BORDER_COLOR)
                b.set(qn('w:sz'), BORDER_SZ)
                b.set(qn('w:space'), '0')

            # Shading
            shd = tc_pr.find(qn('w:shd'))
            if shd is None:
                shd = etree.SubElement(tc_pr, qn('w:shd'))
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), shading if shading else 'auto')

            # Grid span
            if span > 1:
                gs = tc_pr.find(qn('w:gridSpan'))
                if gs is None:
                    gs = etree.SubElement(tc_pr, qn('w:gridSpan'))
                gs.set(qn('w:val'), str(span))

            # Vertical merge
            vm = tc_pr.find(qn('w:vMerge'))
            if vmerge == 'restart':
                if vm is None:
                    vm = etree.SubElement(tc_pr, qn('w:vMerge'))
                vm.set(qn('w:val'), 'restart')
            elif vmerge == 'continue':
                if vm is None:
                    vm = etree.SubElement(tc_pr, qn('w:vMerge'))
                # No val attribute for continue
                if qn('w:val') in vm.attrib:
                    del vm.attrib[qn('w:val')]

            # Vertical align
            va = tc_pr.find(qn('w:vAlign'))
            if va is None:
                va = etree.SubElement(tc_pr, qn('w:vAlign'))
            va.set(qn('w:val'), 'center')

            # Set text
            _set_cell_text(cell, text)

            # Remove covered cells for gridSpan
            _remove_grid_span_cells(tr, physical_col, span)

            physical_col += 1
            grid_col += span

    return table


# ============================================================================
# Document Renderer
# ============================================================================

class DocumentRenderer:
    """Render JSON content into a DOCX file."""

    def __init__(self, template_path):
        self.doc = Document(template_path)

        # Remove all existing content from template body
        body = self.doc.element.body
        for child in list(body):
            if child.tag != qn('w:sectPr'):
                body.remove(child)

    def render(self, data, output_path):
        """Render the complete document."""
        # Render cover page info
        self._render_cover(data.get('document', {}))

        # Render sections
        for section in data.get('sections', []):
            self._render_section(section)

        self.doc.save(output_path)

    def _render_cover(self, doc_info):
        """Render cover page fields."""
        # Material ID
        if doc_info.get('material_id'):
            p = self.doc.add_paragraph(f"材料编号：{doc_info['material_id']}")

        self.doc.add_paragraph('')
        self.doc.add_paragraph('')

        # Title
        if doc_info.get('title'):
            p = self.doc.add_paragraph(doc_info['title'])
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(22)
                run.font.bold = True

        # Subtitle
        if doc_info.get('subtitle'):
            p = self.doc.add_paragraph(doc_info['subtitle'])
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(18)

        self.doc.add_paragraph('')
        self.doc.add_paragraph('')

        # Task info fields
        for field, label in [('task_name', '任务名称'), ('task_number', '任务编号'),
                             ('organization', '任务承担单位'), ('task_period', '任务起止时间')]:
            if doc_info.get(field):
                self.doc.add_paragraph(f'{label}：{doc_info[field]}')
            else:
                self.doc.add_paragraph(f'{label}：')

        self.doc.add_paragraph('')
        self.doc.add_paragraph('')

        if doc_info.get('date'):
            p = self.doc.add_paragraph(doc_info['date'])
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Page break after cover
        self.doc.add_page_break()

    def _render_section(self, section):
        """Render a section (heading + content + children)."""
        level = section.get('level', 1)
        title = section.get('title', '')

        # Add heading
        try:
            heading = self.doc.add_heading(title, level=level)
        except Exception:
            heading = self.doc.add_paragraph(title)
            style_name = f'heading {level}'
            style = _get_style(self.doc, style_name)
            if style:
                heading.style = style

        # Render paragraphs
        for para_data in section.get('paragraphs', []):
            self._render_paragraph(para_data)

        # Render tables
        for table_data in section.get('tables', []):
            self._render_table(table_data)

        # Render code blocks
        for code_data in section.get('code_blocks', []):
            self._render_code_block(code_data)

        # Render figures
        for fig_data in section.get('figures', []):
            self._render_figure(fig_data)

        # Render children
        for child in section.get('children', []):
            self._render_section(child)

    def _render_paragraph(self, para_data):
        """Render a paragraph."""
        para_type = para_data.get('type', 'body')
        text = para_data.get('text', '')

        if para_type == 'page_break':
            self.doc.add_page_break()
            return

        if para_type == 'caption':
            style = _get_style(self.doc, CAPTION_STYLE_NAME, CAPTION_STYLE_ID)
            if style:
                p = self.doc.add_paragraph(text, style=style)
            else:
                p = self.doc.add_paragraph(text)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.bold = True
            return

        p = self.doc.add_paragraph(text)

    def _render_table(self, table_data):
        """Render a table based on its type."""
        table_type = table_data.get('type')

        # Render caption if present
        caption = table_data.get('caption')
        if caption:
            style = _get_style(self.doc, CAPTION_STYLE_NAME, CAPTION_STYLE_ID)
            if style:
                p = self.doc.add_paragraph(caption, style=style)
            else:
                p = self.doc.add_paragraph(caption)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.bold = True

        if table_type == 'interface_spec_table':
            render_interface_spec_table(self.doc, table_data)
        elif table_type in ('macro_definition_table', 'struct_member_table',
                           'global_variable_table'):
            render_simple_table(self.doc, table_data, table_type)
        else:
            print(f'Warning: unknown table type "{table_type}", skipping',
                  file=sys.stderr)

    def _render_code_block(self, code_data):
        """Render a code block."""
        caption = code_data.get('caption')
        if caption:
            style = _get_style(self.doc, CAPTION_STYLE_NAME, CAPTION_STYLE_ID)
            if style:
                p = self.doc.add_paragraph(caption, style=style)
            else:
                p = self.doc.add_paragraph(caption)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        code = code_data.get('code', '')
        p = self.doc.add_paragraph()
        run = p.add_run(code)
        run.font.name = 'Courier New'
        run.font.size = Pt(9)

    def _render_figure(self, fig_data):
        """Render a figure placeholder."""
        source = fig_data.get('source', '')
        caption = fig_data.get('caption', '')

        # Figure placeholder
        p = self.doc.add_paragraph(f'[Figure: {source}]')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Caption after figure
        if caption:
            style = _get_style(self.doc, CAPTION_STYLE_NAME, CAPTION_STYLE_ID)
            if style:
                p = self.doc.add_paragraph(caption, style=style)
            else:
                p = self.doc.add_paragraph(caption)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.bold = True


# ============================================================================
# Main
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Render JSON content to DOCX using template formatting')
    parser.add_argument('--template', '-t', required=True, help='Template DOCX file')
    parser.add_argument('--input', '-i', required=True, help='Input JSON file')
    parser.add_argument('--output', '-o', required=True, help='Output DOCX file')
    args = parser.parse_args()

    with open(args.input, 'r', encoding='utf-8') as f:
        data = json.load(f)

    renderer = DocumentRenderer(args.template)
    renderer.render(data, args.output)

    print(f'Rendered {args.input} -> {args.output}')


if __name__ == '__main__':
    main()
