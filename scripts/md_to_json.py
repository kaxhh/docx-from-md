#!/usr/bin/env python3
"""
md_to_json.py — Markdown to JSON converter for AI+工业操作系统详细设计规格说明书

Template-specific converter derived from template analysis:
- template_blueprint.md: heading-to-schema mapping

Table type detection rules and JSON schema are embedded in this script
(derived from the template's table_catalog.json at generation time).

Usage:
    python3 md_to_json.py --input content.md --output data.json
"""

import argparse
import json
import re
import sys
from html.parser import HTMLParser


# ============================================================================
# Embedded JSON Schema (derived from content_contract.schema.json)
# ============================================================================

CONTENT_SCHEMA = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "title": "DetailedDesignSpecContentContract",
    "description": "AI+工业操作系统详细设计规格说明书 - JSON内容契约",
    "type": "object",
    "required": ["document", "sections"],
    "properties": {
        "document": {
            "type": "object",
            "required": ["title"],
            "properties": {
                "title": {"type": "string"},
                "subtitle": {"type": "string"},
                "material_id": {"type": "string"},
                "task_name": {"type": "string"},
                "task_number": {"type": "string"},
                "organization": {"type": "string"},
                "task_period": {"type": "string"},
                "date": {"type": "string"},
                "metadata": {"type": "object", "additionalProperties": True}
            }
        },
        "sections": {
            "type": "array",
            "items": {"$ref": "#/$defs/section"}
        }
    },
    "$defs": {
        "section": {
            "type": "object",
            "required": ["title", "level"],
            "properties": {
                "title": {"type": "string"},
                "level": {"type": "integer", "minimum": 1, "maximum": 5},
                "paragraphs": {
                    "type": "array",
                    "items": {"$ref": "#/$defs/paragraph"}
                },
                "tables": {
                    "type": "array",
                    "items": {"$ref": "#/$defs/table"}
                },
                "code_blocks": {
                    "type": "array",
                    "items": {"$ref": "#/$defs/code_block"}
                },
                "figures": {
                    "type": "array",
                    "items": {"$ref": "#/$defs/figure"}
                },
                "children": {
                    "type": "array",
                    "items": {"$ref": "#/$defs/section"}
                }
            }
        },
        "paragraph": {
            "type": "object",
            "required": ["type", "text"],
            "properties": {
                "type": {
                    "type": "string",
                    "enum": ["body", "caption", "page_break"]
                },
                "text": {"type": "string"},
                "caption_kind": {
                    "type": "string",
                    "enum": ["table", "figure", "code"]
                }
            }
        },
        "table": {
            "type": "object",
            "required": ["type"],
            "properties": {
                "type": {
                    "type": "string",
                    "enum": [
                        "macro_definition_table",
                        "struct_member_table",
                        "global_variable_table",
                        "interface_spec_table"
                    ]
                },
                "caption": {"type": "string"},
                "rows": {
                    "description": "For simple tables: array of row objects. For interface_spec_table: structured object.",
                    "type": ["array", "object"]
                }
            }
        },
        "code_block": {
            "type": "object",
            "required": ["language", "code"],
            "properties": {
                "language": {"type": "string"},
                "code": {"type": "string"},
                "caption": {"type": "string"}
            }
        },
        "figure": {
            "type": "object",
            "required": ["id", "source"],
            "properties": {
                "id": {"type": "string"},
                "source": {"type": "string"},
                "caption": {"type": "string"},
                "description": {"type": "string"}
            }
        }
    }
}


# ============================================================================
# Table Type Detection Rules (embedded from template analysis)
# ============================================================================

# Header text → table type mapping.
# Each entry maps a set of header row texts (synonyms) to a table type.
TABLE_TYPE_RULES = {
    'macro_definition_table': {
        'headers': [{'宏名称', '宏内容', '说明'}],
        'keys': ['name', 'value', 'description'],
    },
    'struct_member_table': {
        'headers': [{'成员变量', '类型', '说明'}],
        'keys': ['name', 'type', 'description'],
    },
    'global_variable_table': {
        'headers': [{'全局变量名称', '类型', '说明'}, {'全局变量名称', '数据类型', '说明'}],
        'keys': ['name', 'type', 'description'],
    },
    'interface_spec_table': {
        # Detected by first-column label text, not header row
        'row_labels': {
            '接口原型': {'field': 'prototype', 'kind': 'scalar'},
            '接口描述': {'field': 'description', 'kind': 'scalar'},
            '接口参数': {
                'field': 'parameters',
                'kind': 'list',
                'sub_fields': ['type', 'name', 'description'],
            },
            '返回值': {
                'field': 'returns',
                'kind': 'list',
                'sub_fields': ['type', 'description'],
            },
            '使用限制': {'field': 'constraints', 'kind': 'scalar'},
            '其它说明': {'field': 'notes', 'kind': 'scalar'},
        },
    },
}

# Build a lookup from normalized header set → table type
HEADER_TO_TYPE = {}
for ttype, rules in TABLE_TYPE_RULES.items():
    for header_set in rules.get('headers', []):
        key = frozenset(h.strip() for h in header_set)
        HEADER_TO_TYPE[key] = ttype


def infer_table_type(header_row):
    """Infer table type from header row text."""
    normalized = frozenset(h.strip() for h in header_row)
    if normalized in HEADER_TO_TYPE:
        return HEADER_TO_TYPE[normalized]
    return None


def is_interface_spec_table(labels):
    """Check if a table is an interface_spec_table by examining first-column labels.

    Args:
        labels: list of first-column label strings (e.g. ['接口原型', '接口描述', ...])
    """
    row_labels = TABLE_TYPE_RULES['interface_spec_table']['row_labels']
    for label in labels:
        if label and label.strip() in row_labels:
            return True
    return False


# ============================================================================
# Caption Detection Patterns
# ============================================================================

# Table caption: "表 X-Y ..." or "表X-Y ..."
TABLE_CAPTION_RE = re.compile(r'^\*{0,2}表\s*[\w\-]+\s*.*?\*{0,2}$')
TABLE_CAPTION_PREFIX_RE = re.compile(r'^\*{0,2}表\s*')

# Figure caption: "图 X-Y ..." or "图X-Y ..."
FIGURE_CAPTION_RE = re.compile(r'^\*{0,2}图\s*[\w\-]+\s*.*?\*{0,2}$')
FIGURE_CAPTION_PREFIX_RE = re.compile(r'^\*{0,2}图\s*')


def strip_bold_markers(text):
    """Strip **...** markdown bold markers."""
    return text.replace('**', '').strip()


def is_table_caption(text):
    """Check if text is a table caption."""
    stripped = strip_bold_markers(text)
    return bool(TABLE_CAPTION_PREFIX_RE.match(stripped))


def is_figure_caption(text):
    """Check if text is a figure caption."""
    stripped = strip_bold_markers(text)
    return bool(FIGURE_CAPTION_PREFIX_RE.match(stripped))


# ============================================================================
# HTML Table Parser
# ============================================================================

class HTMLTableParser(HTMLParser):
    """Parse HTML <table> blocks into a virtual grid handling rowspan."""

    def __init__(self):
        super().__init__()
        self.tables = []
        self.current_table = None
        self.current_row = None
        self.current_cell = None
        self.in_table = False
        self.in_row = False
        self.in_cell = False

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)
        if tag == 'table':
            self.in_table = True
            self.current_table = []
        elif tag == 'tr' and self.in_table:
            self.in_row = True
            self.current_row = []
        elif tag in ('td', 'th') and self.in_row:
            self.in_cell = True
            colspan = int(attrs_dict.get('colspan', 1))
            rowspan = int(attrs_dict.get('rowspan', 1))
            self.current_cell = {
                'text': '',
                'colspan': colspan,
                'rowspan': rowspan,
                'is_header': (tag == 'th'),
            }

    def handle_endtag(self, tag):
        if tag in ('td', 'th') and self.in_cell:
            self.in_cell = False
            self.current_cell['text'] = self.current_cell['text'].strip()
            self.current_row.append(self.current_cell)
            self.current_cell = None
        elif tag == 'tr' and self.in_row:
            self.in_row = False
            self.current_table.append(self.current_row)
            self.current_row = None
        elif tag == 'table' and self.in_table:
            self.in_table = False
            self.tables.append(self.current_table)
            self.current_table = None

    def handle_data(self, data):
        if self.in_cell and self.current_cell is not None:
            self.current_cell['text'] += data


def parse_html_tables(text):
    """Parse all HTML <table> blocks from text."""
    parser = HTMLTableParser()
    parser.feed(text)
    return parser.tables


def build_virtual_grid(physical_rows):
    """Build a virtual grid that accounts for rowspan carryover.

    Every row in the output has the same number of logical columns.
    Cells carried from a previous row's rowspan are marked with
    rowspan_continue=True.
    """
    virtual_grid = []
    rowspan_cells = {}  # col -> {text, remaining, colspan, is_header}

    for physical_row in physical_rows:
        virtual_row = []
        col = 0
        cell_idx = 0

        while cell_idx < len(physical_row):
            # Skip columns occupied by rowspan from previous rows
            while col in rowspan_cells:
                rs = rowspan_cells[col]
                virtual_row.append({
                    'text': rs['text'],
                    'rowspan_continue': True,
                    'colspan': 1,
                    'is_header': rs.get('is_header', False),
                })
                rs['remaining'] -= 1
                if rs['remaining'] <= 0:
                    del rowspan_cells[col]
                col += 1

            cell = physical_row[cell_idx]
            # Handle colspan by adding empty cells
            for cs in range(cell.get('colspan', 1)):
                if cs == 0:
                    virtual_row.append({
                        'text': cell['text'],
                        'colspan': 1,
                        'is_header': cell.get('is_header', False),
                        'rowspan_continue': False,
                    })
                else:
                    virtual_row.append({
                        'text': '',
                        'colspan': 1,
                        'is_header': cell.get('is_header', False),
                        'rowspan_continue': False,
                    })

            if cell.get('rowspan', 1) > 1:
                for cs in range(cell.get('colspan', 1)):
                    rowspan_cells[col + cs] = {
                        'text': cell['text'] if cs == 0 else '',
                        'remaining': cell['rowspan'] - 1,
                        'is_header': cell.get('is_header', False),
                    }

            col += cell.get('colspan', 1)
            cell_idx += 1

        # Fill remaining rowspan columns at end of row
        while col in rowspan_cells:
            rs = rowspan_cells[col]
            virtual_row.append({
                'text': rs['text'],
                'rowspan_continue': True,
                'colspan': 1,
                'is_header': rs.get('is_header', False),
            })
            rs['remaining'] -= 1
            if rs['remaining'] <= 0:
                del rowspan_cells[col]
            col += 1

        virtual_grid.append(virtual_row)

    return virtual_grid


# ============================================================================
# Markdown Pipe Table Parser
# ============================================================================

def parse_pipe_table(lines):
    """Parse a Markdown pipe table into a list of rows (each a list of cell texts)."""
    rows = []
    for line in lines:
        line = line.strip()
        if line.startswith('|') and not re.match(r'^\|[\s\-:]+\|', line):
            cells = [c.strip() for c in line.split('|')[1:-1]]
            if cells:
                rows.append(cells)
    return rows


# ============================================================================
# Interface Spec Table Converter
# ============================================================================

def convert_interface_spec_table(grid):
    """Convert a virtual grid merged table to structured JSON for interface_spec_table."""
    row_labels = TABLE_TYPE_RULES['interface_spec_table']['row_labels']
    all_labels = set(row_labels.keys())

    result = {
        'prototype': '',
        'description': '',
        'parameters': [],
        'returns': [],
        'constraints': '',
        'notes': '',
    }

    i = 0
    while i < len(grid):
        row = grid[i]
        label = row[0].get('text', '').strip() if row else ''

        if label not in row_labels:
            i += 1
            continue

        meta = row_labels[label]

        if meta['kind'] == 'scalar':
            # For scalar fields, the content is in the merged right portion
            content_parts = []
            for cell in row[1:]:
                if not cell.get('rowspan_continue', False):
                    content_parts.append(cell.get('text', ''))
            result[meta['field']] = '\n'.join(content_parts).strip()
            i += 1

        elif meta['kind'] == 'list':
            i += 1  # move past label row
            # Check for sub-header row
            if i < len(grid):
                sub_row = grid[i]
                sub_texts = [c.get('text', '').strip() for c in sub_row
                             if not c.get('rowspan_continue', False)]
                # Skip if this looks like a sub-header
                if any(t in ('数据类型', '参数名称', '参数说明', '返回值说明')
                       for t in sub_texts):
                    i += 1

            # Collect data rows until next label
            while i < len(grid):
                first_cell = grid[i][0]
                first_text = first_cell.get('text', '').strip()
                is_rowspan_cont = first_cell.get('rowspan_continue', False)

                # Only break if this is a real label row (not rowspan continuation)
                if not is_rowspan_cont and first_text in all_labels:
                    break
                if not is_rowspan_cont and first_text == label:
                    # Repeated group - skip label + sub-header
                    i += 1
                    if i < len(grid):
                        sub_texts = [c.get('text', '').strip()
                                     for c in grid[i]
                                     if not c.get('rowspan_continue', False)]
                        if any(t in ('数据类型', '参数名称', '参数说明', '返回值说明')
                               for t in sub_texts):
                            i += 1
                    continue

                # Extract data row - use full grid row to preserve column indices
                if len(grid[i]) < 4:
                    i += 1
                    continue

                if meta['field'] == 'parameters':
                    # parameters: type, name, description (cols 1, 2, 3)
                    entry = {
                        'type': grid[i][1].get('text', ''),
                        'name': grid[i][2].get('text', ''),
                        'description': grid[i][3].get('text', ''),
                    }
                    result['parameters'].append(entry)
                elif meta['field'] == 'returns':
                    # returns: type, description (cols 1, 2+3 merged)
                    desc_parts = [grid[i][j].get('text', '')
                                  for j in range(2, len(grid[i]))]
                    entry = {
                        'type': grid[i][1].get('text', ''),
                        'description': '\n'.join(desc_parts),
                    }
                    result['returns'].append(entry)

                i += 1

    return result


# ============================================================================
# Simple Table Converter
# ============================================================================

def convert_simple_table(rows, table_type):
    """Convert simple (non-merged) table rows to structured JSON."""
    rules = TABLE_TYPE_RULES[table_type]
    keys = rules['keys']
    result = []
    # Skip header row
    for row in rows[1:]:
        entry = {}
        for i, key in enumerate(keys):
            entry[key] = row[i] if i < len(row) else ''
        result.append(entry)
    return result


# ============================================================================
# MD-to-JSON Converter
# ============================================================================

class MarkdownConverter:
    """Convert Markdown to JSON following the template structure."""

    def __init__(self):
        self.result = {
            'document': {
                'title': '',
                'subtitle': '',
                'material_id': '',
                'task_name': '',
                'task_number': '',
                'organization': '',
                'task_period': '',
                'date': '',
            },
            'sections': [],
        }
        self.section_stack = []  # {level, section_dict}
        self.pending_caption = None  # (text, kind) for next table/code block
        self.in_code_block = False
        self.code_buffer = []
        self.code_language = ''
        self.table_buffer = []
        self.in_table = False
        self.cover_parsed = False

    def convert(self, md_text):
        """Convert markdown text to JSON structure."""
        lines = md_text.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i]

            # Code block
            if line.strip().startswith('```'):
                if not self.in_code_block:
                    self.in_code_block = True
                    self.code_language = line.strip()[3:].strip()
                    self.code_buffer = []
                else:
                    self.in_code_block = False
                    code_text = '\n'.join(self.code_buffer)
                    self._add_code_block(code_text, self.code_language)
                    self.code_buffer = []
                i += 1
                continue

            if self.in_code_block:
                self.code_buffer.append(line)
                i += 1
                continue

            # HTML table block
            if '<table' in line.lower():
                html_lines = [line]
                while i < len(lines) and '</table>' not in line.lower():
                    i += 1
                    if i < len(lines):
                        line = lines[i]
                        html_lines.append(line)
                html_text = '\n'.join(html_lines)
                self._process_html_table(html_text)
                i += 1
                continue

            # Heading
            heading_match = re.match(r'^(#{1,5})\s+(.+)$', line)
            if heading_match:
                level = len(heading_match.group(1))
                title = heading_match.group(2).strip()
                self._add_heading(level, title)
                i += 1
                continue

            # Pipe table
            if line.strip().startswith('|'):
                table_lines = []
                while i < len(lines) and lines[i].strip().startswith('|'):
                    table_lines.append(lines[i])
                    i += 1
                self._process_pipe_table(table_lines)
                continue

            # Cover page fields
            if not self.cover_parsed:
                cover_match = re.match(r'^材料编号[：:]\s*(.*)$', line.strip())
                if cover_match:
                    self.result['document']['material_id'] = cover_match.group(1).strip()
                    i += 1
                    continue
                task_name_match = re.match(r'^任务名称[：:]\s*(.*)$', line.strip())
                if task_name_match:
                    self.result['document']['task_name'] = task_name_match.group(1).strip()
                    i += 1
                    continue
                task_num_match = re.match(r'^任务编号[：:]\s*(.*)$', line.strip())
                if task_num_match:
                    self.result['document']['task_number'] = task_num_match.group(1).strip()
                    i += 1
                    continue
                org_match = re.match(r'^任务承担单位[：:]\s*(.*)$', line.strip())
                if org_match:
                    self.result['document']['organization'] = org_match.group(1).strip()
                    i += 1
                    continue
                period_match = re.match(r'^任务起止时间[：:]\s*(.*)$', line.strip())
                if period_match:
                    self.result['document']['task_period'] = period_match.group(1).strip()
                    i += 1
                    continue

            # Empty line
            if not line.strip():
                i += 1
                continue

            # Figure reference: ![alt](source)
            fig_match = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)\s*$', line.strip())
            if fig_match:
                alt = fig_match.group(1)
                source = fig_match.group(2)
                self._add_figure(alt, source)
                i += 1
                continue

            # Caption detection
            stripped = line.strip()
            if is_table_caption(stripped):
                caption_text = strip_bold_markers(stripped)
                self.pending_caption = (caption_text, 'table')
                i += 1
                continue

            if is_figure_caption(stripped):
                caption_text = strip_bold_marker(stripped)
                # Attach to last figure
                self._attach_figure_caption(caption_text)
                i += 1
                continue

            # Regular paragraph
            self._add_paragraph(stripped)
            i += 1

        return self.result

    def _add_heading(self, level, title):
        """Add a heading as a new section."""
        # Parse cover page title/subtitle from first H1-level content
        if level == 1 and not self.result['document'].get('title'):
            # First H1 is likely part of cover, check if it's a real chapter
            known_h1 = {'引言', '术语和缩略语', '概述', '内核详细设计'}
            if title not in known_h1:
                self.result['document']['title'] = title
                return

        section = {
            'title': title,
            'level': level,
            'paragraphs': [],
            'tables': [],
            'code_blocks': [],
            'figures': [],
            'children': [],
        }

        # Pop stack until we find a parent with lower level
        while self.section_stack and self.section_stack[-1]['level'] >= level:
            self.section_stack.pop()

        if self.section_stack:
            parent = self.section_stack[-1]
            parent_section = parent['section']
            parent_section['children'].append(section)
        else:
            self.result['sections'].append(section)

        self.section_stack.append({'level': level, 'section': section})

    def _current_section(self):
        """Get the current (deepest) section."""
        if self.section_stack:
            return self.section_stack[-1]['section']
        return None

    def _add_paragraph(self, text):
        """Add a paragraph to the current section."""
        section = self._current_section()
        if section is None:
            # Before any heading - might be cover page content
            if not self.cover_parsed:
                if text and not self.result['document'].get('title'):
                    self.result['document']['title'] = text
                elif text and not self.result['document'].get('subtitle'):
                    self.result['document']['subtitle'] = text
                elif text:
                    # Date line
                    if re.match(r'^[一二三四五六七八九十〇]+年', text):
                        self.result['document']['date'] = text
                        self.cover_parsed = True
            return

        if self.pending_caption and self.pending_caption[1] == 'table':
            # This caption is for a table - add as caption paragraph
            section['paragraphs'].append({
                'type': 'caption',
                'text': self.pending_caption[0],
                'caption_kind': 'table',
            })
            self.pending_caption = None
            return

        section['paragraphs'].append({
            'type': 'body',
            'text': text,
        })

    def _add_code_block(self, code, language):
        """Add a code block to the current section."""
        section = self._current_section()
        if section is None:
            return

        block = {
            'language': language,
            'code': code,
        }
        if self.pending_caption and self.pending_caption[1] == 'code':
            block['caption'] = self.pending_caption[0]
            self.pending_caption = None

        section['code_blocks'].append(block)

    def _add_figure(self, alt, source):
        """Add a figure reference."""
        section = self._current_section()
        if section is None:
            return

        fig = {
            'id': f'fig_{len(section["figures"]) + 1}',
            'source': source,
            'description': alt,
        }
        section['figures'].append(fig)

    def _attach_figure_caption(self, caption_text):
        """Attach a caption to the last figure in the current section."""
        section = self._current_section()
        if section and section['figures']:
            section['figures'][-1]['caption'] = caption_text

    def _process_pipe_table(self, lines):
        """Process a Markdown pipe table."""
        rows = parse_pipe_table(lines)
        if not rows:
            return

        header = rows[0]
        table_type = infer_table_type(header)

        if table_type:
            table_data = {
                'type': table_type,
                'rows': convert_simple_table(rows, table_type),
            }
        else:
            # Unknown table type - store as raw rows
            table_data = {
                'type': 'unknown',
                'rows': [{'cells': row} for row in rows],
            }

        if self.pending_caption and self.pending_caption[1] == 'table':
            table_data['caption'] = self.pending_caption[0]
            self.pending_caption = None

        section = self._current_section()
        if section:
            section['tables'].append(table_data)

    def _process_html_table(self, html_text):
        """Process an HTML table block."""
        tables = parse_html_tables(html_text)
        for physical_rows in tables:
            grid = build_virtual_grid(physical_rows)

            # Check if this is an interface_spec_table
            first_col_labels = [row[0].get('text', '').strip() for row in grid if row]
            if is_interface_spec_table(first_col_labels):
                table_data = {
                    'type': 'interface_spec_table',
                    'rows': convert_interface_spec_table(grid),
                }
            else:
                # Try to detect from header row
                header = [cell.get('text', '').strip() for cell in grid[0]] if grid else []
                table_type = infer_table_type(header)
                if table_type:
                    simple_rows = [[cell.get('text', '') for cell in row] for row in grid]
                    table_data = {
                        'type': table_type,
                        'rows': convert_simple_table(simple_rows, table_type),
                    }
                else:
                    table_data = {
                        'type': 'unknown',
                        'rows': [[cell.get('text', '') for cell in row] for row in grid],
                    }

            if self.pending_caption and self.pending_caption[1] == 'table':
                table_data['caption'] = self.pending_caption[0]
                self.pending_caption = None

            section = self._current_section()
            if section:
                section['tables'].append(table_data)


def strip_bold_marker(text):
    """Alias for strip_bold_markers."""
    return strip_bold_markers(text)


# ============================================================================
# Validation
# ============================================================================

def validate_json(data):
    """Validate output JSON against embedded schema."""
    try:
        import jsonschema
        jsonschema.validate(data, CONTENT_SCHEMA)
        return True
    except ImportError:
        # Fallback: basic structural checks
        errors = []
        if 'document' not in data:
            errors.append('Missing "document" field')
        if 'sections' not in data:
            errors.append('Missing "sections" field')
        if errors:
            print(f'Validation errors: {errors}', file=sys.stderr)
            return False
        return True
    except Exception as e:
        print(f'Validation error: {e}', file=sys.stderr)
        return False


# ============================================================================
# Main
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Convert Markdown to JSON for detailed design spec document')
    parser.add_argument('--input', '-i', required=True, help='Input Markdown file')
    parser.add_argument('--output', '-o', required=True, help='Output JSON file')
    args = parser.parse_args()

    with open(args.input, 'r', encoding='utf-8') as f:
        md_text = f.read()

    converter = MarkdownConverter()
    data = converter.convert(md_text)

    if not validate_json(data):
        print('Warning: output JSON failed validation', file=sys.stderr)

    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f'Converted {args.input} -> {args.output}')


if __name__ == '__main__':
    main()
