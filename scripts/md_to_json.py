#!/usr/bin/env python3
"""MD-to-JSON converter for AI+工业设计文档模板.

Derived from template: skill-test/AI+template.docx
Template analysis: skill-test/docx-template-analysis/

Usage:
    python3 md_to_json.py --input content.md --output content_data.json
"""

import argparse
import json
import re
import sys
from html.parser import HTMLParser

# ---------------------------------------------------------------------------
# Embedded JSON schema (content contract)
# ---------------------------------------------------------------------------
CONTENT_SCHEMA = {
    "type": "object",
    "required": ["document", "sections"],
    "properties": {
        "document": {
            "type": "object",
            "required": ["title"],
            "properties": {
                "title": {"type": "string"},
                "metadata": {"type": "object", "additionalProperties": True},
            },
        },
        "sections": {
            "type": "array",
            "items": {"$ref": "#/$defs/section"},
        },
        "figures": {
            "type": "array",
            "items": {"$ref": "#/$defs/figure"},
        },
    },
    "$defs": {
        "section": {
            "type": "object",
            "required": ["title", "level"],
            "properties": {
                "title": {"type": "string"},
                "level": {"type": "integer", "minimum": 1, "maximum": 9},
                "paragraphs": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "required": ["type", "text"],
                        "properties": {
                            "type": {"type": "string"},
                            "text": {"type": "string"},
                        },
                    },
                },
                "code_blocks": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "caption": {"type": "string"},
                            "language": {"type": "string"},
                            "code": {"type": "string"},
                        },
                    },
                },
                "tables": {
                    "type": "array",
                    "items": {"$ref": "#/$defs/table"},
                },
                "figures": {
                    "type": "array",
                    "items": {"$ref": "#/$defs/figure"},
                },
                "children": {
                    "type": "array",
                    "items": {"$ref": "#/$defs/section"},
                },
            },
        },
        "table": {
            "type": "object",
            "required": ["type", "rows"],
            "properties": {
                "type": {
                    "type": "string",
                    "enum": [
                        "macro_definition_table",
                        "struct_member_table",
                        "global_variable_table",
                        "interface_spec_table",
                    ],
                },
                "caption": {"type": "string"},
                "rows": {
                    "oneOf": [
                        {"type": "array"},
                        {"type": "object"},
                    ]
                },
            },
        },
        "figure": {
            "type": "object",
            "required": ["id", "source"],
            "properties": {
                "id": {"type": "string"},
                "source": {"type": "string"},
                "caption": {"type": "string"},
                "description": {"type": "string"},
            },
        },
    },
}

# ---------------------------------------------------------------------------
# Table type detection rules — derived from table_catalog.json
# ---------------------------------------------------------------------------
TABLE_TYPE_RULES = {
    "macro_definition_table": {
        "headers": [
            ("宏名称", "宏内容", "说明"),
        ],
        "keys": ("name", "value", "description"),
    },
    "struct_member_table": {
        "headers": [
            ("成员变量", "类型", "说明"),
        ],
        "keys": ("name", "type", "description"),
    },
    "global_variable_table": {
        "headers": [
            ("全局变量名称", "类型", "说明"),
            ("全局变量名称", "数据类型", "说明"),
        ],
        "keys": ("name", "type", "description"),
    },
    "interface_spec_table": {
        "headers": [
            ("接口原型",),  # first row label (merged table)
        ],
        "keys": None,  # merged table uses row-label parser
    },
}

# Row labels for interface_spec_table — from catalog prototype
INTERFACE_ROW_LABELS = {
    "接口原型": {"field": "prototype", "kind": "scalar"},
    "接口描述": {"field": "description", "kind": "scalar"},
    "接口参数": {
        "field": "parameters",
        "kind": "list",
        "sub_fields": ("type", "name", "description"),
        "sub_header_texts": ("数据类型", "参数名称", "参数说明"),
    },
    "返回值": {
        "field": "returns",
        "kind": "list",
        "sub_fields": ("type", "description"),
        "sub_header_texts": ("数据类型", "返回值说明"),
    },
    "使用限制": {"field": "constraints", "kind": "scalar"},
    "其它说明": {"field": "notes", "kind": "scalar"},
}

# ---------------------------------------------------------------------------
# Caption patterns — 表/图/代码 (use [\w\-]* not [\d\-]*)
# ---------------------------------------------------------------------------
TABLE_CAPTION_RE = re.compile(r"^\*?\*?表\s*[\w\-]*\s+(.+?)\*?\*?$")
FIGURE_CAPTION_RE = re.compile(r"^\*?\*?图\s*[\w\-]*\s+(.+?)\*?\*?$")
CODE_CAPTION_RE = re.compile(r"^\*?\*?代码\s*[\w\-]*\s+(.+?)\*?\*?$")


# ---------------------------------------------------------------------------
# HTML table parser (with rowspan virtual grid)
# ---------------------------------------------------------------------------
class HTMLTableParser(HTMLParser):
    """Parse HTML <table> into virtual grid with rowspan tracking."""

    def __init__(self):
        super().__init__()
        self.tables = []
        self._in_table = False
        self._in_row = False
        self._in_cell = False
        self._current_table = []
        self._current_row = []
        self._current_cell = {}
        self._cell_text = ""

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)
        if tag == "table":
            self._in_table = True
            self._current_table = []
        elif tag == "tr" and self._in_table:
            self._in_row = True
            self._current_row = []
        elif tag in ("td", "th") and self._in_row:
            self._in_cell = True
            self._cell_text = ""
            self._current_cell = {}
            if "rowspan" in attrs_dict:
                self._current_cell["rowspan"] = int(attrs_dict["rowspan"])
            if "colspan" in attrs_dict:
                self._current_cell["colspan"] = int(attrs_dict["colspan"])

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._in_cell:
            self._in_cell = False
            self._current_cell["text"] = self._cell_text.strip()
            self._current_row.append(self._current_cell)
        elif tag == "tr" and self._in_row:
            self._in_row = False
            self._current_table.append(self._current_row)
        elif tag == "table" and self._in_table:
            self._in_table = False
            self.tables.append(self._current_table)

    def handle_data(self, data):
        if self._in_cell:
            self._cell_text += data


def build_virtual_grid(physical_rows):
    """Build virtual grid with rowspan carryover."""
    virtual_grid = []
    rowspan_cells = {}  # col -> {text, remaining, cell}

    for physical_row in physical_rows:
        virtual_row = []
        col = 0
        cell_idx = 0

        while cell_idx < len(physical_row):
            # Skip columns occupied by rowspan from previous rows
            while col in rowspan_cells:
                rs = rowspan_cells[col]
                virtual_row.append({
                    "text": rs["text"],
                    "rowspan_continue": True,
                    "grid_span": rs.get("colspan", 1),
                })
                rs["remaining"] -= 1
                if rs["remaining"] <= 0:
                    del rowspan_cells[col]
                col += 1

            cell = physical_row[cell_idx]
            cell_copy = dict(cell)
            cell_copy.pop("rowspan", None)
            cell_copy.pop("colspan", None)
            virtual_row.append(cell_copy)

            rowspan = cell.get("rowspan", 1)
            if rowspan > 1:
                rowspan_cells[col] = {
                    "text": cell.get("text", ""),
                    "remaining": rowspan - 1,
                    "colspan": cell.get("colspan", 1),
                }

            col += 1
            cell_idx += 1

        # Fill remaining rowspan columns at end of row
        while col in rowspan_cells:
            rs = rowspan_cells[col]
            virtual_row.append({
                "text": rs["text"],
                "rowspan_continue": True,
                "grid_span": rs.get("colspan", 1),
            })
            rs["remaining"] -= 1
            if rs["remaining"] <= 0:
                del rowspan_cells[col]
            col += 1

        virtual_grid.append(virtual_row)

    return virtual_grid


def parse_html_tables(md_text):
    """Extract and parse all HTML <table> blocks from markdown."""
    parser = HTMLTableParser()
    # Find <table>...</table> blocks
    table_pattern = re.compile(r"<table[^>]*>.*?</table>", re.DOTALL | re.IGNORECASE)
    tables = []
    for match in table_pattern.finditer(md_text):
        parser.tables = []
        parser.feed(match.group())
        for raw_table in parser.tables:
            grid = build_virtual_grid(raw_table)
            tables.append(grid)
    return tables


# ---------------------------------------------------------------------------
# Pipe table parser
# ---------------------------------------------------------------------------
def parse_pipe_table(lines):
    """Parse Markdown pipe table lines into list of rows (list of cell texts)."""
    rows = []
    for line in lines:
        line = line.strip()
        if not line.startswith("|"):
            continue
        # Skip separator rows like |---|---|
        if re.match(r"^\|[\s\-:|]+\|$", line):
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        rows.append(cells)
    return rows


# ---------------------------------------------------------------------------
# Table type inference
# ---------------------------------------------------------------------------
def infer_table_type(header_row):
    """Infer table type from header row texts."""
    normalized = tuple(h.strip() for h in header_row)
    for ttype, rules in TABLE_TYPE_RULES.items():
        for variant in rules["headers"]:
            if len(variant) <= len(normalized):
                if normalized[:len(variant)] == variant:
                    return ttype
    return None


# ---------------------------------------------------------------------------
# Interface spec table parser (merged, from virtual grid)
# ---------------------------------------------------------------------------
def convert_interface_spec_table(grid):
    """Convert virtual grid to interface_spec_table structured data."""
    result = {
        "type": "interface_spec_table",
        "rows": {
            "prototype": "",
            "description": "",
            "parameters": [],
            "returns": [],
            "constraints": "",
            "notes": "",
        },
    }

    all_labels = set(INTERFACE_ROW_LABELS.keys())
    i = 0

    def is_subheader_row(row, meta):
        """Check if row matches sub-header texts for a list-type label."""
        texts = tuple(c.get("text", "").strip() for c in row[1:])
        return texts == meta.get("sub_header_texts")

    def is_label_subheader_row(row, meta):
        """Check if a label row also carries sub-header texts (rowspan=1 case)."""
        sub_hdrs = meta.get("sub_header_texts")
        if not sub_hdrs or len(row) < 1 + len(sub_hdrs):
            return False
        texts = tuple(c.get("text", "").strip() for c in row[1:1 + len(sub_hdrs)])
        return texts == sub_hdrs

    def extract_row_data(row, sub_fields, start_col=1):
        """Extract data from a row using sub_fields for column mapping."""
        entry = {}
        for idx, field in enumerate(sub_fields):
            col = start_col + idx
            if col < len(row):
                entry[field] = row[col].get("text", "").strip()
            else:
                entry[field] = ""
        return entry

    while i < len(grid):
        row = grid[i]
        if not row:
            i += 1
            continue

        label = row[0].get("text", "").strip()
        # Skip rowspan-continue label cells
        if row[0].get("rowspan_continue"):
            label = ""

        if label in INTERFACE_ROW_LABELS:
            meta = INTERFACE_ROW_LABELS[label]

            if meta["kind"] == "scalar":
                # Content is in the last cell (which may span multiple grid cols)
                if len(row) > 1:
                    result["rows"][meta["field"]] = row[-1].get("text", "").strip()
                i += 1

            elif meta["kind"] == "list":
                i += 1  # move past label row
                # Check if next row has rowspan continuation (rowspan > 1 case)
                next_has_continuation = (
                    i < len(grid) and grid[i] and
                    grid[i][0].get("rowspan_continue", False)
                )
                # Skip sub-header row if present on separate row
                if not next_has_continuation and i < len(grid) and is_subheader_row(grid[i], meta):
                    i += 1
                # Data column offset: 0 for rowspan=1, 1 for rowspan>1
                data_start_col = 0 if not next_has_continuation else 1
                # Collect data rows until next label
                while i < len(grid):
                    first = grid[i][0] if grid[i] else {}
                    first_text = first.get("text", "").strip()
                    first_is_continue = first.get("rowspan_continue", False)

                    # Stop if we hit a new label
                    if first_text in all_labels and not first_is_continue:
                        break
                    # Skip stray sub-headers
                    if is_subheader_row(grid[i], meta):
                        i += 1
                        continue
                    # Extract data
                    entry = extract_row_data(grid[i], meta["sub_fields"], start_col=data_start_col)
                    if any(v for v in entry.values()):
                        result["rows"][meta["field"]].append(entry)
                    i += 1
        else:
            i += 1

    return result


# ---------------------------------------------------------------------------
# Simple table parser (pipe tables → flat list)
# ---------------------------------------------------------------------------
def convert_simple_table(rows, table_type):
    """Convert pipe table rows to simple table JSON."""
    rules = TABLE_TYPE_RULES[table_type]
    keys = rules["keys"]
    if not keys:
        raise ValueError(f"Table type {table_type} requires merged table parser")

    result = {
        "type": table_type,
        "rows": [],
    }
    # Skip header row (first row)
    for row in rows[1:]:
        entry = {}
        for idx, key in enumerate(keys):
            entry[key] = row[idx] if idx < len(row) else ""
        result["rows"].append(entry)
    return result


# ---------------------------------------------------------------------------
# MD → JSON Converter
# ---------------------------------------------------------------------------
class MarkdownConverter:
    """Template-aware MD → JSON converter."""

    def __init__(self):
        self.result = {
            "document": {"title": ""},
            "sections": [],
            "figures": [],
        }
        self.section_stack = []  # (level, section_dict)
        self.current_paragraphs = []
        self.pending_caption = None
        self._html_tables = []
        self._html_table_positions = {}  # line_no -> grid index
        self._figure_counter = 0

    def convert(self, md_text):
        """Parse markdown text and return JSON-serializable dict."""
        # Pre-parse HTML tables
        self._html_tables = parse_html_tables(md_text)
        # Track which lines are inside HTML table blocks
        self._html_table_lines = set()
        html_idx = 0
        for match in re.finditer(r"<table[^>]*>", md_text, re.IGNORECASE):
            line_no = md_text[:match.start()].count("\n")
            self._html_table_positions[line_no] = html_idx
            html_idx += 1
            # Mark all lines until </table>
            end_match = re.search(
                r"</table>",
                md_text[match.start():],
                re.IGNORECASE,
            )
            if end_match:
                end_pos = match.start() + end_match.end()
                end_line = md_text[:end_pos].count("\n")
                for ln in range(line_no, end_line + 1):
                    self._html_table_lines.add(ln)

        lines = md_text.split("\n")
        i = 0
        while i < len(lines):
            line = lines[i]
            line_no = i

            # Skip lines inside HTML table blocks (handled separately)
            if line_no in self._html_table_lines:
                # Check if this is the start of an HTML table
                if line_no in self._html_table_positions:
                    grid_idx = self._html_table_positions[line_no]
                    grid = self._html_tables[grid_idx]
                    self._handle_html_table(grid)
                i += 1
                continue

            # Heading
            heading_match = re.match(r"^(#{1,6})\s+(.+)$", line)
            if heading_match:
                self._flush_paragraphs()
                level = len(heading_match.group(1))
                title = heading_match.group(2).strip()
                self._handle_heading(level, title)
                i += 1
                continue

            # Code block
            if line.strip().startswith("```"):
                self._flush_paragraphs()
                lang = line.strip()[3:].strip()
                code_lines = []
                i += 1
                while i < len(lines) and not lines[i].strip().startswith("```"):
                    code_lines.append(lines[i])
                    i += 1
                i += 1  # skip closing ```
                self._handle_code_block(lang, "\n".join(code_lines))
                continue

            # Pipe table
            if line.strip().startswith("|"):
                self._flush_paragraphs()
                table_lines = []
                while i < len(lines) and lines[i].strip().startswith("|"):
                    table_lines.append(lines[i])
                    i += 1
                self._handle_pipe_table(table_lines)
                continue

            # Figure: ![alt](src)
            fig_match = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)", line.strip())
            if fig_match:
                self._flush_paragraphs()
                alt = fig_match.group(1)
                src = fig_match.group(2)
                self._handle_figure(alt, src)
                i += 1
                continue

            # Regular paragraph
            stripped = line.strip()
            if stripped:
                self._handle_paragraph_line(stripped)
            else:
                # Blank line — flush accumulated paragraphs
                self._flush_paragraphs()
            i += 1

        self._flush_paragraphs()
        return self.result

    def _current_section(self):
        """Return the current leaf section dict, or None."""
        if self.section_stack:
            return self.section_stack[-1]
        return None

    def _handle_heading(self, level, title):
        """Process a heading line."""
        # Pop sections at same or deeper level
        while self.section_stack and self.section_stack[-1]["level"] >= level:
            self.section_stack.pop()

        section = {
            "title": title,
            "level": level,
            "paragraphs": [],
            "code_blocks": [],
            "tables": [],
            "figures": [],
            "children": [],
        }

        if self.section_stack:
            parent = self.section_stack[-1]
            parent["children"].append(section)
        else:
            self.result["sections"].append(section)

        self.section_stack.append(section)

        # Set document title from first H1
        if level == 1 and not self.result["document"]["title"]:
            self.result["document"]["title"] = title

    def _handle_paragraph_line(self, text):
        """Accumulate a paragraph line."""
        # Check for caption patterns
        table_cap = TABLE_CAPTION_RE.match(text)
        fig_cap = FIGURE_CAPTION_RE.match(text)
        code_cap = CODE_CAPTION_RE.match(text)

        if table_cap:
            cap_text = table_cap.group(1).strip().strip("*")
            self.pending_caption = {"type": "table", "text": f"表 {cap_text}"}
            return
        if fig_cap:
            cap_text = fig_cap.group(1).strip().strip("*")
            caption_text = f"图 {cap_text}"
            # Attach to most recent figure in current section + global
            section = self._current_section()
            if section and section.get("figures"):
                section["figures"][-1]["caption"] = caption_text
            if self.result.get("figures"):
                self.result["figures"][-1]["caption"] = caption_text
            # Do NOT set pending_caption or emit as paragraph
            return
        if code_cap:
            cap_text = code_cap.group(1).strip().strip("*")
            self.pending_caption = {"type": "code", "text": f"代码 {cap_text}"}
            return

        self.current_paragraphs.append(text)

    def _flush_paragraphs(self):
        """Flush accumulated paragraph lines to current section."""
        if not self.current_paragraphs:
            return
        section = self._current_section()
        if section is None:
            self.current_paragraphs = []
            return

        text = "\n".join(self.current_paragraphs)
        self.current_paragraphs = []

        section["paragraphs"].append({"type": "text", "text": text})

    def _handle_code_block(self, lang, code):
        """Process a code block."""
        section = self._current_section()
        if section is None:
            return

        block = {"language": lang, "code": code}
        if self.pending_caption and self.pending_caption["type"] == "code":
            block["caption"] = self.pending_caption["text"]
            section["paragraphs"].append({
                "type": "caption",
                "text": self.pending_caption["text"],
            })
            self.pending_caption = None
        section["code_blocks"].append(block)

    def _handle_figure(self, alt, src):
        """Process a figure reference."""
        section = self._current_section()
        if section is None:
            return

        self._figure_counter += 1
        fig = {
            "id": f"fig_{self._figure_counter}",
            "source": src,
            "description": alt,
        }

        # Figure caption comes AFTER the figure in MD, so don't consume
        # pending_caption here — it will be set by the next paragraph line
        section["figures"].append(fig)
        self.result["figures"].append(fig)

    def _handle_pipe_table(self, lines):
        """Process a Markdown pipe table."""
        rows = parse_pipe_table(lines)
        if len(rows) < 2:
            return

        header = rows[0]
        table_type = infer_table_type(header)
        if table_type is None:
            print(f"WARNING: Unknown table header: {header}", file=sys.stderr)
            return

        table_data = convert_simple_table(rows, table_type)

        # Attach caption if pending
        if self.pending_caption and self.pending_caption["type"] == "table":
            table_data["caption"] = self.pending_caption["text"]
            section = self._current_section()
            if section:
                section["paragraphs"].append({
                    "type": "caption",
                    "text": self.pending_caption["text"],
                })
            self.pending_caption = None

        section = self._current_section()
        if section:
            section["tables"].append(table_data)

    def _handle_html_table(self, grid):
        """Process an HTML table (virtual grid)."""
        if not grid:
            return

        # Check if this is an interface_spec_table
        first_label = grid[0][0].get("text", "").strip() if grid[0] else ""
        if first_label in INTERFACE_ROW_LABELS:
            table_data = convert_interface_spec_table(grid)
        else:
            # Try as simple table: first row is header
            header = [c.get("text", "").strip() for c in grid[0]]
            table_type = infer_table_type(header)
            if table_type is None:
                print(f"WARNING: Unknown HTML table header: {header}", file=sys.stderr)
                return
            rules = TABLE_TYPE_RULES[table_type]
            keys = rules["keys"]
            result_rows = []
            for row in grid[1:]:
                entry = {}
                for idx, key in enumerate(keys):
                    entry[key] = row[idx].get("text", "").strip() if idx < len(row) else ""
                result_rows.append(entry)
            table_data = {"type": table_type, "rows": result_rows}

        # Attach caption if pending
        if self.pending_caption and self.pending_caption["type"] == "table":
            table_data["caption"] = self.pending_caption["text"]
            section = self._current_section()
            if section:
                section["paragraphs"].append({
                    "type": "caption",
                    "text": self.pending_caption["text"],
                })
            self.pending_caption = None

        section = self._current_section()
        if section:
            section["tables"].append(table_data)

    def validate(self):
        """Validate output against CONTENT_SCHEMA."""
        try:
            import jsonschema
            jsonschema.validate(self.result, CONTENT_SCHEMA)
        except ImportError:
            # Fallback: basic structural checks
            assert "document" in self.result
            assert "title" in self.result["document"]
            assert "sections" in self.result
            for section in self.result["sections"]:
                assert "title" in section
                assert "level" in section
        except Exception as e:
            print(f"Schema validation failed: {e}", file=sys.stderr)
            sys.exit(1)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="MD → JSON converter for AI+设计文档模板")
    parser.add_argument("--input", required=True, help="Input Markdown file")
    parser.add_argument("--output", required=True, help="Output JSON file")
    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        md_text = f.read()

    converter = MarkdownConverter()
    result = converter.convert(md_text)
    converter.validate()

    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Converted: {args.input} → {args.output}")
    # Summary
    def count_sections(sections):
        n = len(sections)
        for s in sections:
            n += count_sections(s.get("children", []))
        return n

    total_sections = count_sections(result["sections"])
    total_tables = sum(
        len(s.get("tables", []))
        for s in result["sections"]
        for _ in [None]  # only top-level; walk recursively below
    )

    def count_tables(sections):
        n = sum(len(s.get("tables", [])) for s in sections)
        for s in sections:
            n += count_tables(s.get("children", []))
        return n

    total_figs = len(result.get("figures", []))
    print(f"  Sections: {total_sections}, Tables: {count_tables(result['sections'])}, Figures: {total_figs}")


if __name__ == "__main__":
    main()
