#!/usr/bin/env python3
"""
md_to_json.py — Markdown → JSON converter for the detailed design spec template.

Template: AI+工业操作系统研发项目详细设计规格说明书.docx
Schema:   content_contract.schema.json (in docx-template-analysis/)

Parses a Markdown document and emits a JSON object conforming to the schema.
The heading-to-schema mapping and table-type detection rules are derived from
the template's table_catalog.json and template_blueprint.md.

Usage:
    python3 md_to_json.py --input content.md --schema schema.json --output content.json

Limitations:
    - Only the 4 known table types are accepted; unknown header rows fail fast.
    - HTML tables must be well-formed XHTML-style (<tr>, <td>, <th> lower-case).
    - Inline Markdown formatting (bold/italic) is stripped in captions only;
      elsewhere it is preserved verbatim as paragraph text.
"""

import argparse
import json
import re
import sys
from html.parser import HTMLParser
from pathlib import Path

# ---------------------------------------------------------------------------
# Table type detection — derived from table_catalog.json header signatures.
#
# The catalog shows 3 "simple" table families with distinct header rows, plus
# the merged interface_spec_table whose first row is always "接口原型".
# ---------------------------------------------------------------------------

# Header row text -> (table_type, ordered column keys)
SIMPLE_TABLE_RULES = {
    ("宏名称", "宏内容", "说明"): (
        "macro_definition_table",
        ("name", "value", "desc"),
    ),
    ("成员变量", "类型", "说明"): (
        "struct_member_table",
        ("member", "type", "desc"),
    ),
    ("全局变量名称", "数据类型", "说明"): (
        "global_variable_table",
        ("name", "type", "desc"),
    ),
    ("字段", "类型", "用途"): (
        "struct_member_table",
        ("member", "type", "desc"),
    ),
}

# Row-label metadata for interface_spec_table (first-column text).
# Derived from the two prototype tables (table_5, table_6) in the catalog.
INTERFACE_ROW_LABELS = {
    "接口原型": {"field": "prototype", "kind": "scalar"},
    "接口描述": {"field": "description", "kind": "scalar"},
    "接口参数": {
        "field": "parameters",
        "kind": "list",
        "sub_fields": ("type", "name", "desc"),
        "sub_header": ("数据类型", "参数名称", "参数说明"),
    },
    "返回值": {
        "field": "returns",
        "kind": "list",
        "sub_fields": ("type", "desc"),
        "sub_header": ("数据类型", "返回值说明"),
    },
    "使用限制": {"field": "constraints", "kind": "scalar"},
    "其它说明": {"field": "notes", "kind": "scalar"},
}

# Heading regex patterns (chapter numbers are stripped).
# Heading level comes from the Markdown '#' count; the textual chapter number
# prefix (e.g. "4.7.6.1") is ignored for mapping purposes.
HEADING_NUMBER_RE = re.compile(r"^\s*[\d.]+\s*")

# Caption prefixes — template uses "表" for tables and "图" for figures.
TABLE_CAPTION_RE = re.compile(r"^\s*表\s*[\w\-]*\s*")
FIGURE_CAPTION_RE = re.compile(r"^\s*图\s*[\w\-]*\s*")
BOLD_CAPTION_RE = re.compile(r"^\s*\*\*\s*(?:表|图)\s*[\w\-]*\s*.*\*\*\s*$")

# ---------------------------------------------------------------------------
# Markdown tokenizer (minimal, sufficient for this template).
# ---------------------------------------------------------------------------

MD_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
MD_FENCE_RE = re.compile(r"^```(\w*)\s*$")
MD_PIPE_ROW_RE = re.compile(r"^\s*\|.*\|\s*$")
MD_PIPE_SEP_RE = re.compile(r"^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)+\|?\s*$")
MD_IMAGE_RE = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)\s*$")
MD_HTML_TAG_RE = re.compile(r"^\s*<(/?)(\w+)")


def strip_chapter_number(text):
    """Drop leading '4.7.6.1 '-style prefix from heading text."""
    return HEADING_NUMBER_RE.sub("", text).strip()


def strip_inline_bold(text):
    return re.sub(r"\*\*([^*]+)\*\*", r"\1", text).strip()


# ---------------------------------------------------------------------------
# HTML table parser with virtual-grid rowspan handling.
# ---------------------------------------------------------------------------


class _HTMLTableParser(HTMLParser):
    """Parse one <table>...</table> block into a virtual grid."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.physical_rows = []  # list of list of cell dicts
        self._cur_row = None
        self._cur_cell = None
        self._cur_text = []
        self._in_table = False
        self._in_thead = False

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "table":
            self._in_table = True
        elif tag == "thead":
            self._in_thead = True
        elif tag == "tbody":
            self._in_thead = False
        elif tag == "tr":
            self._cur_row = []
        elif tag in ("td", "th"):
            self._cur_cell = {
                "is_header": tag == "th" or self._in_thead,
                "colspan": int(attrs.get("colspan", 1)),
                "rowspan": int(attrs.get("rowspan", 1)),
            }
            self._cur_text = []

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cur_cell is not None:
            self._cur_cell["text"] = "".join(self._cur_text).strip()
            self._cur_row.append(self._cur_cell)
            self._cur_cell = None
        elif tag == "tr" and self._cur_row is not None:
            self.physical_rows.append(self._cur_row)
            self._cur_row = None
        elif tag == "table":
            self._in_table = False

    def handle_data(self, data):
        if self._cur_cell is not None:
            self._cur_text.append(data)


def _expand_colspan(row):
    """Expand each cell by its colspan so column indices line up."""
    expanded = []
    for cell in row:
        for i in range(cell["colspan"]):
            expanded.append(
                {
                    "text": cell["text"] if i == 0 else "",
                    "is_header": cell["is_header"],
                    "rowspan": cell["rowspan"] if i == 0 else 1,
                    "colspan_continued": i > 0,
                }
            )
    return expanded


def parse_html_table(html):
    """Parse an HTML <table> block into a uniform virtual grid.

    Each row in the returned grid has the same number of cells; rowspan
    carryover cells are injected with rowspan_continue=True.
    """
    parser = _HTMLTableParser()
    parser.feed(html)
    physical = [_expand_colspan(r) for r in parser.physical_rows]
    if not physical:
        return []

    grid = []
    pending = {}  # col -> {"text": ..., "remaining": int}
    num_cols = max(len(r) for r in physical) if physical else 0

    for row in physical:
        virtual = []
        col = 0
        idx = 0
        while col < num_cols:
            if col in pending:
                p = pending[col]
                virtual.append(
                    {
                        "text": p["text"],
                        "is_header": False,
                        "rowspan_continue": True,
                    }
                )
                p["remaining"] -= 1
                if p["remaining"] <= 0:
                    del pending[col]
                col += 1
                continue
            if idx < len(row):
                cell = row[idx]
                virtual.append(cell)
                if cell["rowspan"] > 1:
                    pending[col] = {
                        "text": cell["text"],
                        "remaining": cell["rowspan"] - 1,
                    }
                idx += 1
                col += 1
            else:
                break
        # fill trailing rowspan carryover
        while col in pending:
            p = pending[col]
            virtual.append(
                {"text": p["text"], "is_header": False, "rowspan_continue": True}
            )
            p["remaining"] -= 1
            if p["remaining"] <= 0:
                del pending[col]
            col += 1
        grid.append(virtual)
    return grid


# ---------------------------------------------------------------------------
# Markdown line-by-line parser.
# ---------------------------------------------------------------------------


class MarkdownParser:
    def __init__(self, text):
        self.lines = text.splitlines()
        self.pos = 0
        self.blocks = []  # list of parsed blocks

    def at_end(self):
        return self.pos >= len(self.lines)

    def peek(self):
        return self.lines[self.pos] if not self.at_end() else None

    def advance(self):
        line = self.lines[self.pos]
        self.pos += 1
        return line

    def parse_all(self):
        while not self.at_end():
            block = self.parse_block()
            if block is not None:
                self.blocks.append(block)
        return self.blocks

    def parse_block(self):
        line = self.peek()

        # blank line
        if line.strip() == "":
            self.advance()
            return None

        # ATX heading
        m = MD_HEADING_RE.match(line)
        if m:
            self.advance()
            level = len(m.group(1))
            text = strip_chapter_number(m.group(2))
            return {"kind": "heading", "level": level, "text": text, "raw": m.group(2)}

        # fenced code block
        m = MD_FENCE_RE.match(line)
        if m:
            lang = m.group(1) or ""
            self.advance()
            code_lines = []
            while not self.at_end() and not MD_FENCE_RE.match(self.peek()):
                code_lines.append(self.advance())
            if not self.at_end():
                self.advance()  # closing fence
            return {
                "kind": "code",
                "language": lang,
                "code": "\n".join(code_lines),
            }

        # HTML <table> block
        if MD_HTML_TAG_RE.match(line) and "<table" in line.lower():
            html_lines = [self.advance()]
            while not self.at_end():
                cur = self.peek()
                html_lines.append(cur)
                self.advance()
                if "</table>" in cur.lower():
                    break
            return {"kind": "html_table", "html": "\n".join(html_lines)}

        # Markdown image
        m = MD_IMAGE_RE.match(line)
        if m:
            self.advance()
            return {"kind": "image", "alt": m.group(1), "source": m.group(2)}

        # Markdown pipe table
        if MD_PIPE_ROW_RE.match(line):
            rows = []
            while not self.at_end() and MD_PIPE_ROW_RE.match(self.peek()):
                row_line = self.advance()
                if MD_PIPE_SEP_RE.match(row_line):
                    continue  # separator row, skip
                cells = [
                    c.strip()
                    for c in row_line.strip().strip("|").split("|")
                ]
                rows.append(cells)
            return {"kind": "pipe_table", "rows": rows}

        # HTML comment (e.g. <!-- FLOWCHART: name -->)
        if line.strip().startswith("<!--"):
            comment_lines = [self.advance()]
            if "-->" not in comment_lines[0]:
                while not self.at_end():
                    cur = self.advance()
                    comment_lines.append(cur)
                    if "-->" in cur:
                        break
            return {"kind": "comment", "text": "\n".join(comment_lines)}

        # plain paragraph (including possible caption line)
        text = self.advance()
        return {"kind": "paragraph", "text": text}


# ---------------------------------------------------------------------------
# Table classifier.
# ---------------------------------------------------------------------------


def _is_sub_header_row(cell_texts, sub_header_tokens):
    """Check if a row's cells contain sub-header tokens (e.g. '数据类型')."""
    return any(
        any(tok in cell for tok in sub_header_tokens)
        for cell in cell_texts
        if cell
    )


def convert_interface_pipe_table(rows):
    """Convert interface spec pipe-table rows to structured dict.

    In the Markdown source, the interface spec table uses embedded separator
    rows to visually separate sub-tables (parameters, returns) within the
    main table.  The tokenizer strips those separator rows, so we receive
    a flat list of data rows.  This function reconstructs the structure
    using the known row labels and sub-header tokens.
    """
    result = {
        "prototype": "",
        "description": "",
        "parameters": [],
        "returns": [],
        "constraints": "",
        "notes": "",
    }
    all_labels = set(INTERFACE_ROW_LABELS.keys())
    i = 0
    n = len(rows)
    while i < n:
        row = rows[i]
        label = row[0].strip() if row else ""

        if label not in INTERFACE_ROW_LABELS:
            i += 1
            continue

        meta = INTERFACE_ROW_LABELS[label]

        if meta["kind"] == "scalar":
            result[meta["field"]] = row[1].strip() if len(row) > 1 else ""
            i += 1
        elif meta["kind"] == "list":
            sub_hdr = meta["sub_header"]
            sf = meta["sub_fields"]
            num_sub = len(sf)

            # After the label, get the first non-empty content cell.
            content = ""
            for cell in row[1:]:
                if cell.strip():
                    content = cell.strip()
                    break

            # If content matches a sub-header token → inline sub-header row.
            is_sub_header_inline = (
                content and any(tok in content for tok in sub_hdr)
            )

            if content and not is_sub_header_inline:
                # Scalar content (e.g. "无", "不接收参数。").
                result[meta["field"]] = content
                i += 1
                continue

            # Sub-header present — skip to next row.
            i += 1

            # If sub-header was NOT inline, the next row is the sub-header.
            if not is_sub_header_inline and i < n:
                next_cells = rows[i]
                if _is_sub_header_row(next_cells, sub_hdr):
                    i += 1

            # Collect data rows until next known label or table end.
            while i < n:
                data_row = rows[i]
                data_label = data_row[0].strip() if data_row else ""
                if data_label in all_labels:
                    break
                entry = {}
                # Data cells start at index 1 (index 0 is empty continuation).
                data_cells = data_row[1:] if len(data_row) > 1 else []
                for j, k in enumerate(sf):
                    entry[k] = data_cells[j].strip() if j < len(data_cells) else ""
                result[meta["field"]].append(entry)
                i += 1
    return result


def classify_pipe_table(rows):
    """Determine table type from header row, or fail."""
    if not rows:
        raise ValueError("Empty pipe table")

    # Detect interface_spec_table — first cell of first row is "接口原型".
    if rows[0][0].strip() == "接口原型":
        return {
            "type": "interface_spec_table",
            "rows": convert_interface_pipe_table(rows),
        }

    header = tuple(rows[0])
    # Try exact match first.
    if header in SIMPLE_TABLE_RULES:
        ttype, keys = SIMPLE_TABLE_RULES[header]
    else:
        # Fall back: header might contain synonyms. Catalog shows header rows
        # are stable, so we also accept subsets by matching on column count
        # and known tokens.
        found = None
        for key, rule in SIMPLE_TABLE_RULES.items():
            if len(key) == len(header) and all(
                k in h or h in k for k, h in zip(key, header)
            ):
                found = rule
                break
        if found is None:
            raise ValueError(f"Unknown pipe table header: {header}")
        ttype, keys = found

    table_rows = []
    for r in rows[1:]:
        obj = {}
        for i, k in enumerate(keys):
            obj[k] = r[i] if i < len(r) else ""
        table_rows.append(obj)
    return {"type": ttype, "rows": table_rows}


def convert_interface_html_table(grid):
    """Convert an interface_spec_table virtual grid to structured JSON."""
    result = {
        "prototype": "",
        "description": "",
        "parameters": [],
        "returns": [],
        "constraints": "",
        "notes": "",
    }
    all_labels = set(INTERFACE_ROW_LABELS.keys())
    i = 0
    n = len(grid)
    while i < n:
        row = grid[i]
        label = row[0].get("text", "").strip() if row else ""

        # Skip pure header row (all cells are is_header) UNLESS cell[0]
        # matches a known label (e.g. "接口原型" in thead).
        if row and all(c.get("is_header") for c in row) and label not in INTERFACE_ROW_LABELS:
            i += 1
            continue

        if label not in INTERFACE_ROW_LABELS:
            i += 1
            continue

        meta = INTERFACE_ROW_LABELS[label]

        # rowspan_continue cell — already handled by prior group, skip.
        if row[0].get("rowspan_continue"):
            i += 1
            continue

        if meta["kind"] == "scalar":
            # Scalar value lives in cell[1] (the first cell after the label).
            # When colspan>1 on the content cell, expanded cells at index 2+
            # have empty text — we must use cell[1].
            value = ""
            if len(row) > 1:
                # Find the first non-continuation cell after the label.
                for c in row[1:]:
                    if not c.get("colspan_continued") and not c.get("rowspan_continue"):
                        value = c.get("text", "").strip()
                        break
            result[meta["field"]] = value
            i += 1

        elif meta["kind"] == "list":
            # Check if cell[1] contains a sub-header token (list group)
            # or scalar content (e.g. "无。").
            content = ""
            if len(row) > 1:
                for c in row[1:]:
                    if not c.get("colspan_continued") and not c.get("rowspan_continue"):
                        content = c.get("text", "").strip()
                        break

            # If content is a sub-header token, this is a list group with
            # inline sub-header (label rowspan=2, sub-header cells on same row).
            is_subheader = (
                content and any(sh in content for sh in meta["sub_header"])
            )

            if content and not is_subheader:
                # Scalar content (e.g. "无。" or "不接收参数。").
                result[meta["field"]] = content
                i += 1
                continue

            # List group — sub-header is either inline (same row) or on
            # the next row (when the label has rowspan=1).
            i += 1  # move past label row

            # If sub-header was NOT inline, check next row for sub-header.
            if not is_subheader and i < n:
                next_row = grid[i]
                sub_texts = [
                    c.get("text", "").strip()
                    for c in next_row
                    if not c.get("rowspan_continue")
                    and not c.get("colspan_continued")
                ]
                if any(t in meta["sub_header"] for t in sub_texts):
                    i += 1

            # Collect data rows until next label appears.
            while i < n:
                nxt = grid[i]
                nxt_label = nxt[0].get("text", "").strip() if nxt else ""
                if nxt_label in all_labels and not nxt[0].get("rowspan_continue"):
                    break
                if nxt[0].get("rowspan_continue"):
                    data_cells = [
                        c
                        for c in nxt
                        if not c.get("rowspan_continue")
                        and not c.get("colspan_continued")
                    ]
                    if data_cells and any(
                        c.get("text", "").strip() in meta["sub_header"]
                        for c in data_cells
                    ):
                        # Stray sub-header (repeated group) — skip.
                        i += 1
                        continue
                    entry = {}
                    sf = meta["sub_fields"]
                    values = [c.get("text", "").strip() for c in data_cells]
                    for j, k in enumerate(sf):
                        entry[k] = values[j] if j < len(values) else ""
                    result[meta["field"]].append(entry)
                else:
                    break
                i += 1
    return result


# ---------------------------------------------------------------------------
# Section tree builder.
# ---------------------------------------------------------------------------


def build_section_tree(blocks):
    """Walk parsed blocks and build the section tree + flat content lists."""
    root = {
        "document": {},
        "sections": [],
    }
    section_stack = []  # each entry: {"level": int, "section": dict}
    current_section = None
    pending_caption = None  # {"text": str, "kind": "table"|"figure"|"code"}

    def current():
        return current_section

    def push_section(level, title):
        nonlocal current_section
        new = {
            "title": title,
            "level": level,
            "paragraphs": [],
            "tables": [],
            "figures": [],
            "code_blocks": [],
            "children": [],
        }
        # Pop entries whose level is >= this one (CORRECT comparison).
        while section_stack and section_stack[-1]["level"] >= level:
            section_stack.pop()
        if section_stack:
            section_stack[-1]["section"]["children"].append(new)
        else:
            root["sections"].append(new)
        section_stack.append({"level": level, "section": new})
        current_section = new

    for block in blocks:
        kind = block["kind"]

        if kind == "heading":
            push_section(block["level"], block["text"])
            continue

        if current() is None:
            # Content before first heading — attach to a synthetic root section.
            push_section(1, "")

        if kind == "paragraph":
            text = block["text"]
            # Detect caption lines (table or figure).
            if BOLD_CAPTION_RE.match(text) or TABLE_CAPTION_RE.match(text):
                cap_kind = "figure" if FIGURE_CAPTION_RE.match(text) else "table"
                pending_caption = {
                    "text": strip_inline_bold(text),
                    "kind": cap_kind,
                }
                if cap_kind == "table":
                    # Table captions are emitted as caption blocks in paragraphs
                    # AND attached to the next table via tables[].caption.
                    current()["paragraphs"].append(
                        {"type": "caption", "text": pending_caption["text"],
                         "caption_kind": "table"}
                    )
                # figure captions are NOT emitted here — they attach to the
                # next image block as figures[].caption.
                continue
            if FIGURE_CAPTION_RE.match(text):
                pending_caption = {
                    "text": strip_inline_bold(text),
                    "kind": "figure",
                }
                continue
            current()["paragraphs"].append({"type": "paragraph", "text": text})
            continue

        if kind == "image":
            fig = {"source": block["source"], "alt": block.get("alt", "")}
            if pending_caption and pending_caption["kind"] == "figure":
                fig["caption"] = pending_caption["text"]
                pending_caption = None
            current()["figures"].append(fig)
            continue

        if kind == "code":
            cb = {
                "language": block.get("language", ""),
                "code": block["code"],
            }
            if pending_caption and pending_caption["kind"] == "code":
                cb["caption"] = pending_caption["text"]
                pending_caption = None
            current()["code_blocks"].append(cb)
            continue

        if kind == "pipe_table":
            try:
                tdata = classify_pipe_table(block["rows"])
            except ValueError as e:
                raise SystemExit(f"ERROR: {e}") from None
            if pending_caption and pending_caption["kind"] == "table":
                tdata["caption"] = pending_caption["text"]
                pending_caption = None
            current()["tables"].append(tdata)
            continue

        if kind == "html_table":
            grid = parse_html_table(block["html"])
            if not grid:
                continue
            # Detect interface_spec_table by first-cell label.
            first_label = grid[0][0].get("text", "").strip() if grid[0] else ""
            if first_label == "接口原型" or first_label in INTERFACE_ROW_LABELS:
                rows_data = convert_interface_html_table(grid)
                tdata = {"type": "interface_spec_table", "rows": rows_data}
            else:
                raise SystemExit(
                    f"ERROR: Unknown HTML table first label: {first_label!r}"
                )
            if pending_caption and pending_caption["kind"] == "table":
                tdata["caption"] = pending_caption["text"]
                pending_caption = None
            current()["tables"].append(tdata)
            continue

        if kind == "comment":
            # <!-- FLOWCHART: name --> — store as a paragraph marker so the
            # renderer knows where to insert a figure placeholder.
            text = block["text"].strip()
            current()["paragraphs"].append({"type": "paragraph", "text": text})
            continue

    return root


# ---------------------------------------------------------------------------
# Validation.
# ---------------------------------------------------------------------------


def validate_output(data, schema):
    try:
        import jsonschema
        jsonschema.validate(data, schema)
    except ImportError:
        # Fallback structural checks.
        assert "document" in data
        assert "sections" in data and isinstance(data["sections"], list)
        valid_types = {
            "macro_definition_table",
            "struct_member_table",
            "global_variable_table",
            "interface_spec_table",
        }

        def walk(sections):
            for s in sections:
                assert "title" in s and "level" in s
                for t in s.get("tables", []):
                    assert t["type"] in valid_types, f"Bad table type: {t['type']}"
                walk(s.get("children", []))

        walk(data["sections"])
    except Exception as e:
        raise SystemExit(f"Schema validation failed: {e}") from None


# ---------------------------------------------------------------------------
# Main.
# ---------------------------------------------------------------------------


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True, help="Input Markdown file")
    ap.add_argument("--schema", required=True, help="JSON schema file")
    ap.add_argument("--output", required=True, help="Output JSON file")
    args = ap.parse_args()

    md_text = Path(args.input).read_text(encoding="utf-8")
    schema = json.loads(Path(args.schema).read_text(encoding="utf-8"))

    parser = MarkdownParser(md_text)
    blocks = parser.parse_all()
    data = build_section_tree(blocks)

    # Top-level metadata defaults (cover page fields).
    data["document"] = {
        "title": data["document"].get("title", "详细设计规格说明书"),
        "subtitle": "",
        "material_number": "",
        "task_name": "",
        "task_number": "",
        "organization": "",
        "date_range": "",
        "date_line": "",
    }

    validate_output(data, schema)

    Path(args.output).write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
