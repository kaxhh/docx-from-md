#!/usr/bin/env python3
"""fix_md_tables.py — 预处理 Markdown 文件，修复管道表格。

Markdown 管道表格不支持单元格合并（colspan/rowspan）。当设计规格中的表格
需要合并单元格时，不同行用不同列数来表达合并意图——这不是格式错误，
而是管道表格格式的表达力不足。

处理策略：
- 矩形管道表格：原样输出
- 非矩形接口规格管道表格（以 '接口原型' 开头）：
  转换为 HTML <table>（带 colspan/rowspan），精确表达合并关系。
  下游 md_to_json.py 的 parse_html_table + convert_interface_html_table
  可以直接处理，与 HTML 格式的接口规格表（如说明书1）统一。
- 非矩形其他管道表格：以最大列数为基准补齐空单元格

用法：
    python3 fix_md_tables.py --input doc.md --output fixed.md
    python3 fix_md_tables.py --input doc.md          # 原地修改
"""

import argparse
import re
import sys
from pathlib import Path

PIPE_ROW_RE = re.compile(r"^\s*\|.*\|\s*$")
SEP_ROW_RE = re.compile(r"^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)+\|?\s*$")


def count_cells(row_line: str) -> int:
    """统计管道行中的单元格数量。"""
    stripped = row_line.strip().strip("|")
    if not stripped:
        return 0
    return len(stripped.split("|"))


def split_cells(row_line: str) -> list[str]:
    """拆分管道行为单元格列表。"""
    stripped = row_line.strip().strip("|")
    if not stripped:
        return []
    return [c.strip() for c in stripped.split("|")]


def join_cells(cells: list[str], leading_space: str = "") -> str:
    """将单元格列表拼回管道行。"""
    inner = " | ".join(cells)
    return f"{leading_space}| {inner} |"


def get_leading_space(row_line: str) -> str:
    """提取行前导空格。"""
    return row_line[: len(row_line) - len(row_line.lstrip())]


def fix_table_block(rows: list[str]) -> list[str]:
    """修复一个管道表格块中的非矩形行。

    策略：以最大列数为基准，补齐空单元格。不截断，避免丢失内容。
    """
    # 计算所有数据行的最大列数
    max_cols = max(count_cells(r) for r in rows if not SEP_ROW_RE.match(r))

    result = []
    for row in rows:
        if SEP_ROW_RE.match(row):
            # 分隔行也补齐到最大列数
            result.append(normalize_sep_row(row, max_cols))
        else:
            result.append(normalize_row(row, max_cols))

    return result


def normalize_sep_row(sep_line: str, target_cols: int) -> str:
    """将分隔行补齐到目标列数。"""
    leading = get_leading_space(sep_line)
    stripped = sep_line.strip().strip("|")
    cells = [c.strip() for c in stripped.split("|")] if stripped else []

    if len(cells) < target_cols:
        # 补齐分隔单元格
        cells.extend([":---"] * (target_cols - len(cells)))

    inner = " | ".join(cells)
    return f"{leading}| {inner} |"


def normalize_row(row_line: str, target_cols: int) -> str:
    """将一行管道行规范化到目标列数。"""
    cells = split_cells(row_line)
    leading = get_leading_space(row_line)

    if len(cells) > target_cols:
        # 截断多余单元格
        cells = cells[:target_cols]
    elif len(cells) < target_cols:
        # 补齐空单元格
        cells.extend([""] * (target_cols - len(cells)))

    return join_cells(cells, leading)


# ---------------------------------------------------------------------------
# Interface spec table: pipe → HTML conversion.
#
# 管道表格不支持 colspan/rowspan。说明书中不同行用不同列数，
# 是用减少 "|" 分隔符的方式表达单元格合并意图，不是格式错误。
# 正确做法：转换为 HTML <table>（带 colspan/rowspan），
# 与 HTML 格式的接口规格表（如说明书1）统一，
# 下游 md_to_json.py 的 parse_html_table + convert_interface_html_table
# 可以直接处理。
# ---------------------------------------------------------------------------

_INTERFACE_LABELS = frozenset([
    "接口原型", "接口描述", "接口参数", "返回值",
    "使用限制", "其它说明",
])

_SUB_HEADER_TOKENS = frozenset(["数据类型", "参数名称", "参数说明", "返回值说明"])


def _split_cells_safe(row_line: str) -> list[str]:
    """拆分管道行单元格，保证至少返回一个单元格。"""
    cells = split_cells(row_line)
    if not cells:
        cells = [""]
    return cells


def _is_interface_spec_block(block: list[str]) -> bool:
    """判断管道表格块是否为接口规格表（以 '接口原型' 开头）。"""
    for row in block:
        if SEP_ROW_RE.match(row):
            continue
        cells = _split_cells_safe(row)
        return cells[0].strip() == "接口原型"
    return False


def _merge_interface_spec_blocks(lines: list[str]) -> list[str]:
    """合并被空行拆散的接口规格管道表格块。

    说明书2 中，一个接口规格表的多段管道表格可能被空行分隔
    （例如参数段前有空行），导致被解析为多个独立块。
    本函数将以 '接口原型' 开头的块与其后续管道表格块合并。
    """
    segments = []  # list of (kind, data)
    # kind: "table" | "other"
    # data: list of lines

    i = 0
    while i < len(lines):
        if PIPE_ROW_RE.match(lines[i]):
            block = []
            while i < len(lines) and PIPE_ROW_RE.match(lines[i]):
                block.append(lines[i])
                i += 1
            segments.append(("table", block))
        else:
            segments.append(("other", [lines[i]]))
            i += 1

    # Merge: 接口规格块 + 空白段 + 后续管道表格块
    merged = []
    buf = None
    for kind, data in segments:
        if kind == "table":
            if buf is not None:
                # 追加到当前接口规格块
                buf.extend(data)
            elif _is_interface_spec_block(data):
                buf = list(data)
            else:
                merged.append(("table", data))
        else:
            if buf is not None:
                # 空白行：暂存，等待后续是否还有管道表格
                if all(line.strip() == "" for line in data):
                    buf.extend(data)
                else:
                    # 非空白非表格行，结束合并
                    merged.append(("table", buf))
                    buf = None
                    merged.append((kind, data))
            else:
                merged.append((kind, data))

    if buf is not None:
        merged.append(("table", buf))

    result = []
    for _, data in merged:
        result.extend(data)
    return result


def _convert_interface_spec_to_html(rows: list[str]) -> list[str]:
    """将接口规格管道表格转换为 HTML <table>（带 colspan/rowspan）。

    转换规则：
    1. 跳过嵌入的分隔行（| :--- | :--- |）
    2. 按首列标签分组连续行（空标签行归入前一组）
    3. 标签单元格设置 rowspan = 组内行数
    4. 列数少于最大列数的行，最后一个单元格设置 colspan 补齐
    """
    # 提取数据行，跳过分隔行
    data_rows = []
    for row in rows:
        if SEP_ROW_RE.match(row):
            continue
        data_rows.append(_split_cells_safe(row))

    if not data_rows:
        return []

    # 计算全局最大列数
    max_cols = max(len(r) for r in data_rows)
    if max_cols < 2:
        max_cols = 2

    # 按首列标签分组：空标签行归入前一组
    groups = []
    for cells in data_rows:
        label = cells[0].strip() if cells else ""
        if groups and not label:
            groups[-1]["rows"].append(cells)
        else:
            groups.append({"label": label, "rows": [cells]})

    # 生成 HTML
    out = ["<table>", "<thead>"]

    # 原型行（第一组，用 <th>）
    if groups:
        g = groups[0]
        row = g["rows"][0]
        content = row[1] if len(row) > 1 else ""
        cs = max_cols - 1
        out.append("<tr>")
        out.append('  <th colspan="1" rowspan="1">接口原型</th>')
        out.append(f'  <th colspan="{cs}" rowspan="1">{_esc(content)}</th>')
        out.append("</tr>")
    out.append("</thead>")
    out.append("<tbody>")

    # 后续行
    for g in groups[1:]:
        label = g["label"]
        g_rows = g["rows"]
        rs = len(g_rows)
        meta = _INTERFACE_LABELS_META.get(label)

        for idx, row in enumerate(g_rows):
            out.append("<tr>")
            if idx == 0:
                # 标签单元格（带 rowspan）
                out.append(f'  <td colspan="1" rowspan="{rs}">{_esc(label)}</td>')
                content_cells = row[1:]
                ncols = len(content_cells)
                # 检测标量数据行：只有第一个内容单元格有值，其余全空
                # 例如 "| 接口描述 | 很长的描述文本 |" 或 "| | 无 | | |"
                is_scalar_row = (
                    content_cells
                    and content_cells[0].strip()
                    and all(not c.strip() for c in content_cells[1:])
                )
                if is_scalar_row:
                    cs = max_cols - 1
                    out.append(
                        f'  <td colspan="{cs}" rowspan="1">'
                        f'{_esc(content_cells[0])}</td>'
                    )
                elif ncols < max_cols - 1:
                    # 列数不足但少于最大：最后一个单元格 colspan 补齐
                    cs = max_cols - ncols
                    for c in content_cells[:-1]:
                        out.append(f'  <td colspan="1" rowspan="1">{_esc(c)}</td>')
                    out.append(
                        f'  <td colspan="{cs}" rowspan="1">'
                        f'{_esc(content_cells[-1])}</td>'
                    )
                else:
                    for c in content_cells:
                        out.append(f'  <td colspan="1" rowspan="1">{_esc(c)}</td>')
            else:
                # 非首行：标签已被 rowspan 覆盖，跳过首列
                content_cells = row[1:]
                is_scalar_row = (
                    content_cells
                    and content_cells[0].strip()
                    and all(not c.strip() for c in content_cells[1:])
                )
                if is_scalar_row:
                    cs = max_cols - 1
                    out.append(
                        f'  <td colspan="{cs}" rowspan="1">'
                        f'{_esc(content_cells[0])}</td>'
                    )
                else:
                    for c in content_cells:
                        out.append(f'  <td colspan="1" rowspan="1">{_esc(c)}</td>')
            out.append("</tr>")

    out.append("</tbody>")
    out.append("</table>")
    return out


_INTERFACE_LABELS_META = {
    "接口参数": {"sub_header": ["数据类型", "参数名称", "参数说明"]},
    "返回值": {"sub_header": ["数据类型", "返回值说明"]},
}


def _esc(text: str) -> str:
    """HTML 转义基本字符。"""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def process_markdown(text: str) -> tuple[str, list[str]]:
    """处理 Markdown 文本，修复管道表格。

    处理策略：
    - 矩形管道表格：原样输出
    - 非矩形接口规格管道表格（以 '接口原型' 开头）：
      转换为 HTML <table>（带 colspan/rowspan），精确表达合并关系
    - 非矩形其他管道表格：补齐空单元格（旧逻辑）

    返回 (fixed_text, warnings)。
    """
    lines = text.splitlines()

    # 预处理：合并被空行拆散的接口规格管道表格块
    lines = _merge_interface_spec_blocks(lines)

    warnings = []
    i = 0
    output_lines = []

    while i < len(lines):
        line = lines[i]

        if PIPE_ROW_RE.match(line):
            # 收集连续的管道行
            block_start = i
            block = []
            while i < len(lines) and PIPE_ROW_RE.match(lines[i]):
                block.append(lines[i])
                i += 1

            # 检查是否矩形
            col_counts = [count_cells(r) for r in block if not SEP_ROW_RE.match(r)]
            unique_counts = set(col_counts)

            if len(unique_counts) <= 1:
                # 矩形表格，原样输出
                output_lines.extend(block)
            elif _is_interface_spec_block(block):
                # 非矩形接口规格表：转换为 HTML
                sep_counts = [count_cells(r) for r in block if SEP_ROW_RE.match(r)]
                warnings.append(
                    f"行 {block_start + 1}-{i}: "
                    f"接口规格表列数不一致 {sorted(unique_counts)}，"
                    f"已转换为 HTML 表格（colspan/rowspan）"
                )
                html_lines = _convert_interface_spec_to_html(block)
                output_lines.extend(html_lines)
            else:
                # 非矩形非接口规格表：补齐空单元格
                sep_counts = [count_cells(r) for r in block if SEP_ROW_RE.match(r)]
                warnings.append(
                    f"行 {block_start + 1}-{i}: "
                    f"数据行列数不一致 {sorted(unique_counts)}，"
                    f"分隔行列数 {sorted(set(sep_counts))}，已补齐"
                )
                fixed_block = fix_table_block(block)
                output_lines.extend(fixed_block)
        else:
            output_lines.append(line)
            i += 1

    return "\n".join(output_lines), warnings


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True, type=Path, help="输入 Markdown 文件")
    ap.add_argument("--output", type=Path, help="输出文件（默认原地修改）")
    ap.add_argument("--dry-run", action="store_true", help="仅检查不修改")
    args = ap.parse_args()

    if not args.input.exists():
        print(f"错误：文件不存在：{args.input}", file=sys.stderr)
        sys.exit(1)

    text = args.input.read_text(encoding="utf-8")
    fixed, warnings = process_markdown(text)

    for w in warnings:
        print(f"警告：{w}", file=sys.stderr)

    if not warnings:
        print("所有管道表格均为矩形，无需修改。")

    if args.dry_run:
        print("（dry-run 模式，未写入文件）")
        sys.exit(1)

    output_path = args.output or args.input
    output_path.write_text(fixed, encoding="utf-8")
    print(f"已写入：{output_path}")


if __name__ == "__main__":
    main()
