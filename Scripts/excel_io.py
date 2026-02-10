# -*- coding: utf-8 -*-
"""
Excel 读写与行/Key 抽象。
只做一件事：从 openpyxl 的 Workbook/Worksheet 读取或写入“按行、按首列 Key”的数据结构，
以及 Key 规范化、行相等判断，不包含合并算法或对比逻辑。
"""

import openpyxl

from config import SKIP_SHEET_PREFIX


# -----------------------------------------------------------------------------
# 单元格与 Key
# -----------------------------------------------------------------------------

def cell_str(c):
    """将单元格值转为用于比较的字符串；None 或空视为 ''。"""
    if c is None:
        return ""
    return str(c).strip()


def key_str_normalized(c):
    """
    合并时统一 Key：数字 1 与 1.0 视为同一行，避免漏检冲突。
    用于冲突检测与合并时以首列作为行唯一标识。
    """
    if c is None:
        return ""
    s = str(c).strip()
    if not s:
        return ""
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
        return str(f)
    except ValueError:
        return s


# -----------------------------------------------------------------------------
# Sheet 名过滤
# -----------------------------------------------------------------------------

def should_skip_sheet(name):
    """仅跳过以配置前缀（如 #）开头的表名；Sheet1/Sheet2 等会参与合并与对比。"""
    s = (name or "").strip()
    return s.startswith(SKIP_SHEET_PREFIX)


def get_sheet_names(wb):
    """返回需要参与合并/对比的 Sheet 名称列表（按需可保持顺序）。"""
    return [n for n in wb.sheetnames if not should_skip_sheet(n)]


# -----------------------------------------------------------------------------
# 合并单元格取值（openpyxl 仅左上格有值，其余为 None）
# -----------------------------------------------------------------------------

def get_merged_cell_value(ws, row, col):
    """
    若 (row, col) 落在合并区域内，返回该区域左上角的值；
    否则返回该格子的值。用于避免合并格导致首列“消失”。
    """
    try:
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                return ws.cell(row=rng.min_row, column=rng.min_col).value
    except Exception:
        pass
    return ws.cell(row=row, column=col).value


# -----------------------------------------------------------------------------
# 按行加载：每行为 list，可选跳过首列为空的行
# -----------------------------------------------------------------------------

def load_sheet_rows(ws, max_col=None):
    """
    加载 Sheet 所有数据行，每行为 list；第一列为空则跳过该行。
    用于对比模式或简单合并（不保留表头行号）。
    """
    if ws is None:
        return []
    rows = []
    if max_col is None:
        max_col = ws.max_column or 1
    for row in ws.iter_rows(min_row=1, max_col=max_col, values_only=True):
        row_list = list(row) if row else []
        while len(row_list) < max_col:
            row_list.append("")
        key = cell_str(row_list[0]) if row_list else ""
        if not key:
            continue
        rows.append(row_list)
    return rows


def load_sheet_rows_full(ws, max_col=None):
    """
    加载 Sheet 所有行（不跳过首列为空的行，表头会保留）。
    合并单元格会用区域左上角的值填满整区域。
    返回 (rows, row_indices)，row_indices 为每行在 Sheet 中的 1-based 行号。
    """
    if ws is None:
        return [], []
    rows = []
    row_indices = []
    if max_col is None:
        max_col = ws.max_column or 1
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_col=max_col, values_only=True), start=1):
        row_list = list(row) if row else []
        while len(row_list) < max_col:
            row_list.append("")
        rows.append(row_list)
        row_indices.append(row_idx)
    # 用合并区域左上角补齐空单元格
    try:
        for i in range(len(rows)):
            for col in range(max_col):
                cur = rows[i][col]
                if cur is None or (isinstance(cur, str) and cur.strip() == ""):
                    v = get_merged_cell_value(ws, row_indices[i], col + 1)
                    if v is not None:
                        rows[i][col] = v
    except Exception:
        pass
    return rows, row_indices


# -----------------------------------------------------------------------------
# 行列表 ↔ 字典 / 顺序（用于合并与对比）
# -----------------------------------------------------------------------------

def rows_to_dict(rows):
    """按第一列 Key 转为 dict[key] = row_list；空 key 跳过。"""
    d = {}
    for r in rows:
        k = cell_str(r[0]) if r else ""
        if k:
            d[k] = r
    return d


def ordered_keys(rows):
    """从 rows 按行出现顺序提取 key 列表（去重）。"""
    keys = []
    seen = set()
    for r in rows:
        k = cell_str(r[0]) if r else ""
        if k and k not in seen:
            seen.add(k)
            keys.append(k)
    return keys


def rows_to_dict_with_indices(rows, row_indices):
    """从 (rows, row_indices) 建 dict 和 key->行号；首列空的行 key 为 __row_N。"""
    d = {}
    key_to_row_index = {}
    for i, r in enumerate(rows):
        if i >= len(row_indices):
            break
        k = cell_str(r[0]) if r else ""
        if not k:
            k = "__row_%d" % row_indices[i]
        d[k] = r
        key_to_row_index[k] = row_indices[i]
    return d, key_to_row_index


def ordered_keys_from_rows_and_indices(rows, row_indices):
    """从 (rows, row_indices) 按行顺序得到 key 列表（首列空用 __row_N）。"""
    keys = []
    seen = set()
    for i, r in enumerate(rows):
        if i >= len(row_indices):
            break
        k = cell_str(r[0]) if r else ""
        if not k:
            k = "__row_%d" % row_indices[i]
        if k not in seen:
            seen.add(k)
            keys.append(k)
    return keys


def rows_to_dict_normalized(rows):
    """按首列建 dict，key 用 key_str_normalized，供冲突检测用。"""
    d = {}
    for r in rows:
        k = key_str_normalized(r[0]) if r else ""
        if k:
            d[k] = r
    return d


def ordered_keys_normalized(rows):
    """按行顺序返回规范化 key 列表（去重）。"""
    keys, seen = [], set()
    for r in rows:
        k = key_str_normalized(r[0]) if r else ""
        if k and k not in seen:
            seen.add(k)
            keys.append(k)
    return keys


# -----------------------------------------------------------------------------
# 行相等判断（用于冲突检测与合并）
# -----------------------------------------------------------------------------

def row_equal(a, b):
    """两行（list 或 tuple）按单元格字符串逐格比较是否相等。"""
    if a is b:
        return True
    if a is None or b is None:
        return a == b
    return [cell_str(c) for c in a] == [cell_str(c) for c in b]
