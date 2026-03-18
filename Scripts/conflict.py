# -*- coding: utf-8 -*-
"""
冲突检测：三份 Excel（LOCAL、BASE、REMOTE）逐 Sheet 比较，得到“需选择”的冲突项与“仅本地有/仅线上有”项。
只做一件事：给定三个路径，返回 (conflicts, sheet_data, sheet_names)，供合并 GUI 展示与写回；
使用 data_only=False 加载，避免公式未计算导致误判相同。不包含合并写回逻辑。
"""

import openpyxl

from excel_io import (
    cell_str,
    col_equal,
    get_column_values,
    get_sheet_names,
    key_str_normalized,
    load_sheet_header,
    load_sheet_rows_full,
    row_equal,
)
from log_util import log_path


def _dict_and_order(rows, row_indices):
    """
    将 (rows, row_indices) 转为 dict、key->行号、顺序列表。
    首列空或 key 重复时使用 __row_N 作为 key，保证每行唯一且可写回时定位行号。
    """
    d = {}
    key_to_row = {}
    ord_list = []
    seen = set()
    for i, r in enumerate(rows):
        if i >= len(row_indices):
            break
        raw = cell_str(r[0]) if r else ""
        if not raw:
            k = "__row_%d" % row_indices[i]
        else:
            k = key_str_normalized(raw)
        if k in seen:
            k = "__row_%d" % row_indices[i]
        seen.add(k)
        d[k] = r
        key_to_row[k] = row_indices[i]
        ord_list.append(k)
    return d, key_to_row, ord_list


def _log_merge_diagnostic(sheet_name, n_l, n_b, n_r, n_keys, n_conflict):
    """写入单 Sheet 合并诊断到日志，便于排查“0 冲突”等问题。"""
    try:
        from datetime import datetime
        with open(log_path(), "a", encoding="utf-8") as f:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write("%s [MERGE] Sheet=%s 行数 local=%d base=%d remote=%d all_keys=%d 冲突=%d\n" % (
                ts, sheet_name, n_l, n_b, n_r, n_keys, n_conflict))
    except Exception:
        pass


def compute_conflicts(path_local, path_base, path_remote):
    """
    计算三向合并的冲突列表与每 Sheet 的数据结构。
    使用 data_only=False，避免公式单元格缓存为空导致误判。
    返回:
      - conflicts: list of dict，每项含 sheet, key, local_row, remote_row, base_row，
        以及可选 _only_local / _only_remote（仅一方有的行）；
      - sheet_data: dict[sheet_name] = { base_rows, local_rows, remote_rows, base_ordered, local_ordered, remote_ordered, key_to_row_l, key_to_row_r, key_to_row_b, max_col }；
      - sheet_names: list of str。
    """
    wb_l = openpyxl.load_workbook(path_local, data_only=False)
    wb_b = openpyxl.load_workbook(path_base, data_only=False)
    wb_r = openpyxl.load_workbook(path_remote, data_only=False)

    seen = set()
    sheet_names = []
    for n in get_sheet_names(wb_b):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)
    for n in get_sheet_names(wb_l):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)
    for n in get_sheet_names(wb_r):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)

    conflicts = []
    sheet_data = {}

    for sheet_name in sheet_names:
        ws_l = wb_l[sheet_name] if sheet_name in wb_l.sheetnames else None
        ws_b = wb_b[sheet_name] if sheet_name in wb_b.sheetnames else None
        ws_r = wb_r[sheet_name] if sheet_name in wb_r.sheetnames else None

        max_col_sheet = max(
            (ws_l.max_column or 1) if ws_l else 1,
            (ws_b.max_column or 1) if ws_b else 1,
            (ws_r.max_column or 1) if ws_r else 1,
        )
        rows_l, idx_l = load_sheet_rows_full(ws_l, max_col_sheet) if ws_l else ([], [])
        rows_b, idx_b = load_sheet_rows_full(ws_b, max_col_sheet) if ws_b else ([], [])
        rows_r, idx_r = load_sheet_rows_full(ws_r, max_col_sheet) if ws_r else ([], [])

        dict_l, key_to_row_l, ord_l = _dict_and_order(rows_l, idx_l)
        dict_b, key_to_row_b, ord_b = _dict_and_order(rows_b, idx_b)
        dict_r, key_to_row_r, ord_r = _dict_and_order(rows_r, idx_r)

        base_set = set(dict_b)
        local_set = set(dict_l)
        remote_set = set(dict_r)
        all_keys = base_set | local_set | remote_set

        n_conflict_sheet = 0
        for key in all_keys:
            row_l = dict_l.get(key)
            row_r = dict_r.get(key)
            row_b = dict_b.get(key)
            if row_l is not None and row_r is not None:
                if row_equal(row_l, row_r):
                    continue
                if row_b is None or (not row_equal(row_b, row_l) and not row_equal(row_b, row_r)):
                    n_conflict_sheet += 1
                    conflicts.append({
                        "sheet": sheet_name,
                        "key": key,
                        "local_row": row_l,
                        "remote_row": row_r,
                        "base_row": row_b,
                    })

        only_local_set = local_set - remote_set
        only_remote_set = remote_set - local_set
        for k in only_local_set:
            conflicts.append({
                "sheet": sheet_name,
                "key": k,
                "local_row": dict_l.get(k),
                "remote_row": None,
                "base_row": dict_b.get(k),
                "_only_local": True,
            })
        for k in only_remote_set:
            conflicts.append({
                "sheet": sheet_name,
                "key": k,
                "local_row": None,
                "remote_row": dict_r.get(k),
                "base_row": dict_b.get(k),
                "_only_remote": True,
            })

        sheet_data[sheet_name] = {
            "base_rows": dict_b,
            "local_rows": dict_l,
            "remote_rows": dict_r,
            "base_ordered": ord_b,
            "local_ordered": ord_l,
            "remote_ordered": ord_r,
            "key_to_row_l": key_to_row_l,
            "key_to_row_r": key_to_row_r,
            "key_to_row_b": key_to_row_b,
            "max_col": max(
                max(len(r) for r in rows_l) if rows_l else 1,
                max(len(r) for r in rows_b) if rows_b else 1,
                max(len(r) for r in rows_r) if rows_r else 1,
            ),
        }
        _log_merge_diagnostic(sheet_name, len(rows_l), len(rows_b), len(rows_r), len(all_keys), n_conflict_sheet)

    wb_l.close()
    wb_b.close()
    wb_r.close()
    return conflicts, sheet_data, sheet_names


def compute_conflicts_d(path_local, path_remote, path_base=None):
    """
    D 模式：二向冲突检测（冲突行 + 冲突列）。
    两边都有且内容不同即为冲突。path_base 可选，暂不用于过滤。
    返回 (conflict_rows, conflict_cols, sheet_names)。
    - conflict_rows: list of {"sheet", "key", "local_row", "remote_row", "kind": "row"}
    - conflict_cols: list of {"sheet", "key", "local_col", "remote_col", "kind": "column"}，key 为表头（列名）
    """
    wb_l = openpyxl.load_workbook(path_local, data_only=False)
    wb_r = openpyxl.load_workbook(path_remote, data_only=False)

    seen = set()
    sheet_names = []
    for n in get_sheet_names(wb_l):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)
    for n in get_sheet_names(wb_r):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)

    conflict_rows = []
    conflict_cols = []

    for sheet_name in sheet_names:
        ws_l = wb_l[sheet_name] if sheet_name in wb_l.sheetnames else None
        ws_r = wb_r[sheet_name] if sheet_name in wb_r.sheetnames else None
        if not ws_l or not ws_r:
            continue
        max_col = max(ws_l.max_column or 1, ws_r.max_column or 1)
        max_row = max(ws_l.max_row or 1, ws_r.max_row or 1)

        # 使用缓存加速合并单元格查找
        rows_l, idx_l = load_sheet_rows_full(ws_l, max_col, use_cache=True)
        rows_r, idx_r = load_sheet_rows_full(ws_r, max_col, use_cache=True)
        dict_l, _, ord_l = _dict_and_order(rows_l, idx_l)
        dict_r, _, ord_r = _dict_and_order(rows_r, idx_r)
        common_keys = set(dict_l) & set(dict_r)
        for key in common_keys:
            if not row_equal(dict_l[key], dict_r[key]):
                conflict_rows.append({
                    "sheet": sheet_name,
                    "key": key,
                    "local_row": dict_l[key],
                    "remote_row": dict_r[key],
                    "kind": "row",
                })

        header_l = load_sheet_header(ws_l, max_col)
        header_r = load_sheet_header(ws_r, max_col)
        norm_to_col_l = {}
        norm_to_col_r = {}
        for i, h in enumerate(header_l):
            if i >= max_col:
                break
            n = key_str_normalized(h or "")
            if n:
                norm_to_col_l[n] = i + 1
        for i, h in enumerate(header_r):
            if i >= max_col:
                break
            n = key_str_normalized(h or "")
            if n:
                norm_to_col_r[n] = i + 1
        common_headers = set(norm_to_col_l) & set(norm_to_col_r)
        for h_norm in common_headers:
            col_l = get_column_values(ws_l, norm_to_col_l[h_norm], max_row)
            col_r = get_column_values(ws_r, norm_to_col_r[h_norm], max_row)
            if not col_equal(col_l, col_r):
                conflict_cols.append({
                    "sheet": sheet_name,
                    "key": header_l[norm_to_col_l[h_norm] - 1] or header_r[norm_to_col_r[h_norm] - 1],
                    "local_col": col_l,
                    "remote_col": col_r,
                    "kind": "column",
                })

    wb_l.close()
    wb_r.close()
    return conflict_rows, conflict_cols, sheet_names
