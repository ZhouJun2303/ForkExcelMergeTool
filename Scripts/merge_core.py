# -*- coding: utf-8 -*-
"""
多模式合并核心（A 新增行 / B 新增列 / C 新增 Sheet / D 冲突选择 / E 智能）。
给定 LOCAL、BASE、REMOTE、MERGED 路径，按 mode 与 base_side 执行合并并写入 MERGED，可选备份。
A/B/C 仅用 LOCAL 与 REMOTE（base_side 指定基准）；D 可用 BASE 做三向冲突；E 顺序执行 A→B→C→D。
"""

import os
import shutil
import sys

import openpyxl
from openpyxl.styles import Font

from config import BACKUP_SUBDIR
from excel_io import (
    cell_str,
    get_sheet_names,
    get_column_values,
    key_str_normalized,
    load_sheet_header,
    load_sheet_rows,
    load_sheet_rows_full,
    merge_ordered_with_new_cols,
    merge_ordered_with_new_rows,
    ordered_keys,
    ordered_keys_normalized,
    row_equal,
    rows_to_dict,
    rows_to_dict_normalized,
)
from log_util import log


# ---------------------------------------------------------------------------
# 模式 A：新增行插入（基准 + 另一侧新增行，按前缀组末尾插入）
# ---------------------------------------------------------------------------

def _merge_mode_a_impl(path_base_side, path_other_side, path_merged, base_side):
    """以 base_side 文件为基准，将 other 中新增行按前缀插入，写入 path_merged。"""
    wb_base = openpyxl.load_workbook(path_base_side, data_only=True)
    wb_other = openpyxl.load_workbook(path_other_side, data_only=True)
    font_new = Font(color="008000")

    seen = set()
    sheet_names = []
    for n in get_sheet_names(wb_base):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)
    for n in get_sheet_names(wb_other):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    for sheet_name in sheet_names:
        ws_base = wb_base[sheet_name] if sheet_name in wb_base.sheetnames else None
        ws_other = wb_other[sheet_name] if sheet_name in wb_other.sheetnames else None
        if not ws_base and not ws_other:
            continue
        max_col = max(
            (ws_base.max_column or 1) if ws_base else 1,
            (ws_other.max_column or 1) if ws_other else 1,
        )
        base_rows = {}
        base_ordered = []
        other_rows = {}
        other_ordered = []
        if ws_base:
            rows = load_sheet_rows(ws_base, max_col)
            base_rows = rows_to_dict(rows)
            base_ordered = ordered_keys(rows)
        if ws_other:
            rows = load_sheet_rows(ws_other, max_col)
            other_rows = rows_to_dict(rows)
            other_ordered = ordered_keys(rows)

        base_set = set(base_rows)
        new_keys = [k for k in other_ordered if k not in base_set]
        merged_ordered = merge_ordered_with_new_rows(base_ordered, new_keys)
        base_set_plus_new = set(merged_ordered)

        merged_rows = []
        for k in merged_ordered:
            if k in base_rows:
                merged_rows.append(list(base_rows[k]))
            else:
                merged_rows.append(list(other_rows.get(k, [])))
            while len(merged_rows[-1]) < max_col:
                merged_rows[-1].append("")

        if not merged_rows and not base_rows and not other_rows:
            continue
        ws_out = wb_out.create_sheet(sheet_name)
        for r, row_list in enumerate(merged_rows, start=1):
            key_str = cell_str(row_list[0]) if row_list else ""
            is_new = key_str in base_set_plus_new and key_str in new_keys
            for c, val in enumerate(row_list, start=1):
                cell = ws_out.cell(row=r, column=c, value=val)
                if is_new:
                    cell.font = font_new

    wb_base.close()
    wb_other.close()
    if not wb_out.sheetnames:
        wb_out.create_sheet("Data")
    os.makedirs(os.path.dirname(os.path.abspath(path_merged)) or ".", exist_ok=True)
    wb_out.save(path_merged)
    wb_out.close()
    log("[Mode A] 新增行插入完成 MERGED=%s" % path_merged)


# ---------------------------------------------------------------------------
# 模式 B：新增列插入（基准 + 另一侧新增列，按列名前缀组末尾插入）
# ---------------------------------------------------------------------------

def _merge_mode_b_impl(path_base_side, path_other_side, path_merged, base_side):
    """以 base_side 为基准，将 other 中新增列按表头前缀插入，写入 path_merged。"""
    wb_base = openpyxl.load_workbook(path_base_side, data_only=True)
    wb_other = openpyxl.load_workbook(path_other_side, data_only=True)
    font_new = Font(color="008000")

    seen = set()
    sheet_names = []
    for n in get_sheet_names(wb_base):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)
    for n in get_sheet_names(wb_other):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    for sheet_name in sheet_names:
        ws_base = wb_base[sheet_name] if sheet_name in wb_base.sheetnames else None
        ws_other = wb_other[sheet_name] if sheet_name in wb_other.sheetnames else None
        if not ws_base and not ws_other:
            continue
        max_col_b = (ws_base.max_column or 1) if ws_base else 1
        max_col_o = (ws_other.max_column or 1) if ws_other else 1
        max_col = max(max_col_b, max_col_o)

        header_b = load_sheet_header(ws_base, max_col) if ws_base else []
        header_o = load_sheet_header(ws_other, max_col) if ws_other else []
        while len(header_b) < max_col:
            header_b.append("")
        while len(header_o) < max_col:
            header_o.append("")
        base_cols_ordered = [c for c in header_b if c]
        if not base_cols_ordered and header_b:
            base_cols_ordered = [cell_str(header_b[i]) or ("_col_%d" % (i + 1)) for i in range(len(header_b))]
        other_cols_ordered = [c for c in header_o if c]
        if not other_cols_ordered and header_o:
            other_cols_ordered = [cell_str(header_o[i]) or ("_col_%d" % (i + 1)) for i in range(len(header_o))]
        base_set = set(base_cols_ordered)
        new_cols = [c for c in other_cols_ordered if c not in base_set]
        merged_col_order = merge_ordered_with_new_cols(base_cols_ordered, new_cols)

        col_to_idx_base = {}
        for i, h in enumerate(header_b):
            if h:
                col_to_idx_base[h] = i + 1
        col_to_idx_other = {}
        for i, h in enumerate(header_o):
            if h:
                col_to_idx_other[h] = i + 1

        max_row = max(
            (ws_base.max_row or 1) if ws_base else 1,
            (ws_other.max_row or 1) if ws_other else 1,
        )
        ws_out = wb_out.create_sheet(sheet_name)
        for c, col_key in enumerate(merged_col_order, start=1):
            is_new_col = col_key in new_cols
            idx_b = col_to_idx_base.get(col_key)
            idx_o = col_to_idx_other.get(col_key)
            for r in range(1, max_row + 1):
                if idx_o is not None and (idx_b is None or is_new_col):
                    val = get_column_values(ws_other, idx_o, max_row)[r - 1] if ws_other else ""
                else:
                    val = get_column_values(ws_base, idx_b, max_row)[r - 1] if ws_base and idx_b is not None else ""
                cell = ws_out.cell(row=r, column=c, value=val)
                if is_new_col:
                    cell.font = font_new

    wb_base.close()
    wb_other.close()
    if not wb_out.sheetnames:
        wb_out.create_sheet("Data")
    os.makedirs(os.path.dirname(os.path.abspath(path_merged)) or ".", exist_ok=True)
    wb_out.save(path_merged)
    wb_out.close()
    log("[Mode B] 新增列插入完成 MERGED=%s" % path_merged)


# ---------------------------------------------------------------------------
# 模式 C：新增 Sheet 插入（基准 + 另一侧新增的 Sheet 整表追加）
# ---------------------------------------------------------------------------

def _merge_mode_c_impl(path_base_side, path_other_side, path_merged, base_side):
    """以 base_side 为基准，将 other 中多出的 Sheet 整表复制追加，写入 path_merged。"""
    wb_base = openpyxl.load_workbook(path_base_side, data_only=True)
    wb_other = openpyxl.load_workbook(path_other_side, data_only=True)

    base_sheets = set(get_sheet_names(wb_base))
    other_sheets = get_sheet_names(wb_other)
    new_sheets = [n for n in other_sheets if n not in base_sheets]

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    for name in get_sheet_names(wb_base):
        ws = wb_base[name]
        ws_new = wb_out.create_sheet(name)
        for r in ws.iter_rows():
            for c in r:
                ws_new.cell(row=c.row, column=c.column, value=c.value)
    for name in new_sheets:
        ws = wb_other[name]
        ws_new = wb_out.create_sheet(name)
        for r in ws.iter_rows():
            for c in r:
                ws_new.cell(row=c.row, column=c.column, value=c.value)

    wb_base.close()
    wb_other.close()
    if not wb_out.sheetnames:
        wb_out.create_sheet("Data")
    os.makedirs(os.path.dirname(os.path.abspath(path_merged)) or ".", exist_ok=True)
    wb_out.save(path_merged)
    wb_out.close()
    log("[Mode C] 新增 Sheet 插入完成 MERGED=%s new_sheets=%s" % (path_merged, new_sheets))


# ---------------------------------------------------------------------------
# 模式 D：冲突行/列选择（按 d_choices 写回 MERGED）
# ---------------------------------------------------------------------------

def _merge_mode_d_impl(path_local, path_remote, path_merged, path_base, d_choices, path_initial_merged=None):
    """
    d_choices: list of dict 每项 {"sheet", "key", "choice": "local"|"remote", "kind": "row"|"column"}。
    path_initial_merged: 若提供则以此为底写 path_merged，否则以 local 为底。E 模式传入 A+B+C 的结果。
    """
    if path_initial_merged and os.path.isfile(path_initial_merged):
        wb_out = openpyxl.load_workbook(path_initial_merged, data_only=False)
    else:
        wb_out = openpyxl.Workbook()
        wb_out.remove(wb_out.active)
        wb_local = openpyxl.load_workbook(path_local, data_only=False)
        for sheet_name in get_sheet_names(wb_local):
            ws_l = wb_local[sheet_name]
            ws_new = wb_out.create_sheet(sheet_name)
            for r in ws_l.iter_rows():
                for c in r:
                    ws_new.cell(row=c.row, column=c.column, value=c.value)
        wb_local.close()

    wb_local = openpyxl.load_workbook(path_local, data_only=False)
    wb_remote = openpyxl.load_workbook(path_remote, data_only=False)
    font_modified = Font(color="CC6600")
    font_conflict = Font(color="CC0000", bold=True)

    sheet_names = list(wb_out.sheetnames)

    choice_map = {}
    for item in d_choices or []:
        key = (item.get("sheet"), item.get("key"), item.get("kind"))
        choice_map[key] = item.get("choice", "local")

    for item in d_choices or []:
        sheet_name = item.get("sheet")
        key = item.get("key")
        choice = item.get("choice", "local")
        kind = item.get("kind", "row")
        if sheet_name not in wb_out.sheetnames:
            continue
        ws_out = wb_out[sheet_name]
        ws_l = wb_local[sheet_name] if sheet_name in wb_local.sheetnames else None
        ws_r = wb_remote[sheet_name] if sheet_name in wb_remote.sheetnames else None
        max_col = ws_out.max_column or 1
        max_row = ws_out.max_row or 1
        if kind == "row":
            rows_l, idx_l = load_sheet_rows_full(ws_l, max_col) if ws_l else ([], [])
            rows_r, idx_r = load_sheet_rows_full(ws_r, max_col) if ws_r else ([], [])
            key_to_row_l = {}
            key_to_row_r = {}
            for i, r in enumerate(rows_l):
                k = key_str_normalized(r[0]) if r else ""
                if k:
                    key_to_row_l[k] = idx_l[i] if i < len(idx_l) else i + 1
            for i, r in enumerate(rows_r):
                k = key_str_normalized(r[0]) if r else ""
                if k:
                    key_to_row_r[k] = idx_r[i] if i < len(idx_r) else i + 1
            row_idx = key_to_row_l.get(key) or key_to_row_r.get(key)
            if row_idx is None:
                continue
            source = ws_r if choice == "remote" else ws_l
            if not source:
                continue
            for c in range(1, max_col + 1):
                val = source.cell(row=row_idx, column=c).value
                ws_out.cell(row=row_idx, column=c, value=val)
                ws_out.cell(row=row_idx, column=c).font = font_modified
        else:
            header_l = load_sheet_header(ws_l, max_col) if ws_l else []
            header_r = load_sheet_header(ws_r, max_col) if ws_r else []
            col_idx_l = None
            col_idx_r = None
            for i, h in enumerate(header_l):
                if key_str_normalized(h) == key_str_normalized(key):
                    col_idx_l = i + 1
                    break
            for i, h in enumerate(header_r):
                if key_str_normalized(h) == key_str_normalized(key):
                    col_idx_r = i + 1
                    break
            col_idx = col_idx_l if choice == "local" else col_idx_r
            source = ws_l if choice == "local" else ws_r
            if col_idx is None or not source:
                continue
            for r in range(1, max_row + 1):
                val = source.cell(row=r, column=col_idx).value
                ws_out.cell(row=r, column=col_idx, value=val)
                ws_out.cell(row=r, column=col_idx).font = font_modified

    wb_local.close()
    wb_remote.close()
    if not wb_out.sheetnames:
        wb_out.create_sheet("Data")
    os.makedirs(os.path.dirname(os.path.abspath(path_merged)) or ".", exist_ok=True)
    wb_out.save(path_merged)
    wb_out.close()
    log("[Mode D] 冲突选择写回完成 MERGED=%s" % path_merged)


# ---------------------------------------------------------------------------
# 模式 E：智能（A → B → C → D，D 需 d_choices）
# ---------------------------------------------------------------------------

def _merge_mode_e_impl(path_local, path_base, path_remote, path_merged, base_side, d_choices):
    """顺序执行 A、B、C、D；前一步输出作为下一步输入，D 在 C 的结果上应用 d_choices 写入 path_merged。"""
    import tempfile
    path_base_side = path_local if base_side == "local" else path_remote
    path_other_side = path_remote if base_side == "local" else path_local

    fd, path_a = tempfile.mkstemp(suffix=".xlsx", prefix="merge_a_")
    os.close(fd)
    try:
        _merge_mode_a_impl(path_base_side, path_other_side, path_a, base_side)
        fd2, path_b = tempfile.mkstemp(suffix=".xlsx", prefix="merge_b_")
        os.close(fd2)
        try:
            _merge_mode_b_impl(path_a, path_other_side, path_b, base_side)
            fd3, path_c = tempfile.mkstemp(suffix=".xlsx", prefix="merge_c_")
            os.close(fd3)
            try:
                _merge_mode_c_impl(path_b, path_other_side, path_c, base_side)
                _merge_mode_d_impl(
                    path_local, path_remote, path_merged, path_base,
                    d_choices, path_initial_merged=path_c,
                )
            finally:
                try:
                    os.unlink(path_c)
                except Exception:
                    pass
        finally:
            try:
                os.unlink(path_b)
            except Exception:
                pass
    finally:
        try:
            os.unlink(path_a)
        except Exception:
            pass
    log("[Mode E] 智能合并完成 MERGED=%s" % path_merged)


# ---------------------------------------------------------------------------
# 选项驱动：C 删除行 / D 删除列 / F 删除 Sheet（path_in 为当前工作簿，other 为对照）
# ---------------------------------------------------------------------------

def _merge_delete_rows_impl(path_in, path_other_side, path_out):
    """从 path_in 中删除「在 base 侧有而 other 侧没有」的行（按首列 key）。"""
    wb_in = openpyxl.load_workbook(path_in, data_only=False)
    wb_other = openpyxl.load_workbook(path_other_side, data_only=True)
    for sheet_name in list(wb_in.sheetnames):
        if sheet_name not in wb_other.sheetnames:
            continue
        ws_in = wb_in[sheet_name]
        ws_o = wb_other[sheet_name]
        max_col = max(ws_in.max_column or 1, ws_o.max_column or 1)
        rows_o, _ = load_sheet_rows_full(ws_o, max_col)
        keys_other = set(key_str_normalized(r[0]) if r else "" for r in rows_o)
        keys_other.discard("")
        rows_in, idx_in = load_sheet_rows_full(ws_in, max_col)
        to_delete = []
        for i, r in enumerate(rows_in):
            k = key_str_normalized(r[0]) if r else ""
            if k and k not in keys_other:
                to_delete.append(idx_in[i] if i < len(idx_in) else i + 1)
        for row_idx in sorted(set(to_delete), reverse=True):
            ws_in.delete_rows(row_idx, 1)
    wb_other.close()
    os.makedirs(os.path.dirname(os.path.abspath(path_out)) or ".", exist_ok=True)
    wb_in.save(path_out)
    wb_in.close()
    log("[Option C] 删除行完成 %s" % path_out)


def _merge_delete_cols_impl(path_in, path_other_side, path_out):
    """从 path_in 中删除「在 base 有而 other 没有」的列（按表头）。"""
    wb_in = openpyxl.load_workbook(path_in, data_only=False)
    wb_other = openpyxl.load_workbook(path_other_side, data_only=True)
    for sheet_name in list(wb_in.sheetnames):
        if sheet_name not in wb_other.sheetnames:
            continue
        ws_in = wb_in[sheet_name]
        ws_o = wb_other[sheet_name]
        max_col = max(ws_in.max_column or 1, ws_o.max_column or 1)
        header_in = load_sheet_header(ws_in, max_col)
        header_o = load_sheet_header(ws_o, max_col)
        keys_other = set(key_str_normalized(h) for h in header_o if h)
        to_delete = []
        for c in range(len(header_in) - 1, -1, -1):
            h = header_in[c] if c < len(header_in) else ""
            k = key_str_normalized(h) if h else ""
            if k and k not in keys_other:
                to_delete.append(c + 1)
        for col_idx in sorted(to_delete, reverse=True):
            ws_in.delete_cols(col_idx, 1)
    wb_other.close()
    os.makedirs(os.path.dirname(os.path.abspath(path_out)) or ".", exist_ok=True)
    wb_in.save(path_out)
    wb_in.close()
    log("[Option D] 删除列完成 %s" % path_out)


def _merge_delete_sheets_impl(path_in, path_other_side, path_out):
    """从 path_in 中删除「在 base 有而 other 没有」的 Sheet。"""
    wb_in = openpyxl.load_workbook(path_in, data_only=False)
    wb_other = openpyxl.load_workbook(path_other_side, data_only=True)
    other_sheets = set(get_sheet_names(wb_other))
    wb_other.close()
    for name in list(wb_in.sheetnames):
        if name not in other_sheets:
            del wb_in[name]
    if not wb_in.sheetnames:
        wb_in.create_sheet("Data")
    os.makedirs(os.path.dirname(os.path.abspath(path_out)) or ".", exist_ok=True)
    wb_in.save(path_out)
    wb_in.close()
    log("[Option F] 删除 Sheet 完成 %s" % path_out)


# ---------------------------------------------------------------------------
# 按选项集合执行管道（A 不变=不增行 B 不变=不增列 C 删除行 D 删除列 E 新增Sheet F 删除Sheet G 冲突）
# ---------------------------------------------------------------------------

def _do_merge_by_options(path_local, path_base, path_remote, path_merged, options, base_side, d_choices):
    options = set(options or [])
    path_base_side = path_local if base_side == "local" else path_remote
    path_other_side = path_remote if base_side == "local" else path_local
    import tempfile
    fd, path_cur = tempfile.mkstemp(suffix=".xlsx", prefix="merge_opt_")
    os.close(fd)
    try:
        shutil.copy2(path_base_side, path_cur)
        if "A" not in options:
            _merge_mode_a_impl(path_cur, path_other_side, path_cur + ".tmp", base_side)
            os.replace(path_cur + ".tmp", path_cur)
        if "C" in options:
            _merge_delete_rows_impl(path_cur, path_other_side, path_cur + ".tmp")
            os.replace(path_cur + ".tmp", path_cur)
        if "B" not in options:
            _merge_mode_b_impl(path_cur, path_other_side, path_cur + ".tmp", base_side)
            os.replace(path_cur + ".tmp", path_cur)
        if "D" in options:
            _merge_delete_cols_impl(path_cur, path_other_side, path_cur + ".tmp")
            os.replace(path_cur + ".tmp", path_cur)
        if "E" in options:
            _merge_mode_c_impl(path_cur, path_other_side, path_cur + ".tmp", base_side)
            os.replace(path_cur + ".tmp", path_cur)
        if "F" in options:
            _merge_delete_sheets_impl(path_cur, path_other_side, path_cur + ".tmp")
            os.replace(path_cur + ".tmp", path_cur)
        if "G" in options and d_choices:
            _merge_mode_d_impl(path_local, path_remote, path_merged, path_base, d_choices, path_initial_merged=path_cur)
        else:
            os.makedirs(os.path.dirname(os.path.abspath(path_merged)) or ".", exist_ok=True)
            shutil.copy2(path_cur, path_merged)
        log("[Options] 管道完成 MERGED=%s options=%s" % (path_merged, options))
    finally:
        try:
            os.unlink(path_cur)
        except Exception:
            pass
        try:
            os.unlink(path_cur + ".tmp")
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 统一入口与备份
# ---------------------------------------------------------------------------

def do_merge(path_local, path_base, path_remote, path_merged, mode="E", base_side="local", d_choices=None, options=None):
    """
    执行合并并写入 MERGED，可选备份到 BACKUP_SUBDIR。
    options: 若提供则为多选集合 {"A","B","C","D","E","F","G"}，此时忽略 mode。
    A=行不变 B=列不变 C=删除行 D=删除列 E=新增Sheet F=删除Sheet G=冲突。
    mode: 未提供 options 时生效，"A"|"B"|"C"|"D"|"E"。
    base_side: "local"|"remote"
    d_choices: G 或冲突时需要；list of {"sheet", "key", "choice", "kind"}。
    返回 0 成功，2 异常。
    """
    path_base_side = path_local if base_side == "local" else path_remote
    path_other_side = path_remote if base_side == "local" else path_local
    try:
        if options is not None:
            _do_merge_by_options(path_local, path_base, path_remote, path_merged, options, base_side, d_choices or [])
        elif mode == "A":
            _merge_mode_a_impl(path_base_side, path_other_side, path_merged, base_side)
        elif mode == "B":
            _merge_mode_b_impl(path_base_side, path_other_side, path_merged, base_side)
        elif mode == "C":
            _merge_mode_c_impl(path_base_side, path_other_side, path_merged, base_side)
        elif mode == "D":
            _merge_mode_d_impl(path_local, path_remote, path_merged, path_base, d_choices or [])
        elif mode == "E":
            _merge_mode_e_impl(path_local, path_base, path_remote, path_merged, base_side, d_choices or [])
        else:
            log("未知 mode=%s，回退到 A" % mode)
            _merge_mode_a_impl(path_base_side, path_other_side, path_merged, base_side)
    except Exception as e:
        log("合并异常: %s" % e, is_error=True)
        import traceback
        log(traceback.format_exc(), is_error=True)
        print("ERROR: " + str(e), file=sys.stderr)
        return 2

    merged_dir = os.path.dirname(os.path.abspath(path_merged))
    base_name = os.path.splitext(os.path.basename(path_merged))[0]
    backup_dir = os.path.join(merged_dir, BACKUP_SUBDIR)
    os.makedirs(backup_dir, exist_ok=True)
    shutil.copy2(path_local, os.path.join(backup_dir, base_name + "_local.xlsx"))
    shutil.copy2(path_remote, os.path.join(backup_dir, base_name + "_remote.xlsx"))
    shutil.copy2(path_merged, os.path.join(backup_dir, base_name + "_merged.xlsx"))
    log("合并完成 MERGED=%s 备份=%s" % (path_merged, backup_dir))
    print("OK: 合并完成。MERGED=%s 备份=%s" % (path_merged, backup_dir), file=sys.stdout)
    return 0
