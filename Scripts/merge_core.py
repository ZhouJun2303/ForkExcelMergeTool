# -*- coding: utf-8 -*-
"""
三向合并核心（LOCAL + BASE + REMOTE → MERGED）。
只做一件事：给定三个 Excel 路径，执行三向合并并写入 MERGED，同时备份三方到 MergeExcelBackup；
不包含冲突选择（GUI 的“以本地为底、按选择覆盖”在 merge_gui 中实现）。
本模块供命令行回退或无 GUI 时使用；合并结果对新增/修改/冲突行做颜色标记。
"""

import os
import shutil
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill

from config import BACKUP_SUBDIR
from excel_io import (
    cell_str,
    get_sheet_names,
    load_sheet_rows,
    ordered_keys,
    row_equal,
    rows_to_dict,
)
from log_util import log


def _merge_sheet(base_rows, local_rows, remote_rows, base_ordered, local_ordered, remote_ordered, max_col):
    """
    单 Sheet 三向合并。
    顺序：先 BASE 原有 key，再 LOCAL 新增，再 REMOTE 新增。
    返回 (merged_rows, row_types)，row_types 为 {key: "新增"|"修改"|"冲突"}。
    """
    base_set = set(base_rows)
    local_set = set(local_rows)
    remote_set = set(remote_rows)
    all_keys = base_set | local_set | remote_set
    merged = []
    row_types = {}

    def process_key(key):
        base_row = base_rows.get(key)
        local_row = local_rows.get(key)
        remote_row = remote_rows.get(key)
        if local_row is not None and remote_row is not None:
            if row_equal(local_row, remote_row):
                merged.append(list(local_row))
                if base_row is None:
                    row_types[key] = "新增"
                elif not row_equal(base_row, local_row):
                    row_types[key] = "修改"
            elif base_row is None:
                merged.append(list(local_row))
                row_types[key] = "冲突"
            elif row_equal(base_row, local_row):
                merged.append(list(remote_row))
                row_types[key] = "修改"
            elif row_equal(base_row, remote_row):
                merged.append(list(local_row))
                row_types[key] = "修改"
            else:
                merged.append(list(local_row))
                row_types[key] = "冲突"
        elif local_row is not None:
            merged.append(list(local_row))
            row_types[key] = "新增" if base_row is None else ("修改" if not row_equal(base_row, local_row) else None)
        elif remote_row is not None:
            merged.append(list(remote_row))
            row_types[key] = "新增" if base_row is None else ("修改" if not row_equal(base_row, remote_row) else None)

    for key in base_ordered:
        if key in all_keys:
            process_key(key)
    for key in local_ordered:
        if key not in base_set and key in all_keys:
            process_key(key)
    for key in remote_ordered:
        if key not in base_set and key not in local_set and key in all_keys:
            process_key(key)

    for r in merged:
        while len(r) < max_col:
            r.append("")

    return merged, row_types


def do_merge(path_local, path_base, path_remote, path_merged):
    """
    执行三向合并：读取 LOCAL、BASE、REMOTE，写入 MERGED，并备份三方到 MERGED 所在目录的 BACKUP_SUBDIR。
    合并结果中对新增/修改/冲突行做颜色标记。
    返回 0 成功，2 异常（如缺 openpyxl）。
    """
    red_font = Font(color="FF0000", bold=True)
    green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

    merged_dir = os.path.dirname(os.path.abspath(path_merged))
    base_name = os.path.splitext(os.path.basename(path_merged))[0]
    backup_dir = os.path.join(merged_dir, BACKUP_SUBDIR)

    wb_local = openpyxl.load_workbook(path_local, data_only=True)
    wb_base = openpyxl.load_workbook(path_base, data_only=True)
    wb_remote = openpyxl.load_workbook(path_remote, data_only=True)

    seen = set()
    sheet_names = []
    for n in get_sheet_names(wb_base):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)
    for n in get_sheet_names(wb_local):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)
    for n in get_sheet_names(wb_remote):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    for sheet_name in sheet_names:
        ws_local = wb_local[sheet_name] if sheet_name in wb_local.sheetnames else None
        ws_base = wb_base[sheet_name] if sheet_name in wb_base.sheetnames else None
        ws_remote = wb_remote[sheet_name] if sheet_name in wb_remote.sheetnames else None

        max_col = max(
            (ws_local.max_column or 1) if ws_local else 1,
            (ws_base.max_column or 1) if ws_base else 1,
            (ws_remote.max_column or 1) if ws_remote else 1,
        )
        base_rows = {}
        local_rows = {}
        remote_rows = {}
        base_ordered = []
        local_ordered = []
        remote_ordered = []

        if ws_local:
            rows = load_sheet_rows(ws_local, max_col)
            local_rows = rows_to_dict(rows)
            local_ordered = ordered_keys(rows)
        if ws_base:
            rows = load_sheet_rows(ws_base, max_col)
            base_rows = rows_to_dict(rows)
            base_ordered = ordered_keys(rows)
        if ws_remote:
            rows = load_sheet_rows(ws_remote, max_col)
            remote_rows = rows_to_dict(rows)
            remote_ordered = ordered_keys(rows)

        merged_rows, row_types = _merge_sheet(
            base_rows, local_rows, remote_rows,
            base_ordered, local_ordered, remote_ordered,
            max_col,
        )
        if not merged_rows and not local_rows and not remote_rows:
            continue

        ws_out = wb_out.create_sheet(sheet_name)
        for r, row_list in enumerate(merged_rows, start=1):
            key_str = cell_str(row_list[0]) if row_list else ""
            rtype = row_types.get(key_str)
            fill = green_fill if rtype == "新增" else (yellow_fill if rtype == "修改" else (red_fill if rtype == "冲突" else None))
            for c, val in enumerate(row_list, start=1):
                cell = ws_out.cell(row=r, column=c, value=val)
                if fill:
                    cell.fill = fill
                if rtype == "冲突":
                    cell.font = red_font

    if not wb_out.sheetnames:
        wb_out.create_sheet("Data")
    if merged_dir:
        os.makedirs(merged_dir, exist_ok=True)
    wb_out.save(path_merged)
    wb_local.close()
    wb_base.close()
    wb_remote.close()
    wb_out.close()

    os.makedirs(backup_dir, exist_ok=True)
    shutil.copy2(path_local, os.path.join(backup_dir, base_name + "_local.xlsx"))
    shutil.copy2(path_remote, os.path.join(backup_dir, base_name + "_remote.xlsx"))
    shutil.copy2(path_merged, os.path.join(backup_dir, base_name + "_merged.xlsx"))

    log("合并完成 MERGED=%s 备份=%s" % (path_merged, backup_dir))
    print("OK: 合并完成。MERGED=%s 备份=%s" % (path_merged, backup_dir), file=sys.stdout)
    return 0
