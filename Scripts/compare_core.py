# -*- coding: utf-8 -*-
"""
二向对比核心（A vs B → 对比表 Excel）。
只做一件事：给定两个 Excel 路径，计算差异、生成对比 Excel（含颜色标记），可选打开文件。
"""

import os
import subprocess
import sys

import openpyxl
from openpyxl.styles import PatternFill

from config import COMPARE_SUFFIX
from excel_io import cell_str, get_sheet_names, load_sheet_rows, rows_to_dict
from log_util import log


def get_compare_data(path_a, path_b):
    """
    计算 A、B 两个 Excel 的差异。
    返回 (path_out, sheet_names, diff_rows)：
      - path_out: 将生成的对比文件路径（与 path_a 同目录，文件名为 {base}_compare.xlsx）
      - sheet_names: 参与对比的 Sheet 名列表
      - diff_rows: [(sheet_name, key, status, str_a, str_b), ...]，status 为 "A独有"|"B新增"|"修改"|"相同"
    """
    out_dir = os.path.dirname(os.path.abspath(path_a))
    base_name = os.path.splitext(os.path.basename(path_a))[0]
    path_out = os.path.join(out_dir, base_name + COMPARE_SUFFIX + ".xlsx")

    wb_a = openpyxl.load_workbook(path_a, data_only=True)
    wb_b = openpyxl.load_workbook(path_b, data_only=True)

    seen = set()
    sheet_names = []
    for n in get_sheet_names(wb_a):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)
    for n in get_sheet_names(wb_b):
        if n not in seen:
            seen.add(n)
            sheet_names.append(n)
    if not sheet_names:
        for wb in (wb_a, wb_b):
            for n in (wb.sheetnames or []):
                sheet_names.append(n)
                break
            if sheet_names:
                break
        if not sheet_names:
            sheet_names = ["Sheet1"]

    diff_rows = []
    for sheet_name in sheet_names:
        ws_a = wb_a[sheet_name] if sheet_name in wb_a.sheetnames else None
        ws_b = wb_b[sheet_name] if sheet_name in wb_b.sheetnames else None
        rows_a = load_sheet_rows(ws_a) if ws_a else []
        rows_b = load_sheet_rows(ws_b) if ws_b else []
        dict_a = rows_to_dict(rows_a)
        dict_b = rows_to_dict(rows_b)
        all_keys = sorted(set(dict_a) | set(dict_b))
        max_col = max(
            max(len(r) for r in rows_a) if rows_a else 0,
            max(len(r) for r in rows_b) if rows_b else 0,
            1,
        )
        for key in all_keys:
            row_a = dict_a.get(key)
            row_b = dict_b.get(key)
            vals_a = [cell_str(c) for c in (row_a or [])]
            vals_b = [cell_str(c) for c in (row_b or [])]
            while len(vals_a) < max_col:
                vals_a.append("")
            while len(vals_b) < max_col:
                vals_b.append("")
            str_a = " | ".join(vals_a)
            str_b = " | ".join(vals_b)
            if row_a is None:
                status = "B新增"
            elif row_b is None:
                status = "A独有"
            elif str_a != str_b:
                status = "修改"
            else:
                status = "相同"
            diff_rows.append((sheet_name, key, status, str_a, str_b))

    wb_a.close()
    wb_b.close()
    return path_out, sheet_names, diff_rows


def write_compare_excel(path_out, sheet_names, diff_rows, open_file=False):
    """根据 diff 数据写入对比 Excel（带颜色），可选用系统默认程序打开。"""
    green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    row_idx = {}
    for sheet_name in sheet_names:
        ws_out = wb_out.create_sheet(sheet_name[:31])
        ws_out.cell(row=1, column=1, value="[Key]")
        ws_out.cell(row=1, column=2, value="[A-LEFT]")
        ws_out.cell(row=1, column=3, value="[B-RIGHT]")
        ws_out.cell(row=1, column=4, value="[Status]")
        row_idx[sheet_name] = 2

    for sheet_name, key, status, str_a, str_b in diff_rows:
        ws_out = wb_out[sheet_name[:31]]
        r = row_idx[sheet_name]
        ws_out.cell(row=r, column=1, value=key)
        ws_out.cell(row=r, column=2, value=str_a)
        ws_out.cell(row=r, column=3, value=str_b)
        ws_out.cell(row=r, column=4, value=status)
        fill = green_fill if status == "B新增" else (red_fill if status == "A独有" else (yellow_fill if status == "修改" else None))
        if fill:
            for col in range(1, 5):
                ws_out.cell(row=r, column=col).fill = fill
        row_idx[sheet_name] = r + 1

    wb_out.save(path_out)
    wb_out.close()

    if open_file:
        try:
            if sys.platform == "win32":
                os.startfile(path_out)
            elif sys.platform == "darwin":
                subprocess.run(["open", path_out], check=False)
            else:
                subprocess.run(["xdg-open", path_out], check=False)
        except Exception:
            pass
    return 0


def do_compare(path_a, path_b, open_file=True):
    """
    二向对比：生成对比 Excel 并可选打开。
    返回 0 成功，2 异常（如缺 openpyxl 或 get_compare_data 失败）。
    """
    result = get_compare_data(path_a, path_b)
    if result[0] is None:
        return 2
    path_out, sheet_names, diff_rows = result
    log("对比模式 输出: %s" % path_out)
    write_compare_excel(path_out, sheet_names, diff_rows, open_file=open_file)

    if open_file:
        try:
            if sys.platform == "win32":
                os.startfile(path_out)
            elif sys.platform == "darwin":
                subprocess.run(["open", path_out], check=False)
            else:
                subprocess.run(["xdg-open", path_out], check=False)
        except Exception:
            pass

    print("OK: 对比已生成 %s" % path_out, file=sys.stdout)
    return 0
