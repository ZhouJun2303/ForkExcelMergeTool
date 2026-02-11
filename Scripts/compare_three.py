# -*- coding: utf-8 -*-
"""临时脚本：读取 local / remote / merged 三个 Excel，对比并输出差异。"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from excel_io import (
    cell_str,
    get_sheet_names,
    key_str_normalized,
    load_sheet_rows,
    load_sheet_rows_full,
    ordered_keys_normalized,
    rows_to_dict_normalized,
)

def main():
    base_dir = r"d:\MyGit\ExcelMergeDiffTools_B\MergeExcelBackup"
    name = "Y英雄齿轮养成表"
    path_local = os.path.join(base_dir, name + "_local.xlsx")
    path_remote = os.path.join(base_dir, name + "_remote.xlsx")
    path_merged = os.path.join(base_dir, name + "_merged.xlsx")

    for p in (path_local, path_remote, path_merged):
        if not os.path.isfile(p):
            print("文件不存在:", p)
            return

    def load_book(path, full=False):
        wb = openpyxl.load_workbook(path, data_only=True)
        info = {}
        for sn in get_sheet_names(wb):
            ws = wb[sn]
            if full:
                rows, indices = load_sheet_rows_full(ws)
            else:
                rows = load_sheet_rows(ws)
                indices = list(range(1, len(rows) + 1))
            info[sn] = {
                "rows": rows,
                "indices": indices,
                "dict_norm": rows_to_dict_normalized(rows),
                "ordered_norm": ordered_keys_normalized(rows),
            }
        return info

    print("=== 使用 load_sheet_rows（跳过首列为空）===")
    local_sheets = load_book(path_local)
    remote_sheets = load_book(path_remote)
    merged_sheets = load_book(path_merged)

    all_sheets = sorted(set(local_sheets) | set(remote_sheets) | set(merged_sheets))
    for sn in all_sheets:
        loc = local_sheets.get(sn, {})
        rem = remote_sheets.get(sn, {})
        mrg = merged_sheets.get(sn, {})
        loc_keys = set(loc.get("ordered_norm", []))
        rem_keys = set(rem.get("ordered_norm", []))
        mrg_keys = set(mrg.get("ordered_norm", []))
        loc_rows = len(loc.get("rows", []))
        rem_rows = len(rem.get("rows", []))
        mrg_rows = len(mrg.get("rows", []))

        print("\n--- Sheet:", sn, "---")
        print("  local 行数:", loc_rows, "  key数:", len(loc_keys))
        print("  remote 行数:", rem_rows, "  key数:", len(rem_keys))
        print("  merged 行数:", mrg_rows, "  key数:", len(mrg_keys))

        only_local = loc_keys - rem_keys - mrg_keys
        only_remote = rem_keys - loc_keys - mrg_keys
        in_both_not_merged = (loc_keys | rem_keys) - mrg_keys
        in_merged_extra = mrg_keys - loc_keys - rem_keys
        if only_local:
            print("  仅在 local 的 key（且不在 merged）:", list(only_local)[:15], "..." if len(only_local) > 15 else "")
        if only_remote:
            print("  仅在 remote 的 key（且不在 merged）:", list(only_remote)[:15], "..." if len(only_remote) > 15 else "")
        if in_both_not_merged:
            print("  在 local 或 remote 但不在 merged:", list(in_both_not_merged)[:15], "..." if len(in_both_not_merged) > 15 else "")
        if in_merged_extra:
            print("  仅在 merged 的 key:", list(in_merged_extra)[:15], "..." if len(in_merged_extra) > 15 else "")

    print("\n=== 使用 load_sheet_rows_full（不跳过首列空，含表头）===")
    local_full = load_book(path_local, full=True)
    remote_full = load_book(path_remote, full=True)
    merged_full = load_book(path_merged, full=True)
    for sn in all_sheets[:1]:  # 只看第一个 sheet 的详情
        loc = local_full.get(sn, {})
        rem = remote_full.get(sn, {})
        mrg = merged_full.get(sn, {})
        print("\n--- Sheet (full):", sn, "---")
        print("  local 总行数:", len(loc.get("rows", [])))
        print("  remote 总行数:", len(rem.get("rows", [])))
        print("  merged 总行数:", len(mrg.get("rows", [])))
        # 首行是否为空
        for label, data in [("local", loc), ("remote", rem), ("merged", mrg)]:
            rows = data.get("rows", [])
            if rows:
                first = rows[0]
                print("  %s 首行首列:" % label, repr(first[0]) if first else "empty")

    # 重点：Data_HeroGearSkin 的 key 分布（是否有重复 key）
    sn = "Data_HeroGearSkin"
    if sn in all_sheets:
        print("\n=== Data_HeroGearSkin key 重复检查 ===")
        for label, sheets in [("local", local_sheets), ("remote", remote_sheets), ("merged", merged_sheets)]:
            rows = sheets.get(sn, {}).get("rows", [])
            keys = [key_str_normalized(r[0]) for r in rows if r and cell_str(r[0])]
            from collections import Counter
            cnt = Counter(keys)
            dups = {k: v for k, v in cnt.items() if v > 1}
            print("  %s: 总行数=%d, 唯一key=%d, 重复key及次数=%s" % (label, len(rows), len(cnt), dups))

if __name__ == "__main__":
    main()
