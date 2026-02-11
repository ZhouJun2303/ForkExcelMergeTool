# -*- coding: utf-8 -*-
"""
按模式 A/B/C/D/E 各执行一次合并，使用 gen_merge_mode_tests 生成的 10 张 Excel。
输出到 TestData/_output/mode_*_merged.xlsx。
"""

import os
import sys


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root = os.path.dirname(script_dir)
    if root not in sys.path:
        sys.path.insert(0, root)
    sys.path.insert(0, os.path.join(root, "Scripts"))
    os.chdir(root)

    from merge_core import do_merge

    out_dir = os.path.join(script_dir, "_output")
    os.makedirs(out_dir, exist_ok=True)
    base_dir = script_dir

    modes = [
        ("A", "mode_a_local.xlsx", "mode_a_remote.xlsx", "mode_a_merged.xlsx", "local"),
        ("B", "mode_b_local.xlsx", "mode_b_remote.xlsx", "mode_b_merged.xlsx", "local"),
        ("C", "mode_c_local.xlsx", "mode_c_remote.xlsx", "mode_c_merged.xlsx", "local"),
        ("D", "mode_d_local.xlsx", "mode_d_remote.xlsx", "mode_d_merged.xlsx", "local"),
        ("E", "mode_e_local.xlsx", "mode_e_remote.xlsx", "mode_e_merged.xlsx", "local"),
    ]
    for mode, local_name, remote_name, merged_name, base_side in modes:
        path_local = os.path.join(base_dir, local_name)
        path_remote = os.path.join(base_dir, remote_name)
        path_merged = os.path.join(out_dir, merged_name)
        path_base = path_local
        if not os.path.isfile(path_local) or not os.path.isfile(path_remote):
            print("SKIP mode %s: 缺少 %s 或 %s，请先运行 gen_merge_mode_tests.py" % (mode, local_name, remote_name))
            continue
        d_choices = [] if mode in ("A", "B", "C") else [
            {"sheet": "Data", "key": "k1", "choice": "local", "kind": "row"},
            {"sheet": "Data", "key": "k2", "choice": "remote", "kind": "row"},
        ]
        if mode == "D" or mode == "E":
            d_choices = [
                {"sheet": "Data", "key": "k1", "choice": "local", "kind": "row"},
                {"sheet": "Data", "key": "k2", "choice": "remote", "kind": "row"},
            ]
        code = do_merge(path_local, path_base, path_remote, path_merged, mode=mode, base_side=base_side, d_choices=d_choices)
        if code == 0:
            print("OK mode %s -> %s" % (mode, path_merged))
        else:
            print("FAIL mode %s code=%d" % (mode, code))
            sys.exit(code)
    print("All mode tests passed.")
    sys.exit(0)


if __name__ == "__main__":
    main()
