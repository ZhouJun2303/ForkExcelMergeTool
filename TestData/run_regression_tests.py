# -*- coding: utf-8 -*-
"""
针对核心 bugfix 的轻量回归测试。
不依赖 pytest，输出写入 TestData/_output/regression。
"""

import os
import shutil
import sys
import time

import openpyxl


def _setup_paths():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root = os.path.dirname(script_dir)
    if root not in sys.path:
        sys.path.insert(0, root)
    scripts_dir = os.path.join(root, "Scripts")
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    os.chdir(root)
    out_dir = os.path.join(script_dir, "_output", "regression")
    if os.path.isdir(out_dir):
        shutil.rmtree(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    return root, out_dir


def _write_xlsx(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=val)
    wb.save(path)
    wb.close()


def _rows(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Data"]
    values = [
        [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        for r in range(1, ws.max_row + 1)
    ]
    wb.close()
    return values


def _assert(cond, msg):
    if not cond:
        raise AssertionError(msg)


def test_compare_opens_once(out_dir):
    import compare_core

    left = os.path.join(out_dir, "compare_left.xlsx")
    right = os.path.join(out_dir, "compare_right.xlsx")
    _write_xlsx(left, [["Key", "V"], ["a", "1"]])
    _write_xlsx(right, [["Key", "V"], ["a", "2"]])

    calls = []
    original = compare_core.write_compare_excel

    def wrapped(path_out, sheet_names, diff_rows, open_file=False):
        calls.append(open_file)
        return original(path_out, sheet_names, diff_rows, open_file=False)

    compare_core.write_compare_excel = wrapped
    try:
        code = compare_core.do_compare(left, right, open_file=True)
    finally:
        compare_core.write_compare_excel = original

    _assert(code == 0, "compare should succeed")
    _assert(calls == [True], "do_compare should delegate opening exactly once")


def test_backup_same_second(out_dir):
    from backup_util import create_merge_backup

    local = os.path.join(out_dir, "backup_local.xlsx")
    remote = os.path.join(out_dir, "backup_remote.xlsx")
    merged = os.path.join(out_dir, "backup_merged.xlsx")
    backup_root = os.path.join(out_dir, "backup_root")
    for path, value in [(local, "l"), (remote, "r"), (merged, "m")]:
        _write_xlsx(path, [["Key", "V"], ["a", value]])

    infos = [
        create_merge_backup(local, remote, merged, backup_root=backup_root, timestamp="20260101_010101")
        for _ in range(3)
    ]
    dirs = [info["dir"] for info in infos]
    _assert(len(set(dirs)) == 3, "same-second backups should create unique dirs")
    _assert(all(os.path.isdir(d) for d in dirs), "backup dirs should exist")


def test_delete_conflict_choices(out_dir):
    from merge_core import do_merge

    base = os.path.join(out_dir, "delete_base.xlsx")
    local_deleted = os.path.join(out_dir, "delete_local_deleted.xlsx")
    remote_kept = os.path.join(out_dir, "delete_remote_kept.xlsx")
    local_kept = os.path.join(out_dir, "delete_local_kept.xlsx")
    remote_deleted = os.path.join(out_dir, "delete_remote_deleted.xlsx")

    _write_xlsx(base, [["Key", "V"], ["keep", "base"], ["gone", "base"]])
    _write_xlsx(local_deleted, [["Key", "V"], ["keep", "local"]])
    _write_xlsx(remote_kept, [["Key", "V"], ["keep", "remote"], ["gone", "remote"]])
    _write_xlsx(local_kept, [["Key", "V"], ["keep", "local"], ["gone", "local"]])
    _write_xlsx(remote_deleted, [["Key", "V"], ["keep", "remote"]])

    out_keep_remote = os.path.join(out_dir, "delete_keep_remote.xlsx")
    code = do_merge(
        local_deleted, base, remote_kept, out_keep_remote,
        mode="D",
        d_choices=[{"sheet": "Data", "key": "gone", "choice": "remote", "kind": "row"}],
        backup_root=os.path.join(out_dir, "backup_delete_keep_remote"),
    )
    _assert(code == 0, "remote keep delete-conflict merge should succeed")
    rows = _rows(out_keep_remote)
    _assert(["gone", "remote"] in rows, "remote-kept row should be copied by key")

    out_delete_local = os.path.join(out_dir, "delete_choose_local_deleted.xlsx")
    code = do_merge(
        local_deleted, base, remote_kept, out_delete_local,
        mode="D",
        d_choices=[{"sheet": "Data", "key": "gone", "choice": "local", "kind": "row"}],
        backup_root=os.path.join(out_dir, "backup_delete_choose_local"),
    )
    _assert(code == 0, "local delete choice should succeed")
    keys = [row[0] for row in _rows(out_delete_local)]
    _assert("gone" not in keys, "choosing deleted local side should remove output row")

    out_delete_remote = os.path.join(out_dir, "delete_choose_remote_deleted.xlsx")
    code = do_merge(
        local_kept, base, remote_deleted, out_delete_remote,
        mode="D",
        d_choices=[{"sheet": "Data", "key": "gone", "choice": "remote", "kind": "row"}],
        backup_root=os.path.join(out_dir, "backup_delete_choose_remote"),
    )
    _assert(code == 0, "remote delete choice should succeed")
    keys = [row[0] for row in _rows(out_delete_remote)]
    _assert("gone" not in keys, "choosing deleted remote side should remove output row")

    out_keep_local = os.path.join(out_dir, "delete_keep_local.xlsx")
    code = do_merge(
        local_kept, base, remote_deleted, out_keep_local,
        mode="D",
        d_choices=[{"sheet": "Data", "key": "gone", "choice": "local", "kind": "row"}],
        backup_root=os.path.join(out_dir, "backup_delete_keep_local"),
    )
    _assert(code == 0, "local keep delete-conflict merge should succeed")
    rows = _rows(out_keep_local)
    _assert(["gone", "local"] in rows, "local-kept row should be copied by key")


def test_multi_new_columns(out_dir):
    from merge_core import do_merge

    local = os.path.join(out_dir, "cols_local.xlsx")
    remote = os.path.join(out_dir, "cols_remote.xlsx")
    merged = os.path.join(out_dir, "cols_merged.xlsx")
    _write_xlsx(local, [["Key", "Name"], ["1", "Apple"], ["2", "Pear"]])
    _write_xlsx(
        remote,
        [
            ["Key", "Name", "Price", "Qty", "Total"],
            ["1", "Apple", "2", "3", "6"],
            ["2", "Pear", "4", "5", "20"],
        ],
    )
    code = do_merge(
        local, local, remote, merged,
        mode="B",
        backup_root=os.path.join(out_dir, "backup_cols"),
    )
    _assert(code == 0, "mode B merge should succeed")
    rows = _rows(merged)
    _assert(rows[0] == ["Key", "Name", "Price", "Qty", "Total"], "new columns should keep expected order")
    _assert(rows[1] == ["1", "Apple", "2", "3", "6"], "first data row should include all new col values")
    _assert(rows[2] == ["2", "Pear", "4", "5", "20"], "second data row should include all new col values")


def test_performance_smoke(out_dir):
    from merge_core import do_merge

    local = os.path.join(out_dir, "perf_local.xlsx")
    remote = os.path.join(out_dir, "perf_remote.xlsx")
    merged = os.path.join(out_dir, "perf_merged.xlsx")
    headers_local = ["Key"] + ["B%02d" % i for i in range(1, 16)]
    headers_remote = headers_local + ["N%02d" % i for i in range(1, 16)]
    rows_local = [headers_local]
    rows_remote = [headers_remote]
    for i in range(1, 301):
        rows_local.append(["k%03d" % i] + ["v%d_%02d" % (i, c) for c in range(1, 16)])
        rows_remote.append(
            ["k%03d" % i]
            + ["v%d_%02d" % (i, c) for c in range(1, 16)]
            + ["n%d_%02d" % (i, c) for c in range(1, 16)]
        )
    _write_xlsx(local, rows_local)
    _write_xlsx(remote, rows_remote)

    started = time.time()
    code = do_merge(
        local, local, remote, merged,
        mode="B",
        backup_root=os.path.join(out_dir, "backup_perf"),
    )
    elapsed = time.time() - started
    _assert(code == 0, "performance smoke merge should succeed")
    _assert(elapsed < 20, "performance smoke should finish quickly enough, elapsed=%.2fs" % elapsed)
    print("Performance smoke elapsed: %.2fs" % elapsed)


def main():
    _, out_dir = _setup_paths()
    tests = [
        test_compare_opens_once,
        test_backup_same_second,
        test_delete_conflict_choices,
        test_multi_new_columns,
        test_performance_smoke,
    ]
    for test in tests:
        test(out_dir)
        print("OK %s" % test.__name__)
    print("All regression tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
