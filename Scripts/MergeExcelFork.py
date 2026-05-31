# -*- coding: utf-8 -*-
"""
Fork Excel Merge Tool 入口。
只做一件事：解析命令行参数（支持 Fork 传入的逗号拼接路径），根据参数个数启动合并 GUI 或对比 GUI，
失败时回退到命令行合并/对比；不包含具体合并、对比、冲突检测逻辑。

用法:
  合并（4 参数）: python MergeExcelFork.py <local> <base> <remote> <merged>
  对比（2 参数，Fork 推荐）: python MergeExcelFork.py <remote> <local>

退出码: 0 成功, 1 参数/文件错误, 2 合并/对比异常
日志: 与 exe/脚本同目录下的 MergeExcelFork.log
"""

import os
import csv
import sys
import traceback

from tkinter import messagebox
from log_util import (
    log,
    release_compare_lock,
    release_merge_lock,
    try_acquire_compare_lock,
    try_acquire_merge_lock,
)


def _normalize_args():
    """
    Fork 可能将多个路径合并为单个参数 "path1,path2" 或 "path1,path2,path3,path4"，
    按逗号拆分为列表。返回 (mode, args)，mode 为 'merge'|'compare'|None。
    """
    raw = sys.argv[1:]
    log("启动 原始 args=%s" % raw)
    if raw and raw[0] == "--git-merge-driver":
        return "git-driver", raw[1:]
    raw = list(raw)
    flags = []
    if "--include-same" in raw:
        flags.append("--include-same")
        raw = [a for a in raw if a != "--include-same"]
    if len(raw) in (2, 4):
        args = [a.strip().strip('"').strip("'") for a in raw if a.strip()]
    elif len(raw) == 1:
        try:
            args = [p.strip().strip('"').strip("'") for p in next(csv.reader([raw[0]], skipinitialspace=True)) if p.strip()]
        except Exception:
            args = [p.strip().strip('"').strip("'") for p in raw[0].split(",") if p.strip()]
    else:
        args = []
        for a in raw:
            if "," in a:
                try:
                    parts = next(csv.reader([a], skipinitialspace=True))
                except Exception:
                    parts = a.split(",")
            else:
                parts = [a]
            for p in parts:
                p = p.strip().strip('"').strip("'")
                if p:
                    args.append(p)
    argc = len(args)
    if argc == 4:
        return "merge", flags + args
    if argc == 2:
        return "compare", flags + args
    return None, flags + args


def main():
    try:
        mode, args = _normalize_args()
        argc = len(args)
        log("解析后 argc=%d args=%s" % (argc, args))

        include_same = False
        if "--include-same" in args:
            include_same = True
            args = [a for a in args if a != "--include-same"]
            argc = len(args)
            if len(args) == 4:
                mode = "merge"
            elif len(args) == 2:
                mode = "compare"

        if mode == "git-driver":
            if argc != 4:
                msg = "Usage: --git-merge-driver <base> <current> <other> <repo-path>"
                log(msg, is_error=True)
                print(msg, file=sys.stderr)
                sys.exit(1)
            from git_merge_driver import run_git_merge_driver
            sys.exit(run_git_merge_driver(args[0], args[1], args[2], args[3]))

        if mode == "merge":
            path_local, path_base, path_remote, path_merged = args[0], args[1], args[2], args[3]
            for p, name in [(path_local, "LOCAL"), (path_base, "BASE"), (path_remote, "REMOTE")]:
                if not os.path.isfile(p):
                    msg = "%s 不存在: %s" % (name, p)
                    log(msg, is_error=True)
                    print("ERROR: " + msg, file=sys.stderr)
                    sys.exit(1)
            try:
                from merge_gui import MergeWindow, get_existing_merge_window
                existing = get_existing_merge_window()
                if existing is not None:
                    existing.activate_and_refresh(path_local, path_base, path_remote, path_merged)
                    messagebox.showinfo("提示", "合并窗口已存在，已激活并刷新。")
                    sys.exit(0)
                lock_acquired = False
                if not try_acquire_merge_lock():
                    messagebox.showwarning("提示", "合并窗口已在其他进程中打开，请先关闭后再试。")
                    sys.exit(0)
                lock_acquired = True
                try:
                    win = MergeWindow(path_local, path_base, path_remote, path_merged)
                    win.run()
                finally:
                    if lock_acquired:
                        release_merge_lock()
                sys.exit(0)
            except Exception as gui_err:
                log("GUI 合并失败，回退命令行: %s" % gui_err)
                from merge_core import do_merge
                code = do_merge(path_local, path_base, path_remote, path_merged)
                sys.exit(code)

        elif mode == "compare":
            path_remote, path_local = args[0], args[1]
            if not os.path.isfile(path_remote):
                msg = "REMOTE 不存在: %s" % path_remote
                log(msg, is_error=True)
                print("ERROR: " + msg, file=sys.stderr)
                sys.exit(1)
            if not os.path.isfile(path_local):
                msg = "LOCAL 不存在: %s" % path_local
                log(msg, is_error=True)
                print("ERROR: " + msg, file=sys.stderr)
                sys.exit(1)
            try:
                from diff_gui import DiffWindow, get_existing_diff_window
                existing = get_existing_diff_window()
                if existing is not None:
                    existing.activate_and_refresh(path_local, path_remote)
                    messagebox.showinfo("提示", "对比窗口已存在，已激活并刷新。")
                    sys.exit(0)
                lock_acquired = False
                if not try_acquire_compare_lock():
                    messagebox.showwarning("提示", "对比窗口已在其他进程中打开，请先关闭后再试。")
                    sys.exit(0)
                lock_acquired = True
                try:
                    win = DiffWindow(path_local, path_remote)
                    win.run()
                finally:
                    if lock_acquired:
                        release_compare_lock()
                sys.exit(0)
            except Exception as gui_err:
                log("GUI 对比失败，回退命令行: %s" % gui_err)
                from compare_core import do_compare
                code = do_compare(path_local, path_remote, include_same=include_same)
                sys.exit(code if isinstance(code, int) else 0)

        else:
            msg = "Usage: Merge (4 args) | Compare (2 args). Fork 可能传 path1,path2 单参数，已支持拆分"
            log(msg)
            print(msg, file=sys.stderr)
            sys.exit(1)

    except Exception as e:
        tb = traceback.format_exc()
        log("异常: %s\n%s" % (e, tb), is_error=True)
        print("ERROR: %s\n%s" % (e, tb), file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
