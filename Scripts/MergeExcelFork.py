# -*- coding: utf-8 -*-
"""
Fork Excel Merge Tool 入口。
只做一件事：解析命令行参数（支持 Fork 传入的逗号拼接路径），根据默认运行模式启动设置中心、快速备份、
合并 GUI 或对比 GUI；失败时回退到命令行合并/对比。

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
from app_settings import STARTUP_FEATURE_BACKUP_ONLY, load_startup_feature
from excel_format import merge_diff_extension_text, merge_diff_supported
from log_util import (
    log,
    release_compare_lock,
    release_merge_lock,
    try_acquire_compare_lock,
    try_acquire_merge_lock,
)


def _unsupported_merge_diff_files(paths):
    return [p for p in paths if p and not merge_diff_supported(p)]


def _show_unsupported_format(mode, paths):
    from quick_backup_gui import show_quick_backup_panel

    msg = (
        "当前默认运行模式是合并对比模式，但这些 Excel 后缀暂不支持解析：\n%s\n\n"
        "合并对比模式当前支持：%s。\n"
        "可以在设置中心切换到快速备份模式，或先把文件转换为支持的格式。"
    ) % ("\n".join(paths), merge_diff_extension_text())
    show_quick_backup_panel(mode, error=msg)


def _normalize_args():
    """
    Fork 可能将多个路径合并为单个参数 "path1,path2" 或 "path1,path2,path3,path4"，
    按逗号拆分为列表。返回 (mode, args)，mode 为 'merge'|'compare'|None。
    """
    raw = sys.argv[1:]
    log("启动 原始 args=%s" % raw)
    if not raw:
        return "main", []
    if raw and raw[0] == "--main":
        return "main", raw[1:]
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

        if mode == "main":
            from main_gui import MainWindow
            win = MainWindow()
            win.run()
            sys.exit(0)

        if mode == "git-driver":
            if argc != 4:
                msg = "Usage: --git-merge-driver <base> <current> <other> <repo-path>"
                log(msg, is_error=True)
                print(msg, file=sys.stderr)
                sys.exit(1)
            if load_startup_feature() == STARTUP_FEATURE_BACKUP_ONLY:
                from quick_backup import quick_backup_git_driver
                from quick_backup_gui import show_quick_backup_panel
                try:
                    try:
                        from git_merge_driver import _resolve_context_path
                        context_path = _resolve_context_path(args[1], args[3])
                    except Exception:
                        context_path = args[3] or args[1]
                    info = quick_backup_git_driver(args[0], args[1], args[2], context_path=context_path)
                    log("[QuickBackup] git-driver 已备份，保持未解决状态: %s" % info.get("dir"))
                    print("OK: 已快速备份，Git 冲突保持未解决。备份=%s" % info.get("dir"), file=sys.stdout)
                    show_quick_backup_panel("git-driver", backup_info=info)
                    sys.exit(1)
                except Exception as e:
                    log("[QuickBackup] git-driver 备份失败: %s" % e, is_error=True)
                    print("ERROR: 快速备份失败: %s" % e, file=sys.stderr)
                    show_quick_backup_panel("git-driver", error=e)
                    sys.exit(2)
            unsupported = _unsupported_merge_diff_files([args[0], args[1], args[2]])
            if unsupported:
                _show_unsupported_format("git-driver", unsupported)
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
            if load_startup_feature() == STARTUP_FEATURE_BACKUP_ONLY:
                from quick_backup import quick_backup_merge
                from quick_backup_gui import show_quick_backup_panel
                try:
                    info = quick_backup_merge(path_local, path_base, path_remote, path_merged)
                    print("OK: 已快速备份。备份=%s" % info.get("dir"), file=sys.stdout)
                    show_quick_backup_panel("merge", backup_info=info)
                    sys.exit(0)
                except Exception as e:
                    log("快速备份失败: %s" % e, is_error=True)
                    print("ERROR: 快速备份失败: %s" % e, file=sys.stderr)
                    show_quick_backup_panel("merge", error=e)
                    sys.exit(2)
            unsupported = _unsupported_merge_diff_files([path_local, path_base, path_remote, path_merged])
            if unsupported:
                _show_unsupported_format("merge", unsupported)
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
            if load_startup_feature() == STARTUP_FEATURE_BACKUP_ONLY:
                from quick_backup import quick_backup_compare
                from quick_backup_gui import show_quick_backup_panel
                try:
                    info = quick_backup_compare(path_remote, path_local)
                    print("OK: 已快速备份。备份=%s" % info.get("dir"), file=sys.stdout)
                    show_quick_backup_panel("compare", backup_info=info)
                    sys.exit(0)
                except Exception as e:
                    log("快速备份失败: %s" % e, is_error=True)
                    print("ERROR: 快速备份失败: %s" % e, file=sys.stderr)
                    show_quick_backup_panel("compare", error=e)
                    sys.exit(2)
            unsupported = _unsupported_merge_diff_files([path_remote, path_local])
            if unsupported:
                _show_unsupported_format("compare", unsupported)
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
            msg = "Usage: Settings Center (no args/--main) | Merge (4 args) | Compare (2 args). Fork 可能传 path1,path2 单参数，已支持拆分"
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
