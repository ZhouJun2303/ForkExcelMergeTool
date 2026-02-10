# -*- coding: utf-8 -*-
"""
Fork Excel Merge Tool 入口。
只做一件事：解析命令行参数（支持 Fork 传入的逗号拼接路径），根据参数个数启动合并 GUI 或对比 GUI，
失败时回退到命令行合并/对比；不包含具体合并、对比、冲突检测逻辑。

用法:
  合并（4 参数）: python MergeExcelFork.py <local> <base> <remote> <merged>
  对比（2 参数）: python MergeExcelFork.py <local> <remote>

退出码: 0 成功, 1 参数/文件错误, 2 合并/对比异常
日志: 与 exe/脚本同目录下的 MergeExcelFork.log
"""

import os
import sys
import traceback

from log_util import log


def _normalize_args():
    """
    Fork 可能将多个路径合并为单个参数 "path1,path2" 或 "path1,path2,path3,path4"，
    按逗号拆分为列表。返回 (mode, args)，mode 为 'merge'|'compare'|None。
    """
    raw = sys.argv[1:]
    log("启动 原始 args=%s" % raw)
    args = []
    for a in raw:
        for p in a.split(","):
            p = p.strip().strip('"').strip("'")
            if p:
                args.append(p)
    argc = len(args)
    if argc == 4:
        return "merge", args
    if argc == 2:
        return "compare", args
    return None, args


def main():
    try:
        mode, args = _normalize_args()
        argc = len(args)
        log("解析后 argc=%d args=%s" % (argc, args))

        if mode == "merge":
            path_local, path_base, path_remote, path_merged = args[0], args[1], args[2], args[3]
            for p, name in [(path_local, "LOCAL"), (path_base, "BASE"), (path_remote, "REMOTE")]:
                if not os.path.isfile(p):
                    msg = "%s 不存在: %s" % (name, p)
                    log(msg, is_error=True)
                    print("ERROR: " + msg, file=sys.stderr)
                    sys.exit(1)
            try:
                from merge_gui import MergeWindow
                win = MergeWindow(path_local, path_base, path_remote, path_merged)
                win.run()
                sys.exit(0)
            except Exception as gui_err:
                log("GUI 合并失败，回退命令行: %s" % gui_err)
                from merge_core import do_merge
                code = do_merge(path_local, path_base, path_remote, path_merged)
                sys.exit(code)

        elif mode == "compare":
            path_a, path_b = args[0], args[1]
            if not os.path.isfile(path_a):
                msg = "文件 A 不存在: %s" % path_a
                log(msg, is_error=True)
                print("ERROR: " + msg, file=sys.stderr)
                sys.exit(1)
            if not os.path.isfile(path_b):
                msg = "文件 B 不存在: %s" % path_b
                log(msg, is_error=True)
                print("ERROR: " + msg, file=sys.stderr)
                sys.exit(1)
            try:
                from diff_gui import DiffWindow
                win = DiffWindow(path_a, path_b)
                win.run()
                sys.exit(0)
            except Exception as gui_err:
                log("GUI 对比失败，回退命令行: %s" % gui_err)
                from compare_core import do_compare
                code = do_compare(path_a, path_b)
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
