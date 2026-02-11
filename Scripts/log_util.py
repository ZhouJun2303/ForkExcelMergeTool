# -*- coding: utf-8 -*-
"""
日志工具：只做“写日志到文件”和“解析日志所在目录/路径”。
不包含 GUI 状态栏更新（由 gui_common 负责）。
"""

import os
import sys
from datetime import datetime

from config import LOG_FILE, MERGE_OPTIONS_FILE


def log_dir():
    """日志文件所在目录：打包为 exe 时为 exe 所在目录，否则为当前脚本所在目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def log_path():
    """日志文件的完整路径。"""
    return os.path.join(log_dir(), LOG_FILE)


def merge_options_path():
    """合并选项持久化文件的完整路径（与 exe/脚本同目录）。"""
    return os.path.join(log_dir(), MERGE_OPTIONS_FILE)


def log(msg, is_error=False):
    """追加一行到 MergeExcelFork.log。"""
    try:
        with open(log_path(), "a", encoding="utf-8") as f:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            prefix = "[ERROR] " if is_error else ""
            f.write("%s %s%s\n" % (ts, prefix, msg))
    except Exception:
        pass
