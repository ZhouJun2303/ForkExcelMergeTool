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


def _lock_path(name):
    return os.path.join(log_dir(), ".merge_excel_%s.lock" % name)


def _is_pid_alive(pid):
    try:
        os.kill(pid, 0)
        return True
    except (OSError, ValueError):
        return False


def try_acquire_compare_lock():
    """尝试占用对比窗口锁。返回 True 表示成功，False 表示已有其他进程占用。"""
    path = _lock_path("compare")
    try:
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                pid = int(f.read().strip())
            if _is_pid_alive(pid):
                return False
        with open(path, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
        return True
    except Exception:
        try:
            if os.path.isfile(path):
                os.remove(path)
        except Exception:
            pass
        with open(path, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
        return True


def release_compare_lock():
    try:
        path = _lock_path("compare")
        if os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass


def try_acquire_merge_lock():
    """尝试占用合并窗口锁。返回 True 表示成功，False 表示已有其他进程占用。"""
    path = _lock_path("merge")
    try:
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                pid = int(f.read().strip())
            if _is_pid_alive(pid):
                return False
        with open(path, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
        return True
    except Exception:
        try:
            if os.path.isfile(path):
                os.remove(path)
        except Exception:
            pass
        with open(path, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
        return True


def release_merge_lock():
    try:
        path = _lock_path("merge")
        if os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass


def log(msg, is_error=False):
    """追加一行到 MergeExcelFork.log。"""
    try:
        with open(log_path(), "a", encoding="utf-8") as f:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            prefix = "[ERROR] " if is_error else ""
            f.write("%s %s%s\n" % (ts, prefix, msg))
    except Exception:
        pass
