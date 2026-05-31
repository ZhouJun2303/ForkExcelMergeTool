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


def _try_acquire_lock(name):
    """原子方式创建锁文件；遇到僵尸锁会清理后重试。"""
    path = _lock_path(name)
    for _ in range(3):
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write(str(os.getpid()))
            return True
        except FileExistsError:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    pid = int(f.read().strip())
                if _is_pid_alive(pid):
                    return False
            except Exception:
                pass
            try:
                os.remove(path)
            except FileNotFoundError:
                pass
            except Exception:
                return False
        except Exception:
            return False
    return False


def _release_lock(name):
    path = _lock_path(name)
    try:
        if not os.path.isfile(path):
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                pid = int(f.read().strip())
            if pid != os.getpid():
                return
        except Exception:
            pass
        os.remove(path)
    except Exception:
        pass


def try_acquire_compare_lock():
    """尝试占用对比窗口锁。返回 True 表示成功，False 表示已有其他进程占用。"""
    return _try_acquire_lock("compare")


def release_compare_lock():
    _release_lock("compare")


def try_acquire_merge_lock():
    """尝试占用合并窗口锁。返回 True 表示成功，False 表示已有其他进程占用。"""
    return _try_acquire_lock("merge")


def release_merge_lock():
    _release_lock("merge")


def log(msg, is_error=False):
    """追加一行到 MergeExcelFork.log。"""
    try:
        with open(log_path(), "a", encoding="utf-8") as f:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            prefix = "[ERROR] " if is_error else ""
            f.write("%s %s%s\n" % (ts, prefix, msg))
    except Exception:
        pass
