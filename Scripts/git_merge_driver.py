# -*- coding: utf-8 -*-
"""
Git 底层 merge driver 适配层。
只负责把 Git 的 %O/%A/%B/%P 契约隔离成临时工作目录，复用现有 GUI 合并，
用户确认后再把结果原子写回 %A。driver 模式不执行 git add，也不删除 Git 管理文件。
"""

import os
import shutil
import sys
import tempfile
import traceback
from dataclasses import dataclass, field

from git_util import discover_git_worktree_root
from log_util import log, release_merge_lock, try_acquire_merge_lock


def _same_path(left, right):
    try:
        return os.path.normcase(os.path.abspath(left)) == os.path.normcase(os.path.abspath(right))
    except Exception:
        return False


@dataclass
class GitDriverCompletionResult:
    success: bool = False
    message: str = ""
    errors: list = field(default_factory=list)


class GitDriverCompletionStrategy:
    def __init__(self, target_current, temp_merged):
        self.target_current = os.path.abspath(target_current)
        self.temp_merged = os.path.abspath(temp_merged)
        self.completed = False

    def complete(self, _window):
        result = GitDriverCompletionResult()
        try:
            if not os.path.isfile(self.temp_merged):
                result.errors.append("合并结果不存在: %s" % self.temp_merged)
                return result
            target_dir = os.path.dirname(self.target_current) or "."
            os.makedirs(target_dir, exist_ok=True)
            fd, tmp_path = tempfile.mkstemp(
                suffix=".xlsx",
                prefix=".excelmergefork_driver_",
                dir=target_dir,
            )
            os.close(fd)
            try:
                shutil.copy2(self.temp_merged, tmp_path)
                os.replace(tmp_path, self.target_current)
            finally:
                try:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)
                except Exception:
                    pass
            self.completed = True
            result.success = True
            result.message = "Git merge driver 已写回当前文件，Git 将继续合并。"
            return result
        except Exception as e:
            result.errors.append("写回 Git 当前文件失败: %s" % e)
            return result


def _copy_input(src, dst, label):
    if not os.path.isfile(src):
        raise FileNotFoundError("%s 不存在: %s" % (label, src))
    shutil.copy2(src, dst)


def _resolve_context_path(path_current, repo_path):
    if not repo_path:
        return path_current
    try:
        root = _git_root(os.path.dirname(os.path.abspath(path_current)) or ".")
        candidate = os.path.abspath(os.path.join(root, repo_path))
        root_abs = os.path.abspath(root)
        if os.path.commonpath([root_abs, candidate]) == root_abs:
            return candidate
    except Exception:
        pass
    return path_current


def _git_root(cwd):
    root, _ = discover_git_worktree_root(cwd)
    if root:
        return root
    return cwd


def _excel_temp_ext(path_current, repo_path):
    for path in (repo_path, path_current):
        ext = os.path.splitext(path or "")[1]
        if ext:
            return ext
    return ".xlsx"


def run_git_merge_driver(path_base, path_current, path_other, repo_path):
    """
    Git driver 入口。
    返回 0 表示用户确认且已写回 %A；1 表示取消/无法完成；2 表示内部异常。
    """
    temp_dir = None
    strategy = None
    lock_acquired = False
    try:
        path_base = os.path.abspath(path_base)
        path_current = os.path.abspath(path_current)
        path_other = os.path.abspath(path_other)
        log("[GitDriver] start repo_path=%s base=%s current=%s other=%s" % (
            repo_path, path_base, path_current, path_other,
        ))

        temp_dir = tempfile.mkdtemp(prefix="ExcelMergeForkDriver_")
        ext = _excel_temp_ext(path_current, repo_path)
        temp_base = os.path.join(temp_dir, "base" + ext)
        temp_local = os.path.join(temp_dir, "local" + ext)
        temp_remote = os.path.join(temp_dir, "remote" + ext)
        temp_merged = os.path.join(temp_dir, "merged" + ext)

        _copy_input(path_base, temp_base, "BASE")
        _copy_input(path_current, temp_local, "CURRENT")
        if _same_path(path_other, path_current):
            shutil.copy2(temp_local, temp_remote)
        else:
            _copy_input(path_other, temp_remote, "OTHER")
        shutil.copy2(temp_local, temp_merged)

        if not try_acquire_merge_lock():
            log("[GitDriver] merge window lock is already held", is_error=True)
            return 1
        lock_acquired = True

        from merge_gui import MergeWindow

        context_path = _resolve_context_path(path_current, repo_path)
        strategy = GitDriverCompletionStrategy(path_current, temp_merged)
        strategy.target_current = path_current
        strategy.context_path = context_path
        win = MergeWindow(temp_local, temp_base, temp_remote, temp_merged, completion_strategy=strategy)
        win.root.title("Excel Git 合并驱动 - %s" % (repo_path or os.path.basename(path_current)))
        win.run()
        return 0 if strategy.completed else 1
    except SystemExit as e:
        code = e.code
        if strategy is not None and strategy.completed:
            return 0
        return code if isinstance(code, int) and code in (0, 1, 2) else 1
    except Exception as e:
        log("[GitDriver] exception: %s\n%s" % (e, traceback.format_exc()), is_error=True)
        print("ERROR: Git merge driver failed: %s" % e, file=sys.stderr)
        return 2
    finally:
        if lock_acquired:
            release_merge_lock()
        if temp_dir:
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
                log("[GitDriver] cleaned temp dir: %s" % temp_dir)
            except Exception as e:
                log("[GitDriver] cleanup temp dir failed: %s" % e, is_error=True)
