# -*- coding: utf-8 -*-
"""
Git 相关操作：解决冲突后标记已解决、清理临时文件与备份。
只做一件事：给定 MERGED 与 LOCAL/BASE/REMOTE 路径，执行 git add、删除临时文件、
删除当前文件在 MergeExcelBackup 中的备份（并视情况从 git 索引移除）。
另提供 get_git_merge_info 用于在合并界面展示本地/线上提交信息。
"""

import os
import subprocess
import tempfile
from dataclasses import dataclass, field

from config import BACKUP_SUBDIR
from log_util import log


GIT_DISCOVERY_TIMEOUT = 15
GIT_ADD_TIMEOUT = 30
GIT_CLEANUP_TIMEOUT = 15


@dataclass
class CompletionResult:
    """合并确认动作的结构化结果。"""

    success: bool = False
    staged: bool = False
    cleaned: list = field(default_factory=list)
    skipped: list = field(default_factory=list)
    errors: list = field(default_factory=list)


class CleanupPolicy:
    """只允许清理显式登记目录下的临时文件。"""

    def __init__(self, allowed_roots=None):
        roots = []
        for root in allowed_roots or []:
            if not root:
                continue
            try:
                root_abs = os.path.abspath(root)
                if os.path.isdir(root_abs):
                    roots.append(os.path.normcase(root_abs))
            except Exception:
                pass
        seen = set()
        self.allowed_roots = []
        for root in roots:
            if root not in seen:
                seen.add(root)
                self.allowed_roots.append(root)

    @classmethod
    def default(cls):
        roots = [tempfile.gettempdir(), os.environ.get("TEMP"), os.environ.get("TMP")]
        return cls(roots)

    def allows(self, path):
        if not path or not os.path.isfile(path):
            return False
        try:
            abs_path = os.path.normcase(os.path.abspath(path))
            for root in self.allowed_roots:
                if os.path.commonpath([root, abs_path]) == root:
                    return True
        except (OSError, ValueError):
            return False
        return False


def _path_inside(path, root):
    try:
        abs_path = os.path.normcase(os.path.abspath(path))
        abs_root = os.path.normcase(os.path.abspath(root))
        return os.path.commonpath([abs_root, abs_path]) == abs_root
    except (OSError, ValueError):
        return False


def _find_git_marker_root(start_path):
    """Walk upward and return the nearest worktree root containing a .git marker."""
    try:
        current = os.path.abspath(start_path or os.getcwd())
        if os.path.isfile(current):
            current = os.path.dirname(current)
    except Exception:
        current = os.getcwd()

    while current:
        marker = os.path.join(current, ".git")
        if os.path.isdir(marker) or os.path.isfile(marker):
            return current
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent
    return None


def discover_git_worktree_root(start_path, timeout=GIT_DISCOVERY_TIMEOUT):
    """
    Return (repo_root, error_message).

    Prefer the local .git marker so confirmation is not blocked by slow
    `git rev-parse` calls in Windows GUI/Fork environments.
    """
    marker_root = _find_git_marker_root(start_path)
    if marker_root:
        return os.path.abspath(marker_root), None

    try:
        cwd = os.path.abspath(start_path or os.getcwd())
        if os.path.isfile(cwd):
            cwd = os.path.dirname(cwd)
        if not os.path.isdir(cwd):
            cwd = os.getcwd()
        rr = subprocess.run(
            ["git", "rev-parse", "--show-toplevel"],
            cwd=cwd,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        if rr.returncode == 0 and rr.stdout.strip():
            return os.path.abspath(rr.stdout.strip()), None
        return None, "git rev-parse 失败，无法确认仓库根目录: %s" % (rr.stderr or rr.stdout or "未知")
    except Exception as e:
        return None, "无法执行 git rev-parse: %s" % e


def stage_merged_and_cleanup(path_merged, path_local, path_base, path_remote, log_callback=None, cleanup_policy=None):
    """
    解决冲突后：
    1) 对合并文件执行 git add，使 Fork 识别为已解决；
    2) 只删除显式允许临时根目录内的 LOCAL/BASE/REMOTE 文件；
    3) 删除 MERGED 所在目录下 MergeExcelBackup 中当前文件的 _local/_remote/_merged 备份，
       并从 git 索引移除（若在仓库内）。
    log_callback(msg, is_error=False) 可选，用于 GUI 状态栏等。
    """
    def _log_cb(msg, is_err=False):
        log(msg, is_error=is_err)
        if log_callback:
            try:
                log_callback(msg, is_err)
            except Exception:
                pass

    result = CompletionResult()
    work_dir = os.path.dirname(os.path.abspath(path_merged))
    if not work_dir:
        work_dir = "."
    abs_merged = os.path.abspath(path_merged)
    repo_root, msg = discover_git_worktree_root(work_dir)
    if not repo_root:
        result.errors.append(msg)
        _log_cb(msg, True)
        return result

    if not _path_inside(abs_merged, repo_root):
        msg = "MERGED 不在当前 Git 仓库中，已停止确认流程: %s" % abs_merged
        result.errors.append(msg)
        _log_cb(msg, True)
        return result

    rel_path = os.path.relpath(abs_merged, repo_root).replace("\\", "/")

    try:
        r = subprocess.run(
            ["git", "add", "--", rel_path],
            cwd=repo_root,
            capture_output=True,
            text=True,
            timeout=GIT_ADD_TIMEOUT,
        )
        if r.returncode == 0:
            result.staged = True
            _log_cb("已执行 git add，冲突已标记为已解决")
        else:
            msg = "git add 失败: %s" % (r.stderr or r.stdout or "未知")
            result.errors.append(msg)
            _log_cb(msg, is_err=True)
            return result
    except Exception as e:
        msg = "git add 异常: %s" % e
        result.errors.append(msg)
        _log_cb(msg, is_err=True)
        return result

    cleanup_policy = cleanup_policy or CleanupPolicy.default()
    for label, p in [("LOCAL", path_local), ("BASE", path_base), ("REMOTE", path_remote)]:
        if not p or not os.path.isfile(p):
            continue
        if cleanup_policy.allows(p):
            try:
                os.remove(p)
                result.cleaned.append(p)
                _log_cb("已清理临时文件 %s: %s" % (label, p))
            except Exception as e:
                msg = "清理 %s 失败: %s" % (label, e)
                result.errors.append(msg)
                _log_cb(msg, is_err=True)
        else:
            result.skipped.append(p)
            _log_cb("跳过清理 %s：路径不在允许的临时目录内 %s" % (label, p))

    merged_dir = os.path.dirname(os.path.abspath(path_merged))
    base_name = os.path.splitext(os.path.basename(path_merged))[0]
    backup_dir = os.path.join(merged_dir, BACKUP_SUBDIR)
    for suf in ["_local.xlsx", "_remote.xlsx", "_merged.xlsx"]:
        bp = os.path.join(backup_dir, base_name + suf)
        if not os.path.isfile(bp):
            continue
        bp_rel = os.path.relpath(bp, repo_root).replace("\\", "/")
        try:
            if _path_inside(bp, repo_root):
                r = subprocess.run(["git", "rm", "-f", "--", bp_rel], cwd=repo_root, capture_output=True, text=True, timeout=GIT_CLEANUP_TIMEOUT)
                if r.returncode == 0:
                    result.cleaned.append(bp)
                    _log_cb("已删除备份(含从 git 移除): %s" % bp)
                else:
                    os.remove(bp)
                    result.cleaned.append(bp)
                    _log_cb("已删除备份: %s" % bp)
            else:
                os.remove(bp)
                result.cleaned.append(bp)
                _log_cb("已删除备份: %s" % bp)
        except Exception as e:
            try:
                os.remove(bp)
                result.cleaned.append(bp)
                _log_cb("已删除备份: %s" % bp)
            except Exception as e2:
                msg = "删除备份失败 %s: %s" % (bp, e2)
                result.errors.append(msg)
                _log_cb(msg, is_err=True)

    if os.path.isdir(backup_dir) and not os.listdir(backup_dir):
        try:
            os.rmdir(backup_dir)
            result.cleaned.append(backup_dir)
            _log_cb("已删除空备份目录: %s" % backup_dir)
        except Exception:
            pass

    result.success = result.staged and not result.errors
    return result


def get_git_merge_info(path_merged):
    """
    从 MERGED 文件所在仓库获取 LOCAL（HEAD）与 REMOTE（MERGE_HEAD 等）的最近一次提交信息，
    用于合并界面展示“本地/线上”是谁在何时改的。
    返回 (local_info, remote_info)，每个为 dict: hash, short_hash, author, email, date, message；
    获取失败时返回 (None, None) 或 remote_info 为 {}。
    """
    local_info = None
    remote_info = None
    try:
        work_dir = os.path.dirname(os.path.abspath(path_merged))
        if not work_dir:
            work_dir = "."
        rel_path = os.path.relpath(path_merged, work_dir).replace("\\", "/")
        if rel_path.startswith(".."):
            return None, None

        fmt = "%H%n%h%n%an%n%ae%n%ci%n%s"

        def run_git(ref):
            r = subprocess.run(
                ["git", "log", "-1", "--format=" + fmt, ref, "--", rel_path],
                cwd=work_dir,
                capture_output=True,
                text=True,
                timeout=5,
            )
            if r.returncode != 0 or not r.stdout.strip():
                return None
            parts = r.stdout.strip().split("\n")
            if len(parts) >= 6:
                return {
                    "hash": parts[0],
                    "short_hash": parts[1],
                    "author": parts[2],
                    "email": parts[3],
                    "date": parts[4],
                    "message": parts[5],
                }
            return None

        local_info = run_git("HEAD")
        for ref in ["MERGE_HEAD", "REBASE_HEAD", "CHERRY_PICK_HEAD"]:
            remote_info = run_git(ref)
            if remote_info:
                break
        if remote_info is None:
            remote_info = {}
    except Exception:
        pass
    return local_info, remote_info
