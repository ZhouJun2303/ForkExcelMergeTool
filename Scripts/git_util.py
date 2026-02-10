# -*- coding: utf-8 -*-
"""
Git 相关操作：解决冲突后标记已解决、清理临时文件与备份。
只做一件事：给定 MERGED 与 LOCAL/BASE/REMOTE 路径，执行 git add、删除临时文件、
删除当前文件在 MergeExcelBackup 中的备份（并视情况从 git 索引移除）。
另提供 get_git_merge_info 用于在合并界面展示本地/线上提交信息。
"""

import os
import subprocess

from config import BACKUP_SUBDIR
from log_util import log


def stage_merged_and_cleanup(path_merged, path_local, path_base, path_remote, log_callback=None):
    """
    解决冲突后：
    1) 对合并文件执行 git add，使 Fork 识别为已解决；
    2) 删除 LOCAL/BASE/REMOTE 中的临时文件（路径含 temp/tmp/fork/appdata 等）；
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

    work_dir = os.path.dirname(os.path.abspath(path_merged))
    if not work_dir:
        work_dir = "."
    abs_merged = os.path.abspath(path_merged)
    try:
        rr = subprocess.run(
            ["git", "rev-parse", "--show-toplevel"],
            cwd=work_dir,
            capture_output=True,
            text=True,
            timeout=5,
        )
        repo_root = rr.stdout.strip() if rr.returncode == 0 and rr.stdout else work_dir
    except Exception:
        repo_root = work_dir

    rel_path = os.path.relpath(abs_merged, repo_root).replace("\\", "/")
    if rel_path.startswith(".."):
        rel_path = os.path.basename(path_merged)

    try:
        r = subprocess.run(
            ["git", "add", rel_path],
            cwd=repo_root,
            capture_output=True,
            text=True,
            timeout=10,
        )
        if r.returncode == 0:
            _log_cb("已执行 git add，冲突已标记为已解决")
        else:
            _log_cb("git add 失败: %s" % (r.stderr or r.stdout or "未知"), is_err=True)
    except Exception as e:
        _log_cb("git add 异常: %s" % e, is_err=True)

    def _is_temp(p):
        if not p or not os.path.isfile(p):
            return False
        pn = p.lower()
        return "temp" in pn or "tmp" in pn or "fork" in pn or "appdata" in pn

    for label, p in [("LOCAL", path_local), ("BASE", path_base), ("REMOTE", path_remote)]:
        if _is_temp(p):
            try:
                os.remove(p)
                _log_cb("已清理临时文件 %s: %s" % (label, p))
            except Exception as e:
                _log_cb("清理 %s 失败: %s" % (label, e), is_err=True)

    merged_dir = os.path.dirname(os.path.abspath(path_merged))
    base_name = os.path.splitext(os.path.basename(path_merged))[0]
    backup_dir = os.path.join(merged_dir, BACKUP_SUBDIR)
    for suf in ["_local.xlsx", "_remote.xlsx", "_merged.xlsx"]:
        bp = os.path.join(backup_dir, base_name + suf)
        if not os.path.isfile(bp):
            continue
        bp_rel = os.path.relpath(bp, repo_root).replace("\\", "/")
        try:
            r = subprocess.run(["git", "rm", "-f", bp_rel], cwd=repo_root, capture_output=True, text=True, timeout=5)
            if r.returncode == 0:
                _log_cb("已删除备份(含从 git 移除): %s" % bp)
            else:
                os.remove(bp)
                _log_cb("已删除备份: %s" % bp)
        except Exception as e:
            try:
                os.remove(bp)
                _log_cb("已删除备份: %s" % bp)
            except Exception as e2:
                _log_cb("删除备份失败 %s: %s" % (bp, e2), is_err=True)

    if os.path.isdir(backup_dir) and not os.listdir(backup_dir):
        try:
            os.rmdir(backup_dir)
            _log_cb("已删除空备份目录: %s" % backup_dir)
        except Exception:
            pass


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
