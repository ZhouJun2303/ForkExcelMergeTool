# -*- coding: utf-8 -*-
"""
合并备份工具：解析备份根目录、按 项目/时间 创建快照目录，并复制本地/线上/合并结果。
"""

import json
import hashlib
import os
import re
import shutil
import subprocess
from datetime import datetime

from config import BACKUP_SUBDIR
from git_util import discover_git_worktree_root
from log_util import merge_options_path


BACKUP_ROOT_OPTION = "backup_root_dir"
INVALID_DIR_CHARS = r'[<>:"/\\|?*\x00-\x1f]'
MAX_NAME_PART_LEN = 48
MAX_BACKUP_DIR_LABEL_LEN = 96
MAX_BACKUP_FILE_STEM_LEN = 96
MAX_BACKUP_PATH_LEN = 240


def _load_options_data():
    try:
        path = merge_options_path()
        if path and os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return {}


def _save_options_data(data):
    path = merge_options_path()
    if not path:
        return
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_saved_backup_root():
    """读取用户选择的备份根目录；未设置时返回空字符串。"""
    data = _load_options_data()
    value = data.get(BACKUP_ROOT_OPTION, "")
    return value.strip() if isinstance(value, str) else ""


def save_backup_root(root_dir):
    """保存用户选择的备份根目录；空值表示恢复默认目录。"""
    data = _load_options_data()
    root_dir = (root_dir or "").strip()
    if root_dir:
        data[BACKUP_ROOT_OPTION] = os.path.normpath(root_dir)
    elif BACKUP_ROOT_OPTION in data:
        del data[BACKUP_ROOT_OPTION]
    _save_options_data(data)


def _sanitize_dir_name(name, fallback, max_len=None):
    cleaned = re.sub(INVALID_DIR_CHARS, "_", name or "")
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" ._")
    if max_len and len(cleaned) > max_len:
        cleaned = cleaned[:max_len].rstrip(" ._")
    return cleaned or fallback


def _safe_name_part(value, fallback, max_len=MAX_NAME_PART_LEN):
    return _sanitize_dir_name(value, fallback, max_len=max_len)


def project_name_for_backup(path_merged):
    """优先使用 Git 仓库目录名；不在仓库中时使用 MERGED 所在目录名。"""
    merged_dir = os.path.dirname(os.path.abspath(path_merged)) or os.getcwd()
    project_dir = merged_dir
    repo_root, _ = discover_git_worktree_root(merged_dir)
    if repo_root:
        project_dir = repo_root
    return _sanitize_dir_name(os.path.basename(os.path.normpath(project_dir)), "Project")


def default_backup_root(path_merged):
    """默认备份根目录仍放在 MERGED 同目录下，便于兼容旧使用方式。"""
    merged_dir = os.path.dirname(os.path.abspath(path_merged)) or "."
    return os.path.join(merged_dir, BACKUP_SUBDIR)


def resolve_backup_root(path_merged, backup_root=None):
    root = (backup_root or "").strip()
    if not root:
        root = load_saved_backup_root()
    if not root:
        root = default_backup_root(path_merged)
    return os.path.normpath(os.path.abspath(root))


def backup_project_parent(path_merged, backup_root=None):
    root = resolve_backup_root(path_merged, backup_root)
    return os.path.join(root, project_name_for_backup(path_merged))


def _create_unique_backup_dir(project_parent, timestamp=None):
    """
    创建唯一备份目录。存在检查和 mkdir 之间可能有并发竞态，
    所以实际创建失败时继续尝试下一个后缀。
    """
    os.makedirs(project_parent, exist_ok=True)
    ts = timestamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    candidates = [ts]
    candidates.extend("%s_%02d" % (ts, i) for i in range(2, 1000))
    candidates.append(datetime.now().strftime("%Y%m%d_%H%M%S_%f"))

    for used_timestamp in candidates:
        backup_dir = os.path.join(project_parent, used_timestamp)
        try:
            os.makedirs(backup_dir, exist_ok=False)
            return backup_dir, used_timestamp
        except FileExistsError:
            continue
    backup_dir = os.path.join(project_parent, datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
    os.makedirs(backup_dir, exist_ok=False)
    return backup_dir, os.path.basename(backup_dir)


def _create_unique_named_backup_dir(project_parent, label=None, timestamp=None):
    """
    创建唯一备份目录，目录名格式为 时间__label。label 为空时保持旧时间戳格式。
    """
    os.makedirs(project_parent, exist_ok=True)
    ts = timestamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    reserved_file_len = 52
    max_label_len = max(0, min(
        MAX_BACKUP_DIR_LABEL_LEN,
        MAX_BACKUP_PATH_LEN - len(os.path.abspath(project_parent)) - len(ts) - reserved_file_len - 4,
    ))
    safe_label = _sanitize_dir_name(label, "", max_len=max_label_len)
    base_name = "%s__%s" % (ts, safe_label) if safe_label else ts
    candidates = [base_name]
    candidates.extend("%s_%02d" % (base_name, i) for i in range(2, 1000))
    fallback = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    candidates.append("%s__%s" % (fallback, safe_label) if safe_label else fallback)

    for used_name in candidates:
        backup_dir = os.path.join(project_parent, used_name)
        try:
            os.makedirs(backup_dir, exist_ok=False)
            return backup_dir, used_name
        except FileExistsError:
            continue
    backup_dir = os.path.join(project_parent, fallback)
    os.makedirs(backup_dir, exist_ok=False)
    return backup_dir, os.path.basename(backup_dir)


def _git_log_info(repo_root, ref, rel_path):
    fmt = "%H%n%h%n%an%n%ae%n%ci%n%s"
    try:
        r = subprocess.run(
            ["git", "log", "-1", "--format=" + fmt, ref, "--", rel_path],
            cwd=repo_root,
            capture_output=True,
            text=True,
            timeout=5,
        )
        if r.returncode != 0 or not r.stdout.strip():
            return None
        parts = r.stdout.strip().split("\n")
        if len(parts) < 6:
            return None
        return {
            "hash": parts[0],
            "short_hash": parts[1],
            "author": parts[2],
            "email": parts[3],
            "date": parts[4],
            "message": parts[5],
            "ref": ref,
        }
    except Exception:
        return None


def get_backup_git_info(context_path):
    """
    根据真实 Excel 路径读取冲突两侧提交信息。
    返回 (local_info, remote_info)，获取失败时字段可能为 None / {}。
    """
    try:
        context_abs = os.path.abspath(context_path)
        start_dir = os.path.dirname(context_abs) if os.path.splitext(context_abs)[1] else context_abs
        repo_root, _ = discover_git_worktree_root(start_dir)
        if not repo_root:
            return None, {}
        rel_path = os.path.relpath(context_abs, repo_root).replace("\\", "/")
        if rel_path.startswith(".."):
            return None, {}
        local_info = _git_log_info(repo_root, "HEAD", rel_path)
        remote_info = None
        for ref in ("MERGE_HEAD", "REBASE_HEAD", "CHERRY_PICK_HEAD"):
            remote_info = _git_log_info(repo_root, ref, rel_path)
            if remote_info:
                break
        return local_info, remote_info or {}
    except Exception:
        return None, {}


def _commit_part(info, prefix, include_message=True):
    if not info:
        return "%s-unknown-nohash-nomsg" % prefix
    author = _safe_name_part(info.get("author"), "unknown", 28)
    short_hash = _safe_name_part(info.get("short_hash"), "nohash", 16)
    message = _safe_name_part(info.get("message"), "nomsg", 54)
    if include_message:
        return "%s-%s-%s-%s" % (prefix, author, short_hash, message)
    return "%s-%s-%s" % (prefix, author, short_hash)


def _backup_label(context_path, local_info, remote_info):
    excel_name = _safe_name_part(os.path.splitext(os.path.basename(context_path))[0], "Excel", 42)
    parts = [
        excel_name,
        _commit_part(local_info, "L", include_message=True),
        _commit_part(remote_info, "R", include_message=True),
    ]
    return _sanitize_dir_name("__".join(parts), "ExcelBackup", max_len=MAX_BACKUP_DIR_LABEL_LEN)


def _backup_file_name_with_commit(context_path, label, local_info, remote_info):
    excel_name, ext = os.path.splitext(os.path.basename(context_path))
    excel_name = _safe_name_part(excel_name, "Excel", 54)
    ext = ext or ".xlsx"
    if label == "local":
        commit_label = _commit_part(local_info, "L", include_message=True)
    elif label == "remote":
        commit_label = _commit_part(remote_info, "R", include_message=True)
    elif label == "merged":
        commit_label = "%s__%s" % (
            _commit_part(local_info, "L", include_message=False),
            _commit_part(remote_info, "R", include_message=False),
        )
    else:
        commit_label = _safe_name_part(label, label or "backup", 24)
    stem = _sanitize_dir_name(
        "%s__%s__%s" % (excel_name, label, commit_label),
        "%s_%s" % (excel_name, label),
        max_len=MAX_BACKUP_FILE_STEM_LEN,
    )
    return stem + ext


def _backup_path_for_dir(backup_dir, context_path, label, local_info, remote_info):
    name = _backup_file_name_with_commit(context_path, label, local_info, remote_info)
    full_path = os.path.join(backup_dir, name)
    if len(os.path.abspath(full_path)) <= MAX_BACKUP_PATH_LEN:
        return full_path
    base, ext = os.path.splitext(name)
    budget = MAX_BACKUP_PATH_LEN - len(os.path.abspath(backup_dir)) - len(ext) - 2
    digest = hashlib.sha1(os.path.abspath(full_path).encode("utf-8", "ignore")).hexdigest()[:10]
    min_name = "%s__%s%s" % (_safe_name_part(label, "backup", 18), digest, ext)
    if budget < len(os.path.splitext(min_name)[0]):
        return os.path.join(backup_dir, min_name)
    short_stem = _sanitize_dir_name(base, label or "backup", max_len=budget)
    if len(short_stem) + 2 + len(digest) <= budget:
        short_stem = "%s__%s" % (short_stem, digest)
    else:
        short_stem = "%s__%s" % (_sanitize_dir_name(short_stem, label or "backup", max_len=max(1, budget - len(digest) - 2)), digest)
    return os.path.join(backup_dir, short_stem + ext)


def _backup_file_name(path_merged, suffix):
    base_name, ext = os.path.splitext(os.path.basename(path_merged))
    return (base_name or "merged") + suffix + (ext or ".xlsx")


def create_merge_backup(path_local, path_remote, path_merged, backup_root=None, timestamp=None, context_path=None):
    """
    创建一次完整合并备份，目录结构：
    备份根目录 / 项目名 / 时间戳 / {合并文件名}_{local|remote|merged}.xlsx
    """
    context_path = context_path or path_merged
    project_parent = backup_project_parent(context_path, backup_root)
    local_info, remote_info = get_backup_git_info(context_path)
    label = _backup_label(context_path, local_info, remote_info)
    backup_dir, used_timestamp = _create_unique_named_backup_dir(project_parent, label, timestamp)

    paths = {
        "local": _backup_path_for_dir(backup_dir, context_path, "local", local_info, remote_info),
        "remote": _backup_path_for_dir(backup_dir, context_path, "remote", local_info, remote_info),
        "merged": _backup_path_for_dir(backup_dir, context_path, "merged", local_info, remote_info),
    }
    shutil.copy2(path_local, paths["local"])
    shutil.copy2(path_remote, paths["remote"])
    shutil.copy2(path_merged, paths["merged"])
    return {
        "root": resolve_backup_root(context_path, backup_root),
        "project": project_name_for_backup(context_path),
        "time": used_timestamp,
        "dir": backup_dir,
        "local": paths["local"],
        "remote": paths["remote"],
        "merged": paths["merged"],
        "label": label,
        "local_commit": local_info,
        "remote_commit": remote_info,
    }
