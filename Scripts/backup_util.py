# -*- coding: utf-8 -*-
"""
合并备份工具：解析备份根目录、按 项目/时间 创建快照目录，并复制本地/线上/合并结果。
"""

import json
import os
import re
import shutil
import subprocess
from datetime import datetime

from config import BACKUP_SUBDIR
from log_util import merge_options_path


BACKUP_ROOT_OPTION = "backup_root_dir"
INVALID_DIR_CHARS = r'[<>:"/\\|?*\x00-\x1f]'


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


def _sanitize_dir_name(name, fallback):
    cleaned = re.sub(INVALID_DIR_CHARS, "_", name or "")
    cleaned = cleaned.strip(" ._")
    return cleaned or fallback


def project_name_for_backup(path_merged):
    """优先使用 Git 仓库目录名；不在仓库中时使用 MERGED 所在目录名。"""
    merged_dir = os.path.dirname(os.path.abspath(path_merged)) or os.getcwd()
    project_dir = merged_dir
    try:
        r = subprocess.run(
            ["git", "rev-parse", "--show-toplevel"],
            cwd=merged_dir,
            capture_output=True,
            text=True,
            timeout=5,
        )
        if r.returncode == 0 and r.stdout.strip():
            project_dir = r.stdout.strip()
    except Exception:
        pass
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


def _unique_backup_dir(project_parent, timestamp=None):
    ts = timestamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = os.path.join(project_parent, ts)
    if not os.path.exists(backup_dir):
        return backup_dir, ts
    for i in range(2, 1000):
        ts_i = "%s_%02d" % (ts, i)
        backup_dir = os.path.join(project_parent, ts_i)
        if not os.path.exists(backup_dir):
            return backup_dir, ts_i
    ts_i = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return os.path.join(project_parent, ts_i), ts_i


def _backup_file_name(path_merged, suffix):
    base_name, ext = os.path.splitext(os.path.basename(path_merged))
    return (base_name or "merged") + suffix + (ext or ".xlsx")


def create_merge_backup(path_local, path_remote, path_merged, backup_root=None, timestamp=None):
    """
    创建一次完整合并备份，目录结构：
    备份根目录 / 项目名 / 时间戳 / {合并文件名}_{local|remote|merged}.xlsx
    """
    project_parent = backup_project_parent(path_merged, backup_root)
    backup_dir, used_timestamp = _unique_backup_dir(project_parent, timestamp)
    os.makedirs(backup_dir, exist_ok=False)

    paths = {
        "local": os.path.join(backup_dir, _backup_file_name(path_merged, "_local")),
        "remote": os.path.join(backup_dir, _backup_file_name(path_merged, "_remote")),
        "merged": os.path.join(backup_dir, _backup_file_name(path_merged, "_merged")),
    }
    shutil.copy2(path_local, paths["local"])
    shutil.copy2(path_remote, paths["remote"])
    shutil.copy2(path_merged, paths["merged"])
    return {
        "root": resolve_backup_root(path_merged, backup_root),
        "project": project_name_for_backup(path_merged),
        "time": used_timestamp,
        "dir": backup_dir,
        "local": paths["local"],
        "remote": paths["remote"],
        "merged": paths["merged"],
    }
