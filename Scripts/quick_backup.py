# -*- coding: utf-8 -*-
"""
快速备份：只复制输入文件，不读取 Excel，也不生成合并/对比结果。
"""

import os
import shutil

from backup_util import (
    _backup_label,
    _backup_path_for_dir,
    _create_unique_named_backup_dir,
    _safe_name_part,
    backup_project_parent,
    get_backup_git_info,
)
from log_util import log


def _backup_file_name(src_path, label, context_path=None, local_info=None, remote_info=None):
    base_name, ext = os.path.splitext(os.path.basename(src_path))
    base_name = _safe_name_part(base_name, label or "backup", 72)
    label = _safe_name_part(label, "backup", 24)
    return "%s_%s%s" % (base_name, label, ext or ".xlsx")


def _copy_existing_files(file_items, backup_dir, context_path=None, local_info=None, remote_info=None):
    copied = {}
    used_names = set()
    for label, src_path in file_items:
        if not src_path or not os.path.isfile(src_path):
            continue
        commit_label = {
            "current": "local",
            "other": "remote",
        }.get(label, label)
        if context_path and commit_label in ("local", "remote", "merged"):
            dst = _backup_path_for_dir(backup_dir, context_path, commit_label, local_info, remote_info)
            candidate = os.path.basename(dst)
        else:
            candidate = _backup_file_name(src_path, label, context_path, local_info, remote_info)
            dst = os.path.join(backup_dir, candidate)
        base, ext = os.path.splitext(candidate)
        i = 2
        while candidate.lower() in used_names:
            candidate = "%s_%02d%s" % (base, i, ext)
            i += 1
        used_names.add(candidate.lower())
        dst = os.path.join(backup_dir, candidate)
        shutil.copy2(src_path, dst)
        copied[label] = dst
    return copied


def create_quick_backup(file_items, context_path, backup_root=None, timestamp=None):
    """
    file_items: [("label", path), ...]。不存在的文件会跳过。
    context_path: 用于决定备份根目录/项目名的逻辑路径。
    """
    project_parent = backup_project_parent(context_path, backup_root)
    local_info, remote_info = get_backup_git_info(context_path)
    label = _backup_label(context_path, local_info, remote_info)
    backup_dir, used_timestamp = _create_unique_named_backup_dir(project_parent, label, timestamp)
    copied = _copy_existing_files(file_items, backup_dir, context_path, local_info, remote_info)
    if not copied:
        raise FileNotFoundError("没有可备份的输入文件")
    info = {
        "dir": backup_dir,
        "time": used_timestamp,
        "files": copied,
        "label": label,
        "local_commit": local_info,
        "remote_commit": remote_info,
    }
    log("[QuickBackup] 完成 dir=%s files=%s" % (backup_dir, sorted(copied)))
    return info


def quick_backup_merge(path_local, path_base, path_remote, path_merged, backup_root=None, context_path=None):
    context = context_path or path_merged or path_local
    items = [
        ("local", path_local),
        ("base", path_base),
        ("remote", path_remote),
        ("merged", path_merged),
    ]
    return create_quick_backup(items, context, backup_root=backup_root)


def quick_backup_compare(path_a, path_b, backup_root=None):
    items = [
        ("a", path_a),
        ("b", path_b),
    ]
    return create_quick_backup(items, path_a or path_b, backup_root=backup_root)


def quick_backup_git_driver(path_base, path_current, path_other, context_path=None, backup_root=None):
    for label, path in (("base", path_base), ("current", path_current), ("other", path_other)):
        if not path or not os.path.isfile(path):
            raise FileNotFoundError("%s 不存在: %s" % (label, path))
    items = [
        ("base", path_base),
        ("current", path_current),
        ("other", path_other),
    ]
    return create_quick_backup(items, context_path or path_current, backup_root=backup_root)
