# -*- coding: utf-8 -*-
"""
应用级偏好设置：读取/保存启动默认功能等轻量配置。
"""

import json
import os

from log_util import merge_options_path


STARTUP_FEATURE_OPTION = "startup_feature"
STARTUP_FEATURE_BACKUP_ONLY = "backup_only"
STARTUP_FEATURE_MERGE_DIFF = "merge_diff"
STARTUP_FEATURE_ASK_EACH_TIME = "ask_each_time"
STARTUP_FEATURE_DEFAULT = STARTUP_FEATURE_MERGE_DIFF


def _load_options_data():
    try:
        path = merge_options_path()
        if path and os.path.isfile(path):
            with open(path, "r", encoding="utf-8-sig") as f:
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


def load_startup_feature():
    data = _load_options_data()
    value = data.get(STARTUP_FEATURE_OPTION, STARTUP_FEATURE_DEFAULT)
    if value == STARTUP_FEATURE_BACKUP_ONLY:
        return STARTUP_FEATURE_BACKUP_ONLY
    if value == STARTUP_FEATURE_ASK_EACH_TIME:
        return STARTUP_FEATURE_ASK_EACH_TIME
    return STARTUP_FEATURE_MERGE_DIFF


def save_startup_feature(value):
    data = _load_options_data()
    if value == STARTUP_FEATURE_BACKUP_ONLY:
        data[STARTUP_FEATURE_OPTION] = STARTUP_FEATURE_BACKUP_ONLY
    elif value == STARTUP_FEATURE_ASK_EACH_TIME:
        data[STARTUP_FEATURE_OPTION] = STARTUP_FEATURE_ASK_EACH_TIME
    else:
        data[STARTUP_FEATURE_OPTION] = STARTUP_FEATURE_MERGE_DIFF
    _save_options_data(data)
