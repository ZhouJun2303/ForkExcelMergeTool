# -*- coding: utf-8 -*-
"""
Fork 客户端 Integration 配置注入：写入 Fork settings.json 中的 Merge/Diff Tool。
"""

import json
import os
import shutil
import sys
from datetime import datetime


TOOL_NAME = "ExcelMergeFork"
MERGE_ARGS = "$LOCAL,$BASE,$REMOTE,$MERGED"
DIFF_ARGS = '"$REMOTE" "$LOCAL"'


def fork_settings_path():
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise EnvironmentError("无法读取 LOCALAPPDATA 环境变量，不能定位 Fork settings.json。")
    return os.path.join(local_app_data, "Fork", "settings.json")


def default_tool_path():
    launcher = os.environ.get("EXCEL_MERGE_FORK_LAUNCHER_EXE")
    if launcher and os.path.isfile(launcher):
        return os.path.abspath(launcher)

    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)

    root = os.environ.get("EXCEL_MERGE_FORK_HOME") or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for name in ("ExcelMergeFork.exe", "ExcelMergeFork-lite.exe", "ExcelMergeFork-python.cmd"):
        path = os.path.join(root, name)
        if os.path.isfile(path):
            return path
    return os.path.join(root, "ExcelMergeFork.exe")


def _norm_path(path):
    return os.path.normcase(os.path.abspath(path or ""))


def _read_settings(path):
    with open(path, "r", encoding="utf-8-sig") as f:
        return json.load(f)


def _write_settings(path, settings):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
        f.write("\n")


def _ensure_list(settings, key):
    value = settings.get(key)
    if isinstance(value, list):
        return value
    settings[key] = []
    return settings[key]


def _upsert_tool(tools, name, path, arguments, set_primary=False):
    result = []
    found = False
    target_path = _norm_path(path)

    for tool in tools:
        if not isinstance(tool, dict):
            continue

        same_tool = (
            tool.get("Type") == "Custom"
            and (tool.get("Name") == name or _norm_path(tool.get("Path")) == target_path)
        )
        if same_tool:
            if found:
                continue
            tool["Type"] = "Custom"
            tool["Name"] = name
            tool["Path"] = path
            tool["Arguments"] = arguments
            if set_primary:
                tool["IsPrimary"] = True
            found = True
        elif set_primary and "IsPrimary" in tool:
            tool["IsPrimary"] = False
        result.append(tool)

    if not found:
        tool = {
            "Type": "Custom",
            "Name": name,
            "Path": path,
            "Arguments": arguments,
        }
        if set_primary:
            tool["IsPrimary"] = True
        result.append(tool)

    return result


def _is_managed_tool(tool, name, path):
    if not isinstance(tool, dict):
        return False
    return (
        tool.get("Type") == "Custom"
        and (tool.get("Name") == name or _norm_path(tool.get("Path")) == _norm_path(path))
    )


def _is_selected_tool(tool, path, arguments):
    return (
        isinstance(tool, dict)
        and tool.get("Type") == "Custom"
        and _norm_path(tool.get("ApplicationPath")) == _norm_path(path)
        and tool.get("Arguments") == arguments
    )


def _looks_like_excelmergefork_path(path):
    name = os.path.basename(path or "").lower()
    return name in (
        "excelmergefork.exe",
        "excelmergefork-lite.exe",
        "excelmergefork-python.cmd",
    )


def _is_managed_selection(tool, paths, arguments):
    if not isinstance(tool, dict) or tool.get("Type") != "Custom" or tool.get("Arguments") != arguments:
        return False
    app_path = tool.get("ApplicationPath")
    if _looks_like_excelmergefork_path(app_path):
        return True
    return _norm_path(app_path) in set(_norm_path(path) for path in paths if path)


def install_fork_integration(tool_path=None, settings_path=None, tool_name=TOOL_NAME, backup=True):
    tool_path = os.path.abspath(tool_path or default_tool_path())
    settings_path = os.path.abspath(settings_path or fork_settings_path())

    if not os.path.isfile(tool_path):
        raise FileNotFoundError("未找到工具文件: %s" % tool_path)
    if not os.path.isfile(settings_path):
        raise FileNotFoundError("未找到 Fork 配置文件: %s。请先启动并退出一次 Fork。" % settings_path)

    backup_path = None
    if backup:
        backup_path = "%s.ExcelMergeForkBackup.%s" % (
            settings_path,
            datetime.now().strftime("%Y%m%d_%H%M%S"),
        )
        shutil.copy2(settings_path, backup_path)

    settings = _read_settings(settings_path)
    merge_tools = _ensure_list(settings, "ExternalMergeTools")
    diff_tools = _ensure_list(settings, "ExternalDiffTools")

    settings["ExternalMergeTools"] = _upsert_tool(
        merge_tools,
        tool_name,
        tool_path,
        MERGE_ARGS,
        set_primary=True,
    )
    settings["ExternalDiffTools"] = _upsert_tool(
        diff_tools,
        tool_name,
        tool_path,
        DIFF_ARGS,
        set_primary=False,
    )
    settings["MergeTool"] = {
        "Type": "Custom",
        "ApplicationPath": tool_path,
        "Arguments": MERGE_ARGS,
    }
    settings["ExternalDiffTool"] = {
        "Type": "Custom",
        "ApplicationPath": tool_path,
        "Arguments": DIFF_ARGS,
    }

    _write_settings(settings_path, settings)
    return integration_status(tool_path=tool_path, settings_path=settings_path, backup_path=backup_path)


def uninstall_fork_integration(tool_path=None, settings_path=None, tool_name=TOOL_NAME, backup=True):
    tool_path = os.path.abspath(tool_path or default_tool_path())
    settings_path = os.path.abspath(settings_path or fork_settings_path())

    if not os.path.isfile(settings_path):
        raise FileNotFoundError("未找到 Fork 配置文件: %s。请先启动并退出一次 Fork。" % settings_path)

    backup_path = None
    if backup:
        backup_path = "%s.ExcelMergeForkBackup.%s" % (
            settings_path,
            datetime.now().strftime("%Y%m%d_%H%M%S"),
        )
        shutil.copy2(settings_path, backup_path)

    settings = _read_settings(settings_path)
    merge_tools = _ensure_list(settings, "ExternalMergeTools")
    diff_tools = _ensure_list(settings, "ExternalDiffTools")
    managed_paths = [tool_path]
    for tool in list(merge_tools) + list(diff_tools):
        if _is_managed_tool(tool, tool_name, tool_path) and tool.get("Path"):
            managed_paths.append(tool.get("Path"))

    settings["ExternalMergeTools"] = [
        tool for tool in merge_tools if not _is_managed_tool(tool, tool_name, tool_path)
    ]
    settings["ExternalDiffTools"] = [
        tool for tool in diff_tools if not _is_managed_tool(tool, tool_name, tool_path)
    ]

    if _is_managed_selection(settings.get("MergeTool"), managed_paths, MERGE_ARGS):
        settings["MergeTool"] = None
    if _is_managed_selection(settings.get("ExternalDiffTool"), managed_paths, DIFF_ARGS):
        settings["ExternalDiffTool"] = None

    _write_settings(settings_path, settings)
    return integration_status(tool_path=tool_path, settings_path=settings_path, backup_path=backup_path)


def integration_status(tool_path=None, settings_path=None, backup_path=None):
    tool_path = os.path.abspath(tool_path or default_tool_path())
    settings_path = os.path.abspath(settings_path or fork_settings_path())
    status = {
        "installed": False,
        "settings_exists": os.path.isfile(settings_path),
        "tool_exists": os.path.isfile(tool_path),
        "tool_path": tool_path,
        "settings_path": settings_path,
        "backup_path": backup_path,
        "merge_configured": False,
        "diff_configured": False,
    }
    if not status["settings_exists"]:
        return status

    settings = _read_settings(settings_path)
    merge_tool = settings.get("MergeTool") or {}
    diff_tool = settings.get("ExternalDiffTool") or {}

    status["merge_configured"] = (
        isinstance(merge_tool, dict)
        and merge_tool.get("Type") == "Custom"
        and _norm_path(merge_tool.get("ApplicationPath")) == _norm_path(tool_path)
        and merge_tool.get("Arguments") == MERGE_ARGS
    )
    status["diff_configured"] = (
        isinstance(diff_tool, dict)
        and diff_tool.get("Type") == "Custom"
        and _norm_path(diff_tool.get("ApplicationPath")) == _norm_path(tool_path)
        and diff_tool.get("Arguments") == DIFF_ARGS
    )
    status["installed"] = status["merge_configured"] and status["diff_configured"]
    return status
