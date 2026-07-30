# -*- coding: utf-8 -*-
"""Register ExcelMergeFork and launch its ExternalMergeTools companion."""

import json
import os
import subprocess
import sys
import tempfile
import time
import webbrowser
from datetime import datetime


APP_EXTERNAL = "external_merge_tools"
APP_EXCEL = "excel_merge_fork"
EXTERNAL_REPO_URL = "https://github.com/ZhouJun2303/ExternalMergeTools"
EXTERNAL_FORK_ARGS = '--fork-merge "$LOCAL" "$BASE" "$REMOTE" "$MERGED"'


def _registry_dir():
    override = os.environ.get("MERGE_TOOLS_HUB_DIR")
    if override:
        return os.path.abspath(override)
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "MergeToolsHub")


def _registry_path():
    return os.path.join(_registry_dir(), "registry.json")


def _read_registry():
    try:
        with open(_registry_path(), "r", encoding="utf-8-sig") as stream:
            value = json.load(stream)
        return value if isinstance(value, dict) else {}
    except (OSError, ValueError):
        return {}


def _write_registry(value):
    parent = _registry_dir()
    os.makedirs(parent, exist_ok=True)
    handle, temporary = tempfile.mkstemp(prefix="registry-", suffix=".tmp", dir=parent)
    try:
        with os.fdopen(handle, "w", encoding="utf-8") as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
        os.replace(temporary, _registry_path())
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


class _RegistryLock:
    def __enter__(self):
        os.makedirs(_registry_dir(), exist_ok=True)
        self.path = os.path.join(_registry_dir(), "registry.lock")
        deadline = time.time() + 2.0
        while True:
            try:
                self.fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                return self
            except FileExistsError:
                if time.time() >= deadline:
                    raise RuntimeError("伴随工具登记文件正被占用")
                time.sleep(0.05)

    def __exit__(self, exc_type, exc_value, traceback):
        try:
            os.close(self.fd)
        finally:
            try:
                os.unlink(self.path)
            except OSError:
                pass


def current_excel_launcher():
    launcher = os.environ.get("EXCEL_MERGE_FORK_LAUNCHER_EXE")
    if launcher and os.path.isfile(launcher):
        return os.path.abspath(launcher)
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)
    root = os.environ.get("EXCEL_MERGE_FORK_HOME") or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for name in ("ExcelMergeFork.exe", "ExcelMergeFork-lite.exe", "ExcelMergeFork-python.cmd", "MergeExcelFork.py"):
        path = os.path.join(root, name)
        if os.path.isfile(path):
            return os.path.abspath(path)
    return ""


def register_current_app(version):
    launcher = current_excel_launcher()
    if not launcher:
        return False
    with _RegistryLock():
        value = _read_registry()
        value["schema_version"] = 1
        value.setdefault("apps", {})[APP_EXCEL] = {
            "launcher_path": launcher,
            "version": version,
            "updated_at": datetime.utcnow().replace(microsecond=0).isoformat() + "Z",
        }
        _write_registry(value)
    return True


def _registered_external_path():
    item = (_read_registry().get("apps") or {}).get(APP_EXTERNAL) or {}
    path = item.get("launcher_path") or ""
    return os.path.abspath(path) if path and os.path.isfile(path) else ""


def _fork_external_paths():
    local = os.environ.get("LOCALAPPDATA")
    if not local:
        return []
    path = os.path.join(local, "Fork", "settings.json")
    try:
        with open(path, "r", encoding="utf-8-sig") as stream:
            settings = json.load(stream)
    except (OSError, ValueError):
        return []
    result = []
    for tool in settings.get("ExternalMergeTools") or []:
        if isinstance(tool, dict) and tool.get("Name") == "ExternalMergeTools" and tool.get("Path"):
            result.append(tool["Path"])
    selection = settings.get("MergeTool") or {}
    if isinstance(selection, dict) and selection.get("Arguments") == EXTERNAL_FORK_ARGS:
        result.append(selection.get("ApplicationPath"))
    return result


def find_external_merge_tools():
    current = current_excel_launcher()
    current_dir = os.path.dirname(current) if current else os.getcwd()
    parent = os.path.dirname(current_dir)
    candidates = [_registered_external_path()]
    candidates.extend(_fork_external_paths())
    for root in (current_dir, os.path.join(parent, "ExternalMergeTools")):
        candidates.extend(os.path.join(root, name) for name in ("ExternalMergeTools.exe", "ExternalMergeTools-lite.exe", "ExternalMergeTools-python.cmd", "ExternalMergeTools.py"))
    for path in candidates:
        if path and os.path.isfile(path):
            return os.path.abspath(path)
    return ""


def launch_external_or_repo():
    path = find_external_merge_tools()
    if not path:
        webbrowser.open(EXTERNAL_REPO_URL)
        return {"launched": False, "url": EXTERNAL_REPO_URL, "path": ""}
    command = [path, "--main"]
    if os.path.splitext(path)[1].lower() == ".py":
        command.insert(0, sys.executable)
    elif os.path.splitext(path)[1].lower() in (".cmd", ".bat"):
        command = [os.environ.get("COMSPEC") or "cmd.exe", "/c"] + command
    subprocess.Popen(command, cwd=os.path.dirname(path))
    return {"launched": True, "url": "", "path": path}


def dispatcher_selection(settings):
    selection = settings.get("MergeTool") or {}
    path = selection.get("ApplicationPath") if isinstance(selection, dict) else ""
    return (
        isinstance(selection, dict)
        and selection.get("Type") == "Custom"
        and selection.get("Arguments") == EXTERNAL_FORK_ARGS
        and path
        and os.path.isfile(path)
    )
