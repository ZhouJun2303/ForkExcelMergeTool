# -*- coding: utf-8 -*-
"""
全局 Git merge driver 注入/移除：让任意 Git 工具遇到常见 Excel 后缀冲突时调用本工具。
"""

import os
import subprocess
import sys

from excel_format import COMMON_EXCEL_EXTENSIONS, extension_text, git_attr_lines


DRIVER_NAME = "excelmergefork"
EXCEL_EXTENSIONS = COMMON_EXCEL_EXTENSIONS
EXCEL_EXTENSION_TEXT = extension_text()
ATTR_LINES = git_attr_lines()
LEGACY_ATTR_LINES = [
    "*.xlsx merge=excelmergefork",
    "*.XLSX merge=excelmergefork",
]


def _run_git(args, check=False):
    return subprocess.run(
        ["git"] + list(args),
        capture_output=True,
        text=True,
        check=check,
    )


def _to_git_path(path):
    return os.path.abspath(path).replace("\\", "/")


def current_executable_path():
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)
    return os.path.abspath(os.path.join(os.path.dirname(os.path.dirname(__file__)), "ExcelMergeFork.exe"))


def driver_command(exe_path=None):
    exe = _to_git_path(exe_path or current_executable_path())
    return '"%s" --git-merge-driver "%%O" "%%A" "%%B" "%%P"' % exe


def attributes_file_path():
    r = _run_git(["config", "--global", "--get", "core.attributesFile"])
    if r.returncode == 0 and r.stdout.strip():
        return os.path.expandvars(os.path.expanduser(r.stdout.strip()))
    return os.path.join(os.path.expanduser("~"), ".config", "git", "attributes")


def _read_lines(path):
    if not os.path.isfile(path):
        return []
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            return f.read().splitlines()
    except UnicodeDecodeError:
        with open(path, "r") as f:
            return f.read().splitlines()


def _write_lines(path, lines):
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
        if lines:
            f.write("\n")


def git_available():
    try:
        r = _run_git(["--version"])
        return r.returncode == 0
    except OSError:
        return False


def integration_status(exe_path=None):
    attrs_path = attributes_file_path()
    attr_lines = _read_lines(attrs_path)
    driver = _run_git(["config", "--global", "--get", "merge.%s.driver" % DRIVER_NAME])
    name = _run_git(["config", "--global", "--get", "merge.%s.name" % DRIVER_NAME])
    recursive = _run_git(["config", "--global", "--get", "merge.%s.recursive" % DRIVER_NAME])
    expected = driver_command(exe_path)
    driver_value = driver.stdout.strip() if driver.returncode == 0 else ""
    installed = (
        driver_value == expected
        and name.returncode == 0
        and recursive.returncode == 0
        and all(line in attr_lines for line in ATTR_LINES)
    )
    return {
        "installed": installed,
        "git_available": git_available(),
        "driver": driver_value,
        "expected_driver": expected,
        "attributes_file": attrs_path,
        "attributes_installed": all(line in attr_lines for line in ATTR_LINES),
    }


def install_global_integration(exe_path=None):
    exe_path = os.path.abspath(exe_path or current_executable_path())
    if not os.path.isfile(exe_path):
        raise FileNotFoundError("ExcelMergeFork.exe 不存在: %s" % exe_path)
    command = driver_command(exe_path)
    _run_git(["config", "--global", "merge.%s.name" % DRIVER_NAME, "ExcelMergeFork workbook merge driver"], check=True)
    _run_git(["config", "--global", "merge.%s.driver" % DRIVER_NAME, command], check=True)
    _run_git(["config", "--global", "merge.%s.recursive" % DRIVER_NAME, "binary"], check=True)
    attrs_path = attributes_file_path()
    lines = _read_lines(attrs_path)
    changed = False
    lines = [line for line in lines if line not in LEGACY_ATTR_LINES]
    for line in ATTR_LINES:
        if line not in lines:
            lines.append(line)
            changed = True
    if changed or not os.path.isfile(attrs_path):
        _write_lines(attrs_path, lines)
    return integration_status(exe_path)


def uninstall_global_integration():
    for key in ("name", "driver", "recursive"):
        _run_git(["config", "--global", "--unset-all", "merge.%s.%s" % (DRIVER_NAME, key)])
    attrs_path = attributes_file_path()
    if os.path.isfile(attrs_path):
        remove = set(ATTR_LINES + LEGACY_ATTR_LINES)
        lines = [line for line in _read_lines(attrs_path) if line not in remove]
        _write_lines(attrs_path, lines)
    return integration_status()
