# -*- coding: utf-8 -*-
"""
GitHub Release 更新器。
只负责检查版本、下载 release 资产、生成退出后替换 exe 的临时脚本。
"""

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request

from config import (
    GITHUB_REPO,
    UPDATE_ASSET_NAME,
    UPDATE_CHECK_INTERVAL_SECONDS,
    UPDATE_CHECK_TIMEOUT,
    UPDATE_SHA256_ASSET_NAME,
    UPDATE_STATE_FILE,
)
from log_util import log, log_dir
from version import __version__ as APP_VERSION


LATEST_RELEASE_API = "https://api.github.com/repos/%s/releases/latest" % GITHUB_REPO


def update_asset_name():
    return os.environ.get("EXCEL_MERGE_FORK_UPDATE_ASSET_NAME") or UPDATE_ASSET_NAME


def update_sha256_asset_name():
    return os.environ.get("EXCEL_MERGE_FORK_UPDATE_SHA256_ASSET_NAME") or UPDATE_SHA256_ASSET_NAME


class UpdateError(Exception):
    """更新流程异常。"""


def update_state_path():
    """返回自动更新检查缓存文件路径。"""
    return os.path.join(log_dir(), UPDATE_STATE_FILE)


def _read_update_state():
    try:
        path = update_state_path()
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return {}


def _write_update_state(state):
    try:
        path = update_state_path()
        parent = os.path.dirname(path)
        if parent:
            os.makedirs(parent, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log("写入更新检查缓存失败: %s" % e, is_error=True)


def _info_for_cache(info):
    keep = {}
    for key in (
        "available",
        "current_version",
        "latest_version",
        "html_url",
        "body",
        "published_at",
        "missing_asset",
    ):
        keep[key] = info.get(key)
    return keep


def cached_update_info():
    """读取上次检查结果。只有当前版本仍匹配时才返回缓存 info。"""
    state = _read_update_state()
    info = state.get("info")
    if not isinstance(info, dict):
        return None
    if info.get("current_version") != APP_VERSION:
        return None
    return info


def should_auto_check_update(now=None):
    """自动检查最多每 7 天触发一次；手动点击不受限制。"""
    state = _read_update_state()
    checked_at = state.get("checked_at")
    if not isinstance(checked_at, (int, float)):
        return True
    return (now or time.time()) - checked_at >= UPDATE_CHECK_INTERVAL_SECONDS


def remember_update_check(info, checked_at=None):
    """保存检查时间和轻量结果，用于启动时显示新版标识。"""
    _write_update_state({
        "checked_at": checked_at or time.time(),
        "info": _info_for_cache(info or {}),
    })


def _parse_version(text):
    parts = re.findall(r"\d+", text or "")
    return tuple(int(p) for p in parts)


def _is_newer(latest, current):
    latest_parts = _parse_version(latest)
    current_parts = _parse_version(current)
    max_len = max(len(latest_parts), len(current_parts), 1)
    latest_parts += (0,) * (max_len - len(latest_parts))
    current_parts += (0,) * (max_len - len(current_parts))
    return latest_parts > current_parts


def _request_json(url, timeout=UPDATE_CHECK_TIMEOUT):
    req = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "ExcelMergeFork/%s" % APP_VERSION,
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _download(url, path, timeout=UPDATE_CHECK_TIMEOUT):
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "ExcelMergeFork/%s" % APP_VERSION},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        with open(path, "wb") as f:
            shutil.copyfileobj(resp, f)


def _download_with_progress(url, path, timeout=UPDATE_CHECK_TIMEOUT, progress_callback=None):
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "ExcelMergeFork/%s" % APP_VERSION},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        total = resp.headers.get("Content-Length")
        total = int(total) if total and total.isdigit() else 0
        downloaded = 0
        with open(path, "wb") as f:
            while True:
                chunk = resp.read(1024 * 256)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)
                if progress_callback:
                    progress_callback(downloaded, total)


def _find_asset(release, asset_name):
    for asset in release.get("assets") or []:
        if asset.get("name") == asset_name and asset.get("browser_download_url"):
            return asset
    return None


def get_current_executable():
    """返回当前可更新目标。脚本运行时返回 None，避免覆盖 .py。"""
    launcher = os.environ.get("EXCEL_MERGE_FORK_LAUNCHER_EXE")
    if launcher:
        return os.path.abspath(launcher)
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)
    return None


def check_for_update():
    """
    检查 GitHub 最新 release。
    返回 dict: available/latest_version/html_url/body/published_at/asset。
    """
    try:
        release = _request_json(LATEST_RELEASE_API)
        latest = (release.get("tag_name") or release.get("name") or "").lstrip("vV")
        asset = _find_asset(release, update_asset_name())
        sha_asset = _find_asset(release, update_sha256_asset_name())
        available = bool(latest and asset and _is_newer(latest, APP_VERSION))
        return {
            "available": available,
            "current_version": APP_VERSION,
            "latest_version": latest,
            "html_url": release.get("html_url") or "",
            "body": release.get("body") or "",
            "published_at": release.get("published_at") or "",
            "asset": asset,
            "sha_asset": sha_asset,
            "missing_asset": asset is None,
        }
    except (urllib.error.URLError, urllib.error.HTTPError, ValueError, OSError) as e:
        log("检查更新失败: %s" % e, is_error=True)
        raise UpdateError("检查更新失败: %s" % e)


def _read_expected_sha256(asset, tmp_dir):
    if not asset:
        return None
    sha_path = os.path.join(tmp_dir, update_sha256_asset_name())
    _download(asset["browser_download_url"], sha_path)
    with open(sha_path, "r", encoding="utf-8") as f:
        text = f.read().strip()
    m = re.search(r"\b[a-fA-F0-9]{64}\b", text)
    return m.group(0).lower() if m else None


def _sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest().lower()


def download_update(info, progress_callback=None):
    """下载最新版 exe，校验 sha256（如果 release 提供），返回临时 exe 路径。"""
    asset = info.get("asset")
    if not asset:
        raise UpdateError("最新 Release 未找到 %s" % update_asset_name())

    tmp_dir = tempfile.mkdtemp(prefix="ExcelMergeForkUpdate_")
    exe_path = os.path.join(tmp_dir, update_asset_name())
    try:
        _download_with_progress(asset["browser_download_url"], exe_path, timeout=60, progress_callback=progress_callback)
        expected = _read_expected_sha256(info.get("sha_asset"), tmp_dir)
        if expected:
            actual = _sha256_file(exe_path)
            if actual != expected:
                raise UpdateError("更新包校验失败：sha256 不一致")
        return exe_path
    except Exception:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise


def make_update_script(new_exe, target_exe, restart_after=False):
    """生成替换脚本。脚本会等待当前 exe 退出后覆盖并重新启动。"""
    if not target_exe:
        raise UpdateError("当前不是 exe 运行模式，无法原地更新。")
    tmp_dir = os.path.dirname(os.path.abspath(new_exe))
    script_path = os.path.join(tmp_dir, "apply_update.bat")
    bak_path = target_exe + ".bak"
    log_path = os.path.join(log_dir(), "MergeExcelFork.update.log")
    restart_line = 'start "" "%TARGET_EXE%"' if restart_after else "rem restart skipped"
    content = """@echo off
chcp 65001 >nul
set "NEW_EXE=%s"
set "TARGET_EXE=%s"
set "BAK_EXE=%s"
set "UPDATE_LOG=%s"
echo [%%date%% %%time%%] waiting target exit... >> "%%UPDATE_LOG%%"
timeout /t 2 /nobreak >nul
for /l %%%%i in (1,1,30) do (
  copy /Y "%%TARGET_EXE%%" "%%BAK_EXE%%" >nul 2>nul
  copy /Y "%%NEW_EXE%%" "%%TARGET_EXE%%" >nul 2>nul
  if not errorlevel 1 goto updated
  timeout /t 1 /nobreak >nul
)
echo [%%date%% %%time%%] update failed: target locked >> "%%UPDATE_LOG%%"
exit /b 1
:updated
echo [%%date%% %%time%%] update ok >> "%%UPDATE_LOG%%"
%s
exit /b 0
""" % (new_exe, target_exe, bak_path, log_path, restart_line)
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(content)
    return script_path


def launch_update_script(script_path):
    """启动替换脚本。"""
    creationflags = 0
    if sys.platform == "win32":
        creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        creationflags |= getattr(subprocess, "CREATE_NO_WINDOW", 0)
    subprocess.Popen(
        ["cmd", "/c", script_path],
        cwd=os.path.dirname(script_path),
        close_fds=True,
        creationflags=creationflags,
    )
    time.sleep(0.2)
