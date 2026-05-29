# -*- coding: utf-8 -*-
"""
GUI 公共组件与样式：日志输出到文件并更新状态栏、颜色图例、用系统默认程序打开文件、合并/对比窗口统一样式。
只做一件事：为 Merge/Diff 窗口提供统一的写日志、界面小部件和样式，不包含业务逻辑。
"""

import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

from log_util import log_path, log
from update_manager import (
    UpdateError,
    check_for_update,
    download_update,
    get_current_executable,
    launch_update_script,
    make_update_script,
)


def gui_log_path():
    """GUI 下使用的日志文件路径（与主入口一致）。"""
    return log_path()


def gui_log(msg, log_widget=None, is_error=False):
    """
    写一行到日志文件，并可选更新 GUI 上的日志控件（如状态栏 StringVar）。
    log_widget: 若为 tk.StringVar 则 set 为单行；若为 Text 则 insert 并滚动到底。
    """
    ts = datetime.now().strftime("%H:%M:%S")
    prefix = "[ERROR] " if is_error else ""
    line = "%s %s%s" % (ts, prefix, msg)
    try:
        with open(log_path(), "a", encoding="utf-8") as f:
            f.write(datetime.now().strftime("%Y-%m-%d ") + line + "\n")
    except Exception:
        pass
    if log_widget:
        try:
            if isinstance(log_widget, tk.StringVar):
                log_widget.set(line.strip())
            else:
                log_widget.insert(tk.END, line + "\n")
                log_widget.see(tk.END)
            if hasattr(log_widget, "update_idletasks"):
                log_widget.update_idletasks()
        except Exception:
            pass


def make_color_legend(parent, items, bg="#f0f2f5"):
    """
    在 parent 下生成一行颜色图例。
    items: [(颜色 hex 或 None, "说明文字"), ...]
    """
    frame = tk.Frame(parent, bg=bg)
    for i, (color, text) in enumerate(items):
        if i > 0:
            tk.Label(frame, text="  ", bg=bg, font=("Segoe UI", 8)).pack(side=tk.LEFT)
        if color:
            patch = tk.Frame(frame, width=14, height=14, bg=color, relief=tk.SOLID, borderwidth=1)
            patch.pack(side=tk.LEFT, padx=(0, 4))
            patch.pack_propagate(False)
        tk.Label(frame, text=text, bg=bg, font=("Segoe UI", 8), fg="#1c1e21").pack(side=tk.LEFT)
    return frame


def open_excel_file(path):
    """用系统默认程序打开 Excel 文件；路径不存在或打开失败返回 False。"""
    if not path or not os.path.isfile(path):
        return False
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
        return True
    except Exception:
        return False


def open_containing_folder(path, select_file=True):
    """
    打开文件所在文件夹。
    - Windows: 优先 explorer /select 选中文件；失败则打开目录
    - macOS: open -R
    - Linux: xdg-open 目录
    """
    if not path:
        return False
    try:
        abs_path = os.path.abspath(path)
        folder = abs_path if os.path.isdir(abs_path) else os.path.dirname(abs_path)
        if not folder:
            return False
        if sys.platform == "win32":
            if select_file and os.path.isfile(abs_path):
                try:
                    subprocess.run(["explorer", "/select,", abs_path], check=False)
                    return True
                except Exception:
                    pass
            os.startfile(folder)
            return True
        if sys.platform == "darwin":
            if select_file and os.path.exists(abs_path):
                subprocess.run(["open", "-R", abs_path], check=False)
            else:
                subprocess.run(["open", folder], check=False)
            return True
        subprocess.run(["xdg-open", folder], check=False)
        return True
    except Exception:
        return False


def setup_merge_styles(root):
    """为合并/对比窗口配置统一样式：Segoe UI、浅灰背景、主按钮突出。"""
    style = ttk.Style(root)
    try:
        style.theme_use("vista")
    except tk.TclError:
        pass
    bg = "#f0f2f5"
    fg = "#1c1e21"
    style.configure("TFrame", background=bg)
    style.configure("TLabel", background=bg, foreground=fg, font=("Segoe UI", 9))
    style.configure("TLabelframe", background=bg, font=("Segoe UI", 9))
    style.configure("TLabelframe.Label", background=bg, foreground=fg, font=("Segoe UI", 9, "bold"))
    style.configure("TButton", font=("Segoe UI", 9), padding=(12, 6))
    style.configure("Accent.TButton", font=("Segoe UI", 9, "bold"), padding=(14, 7))
    style.map("Accent.TButton", background=[("active", "#166fe5")])
    style.configure("Treeview", font=("Segoe UI", 9), rowheight=22, fieldbackground="#fff")
    style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
    root.configure(bg=bg)


class UpdateButtonController:
    """把 GitHub Release 更新检查接到一个 Tk 按钮上。"""

    def __init__(self, root, button, status_var=None, on_quit=None):
        self.root = root
        self.button = button
        self.status_var = status_var
        self.on_quit = on_quit
        self.progress_container = None
        self.progress_var = None
        self.progress_bar = None
        self.progress_label = None
        self.info = None
        self.checking = False
        self.installing = False
        self.button.config(text="检查更新", command=self.on_click)

    def bind_progress_widgets(self, progress_bar, progress_var, progress_label=None, progress_container=None):
        self.progress_container = progress_container
        self.progress_bar = progress_bar
        self.progress_var = progress_var
        self.progress_label = progress_label
        self._set_progress_visible(False)

    def start_background_check(self):
        if self.checking:
            return
        self.checking = True
        self.button.config(text="检查中...", state=tk.DISABLED)
        self._set_progress_visible(True, mode="indeterminate", text="检查更新中...")

        def worker():
            try:
                info = check_for_update()
                self.root.after(0, lambda: self._set_check_result(info, silent=True))
            except Exception as e:
                self.root.after(0, lambda err=e: self._set_check_error(err, silent=True))

        threading.Thread(target=worker, daemon=True).start()

    def on_click(self):
        if self.installing:
            return
        if self.info and self.info.get("available"):
            self._confirm_and_install(self.info)
            return
        self._manual_check()

    def _manual_check(self):
        if self.checking:
            return
        self.checking = True
        self.button.config(text="检查中...", state=tk.DISABLED)
        self._set_progress_visible(True, mode="indeterminate", text="检查更新中...")
        gui_log("正在检查更新...", self.status_var)

        def worker():
            try:
                info = check_for_update()
                self.root.after(0, lambda: self._set_check_result(info, silent=False))
            except Exception as e:
                self.root.after(0, lambda err=e: self._set_check_error(err, silent=False))

        threading.Thread(target=worker, daemon=True).start()

    def _set_check_result(self, info, silent):
        self.checking = False
        self.info = info
        self.button.config(state=tk.NORMAL)
        self._set_progress_visible(False)
        if info.get("available"):
            self.button.config(text="有新版本 v%s" % info.get("latest_version"))
            gui_log("发现新版本 v%s，请点击按钮更新。" % info.get("latest_version"), self.status_var)
            if not silent:
                self._confirm_and_install(info)
            return
        self.button.config(text="检查更新")
        if not silent:
            if info.get("missing_asset"):
                messagebox.showwarning("检查更新", "最新 Release 未找到 ExcelMergeFork.exe。")
            else:
                messagebox.showinfo("检查更新", "当前已是最新版本 v%s。" % info.get("current_version"))

    def _set_check_error(self, err, silent):
        self.checking = False
        self.button.config(text="检查更新", state=tk.NORMAL)
        self._set_progress_visible(False)
        if not silent:
            messagebox.showerror("检查更新失败", str(err))
        else:
            log("后台检查更新失败: %s" % err, is_error=True)

    def _confirm_and_install(self, info):
        target = get_current_executable()
        if not target:
            messagebox.showinfo("更新提示", "当前是 Python 脚本运行模式，只有打包后的 exe 支持原地更新。")
            return
        msg = (
            "发现新版本 v%s，当前版本 v%s。\n\n"
            "点击确定后会下载新版 exe；下载完成后需要关闭当前窗口，工具会在退出后替换文件。\n"
            "如果正在合并冲突，建议先完成或取消当前合并后再更新。"
        ) % (info.get("latest_version"), info.get("current_version"))
        if not messagebox.askokcancel("发现新版本", msg):
            return
        self.installing = True
        self.button.config(text="下载中...", state=tk.DISABLED)
        self._set_progress_visible(True, mode="indeterminate", text="准备下载...")
        gui_log("正在下载 v%s..." % info.get("latest_version"), self.status_var)

        def worker():
            try:
                def on_progress(downloaded, total):
                    self.root.after(0, lambda d=downloaded, t=total: self._update_download_progress(d, t))

                new_exe = download_update(info, progress_callback=on_progress)
                script = make_update_script(new_exe, target, restart_after=False)
                self.root.after(0, lambda: self._finish_install(script))
            except Exception as e:
                self.root.after(0, lambda err=e: self._install_error(err))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_install(self, script):
        self.installing = False
        self._set_progress_visible(True, mode="determinate", value=100, text="下载完成，等待替换...")
        self.button.config(text="有新版本 v%s" % self.info.get("latest_version"), state=tk.NORMAL)
        messagebox.showinfo("更新已准备好", "更新包已下载。点击确定后会关闭当前窗口，并自动替换 ExcelMergeFork.exe。")
        launch_update_script(script)
        if self.on_quit:
            self.on_quit()
        else:
            self.root.quit()
            self.root.destroy()

    def _install_error(self, err):
        self.installing = False
        self._set_progress_visible(False)
        self.button.config(text="有新版本 v%s" % self.info.get("latest_version"), state=tk.NORMAL)
        if isinstance(err, UpdateError):
            msg = str(err)
        else:
            msg = "更新失败: %s" % err
        gui_log(msg, self.status_var, is_error=True)
        messagebox.showerror("更新失败", msg)

    def _set_progress_visible(self, visible, mode="determinate", value=0, text=""):
        if self.progress_bar is None:
            return
        if visible:
            try:
                self.progress_bar.config(mode=mode)
                if mode == "indeterminate":
                    self.progress_bar.start(12)
                else:
                    self.progress_bar.stop()
                    if self.progress_var is not None:
                        self.progress_var.set(value)
                if self.progress_label is not None:
                    self.progress_label.config(text=text)
                if self.progress_container is not None:
                    self.progress_container.pack(fill=tk.X, pady=(6, 0))
                self.progress_bar.grid()
                if self.progress_label is not None:
                    self.progress_label.grid()
            except Exception:
                pass
        else:
            try:
                self.progress_bar.stop()
                if self.progress_var is not None:
                    self.progress_var.set(0)
                if self.progress_label is not None:
                    self.progress_label.config(text="")
                    self.progress_label.grid_remove()
                self.progress_bar.grid_remove()
                if self.progress_container is not None:
                    self.progress_container.pack_forget()
            except Exception:
                pass

    def _update_download_progress(self, downloaded, total):
        if self.progress_bar is None:
            return
        if total > 0:
            percent = int(downloaded * 100 / total)
            self._set_progress_visible(True, mode="determinate", value=percent, text="下载中 %d%%" % percent)
        else:
            self._set_progress_visible(True, mode="indeterminate", text="下载中...")
