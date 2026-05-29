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


UI = {
    "bg": "#F4F7FB",
    "panel": "#FFFFFF",
    "panel_alt": "#F8FAFC",
    "border": "#D8E0EA",
    "text": "#172033",
    "muted": "#64748B",
    "primary": "#2563EB",
    "primary_active": "#1D4ED8",
    "success": "#15803D",
    "success_bg": "#DCFCE7",
    "warning": "#B45309",
    "warning_bg": "#FEF3C7",
    "danger": "#B91C1C",
    "danger_bg": "#FEE2E2",
    "deleted": "#64748B",
    "deleted_bg": "#EEF2F7",
    "row_alt": "#F8FAFC",
}


def ui_font(size=9, weight="normal"):
    """统一 UI 字体，优先使用 Windows 中文界面更自然的微软雅黑。"""
    return ("Microsoft YaHei UI", size, weight)


def make_color_legend(parent, items, bg=None):
    """
    在 parent 下生成一行颜色图例。
    items: [(颜色 hex 或 None, "说明文字"), ...]
    """
    bg = bg or UI["panel"]
    frame = tk.Frame(parent, bg=bg)
    for i, (color, text) in enumerate(items):
        if i > 0:
            tk.Label(frame, text="  ", bg=bg, font=ui_font(8)).pack(side=tk.LEFT)
        if color:
            patch = tk.Frame(frame, width=14, height=14, bg=color, relief=tk.SOLID, borderwidth=1)
            patch.pack(side=tk.LEFT, padx=(0, 4))
            patch.pack_propagate(False)
        tk.Label(frame, text=text, bg=bg, font=ui_font(8), fg=UI["muted"]).pack(side=tk.LEFT)
    return frame


def make_badge(parent, text, tone="neutral"):
    """创建一个小状态徽标。tone: primary|success|warning|danger|neutral。"""
    colors = {
        "primary": (UI["primary"], "#DBEAFE"),
        "success": (UI["success"], UI["success_bg"]),
        "warning": (UI["warning"], UI["warning_bg"]),
        "danger": (UI["danger"], UI["danger_bg"]),
        "neutral": (UI["muted"], UI["deleted_bg"]),
    }
    fg, bg = colors.get(tone, colors["neutral"])
    return tk.Label(
        parent,
        text=text,
        fg=fg,
        bg=bg,
        font=ui_font(8, "bold"),
        padx=8,
        pady=2,
        relief=tk.FLAT,
    )


def make_separator(parent):
    line = tk.Frame(parent, bg=UI["border"], height=1)
    line.pack_propagate(False)
    return line


class ToolTip:
    """轻量 tooltip，避免引入额外依赖。"""

    def __init__(self, widget, text, delay=450):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._after_id = None
        self._tip = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self):
        if self._tip is not None or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        except Exception:
            return
        tip = tk.Toplevel(self.widget)
        tip.wm_overrideredirect(True)
        tip.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(
            tip,
            text=self.text,
            justify=tk.LEFT,
            bg="#111827",
            fg="#F9FAFB",
            font=ui_font(8),
            padx=8,
            pady=5,
            wraplength=360,
        )
        label.pack()
        self._tip = tip

    def _hide(self, _event=None):
        self._cancel()
        if self._tip is not None:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


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
    """为合并/对比窗口配置统一现代桌面样式。"""
    style = ttk.Style(root)
    try:
        style.theme_use("vista")
    except tk.TclError:
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

    bg = UI["bg"]
    panel = UI["panel"]
    fg = UI["text"]
    muted = UI["muted"]

    style.configure(".", font=ui_font(9))
    style.configure("TFrame", background=bg)
    style.configure("App.TFrame", background=bg)
    style.configure("Panel.TFrame", background=panel, relief=tk.FLAT)
    style.configure("Card.TFrame", background=UI["panel_alt"], relief=tk.FLAT)
    style.configure("Toolbar.TFrame", background=panel)
    style.configure("BottomBar.TFrame", background=panel)
    style.configure("TLabel", background=bg, foreground=fg, font=ui_font(9))
    style.configure("Panel.TLabel", background=panel, foreground=fg, font=ui_font(9))
    style.configure("Card.TLabel", background=UI["panel_alt"], foreground=fg, font=ui_font(9))
    style.configure("CardMuted.TLabel", background=UI["panel_alt"], foreground=muted, font=ui_font(8))
    style.configure("CardSection.TLabel", background=UI["panel_alt"], foreground=fg, font=ui_font(10, "bold"))
    style.configure("Muted.TLabel", background=panel, foreground=muted, font=ui_font(8))
    style.configure("Title.TLabel", background=bg, foreground=fg, font=ui_font(16, "bold"))
    style.configure("PanelTitle.TLabel", background=panel, foreground=fg, font=ui_font(16, "bold"))
    style.configure("Section.TLabel", background=panel, foreground=fg, font=ui_font(10, "bold"))
    style.configure("TLabelframe", background=panel, borderwidth=1, relief=tk.SOLID)
    style.configure("TLabelframe.Label", background=panel, foreground=fg, font=ui_font(9, "bold"))
    style.configure("TButton", font=ui_font(9), padding=(12, 6))
    style.configure("Secondary.TButton", font=ui_font(9), padding=(10, 6))
    style.configure("Tiny.TButton", font=ui_font(8), padding=(7, 3))
    style.configure("Accent.TButton", font=ui_font(9, "bold"), padding=(14, 7))
    style.map(
        "TButton",
        foreground=[("disabled", "#64748B"), ("!disabled", fg)],
        background=[("active", "#E2E8F0"), ("disabled", "#F1F5F9"), ("!disabled", "#FFFFFF")],
    )
    style.map(
        "Secondary.TButton",
        foreground=[("disabled", "#64748B"), ("!disabled", fg)],
        background=[("active", "#E2E8F0"), ("disabled", "#F1F5F9"), ("!disabled", "#FFFFFF")],
    )
    style.map(
        "Tiny.TButton",
        foreground=[("disabled", "#64748B"), ("!disabled", fg)],
        background=[("active", "#E2E8F0"), ("disabled", "#F1F5F9"), ("!disabled", "#FFFFFF")],
    )
    style.map(
        "Accent.TButton",
        foreground=[("disabled", "#64748B"), ("!disabled", fg)],
        background=[("active", "#DBEAFE"), ("disabled", "#F1F5F9"), ("!disabled", "#FFFFFF")],
    )
    style.configure("TCheckbutton", background=panel, foreground=fg, font=ui_font(9), padding=(2, 4))
    style.configure("TRadiobutton", background=panel, foreground=fg, font=ui_font(9), padding=(2, 4))
    style.configure("TCombobox", font=ui_font(9), padding=(8, 4))
    style.configure("TEntry", font=ui_font(9), padding=(8, 4))
    style.configure(
        "Treeview",
        font=ui_font(9),
        rowheight=28,
        fieldbackground=panel,
        background=panel,
        foreground=fg,
        borderwidth=0,
    )
    style.configure("Treeview.Heading", font=ui_font(9, "bold"), background=UI["panel_alt"], foreground=fg)
    style.map(
        "Treeview",
        background=[("selected", "#DBEAFE")],
        foreground=[("selected", fg)],
    )
    root.configure(bg=bg)


class UpdateButtonController:
    """把 GitHub Release 更新检查接到一个 Tk 按钮上。"""

    def __init__(self, root, button, status_var=None, on_quit=None, compact=False):
        self.root = root
        self.button = button
        self.status_var = status_var
        self.on_quit = on_quit
        self.compact = compact
        self.default_text = "更新" if compact else "检查更新"
        self.checking_text = "检查中" if compact else "检查中..."
        self.available_text = "新版" if compact else None
        self.downloading_text = "下载中" if compact else "下载中..."
        self.progress_container = None
        self.progress_var = None
        self.progress_bar = None
        self.progress_label = None
        self.info = None
        self.checking = False
        self.installing = False
        self.button.config(text=self.default_text, command=self.on_click)

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
        self.button.config(text=self.checking_text, state=tk.DISABLED)
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
        self.button.config(text=self.checking_text, state=tk.DISABLED)
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
            text = self.available_text or ("有新版本 v%s" % info.get("latest_version"))
            self.button.config(text=text)
            gui_log("发现新版本 v%s，请点击按钮更新。" % info.get("latest_version"), self.status_var)
            if not silent:
                self._confirm_and_install(info)
            return
        self.button.config(text=self.default_text)
        if not silent:
            if info.get("missing_asset"):
                messagebox.showwarning("检查更新", "最新 Release 未找到 ExcelMergeFork.exe。")
            else:
                messagebox.showinfo("检查更新", "当前已是最新版本 v%s。" % info.get("current_version"))

    def _set_check_error(self, err, silent):
        self.checking = False
        self.button.config(text=self.default_text, state=tk.NORMAL)
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
        self.button.config(text=self.downloading_text, state=tk.DISABLED)
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
        self.button.config(text=(self.available_text or ("有新版本 v%s" % self.info.get("latest_version"))), state=tk.NORMAL)
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
        self.button.config(text=(self.available_text or ("有新版本 v%s" % self.info.get("latest_version"))), state=tk.NORMAL)
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
