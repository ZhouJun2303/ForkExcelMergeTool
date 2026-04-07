# -*- coding: utf-8 -*-
"""
GUI 公共组件与样式：日志输出到文件并更新状态栏、颜色图例、用系统默认程序打开文件、合并/对比窗口统一样式。
只做一件事：为 Merge/Diff 窗口提供统一的写日志、界面小部件和样式，不包含业务逻辑。
"""

import os
import subprocess
import sys
import tkinter as tk
from tkinter import ttk
from datetime import datetime

from log_util import log_path, log


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
