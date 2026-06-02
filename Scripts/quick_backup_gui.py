# -*- coding: utf-8 -*-
"""
快速备份模式结果面板：备份完成或失败后停留给用户确认，用户手动关闭后才返回。
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox

from gui_common import (
    ToolTip,
    apply_app_icon,
    gui_log,
    make_badge,
    make_header_icon,
    make_icon_button,
    open_containing_folder,
    setup_merge_styles,
)
from version import __version__ as APP_VERSION


MODE_NOTICES = {
    "merge": "快速备份模式只备份 LOCAL / BASE / REMOTE，不生成合并结果文件。",
    "compare": "快速备份模式只备份传入的两个 Excel，不生成对比工作簿。",
    "git-driver": "快速备份模式已备份 Git 传入文件，但不会标记冲突已解决；关闭后 Git 会保持未解决状态。",
}


class QuickBackupWindow:
    """快速备份结果面板。"""

    def __init__(self, mode, backup_info=None, error=None):
        self.mode = mode
        self.backup_info = backup_info or {}
        self.error = error
        self.root = tk.Tk()
        self.status_var = tk.StringVar(self.root, value="")
        self.root.title("快速备份模式 v%s" % APP_VERSION)
        self.root.minsize(700, 460)
        self.root.geometry("820x540")
        setup_merge_styles(self.root)
        apply_app_icon(self.root)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()

    def _build_ui(self):
        pad = 16
        shell = ttk.Frame(self.root, padding=(pad, pad, pad, 0), style="App.TFrame")
        shell.pack(fill=tk.BOTH, expand=True)

        title_row = ttk.Frame(shell, style="App.TFrame")
        title_row.pack(fill=tk.X)
        make_header_icon(title_row, self.root).pack(side=tk.LEFT, padx=(0, 10))
        title_text = ttk.Frame(title_row, style="App.TFrame")
        title_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        headline = ttk.Frame(title_text, style="App.TFrame")
        headline.pack(fill=tk.X)
        ttk.Label(headline, text="快速备份模式", style="Title.TLabel").pack(side=tk.LEFT)
        make_badge(headline, "成功" if not self.error else "失败", "success" if not self.error else "danger").pack(side=tk.LEFT, padx=(10, 0))
        self.btn_main = make_icon_button(headline, self.root, "设置中心", "app", command=self._open_main, style="Tiny.TButton")
        self.btn_main.pack(side=tk.LEFT, padx=(8, 0))
        ToolTip(self.btn_main, "打开设置中心，设置默认运行模式、备份目录、全局 Git 注入和程序更新。")
        ttk.Label(
            title_text,
            text=MODE_NOTICES.get(self.mode, "快速备份模式只做文件备份，不执行合并或对比。"),
            style="Subtitle.TLabel",
        ).pack(anchor=tk.W, pady=(2, 0))

        panel = ttk.Frame(shell, padding=(14, 14), style="Panel.TFrame")
        panel.pack(fill=tk.BOTH, expand=True, pady=(14, 0))

        if self.error:
            ttk.Label(panel, text="需要处理", style="Section.TLabel").pack(anchor=tk.W)
            msg = tk.Text(panel, height=5, wrap=tk.WORD, relief=tk.FLAT, borderwidth=0)
            msg.insert(tk.END, str(self.error))
            msg.configure(state=tk.DISABLED)
            msg.pack(fill=tk.X, pady=(8, 14))
            ttk.Label(panel, text="请检查文件路径和备份目录权限，然后在设置中心调整配置。", style="Muted.TLabel").pack(anchor=tk.W)
            return

        backup_dir = self.backup_info.get("dir") or ""
        ttk.Label(panel, text="备份目录", style="Section.TLabel").pack(anchor=tk.W)
        path_row = ttk.Frame(panel, style="Panel.TFrame")
        path_row.pack(fill=tk.X, pady=(8, 12))
        self.backup_dir_var = tk.StringVar(self.root, value=backup_dir)
        ttk.Label(path_row, textvariable=self.backup_dir_var, style="Panel.TLabel").pack(side=tk.LEFT, fill=tk.X, expand=True)
        make_icon_button(path_row, self.root, "打开目录", "folder", command=self._open_backup_dir, style="Secondary.TButton").pack(side=tk.RIGHT)

        ttk.Label(panel, text="已备份文件", style="Section.TLabel").pack(anchor=tk.W, pady=(4, 8))
        table_frame = ttk.Frame(panel, style="Panel.TFrame")
        table_frame.pack(fill=tk.BOTH, expand=True)
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(table_frame, columns=("label", "path"), show="headings", height=8)
        self.tree.heading("label", text="类型")
        self.tree.heading("path", text="备份文件")
        self.tree.column("label", width=90, minwidth=70, stretch=False)
        self.tree.column("path", width=560, minwidth=260, stretch=True)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.grid(row=0, column=1, sticky="ns")
        for label, path in sorted((self.backup_info.get("files") or {}).items()):
            self.tree.insert("", tk.END, values=(label, path))

        ttk.Label(
            panel,
            text="请确认备份已生成。关闭此窗口后，本次快速备份流程才会结束。",
            style="Muted.TLabel",
        ).pack(anchor=tk.W, pady=(10, 0))

    def _open_backup_dir(self):
        path = self.backup_info.get("dir")
        if not path or not os.path.isdir(path):
            messagebox.showwarning("提示", "备份目录不存在")
            return
        if open_containing_folder(path, select_file=False):
            gui_log("已打开备份目录：%s" % path, self.status_var)
        else:
            messagebox.showwarning("提示", "无法打开备份目录")

    def _open_main(self):
        try:
            from main_gui import open_main_window
            open_main_window(parent=self.root)
            gui_log("已打开设置中心", self.status_var)
        except Exception as e:
            gui_log("打开设置中心失败: %s" % e, self.status_var, is_error=True)
            messagebox.showerror("错误", "打开设置中心失败：%s" % e)

    def _on_close(self):
        self.root.quit()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


def show_quick_backup_panel(mode, backup_info=None, error=None):
    if os.environ.get("EXCEL_MERGE_FORK_SKIP_BACKUP_PANEL") == "1":
        return
    win = QuickBackupWindow(mode, backup_info=backup_info, error=error)
    win.run()
