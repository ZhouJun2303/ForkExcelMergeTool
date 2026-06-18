# -*- coding: utf-8 -*-
"""
启动模式选择窗口：默认运行模式为“每次询问”时，按本次文件让用户选择快速备份或合并/对比。
"""

import os
import tkinter as tk
from tkinter import ttk

from app_settings import STARTUP_FEATURE_BACKUP_ONLY, STARTUP_FEATURE_MERGE_DIFF
from gui_common import (
    UI,
    apply_app_icon,
    make_header_icon,
    make_icon_button,
    setup_merge_styles,
)
from version import __version__ as APP_VERSION


SCENE_TITLES = {
    "merge": "Excel 三向合并",
    "compare": "Excel 二向对比",
    "git-driver": "Git Excel 冲突",
}

SCENE_NOTICES = {
    "merge": "Fork 或命令行传入了三方合并文件。大表建议先快速备份，小表可进入合并窗口。",
    "compare": "Fork 或命令行传入了两个对比文件。大表建议先快速备份，小表可进入对比窗口。",
    "git-driver": "Git merge driver 传入了冲突文件。快速备份会保留冲突未解决，合并确认后 Git 才会继续。",
}

MERGE_BUTTON_TEXT = {
    "merge": "进入合并模式",
    "compare": "进入对比模式",
    "git-driver": "进入合并模式",
}


def _format_size(path):
    try:
        if not path or not os.path.isfile(path):
            return "未生成/不存在"
        size = float(os.path.getsize(path))
    except Exception:
        return "无法读取"
    units = ["B", "KB", "MB", "GB", "TB"]
    idx = 0
    while size >= 1024 and idx < len(units) - 1:
        size /= 1024.0
        idx += 1
    if idx == 0:
        return "%d %s" % (int(size), units[idx])
    return "%.1f %s" % (size, units[idx])


def _file_name(path):
    name = os.path.basename(path or "")
    return name or "(无文件名)"


def _center_window(root, width=760, height=560):
    try:
        root.update_idletasks()
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        x = max(0, int((screen_w - width) / 2))
        y = max(0, int((screen_h - height) / 2))
        root.geometry("%dx%d+%d+%d" % (width, height, x, y))
    except Exception:
        root.geometry("%dx%d" % (width, height))


class StartupChoiceWindow:
    def __init__(self, scene, file_items, unsupported_paths=None, supported_extensions=""):
        self.scene = scene
        self.file_items = list(file_items or [])
        self.unsupported_paths = list(unsupported_paths or [])
        self.supported_extensions = supported_extensions
        self.choice = None
        self.root = tk.Tk()
        self.root.title("选择运行模式 - ExcelMergeFork v%s" % APP_VERSION)
        self.root.minsize(760, 540)
        setup_merge_styles(self.root)
        apply_app_icon(self.root)
        self.root.protocol("WM_DELETE_WINDOW", self._cancel)
        self._build_ui()
        _center_window(self.root)
        try:
            self.root.focus_force()
        except Exception:
            pass

    def _build_ui(self):
        pad = 16
        btn_row = ttk.Frame(self.root, padding=(pad, 10, pad, 12), style="BottomBar.TFrame")
        btn_row.pack(fill=tk.X, side=tk.BOTTOM)

        content = ttk.Frame(self.root, padding=(pad, pad, pad, 0), style="App.TFrame")
        content.pack(fill=tk.BOTH, expand=True)

        title_row = ttk.Frame(content, style="App.TFrame")
        title_row.pack(fill=tk.X)
        make_header_icon(title_row, self.root).pack(side=tk.LEFT, padx=(0, 10))
        title_text = ttk.Frame(title_row, style="App.TFrame")
        title_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(
            title_text,
            text="本次打开方式",
            style="Title.TLabel",
        ).pack(anchor=tk.W)
        ttk.Label(
            title_text,
            text=SCENE_TITLES.get(self.scene, "Excel 文件处理"),
            style="Subtitle.TLabel",
        ).pack(anchor=tk.W, pady=(2, 0))

        panel = ttk.Frame(content, padding=(14, 14, 14, 14), style="Panel.TFrame")
        panel.pack(fill=tk.BOTH, expand=True, pady=(14, 0))
        ttk.Label(
            panel,
            text=SCENE_NOTICES.get(self.scene, "请选择本次使用快速备份还是合并/对比。"),
            style="Panel.TLabel",
            wraplength=690,
        ).pack(anchor=tk.W)

        table_frame = ttk.Frame(panel, style="Panel.TFrame")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(12, 0))
        columns = ("role", "name", "size", "path")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=7)
        tree.heading("role", text="角色")
        tree.heading("name", text="文件名")
        tree.heading("size", text="大小")
        tree.heading("path", text="路径")
        tree.column("role", width=82, stretch=False, anchor=tk.W)
        tree.column("name", width=180, stretch=False, anchor=tk.W)
        tree.column("size", width=92, stretch=False, anchor=tk.W)
        tree.column("path", width=380, stretch=True, anchor=tk.W)
        scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scroll_y.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        for role, path in self.file_items:
            tree.insert("", tk.END, values=(role, _file_name(path), _format_size(path), path or ""))

        warning = self._unsupported_message()
        if warning:
            ttk.Label(
                panel,
                text=warning,
                style="Muted.TLabel",
                wraplength=690,
            ).pack(anchor=tk.W, pady=(10, 0))
        else:
            ttk.Label(
                panel,
                text="提示：表格较大或不确定时可先快速备份；表格较小时可继续进入合并/对比流程。",
                style="Muted.TLabel",
                wraplength=690,
            ).pack(anchor=tk.W, pady=(10, 0))

        make_icon_button(
            btn_row,
            self.root,
            "快速备份",
            "backup",
            command=self._choose_backup,
            style="Secondary.TButton",
        ).pack(side=tk.LEFT, padx=(0, 8))
        merge_button = make_icon_button(
            btn_row,
            self.root,
            MERGE_BUTTON_TEXT.get(self.scene, "进入合并/对比"),
            "merge",
            command=self._choose_merge_diff,
            style="Accent.TButton",
        )
        if self.unsupported_paths:
            merge_button.config(state=tk.DISABLED)
        merge_button.pack(side=tk.LEFT)
        make_icon_button(
            btn_row,
            self.root,
            "取消",
            "cancel",
            command=self._cancel,
            style="Secondary.TButton",
        ).pack(side=tk.RIGHT)

    def _unsupported_message(self):
        if not self.unsupported_paths:
            return ""
        text = "\n".join(self.unsupported_paths)
        return (
            "这些文件后缀暂不支持合并/对比，已禁用合并/对比入口：\n%s\n"
            "当前合并/对比支持：%s。请使用快速备份保留文件。"
        ) % (text, self.supported_extensions or ".xlsx, .xltx")

    def _choose_backup(self):
        self.choice = STARTUP_FEATURE_BACKUP_ONLY
        self.root.destroy()

    def _choose_merge_diff(self):
        self.choice = STARTUP_FEATURE_MERGE_DIFF
        self.root.destroy()

    def _cancel(self):
        self.choice = None
        self.root.destroy()

    def run(self):
        self.root.mainloop()
        return self.choice


def choose_startup_feature(scene, file_items, unsupported_paths=None, supported_extensions=""):
    win = StartupChoiceWindow(
        scene,
        file_items,
        unsupported_paths=unsupported_paths,
        supported_extensions=supported_extensions,
    )
    return win.run()
