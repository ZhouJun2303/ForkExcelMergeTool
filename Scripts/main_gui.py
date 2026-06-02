# -*- coding: utf-8 -*-
"""
设置中心窗口：集中管理默认功能、备份目录、全局 Git 注入和程序更新。
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from app_settings import (
    STARTUP_FEATURE_BACKUP_ONLY,
    STARTUP_FEATURE_MERGE_DIFF,
    load_startup_feature,
    save_startup_feature,
)
from backup_util import load_saved_backup_root, save_backup_root
from excel_format import merge_diff_extension_text
from git_integration import (
    EXCEL_EXTENSION_TEXT,
    current_executable_path,
    install_global_integration,
    integration_status,
    uninstall_global_integration,
)
from gui_common import (
    ToolTip,
    UpdateButtonController,
    apply_app_icon,
    gui_log,
    make_badge,
    make_header_icon,
    make_icon_button,
    make_separator,
    make_update_card,
    open_containing_folder,
    setup_merge_styles,
)
from version import __version__ as APP_VERSION


_main_instance = None


def get_existing_main_window():
    global _main_instance
    if _main_instance is not None and _main_instance.root.winfo_exists():
        return _main_instance
    _main_instance = None
    return None


def open_main_window(parent=None, on_update_quit=None):
    existing = get_existing_main_window()
    if existing is not None:
        if on_update_quit is not None:
            existing.on_update_quit = on_update_quit
        existing.activate()
        return existing
    win = MainWindow(parent=parent, on_update_quit=on_update_quit)
    return win


class MainWindow:
    """独立设置中心 UI。parent 存在时使用 Toplevel，否则创建 Tk 主窗口。"""

    FEATURE_LABELS = {
        STARTUP_FEATURE_BACKUP_ONLY: "快速备份模式",
        STARTUP_FEATURE_MERGE_DIFF: "合并对比模式",
    }

    def __init__(self, parent=None, on_update_quit=None):
        global _main_instance
        self.parent = parent
        self.on_update_quit = on_update_quit
        self.root = tk.Toplevel(parent) if parent is not None else tk.Tk()
        _main_instance = self
        self.status_var = tk.StringVar(self.root, value="")
        self.feature_var = tk.StringVar(self.root, value=load_startup_feature())
        self.backup_root_var = tk.StringVar(self.root, value=load_saved_backup_root())
        self.git_status_var = tk.StringVar(self.root, value="")
        self.git_detail_var = tk.StringVar(self.root, value="")
        self.update_controller = None
        self.root.title("ExcelMergeFork 设置中心 v%s" % APP_VERSION)
        self.root.minsize(780, 620)
        self.root.geometry("900x680")
        setup_merge_styles(self.root)
        apply_app_icon(self.root)
        if parent is not None:
            self.root.transient(parent)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()
        self._refresh_git_status()
        if self.update_controller:
            self.update_controller.start_background_check()

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
        ttk.Label(headline, text="ExcelMergeFork 设置中心", style="Title.TLabel").pack(side=tk.LEFT)
        make_badge(headline, "v%s" % APP_VERSION, "primary").pack(side=tk.LEFT, padx=(10, 0))
        ttk.Label(
            title_text,
            text="管理默认模式、备份目录、全局 Git 注入和程序更新。",
            style="Subtitle.TLabel",
        ).pack(anchor=tk.W, pady=(2, 0))

        body = ttk.Frame(shell, style="App.TFrame")
        body.pack(fill=tk.BOTH, expand=True, pady=(14, 0))
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        left = ttk.Frame(body, padding=(14, 14), style="Panel.TFrame")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        right = ttk.Frame(body, padding=(14, 14), style="Panel.TFrame")
        right.grid(row=0, column=1, sticky="nsew")

        ttk.Label(left, text="默认运行模式", style="Section.TLabel").pack(anchor=tk.W)
        feature_a = ttk.Radiobutton(
            left,
            text="快速备份模式",
            value=STARTUP_FEATURE_BACKUP_ONLY,
            variable=self.feature_var,
            command=self._on_feature_changed,
        )
        feature_a.pack(anchor=tk.W, pady=(10, 0))
        ToolTip(feature_a, "Fork、命令行或全局 Git 注入传入文件时只快速复制备份，不做合并或对比。")
        ttk.Label(
            left,
            text="只备份输入文件，不生成合并结果或对比工作簿。",
            style="Muted.TLabel",
        ).pack(anchor=tk.W, padx=(24, 0), pady=(2, 0))

        feature_b = ttk.Radiobutton(
            left,
            text="合并对比模式",
            value=STARTUP_FEATURE_MERGE_DIFF,
            variable=self.feature_var,
            command=self._on_feature_changed,
        )
        feature_b.pack(anchor=tk.W, pady=(12, 0))
        ToolTip(feature_b, "Fork、命令行或全局 Git 注入传入文件时打开合并/对比窗口。")
        ttk.Label(
            left,
            text="打开完整 Excel 三向合并、二向对比、冲突选择和备份流程。",
            style="Muted.TLabel",
        ).pack(anchor=tk.W, padx=(24, 0), pady=(2, 0))

        make_separator(left).pack(fill=tk.X, pady=16)

        ttk.Label(left, text="备份根目录", style="Section.TLabel").pack(anchor=tk.W)
        backup_row = ttk.Frame(left, style="Panel.TFrame")
        backup_row.pack(fill=tk.X, pady=(10, 0))
        entry = ttk.Entry(backup_row, textvariable=self.backup_root_var)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        make_icon_button(backup_row, self.root, "选择", "folder", command=self._choose_backup_root, style="Secondary.TButton").pack(side=tk.LEFT)
        ToolTip(entry, "留空时使用目标文件同目录下的 MergeExcelBackup。")
        button_row = ttk.Frame(left, style="Panel.TFrame")
        button_row.pack(fill=tk.X, pady=(8, 0))
        make_icon_button(button_row, self.root, "保存设置", "backup", command=self._save_backup_root, style="Secondary.TButton").pack(side=tk.LEFT, padx=(0, 8))
        make_icon_button(button_row, self.root, "打开目录", "folder", command=self._open_backup_root, style="Secondary.TButton").pack(side=tk.LEFT)

        current = self.FEATURE_LABELS.get(self.feature_var.get(), self.FEATURE_LABELS[STARTUP_FEATURE_MERGE_DIFF])
        self.current_feature_var = tk.StringVar(self.root, value="当前默认：%s" % current)
        ttk.Label(left, textvariable=self.current_feature_var, style="Panel.TLabel").pack(anchor=tk.W, pady=(16, 0))

        make_separator(left).pack(fill=tk.X, pady=16)
        ttk.Label(left, text="全局 Git 注入", style="Section.TLabel").pack(anchor=tk.W)
        ttk.Label(
            left,
            text="安装后，任意 Git 工具在常见 Excel 后缀冲突时都会调用本工具，并按上方默认运行模式处理。",
            style="Muted.TLabel",
            wraplength=430,
        ).pack(anchor=tk.W, pady=(8, 0))
        ttk.Label(left, textvariable=self.git_status_var, style="Panel.TLabel").pack(anchor=tk.W, pady=(8, 0))
        ttk.Label(left, textvariable=self.git_detail_var, style="Muted.TLabel", wraplength=430).pack(anchor=tk.W, pady=(2, 0))
        git_btn_row = ttk.Frame(left, style="Panel.TFrame")
        git_btn_row.pack(fill=tk.X, pady=(8, 0))
        make_icon_button(git_btn_row, self.root, "安装注入", "check", command=self._install_git_integration, style="Secondary.TButton").pack(side=tk.LEFT, padx=(0, 8))
        make_icon_button(git_btn_row, self.root, "移除注入", "cancel", command=self._uninstall_git_integration, style="Secondary.TButton").pack(side=tk.LEFT, padx=(0, 8))
        make_icon_button(git_btn_row, self.root, "刷新状态", "refresh", command=self._refresh_git_status, style="Secondary.TButton").pack(side=tk.LEFT)

        (
            update_card,
            btn_update,
            self.update_state_var,
            self.update_state_label,
            self.update_state_icon,
            update_progress_row,
            self.update_progress_var,
            self.update_progress_label,
            self.update_progress_bar,
        ) = make_update_card(right, self.root, include_button=True)
        update_card.pack(fill=tk.X)
        self.update_controller = UpdateButtonController(
            self.root,
            btn_update,
            status_var=self.status_var,
            on_quit=self._quit_for_update,
        )
        self.update_controller.bind_state_widget(self.update_state_var, self.update_state_label, self.update_state_icon)
        self.update_controller.bind_progress_widgets(
            self.update_progress_bar,
            self.update_progress_var,
            self.update_progress_label,
            update_progress_row,
        )

        make_separator(right).pack(fill=tk.X, pady=16)
        ttk.Label(right, text="启动方式", style="Section.TLabel").pack(anchor=tk.W)
        ttk.Label(
            right,
            text="双击 exe 会打开设置中心。有文件参数时，工具会按左侧默认运行模式进入快速备份或合并对比。",
            style="Muted.TLabel",
            wraplength=260,
        ).pack(anchor=tk.W, pady=(8, 0))

        bottom = ttk.Frame(self.root, padding=(pad, 10), style="BottomBar.TFrame")
        bottom.pack(fill=tk.X)
        ttk.Label(bottom, textvariable=self.status_var, style="Panel.TLabel").pack(side=tk.LEFT, fill=tk.X, expand=True)
        make_icon_button(bottom, self.root, "关闭", "cancel", command=self._on_close, style="Secondary.TButton").pack(side=tk.RIGHT)

    def _on_feature_changed(self):
        value = self.feature_var.get()
        save_startup_feature(value)
        label = self.FEATURE_LABELS.get(value, self.FEATURE_LABELS[STARTUP_FEATURE_MERGE_DIFF])
        self.current_feature_var.set("当前默认：%s" % label)
        gui_log("默认功能已保存：%s" % label, self.status_var)

    def _choose_backup_root(self):
        initial_dir = self.backup_root_var.get().strip()
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = os.getcwd()
        folder = filedialog.askdirectory(
            parent=self.root,
            title="选择备份根目录",
            initialdir=initial_dir,
        )
        if folder:
            self.backup_root_var.set(os.path.normpath(folder))
            self._save_backup_root(show_message=False)

    def _save_backup_root(self, show_message=True):
        root_dir = self.backup_root_var.get().strip()
        try:
            save_backup_root(root_dir)
            msg = "备份根目录已保存：%s" % (root_dir or "默认目录")
            gui_log(msg, self.status_var)
            if show_message:
                messagebox.showinfo("已保存", msg)
        except Exception as e:
            gui_log("保存备份根目录失败: %s" % e, self.status_var, is_error=True)
            messagebox.showerror("错误", "保存备份根目录失败：%s" % e)

    def _open_backup_root(self):
        root_dir = self.backup_root_var.get().strip()
        if not root_dir:
            messagebox.showinfo("备份目录", "当前使用默认目录：目标文件同目录下的 MergeExcelBackup。")
            return
        try:
            os.makedirs(root_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", "无法创建备份目录：%s" % e)
            return
        if not open_containing_folder(root_dir, select_file=False):
            messagebox.showwarning("提示", "无法打开备份目录")

    def _refresh_git_status(self):
        try:
            status = integration_status()
            exe_path = current_executable_path()
            if not status.get("git_available"):
                self.git_status_var.set("状态：未检测到 git 命令")
            elif status.get("installed"):
                self.git_status_var.set("状态：已安装全局 Git 注入")
            else:
                self.git_status_var.set("状态：未安装或配置不完整")
            self.git_detail_var.set("目标 exe：%s\nAttributes：%s\n注入后缀：%s\n合并对比支持：%s" % (
                exe_path,
                status.get("attributes_file") or "",
                EXCEL_EXTENSION_TEXT,
                merge_diff_extension_text(),
            ))
        except Exception as e:
            self.git_status_var.set("状态：读取失败")
            self.git_detail_var.set(str(e))

    def _install_git_integration(self):
        exe_path = current_executable_path()
        if not os.path.isfile(exe_path):
            messagebox.showwarning(
                "无法安装",
                "未找到 ExcelMergeFork.exe：\n%s\n\n请先打包 exe，或把 exe 放在项目根目录后再安装全局 Git 注入。" % exe_path,
            )
            self._refresh_git_status()
            return
        if not messagebox.askokcancel(
            "安装全局 Git 注入",
            "这会写入用户级 Git 配置，并让所有 Git 工具在常见 Excel 后缀冲突时调用：\n%s\n\n快速备份模式可直接备份这些文件；合并对比模式依赖当前 Excel 解析能力。" % exe_path,
        ):
            return
        try:
            install_global_integration(exe_path)
            gui_log("已安装全局 Git 注入", self.status_var)
            self._refresh_git_status()
            messagebox.showinfo("已安装", "全局 Git 注入已安装。之后常见 Excel 后缀冲突会自动打开本工具。")
        except Exception as e:
            gui_log("安装全局 Git 注入失败: %s" % e, self.status_var, is_error=True)
            self._refresh_git_status()
            messagebox.showerror("安装失败", str(e))

    def _uninstall_git_integration(self):
        if not messagebox.askokcancel("移除全局 Git 注入", "这会移除 ExcelMergeFork 写入的用户级 Git 配置和 attributes 条目。"):
            return
        try:
            uninstall_global_integration()
            gui_log("已移除全局 Git 注入", self.status_var)
            self._refresh_git_status()
            messagebox.showinfo("已移除", "全局 Git 注入已移除。")
        except Exception as e:
            gui_log("移除全局 Git 注入失败: %s" % e, self.status_var, is_error=True)
            self._refresh_git_status()
            messagebox.showerror("移除失败", str(e))

    def activate(self):
        self.root.lift()
        self.root.focus_force()

    def _on_close(self):
        global _main_instance
        if _main_instance is self:
            _main_instance = None
        if self.parent is None:
            self.root.quit()
        self.root.destroy()

    def _quit_for_update(self):
        global _main_instance
        if _main_instance is self:
            _main_instance = None
        if self.on_update_quit:
            try:
                self.root.destroy()
            except Exception:
                pass
            self.on_update_quit()
        else:
            self.root.quit()
            self.root.destroy()

    def run(self):
        self.root.mainloop()
