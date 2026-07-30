# -*- coding: utf-8 -*-
"""
设置中心窗口：集中管理默认功能、备份目录、全局 Git 注入和程序更新。
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from app_settings import (
    STARTUP_FEATURE_ASK_EACH_TIME,
    STARTUP_FEATURE_BACKUP_ONLY,
    STARTUP_FEATURE_MERGE_DIFF,
    load_startup_feature,
    save_startup_feature,
)
from backup_util import load_saved_backup_root, save_backup_root
from companion_tools import find_external_merge_tools, launch_external_or_repo
from excel_format import merge_diff_extension_text
from fork_integration import (
    install_fork_integration,
    integration_status as fork_integration_status,
    uninstall_fork_integration,
)
from git_integration import (
    ATTR_LINES,
    EXCEL_EXTENSION_TEXT,
    current_executable_path,
    driver_command,
    install_global_integration,
    integration_status,
    uninstall_global_integration,
)
from gui_common import (
    ToolTip,
    UpdateButtonController,
    UI,
    apply_app_icon,
    gui_log,
    install_global_button_loading,
    make_badge,
    make_header_icon,
    make_icon_button,
    make_separator,
    make_update_card,
    open_containing_folder,
    run_loading_task,
    set_global_busy,
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
        STARTUP_FEATURE_ASK_EACH_TIME: "每次询问",
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
        self.fork_status_var = tk.StringVar(self.root, value="")
        self.fork_detail_var = tk.StringVar(self.root, value="")
        self.git_status_var = tk.StringVar(self.root, value="")
        self.git_detail_var = tk.StringVar(self.root, value="")
        self.integration_progress_var = tk.IntVar(self.root, value=0)
        self.integration_busy = False
        self._copy_vars = []
        self.update_controller = None
        self.root.title("ExcelMergeFork 设置中心 v%s" % APP_VERSION)
        self.root.minsize(820, 680)
        self.root.geometry("960x740")
        setup_merge_styles(self.root)
        apply_app_icon(self.root)
        if parent is not None:
            self.root.transient(parent)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()
        install_global_button_loading(self.root)
        self._refresh_integration_status()
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
        self.btn_external_merge_tools = make_icon_button(
            headline,
            self.root,
            "打开 ExternalMergeTools",
            "open",
            command=self._open_external_merge_tools,
            style="Secondary.TButton",
        )
        self.btn_external_merge_tools.pack(side=tk.RIGHT)
        ttk.Label(
            title_text,
            text="管理默认模式、备份目录、Fork/Git 注入和程序更新。",
            style="Subtitle.TLabel",
        ).pack(anchor=tk.W, pady=(2, 0))

        body = ttk.Frame(shell, style="App.TFrame")
        body.pack(fill=tk.BOTH, expand=True, pady=(14, 0))
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        left = self._make_scroll_panel(body, 0, padx=(0, 12))
        right = self._make_scroll_panel(body, 1)

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

        feature_c = ttk.Radiobutton(
            left,
            text="每次询问",
            value=STARTUP_FEATURE_ASK_EACH_TIME,
            variable=self.feature_var,
            command=self._on_feature_changed,
        )
        feature_c.pack(anchor=tk.W, pady=(12, 0))
        ToolTip(feature_c, "Fork、命令行或全局 Git 注入传入文件时，先选择快速备份或合并/对比。")
        ttk.Label(
            left,
            text="适合大表先打开备份、小表再进入合并或对比。",
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
        ttk.Label(left, text="Fork 一键注入", style="Section.TLabel").pack(anchor=tk.W)
        ttk.Label(
            left,
            text="安装后，Fork 的 Merge Tool 和 External Diff Tool 会直接指向本工具。请先关闭 Fork 再安装或移除。",
            style="Muted.TLabel",
            wraplength=430,
        ).pack(anchor=tk.W, pady=(8, 0))
        ttk.Label(left, textvariable=self.fork_status_var, style="Panel.TLabel").pack(anchor=tk.W, pady=(8, 0))
        ttk.Label(left, textvariable=self.fork_detail_var, style="Muted.TLabel", wraplength=430).pack(anchor=tk.W, pady=(2, 0))
        fork_btn_row = ttk.Frame(left, style="Panel.TFrame")
        fork_btn_row.pack(fill=tk.X, pady=(8, 0))
        self.btn_fork_install = make_icon_button(fork_btn_row, self.root, "安装注入", "check", command=self._install_fork_integration, style="Secondary.TButton")
        self.btn_fork_install.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_fork_uninstall = make_icon_button(fork_btn_row, self.root, "移除注入", "cancel", command=self._uninstall_fork_integration, style="Secondary.TButton")
        self.btn_fork_uninstall.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_integration_refresh = make_icon_button(fork_btn_row, self.root, "刷新状态", "refresh", command=self._refresh_integration_status, style="Secondary.TButton")
        self.btn_integration_refresh.pack(side=tk.LEFT)

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
        self.btn_git_install = make_icon_button(git_btn_row, self.root, "安装注入", "check", command=self._install_git_integration, style="Secondary.TButton")
        self.btn_git_install.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_git_uninstall = make_icon_button(git_btn_row, self.root, "移除注入", "cancel", command=self._uninstall_git_integration, style="Secondary.TButton")
        self.btn_git_uninstall.pack(side=tk.LEFT, padx=(0, 8))
        self.integration_progress_row = ttk.Frame(left, style="Panel.TFrame")
        self.integration_progress_label = ttk.Label(self.integration_progress_row, text="", style="Muted.TLabel")
        self.integration_progress_label.pack(side=tk.LEFT, padx=(0, 8))
        self.integration_progress_bar = ttk.Progressbar(self.integration_progress_row, variable=self.integration_progress_var, mode="indeterminate", length=180)
        self.integration_progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)

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
        self._build_integration_guide(right)

        make_separator(right).pack(fill=tk.X, pady=16)
        ttk.Label(right, text="启动方式", style="Section.TLabel").pack(anchor=tk.W)
        ttk.Label(
            right,
            text="双击 exe 会打开设置中心。有文件参数时，工具会按左侧默认运行模式处理；选择“每次询问”时会先弹窗确认。",
            style="Muted.TLabel",
            wraplength=260,
        ).pack(anchor=tk.W, pady=(8, 0))

        bottom = ttk.Frame(self.root, padding=(pad, 10), style="BottomBar.TFrame")
        bottom.pack(fill=tk.X)
        ttk.Label(bottom, textvariable=self.status_var, style="Panel.TLabel").pack(side=tk.LEFT, fill=tk.X, expand=True)
        make_icon_button(bottom, self.root, "关闭", "cancel", command=self._on_close, style="Secondary.TButton").pack(side=tk.RIGHT)

    def _make_scroll_panel(self, parent, column, padx=(0, 0)):
        outer = ttk.Frame(parent, style="Panel.TFrame")
        outer.grid(row=0, column=column, sticky="nsew", padx=padx)
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        canvas = tk.Canvas(outer, bg=UI["panel"], highlightthickness=0, bd=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)

        inner = ttk.Frame(canvas, padding=(14, 14), style="Panel.TFrame")
        canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")

        def on_frame_configure(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas_configure(event):
            canvas.itemconfigure(canvas_window, width=event.width)

        def on_mousewheel(event):
            delta = -1 * int(event.delta / 120) if event.delta else 0
            if delta:
                canvas.yview_scroll(delta, "units")

        inner.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        canvas.bind("<MouseWheel>", on_mousewheel)
        inner.bind("<MouseWheel>", on_mousewheel)
        return inner

    def _build_integration_guide(self, parent):
        ttk.Label(parent, text="接入教程", style="Section.TLabel").pack(anchor=tk.W)
        ttk.Label(
            parent,
            text="Fork 优先使用一键注入；需要手动配置时复制下面的路径和参数。",
            style="Muted.TLabel",
            wraplength=280,
        ).pack(anchor=tk.W, pady=(8, 0))

        tabs = ttk.Notebook(parent)
        tabs.pack(fill=tk.X, pady=(8, 0))

        fork_tab = ttk.Frame(tabs, padding=(8, 8), style="Panel.TFrame")
        github_tab = ttk.Frame(tabs, padding=(8, 8), style="Panel.TFrame")
        tabs.add(fork_tab, text="Fork")
        tabs.add(github_tab, text="GitHub Desktop")

        exe_path = current_executable_path()
        self._make_copy_row(fork_tab, "工具路径", exe_path, "Merge/Diff Tool Path")
        ttk.Label(
            fork_tab,
            text="推荐点击左侧「Fork 一键注入」。手动配置时：Merge Tool 选 Custom。",
            style="Muted.TLabel",
            wraplength=285,
        ).pack(anchor=tk.W, pady=(6, 0))
        self._make_copy_row(fork_tab, "合并参数", "$LOCAL,$BASE,$REMOTE,$MERGED", "Arguments")
        ttk.Label(
            fork_tab,
            text="External Diff Tool: Diff Tool 选 Custom，路径同上。",
            style="Muted.TLabel",
            wraplength=285,
        ).pack(anchor=tk.W, pady=(6, 0))
        self._make_copy_row(fork_tab, "对比参数", "\"$REMOTE\" \"$LOCAL\"", "Arguments")

        ttk.Label(
            github_tab,
            text="GitHub Desktop 没有单独参数输入框。优先点击左侧「全局 Git 注入」安装；需要手动配置时复制下面的 driver 值。",
            style="Muted.TLabel",
            wraplength=285,
        ).pack(anchor=tk.W)
        self._make_copy_row(github_tab, "driver 值", driver_command(exe_path), "git config value")
        self._make_copy_row(github_tab, "driver 参数", "--git-merge-driver \"%O\" \"%A\" \"%B\" \"%P\"", "args")
        attrs_row = ttk.Frame(github_tab, style="Panel.TFrame")
        attrs_row.pack(fill=tk.X, pady=(8, 0))
        make_icon_button(
            attrs_row,
            self.root,
            "复制 attributes",
            "copy",
            command=lambda: self._copy_text("\n".join(ATTR_LINES), "attributes"),
            style="Tiny.TButton",
        ).pack(side=tk.LEFT)
        ttk.Label(
            attrs_row,
            text="匹配 %s" % EXCEL_EXTENSION_TEXT,
            style="Muted.TLabel",
            wraplength=170,
        ).pack(side=tk.LEFT, padx=(8, 0))

    def _make_copy_row(self, parent, label_text, value, placeholder):
        ttk.Label(parent, text=label_text, style="Panel.TLabel").pack(anchor=tk.W, pady=(8, 2))
        row = ttk.Frame(parent, style="Panel.TFrame")
        row.pack(fill=tk.X)
        value_var = tk.StringVar(self.root, value=value)
        self._copy_vars.append(value_var)
        entry = ttk.Entry(row, textvariable=value_var, state="readonly")
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ToolTip(entry, placeholder)
        make_icon_button(
            row,
            self.root,
            "复制",
            "copy",
            command=lambda: self._copy_text(value, label_text),
            style="Tiny.TButton",
        ).pack(side=tk.LEFT)

    def _copy_text(self, value, label_text="内容"):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(value)
            self.root.update_idletasks()
            gui_log("已复制：%s" % label_text, self.status_var)
        except Exception as e:
            gui_log("复制失败: %s" % e, self.status_var, is_error=True)
            messagebox.showerror("复制失败", str(e))

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

    def _refresh_integration_status(self):
        if self.integration_busy:
            return
        self._run_integration_task("正在刷新注入状态...", self._read_integration_status, self._finish_integration_refresh)

    def _read_integration_status(self):
        return {
            "fork": fork_integration_status(),
            "git": integration_status(),
        }

    def _apply_integration_status(self, statuses):
        self._apply_fork_status(statuses.get("fork") or {})
        self._apply_git_status(statuses.get("git") or {})

    def _apply_fork_status(self, status):
        try:
            if not status.get("settings_exists"):
                self.fork_status_var.set("状态：未检测到 Fork 配置")
            elif status.get("installed"):
                self.fork_status_var.set("状态：已安装 Fork 注入")
            else:
                self.fork_status_var.set("状态：未安装或配置不完整")
            self.fork_detail_var.set("目标程序：%s\nFork 配置：%s\n合并参数：%s\n对比参数：%s" % (
                status.get("tool_path") or current_executable_path(),
                status.get("settings_path") or "",
                "$LOCAL,$BASE,$REMOTE,$MERGED",
                "\"$REMOTE\" \"$LOCAL\"",
            ))
        except Exception as e:
            self.fork_status_var.set("状态：读取失败")
            self.fork_detail_var.set(str(e))

    def _apply_git_status(self, status):
        try:
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

    def _finish_integration_refresh(self, statuses):
        self._apply_integration_status(statuses)
        gui_log("注入状态已刷新", self.status_var)

    def _install_fork_integration(self):
        exe_path = current_executable_path()
        if not os.path.isfile(exe_path):
            messagebox.showwarning(
                "无法安装",
                "未找到工具程序：\n%s\n\n请确认当前分发包完整，或重新下载 exe。" % exe_path,
            )
            self._refresh_integration_status()
            return
        if not messagebox.askokcancel(
            "安装 Fork 注入",
            "这会备份并修改 Fork 的 settings.json，把 Merge Tool 和 External Diff Tool 指向：\n%s\n\n请先关闭 Fork，避免 Fork 退出时覆盖本次写入。" % exe_path,
        ):
            return
        self._run_integration_task(
            "正在安装 Fork 注入...",
            lambda: install_fork_integration(exe_path),
            self._finish_fork_install,
            error_title="安装 Fork 注入失败",
        )

    def _uninstall_fork_integration(self):
        if not messagebox.askokcancel(
            "移除 Fork 注入",
            "这会备份并修改 Fork 的 settings.json，只移除 ExcelMergeFork 写入的 Fork 工具配置。请先关闭 Fork。",
        ):
            return
        self._run_integration_task(
            "正在移除 Fork 注入...",
            lambda: uninstall_fork_integration(current_executable_path()),
            self._finish_fork_uninstall,
            error_title="移除 Fork 注入失败",
        )

    def _install_git_integration(self):
        exe_path = current_executable_path()
        if not os.path.isfile(exe_path):
            messagebox.showwarning(
                "无法安装",
                "未找到 ExcelMergeFork.exe：\n%s\n\n请先打包 exe，或把 exe 放在项目根目录后再安装全局 Git 注入。" % exe_path,
            )
            self._refresh_integration_status()
            return
        if not messagebox.askokcancel(
            "安装全局 Git 注入",
            "这会写入用户级 Git 配置，并让所有 Git 工具在常见 Excel 后缀冲突时调用：\n%s\n\n快速备份模式可直接备份这些文件；合并对比模式依赖当前 Excel 解析能力。" % exe_path,
        ):
            return
        self._run_integration_task(
            "正在安装全局 Git 注入...",
            lambda: install_global_integration(exe_path),
            self._finish_git_install,
            error_title="安装失败",
        )

    def _uninstall_git_integration(self):
        if not messagebox.askokcancel("移除全局 Git 注入", "这会移除 ExcelMergeFork 写入的用户级 Git 配置和 attributes 条目。"):
            return
        self._run_integration_task(
            "正在移除全局 Git 注入...",
            uninstall_global_integration,
            self._finish_git_uninstall,
            error_title="移除失败",
        )

    def _finish_fork_install(self, status):
        self._apply_fork_status(status)
        gui_log("已安装 Fork 注入", self.status_var)
        messagebox.showinfo("已安装", "Fork 注入已安装。重新打开 Fork 后即可使用。")

    def _finish_fork_uninstall(self, status):
        self._apply_fork_status(status)
        gui_log("已移除 Fork 注入", self.status_var)
        messagebox.showinfo("已移除", "Fork 注入已移除。")

    def _finish_git_install(self, status):
        self._apply_git_status(status)
        gui_log("已安装全局 Git 注入", self.status_var)
        messagebox.showinfo("已安装", "全局 Git 注入已安装。之后常见 Excel 后缀冲突会自动打开本工具。")

    def _finish_git_uninstall(self, status):
        self._apply_git_status(status)
        gui_log("已移除全局 Git 注入", self.status_var)
        messagebox.showinfo("已移除", "全局 Git 注入已移除。")

    def _set_integration_busy(self, busy, text=""):
        self.integration_busy = busy
        set_global_busy(self.root, "integration", busy, text or "正在处理注入...")
        state = tk.DISABLED if busy else tk.NORMAL
        for btn in (
            getattr(self, "btn_fork_install", None),
            getattr(self, "btn_fork_uninstall", None),
            getattr(self, "btn_integration_refresh", None),
            getattr(self, "btn_git_install", None),
            getattr(self, "btn_git_uninstall", None),
        ):
            if btn is not None:
                btn.config(state=state)
        if busy:
            self.integration_progress_label.config(text=text)
            self.integration_progress_row.pack(fill=tk.X, pady=(8, 0))
            self.integration_progress_bar.start(12)
        else:
            self.integration_progress_bar.stop()
            self.integration_progress_var.set(0)
            self.integration_progress_label.config(text="")
            self.integration_progress_row.pack_forget()

    def _run_integration_task(self, busy_text, worker_func, success_func, error_title="操作失败"):
        if self.integration_busy:
            return
        self._set_integration_busy(True, busy_text)
        run_loading_task(
            self.root,
            "integration_worker",
            busy_text,
            worker_func,
            lambda result: self._finish_integration_task(result, success_func),
            lambda err: self._finish_integration_task_error(err, error_title),
        )

    def _finish_integration_task(self, result, success_func):
        self._set_integration_busy(False)
        success_func(result)

    def _finish_integration_task_error(self, err, error_title):
        self._set_integration_busy(False)
        gui_log("%s: %s" % (error_title, err), self.status_var, is_error=True)
        messagebox.showerror(error_title, str(err))

    def _open_external_merge_tools(self):
        result = launch_external_or_repo()
        target = result.get("path") or result.get("url")
        gui_log("已打开：%s" % target, self.status_var)
        if not result.get("launched"):
            messagebox.showinfo("获取 ExternalMergeTools", "未检测到 ExternalMergeTools，已打开对应 GitHub 仓库。")

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
