# -*- coding: utf-8 -*-
"""
对比窗口：二向对比的 GUI。与合并工具同步：无右下角版本、列表行按状态着色、对比后自动打开可持久化。
"""

import json
import os
import re
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox

from config import MAX_TREEVIEW_ROWS
from compare_core import get_compare_data, write_compare_excel
from gui_common import (
    ToolTip,
    UI,
    UpdateButtonController,
    apply_app_icon,
    configure_button_icon,
    gui_log,
    make_header_icon,
    make_badge,
    make_color_legend,
    make_icon_button,
    make_update_card,
    open_containing_folder,
    open_excel_file,
    setup_merge_styles,
)
from log_util import log
from log_util import merge_options_path, release_compare_lock
from version import __version__ as APP_VERSION


def _load_auto_open_compare():
    """从 merge_options.json 读取「对比后自动打开」是否勾选，默认 True。"""
    try:
        path = merge_options_path()
        if path and os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return bool(data.get("auto_open_compare", True))
    except Exception:
        pass
    return True


def _save_auto_open_compare(value):
    """将「对比后自动打开」写入本地，与 merge_options.json 合并。"""
    try:
        path = merge_options_path()
        if not path:
            return
        data = {}
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        data["auto_open_compare"] = bool(value)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# 对比 GUI 默认只加载差异行，避免大文件中海量“相同”行拖慢界面。
DIFF_STATUS_OPTIONS = [
    ("新增行", "新增行"),
    ("删除行", "删除行"),
    ("新增列", "新增列"),
    ("删除列", "删除列"),
    ("修改", "修改"),
]

# 旧配置键名兼容
_LEGACY_DIFF_FILTER_MAP = {"B新增": "新增行", "A独有": "删除行"}


def _load_diff_filter():
    """从 merge_options.json 读取「对比筛选显示」勾选，默认全部 True。"""
    try:
        path = merge_options_path()
        if path and os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            raw = data.get("diff_filter") or {}
            out = {}
            for k, _ in DIFF_STATUS_OPTIONS:
                val = raw.get(k)
                if val is None:
                    old_key = next((ok for ok, nk in _LEGACY_DIFF_FILTER_MAP.items() if nk == k), None)
                    val = raw.get(old_key) if old_key else None
                out[k] = bool(val if val is not None else k != "相同")
            return out
    except Exception:
        pass
    return {k: k != "相同" for k, _ in DIFF_STATUS_OPTIONS}


def _save_diff_filter(visible):
    """将「对比筛选显示」写入本地，与 merge_options.json 合并。"""
    try:
        path = merge_options_path()
        if not path:
            return
        data = {}
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        data["diff_filter"] = visible
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# 单实例：当前对比窗口引用，关闭时清空
_diff_instance = None


def get_existing_diff_window():
    """返回当前已存在的对比窗口（若存在且未关闭），否则返回 None。"""
    global _diff_instance
    if _diff_instance is not None and _diff_instance.root.winfo_exists():
        return _diff_instance
    _diff_instance = None
    return None


class DiffWindow:
    """对比 GUI：内嵌差异表格，关闭时可选删除临时对比文件。支持基准互换、单实例。"""

    # 表头列名（用于根据基准更新）
    COL_SHEET, COL_KEY, COL_STATUS, COL_A, COL_B = "Sheet", "Key", "状态", "A侧", "B侧"

    def __init__(self, path_a, path_b):
        global _diff_instance
        # 约定：path_a=本地 path_b=线上
        self.path_local = path_a
        self.path_online = path_b
        self.baseline = "local"  # "local" | "remote"，基准为本地时 A=本地 B=线上，基准为线上时 A=线上 B=本地
        self.path_a = path_a
        self.path_b = path_b
        self.path_out = None
        self.diff_rows = []
        self.update_controller = None
        self.update_progress_var = None
        self._visible_diff_rows = []
        self._diff_page = 0
        self.diff_page_var = None
        self.is_temp = False
        self._compare_request_id = 0
        self._compare_running = False
        self.root = tk.Tk()
        self.auto_open_var = tk.BooleanVar(self.root, value=_load_auto_open_compare())
        self.root.title("Excel 二向对比 v%s" % APP_VERSION)
        self.root.minsize(920, 560)
        self.root.geometry("1120x680")
        setup_merge_styles(self.root)
        apply_app_icon(self.root)
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        if self.update_controller:
            self.update_controller.start_background_check()
        _diff_instance = self
        self.root.after(80, self._start_compare_async)

    def _apply_baseline(self):
        """根据当前基准设置 path_a、path_b（供对比逻辑使用）。"""
        if self.baseline == "local":
            self.path_a, self.path_b = self.path_local, self.path_online
        else:
            self.path_a, self.path_b = self.path_online, self.path_local

    @staticmethod
    def _short_path(path, max_len=118):
        if not path or len(path) <= max_len:
            return path
        keep_tail = max(48, max_len - 34)
        return path[:28] + "..." + path[-keep_tail:]

    def _update_baseline_display(self):
        """根据基准更新标题、路径标签、表格列头。"""
        path_a = self._short_path(self.path_a)
        path_b = self._short_path(self.path_b)
        if self.path_a_tip is not None:
            self.path_a_tip.text = self.path_a
        if self.path_b_tip is not None:
            self.path_b_tip.text = self.path_b
        if self.baseline == "local":
            self.title_var.set("本地 (A) vs 线上 (B)")
            self.path_a_var.set("A: %s" % path_a)
            self.path_b_var.set("B: %s" % path_b)
            self.tree.heading(self.COL_A, text="A(本地)")
            self.tree.heading(self.COL_B, text="B(线上)")
        else:
            self.title_var.set("线上 (A) vs 本地 (B)")
            self.path_a_var.set("A: %s" % path_a)
            self.path_b_var.set("B: %s" % path_b)
            self.tree.heading(self.COL_A, text="A(线上)")
            self.tree.heading(self.COL_B, text="B(本地)")

    def _set_compare_busy(self, busy, text=None):
        self._compare_running = busy
        state = tk.DISABLED if busy else tk.NORMAL
        if hasattr(self, "btn_swap"):
            self.btn_swap.config(state=state)
        if hasattr(self, "btn_open_excel"):
            self.btn_open_excel.config(state=state)
        if self.status_var is not None:
            self.status_var.set(text or ("正在计算差异..." if busy else ""))

    def _start_compare_async(self):
        """按当前 path_a/path_b 后台执行对比并刷新列表。"""
        self._apply_baseline()
        self._compare_request_id += 1
        request_id = self._compare_request_id
        path_a = self.path_a
        path_b = self.path_b
        self.diff_rows = []
        self._visible_diff_rows = []
        if self.tree is not None:
            self.tree.delete(*self.tree.get_children())
        self._set_compare_busy(True, "正在计算差异...")

        def worker():
            started = time.time()
            try:
                out_dir = os.path.dirname(os.path.abspath(path_a))
                base_name = os.path.splitext(os.path.basename(path_a))[0]
                from config import COMPARE_SUFFIX
                expected_out = os.path.join(out_dir, base_name + COMPARE_SUFFIX + ".xlsx")
                is_temp = "Temp" in path_a or "Fork" in path_a or "tmp" in path_a.lower()
                path_out, sheet_names, diff_rows = get_compare_data(path_a, path_b, include_same=False)
                if path_out is None:
                    raise RuntimeError("get_compare_data 失败")
                write_compare_excel(path_out, sheet_names, diff_rows, open_file=False)
                elapsed_ms = int((time.time() - started) * 1000)
                result = {
                    "path_out": path_out or expected_out,
                    "sheet_names": sheet_names,
                    "diff_rows": diff_rows,
                    "is_temp": is_temp,
                    "elapsed_ms": elapsed_ms,
                }
            except Exception as e:
                result = {"error": e}
            self.root.after(0, lambda: self._finish_compare(request_id, result))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_compare(self, request_id, result):
        if request_id != self._compare_request_id:
            return
        if result.get("error"):
            err = result["error"]
            self._set_compare_busy(False, "对比失败")
            gui_log("对比失败: " + str(err), self.status_var, is_error=True)
            messagebox.showerror("错误", str(err))
            return
        self.path_out = result["path_out"]
        self.is_temp = bool(result.get("is_temp"))
        self.diff_rows = list(result.get("diff_rows") or [])
        self._refresh_diff_tree()
        _save_diff_filter({k: self.filter_vars[k].get() for k, _ in DIFF_STATUS_OPTIONS})
        _save_auto_open_compare(self.auto_open_var.get())
        elapsed_ms = result.get("elapsed_ms", 0)
        log("对比预览完成: sheets=%d rows=%d elapsed=%dms" % (
            len(result.get("sheet_names") or []), len(self.diff_rows), elapsed_ms
        ))
        self._set_compare_busy(False, "对比完成")
        gui_log("已生成对比文件，共 %d 行差异，耗时 %dms" % (len(self.diff_rows), elapsed_ms), self.status_var)
        if self.auto_open_var.get() and self.path_out and os.path.isfile(self.path_out):
            open_excel_file(self.path_out)

    def _on_swap_baseline(self):
        """互换基准（本地↔线上）并重新对比。"""
        self.baseline = "remote" if self.baseline == "local" else "local"
        self._update_baseline_display()
        self._start_compare_async()

    def activate_and_refresh(self, path_a, path_b):
        """单实例复用：用新路径刷新对比并置前。约定 path_a=本地 path_b=线上。"""
        self.path_local, self.path_online = path_a, path_b
        self._apply_baseline()
        self._update_baseline_display()
        self._start_compare_async()
        self.root.lift()
        self.root.focus_force()

    def _build_ui(self):
        pad = 14
        self.status_var = tk.StringVar(self.root, value="")
        shell = ttk.Frame(self.root, padding=(pad, pad, pad, 0), style="App.TFrame")
        shell.pack(fill=tk.BOTH, expand=True)

        top = ttk.Frame(shell, padding=(12, 12), style="Panel.TFrame")
        top.pack(fill=tk.X)
        title_row = ttk.Frame(top, style="Panel.TFrame")
        title_row.pack(fill=tk.X)
        make_header_icon(title_row, self.root, style_name="PanelIcon.TLabel").pack(side=tk.LEFT, padx=(0, 10))
        title_text = ttk.Frame(title_row, style="Panel.TFrame")
        title_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        headline = ttk.Frame(title_text, style="Panel.TFrame")
        headline.pack(fill=tk.X)
        self.title_var = tk.StringVar(self.root, value="本地 (A) vs 线上 (B)")
        ttk.Label(headline, textvariable=self.title_var, style="PanelTitle.TLabel").pack(side=tk.LEFT)
        make_badge(headline, "v%s" % APP_VERSION, "primary").pack(side=tk.LEFT, padx=(10, 0))
        ttk.Label(title_text, text="快速查看两个 Excel 的新增、删除和修改，并导出对比工作簿。", style="Muted.TLabel").pack(anchor=tk.W, pady=(2, 0))
        self.btn_swap = make_icon_button(title_row, self.root, "互换基准", "swap", command=self._on_swap_baseline, style="Secondary.TButton")
        self.btn_swap.pack(side=tk.LEFT, padx=(16, 10))
        ToolTip(self.btn_swap, "交换 A/B 基准并重新计算差异。")
        (
            update_card,
            self.btn_update,
            self.update_state_var,
            self.update_state_label,
            self.update_state_icon,
            update_progress_row,
            self.update_progress_var,
            self.update_progress_label,
            self.update_progress_bar,
        ) = make_update_card(title_row, self.root)
        update_card.pack(side=tk.RIGHT, fill=tk.Y)
        ToolTip(self.btn_update, "检查 GitHub Release 是否有新版本；exe 运行模式支持自动下载替换。")
        self.update_controller = UpdateButtonController(
            self.root, self.btn_update, status_var=self.status_var, on_quit=self._on_close, compact=True,
        )
        self.update_controller.bind_state_widget(self.update_state_var, self.update_state_label, self.update_state_icon)
        self.path_a_var = tk.StringVar(self.root)
        self.path_b_var = tk.StringVar(self.root)
        self.path_a_label = None
        self.path_b_label = None
        self.path_a_tip = None
        self.path_b_tip = None

        self.update_controller.bind_progress_widgets(
            self.update_progress_bar, self.update_progress_var, self.update_progress_label, update_progress_row,
        )

        path_panel = ttk.Frame(top, style="Panel.TFrame")
        path_panel.pack(fill=tk.X, pady=(10, 0))
        path_row_a = ttk.Frame(path_panel, style="Panel.TFrame")
        path_row_a.pack(fill=tk.X, pady=(0, 4))
        make_badge(path_row_a, "A", "neutral").pack(side=tk.LEFT, padx=(0, 8))
        self.path_a_label = ttk.Label(path_row_a, textvariable=self.path_a_var, style="Panel.TLabel")
        self.path_a_label.pack(side=tk.LEFT, anchor=tk.W, fill=tk.X, expand=True)
        self.path_a_tip = ToolTip(self.path_a_label, self.path_a)
        make_icon_button(path_row_a, self.root, "所在位置", "folder", command=lambda: self._open_side_folder("a"), style="Secondary.TButton").pack(side=tk.RIGHT)
        make_icon_button(path_row_a, self.root, "打开", "open", command=lambda: self._open_side_file("a"), style="Secondary.TButton").pack(side=tk.RIGHT, padx=(6, 6))

        path_row_b = ttk.Frame(path_panel, style="Panel.TFrame")
        path_row_b.pack(fill=tk.X)
        make_badge(path_row_b, "B", "neutral").pack(side=tk.LEFT, padx=(0, 8))
        self.path_b_label = ttk.Label(path_row_b, textvariable=self.path_b_var, style="Panel.TLabel")
        self.path_b_label.pack(side=tk.LEFT, anchor=tk.W, fill=tk.X, expand=True)
        self.path_b_tip = ToolTip(self.path_b_label, self.path_b)
        make_icon_button(path_row_b, self.root, "所在位置", "folder", command=lambda: self._open_side_folder("b"), style="Secondary.TButton").pack(side=tk.RIGHT)
        make_icon_button(path_row_b, self.root, "打开", "open", command=lambda: self._open_side_file("b"), style="Secondary.TButton").pack(side=tk.RIGHT, padx=(6, 6))

        toolbar = ttk.Frame(shell, padding=(12, 10), style="Panel.TFrame")
        toolbar.pack(fill=tk.X, pady=(10, 10))
        ttk.Label(toolbar, text="筛选", style="Section.TLabel").pack(side=tk.LEFT, padx=(0, 10))
        loaded_filter = _load_diff_filter()
        self.filter_vars = {}
        for key, label in DIFF_STATUS_OPTIONS:
            var = tk.BooleanVar(self.root, value=loaded_filter.get(key, True))
            self.filter_vars[key] = var
            cb = ttk.Checkbutton(
                toolbar, text=label, variable=var,
                command=self._on_diff_filter_changed,
            )
            cb.configure(text=label)
            cb.pack(side=tk.LEFT, padx=(0, 12))
        make_color_legend(toolbar, [
            (UI["success_bg"], "新增"),
            (UI["danger_bg"], "删除"),
            (UI["warning_bg"], "修改"),
        ]).pack(side=tk.RIGHT)

        table_frame = ttk.Frame(shell, padding=(12, 12), style="Panel.TFrame")
        table_frame.pack(fill=tk.BOTH, expand=True)
        table_frame.grid_rowconfigure(1, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        ttk.Label(table_frame, text="差异列表", style="Section.TLabel").grid(row=0, column=0, sticky=tk.W, pady=(0, 8))

        tree_frame = ttk.Frame(table_frame, style="Panel.TFrame")
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        cols = (self.COL_SHEET, self.COL_KEY, self.COL_STATUS, self.COL_A, self.COL_B)
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=14)
        for c, w, min_w in [
            (self.COL_SHEET, 150, 110),
            (self.COL_KEY, 120, 90),
            (self.COL_STATUS, 76, 66),
            (self.COL_A, 560, 260),
            (self.COL_B, 560, 260),
        ]:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, minwidth=min_w, stretch=False)
        sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        sbh = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb.set, xscrollcommand=sbh.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        sbh.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.tag_configure("新增行", foreground=UI["success"], background=UI["success_bg"])
        self.tree.tag_configure("删除行", foreground=UI["danger"], background=UI["danger_bg"])
        self.tree.tag_configure("新增列", foreground=UI["success"], background=UI["success_bg"])
        self.tree.tag_configure("删除列", foreground=UI["danger"], background=UI["danger_bg"])
        self.tree.tag_configure("修改", foreground=UI["warning"], background=UI["warning_bg"])
        self.tree.configure(cursor="hand2")
        page_row = ttk.Frame(table_frame, style="Panel.TFrame")
        page_row.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        self.diff_page_var = tk.StringVar(self.root, value="")
        make_icon_button(page_row, self.root, "上一页", "prev", command=lambda: self._change_diff_page(-1), style="Secondary.TButton").pack(side=tk.LEFT, padx=(0, 4))
        make_icon_button(page_row, self.root, "下一页", "next", command=lambda: self._change_diff_page(1), style="Secondary.TButton").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(page_row, textvariable=self.diff_page_var, style="Panel.TLabel").pack(side=tk.LEFT)
        self._update_baseline_display()

        btn_frame = ttk.Frame(self.root, padding=(pad, 10), style="BottomBar.TFrame")
        btn_frame.pack(fill=tk.X)
        ttk.Label(btn_frame, textvariable=self.status_var, style="Panel.TLabel").pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.btn_open_excel = ttk.Button(btn_frame, text="打开 Excel", command=self._open_excel, style="Accent.TButton")
        configure_button_icon(self.root, self.btn_open_excel, "open")
        self.btn_open_excel.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Checkbutton(
            btn_frame, text="对比后自动打开", variable=self.auto_open_var,
            command=lambda: _save_auto_open_compare(self.auto_open_var.get()),
        ).pack(side=tk.LEFT, padx=(0, 12))
        make_icon_button(btn_frame, self.root, "关闭", "cancel", command=self._on_close, style="Secondary.TButton").pack(side=tk.LEFT)

    def _open_side_file(self, side):
        """打开当前 A/B 的原始 Excel 文件。side: 'a' | 'b'"""
        self._apply_baseline()
        path = self.path_a if side == "a" else self.path_b
        if path and os.path.isfile(path):
            open_excel_file(path)
            gui_log("已打开 %s 文件" % ("A" if side == "a" else "B"), self.status_var)
        else:
            messagebox.showwarning("提示", "文件不存在")

    def _open_side_folder(self, side):
        """打开当前 A/B 的所在位置。side: 'a' | 'b'"""
        self._apply_baseline()
        path = self.path_a if side == "a" else self.path_b
        ok = open_containing_folder(path, select_file=True)
        if not ok:
            messagebox.showwarning("提示", "无法打开所在位置")

    def _on_diff_filter_changed(self):
        """筛选勾选变化：持久化并刷新列表。"""
        visible = {k: self.filter_vars[k].get() for k, _ in DIFF_STATUS_OPTIONS}
        _save_diff_filter(visible)
        self._refresh_diff_tree()

    def _refresh_diff_tree(self):
        """按当前筛选勾选刷新差异列表（仅插入选中状态的行）。"""
        for i in self.tree.get_children(""):
            self.tree.delete(i)
        if not hasattr(self, "filter_vars"):
            return
        visible = {k: self.filter_vars[k].get() for k, _ in DIFF_STATUS_OPTIONS}
        self._visible_diff_rows = [r for r in self.diff_rows if visible.get(r[2], True)]
        self._diff_page = 0
        shown = self._refresh_diff_tree_page()
        if self.diff_rows and hasattr(self, "status_var") and self.status_var:
            extra = "，每页显示 %d 条" % MAX_TREEVIEW_ROWS if len(self._visible_diff_rows) > MAX_TREEVIEW_ROWS else ""
            self.status_var.set("已生成对比文件，共 %d 行，当前显示 %d 条%s" % (len(self.diff_rows), shown, extra))

    def _refresh_diff_tree_page(self):
        """只渲染当前页差异，避免大文件下 Treeview 卡顿。"""
        self.tree.delete(*self.tree.get_children())
        total = len(self._visible_diff_rows)
        if total <= 0:
            if self.diff_page_var is not None:
                self.diff_page_var.set("")
            return 0
        page_size = MAX_TREEVIEW_ROWS
        max_page = max((total - 1) // page_size, 0)
        if self._diff_page < 0:
            self._diff_page = 0
        if self._diff_page > max_page:
            self._diff_page = max_page
        start = self._diff_page * page_size
        end = min(start + page_size, total)
        for sheet_name, key, status, str_a, str_b in self._visible_diff_rows[start:end]:
            sa = "" if str_a is None else str(str_a)
            sb_val = "" if str_b is None else str(str_b)
            self.tree.insert("", tk.END, values=(sheet_name, key, status, sa, sb_val), tags=(status,))
        if self.diff_page_var is not None:
            self.diff_page_var.set("第 %d/%d 页，显示 %d-%d / %d" % (
                self._diff_page + 1, max_page + 1, start + 1, end, total))
        return end - start

    def _change_diff_page(self, delta):
        """切换对比列表页。"""
        if not self._visible_diff_rows:
            return
        old_page = self._diff_page
        self._diff_page += delta
        shown = self._refresh_diff_tree_page()
        if self._diff_page != old_page:
            gui_log("对比列表翻页：第 %d 页，显示 %d 条" % (self._diff_page + 1, shown), self.status_var)

    def _on_tree_double_click(self, event):
        """双击列表行：打开详情面板，左右对比显示每个参数完整内容。"""
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        vals = self.tree.item(item, "values")
        if len(vals) < 5:
            return
        sheet_name, key, status = vals[0], vals[1], vals[2]
        row = next(
            (r for r in self.diff_rows if r[0] == sheet_name and r[1] == key and r[2] == status),
            None,
        )
        if not row:
            return
        _, _, _, str_a, str_b = row
        self._show_detail_panel(sheet_name, key, status, str_a, str_b)

    def _show_detail_panel(self, sheet_name, key, status, str_a, str_b):
        """弹出 Toplevel：左右两栏显示 A/B 完整参数，每参数一行。"""
        win = tk.Toplevel(self.root)
        win.title("详情 — %s / %s / %s" % (sheet_name, key, status))
        win.minsize(640, 400)
        win.geometry("900x500")
        pad = 8
        header = ttk.Frame(win, padding=(pad, pad))
        header.pack(fill=tk.X)
        label_a = "A(本地)" if self.baseline == "local" else "A(线上)"
        label_b = "B(线上)" if self.baseline == "local" else "B(本地)"
        ttk.Label(header, text=label_a, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="  |  ", font=("Segoe UI", 10)).pack(side=tk.LEFT)
        ttk.Label(header, text=label_b, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        paned = ttk.PanedWindow(win, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=pad, pady=(0, pad))
        params_a = str_a.split(" | ") if str_a else [""]
        params_b = str_b.split(" | ") if str_b else [""]
        n = max(len(params_a), len(params_b), 1)
        lines_a = "\n".join("%d: %s" % (i + 1, (params_a[i] if i < len(params_a) else "")) for i in range(n))
        lines_b = "\n".join("%d: %s" % (i + 1, (params_b[i] if i < len(params_b) else "")) for i in range(n))
        f_a = ttk.Frame(paned)
        f_b = ttk.Frame(paned)
        paned.add(f_a, weight=1)
        paned.add(f_b, weight=1)
        txt_a = tk.Text(f_a, wrap=tk.WORD, font=("Consolas", 9), padx=6, pady=6)
        txt_b = tk.Text(f_b, wrap=tk.WORD, font=("Consolas", 9), padx=6, pady=6)
        sb_a = ttk.Scrollbar(f_a, orient=tk.VERTICAL, command=txt_a.yview)
        sb_b = ttk.Scrollbar(f_b, orient=tk.VERTICAL, command=txt_b.yview)
        txt_a.insert(tk.END, lines_a)
        txt_b.insert(tk.END, lines_b)
        txt_a.tag_configure("diff", background="#FFF3E0", foreground="#CC6600")
        txt_b.tag_configure("diff", background="#FFF3E0", foreground="#CC6600")
        for i in range(n):
            va = str(params_a[i]).strip() if i < len(params_a) else ""
            vb = str(params_b[i]).strip() if i < len(params_b) else ""
            if va != vb:
                line_start = "%d.0" % (i + 1)
                line_end = "%d.0" % (i + 2)
                try:
                    txt_a.tag_add("diff", line_start, line_end)
                    txt_b.tag_add("diff", line_start, line_end)
                except tk.TclError:
                    pass

        def _sync_a(first, last):
            sb_a.set(first, last)
            sb_b.set(first, last)
            txt_b.yview_moveto(first)

        def _sync_b(first, last):
            sb_b.set(first, last)
            sb_a.set(first, last)
            txt_a.yview_moveto(first)

        txt_a.config(yscrollcommand=_sync_a)
        txt_b.config(yscrollcommand=_sync_b)
        txt_a.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_a.pack(side=tk.RIGHT, fill=tk.Y)
        txt_b.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_b.pack(side=tk.RIGHT, fill=tk.Y)

        def _strip_line_numbers(text):
            """去掉每行开头的 'N: ' 序号，便于粘贴到 Excel。"""
            return "\n".join(re.sub(r"^\d+:\s*", "", line) for line in text.splitlines())

        def _copy_without_numbers(widget):
            try:
                sel = widget.get(tk.SEL_FIRST, tk.SEL_LAST)
                win.clipboard_clear()
                win.clipboard_append(_strip_line_numbers(sel))
            except tk.TclError:
                pass
            return "break"

        txt_a.bind("<Control-c>", lambda e: _copy_without_numbers(txt_a))
        txt_b.bind("<Control-c>", lambda e: _copy_without_numbers(txt_b))

        def _copy_left():
            win.clipboard_clear()
            win.clipboard_append("\n".join(str(params_a[i]) if i < len(params_a) else "" for i in range(n)))
        def _copy_right():
            win.clipboard_clear()
            win.clipboard_append("\n".join(str(params_b[i]) if i < len(params_b) else "" for i in range(n)))
        def _copy_left_as_row():
            win.clipboard_clear()
            win.clipboard_append("\t".join(str(params_a[i]) if i < len(params_a) else "" for i in range(n)))
        def _copy_right_as_row():
            win.clipboard_clear()
            win.clipboard_append("\t".join(str(params_b[i]) if i < len(params_b) else "" for i in range(n)))

        btn_row = ttk.Frame(win, padding=(pad, 4))
        btn_row.pack(fill=tk.X)
        ttk.Button(btn_row, text="复制左侧成列", command=_copy_left).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="复制右侧成列", command=_copy_right).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="复制左侧成行", command=_copy_left_as_row).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="复制右侧成行", command=_copy_right_as_row).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="关闭", command=win.destroy).pack(side=tk.LEFT, padx=8)

    def _open_excel(self):
        if self.path_out and os.path.isfile(self.path_out):
            open_excel_file(self.path_out)
            gui_log("已打开对比文件", self.status_var)
        else:
            messagebox.showwarning("提示", "对比文件不存在")

    def _on_close(self):
        global _diff_instance
        release_compare_lock()
        if _diff_instance is self:
            _diff_instance = None
        if self.is_temp and self.path_out and os.path.isfile(self.path_out):
            try:
                os.remove(self.path_out)
            except Exception:
                pass
        self.root.quit()
        self.root.destroy()
        sys.exit(0)

    def run(self):
        self.root.mainloop()
