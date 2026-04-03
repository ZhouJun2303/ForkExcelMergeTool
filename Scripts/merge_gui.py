# -*- coding: utf-8 -*-
"""
合并窗口：A 行不变 / B 列不变 / C 删除行 / D 删除列 / E 新增 Sheet / F 删除 Sheet / G 冲突行或列。
多选框选中项参与合并逻辑，选项持久化到本地 merge_options.json。
"""

import json
import os
import re
import sys
import tkinter as tk
from tkinter import ttk, messagebox

import openpyxl

from config import BACKUP_SUBDIR
from conflict import compute_conflicts_d
from excel_io import (
    cell_str,
    get_sheet_names,
    get_column_values,
    header_normalize_for_compare,
    key_str_normalized,
    load_sheet_header,
    load_sheet_rows,
    load_sheet_rows_full,
    ordered_keys,
    ordered_keys_normalized,
    rows_to_dict,
)
from version import __version__ as APP_VERSION
from gui_common import (
    gui_log,
    make_color_legend,
    open_excel_file,
    setup_merge_styles,
)
from git_util import get_git_merge_info, stage_merged_and_cleanup
from log_util import merge_options_path, release_merge_lock
from merge_core import do_merge


# 选项默认值（勾选=参与逻辑）
DEFAULT_OPTIONS = {"A": True, "B": True, "C": False, "D": False, "E": True, "F": False, "G": True}


def _load_merge_options():
    """从本地文件加载合并选项，不存在或异常则返回默认。"""
    try:
        path = merge_options_path()
        if path and os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {k: bool(data.get(k, DEFAULT_OPTIONS.get(k, False))) for k in "ABCDEFG"}
    except Exception:
        pass
    return dict(DEFAULT_OPTIONS)


def _save_merge_options(opts):
    """将合并选项写入本地文件。"""
    try:
        path = merge_options_path()
        if path:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(opts, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _load_auto_open_merged():
    """从 merge_options.json 读取「合并后自动打开」是否勾选，默认 True。"""
    try:
        path = merge_options_path()
        if path and os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return bool(data.get("auto_open_merged", True))
    except Exception:
        pass
    return True


def _save_auto_open_merged(value):
    """将「合并后自动打开」写入本地，与 merge_options.json 合并。"""
    try:
        path = merge_options_path()
        if not path:
            return
        data = {}
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        data["auto_open_merged"] = bool(value)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# 单实例：当前合并窗口引用，关闭时清空
_merge_instance = None


def get_existing_merge_window():
    """返回当前已存在的合并窗口（若存在且未关闭），否则返回 None。"""
    global _merge_instance
    if _merge_instance is not None and _merge_instance.root.winfo_exists():
        return _merge_instance
    _merge_instance = None
    return None


class MergeWindow:
    """
    合并 GUI：7 项多选（A～G）、基准选择、根据选项展示将删除/新增/冲突列表，生成合并结果。
    """

    OPTIONS = [
        ("A", "行不变"),
        ("B", "列不变"),
        ("C", "删除行"),
        ("D", "删除列"),
        ("E", "新增 Sheet"),
        ("F", "删除 Sheet"),
        ("G", "冲突行/列"),
    ]
    BASE_SIDES = [("local", "本地 (Local)"), ("remote", "线上 (Remote)")]

    def __init__(self, path_local, path_base, path_remote, path_merged):
        global _merge_instance
        self.path_local = path_local
        self.path_base = path_base
        self.path_remote = path_remote
        self.path_merged = path_merged
        self._merged_file_path = None
        self._backup_merged_path = None
        self.merge_done = False
        self.base_side_var = None
        self.status_var = None
        self.tree = None
        self.option_vars = {}
        self.conflict_vars = []
        self.conflict_rows = []
        self.conflict_cols = []
        self.root = tk.Tk()
        self.root.title("Excel 多模式合并 v%s" % APP_VERSION)
        _merge_instance = self
        self.root.minsize(820, 760)
        self.root.geometry("1024x820")
        setup_merge_styles(self.root)

        self.local_info, self.remote_info = get_git_merge_info(path_merged)
        self._build_ui()
        self._on_options_or_base_changed()

    def _build_ui(self):
        pad = 12
        self.status_var = tk.StringVar(self.root, value="")
        bottom_bar = ttk.Frame(self.root, padding=(pad, 10))
        bottom_bar.pack(side=tk.BOTTOM, fill=tk.X)
        bottom_bar.pack_propagate(True)
        status_row = ttk.Frame(bottom_bar)
        status_row.pack(fill=tk.X)
        ttk.Label(status_row, textvariable=self.status_var, font=("Segoe UI", 9)).pack(side=tk.LEFT, anchor=tk.W)
        btn_row = ttk.Frame(bottom_bar)
        btn_row.pack(fill=tk.X, pady=(8, 0))
        self.btn_merge = ttk.Button(btn_row, text="生成合并结果", command=self._on_generate_merge, style="Accent.TButton")
        self.btn_merge.pack(side=tk.LEFT, padx=(0, 8))
        self.auto_open_var = tk.BooleanVar(self.root, value=_load_auto_open_merged())
        ttk.Checkbutton(
            btn_row, text="合并后自动打开", variable=self.auto_open_var,
            command=lambda: _save_auto_open_merged(self.auto_open_var.get()),
        ).pack(side=tk.LEFT, padx=(0, 12))
        self.btn_open_merged = ttk.Button(btn_row, text="打开合并结果", command=self._on_open_merged)
        self.btn_open_merged.pack(side=tk.LEFT, padx=8)
        self.btn_open_backup_merged = ttk.Button(btn_row, text="打开备份的合并文件", command=self._on_open_backup_merged)
        self.btn_open_backup_merged.pack(side=tk.LEFT, padx=8)
        self.btn_confirm = ttk.Button(btn_row, text="确认无误并解决冲突", command=self._on_confirm_done)
        self.btn_confirm.pack(side=tk.LEFT, padx=8)
        self.btn_confirm.config(state=tk.DISABLED)
        ttk.Button(btn_row, text="取消", command=self._on_cancel).pack(side=tk.LEFT)

        center = ttk.Frame(self.root)
        center.pack(fill=tk.BOTH, expand=True, padx=pad, pady=(0, pad))
        top = ttk.Frame(center, padding=(0, 0, 0, 6))
        top.pack(fill=tk.X)
        title_row = ttk.Frame(top)
        title_row.pack(fill=tk.X)
        ttk.Label(title_row, text="多选项合并：勾选参与逻辑，选项已保存到本地", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT)
        tk.Label(title_row, text="  版本 v%s" % APP_VERSION, font=("Segoe UI", 9, "bold"), fg="#1877f2", bg="#f0f2f5").pack(side=tk.LEFT)
        path_short = self.path_merged if len(self.path_merged) <= 72 else "…" + self.path_merged[-68:]
        ttk.Label(top, text=path_short, font=("Segoe UI", 8)).pack(anchor=tk.W)

        opts_row = ttk.Frame(top)
        opts_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(opts_row, text="参与项：").pack(side=tk.LEFT, padx=(0, 8))
        loaded = _load_merge_options()
        self.option_vars = {}
        for key, label in self.OPTIONS:
            var = tk.BooleanVar(self.root, value=loaded.get(key, DEFAULT_OPTIONS.get(key, False)))
            self.option_vars[key] = var
            cb = ttk.Checkbutton(opts_row, text="%s %s" % (key, label), variable=var, command=self._on_option_click)
            cb.pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(opts_row, text="  基准：").pack(side=tk.LEFT, padx=(12, 4))
        self.base_side_var = tk.StringVar(self.root, value="local")
        base_cb = ttk.Combobox(opts_row, textvariable=self.base_side_var, state="readonly", width=16)
        base_cb["values"] = [b[1] for b in self.BASE_SIDES]
        base_cb.current(0)
        base_cb.pack(side=tk.LEFT)
        base_cb.bind("<<ComboboxSelected>>", lambda e: self._on_options_or_base_changed())
        hint_opts = ttk.Label(top, text="提示：不勾选A则插入线上新增行，不勾选B则插入线上新增列。勾选C/D则删除本地独有的行/列。勾选G则显示所有冲突项供选择。", font=("Segoe UI", 8), foreground="gray")
        hint_opts.pack(anchor=tk.W, pady=(2, 0))

        info_frame = ttk.LabelFrame(center, text="版本说明", padding=pad)
        info_frame.pack(fill=tk.X, pady=(0, 6))
        row1 = ttk.Frame(info_frame)
        row1.pack(fill=tk.X)
        left_box = ttk.LabelFrame(row1, text="本地 (Local)", padding=8)
        left_box.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self._fill_commit_info(left_box, self.local_info)
        ttk.Button(left_box, text="打开本地 Excel", command=lambda: self._open_local()).pack(anchor=tk.W, pady=(6, 0))
        right_box = ttk.LabelFrame(row1, text="线上 (Remote)", padding=8)
        right_box.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
        self._fill_commit_info(right_box, self.remote_info)
        ttk.Button(right_box, text="打开线上 Excel", command=lambda: self._open_remote()).pack(anchor=tk.W, pady=(6, 0))

        legend_merge = make_color_legend(center, [
            ("#008000", "绿色=新增"),
            ("#FF6600", "橙色=删除冲突"),
            ("#CC0000", "红色=修改冲突"),
            ("#808080", "灰色=将删除"),
        ])
        legend_merge.pack(anchor=tk.W, pady=(0, 6))

        self.content_frame = ttk.Frame(center)
        self.content_frame.pack(fill=tk.BOTH, expand=True)
        self.hint_label = ttk.Label(self.content_frame, text="", font=("Segoe UI", 9))
        self.hint_label.pack(anchor=tk.W)
        table_frame = ttk.Frame(self.content_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        cols = ("Sheet", "Key / 说明", "选择")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=14)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120 if c != "Key / 说明" else 280)
        self.tree.tag_configure("new", foreground="#008000", background="#E8F5E9")
        self.tree.tag_configure("del", foreground="#808080", background="#F5F5F5")
        self.tree.tag_configure("del_conflict", foreground="#FF6600", background="#FFF3E0")
        self.tree.tag_configure("conflict", foreground="#CC0000", background="#FFEBEE")
        self.tree.bind("<Double-1>", self._on_merge_tree_double_click)
        sb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        sel_frame = ttk.Frame(self.content_frame)
        sel_frame.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(sel_frame, text="当前选中项：").pack(side=tk.LEFT)
        ttk.Button(sel_frame, text="取本地", command=lambda: self._set_choice("本地")).pack(side=tk.LEFT, padx=4)
        ttk.Button(sel_frame, text="取线上", command=lambda: self._set_choice("线上")).pack(side=tk.LEFT)

    def _get_options(self):
        """返回当前勾选的选项集合 {"A","B",...}。"""
        return {k for k, v in self.option_vars.items() if v.get()}

    def _get_base_side(self):
        raw = self.base_side_var.get()
        for sid, label in self.BASE_SIDES:
            if label == raw:
                return sid
        return "local"

    def _on_option_click(self):
        opts = {k: self.option_vars[k].get() for k in self.option_vars}
        _save_merge_options(opts)
        self._on_options_or_base_changed()

    def _on_options_or_base_changed(self):
        for i in self.tree.get_children(""):
            self.tree.delete(i)
        self.conflict_vars.clear()
        options = self._get_options()
        base_side = self._get_base_side()
        try:
            self._fill_content_by_options(options, base_side)
        except Exception as e:
            gui_log("加载数据失败: " + str(e), self.status_var, is_error=True)
            self.hint_label.config(text="加载失败: " + str(e) + " 详见日志。")

    def _fill_content_by_options(self, options, base_side):
        """根据勾选项填充表格：A 未选则显示将新增行，B 未选则显示将新增列，C/D/E/F/G 显示删除行/列、Sheet、冲突。基准切换后删除↔新增对称。"""
        path_base = self.path_local if base_side == "local" else self.path_remote
        path_other = self.path_remote if base_side == "local" else self.path_local

        # 一次性加载两个文件的所有数据，避免重复读取
        wb_base = openpyxl.load_workbook(path_base, data_only=True)
        wb_other = openpyxl.load_workbook(path_other, data_only=True)

        base_sheets = set(get_sheet_names(wb_base))
        other_sheets = get_sheet_names(wb_other)
        sheet_names_common = [n for n in get_sheet_names(wb_base) if n in wb_other.sheetnames]

        # 预加载每个 Sheet 的数据（每个文件只读一次）
        sheet_data_base = {}
        sheet_data_other = {}
        # 提前保存 sheetnames，避免关闭 workbook 后访问
        wb_base_sheetnames = list(wb_base.sheetnames)
        wb_other_sheetnames = list(wb_other.sheetnames)
        for sheet_name in sheet_names_common:
            ws_b = wb_base[sheet_name]
            ws_o = wb_other[sheet_name]
            max_col = max(ws_b.max_column or 1, ws_o.max_column or 1)
            rows_b, _ = load_sheet_rows_full(ws_b, max_col, use_cache=True)
            rows_o, _ = load_sheet_rows_full(ws_o, max_col, use_cache=True)
            sheet_data_base[sheet_name] = {
                'rows': rows_b,
                'keys': set(key_str_normalized(r[0]) if r else "" for r in rows_b),
                'keys_ordered': ordered_keys_normalized(rows_b),
                'header': load_sheet_header(ws_b, max_col),
                'header_other': load_sheet_header(ws_o, max_col),
            }
            sheet_data_other[sheet_name] = {
                'rows': rows_o,
                'keys': set(key_str_normalized(r[0]) if r else "" for r in rows_o),
                'keys_ordered': ordered_keys_normalized(rows_o),
                'header': load_sheet_header(ws_o, max_col),
            }

        wb_base.close()
        wb_other.close()

        # 禁用 Treeview 重绘，批量插入后再启用
        self.tree.delete(*self.tree.get_children())
        self.conflict_vars.clear()
        items_to_insert = []  # 暂存待插入项

        total = 0
        if "A" not in options:
            for sheet_name in sheet_names_common:
                data_b = sheet_data_base[sheet_name]
                data_o = sheet_data_other[sheet_name]
                base_keys = data_b['keys'] - {""}
                for k in data_o['keys_ordered']:
                    if k and k not in base_keys:
                        items_to_insert.append((sheet_name, k, "（将新增行）", "new"))
                        total += 1

        if "C" in options:
            for sheet_name in sheet_names_common:
                data_b = sheet_data_base[sheet_name]
                data_o = sheet_data_other[sheet_name]
                base_keys = data_b['keys'] - {""}
                other_keys = data_o['keys'] - {""}
                for k in base_keys:
                    if k not in other_keys:
                        items_to_insert.append((sheet_name, k, "（将删除行）", "del"))
                        total += 1

        if "B" not in options:
            for sheet_name in sheet_names_common:
                data_b = sheet_data_base[sheet_name]
                data_o = sheet_data_other[sheet_name]
                header_b_norm = set(header_normalize_for_compare(h) for h in data_b['header'] if h)
                for h in data_o['header']:
                    if h and header_normalize_for_compare(h) not in header_b_norm:
                        items_to_insert.append((sheet_name, h, "（将新增列）", "new"))
                        total += 1

        if "D" in options:
            for sheet_name in sheet_names_common:
                data_b = sheet_data_base[sheet_name]
                data_o = sheet_data_other[sheet_name]
                header_o_norm = set(header_normalize_for_compare(h) for h in data_o['header'] if h)
                for h in data_b['header']:
                    if h and header_normalize_for_compare(h) not in header_o_norm:
                        items_to_insert.append((sheet_name, h, "（将删除列）", "del"))
                        total += 1

        if "E" in options:
            for name in other_sheets:
                if name not in base_sheets:
                    items_to_insert.append((name, "（将新增 Sheet）", "—", "new"))
                    total += 1

        if "F" in options:
            for name in wb_base_sheetnames:
                if name not in set(get_sheet_names(wb_other)):
                    items_to_insert.append((name, "（将删除 Sheet）", "—", "del"))
                    total += 1

        if "G" in options:
            # 使用三向冲突检测（包含删除冲突识别）
            from conflict import compute_conflicts
            conflicts, _, _ = compute_conflicts(self.path_local, self.path_base, self.path_remote)
            
            # 定义冲突类型的显示配置
            type_display = {
                "add_local": ("（仅本地新增）", "new", False),  # 不需要选择，由A选项控制
                "add_remote": ("（仅线上新增）", "new", False),  # 不需要选择，由A选项控制
                "add_conflict": (" (新增冲突)", "conflict", True),
                "delete_conflict_local": (" (删除冲突：本地删)", "del_conflict", True),
                "delete_conflict_remote": (" (删除冲突：线上删)", "del_conflict", True),
                "modify_conflict": (" (修改冲突)", "conflict", True),
            }
            
            for c in conflicts:
                conflict_type = c.get("type", "modify_conflict")
                suffix, tag, need_choice = type_display.get(conflict_type, (" (冲突)", "conflict", True))
                
                if need_choice:
                    # 需要用户选择的冲突
                    var = tk.StringVar(value="本地")
                    idx = len(self.conflict_vars)
                    self.conflict_vars.append((var, c, "row"))
                    
                    # 根据冲突类型显示不同的描述和默认值
                    if conflict_type == "delete_conflict_local":
                        choice_text = "将保留线上（本地已删）"
                        var.set("线上")  # 默认保留线上
                    elif conflict_type == "delete_conflict_remote":
                        choice_text = "将保留本地（线上已删）"
                        var.set("本地")  # 默认保留本地
                    else:
                        choice_text = "将保留本地"
                        var.set("本地")
                    
                    items_to_insert.append((c["sheet"], c["key"] + suffix, choice_text, tag))
                    total += 1
                else:
                    # 仅信息展示的项（单方新增，由A选项控制是否插入）
                    if conflict_type == "add_local":
                        choice_text = "（信息）本地新增"
                    else:
                        choice_text = "（信息）线上新增"
                    items_to_insert.append((c["sheet"], c["key"] + suffix, choice_text, tag))
                    total += 1

        # 批量插入（禁用重绘后一次插入）
        for sheet_name, key, choice, tag in items_to_insert:
            self.tree.insert("", tk.END, values=(sheet_name, key, choice), tags=(tag,))

        self.hint_label.config(text="基准=%s。勾选参与项：%s。下列为将新增/将删除/冲突项，共 %d 条。（切换基准后删除↔新增会互换）\n提示：单方新增行由【A 行不变】选项控制（不勾选A则插入新增行），冲突项可在下方选择保留哪方。" % (
            "本地" if base_side == "local" else "线上", ", ".join(sorted(options)) or "无", total))
        gui_log("已加载选项数据，共 %d 条" % total, self.status_var)

    def _fill_commit_info(self, parent, info):
        parts = []
        if info and isinstance(info, dict):
            if info.get("short_hash"):
                parts.append("Hash: %s" % info["short_hash"])
            if info.get("author"):
                parts.append("提交人: %s" % info["author"])
            if info.get("date"):
                parts.append(info["date"][:19] if len(info["date"]) >= 19 else info["date"])
            if info.get("message"):
                msg = info["message"]
                parts.append("事件: %s" % (msg[:50] + "…" if len(msg) > 50 else msg))
        if not parts:
            ttk.Label(parent, text="(无法获取 Git 信息)", foreground="gray", font=("", 8)).pack(anchor=tk.W)
        else:
            for p in parts:
                ttk.Label(parent, text=p, font=("", 8)).pack(anchor=tk.W)

    def _open_local(self):
        if open_excel_file(self.path_local):
            gui_log("已打开本地 Excel", self.status_var)
        else:
            messagebox.showwarning("提示", "文件不存在或无法打开")

    def _open_remote(self):
        if open_excel_file(self.path_remote):
            gui_log("已打开线上 Excel", self.status_var)
        else:
            messagebox.showwarning("提示", "文件不存在或无法打开")

    def _on_open_merged(self):
        if not self.merge_done:
            messagebox.showwarning("提示", "请先点击「生成合并结果」")
            return
        path = self._merged_file_path or os.path.normpath(os.path.abspath(self.path_merged))
        if not os.path.isfile(path):
            messagebox.showwarning("提示", "合并文件不存在：%s" % path)
            return
        if open_excel_file(path):
            gui_log("已打开合并结果：%s" % path, self.status_var)
        else:
            messagebox.showwarning("提示", "无法打开合并文件")

    def _on_open_backup_merged(self):
        if not self.merge_done or not self._backup_merged_path:
            messagebox.showwarning("提示", "请先点击「生成合并结果」")
            return
        path = os.path.normpath(os.path.abspath(self._backup_merged_path))
        if not os.path.isfile(path):
            messagebox.showwarning("提示", "备份的合并文件不存在：%s" % path)
            return
        if open_excel_file(path):
            gui_log("已打开备份的合并文件：%s" % path, self.status_var)
        else:
            messagebox.showwarning("提示", "无法打开备份文件")

    def _on_generate_merge(self):
        path_out = os.path.normpath(os.path.abspath(self.path_merged))
        if os.path.isfile(path_out):
            try:
                with open(path_out, "ab") as _:
                    pass
            except PermissionError:
                messagebox.showerror(
                    "无法写入合并结果",
                    "合并结果文件正在被占用（如 Excel 已打开），请先关闭后再点击「生成合并结果」。"
                )
                return
        options = self._get_options()
        gui_log("生成合并结果，勾选项: %s（未勾选 C 则不删除行）" % ", ".join(sorted(options)) or "无", self.status_var)
        base_side = self._get_base_side()
        d_choices = []
        if "G" in options:
            for var, obj, kind in self.conflict_vars:
                choice = "local" if var.get() == "本地" else "remote"
                d_choices.append({
                    "sheet": obj["sheet"],
                    "key": obj["key"],
                    "choice": choice,
                    "kind": "row" if kind == "row" else "column",
                })
        try:
            code = do_merge(
                self.path_local, self.path_base, self.path_remote, self.path_merged,
                base_side=base_side, d_choices=d_choices if d_choices else None,
                options=options,
            )
            if code != 0:
                raise RuntimeError("合并返回码 %d" % code)
            self.merge_done = True
            self._merged_file_path = os.path.normpath(os.path.abspath(self.path_merged))
            merged_dir = os.path.dirname(self._merged_file_path)
            base_name = os.path.splitext(os.path.basename(self.path_merged))[0]
            self._backup_merged_path = os.path.join(merged_dir, BACKUP_SUBDIR, base_name + "_merged.xlsx")
            gui_log("合并结果已生成：%s" % self._merged_file_path, self.status_var)
            _save_auto_open_merged(self.auto_open_var.get())
            self.btn_confirm.config(state=tk.NORMAL)
            if self.auto_open_var.get() and os.path.isfile(self._merged_file_path):
                open_excel_file(self._merged_file_path)
        except PermissionError as e:
            gui_log("合并失败（文件可能被占用）: " + str(e), self.status_var, is_error=True)
            messagebox.showerror(
                "无法写入合并结果",
                "合并结果文件可能正在被 Excel 打开，请先关闭该文件后再点击「生成合并结果」。"
            )
        except Exception as e:
            import traceback
            gui_log("合并失败: " + str(e), self.status_var, is_error=True)
            messagebox.showerror("错误", str(e) + "\n" + traceback.format_exc())

    def activate_and_refresh(self, path_local, path_base, path_remote, path_merged):
        """单实例复用：用新路径刷新列表并置前。"""
        self.path_local = path_local
        self.path_base = path_base
        self.path_remote = path_remote
        self.path_merged = path_merged
        self._on_options_or_base_changed()
        self.root.lift()
        self.root.focus_force()

    def _on_confirm_done(self):
        global _merge_instance
        if not self.merge_done:
            messagebox.showwarning("提示", "请先点击「生成合并结果」")
            return
        def _log_cb(msg, is_err=False):
            gui_log(msg, self.status_var, is_error=is_err)
        stage_merged_and_cleanup(
            self.path_merged, self.path_local, self.path_base, self.path_remote, log_callback=_log_cb,
        )
        messagebox.showinfo("完成", "冲突已解决：已 git add，已清理临时文件。Fork 将使用合并后的文件。")
        release_merge_lock()
        if _merge_instance is self:
            _merge_instance = None
        self.root.quit()
        self.root.destroy()
        sys.exit(0)

    def _set_choice(self, choice):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        tags = self.tree.item(item, "tags")
        if not tags:
            return
        try:
            idx = int(tags[0])
        except ValueError:
            return
        if "G" not in self._get_options() or idx < 0 or idx >= len(self.conflict_vars):
            return
        var, obj, kind = self.conflict_vars[idx]
        var.set(choice)
        display = "将保留本地" if choice == "本地" else "将保留线上"
        vals = list(self.tree.item(item, "values"))
        vals[2] = display
        self.tree.item(item, values=vals)

    def _on_merge_tree_double_click(self, event):
        """双击列表行：打开详情面板，显示 BASE | LOCAL | REMOTE 三列对比。"""
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        vals = self.tree.item(item, "values")
        tags = self.tree.item(item, "tags") or ()
        if len(vals) < 3:
            return
        sheet_or_name, key_or_col, choice = vals[0], vals[1], vals[2]
        
        # 处理冲突项（带索引tag）
        if tags and tags[0] not in ("new", "del", "del_conflict", "conflict"):
            try:
                idx = int(tags[0])
                if 0 <= idx < len(self.conflict_vars):
                    var, c, kind = self.conflict_vars[idx]
                    conflict_type = c.get("type", "modify_conflict")
                    
                    # 提取三方数据
                    base_vals = [cell_str(x) for x in (c.get("base_row") or [])]
                    local_vals = [cell_str(x) for x in (c.get("local_row") or [])]
                    remote_vals = [cell_str(x) for x in (c.get("remote_row") or [])]
                    
                    # 根据冲突类型设置标题
                    type_names = {
                        "add_conflict": "新增冲突",
                        "delete_conflict_local": "删除冲突（本地删除）",
                        "delete_conflict_remote": "删除冲突（线上删除）",
                        "modify_conflict": "修改冲突",
                    }
                    type_name = type_names.get(conflict_type, "冲突")
                    title = "%s — %s / %s" % (type_name, c["sheet"], c["key"])
                    
                    self._show_merge_detail_panel_three(title, base_vals, local_vals, remote_vals, conflict_type)
                    return
            except (ValueError, IndexError):
                pass
        
        # 处理非冲突项（新增/删除）
        base_side = self._get_base_side()
        if choice == "（将新增行）":
            path_base = self.path_local if base_side == "local" else self.path_remote
            path_other = self.path_remote if base_side == "local" else self.path_local
            left_vals, right_vals = self._load_row_from_workbooks(sheet_or_name, key_or_col, path_base, path_other)
            self._show_merge_detail_panel("将新增行 — %s / %s" % (sheet_or_name, key_or_col), "基准(无)", "另一方", left_vals, right_vals)
        elif choice == "（将删除行）":
            path_base = self.path_local if base_side == "local" else self.path_remote
            path_other = self.path_remote if base_side == "local" else self.path_local
            left_vals, right_vals = self._load_row_from_workbooks(sheet_or_name, key_or_col, path_base, path_other)
            self._show_merge_detail_panel("将删除行 — %s / %s" % (sheet_or_name, key_or_col), "基准", "另一方(无)", left_vals, right_vals)
        elif choice == "（将新增列）":
            left_vals, right_vals = self._load_col_from_workbooks(
                sheet_or_name, key_or_col, base_side, has_in_other=True
            )
            self._show_merge_detail_panel("将新增列 — %s / %s" % (sheet_or_name, key_or_col), "基准(无)", "另一方", left_vals, right_vals)
        elif choice == "（将删除列）":
            left_vals, right_vals = self._load_col_from_workbooks(
                sheet_or_name, key_or_col, base_side, has_in_other=False
            )
            self._show_merge_detail_panel("将删除列 — %s / %s" % (sheet_or_name, key_or_col), "基准", "另一方(无)", left_vals, right_vals)
        elif "将新增 Sheet" in choice or "将删除 Sheet" in choice:
            messagebox.showinfo("详情", "Sheet: %s\n%s" % (sheet_or_name, choice))

    def _load_row_from_workbooks(self, sheet_name, key, path_base, path_other):
        """返回 (base_row_values, other_row_values)，缺失一方为空列表。"""
        left, right = [], []
        try:
            wb_b = openpyxl.load_workbook(path_base, data_only=True)
            wb_o = openpyxl.load_workbook(path_other, data_only=True)
            if sheet_name in wb_b.sheetnames and sheet_name in wb_o.sheetnames:
                ws_b = wb_b[sheet_name]
                ws_o = wb_o[sheet_name]
                max_col = max(ws_b.max_column or 1, ws_o.max_column or 1)
                rows_b, _ = load_sheet_rows_full(ws_b, max_col, use_cache=True)
                rows_o, _ = load_sheet_rows_full(ws_o, max_col, use_cache=True)
                dict_b = rows_to_dict(rows_b)
                dict_o = rows_to_dict(rows_o)
                key_norm = key_str_normalized(key)
                for d, out in [(dict_b, left), (dict_o, right)]:
                    for k, row in d.items():
                        if key_str_normalized(k) == key_norm or k == key:
                            out.extend([cell_str(c) for c in row])
                            break
            wb_b.close()
            wb_o.close()
        except Exception:
            pass
        return (left, right)

    def _load_col_from_workbooks(self, sheet_name, col_header, base_side, has_in_other):
        """返回 (基准列值, 另一方列值)，即 (left_vals, right_vals)。基准=本地时 left 来自 local。"""
        left, right = [], []
        try:
            wb_l = openpyxl.load_workbook(self.path_local, data_only=True)
            wb_r = openpyxl.load_workbook(self.path_remote, data_only=True)
            for wb, path_label in [(wb_l, "local"), (wb_r, "remote")]:
                if sheet_name not in wb.sheetnames:
                    continue
                ws = wb[sheet_name]
                max_col = ws.max_column or 1
                max_row = max(ws.max_row or 1, 1)
                headers = load_sheet_header(ws, max_col)
                col_idx = None
                for i, h in enumerate(headers):
                    if h and header_normalize_for_compare(h) == header_normalize_for_compare(col_header):
                        col_idx = i + 1
                        break
                if col_idx is None:
                    continue
                vals = get_column_values(ws, col_idx, max_row)
                out = [cell_str(v) for v in vals]
                if path_label == "local":
                    left = out
                else:
                    right = out
            wb_l.close()
            wb_r.close()
            # 基准为线上时，left 应为 remote、right 为 local
            if base_side == "remote":
                left, right = right, left
        except Exception:
            pass
        return (left, right)

    def _show_merge_detail_panel_three(self, title, base_vals, local_vals, remote_vals, conflict_type):
        """弹出 Toplevel：三栏显示 BASE | LOCAL | REMOTE，高亮差异。"""
        win = tk.Toplevel(self.root)
        win.title(title)
        win.minsize(900, 500)
        win.geometry("1200x600")
        pad = 8
        
        # 顶部说明
        header = ttk.Frame(win, padding=(pad, pad))
        header.pack(fill=tk.X)
        
        # 根据冲突类型显示不同的说明
        if conflict_type == "delete_conflict_local":
            hint = "本地删除了此行，线上保留/修改了此行"
        elif conflict_type == "delete_conflict_remote":
            hint = "线上删除了此行，本地保留/修改了此行"
        elif conflict_type == "add_conflict":
            hint = "BASE中不存在，双方都新增了此行但内容不同"
        else:
            hint = "BASE中存在，双方都修改了此行但内容不同"
        
        ttk.Label(header, text=hint, font=("Segoe UI", 9), foreground="#666").pack(anchor=tk.W)
        
        # 三栏布局
        cols_frame = ttk.Frame(win)
        cols_frame.pack(fill=tk.BOTH, expand=True, padx=pad, pady=(0, pad))
        
        # 创建三个列
        frames = []
        texts = []
        labels = [("BASE", base_vals), ("LOCAL", local_vals), ("REMOTE", remote_vals)]
        
        for i, (label, vals) in enumerate(labels):
            col = ttk.Frame(cols_frame)
            col.grid(row=0, column=i, sticky="nsew", padx=2)
            cols_frame.columnconfigure(i, weight=1)
            
            # 列标题
            lbl_frame = ttk.Frame(col)
            lbl_frame.pack(fill=tk.X)
            ttk.Label(lbl_frame, text=label, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=4)
            
            # 文本框
            txt = tk.Text(col, wrap=tk.WORD, font=("Consolas", 9), padx=6, pady=6)
            sb = ttk.Scrollbar(col, orient=tk.VERTICAL, command=txt.yview)
            
            # 填充内容
            if vals:
                lines = "\n".join("%d: %s" % (j + 1, v) for j, v in enumerate(vals))
            else:
                lines = "(不存在)"
            txt.insert(tk.END, lines)
            txt.config(state=tk.DISABLED)
            
            txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            sb.pack(side=tk.RIGHT, fill=tk.Y)
            txt.config(yscrollcommand=sb.set)
            
            frames.append(col)
            texts.append(txt)
        
        cols_frame.rowconfigure(0, weight=1)
        
        # 高亮差异
        n = max(len(base_vals), len(local_vals), len(remote_vals))
        for txt in texts:
            txt.tag_configure("diff", background="#FFF3E0", foreground="#CC6600")
            txt.tag_configure("missing", background="#FFEBEE", foreground="#999")
        
        for i in range(n):
            vb = str(base_vals[i]).strip() if i < len(base_vals) else ""
            vl = str(local_vals[i]).strip() if i < len(local_vals) else ""
            vr = str(remote_vals[i]).strip() if i < len(remote_vals) else ""
            
            # 标记差异
            if vb != vl or vb != vr or vl != vr:
                line_start = "%d.0" % (i + 1)
                line_end = "%d.0" % (i + 2)
                for j, (txt, val) in enumerate([(texts[0], vb), (texts[1], vl), (texts[2], vr)]):
                    try:
                        txt.config(state=tk.NORMAL)
                        if not val:
                            txt.tag_add("missing", line_start, line_end)
                        else:
                            txt.tag_add("diff", line_start, line_end)
                        txt.config(state=tk.DISABLED)
                    except tk.TclError:
                        pass

    def _show_merge_detail_panel(self, title, label_left, label_right, left_vals, right_vals):
        """弹出 Toplevel：左右两栏显示完整参数，每参数一行，双栏滚动同步。"""
        win = tk.Toplevel(self.root)
        win.title(title)
        win.minsize(640, 400)
        win.geometry("900x500")
        pad = 8
        header = ttk.Frame(win, padding=(pad, pad))
        header.pack(fill=tk.X)
        ttk.Label(header, text=label_left, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="  |  ", font=("Segoe UI", 10)).pack(side=tk.LEFT)
        ttk.Label(header, text=label_right, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        paned = ttk.PanedWindow(win, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=pad, pady=(0, pad))
        n = max(len(left_vals), len(right_vals), 1)
        lines_left = "\n".join("%d: %s" % (i + 1, (left_vals[i] if i < len(left_vals) else "")) for i in range(n))
        lines_right = "\n".join("%d: %s" % (i + 1, (right_vals[i] if i < len(right_vals) else "")) for i in range(n))
        f_a = ttk.Frame(paned)
        f_b = ttk.Frame(paned)
        paned.add(f_a, weight=1)
        paned.add(f_b, weight=1)
        txt_a = tk.Text(f_a, wrap=tk.WORD, font=("Consolas", 9), padx=6, pady=6)
        txt_b = tk.Text(f_b, wrap=tk.WORD, font=("Consolas", 9), padx=6, pady=6)
        sb_a = ttk.Scrollbar(f_a, orient=tk.VERTICAL, command=txt_a.yview)
        sb_b = ttk.Scrollbar(f_b, orient=tk.VERTICAL, command=txt_b.yview)
        txt_a.insert(tk.END, lines_left)
        txt_b.insert(tk.END, lines_right)
        txt_a.tag_configure("diff", background="#FFF3E0", foreground="#CC6600")
        txt_b.tag_configure("diff", background="#FFF3E0", foreground="#CC6600")
        for i in range(n):
            vl = str(left_vals[i]).strip() if i < len(left_vals) else ""
            vr = str(right_vals[i]).strip() if i < len(right_vals) else ""
            if vl != vr:
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
            win.clipboard_append("\n".join(str(left_vals[i]) if i < len(left_vals) else "" for i in range(n)))
        def _copy_right():
            win.clipboard_clear()
            win.clipboard_append("\n".join(str(right_vals[i]) if i < len(right_vals) else "" for i in range(n)))
        def _copy_left_as_row():
            win.clipboard_clear()
            win.clipboard_append("\t".join(str(left_vals[i]) if i < len(left_vals) else "" for i in range(n)))
        def _copy_right_as_row():
            win.clipboard_clear()
            win.clipboard_append("\t".join(str(right_vals[i]) if i < len(right_vals) else "" for i in range(n)))

        btn_row = ttk.Frame(win, padding=(pad, 4))
        btn_row.pack(fill=tk.X)
        ttk.Button(btn_row, text="复制左侧成列", command=_copy_left).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="复制右侧成列", command=_copy_right).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="复制左侧成行", command=_copy_left_as_row).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="复制右侧成行", command=_copy_right_as_row).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="关闭", command=win.destroy).pack(side=tk.LEFT, padx=8)

    def _on_cancel(self):
        global _merge_instance
        release_merge_lock()
        if _merge_instance is self:
            _merge_instance = None
        self.root.quit()
        self.root.destroy()
        sys.exit(1)

    def run(self):
        self.root.mainloop()
