# -*- coding: utf-8 -*-
"""
合并窗口：A 行不变 / B 列不变 / C 删除行 / D 删除列 / E 新增 Sheet / F 删除 Sheet / G 冲突行或列。
多选框选中项参与合并逻辑，选项持久化到本地 merge_options.json。
"""

import json
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox

import openpyxl

from config import BACKUP_SUBDIR
from conflict import compute_conflicts_d
from excel_io import (
    get_sheet_names,
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
            ("#CC6600", "橙色=修改"),
            ("#CC0000", "红色=冲突"),
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
        self.tree.tag_configure("del", foreground="#CC6600", background="#FFF3E0")
        self.tree.tag_configure("conflict", foreground="#CC0000", background="#FFEBEE")
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
        wb_base = openpyxl.load_workbook(
            self.path_local if base_side == "local" else self.path_remote,
            data_only=True,
        )
        wb_other = openpyxl.load_workbook(
            self.path_remote if base_side == "local" else self.path_local,
            data_only=True,
        )
        base_sheets = set(get_sheet_names(wb_base))
        other_sheets = get_sheet_names(wb_other)
        total = 0
        sheet_names_common = [n for n in get_sheet_names(wb_base) if n in wb_base.sheetnames and n in wb_other.sheetnames]
        if "A" not in options:
            for sheet_name in sheet_names_common:
                ws_b = wb_base[sheet_name]
                ws_o = wb_other[sheet_name]
                max_col = max(ws_b.max_column or 1, ws_o.max_column or 1)
                rows_b, _ = load_sheet_rows_full(ws_b, max_col)
                rows_o, _ = load_sheet_rows_full(ws_o, max_col)
                base_keys = set(key_str_normalized(r[0]) if r else "" for r in rows_b)
                base_keys.discard("")
                other_ordered = ordered_keys_normalized(rows_o)
                for k in other_ordered:
                    if k not in base_keys:
                        self.tree.insert("", tk.END, values=(sheet_name, k, "（将新增行）"), tags=("new",))
                        total += 1
        if "C" in options:
            for sheet_name in sheet_names_common:
                ws_b = wb_base[sheet_name]
                ws_o = wb_other[sheet_name]
                max_col = max(ws_b.max_column or 1, ws_o.max_column or 1)
                rows_b, _ = load_sheet_rows_full(ws_b, max_col)
                rows_o, _ = load_sheet_rows_full(ws_o, max_col)
                base_keys = set(key_str_normalized(r[0]) if r else "" for r in rows_b)
                base_keys.discard("")
                other_keys = set(key_str_normalized(r[0]) if r else "" for r in rows_o)
                other_keys.discard("")
                for k in base_keys:
                    if k not in other_keys:
                        self.tree.insert("", tk.END, values=(sheet_name, k, "（将删除行）"), tags=("del",))
                        total += 1
        if "B" not in options:
            for sheet_name in sheet_names_common:
                ws_b = wb_base[sheet_name]
                ws_o = wb_other[sheet_name]
                max_col = max(ws_b.max_column or 1, ws_o.max_column or 1)
                header_b_norm = set(header_normalize_for_compare(h) for h in load_sheet_header(ws_b, max_col) if h)
                header_o = load_sheet_header(ws_o, max_col)
                for h in header_o:
                    if h and header_normalize_for_compare(h) not in header_b_norm:
                        self.tree.insert("", tk.END, values=(sheet_name, h, "（将新增列）"), tags=("new",))
                        total += 1
        if "D" in options:
            for sheet_name in sheet_names_common:
                ws_b = wb_base[sheet_name]
                ws_o = wb_other[sheet_name]
                max_col = max(ws_b.max_column or 1, ws_o.max_column or 1)
                header_o_norm = set(header_normalize_for_compare(h) for h in load_sheet_header(ws_o, max_col) if h)
                header_b = load_sheet_header(ws_b, max_col)
                for h in header_b:
                    if h and header_normalize_for_compare(h) not in header_o_norm:
                        self.tree.insert("", tk.END, values=(sheet_name, h, "（将删除列）"), tags=("del",))
                        total += 1
        if "E" in options:
            for name in other_sheets:
                if name not in base_sheets and name in wb_other.sheetnames:
                    self.tree.insert("", tk.END, values=(name, "（将新增 Sheet）", "—"), tags=("new",))
                    total += 1
        if "F" in options:
            for name in base_sheets:
                if name not in set(get_sheet_names(wb_other)):
                    self.tree.insert("", tk.END, values=(name, "（将删除 Sheet）", "—"), tags=("del",))
                    total += 1
        if "G" in options:
            self.conflict_rows, self.conflict_cols, _ = compute_conflicts_d(self.path_local, self.path_remote)
            for c in self.conflict_rows:
                var = tk.StringVar(value="本地")
                idx = len(self.conflict_vars)
                self.conflict_vars.append((var, c, "row"))
                self.tree.insert("", tk.END, values=(c["sheet"], c["key"] + " (行)", "将保留本地"), tags=(str(idx), "conflict"))
                total += 1
            for c in self.conflict_cols:
                var = tk.StringVar(value="本地")
                idx = len(self.conflict_vars)
                self.conflict_vars.append((var, c, "column"))
                self.tree.insert("", tk.END, values=(c["sheet"], c["key"] + " (列)", "将保留本地"), tags=(str(idx), "conflict"))
                total += 1
        wb_base.close()
        wb_other.close()
        self.hint_label.config(text="基准=%s。勾选参与项：%s。下列为将新增/将删除/冲突项，共 %d 条。（切换基准后删除↔新增会互换）" % (
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
