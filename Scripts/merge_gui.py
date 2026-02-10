# -*- coding: utf-8 -*-
"""
合并窗口：三向合并的 GUI。只做一件事——展示冲突列表、本地/线上版本信息、生成合并结果并确认。
合并策略：以本地为底，复制本地到 MERGED，再在副本上只修改“取线上”的冲突行并追加仅线上有的行。
"""

import os
import shutil
import sys
import tkinter as tk
from tkinter import ttk, messagebox

import openpyxl
from openpyxl.styles import PatternFill

from config import BACKUP_SUBDIR
from conflict import compute_conflicts
from gui_common import (
    gui_log,
    make_color_legend,
    open_excel_file,
    setup_merge_styles,
)
from git_util import get_git_merge_info, stage_merged_and_cleanup


class MergeWindow:
    """
    合并 GUI：冲突选择、打开本地/线上/合并结果、Git 信息、生成合并结果、确认并解决冲突。
    """

    def __init__(self, path_local, path_base, path_remote, path_merged):
        self.path_local = path_local
        self.path_base = path_base
        self.path_remote = path_remote
        self.path_merged = path_merged
        # 合并完成后实际写入的规范路径，用于“打开合并结果”时保证打开的是刚保存的文件（避免 path_merged 与 path_local 同路径时误开缓存）
        self._merged_file_path = None
        self.conflict_vars = []
        self.merge_done = False
        self.root = tk.Tk()
        self.root.title("Excel 三向合并")
        self.root.minsize(800, 680)
        self.root.geometry("1000x720")
        setup_merge_styles(self.root)

        try:
            gui_log("开始计算冲突...", None)
            self.conflicts, self.sheet_data, self.sheet_names = compute_conflicts(
                path_local, path_base, path_remote
            )
        except Exception as e:
            import traceback
            msg = str(e) + "\n" + traceback.format_exc()
            gui_log(msg, None, is_error=True)
            messagebox.showerror("错误", "加载失败: " + str(e))
            self.root.destroy()
            sys.exit(2)

        self.local_info, self.remote_info = get_git_merge_info(path_merged)
        self._build_ui()

    def _build_ui(self):
        pad = 12
        self.status_var = tk.StringVar(value="")
        bottom_bar = ttk.Frame(self.root, padding=(pad, 8))
        bottom_bar.pack(side=tk.BOTTOM, fill=tk.X)
        status_row = ttk.Frame(bottom_bar)
        status_row.pack(fill=tk.X)
        ttk.Label(status_row, textvariable=self.status_var, font=("Segoe UI", 9)).pack(anchor=tk.W)
        btn_row = ttk.Frame(bottom_bar)
        btn_row.pack(fill=tk.X, pady=(6, 0))
        self.btn_merge = ttk.Button(btn_row, text="生成合并结果", command=self._on_generate_merge, style="Accent.TButton")
        self.btn_merge.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_open_merged = ttk.Button(btn_row, text="打开合并结果", command=self._on_open_merged)
        self.btn_open_merged.pack(side=tk.LEFT, padx=8)
        self.btn_confirm = ttk.Button(btn_row, text="确认无误并解决冲突", command=self._on_confirm_done)
        self.btn_confirm.pack(side=tk.LEFT, padx=8)
        self.btn_confirm.config(state=tk.DISABLED)
        ttk.Button(btn_row, text="取消", command=self._on_cancel).pack(side=tk.LEFT)
        gui_log("已加载 %d 项差异（含仅本地有/仅线上有），请选择后点击「生成合并结果」" % len(self.conflicts), self.status_var)

        center = ttk.Frame(self.root)
        center.pack(fill=tk.BOTH, expand=True, padx=pad, pady=(0, pad))
        top = ttk.Frame(center, padding=(0, 0, 0, 6))
        top.pack(fill=tk.X)
        ttk.Label(top, text="本地 (左) + 线上 (右) → 合并结果", font=("Segoe UI", 11, "bold")).pack(anchor=tk.W)
        path_short = self.path_merged if len(self.path_merged) <= 72 else "…" + self.path_merged[-68:]
        ttk.Label(top, text=path_short, font=("Segoe UI", 8)).pack(anchor=tk.W)
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
            ("#CCFFCC", "绿色=新增"),
            ("#FFFF99", "黄色=修改"),
            ("#FFCCCC", "红色=冲突（需选择）"),
            (None, "无=无变化"),
        ])
        legend_merge.pack(anchor=tk.W, pady=(0, 6))
        n_need_choice = sum(1 for c in self.conflicts if not c.get("_only_local") and not c.get("_only_remote"))
        hint = "下表共 %d 项差异（其中 %d 项需选择取本地/取线上，其余为仅一方有将自动保留）" % (len(self.conflicts), n_need_choice)
        ttk.Label(center, text=hint, font=("Segoe UI", 9)).pack(anchor=tk.W, pady=(0, 2))
        ttk.Label(center, text="说明：首列相同 key 出现多行时只按一行参与合并；若需保留多行请用唯一 key。", font=("Segoe UI", 8)).pack(anchor=tk.W)
        if len(self.conflicts) == 0:
            ttk.Label(center, text="若应有冲突却显示 0：请确认三个文件首列为行关键列、且本地与线上内容确有不同；详见同目录下 MergeExcelFork.log。", font=("Segoe UI", 8)).pack(anchor=tk.W)
        ttk.Label(center, text="", font=("Segoe UI", 1)).pack(anchor=tk.W, pady=(0, 2))
        paned = ttk.PanedWindow(center, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)
        table_frame = ttk.Frame(paned)
        paned.add(table_frame, weight=2)
        cols = ("Sheet", "Key", "选择")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=10)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=100 if c != "Key" else 220)
        sb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        for i, c in enumerate(self.conflicts):
            if c.get("_only_local"):
                choice_label = "将保留本地"
                var = tk.StringVar(value="本地")
            elif c.get("_only_remote"):
                choice_label = "将保留线上"
                var = tk.StringVar(value="线上")
            else:
                choice_label = "本地"
                var = tk.StringVar(value="本地")
            self.conflict_vars.append(var)
            self.tree.insert("", tk.END, values=(c["sheet"], c["key"], choice_label), tags=(str(i),))
        self.tree.tag_configure("conflict", background="#fff3cd")
        sel_frame = ttk.Frame(paned)
        paned.add(sel_frame, weight=0)
        ttk.Label(sel_frame, text="当前选中项：").pack(side=tk.LEFT)
        btn_frame = ttk.Frame(sel_frame)
        btn_frame.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btn_frame, text="取本地", command=lambda: self._set_choice("本地")).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="取线上", command=lambda: self._set_choice("线上")).pack(side=tk.LEFT, padx=2)

    def _fill_commit_info(self, parent, info):
        parts = []
        if info and isinstance(info, dict):
            if info.get("short_hash"):
                parts.append("Hash: %s" % info["short_hash"])
            if info.get("author"):
                parts.append("提交人: %s" % info["author"])
            if info.get("email"):
                parts.append("(%s)" % info["email"])
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
        """打开的是合并结果文件（刚保存的 path_merged），使用保存时记录的规范路径，避免开错或打开缓存。"""
        if not self.merge_done:
            messagebox.showwarning("提示", "请先点击「生成合并结果」")
            return
        # 优先用保存时记录的规范路径，确保打开的就是刚写入的合并文件
        path = self._merged_file_path if self._merged_file_path else os.path.normpath(os.path.abspath(self.path_merged))
        if not os.path.isfile(path):
            messagebox.showwarning("提示", "合并文件不存在：%s" % path)
            return
        if open_excel_file(path):
            gui_log("已打开合并结果：%s" % path, self.status_var)
        else:
            messagebox.showwarning("提示", "无法打开合并文件")

    def _on_generate_merge(self):
        for i, c in enumerate(self.conflicts):
            var = self.conflict_vars[i] if i < len(self.conflict_vars) else None
            c["_choice"] = "local" if (var and var.get() == "本地") else "remote"
        try:
            self._do_merge_with_choices()
            self.merge_done = True
            # 使用保存时记录的路径，与“打开合并结果”按钮一致
            path = self._merged_file_path or os.path.normpath(os.path.abspath(self.path_merged))
            gui_log("合并结果已生成：%s" % path, self.status_var)
            messagebox.showinfo("完成", "合并结果已保存。\n将自动打开合并结果文件供您确认，确认无误后点击「确认无误并解决冲突」完成。")
            self.btn_merge.config(state=tk.DISABLED)
            self.btn_confirm.config(state=tk.NORMAL)
            if path and os.path.isfile(path):
                open_excel_file(path)
                gui_log("已打开合并结果文件", self.status_var)
        except Exception as e:
            import traceback
            gui_log("合并失败: " + str(e), self.status_var, is_error=True)
            messagebox.showerror("错误", str(e))

    def _on_confirm_done(self):
        if not self.merge_done:
            messagebox.showwarning("提示", "请先点击「生成合并结果」")
            return
        def _log_cb(msg, is_err=False):
            gui_log(msg, self.status_var, is_error=is_err)
        stage_merged_and_cleanup(
            self.path_merged, self.path_local, self.path_base, self.path_remote, log_callback=_log_cb,
        )
        messagebox.showinfo("完成", "冲突已解决：已 git add，已清理临时文件。Fork 将使用合并后的文件。")
        self.root.quit()
        self.root.destroy()
        sys.exit(0)

    def _set_choice(self, choice):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        idx = int(self.tree.item(item, "tags")[0])
        if idx < 0 or idx >= len(self.conflicts):
            return
        c = self.conflicts[idx]
        if c.get("_only_local"):
            choice = "本地"
        elif c.get("_only_remote"):
            choice = "线上"
        display = "将保留本地" if c.get("_only_local") else ("将保留线上" if c.get("_only_remote") else choice)
        vals = list(self.tree.item(item, "values"))
        vals[2] = display
        self.tree.item(item, values=vals)
        if idx < len(self.conflict_vars):
            self.conflict_vars[idx].set(choice)

    def _on_cancel(self):
        self.root.quit()
        self.root.destroy()
        sys.exit(1)

    def _do_merge_with_choices(self):
        yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        merged_dir = os.path.dirname(os.path.abspath(self.path_merged))
        base_name = os.path.splitext(os.path.basename(self.path_merged))[0]
        backup_dir = os.path.join(merged_dir, BACKUP_SUBDIR)
        choice_map = {}
        for c in self.conflicts:
            choice_map[(c["key"], c["sheet"])] = c.get("_choice", "local")
        os.makedirs(os.path.dirname(self.path_merged) or ".", exist_ok=True)
        shutil.copy2(self.path_local, self.path_merged)
        wb_merged = openpyxl.load_workbook(self.path_merged, data_only=False)
        for sheet_name in self.sheet_names:
            sd = self.sheet_data[sheet_name]
            if sheet_name not in wb_merged.sheetnames:
                continue
            ws = wb_merged[sheet_name]
            key_to_row_l = sd.get("key_to_row_l") or {}
            remote_rows = sd.get("remote_rows") or {}
            local_rows = sd.get("local_rows") or {}
            remote_ordered = sd.get("remote_ordered") or []
            max_col = sd.get("max_col", 1)
            local_set = set(local_rows)
            remote_set = set(remote_rows)
            only_remote_keys = [k for k in remote_ordered if k in (remote_set - local_set)]
            for c in self.conflicts:
                if c.get("_only_local") or c.get("_only_remote"):
                    continue
                s, key = c.get("sheet"), c.get("key")
                if s != sheet_name:
                    continue
                if choice_map.get((key, s)) != "remote":
                    continue
                row_idx = key_to_row_l.get(key)
                if row_idx is None:
                    continue
                row_data = list(remote_rows.get(key) or [])
                while len(row_data) < max_col:
                    row_data.append("")
                for col_idx, val in enumerate(row_data[:max_col], start=1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)
                    cell.fill = yellow_fill
            next_row = ws.max_row + 1
            for key in only_remote_keys:
                row_data = list(remote_rows.get(key) or [])
                while len(row_data) < max_col:
                    row_data.append("")
                for col_idx, val in enumerate(row_data[:max_col], start=1):
                    ws.cell(row=next_row, column=col_idx, value=val)
                next_row += 1
        wb_merged.save(self.path_merged)
        wb_merged.close()
        # 记录实际写入的规范路径，供“打开合并结果”使用，避免打开错文件或 Excel 缓存
        self._merged_file_path = os.path.normpath(os.path.abspath(self.path_merged))
        os.makedirs(backup_dir, exist_ok=True)
        shutil.copy2(self.path_local, os.path.join(backup_dir, base_name + "_local.xlsx"))
        shutil.copy2(self.path_remote, os.path.join(backup_dir, base_name + "_remote.xlsx"))
        shutil.copy2(self.path_merged, os.path.join(backup_dir, base_name + "_merged.xlsx"))

    def run(self):
        self.root.mainloop()
