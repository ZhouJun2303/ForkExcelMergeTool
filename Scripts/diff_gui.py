# -*- coding: utf-8 -*-
"""
对比窗口：二向对比的 GUI。只做一件事——展示 A/B 差异列表、生成并打开对比 Excel，关闭时可选清理临时文件。
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox

from compare_core import get_compare_data, write_compare_excel
from gui_common import gui_log, make_color_legend, open_excel_file, setup_merge_styles
from version import __version__ as APP_VERSION


class DiffWindow:
    """对比 GUI：内嵌差异表格，关闭时可选删除临时对比文件。"""

    def __init__(self, path_a, path_b):
        self.path_a = path_a
        self.path_b = path_b
        self.path_out = None
        self.diff_rows = []
        self.is_temp = False
        self.root = tk.Tk()
        self.root.title("Excel 二向对比 v%s" % APP_VERSION)
        self.root.minsize(700, 500)
        self.root.geometry("1000x620")
        setup_merge_styles(self.root)
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        pad = 12
        top = ttk.Frame(self.root, padding=(pad, pad, pad, 6))
        top.pack(fill=tk.X)
        title_row = ttk.Frame(top)
        title_row.pack(fill=tk.X)
        ttk.Label(title_row, text="本地 (A) vs 线上 (B)", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT)
        tk.Label(title_row, text="  版本 v%s" % APP_VERSION, font=("Segoe UI", 9, "bold"), fg="#1877f2", bg="#f0f2f5").pack(side=tk.LEFT)
        ttk.Label(top, text="A: %s" % (self.path_a[:70] + "…" if len(self.path_a) > 70 else self.path_a), font=("Segoe UI", 8)).pack(anchor=tk.W)
        ttk.Label(top, text="B: %s" % (self.path_b[:70] + "…" if len(self.path_b) > 70 else self.path_b), font=("Segoe UI", 8)).pack(anchor=tk.W)
        legend_frame = tk.Frame(self.root, bg="#f0f2f5")
        legend_frame.pack(fill=tk.X, padx=pad, pady=(0, 6))
        make_color_legend(legend_frame, [
            ("#CCFFCC", "绿色=B新增"),
            ("#FFCCCC", "红色=A独有"),
            ("#FFFF99", "黄色=修改"),
            (None, "无=相同"),
        ]).pack(anchor=tk.W)
        paned = ttk.PanedWindow(self.root, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=pad, pady=(0, 6))
        table_frame = ttk.LabelFrame(paned, text="差异列表", padding=6)
        paned.add(table_frame, weight=2)
        cols = ("Sheet", "Key", "状态", "A(本地)", "B(线上)")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=14)
        for c, w in [("Sheet", 80), ("Key", 120), ("状态", 70), ("A(本地)", 200), ("B(线上)", 200)]:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w)
        sb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        sbh = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb.set, xscrollcommand=sbh.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        sbh.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.tag_configure("B新增", background="#ccffcc")
        self.tree.tag_configure("A独有", background="#ffcccc")
        self.tree.tag_configure("修改", background="#ffff99")
        self.status_var = tk.StringVar(value="")
        status_bar = ttk.Frame(self.root)
        status_bar.pack(fill=tk.X, padx=pad, pady=(0, 4))
        ttk.Label(status_bar, textvariable=self.status_var, font=("Segoe UI", 9)).pack(side=tk.LEFT, anchor=tk.W)
        _ver_frame = tk.Frame(status_bar, bg="#f0f2f5")
        _ver_frame.pack(side=tk.RIGHT)
        tk.Label(_ver_frame, text="版本 v%s" % APP_VERSION, font=("Segoe UI", 9, "bold"), fg="#1877f2", bg="#f0f2f5").pack()
        try:
            out_dir = os.path.dirname(os.path.abspath(self.path_a))
            base_name = os.path.splitext(os.path.basename(self.path_a))[0]
            from config import COMPARE_SUFFIX
            self.path_out = os.path.join(out_dir, base_name + COMPARE_SUFFIX + ".xlsx")
            self.is_temp = "Temp" in self.path_a or "Fork" in self.path_a or "tmp" in self.path_a.lower()
            gui_log("正在计算差异…", self.status_var)
            path_out, sheet_names, self.diff_rows = get_compare_data(self.path_a, self.path_b)
            if path_out is None:
                raise RuntimeError("get_compare_data 失败")
            self.path_out = path_out
            write_compare_excel(self.path_out, sheet_names, self.diff_rows, open_file=False)
            gui_log("已生成对比文件，共 %d 行差异" % len(self.diff_rows), self.status_var)
            for sheet_name, key, status, str_a, str_b in self.diff_rows:
                sa = (str_a[:60] + "…") if len(str_a) > 60 else str_a
                sb_val = (str_b[:60] + "…") if len(str_b) > 60 else str_b
                self.tree.insert("", tk.END, values=(sheet_name, key, status, sa, sb_val), tags=(status,))
        except Exception as e:
            import traceback
            gui_log("对比失败: " + str(e), self.status_var, is_error=True)
            messagebox.showerror("错误", str(e))
        btn_frame = ttk.Frame(self.root, padding=(pad, 8))
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="打开 Excel", command=self._open_excel, style="Accent.TButton").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="关闭", command=self._on_close).pack(side=tk.LEFT)

    def _open_excel(self):
        if self.path_out and os.path.isfile(self.path_out):
            open_excel_file(self.path_out)
            gui_log("已打开对比文件", self.status_var)
        else:
            messagebox.showwarning("提示", "对比文件不存在")

    def _on_close(self):
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
