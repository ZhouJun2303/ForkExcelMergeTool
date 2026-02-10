# -*- coding: utf-8 -*-
"""
向后兼容层：对外仍提供 MergeWindow、DiffWindow，实际实现已迁至 merge_gui、diff_gui。
直接运行本文件时，用项目根目录的 TestData 打开合并窗口便于预览界面。
"""

import os
import tkinter as tk
from tkinter import messagebox

from merge_gui import MergeWindow
from diff_gui import DiffWindow

__all__ = ["MergeWindow", "DiffWindow"]

if __name__ == "__main__":
    # TestData 在项目根目录（Scripts 的上一级）
    _root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _local = os.path.join(_root, "TestData", "local.xlsx")
    _base = os.path.join(_root, "TestData", "base.xlsx")
    _remote = os.path.join(_root, "TestData", "remote.xlsx")
    _merged = os.path.join(_root, "TestData", "_output", "merged.xlsx")
    if os.path.isfile(_local) and os.path.isfile(_base) and os.path.isfile(_remote):
        os.makedirs(os.path.dirname(_merged) or ".", exist_ok=True)
        win = MergeWindow(_local, _base, _remote, _merged)
        win.run()
    else:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "Excel 合并/对比 GUI",
            "请通过以下方式打开界面：\n\n"
            "• 合并：运行 Scripts\\MergeExcelFork.py 并传入 4 个参数\n"
            "  python Scripts\\MergeExcelFork.py <local> <base> <remote> <merged>\n\n"
            "• 对比：运行 Scripts\\MergeExcelFork.py 并传入 2 个参数\n"
            "  python Scripts\\MergeExcelFork.py <local> <remote>\n\n"
            "或在 Fork 中配置 Merge Tool / Diff Tool 后触发。\n\n"
            "若需本地预览，请确保项目根目录 TestData 下有 local.xlsx、base.xlsx、remote.xlsx。"
        )
        root.destroy()
