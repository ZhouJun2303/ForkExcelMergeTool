# -*- coding: utf-8 -*-
"""
全局常量与配置。
仅做一件事：集中存放日志文件名、备份目录名、对比文件后缀、Sheet 跳过前缀等，
便于修改与打包时保持一致。
"""

# 与 exe/脚本同目录的日志文件名
LOG_FILE = "MergeExcelFork.log"

# 合并完成后，在 MERGED 所在目录下创建的备份子目录名
BACKUP_SUBDIR = "MergeExcelBackup"

# 对比模式生成的 Excel 文件名后缀（不含扩展名）
COMPARE_SUFFIX = "_compare"

# Sheet 名过滤：以该前缀开头的表不参与合并/对比（如 "#说明"）
SKIP_SHEET_PREFIX = "#"

# 合并选项多选框持久化文件名（与 exe/脚本同目录）
MERGE_OPTIONS_FILE = "merge_options.json"

# GUI 表格单次最多渲染行数；大 Excel 仍计算全量结果，但不一次性塞满 Tk Treeview。
MAX_TREEVIEW_ROWS = 500
