# -*- coding: utf-8 -*-
"""
根目录启动器：使用 runpy 执行 Scripts/MergeExcelFork.py，避免与当前模块名冲突。
规范：所有脚本放在 Scripts 文件夹下，根目录仅保留此启动器及批处理、配置、文档。
"""

import os
import runpy
import sys

_SCRIPTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Scripts")
_ENTRY = os.path.join(_SCRIPTS, "MergeExcelFork.py")

if __name__ == "__main__":
    if not os.path.isfile(_ENTRY):
        print("ERROR: Scripts/MergeExcelFork.py 不存在", file=sys.stderr)
        sys.exit(1)
    if _SCRIPTS not in sys.path:
        sys.path.insert(0, _SCRIPTS)
    runpy.run_path(_ENTRY, run_name="__main__")
