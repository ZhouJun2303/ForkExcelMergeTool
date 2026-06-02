# -*- coding: utf-8 -*-
"""
Excel 文件后缀定义。
"""


COMMON_EXCEL_EXTENSIONS = [
    "xls",
    "xlsx",
    "xlsm",
    "xlsb",
    "xlt",
    "xltx",
    "xltm",
    "xla",
    "xlam",
    "xlw",
]

MERGE_DIFF_EXTENSIONS = ["xlsx", "xlsm", "xltx", "xltm"]


def extension_text():
    return ", ".join("." + ext for ext in COMMON_EXCEL_EXTENSIONS)


def git_attr_pattern(ext):
    return "*.%s merge=excelmergefork" % "".join("[%s%s]" % (c.lower(), c.upper()) for c in ext)


def git_attr_lines():
    return ["# ExcelMergeFork managed entry"] + [git_attr_pattern(ext) for ext in COMMON_EXCEL_EXTENSIONS]


def normalized_ext(path):
    ext = (path or "").rsplit(".", 1)
    return ext[-1].lower() if len(ext) == 2 else ""


def merge_diff_supported(path):
    return normalized_ext(path) in MERGE_DIFF_EXTENSIONS


def merge_diff_extension_text():
    return ", ".join("." + ext for ext in MERGE_DIFF_EXTENSIONS)
