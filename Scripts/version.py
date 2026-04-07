# -*- coding: utf-8 -*-
"""
版本号：大版本号手动改 VERSION_MAJOR，小版本号每次打包由 bump_version 自动递增。
界面显示用 __version__（格式 "大.小"）。
"""

VERSION_MAJOR = 2
VERSION_MINOR = 52

__version__ = "%d.%d" % (VERSION_MAJOR, VERSION_MINOR)
