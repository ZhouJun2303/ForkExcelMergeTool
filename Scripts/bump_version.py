# -*- coding: utf-8 -*-
"""
打包前执行：将 version.py 中的 VERSION_MINOR 自增 1，便于确认当前 exe 是否为最新打包。
大版本号需在 version.py 中手动修改 VERSION_MAJOR。
"""

import os
import re

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_VERSION_PY = os.path.join(_SCRIPT_DIR, "version.py")


def main():
    with open(_VERSION_PY, "r", encoding="utf-8") as f:
        content = f.read()
    # 将 VERSION_MINOR = N 改为 N+1
    def replace_minor(m):
        n = int(m.group(1))
        return "VERSION_MINOR = %d" % (n + 1)
    new_content, n = re.subn(r"VERSION_MINOR\s*=\s*(\d+)", replace_minor, content)
    if n == 0:
        print("WARNING: 未找到 VERSION_MINOR，未修改 version.py")
        return
    with open(_VERSION_PY, "w", encoding="utf-8") as f:
        f.write(new_content)
    import importlib.util
    spec = importlib.util.spec_from_file_location("version", _VERSION_PY)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    print("版本号已递增: %s" % getattr(mod, "__version__", "?"))


if __name__ == "__main__":
    main()
