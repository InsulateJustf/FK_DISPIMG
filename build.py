#!/usr/bin/env python3
"""
构建脚本 - 使用 PyInstaller 打包 GUI

用法:
  python build.py

macOS 上如果遇到权限问题，先设置:
  export PYINSTALLER_CONFIG_DIR=/tmp/pyinstaller_cache
"""

import subprocess
import sys
import os
import tempfile


def main():
    # 检查 PyInstaller
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller 未安装，正在安装...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # macOS 上设置临时缓存目录避免权限问题
    if sys.platform == "darwin":
        cache_dir = os.environ.get("PYINSTALLER_CONFIG_DIR")
        if not cache_dir:
            os.environ["PYINSTALLER_CONFIG_DIR"] = os.path.join(tempfile.gettempdir(), "pyinstaller_cache")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--windowed",
        "--name", "FK_DISPIMG",
        "--hidden-import", "converter",
        "--hidden-import", "tkinterdnd2",
        "--hidden-import", "openpyxl",
    ]

    # Windows 上用 --onefile 生成单个 exe
    if sys.platform == "win32":
        cmd.append("--onefile")

    cmd.append("gui.py")

    print("正在构建...")
    print(" ".join(cmd))
    subprocess.check_call(cmd)
    print("\n构建完成！")

    if sys.platform == "darwin":
        print("输出目录: dist/FK_DISPIMG.app")
        print("可直接运行: open dist/FK_DISPIMG.app")
    else:
        print("输出文件: dist/FK_DISPIMG.exe")


if __name__ == "__main__":
    main()
