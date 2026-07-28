#!/usr/bin/env python3
"""
WPS 嵌入图片 XLSX 转换工具 - 命令行版
用法:
  python wps嵌入图片excel转换.py 输入文件.xlsx [输出文件.xlsx]
"""

import sys
from converter import wps_image_converter


def main():
    if len(sys.argv) < 2:
        print("用法: python wps嵌入图片excel转换.py 输入文件.xlsx [输出文件.xlsx]")
        print("\n如果不指定输出文件名，默认在输入文件名后加 _converted")
        sys.exit(1)

    input_file = sys.argv[1]

    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
    else:
        import os
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_converted{ext}"

    success, msg, count = wps_image_converter(input_file, output_file)
    if not success:
        sys.exit(1)


if __name__ == "__main__":
    main()
