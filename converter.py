"""
WPS DISPIMG 图片转换核心逻辑
将 WPS 创建的包含 =@_xlfn.DISPIMG 公式的 XLSX 转换为标准嵌入图片的 XLSX
"""

import zipfile
import re
import os
import shutil
import openpyxl
from openpyxl.drawing.image import Image
import xml.etree.ElementTree as ET


def wps_image_converter(input_xlsx, output_xlsx, log_callback=None):
    """
    转换 WPS DISPIMG 嵌入图片为标准 Excel 嵌入图片。

    Args:
        input_xlsx: 输入文件路径
        output_xlsx: 输出文件路径
        log_callback: 日志回调函数，接收字符串参数

    Returns:
        (success: bool, message: str, count: int)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    temp_dir = "temp_wps_images"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir, ignore_errors=True)
    os.makedirs(temp_dir)

    log(f"正在加载表格: {input_xlsx} ...")

    try:
        wb = openpyxl.load_workbook(input_xlsx, data_only=False)
    except Exception as e:
        msg = f"打开 Excel 失败: {e}"
        log(msg)
        return False, msg, 0

    id_to_target = {}
    total_success_count = 0

    try:
        with zipfile.ZipFile(input_xlsx, 'r') as z:
            if 'xl/cellimages.xml' not in z.namelist():
                msg = "未在文件中找到 WPS 嵌入图片容器，请确认文件是否包含嵌入图片。"
                log(msg)
                return False, msg, 0

            cellimages_bytes = z.read('xl/cellimages.xml')
            rels_bytes = z.read('xl/_rels/cellimages.xml.rels')

            rId_to_target = {}
            rels_root = ET.fromstring(rels_bytes)
            for elem in rels_root.iter():
                if elem.tag.endswith('Relationship'):
                    rid = elem.attrib.get('Id')
                    target = elem.attrib.get('Target')
                    if rid and target:
                        rId_to_target[rid] = target

            cellimages_root = ET.fromstring(cellimages_bytes)
            for child in cellimages_root:
                name, embed = None, None
                for sub in child.iter():
                    for k, v in sub.attrib.items():
                        if k.endswith('name') and str(v).startswith('ID_'):
                            name = v
                        if k.endswith('embed'):
                            embed = v
                if name and embed and embed in rId_to_target:
                    id_to_target[name] = rId_to_target[embed]

            for sheet in wb.worksheets:
                log(f"\n正在处理工作表: [{sheet.title}]")
                sheet_success_count = 0

                cell_to_id = {}
                for row in sheet.iter_rows():
                    for cell in row:
                        cell_val = str(cell.value)
                        if "DISPIMG" in cell_val:
                            match = re.search(r'DISPIMG\s*\(\s*"([^"]+)"', cell_val)
                            if match:
                                cell_to_id[cell.coordinate] = match.group(1)

                if not cell_to_id:
                    log(f"  > 该工作表未检测到嵌入图片，跳过。")
                    continue

                for coord, img_id in cell_to_id.items():
                    target_path = id_to_target.get(img_id)
                    if not target_path:
                        continue

                    if target_path.startswith('/'):
                        full_target_path = target_path[1:]
                    elif target_path.startswith('xl/'):
                        full_target_path = target_path
                    else:
                        full_target_path = f"xl/{target_path}"

                    ext = full_target_path.split('.')[-1]
                    safe_title = re.sub(r'[\\/*?:\[\]]', '_', sheet.title)
                    temp_img_path = os.path.join(temp_dir, f"{safe_title}_{coord}.{ext}")

                    try:
                        with z.open(full_target_path) as source, open(temp_img_path, 'wb') as target_file:
                            shutil.copyfileobj(source, target_file)

                        img = Image(temp_img_path)
                        img.width, img.height = 90, 90
                        sheet.add_image(img, coord)
                        sheet[coord].value = ""

                        sheet_success_count += 1
                        total_success_count += 1
                    except Exception as e:
                        log(f"  > [失败] 单元格 {coord} 报错: {e}")

                log(f"  > 完成！该表修复了 {sheet_success_count} 张图片。")

    except Exception as e:
        msg = f"解析底层数据时发生致命错误: {e}"
        log(msg)
        return False, msg, total_success_count

    log(f"\n正在保存到: {output_xlsx} ...")
    try:
        wb.save(output_xlsx)
    except Exception as e:
        msg = f"保存文件失败: {e}"
        log(msg)
        return False, msg, total_success_count
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    msg = f"全部任务完成！累计修复 {total_success_count} 张图片。"
    log(f"{'=' * 50}")
    log(msg)
    return True, msg, total_success_count
