"""
WPS DISPIMG 图片转换核心逻辑 (ZIP/XML 层面操作)
将 WPS 创建的包含 =@_xlfn.DISPIMG 公式的 XLSX 转换为标准嵌入图片的 XLSX。
直接操作 ZIP 包内的 XML，保留所有原有图片和绘图。
"""

import zipfile
import re
import os
import shutil
import xml.etree.ElementTree as ET


REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'


def _col_letter_to_index(col_str):
    """将列字母转为 0-based 索引: A->0, B->1, ..., G->6"""
    idx = 0
    for ch in col_str:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def _parse_coord(coord):
    """解析单元格坐标如 'D5' -> (col_index=3, row_index=4)，均为 0-based"""
    m = re.match(r'([A-Z]+)(\d+)', coord)
    col = _col_letter_to_index(m.group(1))
    row = int(m.group(2)) - 1
    return col, row


def _build_two_cell_anchor(col, row, r_id, next_id):
    """构建 twoCellAnchor XML 字符串"""
    # 90px ≈ 857250 EMU, 使用一个单元格大小
    return (
        f'<xdr:twoCellAnchor>'
        f'<xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff>'
        f'<xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
        f'<xdr:to><xdr:col>{col}</xdr:col><xdr:colOff>857250</xdr:colOff>'
        f'<xdr:row>{row}</xdr:row><xdr:rowOff>857250</xdr:rowOff></xdr:to>'
        f'<xdr:pic>'
        f'<xdr:nvPicPr><xdr:cNvPr id="{next_id}" name="FK_Converted_{next_id}"/>'
        f'<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
        f'<xdr:blipFill><a:blip r:embed="rId{r_id}"/>'
        f'<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
        f'<xdr:spPr><a:xfrm>'
        f'<a:off x="0" y="0"/><a:ext cx="857250" cy="857250"/>'
        f'</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/><a:ln w="9525"><a:noFill/></a:ln></xdr:spPr>'
        f'</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>'
    )


def _find_next_rid(rels_xml):
    """在 rels XML 中找到下一个可用的 rId 编号"""
    existing = re.findall(r'Id="rId(\d+)"', rels_xml)
    return max(int(n) for n in existing) + 1 if existing else 1


def _find_max_anchor_id(drawing_xml):
    """在 drawing XML 中找到最大的 cNvPr id"""
    ids = re.findall(r'<xdr:cNvPr id="(\d+)"', drawing_xml)
    return max(int(n) for n in ids) + 1 if ids else 1


def _add_rels_entry(rels_xml, rid, target, rel_type):
    """向 rels XML 中添加一条 Relationship"""
    entry = (
        f'<Relationship Id="{rid}" '
        f'Type="{rel_type}" '
        f'Target="{target}"/>'
    )
    # 在 </Relationships> 前插入
    return rels_xml.replace('</Relationships>', f'{entry}\n</Relationships>')


def wps_image_converter(input_xlsx, output_xlsx, log_callback=None):
    """
    转换 WPS DISPIMG 嵌入图片为标准 Excel 嵌入图片。
    直接操作 ZIP/XML，保留所有原有图片和绘图。

    Returns:
        (success: bool, message: str, count: int)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log(f"正在加载表格: {input_xlsx} ...")

    if not os.path.isfile(input_xlsx):
        return False, f"文件不存在: {input_xlsx}", 0

    try:
        z_in = zipfile.ZipFile(input_xlsx, 'r')
    except Exception as e:
        return False, f"打开文件失败: {e}", 0

    # ── 第一步：解析 WPS 图片映射 ──────────────────────────
    if 'xl/cellimages.xml' not in z_in.namelist():
        z_in.close()
        return False, "未找到 WPS 嵌入图片容器 (cellimages.xml)", 0

    cellimages_xml = z_in.read('xl/cellimages.xml').decode('utf-8')
    cellimages_rels = z_in.read('xl/_rels/cellimages.xml.rels').decode('utf-8')

    # rId -> media path (相对 xl/)
    rid_to_media = {}
    for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', cellimages_rels):
        rid_to_media[m.group(1)] = m.group(2)

    # image_name -> media path
    name_to_media = {}
    for m in re.finditer(r'name="(ID_[^"]+)"', cellimages_xml):
        name = m.group(1)
        # 找该 cellImage 中最近的 r:embed
        pos = m.end()
        embed_m = re.search(r'r:embed="(rId\d+)"', cellimages_xml[pos:pos + 500])
        if embed_m:
            rid = embed_m.group(1)
            if rid in rid_to_media:
                name_to_media[name] = rid_to_media[rid]

    log(f"  发现 {len(name_to_media)} 张 WPS 嵌入图片")

    # ── 第二步：建立工作表与绘图文件的映射 ──────────────────
    # sheet_file -> drawing_file (如果有)
    sheet_to_drawing = {}
    drawing_rels_cache = {}  # drawing_file -> rels_xml

    for name in z_in.namelist():
        if re.match(r'xl/worksheets/_rels/sheet\d+\.xml\.rels', name):
            rels_content = z_in.read(name).decode('utf-8')
            sheet_file = name.replace('xl/worksheets/_rels/', '').replace('.xml.rels', '.xml')
            dm = re.search(r'Target="(\.\./drawings/drawing\d+\.xml)"', rels_content)
            if dm:
                drawing_path = 'xl/' + dm.group(1).replace('../', '')
                sheet_to_drawing[sheet_file] = {
                    'path': drawing_path,
                    'rels_path': drawing_path.replace('xl/drawings/', 'xl/drawings/_rels/') + '.rels',
                    'sheet_rels': name,
                    'sheet_rels_content': rels_content,
                }

    # ── 第三步：读取所有 ZIP 条目并修改 ────────────────────
    modified = {}  # path -> new_content (bytes)
    total_count = 0

    for sheet_name in sorted(z_in.namelist()):
        if not re.match(r'xl/worksheets/sheet\d+\.xml', sheet_name):
            continue

        content = z_in.read(sheet_name).decode('utf-8')
        sheet_basename = os.path.basename(sheet_name)

        # 找到所有 DISPIMG 单元格
        dispimg_cells = []
        for cm in re.finditer(
            r'<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>', content, re.DOTALL
        ):
            coord = cm.group(1)
            inner = cm.group(3)
            fm = re.search(r'_xlfn\.DISPIMG\(&quot;([^&]+)&quot;', inner)
            if fm:
                img_name = fm.group(1)
                if img_name in name_to_media:
                    dispimg_cells.append((coord, img_name))

        if not dispimg_cells:
            continue

        log(f"\n正在处理工作表: {sheet_basename} ({len(dispimg_cells)} 张图片)")

        # ── 3a: 清除 DISPIMG 公式 ──
        for coord, img_name in dispimg_cells:
            # 替换整个 <c> 元素，保留标签和样式属性，去掉子元素
            pattern = (
                r'<c r="' + re.escape(coord) + r'"'
                r'([^>]*)>.*?</c>'
            )
            replacement = f'<c r="{coord}"\\1/>'
            content = re.sub(pattern, replacement, content)
            log(f"  > 清除公式: {coord}")

        modified[sheet_name] = content.encode('utf-8')

        # ── 3b: 向绘图添加图片锚点 ──
        drawing_info = sheet_to_drawing.get(sheet_basename)
        if not drawing_info:
            # 没有绘图文件，需要创建
            # 找到下一个可用的 drawing 编号
            existing_drawings = [n for n in z_in.namelist()
                                 if re.match(r'xl/drawings/drawing\d+\.xml', n)]
            next_num = max(
                int(re.search(r'drawing(\d+)', n).group(1))
                for n in existing_drawings
            ) + 1 if existing_drawings else 1

            drawing_path = f'xl/drawings/drawing{next_num}.xml'
            rels_path = f'xl/drawings/_rels/drawing{next_num}.xml.rels'
            sheet_rels = f'xl/worksheets/_rels/{os.path.basename(sheet_name)}.rels'

            # 创建空绘图
            empty_drawing = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
                ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                '</xdr:wsDr>'
            )

            # 创建绘图 rels
            empty_rels = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '</Relationships>'
            )

            # 创建或更新 sheet rels
            if sheet_rels in z_in.namelist():
                sheet_rels_content = z_in.read(sheet_rels).decode('utf-8')
            else:
                sheet_rels_content = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '</Relationships>'
                )

            # 找到下一个可用的 rId
            next_rid = _find_next_rid(sheet_rels_content)
            sheet_rels_content = _add_rels_entry(
                sheet_rels_content,
                f'rId{next_rid}',
                f'../drawings/drawing{next_num}.xml',
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
            )

            # 在 sheet XML 中添加 <drawing> 引用
            sheet_content = content if sheet_name not in modified else modified[sheet_name].decode('utf-8')
            if '<drawing' not in sheet_content:
                sheet_content = sheet_content.replace(
                    '</worksheet>',
                    f'<drawing r:id="rId{next_rid}"/></worksheet>'
                )
                modified[sheet_name] = sheet_content.encode('utf-8')

            modified[sheet_rels] = sheet_rels_content.encode('utf-8')

            drawing_info = {
                'path': drawing_path,
                'rels_path': rels_path,
                'sheet_rels': sheet_rels,
                'sheet_rels_content': sheet_rels_content,
            }
            modified[drawing_path] = empty_drawing.encode('utf-8')
            modified[rels_path] = empty_rels.encode('utf-8')

        # 读取绘图内容（优先用已修改的版本）
        drawing_path = drawing_info['path']
        rels_path = drawing_info['rels_path']

        if drawing_path in modified:
            drawing_xml = modified[drawing_path].decode('utf-8')
        else:
            drawing_xml = z_in.read(drawing_path).decode('utf-8')

        if rels_path in modified:
            drawing_rels = modified[rels_path].decode('utf-8')
        elif rels_path in z_in.namelist():
            drawing_rels = z_in.read(rels_path).decode('utf-8')
        else:
            drawing_rels = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '</Relationships>'
            )

        next_rid = _find_next_rid(drawing_rels)
        next_anchor_id = _find_max_anchor_id(drawing_xml)

        for coord, img_name in dispimg_cells:
            media_path = name_to_media[img_name]
            # media path 格式: "media/imageNNN.png"
            # drawing rels 中用相对路径: "../media/imageNNN.png"
            rel_target = f'../{media_path}'

            # 检查这个 media 文件是否已经在 drawing rels 中
            existing_rid_match = re.search(
                rf'Target="{re.escape(rel_target)}"[^>]*Id="(rId\d+)"',
                drawing_rels
            )
            # 也检查 Id 在 Target 前面的情况
            if not existing_rid_match:
                existing_rid_match = re.search(
                    rf'Id="(rId\d+)"[^>]*Target="{re.escape(rel_target)}"',
                    drawing_rels
                )

            if existing_rid_match:
                img_rid = existing_rid_match.group(1)
            else:
                img_rid = f'rId{next_rid}'
                drawing_rels = _add_rels_entry(
                    drawing_rels, img_rid, rel_target,
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
                )
                next_rid += 1

            col, row = _parse_coord(coord)
            anchor_xml = _build_two_cell_anchor(col, row, img_rid, next_anchor_id)
            next_anchor_id += 1

            # 在 </xdr:wsDr> 前插入
            drawing_xml = drawing_xml.replace('</xdr:wsDr>', f'{anchor_xml}\n</xdr:wsDr>')
            total_count += 1
            log(f"  > 插入图片: {coord} <- {media_path}")

        modified[drawing_path] = drawing_xml.encode('utf-8')
        modified[rels_path] = drawing_rels.encode('utf-8')

    # ── 第四步：写入输出 ZIP ─────────────────────────────
    log(f"\n正在保存到: {output_xlsx} ...")

    try:
        with zipfile.ZipFile(output_xlsx, 'w', zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                if item.filename in modified:
                    z_out.writestr(item, modified[item.filename])
                else:
                    z_out.writestr(item, z_in.read(item.filename))
    except Exception as e:
        z_in.close()
        return False, f"保存文件失败: {e}", total_count

    z_in.close()

    msg = f"全部任务完成！累计修复 {total_count} 张图片。"
    log(f"{'=' * 50}")
    log(msg)
    return True, msg, total_count
