"""
WPS DISPIMG 图片转换核心逻辑 (ZIP/XML 层面操作)
直接操作 ZIP 包内的 XML，保留所有原有图片和绘图。
"""

import zipfile
import re
import os

REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

_RE_RID_NUM = re.compile(r'Id="rId(\d+)"')
_RE_TARGET = re.compile(r'Target="([^"]+)"')
_RE_ANCHOR_ID = re.compile(r'<xdr:cNvPr id="(\d+)"')
_RE_REL_ENTRY = re.compile(r'<Relationship[^>]+/>')
_RE_NAME_ID = re.compile(r'name="(ID_[^"]+)"')
_RE_EMBED_RID = re.compile(r'r:embed="(rId\d+)"')
_RE_DRAWING_NUM = re.compile(r'xl/drawings/drawing(\d+)\.xml$')
_RE_SHEET_RELS = re.compile(r'xl/worksheets/_rels/sheet(\d+)\.xml\.rels$')
_RE_SHEET_XML = re.compile(r'xl/worksheets/sheet(\d+)\.xml$')

WPS_DRAWING_TPL = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '</xdr:wsDr>'
)
EMPTY_RELS_TPL = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<Relationships xmlns="{REL_NS}">'
    '</Relationships>'
)
DRAWING_CT = 'application/vnd.openxmlformats-officedocument.drawing+xml'


def _col_to_idx(s):
    idx = 0
    for ch in s:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def _parse_coord(coord):
    m = re.match(r'([A-Z]+)(\d+)', coord)
    return _col_to_idx(m.group(1)), int(m.group(2)) - 1


def _build_anchor(col, row, r_id, aid):
    return (
        f'<xdr:twoCellAnchor editAs="oneCell">'
        f'<xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff>'
        f'<xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
        f'<xdr:to><xdr:col>{col + 1}</xdr:col><xdr:colOff>0</xdr:colOff>'
        f'<xdr:row>{row + 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
        f'<xdr:pic>'
        f'<xdr:nvPicPr><xdr:cNvPr id="{aid}" name="FK_Converted_{aid}"/>'
        f'<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
        f'<xdr:blipFill><a:blip r:embed="{r_id}"/>'
        f'<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
        f'<xdr:spPr><a:xfrm>'
        f'<a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>'
        f'</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/><a:ln w="9525"><a:noFill/></a:ln></xdr:spPr>'
        f'</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>'
    )


def _next_rid(xml):
    ids = [int(m) for m in _RE_RID_NUM.findall(xml)]
    return max(ids) + 1 if ids else 1


def _next_aid(xml):
    ids = [int(m) for m in _RE_ANCHOR_ID.findall(xml)]
    return max(ids) + 1 if ids else 1


def _add_rel(xml, rid, target, rtype):
    entry = f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>'
    return xml.replace('</Relationships>', entry + '\n</Relationships>')


def _find_rid_for_target(xml, target):
    for m in _RE_REL_ENTRY.finditer(xml):
        entry = m.group(0)
        id_m = _RE_RID_NUM.search(entry)
        t_m = _RE_TARGET.search(entry)
        if id_m and t_m and t_m.group(1) == target:
            return f'rId{id_m.group(1)}'
    return None


def _find_and_clear_dispimg(content, name_to_media):
    """
    找所有 DISPIMG 单元格并清除公式，返回 (cells, new_content)。
    用逐字符扫描代替 split，确保 XML 结构完整。
    """
    result = []
    output = []
    i = 0
    n = len(content)

    while i < n:
        # 找 <c 标签开头
        c_start = content.find('<c ', i)
        if c_start == -1:
            output.append(content[i:])
            break

        # 写入 <c 之前的内容
        output.append(content[i:c_start])

        # 找这个 <c> 的结束标签 </c>
        c_end = content.find('</c>', c_start)
        if c_end == -1:
            # 没有闭合标签，直接写入剩余内容
            output.append(content[c_start:])
            break

        cell_full = content[c_start:c_end + 4]  # 包含 </c>

        # 检查是否包含 DISPIMG
        if 'DISPIMG' in cell_full:
            coord_m = re.search(r'<c r="([A-Z]+\d+)"', cell_full)
            img_m = re.search(r'_xlfn\.DISPIMG\(&quot;([^&]+)&quot;', cell_full)
            if coord_m and img_m:
                coord = coord_m.group(1)
                img_name = img_m.group(1)
                if img_name in name_to_media:
                    result.append((coord, img_name))
                    # 提取 <c r="XX" attrs> 的开头标签，转为自闭合
                    tag_end = cell_full.find('>')
                    open_tag = cell_full[:tag_end]
                    output.append(open_tag + '/>')
                    i = c_end + 4
                    continue

        # 非 DISPIMG 单元格，原样写入
        output.append(cell_full)
        i = c_end + 4

    return result, ''.join(output)


def wps_image_converter(input_xlsx, output_xlsx, log_callback=None):
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

    all_names = set(z_in.namelist())

    if 'xl/cellimages.xml' not in all_names:
        z_in.close()
        return False, "未找到 WPS 嵌入图片容器", 0

    ci_xml = z_in.read('xl/cellimages.xml').decode('utf-8')
    ci_rels = z_in.read('xl/_rels/cellimages.xml.rels').decode('utf-8')

    rid_to_media = {}
    for m in _RE_REL_ENTRY.finditer(ci_rels):
        entry = m.group(0)
        id_m = _RE_RID_NUM.search(entry)
        t_m = _RE_TARGET.search(entry)
        if id_m and t_m:
            rid_to_media[f'rId{id_m.group(1)}'] = t_m.group(1)

    name_to_media = {}
    for m in _RE_NAME_ID.finditer(ci_xml):
        name = m.group(1)
        embed_m = _RE_EMBED_RID.search(ci_xml[m.end():m.end() + 500])
        if embed_m and embed_m.group(1) in rid_to_media:
            name_to_media[name] = rid_to_media[embed_m.group(1)]

    log(f"  发现 {len(name_to_media)} 张 WPS 嵌入图片")

    used_drawing_nums = set()
    for name in all_names:
        dm = _RE_DRAWING_NUM.match(name)
        if dm:
            used_drawing_nums.add(int(dm.group(1)))

    def alloc_num():
        n = 1
        while n in used_drawing_nums:
            n += 1
        used_drawing_nums.add(n)
        return n

    sheet_drawing = {}
    for name in all_names:
        dm = _RE_SHEET_RELS.match(name)
        if not dm:
            continue
        rels_content = z_in.read(name).decode('utf-8')
        t_m = re.search(r'Target="(\.\./drawings/drawing\d+\.xml)"', rels_content)
        if t_m:
            dp = 'xl/' + t_m.group(1).replace('../', '')
            sheet_drawing[f'sheet{dm.group(1)}.xml'] = {
                'drawing': dp,
                'drawing_rels': dp.replace('xl/drawings/', 'xl/drawings/_rels/') + '.xml.rels',
                'sheet_rels': name,
            }

    modified = {}
    new_drawings = []
    total_count = 0

    for sheet_path in sorted(all_names):
        sm = _RE_SHEET_XML.match(sheet_path)
        if not sm:
            continue

        content = z_in.read(sheet_path).decode('utf-8')
        if 'DISPIMG' not in content:
            continue

        sheet_base = os.path.basename(sheet_path)

        dispimg_cells, content = _find_and_clear_dispimg(content, name_to_media)
        if not dispimg_cells:
            continue

        log(f"\n正在处理工作表: [{sheet_base}] ({len(dispimg_cells)} 张图片)")
        for coord, _ in dispimg_cells:
            log(f"  > 清除公式: {coord}")

        dinfo = sheet_drawing.get(sheet_base)

        if dinfo:
            dp = dinfo['drawing']
            rp = dinfo['drawing_rels']
            srp = dinfo['sheet_rels']
        else:
            num = alloc_num()
            dp = f'xl/drawings/drawing{num}.xml'
            rp = f'xl/drawings/_rels/drawing{num}.xml.rels'
            srp = f'xl/worksheets/_rels/{sheet_base}.rels'

            modified[dp] = WPS_DRAWING_TPL.encode('utf-8')
            modified[rp] = EMPTY_RELS_TPL.encode('utf-8')
            new_drawings.append(dp)

            if srp in all_names:
                sr = z_in.read(srp).decode('utf-8')
            else:
                sr = EMPTY_RELS_TPL

            sr_rid = _next_rid(sr)
            sr = _add_rel(
                sr, f'rId{sr_rid}',
                f'../drawings/drawing{num}.xml',
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
            )
            modified[srp] = sr.encode('utf-8')

            content = content.replace(
                '</worksheet>',
                f'<drawing r:id="rId{sr_rid}"/></worksheet>'
            )

        if dp in modified:
            dx = modified[dp].decode('utf-8')
        elif dp in all_names:
            dx = z_in.read(dp).decode('utf-8')
        else:
            dx = WPS_DRAWING_TPL

        if rp in modified:
            dr = modified[rp].decode('utf-8')
        elif rp in all_names:
            dr = z_in.read(rp).decode('utf-8')
        else:
            dr = EMPTY_RELS_TPL

        nrid = _next_rid(dr)
        naid = _next_aid(dx)

        for coord, img_name in dispimg_cells:
            media = name_to_media[img_name]
            target = f'../{media}'

            existing_rid = _find_rid_for_target(dr, target)
            if existing_rid:
                img_rid = existing_rid
            else:
                img_rid = f'rId{nrid}'
                dr = _add_rel(
                    dr, img_rid, target,
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
                )
                nrid += 1

            col, row = _parse_coord(coord)
            dx = dx.replace(
                '</xdr:wsDr>',
                _build_anchor(col, row, img_rid, naid) + '\n</xdr:wsDr>'
            )
            naid += 1
            total_count += 1
            log(f"  > 插入图片: {coord} <- {media}")

        modified[sheet_path] = content.encode('utf-8')
        modified[dp] = dx.encode('utf-8')
        modified[rp] = dr.encode('utf-8')

    if new_drawings:
        ct_path = '[Content_Types].xml'
        ct_xml = z_in.read(ct_path).decode('utf-8')
        for drawing_path in new_drawings:
            part_name = '/' + drawing_path
            if part_name not in ct_xml:
                override = f'<Override PartName="{part_name}" ContentType="{DRAWING_CT}"/>'
                ct_xml = ct_xml.replace('</Types>', override + '\n</Types>')
        modified[ct_path] = ct_xml.encode('utf-8')

    log(f"\n正在保存到: {output_xlsx} ...")

    try:
        with zipfile.ZipFile(output_xlsx, 'w', zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                data = modified.get(item.filename)
                z_out.writestr(item, data if data is not None else z_in.read(item.filename))
            for path, data in modified.items():
                if path not in all_names:
                    z_out.writestr(path, data)
    except Exception as e:
        z_in.close()
        return False, f"保存文件失败: {e}", total_count

    z_in.close()

    msg = f"全部任务完成！累计修复 {total_count} 张图片。"
    log(f"{'=' * 50}")
    log(msg)
    return True, msg, total_count
