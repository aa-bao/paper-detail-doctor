# -*- coding: utf-8 -*-
"""paper-template-extractor 公共工具。

职责：字号换算、OOXML 底层操作（东亚字体槽、首行缩进字符、页码域、节页码格式、三线表边框）。
只依赖 python-docx + lxml。
"""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Pt

# ---------------------------------------------------------------- 字号换算

# 中文字号 -> 磅值（pt）
CN_SIZE_TO_PT = {
    "初号": 42.0, "小初": 36.0,
    "一号": 26.0, "小一": 24.0,
    "二号": 22.0, "小二": 18.0,
    "三号": 16.0, "小三": 15.0,
    "四号": 14.0, "小四": 12.0,
    "五号": 10.5, "小五": 9.0,
    "六号": 7.5, "小六": 6.5,
    "七号": 5.5, "八号": 5.0,
}

# 磅值 -> 最接近的中文字号（用于报告反查）
PT_TO_CN_SIZE = {
    42.0: "初号", 36.0: "小初", 26.0: "一号", 24.0: "小一",
    22.0: "二号", 18.0: "小二", 16.0: "三号", 15.0: "小三",
    14.0: "四号", 12.0: "小四", 10.5: "五号", 9.0: "小五",
    7.5: "六号", 6.5: "小六", 5.5: "七号", 5.0: "八号",
}


def cn_size_to_pt(label, default=None):
    """'三号' -> 16.0；无法识别时返回 default。"""
    if label is None:
        return default
    s = str(label).strip().replace(" ", "")
    return CN_SIZE_TO_PT.get(s, default)


def pt_to_cn_size(pt):
    """16.0 -> '三号'（取最接近的整档）。"""
    if pt is None:
        return None
    try:
        pt = float(pt)
    except (TypeError, ValueError):
        return None
    if pt in PT_TO_CN_SIZE:
        return PT_TO_CN_SIZE[pt]
    best = min(PT_TO_CN_SIZE, key=lambda k: abs(k - pt))
    if abs(best - pt) <= 0.6:
        return PT_TO_CN_SIZE[best]
    return f"{pt:g}磅"


ALIGN_MAP = {
    "left": "LEFT", "left_align": "LEFT", "顶格": "LEFT", "左对齐": "LEFT",
    "center": "CENTER", "centre": "CENTER", "居中": "CENTER",
    "right": "RIGHT", "右对齐": "RIGHT", "右对齐齐": "RIGHT",
    "justify": "JUSTIFY", "两端对齐": "JUSTIFY", "兩端對齊": "JUSTIFY",
}


def resolve_align(value, default="JUSTIFY"):
    """把中英文对齐描述统一成 WD_ALIGN_PARAGRAPH 名称。"""
    if value is None:
        return default
    s = str(value).strip().lower()
    for k, v in ALIGN_MAP.items():
        if k.lower() == s:
            return v
    return default


# ---------------------------------------------------------------- 字体槽

def set_east_asian_font(rpr, cn_font=None, en_font=None):
    """在 rPr 上写入 w:rFonts（ascii/hAnsi/eastAsia/cs）。

    Word/WPS 的中文渲染取决于 w:eastAsia，python-docx 不直接暴露，必须手写。
    """
    rfonts = rpr.get_or_add_rFonts()
    if en_font:
        rfonts.set(qn("w:ascii"), en_font)
        rfonts.set(qn("w:hAnsi"), en_font)
        rfonts.set(qn("w:cs"), en_font)
    if cn_font:
        rfonts.set(qn("w:eastAsia"), cn_font)


# ---------------------------------------------------------------- 段落属性

def set_first_line_indent_chars(pPr, chars):
    """首行缩进 N 个字符（Word 用 w:firstLineChars，单位百分之一字符）。"""
    if not chars:
        return
    ind = pPr.get_or_add_ind()
    ind.set(qn("w:firstLineChars"), str(int(round(float(chars) * 100))))
    ind.set(qn("w:firstLine"), "0")


def set_left_indent_chars(pPr, chars):
    """左缩进 N 个字符。"""
    if not chars:
        return
    ind = pPr.get_or_add_ind()
    ind.set(qn("w:leftChars"), str(int(round(float(chars) * 100))))
    ind.set(qn("w:left"), "0")


def set_snap_to_grid(pPr, on=False):
    """是否对齐文档网格。中文论文常需关闭以精确控制行距。"""
    el = OxmlElement("w:snapToGrid")
    el.set(qn("w:val"), "1" if on else "0")
    pPr.append(el)


def set_outline_level(pPr, level):
    """设置大纲级别（0-8）。

    标题样式自带大纲级别，目录域（TOC \\o）与导航窗格才认得 —— 这样就不必把
    标题样式挂在 Word 内置的 Heading 1/2/3 上，能彻底摆脱内置样式依赖。
    """
    if level is None:
        return
    for old in pPr.findall(qn("w:outlineLvl")):
        pPr.remove(old)
    el = OxmlElement("w:outlineLvl")
    el.set(qn("w:val"), str(int(level)))
    # CT_PPr 顺序：outlineLvl 必须排在 rPr/sectPr 之前
    for tag in ("w:rPr", "w:sectPr", "w:pPrChange"):
        anchor = pPr.find(qn(tag))
        if anchor is not None:
            anchor.addprevious(el)
            return
    pPr.append(el)


# ---------------------------------------------------------------- 节属性

_SECTPR_AFTER_PGNUM = ("w:cols", "w:formProt", "w:vAlign", "w:noEndnote",
                       "w:titlePg", "w:textDirection", "w:bidi", "w:rtlGutter",
                       "w:docGrid", "w:printerSettings", "w:sectPrChange")


def set_section_page_numbering(section, fmt=None, start=None):
    """设置节的页码格式与起始值。

    fmt: 'decimal' | 'lowerRoman' | 'upperRoman' | 'lowerLetter' | 'upperLetter'
    """
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn("w:pgNumType"))
    if pgNumType is None:
        pgNumType = OxmlElement("w:pgNumType")
        inserted = False
        for child in sectPr:
            if child.tag in (qn(t) for t in _SECTPR_AFTER_PGNUM):
                child.addprevious(pgNumType)
                inserted = True
                break
        if not inserted:
            sectPr.append(pgNumType)
    if fmt:
        pgNumType.set(qn("w:fmt"), fmt)
    if start is not None:
        pgNumType.set(qn("w:start"), str(int(start)))


def set_section_no_header(section):
    """让该节页眉为空且与上一节不联动。"""
    try:
        section.header.is_linked_to_previous = False
        for p in list(section.header.paragraphs):
            p.text = ""
    except Exception:
        pass


def _part_has_content(root):
    """页眉/页脚部件里是否真有点东西（文字或域）。"""
    for el in root.iter():
        if el.tag == qn("w:t") and (el.text or "").strip():
            return True
        if el.tag in (qn("w:fldChar"), qn("w:instrText")):
            return True
    return False


def retarget_header_footer_styles(doc, header_style=None, footer_style=None):
    """把页眉/页脚段落改用本模板自己的「页眉」「页码」样式。

    为什么要做：python-docx 的页眉/页脚部件模板自带内置样式（w:pStyle=Header/Footer），
    只要留着这个引用，Word/WPS 就会把「Header / Footer」算作"已使用"，清理时删不掉，
    样式库里就一直多出两条与本文档无关的项。

    空白页眉（规范里"摘要、目录部分无页眉"）不挂样式 —— 否则页眉样式自带的
    下划线会在空白页眉上画出一条线。

    返回被改写的段落数。
    """
    ids = {}
    for key, name in (("header", header_style), ("footer", footer_style)):
        if not name:
            continue
        try:
            ids[key] = doc.styles[name].style_id
        except KeyError:
            pass

    changed = 0
    for rel in list(getattr(doc.part, "rels", {}).values()):
        try:
            if getattr(rel, "is_external", False):
                continue
            rt = rel.reltype
        except Exception:
            continue
        if rt not in (RT.HEADER, RT.FOOTER):
            continue
        try:
            root = rel.target_part.element
        except Exception:
            continue
        want = ids.get("header" if rt == RT.HEADER else "footer")
        if want is not None and not _part_has_content(root):
            want = None
        for p in root.findall(qn("w:p")):
            pPr = p.find(qn("w:pPr"))
            ps = pPr.find(qn("w:pStyle")) if pPr is not None else None
            if want is None:
                if ps is not None:
                    pPr.remove(ps)
                    changed += 1
                continue
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                p.insert(0, pPr)
            if ps is None:
                ps = OxmlElement("w:pStyle")
                pPr.insert(0, ps)
            if ps.get(qn("w:val")) != want:
                ps.set(qn("w:val"), want)
                changed += 1
    return changed


# ---------------------------------------------------------------- 域代码

def add_page_number_field(paragraph):
    """在段落中插入 PAGE 域（页码）。"""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE  \\* MERGEFORMAT "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)
    return run


def add_toc_field(paragraph, levels="1-3"):
    """插入 TOC 域（目录）。Word/WPS 打开后按 F9 或允许更新即可展开。"""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f' TOC \\o "{levels}" \\h \\z \\u '
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "右键此处选择“更新域”生成目录"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    for el in (fld_begin, instr, sep, placeholder, fld_end):
        run._r.append(el)
    return run


# ---------------------------------------------------------------- 三线表

def _border_el(name, val, sz_eighth_pt=0, color="000000"):
    el = OxmlElement(f"w:{name}")
    el.set(qn("w:val"), val)
    el.set(qn("w:sz"), str(int(sz_eighth_pt)))
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), color)
    return el


def apply_three_line_table(table, top_pt=1.5, bottom_pt=1.5, middle_pt=0.5,
                           header_row_index=0, color="000000"):
    """把表格设成三线表：上下线粗、表头下中线细、无竖线与内部横线。"""
    top_sz = int(round(top_pt * 8))
    bottom_sz = int(round(bottom_pt * 8))
    middle_sz = int(round(middle_pt * 8))

    tbl = table._tbl
    tblPr = tbl.tblPr

    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)

    borders = OxmlElement("w:tblBorders")
    borders.append(_border_el("top", "single", top_sz, color))
    borders.append(_border_el("left", "none"))
    borders.append(_border_el("bottom", "single", bottom_sz, color))
    borders.append(_border_el("right", "none"))
    borders.append(_border_el("insideH", "none"))
    borders.append(_border_el("insideV", "none"))
    tblPr.append(borders)

    # 表头行底线：作为三线表的“中线”
    if header_row_index is not None and header_row_index < len(table.rows):
        for cell in table.rows[header_row_index].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for old in tcPr.findall(qn("w:tcBorders")):
                tcPr.remove(old)
            tcBorders = OxmlElement("w:tcBorders")
            tcBorders.append(_border_el("bottom", "single", middle_sz, color))
            tcPr.append(tcBorders)


def set_table_layout_fixed(table, widths_cm=None):
    """固定表格布局，可选按 cm 指定各列宽度。"""
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    if widths_cm:
        from docx.shared import Cm
        table.autofit = False
        for row in table.rows:
            for idx, w in enumerate(widths_cm):
                if idx < len(row.cells):
                    row.cells[idx].width = Cm(w)


def set_repeat_header_row(row):
    """表头行跨页重复。"""
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def set_cell_shading(cell, hex_fill):
    """单元格底纹。hex_fill 形如 'D9D9D9'。"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def set_cell_valign(cell, val="center"):
    """单元格垂直对齐：top | center | bottom。"""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:vAlign")):
        tcPr.remove(old)
    el = OxmlElement("w:vAlign")
    el.set(qn("w:val"), val)
    tcPr.append(el)


# ---------------------------------------------------------------- 多级编号

def ensure_heading_numbering(doc, levels, abstract_ids=None):
    """为标题样式挂多级自动编号（1 / 1.1 / 1.1.1）。

    返回 abstractNumId。失败时返回 None，调用方应退化为“手写编号”。
    """
    try:
        numbering_part = doc.part.numbering_part
    except Exception:
        return None
    try:
        from docx.oxml.ns import nsmap
        numbering = numbering_part.element

        # 计算新的 numId / abstractNumId
        existing_nums = numbering.findall(qn("w:num"))
        existing_abstract = numbering.findall(qn("w:abstractNum"))
        num_id = max([int(n.get(qn("w:numId"))) for n in existing_nums] or [0]) + 1
        abstract_id = max([int(a.get(qn("w:abstractNumId"))) for a in existing_abstract] or [0]) + 1

        abstract = OxmlElement("w:abstractNum")
        abstract.set(qn("w:abstractNumId"), str(abstract_id))

        multi = OxmlElement("w:multiLevelType")
        multi.set(qn("w:val"), "multilevel")
        abstract.append(multi)

        for idx, lvl in enumerate(levels):
            lvl_el = OxmlElement("w:lvl")
            lvl_el.set(qn("w:ilvl"), str(idx))
            start = OxmlElement("w:start")
            start.set(qn("w:val"), "1")
            lvl_el.append(start)
            fmt = OxmlElement("w:numFmt")
            fmt.set(qn("w:val"), "decimal")
            lvl_el.append(fmt)
            # 编号与标题文字之间的分隔（默认空一格）
            suff = OxmlElement("w:suff")
            suff.set(qn("w:val"), lvl.get("suff", "space"))
            lvl_el.append(suff)
            txt = OxmlElement("w:lvlText")
            txt.set(qn("w:val"), lvl.get("lvlText", f"%{idx + 1}"))
            lvl_el.append(txt)
            jc = OxmlElement("w:lvlJc")
            jc.set(qn("w:val"), "left")
            lvl_el.append(jc)
            pPr = OxmlElement("w:pPr")
            ind = OxmlElement("w:ind")
            ind.set(qn("w:left"), "0")
            ind.set(qn("w:firstLine"), "0")
            pPr.append(ind)
            lvl_el.append(pPr)
            abstract.append(lvl_el)

        numbering.insert(0, abstract)

        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(num_id))
        ref = OxmlElement("w:abstractNumId")
        ref.set(qn("w:val"), str(abstract_id))
        num.append(ref)
        numbering.append(num)
        return num_id
    except Exception:
        return None


def bind_style_to_numbering(doc, style_name, num_id, ilvl):
    """把某个段落样式绑定到多级编号的某一级。"""
    if num_id is None:
        return False
    try:
        style = doc.styles[style_name]
        pPr = style.element.get_or_add_pPr()
        for old in pPr.findall(qn("w:numPr")):
            pPr.remove(old)
        numPr = OxmlElement("w:numPr")
        ilvl_el = OxmlElement("w:ilvl")
        ilvl_el.set(qn("w:val"), str(ilvl))
        numPr.append(ilvl_el)
        numId_el = OxmlElement("w:numId")
        numId_el.set(qn("w:val"), str(num_id))
        numPr.append(numId_el)
        pPr.insert(0, numPr)
        return True
    except Exception:
        return False


def disable_paragraph_numbering(paragraph):
    """对单个段落关闭自动编号（w:numId=0），不影响其段落样式。"""
    try:
        pPr = paragraph._p.get_or_add_pPr()
        numPr = None
        try:
            numPr = pPr.get_or_add_numPr()
        except AttributeError:
            numPr = pPr.find(qn("w:numPr"))
            if numPr is None:
                numPr = OxmlElement("w:numPr")
                pStyle = pPr.find(qn("w:pStyle"))
                if pStyle is not None:
                    pStyle.addnext(numPr)
                else:
                    pPr.insert(0, numPr)
        for tag in ("w:ilvl", "w:numId"):
            for old in numPr.findall(qn(tag)):
                numPr.remove(old)
        ilvl = OxmlElement("w:ilvl")
        ilvl.set(qn("w:val"), "0")
        numPr.append(ilvl)
        num_id = OxmlElement("w:numId")
        num_id.set(qn("w:val"), "0")
        numPr.append(num_id)
        return True
    except Exception:
        return False
