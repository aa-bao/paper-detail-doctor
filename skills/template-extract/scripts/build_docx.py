#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""build_docx.py —— 依据 Format Spec 生成带「命名样式」的 Word 模板。

设计要点
--------
* 所有排版落到 **段落样式**（一级标题/正文文本/表格文本 …），不用逐段硬排版。
  这样在 WPS/Word 里改一次样式，全篇生效 —— 即“一键调整”。
* 标题样式基于 Heading 1/2/3，自带大纲级别，目录域才能抓到。
* 中文渲染依赖 w:eastAsia 字体槽，已由 common.set_east_asian_font 处理。

用法
----
    python build_docx.py -s spec.json -o out.docx [--no-numbering] [--no-demo]
"""
import argparse
import json
import os
import sys

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common as C  # noqa: E402
import clean_styles  # noqa: E402


# ------------------------------------------------------------------ 样式

def build_style(doc, s):
    """按 spec 定义一个段落样式，返回样式对象。"""
    name = s["display_name"]
    try:
        st = doc.styles[name]
    except KeyError:
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)

    base = s.get("base_style")
    if base:
        try:
            st.base_style = doc.styles[base]
        except KeyError:
            pass
    st.quick_style = True

    rpr = st.element.get_or_add_rPr()
    C.set_east_asian_font(rpr, s.get("cn_font"), s.get("en_font"))
    if s.get("size_pt"):
        st.font.size = Pt(float(s["size_pt"]))
    st.font.bold = bool(s.get("bold"))
    st.font.italic = bool(s.get("italic"))
    try:
        st.font.color.rgb = RGBColor.from_string(str(s.get("color", "000000")))
    except Exception:
        pass

    pf = st.paragraph_format
    pf.alignment = getattr(WD_ALIGN_PARAGRAPH, C.resolve_align(s.get("align")))

    if s.get("space_before_pt") is not None:
        pf.space_before = Pt(float(s["space_before_pt"]))
    if s.get("space_after_pt") is not None:
        pf.space_after = Pt(float(s["space_after_pt"]))

    ls = s.get("line_spacing")
    if ls:
        if s.get("line_spacing_type") == "exact":
            pf.line_spacing = Pt(float(ls))
        else:
            pf.line_spacing = float(ls)

    if s.get("keep_with_next"):
        pf.keep_with_next = True
    if s.get("keep_lines") or s.get("keep_together"):
        pf.keep_together = True

    pPr = st.element.get_or_add_pPr()
    C.set_first_line_indent_chars(pPr, s.get("first_line_indent_chars"))
    C.set_left_indent_chars(pPr, s.get("left_indent_chars"))
    # 标题样式自带大纲级别：目录域/导航窗格据此识别，无需挂在内置 Heading 上
    C.set_outline_level(pPr, s.get("outline_level"))
    return st


def build_all_styles(doc, spec):
    styles = {}
    for key, s in spec.get("styles", {}).items():
        styles[key] = build_style(doc, s)
    # 让 Normal 也带上默认中英文字体，避免回退到 Calibri
    fonts = spec.get("fonts", {})
    normal = doc.styles["Normal"]
    C.set_east_asian_font(normal.element.get_or_add_rPr(),
                          fonts.get("cn_default"), fonts.get("en_default"))
    if spec.get("styles", {}).get("body", {}).get("size_pt"):
        normal.font.size = Pt(float(spec["styles"]["body"]["size_pt"]))
    return styles


# ------------------------------------------------------------------ 页面 / 页眉页脚

def setup_page(section, page):
    section.page_width = Cm(float(page.get("width_cm", 21.0)))
    section.page_height = Cm(float(page.get("height_cm", 29.7)))
    section.top_margin = Cm(float(page.get("margin_top_cm", 2.0)))
    section.bottom_margin = Cm(float(page.get("margin_bottom_cm", 2.0)))
    section.left_margin = Cm(float(page.get("margin_left_cm", 2.5)))
    section.right_margin = Cm(float(page.get("margin_right_cm", 2.0)))
    if page.get("header_distance_cm"):
        section.header_distance = Cm(float(page["header_distance_cm"]))
    if page.get("footer_distance_cm"):
        section.footer_distance = Cm(float(page["footer_distance_cm"]))
    if page.get("gutter_cm"):
        section.gutter = Cm(float(page["gutter_cm"]))


def set_header_text(section, text, style_name, doc):
    section.header.is_linked_to_previous = False
    p = section.header.paragraphs[0]
    p.text = ""
    if text:
        run = p.add_run(text)
        try:
            p.style = doc.styles[style_name]
        except KeyError:
            pass
        return run
    return None


def set_footer_page_number(section, doc, style_name):
    section.footer.is_linked_to_previous = False
    p = section.footer.paragraphs[0]
    p.text = ""
    try:
        p.style = doc.styles[style_name]
    except KeyError:
        pass
    C.add_page_number_field(p)


# ------------------------------------------------------------------ 演示内容

def _p(doc, style_name, text=""):
    p = doc.add_paragraph(style=style_name)
    if text:
        p.add_run(text)
    return p


def _style_name(spec, key):
    """样式 key -> 样式显示名；允许直接传显示名。"""
    s = spec.get("styles", {}).get(key)
    return s["display_name"] if s else key


HEADING_KEYS = ("h1", "h2", "h3", "h4")


def add_table_block(doc, spec, block):
    tb = spec.get("table", {})
    caption = block.get("caption", "")
    if caption:
        _p(doc, "表题", caption)
    header = block.get("header", [])
    rows = block.get("rows", [])
    ncols = len(header) or (len(rows[0]) if rows else 1)
    table = doc.add_table(rows=1 + len(rows), cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    try:
        table.style = doc.styles["Table Grid"]
    except KeyError:
        pass

    for j, h in enumerate(header):
        cell = table.rows[0].cells[j]
        cell.text = ""
        para = cell.paragraphs[0]
        para.style = doc.styles["表格文本"]
        para.add_run(str(h))
        C.set_cell_valign(cell, tb.get("cell_valign", "center"))

    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            if j >= ncols:
                break
            cell = table.rows[i + 1].cells[j]
            cell.text = ""
            para = cell.paragraphs[0]
            para.style = doc.styles["表格文本"]
            para.add_run(str(v))
            C.set_cell_valign(cell, tb.get("cell_valign", "center"))

    C.apply_three_line_table(
        table,
        top_pt=float(tb.get("top_bottom_pt", 1.5)),
        bottom_pt=float(tb.get("top_bottom_pt", 1.5)),
        middle_pt=float(tb.get("middle_pt", 0.5)),
    )
    if tb.get("header_repeat"):
        C.set_repeat_header_row(table.rows[0])
    if block.get("source_note"):
        _p(doc, "图片来源", block["source_note"])
    return table


def build_demo(doc, spec, numbering_active=False):
    """生成演示骨架：封面 → 摘要/ABSTRACT/目录 → 正文各章 → 参考文献/致谢/附录。"""
    blocks = spec.get("demo", [])
    if not blocks:
        return
    s_title = _style_name(spec, "cover_title")
    s_body = _style_name(spec, "body")

    for b in blocks:
        kind = b.get("block")
        if not kind:
            key = b.get("style")
            text = b.get("text", "")
            if key in HEADING_KEYS and not numbering_active:
                fn = b.get("fallback_number")
                if fn:
                    text = f"{fn} {text}"
            if b.get("runs"):
                p = doc.add_paragraph(style=_style_name(spec, key))
                for r in b["runs"]:
                    run = p.add_run(r.get("text", ""))
                    if r.get("bold"):
                        run.bold = True
            else:
                _p(doc, _style_name(spec, key), text)
            continue

        if kind == "cover":
            p = _p(doc, s_title, b.get("text", ""))
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for line in b.get("lines", []):
                _p(doc, s_body, line)

        elif kind == "new_section":
            sec = doc.add_section(WD_SECTION.NEW_PAGE)
            setup_page(sec, spec.get("page", {}))
            pn = b.get("page_number", {})
            C.set_section_page_numbering(sec, fmt=pn.get("fmt"), start=pn.get("start"))
            if b.get("no_header"):
                C.set_section_no_header(sec)
                sec.footer.is_linked_to_previous = False
                sec.footer.paragraphs[0].text = ""
            if b.get("header") is not None:
                set_header_text(sec, b["header"], "页眉", doc)
            if b.get("page_number_footer", True):
                set_footer_page_number(sec, doc, "页码")

        elif kind == "pagebreak":
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

        elif kind == "toc":
            _p(doc, "目录标题", b.get("text", "目  录"))
            p = doc.add_paragraph()
            C.add_toc_field(p)

        elif kind == "table":
            add_table_block(doc, spec, b)

        elif kind == "figure":
            _p(doc, s_body, b.get("text", "（此处插入图片）"))
            _p(doc, "图题", b.get("caption", "图3-1 图题字体宋体、字号为五号"))
            if b.get("source_note"):
                _p(doc, "图片来源", b["source_note"])

        elif kind == "equation":
            _p(doc, "公式", b.get("text", "F = ma"))
            _p(doc, "公式", b.get("note", "式(1-1)说明了 XX 之间的关系。"))

        elif kind == "style_overview":
            build_style_overview(doc, spec)


def build_style_overview(doc, spec):
    """样式总览页：逐条列出样式名 + 示例，方便肉眼核对与一键调整。"""
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_page(sec, spec.get("page", {}))
    C.set_section_page_numbering(sec, fmt="decimal", start=1)
    set_header_text(sec, spec.get("styles", {}).get("header", {}).get("sample_text", ""), "页眉", doc)
    set_footer_page_number(sec, doc, "页码")

    _p(doc, "一级标题", "样式总览")
    # 这一页是辅助页，不参与章节自动编号
    C.disable_paragraph_numbering(doc.paragraphs[-1])
    intro = "下列每个样式均已在本文档中定义。在 WPS 中打开「样式」面板即可整体调整，全篇同步生效。"
    _p(doc, "正文文本", intro)
    for key, s in spec.get("styles", {}).items():
        name = s.get("display_name")
        if not name:
            continue
        size = s.get("size_pt")
        label = C.pt_to_cn_size(size) or ""
        desc = f"字体 {s.get('cn_font','-')} / {s.get('en_font','-')}　字号 {label or size}　加粗 {'是' if s.get('bold') else '否'}　对齐 {s.get('align','-')}"
        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(3)
        para.paragraph_format.space_after = Pt(0)
        r1 = para.add_run(f"【{name}】")
        r1.bold = True
        para.add_run("　" + desc)


# ------------------------------------------------------------------ 主流程

def main():
    ap = argparse.ArgumentParser(description="Format Spec -> Word 模板")
    ap.add_argument("-s", "--spec", required=True, help="Format Spec JSON 路径")
    ap.add_argument("-o", "--out", required=True, help="输出 .docx 路径")
    ap.add_argument("--no-numbering", action="store_true", help="不启用标题自动编号")
    ap.add_argument("--no-demo", action="store_true", help="只生成样式，不生成演示内容")
    ap.add_argument("--no-clean", action="store_true",
                    help="保留 python-docx 默认模板自带的未使用内置样式（默认会清理）")
    args = ap.parse_args()

    with open(args.spec, "r", encoding="utf-8") as f:
        spec = json.load(f)

    doc = Document()
    build_all_styles(doc, spec)

    page = spec.get("page", {})
    setup_page(doc.sections[0], page)

    # 标题自动编号
    numbering_note = "disabled"
    if not args.no_numbering and spec.get("numbering", {}).get("enabled"):
        levels = spec["numbering"]["levels"]
        num_id = C.ensure_heading_numbering(doc, levels)
        if num_id:
            styles = spec.get("styles", {})
            for key, ilvl in (("h1", 0), ("h2", 1), ("h3", 2), ("h4", 3)):
                name = styles.get(key, {}).get("display_name")
                if name and ilvl < len(levels):
                    C.bind_style_to_numbering(doc, name, num_id, ilvl)
            numbering_note = f"enabled(numId={num_id})"
        else:
            numbering_note = "failed-fallback"

    numbering_active = numbering_note.startswith("enabled")
    if not args.no_demo:
        build_demo(doc, spec, numbering_active=numbering_active)

    # 页眉/页脚段落改用本模板自己的样式：否则 Word 内置的 Header/Footer 会被
    # 当成"已使用"而清理不掉，样式库里就一直多两条无关项。
    retargeted = C.retarget_header_footer_styles(
        doc,
        header_style=_style_name(spec, "header"),
        footer_style=_style_name(spec, "page_number"),
    )

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    doc.save(args.out)

    # 清掉 python-docx 默认模板自带、本文档根本没用到的内置样式（164 个）
    clean_info = None
    if not args.no_clean:
        rep = clean_styles.clean_docx(
            args.out,
            keep_names=[s.get("display_name") for s in spec.get("styles", {}).values()],
            verbose=False,
        )
        clean_info = {
            "styles_before": rep["styles_before"],
            "styles_after": rep["styles_after"],
            "styles_removed": rep["styles_removed"],
            "numbering_removed": len(rep["numbering"]["num_removed"])
                                 + len(rep["numbering"]["abstractNum_removed"]),
            "kept": rep["kept"],
        }

    print(json.dumps({
        "ok": True,
        "out": os.path.abspath(args.out),
        "styles": len(spec.get("styles", {})),
        "numbering": numbering_note,
        "header_footer_retargeted": retargeted,
        "cleanup": clean_info,
        "fonts": spec.get("fonts", {}),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
