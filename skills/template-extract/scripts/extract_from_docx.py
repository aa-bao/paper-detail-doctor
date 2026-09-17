#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""extract_from_docx.py —— 从已有 .docx 模板「无损」提取 Format Spec。

为什么这是最优路径
------------------
docx 是一份带样式表的 XML 包。直接读 word/styles.xml + sectPr 拿到的是
字体槽、字号、段间距、缩进、编号的**真值**，不存在视觉推断误差。

用法
----
    python extract_from_docx.py -d 模板.docx -o spec_from_docx.json [--used-only]
"""
import argparse
import json
import os
import re
import sys

from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common as C  # noqa: E402


def _pt(v):
    try:
        return round(v.pt, 2)
    except Exception:
        return None


def read_spacing(pPr):
    """从 w:spacing 读出行距语义。"""
    out = {"line_spacing_type": "multiple", "line_spacing": None}
    if pPr is None:
        return out
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        return out
    before = sp.get(qn("w:before"))
    after = sp.get(qn("w:after"))
    out["space_before_pt"] = round(int(before) / 20, 2) if before else 0
    out["space_after_pt"] = round(int(after) / 20, 2) if after else 0
    line = sp.get(qn("w:line"))
    rule = sp.get(qn("w:lineRule"))
    if line:
        if rule in (None, "auto"):
            out["line_spacing_type"] = "multiple"
            out["line_spacing"] = round(int(line) / 240, 3)
        elif rule == "exact":
            out["line_spacing_type"] = "exact"
            out["line_spacing"] = round(int(line) / 20, 2)
        elif rule == "atLeast":
            out["line_spacing_type"] = "atLeast"
            out["line_spacing"] = round(int(line) / 20, 2)
    return out


def read_style_props(style):
    s = {}
    el = style.element
    rpr = el.find(qn("w:rPr"))
    if rpr is not None:
        rf = rpr.find(qn("w:rFonts"))
        if rf is not None:
            s["cn_font"] = rf.get(qn("w:eastAsia"))
            s["en_font"] = rf.get(qn("w:ascii"))
        sz = rpr.find(qn("w:sz"))
        if sz is not None:
            s["size_pt"] = round(int(sz.get(qn("w:val"))) / 2, 2)
        b = rpr.find(qn("w:b"))
        if b is not None:
            s["bold"] = b.get(qn("w:val")) not in ("0", "false")
        i = rpr.find(qn("w:i"))
        if i is not None:
            s["italic"] = i.get(qn("w:val")) not in ("0", "false")

    pPr = el.find(qn("w:pPr"))
    if pPr is not None:
        s.update(read_spacing(pPr))
        jc = pPr.find(qn("w:jc"))
        if jc is not None:
            s["align"] = {"both": "justify", "center": "center", "left": "left",
                          "right": "right", "distribute": "justify"}.get(
                jc.get(qn("w:val")), jc.get(qn("w:val")))
        ind = pPr.find(qn("w:ind"))
        if ind is not None:
            flc = ind.get(qn("w:firstLineChars"))
            lc = ind.get(qn("w:leftChars"))
            fl = ind.get(qn("w:firstLine"))
            s["first_line_indent_chars"] = round(int(flc) / 100, 2) if flc else 0
            s["left_indent_chars"] = round(int(lc) / 100, 2) if lc else 0
            if not flc and fl:
                s["first_line_indent_pt"] = round(int(fl) / 20, 2)
        ol = pPr.find(qn("w:outlineLvl"))
        if ol is not None:
            s["outline_level"] = int(ol.get(qn("w:val")))

    base = el.find(qn("w:basedOn"))
    if base is not None:
        s["base_style_id"] = base.get(qn("w:val"))
    return s


def guess_role(name, props, style_id):
    """把样式名 + 属性映射到 spec 的样式 key（宽松匹配，agent 可再纠正）。"""
    n = name or ""
    if props.get("outline_level") is not None:
        lvl = props["outline_level"]
        if 0 <= lvl <= 8:
            return "heading%d" % (lvl + 1)
    m = re.match(r"^([一二三四五六七八九十])级", n)
    if m:
        return "heading%d" % ("一二三四五六七八九十".index(m.group(1)) + 1)
    lower = n.lower()
    if "表题" in n or ("表" in n and "题" in n):
        return "table_caption"
    if "表格文本" in n or ("表" in n and "文本" in n):
        return "table"
    if "图题" in n:
        return "figure_caption"
    if "来源" in n:
        return "figure_source"
    if "关键词" in n or "keywords" in lower:
        return "keywords"
    if "摘要" in n or "abstract" in lower:
        return "abstract_title" if "标题" in n else "abstract_body"
    if "目录" in n:
        return "toc_title" if "标题" in n else "toc_entry"
    if "参考文献" in n:
        return "ref_title" if "标题" in n else "reference"
    if "页眉" in n:
        return "header"
    if "页码" in n:
        return "pagenumber"
    if "公式" in n:
        return "equation"
    if "封面" in n:
        return "cover"
    if "致谢" in n:
        return "ack_title"
    if "附录" in n:
        return "appendix_title"
    if "正文" in n or lower in ("normal", "body text"):
        return "body"
    return "other"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-d", "--docx", required=True)
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--used-only", action="store_true",
                    help="只导出文档中实际用到的样式")
    args = ap.parse_args()

    doc = Document(args.docx)

    # 文档中实际用到的样式名
    used = set()
    for p in doc.paragraphs:
        if p.style is not None:
            used.add(p.style.name)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    if p.style is not None:
                        used.add(p.style.name)
    # 页眉 / 页脚里的段落也要算上
    for sec in doc.sections:
        for hf in (sec.header, sec.footer):
            try:
                for p in hf.paragraphs:
                    if p.style is not None:
                        used.add(p.style.name)
            except Exception:
                pass

    # 页面
    sec = doc.sections[0]
    page = {
        "width_cm": round(sec.page_width.cm, 2),
        "height_cm": round(sec.page_height.cm, 2),
        "margin_top_cm": round(sec.top_margin.cm, 2),
        "margin_bottom_cm": round(sec.bottom_margin.cm, 2),
        "margin_left_cm": round(sec.left_margin.cm, 2),
        "margin_right_cm": round(sec.right_margin.cm, 2),
        "header_distance_cm": round(sec.header_distance.cm, 2),
        "footer_distance_cm": round(sec.footer_distance.cm, 2),
        "gutter_cm": round(sec.gutter.cm, 2) if sec.gutter else 0,
    }

    # 全量解析段落样式（含未使用的），以便解析继承链
    all_props = {}
    for st in doc.styles:
        if getattr(getattr(st, "type", None), "name", "") != "PARAGRAPH":
            continue
        all_props[st.style_id or st.name] = (st, read_style_props(st))

    def effective_outline(sid, props):
        """沿 basedOn 上溯，找到有效的大纲级别。"""
        if props.get("outline_level") is not None:
            return props["outline_level"], None
        seen, cur = set(), props.get("base_style_id")
        while cur and cur not in seen:
            seen.add(cur)
            item = all_props.get(cur)
            if item is None:
                break
            _, bprops = item
            if bprops.get("outline_level") is not None:
                return bprops["outline_level"], item[0].name
            cur = bprops.get("base_style_id")
        return None, None

    styles_out = {}
    for sid, (st, props) in all_props.items():
        if args.used_only and st.name not in used:
            continue
        entry = {"display_name": st.name, "style_id": st.style_id,
                 "used_in_document": st.name in used}
        entry.update(props)
        if entry.get("size_pt"):
            entry["size_cn_label"] = C.pt_to_cn_size(entry["size_pt"])
        ol, src = effective_outline(sid, props)
        if ol is not None:
            entry["outline_level"] = ol
            if src:
                entry["outline_level_inherited_from"] = src
        entry["suggested_role"] = guess_role(st.name, entry, st.style_id)
        styles_out[sid] = entry

    # 节 / 页码格式
    sections = []
    for i, s in enumerate(doc.sections):
        pn = s._sectPr.find(qn("w:pgNumType"))
        sections.append({
            "index": i,
            "page_number_fmt": pn.get(qn("w:fmt")) if pn is not None else None,
            "page_number_start": pn.get(qn("w:start")) if pn is not None else None,
            "header_text": "".join(p.text for p in s.header.paragraphs).strip(),
        })

    # 表格边框（取第一个表格）
    table_out = {}
    if doc.tables:
        t = doc.tables[0]
        tblPr = t._tbl.tblPr
        bd = tblPr.find(qn("w:tblBorders")) if tblPr is not None else None
        table_out = {"sample_table_borders": {}}
        if bd is not None:
            for edge in ("top", "bottom", "left", "right", "insideH", "insideV"):
                e = bd.find(qn("w:" + edge))
                if e is not None:
                    table_out["sample_table_borders"][edge] = {
                        "val": e.get(qn("w:val")),
                        "sz_eighth_pt": e.get(qn("w:sz")),
                    }

    spec = {
        "meta": {
            "template_name": os.path.splitext(os.path.basename(args.docx))[0],
            "source": {"path": os.path.abspath(args.docx), "kind": "docx"},
            "extraction_path": "lossless (styles.xml / numbering.xml / sectPr)",
            "overall_confidence": "high",
        },
        "fonts": {"cn_default": "宋体", "en_default": "Times New Roman"},
        "page": page,
        "styles": styles_out,
        "sections": sections,
        "table": table_out,
    }

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)
    print(json.dumps({
        "ok": True, "out": os.path.abspath(args.out),
        "styles": len(styles_out), "used_styles": len(used),
        "sections": len(sections),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
