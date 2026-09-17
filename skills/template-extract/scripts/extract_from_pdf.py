#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""extract_from_pdf.py —— 从 PDF 模板「半精确」提取格式。

PyMuPDF 能拿到每个文本 span 的字体名、字号、粗斜体标志与坐标，据此可以：
  * 统计字号分布 -> 推断正文基准字号
  * 比正文大的字号档 -> 推断各级标题
  * 用 bbox 判断居中 / 顶格

局限（务必在报告里如实标注）
  * 中文字体常被子集化内嵌，字体名形如 `ABCDEF+SimSun`，需要反查
  * 段前/段后磅值、缩进字符数**无法**从 PDF 稳定还原
  * 页边距可由文本块边界反推，但受页眉页脚干扰

用法
----
    python extract_from_pdf.py -p 模板.pdf -o spec_from_pdf.json [--max-pages 30]
"""
import argparse
import collections
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common as C  # noqa: E402


def _fitz():
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    return fitz


SUBSET_RE = re.compile(r"^[A-Z]{6}\+")


def clean_font(name):
    """去掉子集前缀，如 'ABCDEF+SimSun' -> 'SimSun'。"""
    return SUBSET_RE.sub("", name or "")


# 常见 PDF 内嵌中文字体名 -> 中文字体
FONT_ALIAS = {
    "simsun": "宋体", "nsimsun": "新宋体", "simhei": "黑体", "simkai": "楷体",
    "kaiti": "楷体", "fangsong": "仿宋", "microsoftyahei": "微软雅黑",
    "msyh": "微软雅黑", "stsong": "宋体", "sthei": "黑体", "stkaiti": "楷体",
    "songti": "宋体", "heitisc": "黑体",
}


def map_cn_font(raw):
    key = re.sub(r"[^a-z]", "", (raw or "").lower())
    for k, v in FONT_ALIAS.items():
        if k in key:
            return v
    return raw


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-p", "--pdf", required=True)
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--max-pages", type=int, default=30)
    args = ap.parse_args()

    fitz = _fitz()
    doc = fitz.open(args.pdf)

    size_chars = collections.Counter()       # 字号 -> 字符数
    size_fonts = collections.defaultdict(collections.Counter)
    size_bold = collections.defaultdict(collections.Counter)
    size_center = collections.defaultdict(collections.Counter)
    samples = collections.defaultdict(list)

    for pno in range(min(len(doc), args.max_pages)):
        page = doc[pno]
        pw = page.rect.width
        data = page.get_text("dict")
        for block in data.get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = (span.get("text") or "").strip()
                    if not text:
                        continue
                    size = round(float(span.get("size", 0)), 1)
                    if size <= 0:
                        continue
                    n = len(text)
                    flags = span.get("flags", 0)
                    bold = bool(flags & 2 ** 4) or bool(flags & 2 ** 5)
                    font = clean_font(span.get("font"))
                    bbox = span.get("bbox", [0, 0, 0, 0])
                    cx = (bbox[0] + bbox[2]) / 2
                    centered = abs(cx - pw / 2) < pw * 0.04

                    size_chars[size] += n
                    size_fonts[size][font] += n
                    size_bold[size][bold] += n
                    size_center[size][centered] += n
                    if len(samples[size]) < 3:
                        samples[size].append(text[:40])

    # 正文基准字号 = 出现字符数最多的字号
    body_size = size_chars.most_common(1)[0][0] if size_chars else None
    body_font_raw = size_fonts[body_size].most_common(1)[0][0] if body_size else None

    # 比正文大的档位，从大到小 -> h1/h2/h3
    bigger = sorted([s for s in size_chars if body_size and s > body_size + 0.4], reverse=True)

    def describe(size):
        fonts = size_fonts[size]
        bold_votes = size_bold[size]
        cen_votes = size_center[size]
        total = sum(fonts.values()) or 1
        dominant_font = fonts.most_common(1)[0][0] if fonts else None
        is_bold = bold_votes.get(True, 0) > total * 0.5
        is_center = cen_votes.get(True, 0) > total * 0.5
        return {
            "size_pt": size,
            "size_cn_label": C.pt_to_cn_size(size),
            "cn_font": map_cn_font(dominant_font),
            "raw_font": dominant_font,
            "bold": is_bold,
            "align": "center" if is_center else "left",
            "char_count": size_chars[size],
            "samples": samples[size],
        }

    role_map = {}
    if body_size is not None:
        role_map["body"] = describe(body_size)
    for idx, size in enumerate(bigger[:4]):
        role_map["h%d" % (idx + 1)] = describe(size)

    spec = {
        "meta": {
            "template_name": os.path.splitext(os.path.basename(args.pdf))[0],
            "source": {"path": os.path.abspath(args.pdf), "kind": "pdf",
                       "pages_scanned": min(len(doc), args.max_pages)},
            "extraction_path": "semi-lossless (PyMuPDF span 分析)",
            "overall_confidence": "medium",
        },
        "fonts": {
            "cn_default": map_cn_font(body_font_raw),
            "en_default": "Times New Roman",
        },
        "page": {},
        "styles": role_map,
        "size_histogram": [
            {"size_pt": s, "size_cn_label": C.pt_to_cn_size(s), "char_count": size_chars[s]}
            for s in sorted(size_chars, reverse=True)
        ],
        "limitations": [
            "段前/段后磅值、首行缩进字符数无法从 PDF 稳定还原",
            "中文字体若被子集化，字体名需人工核对",
            "页边距未自动推算，需人工确认或从规范文字读取",
        ],
        "open_questions": [
            {"item": "各级标题的段前/段后间距", "detail": "PDF 无法还原，需人工确认"},
            {"item": "正文首行缩进字符数", "detail": "PDF 无法还原，需人工确认"},
            {"item": "页边距与页眉页脚距离", "detail": "PDF 无法还原，需人工确认"},
            {"item": "表格边框磅值与样式", "detail": "需人工确认（三线表等）"},
        ],
    }

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)
    print(json.dumps({
        "ok": True, "out": os.path.abspath(args.out),
        "body_size_pt": body_size,
        "detected_roles": list(role_map.keys()),
        "size_levels": len(size_chars),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
