#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""probe_input.py —— 判定输入类型并给出推荐提取路径。

输入可以是：单个文件（.docx/.pdf/图片）或一个目录（多为图片序列）。

用法
----
    python probe_input.py -i <文件或目录> [-o probe.json]
"""
import argparse
import json
import os
import sys

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
DOC_EXT = {".docx", ".doc", ".pdf"}


def natural_key(name):
    import re
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def probe(path):
    result = {"input": os.path.abspath(path), "kind": None, "files": [], "notes": []}
    if os.path.isfile(path):
        ext = os.path.splitext(path)[1].lower()
        result["files"] = [os.path.abspath(path)]
        if ext == ".docx":
            result["kind"] = "docx"
            result["recommended"] = "extract_from_docx.py（无损路径，首选）"
        elif ext == ".pdf":
            result["kind"] = "pdf"
            result["recommended"] = "extract_from_pdf.py + agent 视觉复核"
        elif ext in IMAGE_EXT:
            result["kind"] = "image"
            result["recommended"] = "agent 视觉分析（Read 图片）"
        elif ext == ".doc":
            result["kind"] = "legacy-doc"
            result["recommended"] = "先用 LibreOffice/harness-anything 转 .docx，再走无损路径"
            result["notes"].append("旧版 .doc 需先转换")
        else:
            result["kind"] = "unknown"
        return result

    if os.path.isdir(path):
        names = sorted([n for n in os.listdir(path) if not n.startswith(".")], key=natural_key)
        imgs = [n for n in names if os.path.splitext(n)[1].lower() in IMAGE_EXT]
        docs = [n for n in names if os.path.splitext(n)[1].lower() in DOC_EXT]
        result["files"] = [os.path.abspath(os.path.join(path, n)) for n in names]
        if imgs and not docs:
            result["kind"] = "images"
            result["recommended"] = "agent 逐页视觉分析（Read 图片）+ 规范文字对照"
            result["notes"].append(f"共 {len(imgs)} 张图片")
        elif docs:
            result["kind"] = "directory-with-docs"
            result["recommended"] = "优先用其中的 .docx 走无损路径"
            result["notes"].append(f"发现文档 {len(docs)} 个、图片 {len(imgs)} 张")
        else:
            result["kind"] = "directory"
            result["notes"].append("未发现可识别的模板文件")
        return result

    result["kind"] = "missing"
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-i", "--input", required=True)
    ap.add_argument("-o", "--out")
    args = ap.parse_args()

    r = probe(args.input)
    text = json.dumps(r, ensure_ascii=False, indent=2)
    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text)
    print(text)


if __name__ == "__main__":
    main()
