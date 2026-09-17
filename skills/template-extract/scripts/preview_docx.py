#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""preview_docx.py —— 把 .docx 渲染成 PNG，供 agent 视觉自检。

链路：LibreOffice（headless）转 PDF → PyMuPDF 渲染为逐页 PNG。
这样 agent 不用真的打开 WPS/Word，也能核对排版效果。
"""
import argparse
import glob
import json
import os
import shutil
import subprocess
import sys
import uuid

SOFFICE_CANDIDATES = [
    r"F:\Dev\LibreOffice-26.2.3.2\program\soffice.exe",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    r"D:\LibreOffice\program\soffice.exe",
]


def find_soffice():
    for c in SOFFICE_CANDIDATES:
        if os.path.isfile(c):
            return c
    return shutil.which("soffice")


def convert_to_pdf(soffice, docx_path, outdir):
    profile = os.path.join(outdir, "_lo_profile_" + uuid.uuid4().hex[:8])
    cmd = [
        soffice, "--headless", "--norestore", "--invisible",
        f"-env:UserInstallation=file:///{profile.replace(os.sep, '/')}",
        "--convert-to", "pdf", "--outdir", outdir, docx_path,
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", timeout=180)
    shutil.rmtree(profile, ignore_errors=True)
    pdfs = glob.glob(os.path.join(outdir, "*.pdf"))
    for p in pdfs:
        if os.path.splitext(os.path.basename(p))[0] == os.path.splitext(os.path.basename(docx_path))[0]:
            return p
    return pdfs[0] if pdfs else None


def render(pdf_path, outdir, dpi=110):
    try:
        import pymupdf as fitz  # PyMuPDF >= 1.24 推荐写法
    except ImportError:
        import fitz  # 旧版兼容
    doc = fitz.open(pdf_path)
    pages = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        fp = os.path.join(outdir, f"preview_p{i + 1:02d}.png")
        pix.save(fp)
        pages.append(fp)
    doc.close()
    return pages


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-d", "--docx", required=True)
    ap.add_argument("-o", "--outdir", required=True)
    ap.add_argument("--dpi", type=int, default=110)
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    soffice = find_soffice()
    if not soffice:
        print(json.dumps({"ok": False, "error": "soffice not found"}, ensure_ascii=False))
        sys.exit(2)

    pdf = convert_to_pdf(soffice, os.path.abspath(args.docx), os.path.abspath(args.outdir))
    if not pdf:
        print(json.dumps({"ok": False, "error": "pdf conversion failed", "soffice": soffice}, ensure_ascii=False))
        sys.exit(3)

    pages = render(pdf, args.outdir, args.dpi)
    print(json.dumps({"ok": True, "pdf": pdf, "pages": pages, "count": len(pages)},
                     ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
