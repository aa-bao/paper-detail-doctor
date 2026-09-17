#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""clean_styles.py —— 清理 .docx 里「用不上」的样式与编号定义。

为什么需要它
------------
python-docx 的默认模板自带整套 Word 内置样式 **164 个**（其中 100 个是表格样式：
Light Shading Accent 1 … Colorful Grid Accent 6），生成出来的文件在 WPS/Word 的
「样式」面板里会列出一大堆与本文档毫无关系的样式，非常碍事。

本脚本做三件事：
1. **只留用得上的样式** —— 文档/页眉/页脚实际引用的 + 显式指定的（`--keep` 或 spec 里的
   样式名）+ 它们基于（w:basedOn）/链接（w:link）的基样式；
2. **顺带清掉没人引用的编号定义**（w:num / w:abstractNum）——否则它们指向的
   List Bullet 之类样式会被误判成"被引用"而留下来；
3. 把「只是为了继承才保留」的内置样式标记 `w:semiHidden`，不让它在样式库里露脸。

安全边界：默认样式（Normal / Default Paragraph Font / Normal Table / No List，
即 w:default="1" 那几个）永远保留；文档正文、页眉页脚、脚注尾注里引用到的样式永远保留。
所以清理只动"死样式"，不会改变任何排版效果。

用法
----
    python clean_styles.py -d in.docx [-o out.docx]
                           [--spec spec.json] [--keep "一级标题,正文文本"]
                           [--no-hide] [--drop-latent] [--report r.json]

不指定 -o 时**原地重写**。
"""
import argparse
import json
import os
import re
import sys
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def w(tag):
    return "{%s}%s" % (W, tag)


# 可能引用样式的"内容部件"
CONTENT_RE = re.compile(
    r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$")
# 引用样式的元素
STYLE_REF_TAGS = {w("pStyle"), w("rStyle"), w("tblStyle")}
NUMBERING_REF_TAGS = {w("numStyleLink"), w("styleLink"), w("pStyle")}
NUMBERING_PART = "word/numbering.xml"
STYLES_PART = "word/styles.xml"

# w:style 子元素的规范顺序（ECMA-376 CT_Style）——插 w:semiHidden 时必须守序，
# 否则 Word/WPS 可能报文档损坏。
STYLE_CHILD_ORDER = [
    "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
    "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked",
    "personal", "personalCompose", "personalReply", "rsid",
    "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr",
]


# ------------------------------------------------------------------ 工具

def _load(path):
    z = zipfile.ZipFile(path)
    try:
        infos = z.infolist()
        blobs = {i.filename: z.read(i.filename) for i in infos}
    finally:
        z.close()
    return infos, blobs


def _save(path, infos, blobs):
    """把改好的包写回 path。

    Windows 上 os.replace 需要目标文件的 DELETE 权限，而 WPS/Word 打开文档时
    只放行写入、不放行删除 —— 所以先试重命名，失败就退化成原地重写。
    """
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for i in infos:
            zi = zipfile.ZipInfo(i.filename, date_time=i.date_time)
            zi.compress_type = i.compress_type
            zi.external_attr = i.external_attr
            zi.internal_attr = i.internal_attr
            zi.create_system = i.create_system
            zo.writestr(zi, blobs.get(i.filename, b""))

    try:
        os.replace(tmp, path)
        return "replace"
    except PermissionError:
        pass

    # 目标被占用（多半是 WPS 正开着这份文档）：原地重写
    with open(tmp, "rb") as f:
        data = f.read()
    with open(path, "wb") as f:
        f.write(data)
    try:
        os.remove(tmp)
    except OSError:
        pass
    return "inplace"


def _style_name(el):
    nm = el.find(w("name"))
    return nm.get(w("val")) if nm is not None else None


def _add_in_order(parent, child, order=STYLE_CHILD_ORDER):
    """按 CT_Style 规范顺序插入子元素。"""
    tag = child.tag.split("}")[-1]
    idx = order.index(tag) if tag in order else len(order)
    for i, existing in enumerate(parent):
        etag = existing.tag.split("}")[-1]
        eidx = order.index(etag) if etag in order else len(order)
        if eidx > idx:
            existing.addprevious(child)
            return
    parent.append(child)


def _set_semi_hidden(el):
    """标记为半隐藏：样式库里不显示（但文档中被引用时仍然生效）。"""
    for tag in ("unhideWhenUsed", "qFormat"):
        for old in el.findall(w(tag)):
            el.remove(old)
    if el.find(w("semiHidden")) is None:
        _add_in_order(el, etree.Element(w("semiHidden")))


def _resolve_wanted(styles, names):
    """把「希望的样式名」解析成 styleId 集合。

    容忍 WPS 的后缀改写：WPS 保存时会把它自己内置样式同名的自定义样式
    改名成 `正文文本1`，这里一并认。
    """
    by_name = {}
    for sid, el in styles.items():
        nm = _style_name(el)
        if nm:
            by_name.setdefault(nm, sid)
    out = {}
    for raw in names:
        nm = str(raw).strip()
        if not nm:
            continue
        if nm in by_name:
            out[by_name[nm]] = by_name[nm]
            continue
        hit = None
        for cand, sid in by_name.items():
            if re.fullmatch(re.escape(nm) + r"\d+", cand or ""):
                hit = (cand, sid)
                break
        if hit:
            out[hit[1]] = hit[0]
    return out


# ------------------------------------------------------------------ 核心

def clean_docx(src, dst=None, keep_names=(), hide_builtins=True,
               drop_latent=False, verbose=True):
    dst = dst or src
    infos, blobs = _load(src)
    if STYLES_PART not in blobs:
        raise SystemExit("不是有效的 .docx：缺少 word/styles.xml")

    styles_root = etree.fromstring(blobs[STYLES_PART])
    style_els = {}
    style_order = []
    for el in styles_root.findall(w("style")):
        sid = el.get(w("styleId"))
        if sid is None:
            continue
        style_els[sid] = el
        style_order.append(sid)

    defaults = {sid for sid, el in style_els.items() if el.get(w("default")) == "1"}
    wanted = _resolve_wanted(style_els, keep_names)

    # ---- 内容部件里真正被引用的样式 / 编号
    content_refs = set()
    doc_numids = set()
    for name, data in blobs.items():
        if not CONTENT_RE.match(name):
            continue
        root = etree.fromstring(data)
        for el in root.iter():
            if el.tag in STYLE_REF_TAGS:
                v = el.get(w("val"))
                if v:
                    content_refs.add(v)
            elif el.tag == w("numId"):
                v = el.get(w("val"))
                if v and v != "0":
                    doc_numids.add(v)
    content_refs &= set(style_els)

    num_root = etree.fromstring(blobs[NUMBERING_PART]) if NUMBERING_PART in blobs else None

    def _numids_in_styles(sids):
        found = set()
        for sid in sids:
            el = style_els.get(sid)
            if el is None:
                continue
            for nid in el.iter(w("numId")):
                v = nid.get(w("val"))
                if v and v != "0":
                    found.add(v)
        return found

    # ---- 迭代收敛：样式保留集 <-> 编号保留集 互相依赖
    keep = set()
    num_keep = set()
    abs_keep = set()
    removed_nums, removed_abs = [], []
    for _ in range(6):
        keep = set(defaults) | set(wanted) | content_refs
        # 依赖闭包：basedOn / link
        changed = True
        while changed:
            changed = False
            for sid in list(keep):
                el = style_els.get(sid)
                if el is None:
                    continue
                for tag in ("basedOn", "link"):
                    ref = el.find(w(tag))
                    if ref is not None and ref.get(w("val")) in style_els \
                            and ref.get(w("val")) not in keep:
                        keep.add(ref.get(w("val")))
                        changed = True

        new_num_keep, new_abs_keep = _plan_numbering(
            num_root, doc_numids | _numids_in_styles(keep))
        new_refs = _numbering_style_refs(num_root, new_num_keep, new_abs_keep, style_els)
        if new_num_keep == num_keep and new_abs_keep == abs_keep and new_refs <= keep:
            num_keep, abs_keep = new_num_keep, new_abs_keep
            break
        num_keep, abs_keep = new_num_keep, new_abs_keep
        for sid in new_refs:
            if sid in style_els:
                keep.add(sid)

    # ---- 记录 + 执行删除
    removed = []
    for sid in style_order:
        if sid in keep:
            continue
        el = style_els[sid]
        removed.append((sid, _style_name(el), el.get(w("type"))))
        styles_root.remove(el)
        style_els.pop(sid, None)

    # 被保留样式的悬挂引用清理（next/link/basedOn 指向已删样式会让 Word 犯嘀咕）
    for sid in keep:
        el = style_els.get(sid)
        if el is None:
            continue
        for tag in ("next", "link", "basedOn"):
            ref = el.find(w(tag))
            if ref is not None and ref.get(w("val")) not in keep:
                el.remove(ref)

    # ---- 半隐藏「只为继承而留」的内置样式
    hidden = []
    if hide_builtins:
        for sid in sorted(keep - defaults - set(wanted) - content_refs):
            el = style_els.get(sid)
            if el is None or el.find(w("semiHidden")) is not None:
                continue
            _set_semi_hidden(el)
            hidden.append(_style_name(el) or sid)

    if drop_latent:
        lat = styles_root.find(w("latentStyles"))
        if lat is not None:
            styles_root.remove(lat)

    if num_root is not None:
        for n in list(num_root.findall(w("num"))):
            if n.get(w("numId")) not in num_keep:
                removed_nums.append(n.get(w("numId")))
                num_root.remove(n)
        for a in list(num_root.findall(w("abstractNum"))):
            if a.get(w("abstractNumId")) not in abs_keep:
                removed_abs.append(a.get(w("abstractNumId")))
                num_root.remove(a)

    blobs[STYLES_PART] = etree.tostring(
        styles_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    if num_root is not None:
        blobs[NUMBERING_PART] = etree.tostring(
            num_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    if os.path.abspath(dst) != os.path.abspath(src):
        mode = _save(dst, infos, blobs)
    else:
        mode = _save(src, infos, blobs)

    report = {
        "docx": os.path.abspath(dst),
        "write_mode": mode,
        "styles_before": len(style_order),
        "styles_after": len(style_els),
        "styles_removed": len(removed),
        "removed": [{"styleId": s, "name": n, "type": t} for s, n, t in removed],
        "kept": sorted((_style_name(el) or sid) for sid, el in style_els.items()),
        "hidden_builtins": sorted(hidden),
        "numbering": {
            "num_removed": sorted(removed_nums, key=lambda x: int(x)),
            "abstractNum_removed": sorted(removed_abs, key=lambda x: int(x)),
            "num_kept": sorted(num_keep, key=lambda x: int(x)),
        },
    }
    if verbose:
        out = {
            "ok": True, "docx": report["docx"],
            "styles": f"{report['styles_before']} -> {report['styles_after']}",
            "removed": report["styles_removed"],
            "hidden": len(hidden),
            "numbering_removed": len(removed_nums) + len(removed_abs),
        }
        if mode == "inplace":
            out["warning"] = ("目标文件被 WPS/Word 占用，已原地重写；"
                              "请关闭并重新打开文档以生效")
        print(json.dumps(out, ensure_ascii=False, indent=2))
    return report


def _plan_numbering(num_root, used_numids):
    """算出该保留哪些 w:num / w:abstractNum。"""
    if num_root is None:
        return set(), set()
    num_to_abs = {}
    for n in num_root.findall(w("num")):
        ref = n.find(w("abstractNumId"))
        if ref is not None:
            num_to_abs[n.get(w("numId"))] = ref.get(w("val"))
    num_keep = {nid for nid in used_numids if nid in num_to_abs}

    abs_els = {}
    for a in num_root.findall(w("abstractNum")):
        abs_els[a.get(w("abstractNumId"))] = a

    abs_keep = {num_to_abs[n] for n in num_keep}
    # numStyleLink 链：编号定义可以"借用"某个编号样式所引用的 abstractNum
    for aid in list(abs_keep):
        a = abs_els.get(aid)
        if a is None:
            continue
        for link in a.iter(w("numStyleLink")):
            abs_keep.add(link.get(w("val")))
    return num_keep, abs_keep


def _numbering_style_refs(num_root, num_keep, abs_keep, style_els):
    """保留下来的编号定义里引用了哪些样式（这些样式必须留着）。"""
    refs = set()
    if num_root is None:
        return refs
    for a in num_root.findall(w("abstractNum")):
        if a.get(w("abstractNumId")) not in abs_keep:
            continue
        for tag in NUMBERING_REF_TAGS:
            for el in a.iter(w(tag)):
                v = el.get(w("val"))
                if v and v in style_els:
                    refs.add(v)
    # numStyleLink 指向的"编号样式"，再看它引用的 numId -> abstractNum
    for sid in list(refs):
        el = style_els.get(sid)
        if el is None:
            continue
        for nid in el.iter(w("numId")):
            v = nid.get(w("val"))
            if v and v != "0":
                num_keep.add(v)
    return refs


# ------------------------------------------------------------------ CLI

def main():
    ap = argparse.ArgumentParser(description="清理 docx 里未使用的样式与编号定义")
    ap.add_argument("-d", "--docx", required=True, help="输入 .docx")
    ap.add_argument("-o", "--out", help="输出 .docx（默认原地重写）")
    ap.add_argument("-s", "--spec", help="Format Spec：取其 styles 里的全部样式名为白名单")
    ap.add_argument("--keep", default="", help="额外要保留的样式名，逗号分隔")
    ap.add_argument("--no-hide", action="store_true", help="不把保留的内置样式设为半隐藏")
    ap.add_argument("--drop-latent", action="store_true", help="同时删除 w:latentStyles")
    ap.add_argument("--report", help="把详细报告写到该 JSON 路径")
    args = ap.parse_args()

    names = []
    if args.spec:
        with open(args.spec, "r", encoding="utf-8") as f:
            spec = json.load(f)
        for s in spec.get("styles", {}).values():
            if s.get("display_name"):
                names.append(s["display_name"])
    if args.keep:
        names += [x for x in re.split(r"[,，]", args.keep) if x.strip()]

    rep = clean_docx(args.docx, args.out, keep_names=names,
                     hide_builtins=not args.no_hide,
                     drop_latent=args.drop_latent)
    if args.report:
        with open(args.report, "w", encoding="utf-8") as f:
            json.dump(rep, f, ensure_ascii=False, indent=2)
        print("report ->", os.path.abspath(args.report))


if __name__ == "__main__":
    main()
