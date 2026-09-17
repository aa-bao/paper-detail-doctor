# -*- coding: utf-8 -*-
"""init_template_dir.py —— 在项目根目录建立 .template/ 模板资产目录。

用法：
    python init_template_dir.py -p <项目根> [--spec spec.json] [--docx 模板.docx]
                              [--report 提取报告.md] [--source 输入文件...] [--force]

产出结构：
    <项目根>/.template/
    ├── manifest.json        # 状态清单（draft/confirmed/applied 状态机）
    ├── spec.json            # ★ 唯一格式真值（机器可读，可改可重新生成）
    ├── template.docx        # 由 spec 生成（派生物，可随时重建）
    ├── constraints.md       # ★ 写作约束摘要（人+机器双读，给下游写作 agent）
    ├── extraction_report.md # 提取报告（规则 + 来源 + 置信度 + 待确认项）
    ├── preview/             # 验收渲染图
    └── source/              # 输入证据存档（可选）

幂等：已存在的 spec.json / template.docx 不会被覆盖，除非 --force。
"""
import argparse
import datetime
import json
import os
import shutil
import sys

SCHEMA_VERSION = "1.0"

README = """# .template —— 论文模板资产目录

由 `paper-template-extractor` skill 维护。**spec.json 是唯一格式真值**，
template.docx 是它的派生物（改 spec 后重新生成；直接手改 docx 会在下次生成时被覆盖）。

| 文件 | 角色 | 可否手改 |
|---|---|---|
| `manifest.json` | 状态清单（draft → confirmed → applied） | agent 维护 |
| `spec.json` | ★ 唯一格式真值（Format Spec） | 可手改数值 |
| `template.docx` | 派生：命名样式模板（给 WPS/Word 用） | 会被重新生成覆盖 |
| `constraints.md` | ★ 写作约束摘要（给写作 agent / 人看） | 改完同步改 spec |
| `extraction_report.md` | 提取报告（规则来源 + 置信度 + 待确认项） | 追加 |
| `preview/` | 验收渲染图 | — |
| `source/` | 输入证据存档 | — |

重新生成模板 docx：

```bash
"$HOME/.workbuddy/binaries/python/envs/default/Scripts/python.exe" \\
  "$HOME/.workbuddy/skills/paper-template-extractor/scripts/build_docx.py" \\
  -s .template/spec.json -o .template/template.docx
```

校验（styles_with_diff 必须为 0，cleanliness.clean 必须为 true）：

```bash
"$VPY" "$SK/scripts/verify_docx.py" -d .template/template.docx -s .template/spec.json
```

> 多模板：如需"毕业论文版 + 期刊版"，在 manifest.json 的 `templates` 里
> 登记多个子目录（`.template/<profile>/`），`active` 指向当前生效的一个。
"""

CONSTRAINTS_HEADER = """# 写作排版约束（由 spec.json 派生）

> 本文件是 `.template/spec.json` 的人类可读摘要，供写作 / 排版环节直接消费。
> 两者不一致时，**以 spec.json 为准**；修改排版请改 spec 后重新生成本文件。

"""


def load_spec(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def gen_constraints_md(spec):
    """从 spec 派生 constraints.md（人+机器双读的写作约束摘要）。

    字段名与 references/format_spec_schema.md 定义的 spec schema 严格对齐：
    page.margin_*_cm / numbering.levels[].lvlText / styles.*.en_font /
    styles.*.line_spacing_type+line_spacing / sections[].page_number{fmt,start} /
    open_questions[].item+detail+confidence
    """
    L = [CONSTRAINTS_HEADER]

    meta = spec.get("meta", {}) or {}
    src = meta.get("source") or {}
    if meta.get("template_name"):
        kind_label = {"images": "%d 页图片" % src.get("pages", 0),
                      "pdf": "PDF", "docx": "Word 模板"}.get(src.get("kind"), src.get("kind", "?"))
        L.append("**模板**：%s（来源：%s，%s）\n" % (
            meta["template_name"], src.get("path", "?"), kind_label))

    # 页面
    page = spec.get("page", {}) or {}
    if page:
        L.append("## 页面\n")
        L.append("| 项 | 值 |\n|---|---|")
        L.append("| 纸张 | %s %.1f×%.1f cm %s" % (
            page.get("paper", "A4"), page.get("width_cm", 21), page.get("height_cm", 29.7),
            page.get("orientation", "portrait")))
        L.append("| 页边距 上/下/左/右 | %.1f / %.1f / %.1f / %.1f cm" % (
            page.get("margin_top_cm", 2.5), page.get("margin_bottom_cm", 2.5),
            page.get("margin_left_cm", 3), page.get("margin_right_cm", 2.5)))
        g = page.get("gutter_cm")
        if g:
            L.append("| 装订线 | %s，宽 %.1f cm" % (page.get("gutter_position", "left"), g))
        L.append("")

    # 标题编号方案
    num = spec.get("numbering", {}) or {}
    if num:
        L.append("## 标题编号方案（%s）\n" % num.get("scheme", "?"))
        for i, lvl in enumerate(num.get("levels", []), 1):
            suff = {"space": "空一格", "tab": "制表位", "nothing": "无"}.get(lvl.get("suff"), lvl.get("suff"))
            L.append("- **L%d**：`%s`（编号后%s）" % (i, lvl.get("lvlText", "?"), suff))
        ex = num.get("examples") or {}
        if ex:
            L.append("- 示例：" + "；".join("`%s`" % v for v in ex.values()))
        if num.get("alternatives"):
            L.append("- 备选方案：" + "、".join("`%s`" % a for a in num["alternatives"]))
        L.append("")

    # 样式速查
    styles = spec.get("styles", {}) or {}
    if styles:
        L.append("## 样式速查（精确参数见 spec.json）\n")
        L.append("| 样式 | 中文字体 | 西文字体 | 字号 | 加粗 | 对齐 | 行距 | 段前/段后(pt) | 首行缩进 |")
        L.append("|---|---|---|---|---|---|---|---|---|")
        for key, s in styles.items():
            if s.get("line_spacing_type") == "multiple":
                lh = "%g 倍" % s.get("line_spacing", 1)
            elif s.get("line_spacing_type") == "exact":
                lh = "%g 磅(固定)" % s.get("line_spacing", 0)
            else:
                lh = "单倍"
            indent = s.get("first_line_indent_chars")
            L.append("| %s | %s | %s | %s%s | %s | %s | %s | %s/%s | %s |" % (
                s.get("display_name", key),
                s.get("cn_font", "-"), s.get("en_font") or "-",
                s.get("size_cn_label") or "", ("（%gpt）" % s["size_pt"]) if s.get("size_pt") else "",
                "是" if s.get("bold") else "-",
                s.get("align", "-"), lh,
                s.get("space_before_pt", 0), s.get("space_after_pt", 0),
                ("%d 字符" % indent) if indent else "-"))
        L.append("")

    # 分节页码
    fmt_cn = {"lowerRoman": "小写罗马 i,ii", "upperRoman": "大写罗马 I,II", "decimal": "阿拉伯 1,2", "None": "无"}
    secs = spec.get("sections", []) or []
    if secs:
        L.append("## 分节页码方案\n")
        for i, s in enumerate(secs, 1):
            pn = s.get("page_number") or {}
            fmt = pn.get("fmt")
            L.append("- 节 %d（%s）：页码 %s（起始 %s）；页眉 %s" % (
                i, s.get("name", "?"),
                fmt_cn.get(fmt, fmt) if fmt else "无", pn.get("start", "-"),
                "「%s」" % s["header"] if s.get("header") else "无"))
        L.append("")

    # 表格
    tbl = spec.get("table", {}) or {}
    if tbl:
        L.append("## 表格\n")
        L.append("- 风格：%s（上下线 %s 磅、中线 %s 磅、%s竖线）" % (
            {"three-line": "三线表"}.get(tbl.get("border_style"), tbl.get("border_style")),
            tbl.get("top_bottom_pt", "?"), tbl.get("middle_pt", "?"),
            "无" if not tbl.get("vertical_lines") else "有"))
        L.append("- 表题：%s，编号 `%s`，表内文字 %s 对齐" % (
            {"above": "表上方居中"}.get(tbl.get("caption_position"), tbl.get("caption_position")),
            tbl.get("caption_numbering", "?"),
            {"center": "水平居中"}.get(tbl.get("cell_align"), tbl.get("cell_align"))))
        L.append("")

    # 参考文献
    refs = spec.get("references", {}) or {}
    if refs:
        L.append("## 参考文献\n")
        L.append("- 著录标准：%s（%s）；序号形式 `%s`，%s" % (
            refs.get("standard", "?"), refs.get("style", "?"),
            refs.get("number_format", "?"), refs.get("number_indent", "")))
        L.append("")

    # 待确认项
    oq = spec.get("open_questions") or []
    if oq:
        L.append("## ⚠️ 待确认项（manifest 为 draft 时不得下游消费）\n")
        for i, q in enumerate(oq, 1):
            if isinstance(q, dict):
                L.append("%d. **%s**（%s）— %s" % (
                    i, q.get("item", "?"), q.get("confidence", "?"), q.get("detail", "")))
            else:
                L.append("%d. %s" % (i, q))
        L.append("")

    return "\n".join(L)


def copy_if_newer(src, dst, force=False):
    """复制文件；dst 已存在时除非 force 否则跳过。src==dst 时不动。返回是否复制。"""
    if not src or not os.path.isfile(src):
        return False
    if os.path.abspath(src) == os.path.abspath(dst):
        return False
    if os.path.isfile(dst) and not force:
        return False
    shutil.copyfile(src, dst)
    return True


def main():
    ap = argparse.ArgumentParser(description="建立项目 .template/ 目录")
    ap.add_argument("-p", "--project", default=".", help="项目根目录")
    ap.add_argument("--spec", help="Format Spec JSON（将复制为 .template/spec.json）")
    ap.add_argument("--docx", help="生成的模板 docx（将复制为 .template/template.docx）")
    ap.add_argument("--report", help="提取报告 md（将复制为 .template/extraction_report.md）")
    ap.add_argument("--source", nargs="*", default=[], help="输入证据文件（存入 source/）")
    ap.add_argument("--name", help="模板名称（登记进 manifest，如：本科毕业论文）")
    ap.add_argument("--set-status", choices=("draft", "confirmed", "applied"),
                    help="显式推进 manifest 状态（open_questions>0 时 confirmed 会被拒）")
    ap.add_argument("--force", action="store_true", help="覆盖已存在的 spec/docx")
    args = ap.parse_args()

    root = os.path.abspath(args.project)
    tdir = os.path.join(root, ".template")
    for sub in ("preview", "source"):
        os.makedirs(os.path.join(tdir, sub), exist_ok=True)
    with open(os.path.join(tdir, "README.md"), "w", encoding="utf-8", newline="\n") as f:
        f.write(README)

    # manifest：支持多模板登记
    mpath = os.path.join(tdir, "manifest.json")
    manifest = {}
    if os.path.isfile(mpath):
        with open(mpath, encoding="utf-8") as f:
            manifest = json.load(f)
    now = datetime.datetime.now().isoformat(timespec="seconds")
    manifest.setdefault("schema_version", SCHEMA_VERSION)
    manifest.setdefault("templates", {})
    key = args.name or "default"
    entry = manifest["templates"].get(key, {})
    entry.update({
        "status": entry.get("status", "draft"),   # draft -> confirmed -> applied
        "created": entry.get("created", now),
        "updated": now,
    })
    if args.spec:
        entry["spec_source"] = os.path.abspath(args.spec)
    manifest["templates"][key] = entry
    if "active" not in manifest:
        manifest["active"] = key

    copied = []

    # spec：唯一真值
    spec_path = os.path.join(tdir, "spec.json")
    if args.spec and copy_if_newer(args.spec, spec_path, args.force):
        copied.append("spec.json")

    spec = load_spec(spec_path) if os.path.isfile(spec_path) else None

    # constraints.md：从 spec 派生（spec 更新后总是重写，因为它本身是派生物）
    if spec:
        with open(os.path.join(tdir, "constraints.md"), "w", encoding="utf-8", newline="\n") as f:
            f.write(gen_constraints_md(spec))
        copied.append("constraints.md")

    # docx：派生物
    if args.docx and copy_if_newer(args.docx, os.path.join(tdir, "template.docx"), args.force):
        copied.append("template.docx")

    # 报告
    if args.report and copy_if_newer(args.report, os.path.join(tdir, "extraction_report.md"), args.force):
        copied.append("extraction_report.md")

    # 证据存档
    for s in args.source:
        if os.path.isfile(s):
            dst = os.path.join(tdir, "source", os.path.basename(s))
            if copy_if_newer(s, dst, args.force):
                copied.append("source/" + os.path.basename(s))

    # 待确认项数量写入 manifest（confirmed 的门槛）
    if spec:
        n = len(spec.get("open_questions") or [])
        entry["open_questions"] = n

    # 显式状态推进
    oq = (spec.get("open_questions") or []) if spec else []
    if args.set_status:
        if args.set_status in ("confirmed", "applied") and oq:
            print(json.dumps({
                "ok": False,
                "error": "spec 仍有 %d 个 open_questions，拒绝推进到 %s；先解决或清空它们" % (
                    len(oq), args.set_status),
            }, ensure_ascii=False))
            sys.exit(2)
        entry["status"] = args.set_status
    elif oq and entry["status"] in ("confirmed", "applied"):
        entry["status"] = "draft"   # 出现未决问题时退回 draft

    with open(mpath, "w", encoding="utf-8", newline="\n") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    copied.append("manifest.json")

    print(json.dumps({
        "ok": True,
        "template_dir": tdir,
        "status": manifest["templates"][key]["status"],
        "copied_or_updated": copied,
        "active": manifest.get("active"),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    sys.exit(main())
