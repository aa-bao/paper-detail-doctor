---
name: format-audit
description: 审计论文的 L1 格式与版式问题——页面设置（纸张/页边距）、分节页码方案、页眉文字与横线、正文字体字号行距一致性、命名样式合规与直接格式硬改、目录域是否存在与失效。当用户问「页眉没有横线」「页码从哪开始/页码不对」「样式不统一/有硬改」「目录没更新/目录条目对不上」「字体字号行距不对」「帮我看看排版」时路由到本 skill。只读，绝不写盘。
metadata:
  parent_package: paper-detail-doctor
  version: "1.0.0"
  last_updated: "2026-09-17"
  status: active
  related_skills:
    - template-extract
    - cite-doctor
    - text-style
    - structure-length
---

# 格式与版式审计（format-audit，L1）

> **这是 `paper-detail-doctor` 包的子 skill（L1 · 格式与版式）。** 它是**只读**审计：扫描 docx，按 6 条规则产出问题清单（每条带定位 + 期望/实际对照 + 标准来源），**绝不写盘**。要修，请走包根的 `workflow/plan.py` → `workflow/apply.py`。

判定基准来自 `shared/lib/standards.py` 的三级优先：**模板 spec ＞ 配置 ＞ 默认**。没给 spec 时全部落 `default`，报告顶部会明写「未提供模板标准」——不会让人误以为「正文小四」是学校规定。

---

## 这个 skill 管什么 / 不管什么

**管（6 条规则，覆盖 L1 的 1–6 项）**

- 命名样式合规 + 直接格式硬改（style-compliance）
- 页面设置：纸张、页边距（page-setup）
- 分节页码方案：格式/起始、正文/前置是否缺页码（page-numbering）
- 页眉：文字内容、横线有无/一致性（header）
- 正文字体、字号、行距一致性（fonts）
- 目录域：是否存在/失效、条目数 vs 正文标题数（toc）

**不管（反例，防误命中）**

- 引注上标、参考文献顺序、著录格式 → `cite-doctor`
- 破折号太多、标题太虚、引号不统一 → `text-style`
- 章节编号、图表编号、摘要字数、正文字数 → `structure-length`
- 开题报告与正文一致性 → `proposal-consistency`

---

## 检查项清单

每条规则函数里标准来源通过 `source_key` 解析（template ＞ config ＞ default）；无 spec 时绝大多数落 `default`，报告顶部会标注「未提供模板标准」。下表 `rule id` 与源码 `ctx.add('...')` 的第一参数**逐字一致**。

| rule id | 判什么 | 严重度 | 能否自动修复 | 标准来源 |
|---|---|---|---|---|
| `style-compliance` | 命名样式合规 + 直接格式硬改：使用了模板未声明样式（仅 spec 提供时判）、正文误用章节标题样式（长正文）、正文 run 带直接 `w:sz`/`w:rFonts` | warn | 仅 `direct-fmt` 且含直接字号时给动作 `run.clear_size`；`style-unknown`/`heading-misuse` 仅提示 | 默认（命中 `source_key=body_size_pt`；模板提供 `body_size_pt` 时标模板） |
| `page-setup` | 纸张尺寸（默认 A4）；页边距**仅在 spec 提供 `page_margins_cm` 时才判** | error | 仅提示（不自动改分节页面设置，避免误伤其他节） | 模板（`page_size`/`page_margins_cm`；无 spec 落默认） |
| `page-numbering` | 逐节页码：正文节缺页码=error、前置节缺页码=warn；有标准时逐节比对格式；无标准时做自洽检查（前置用阿拉伯数字、别节都设了本节没设） | error / warn | 仅提示 | 模板（`page_numbering_scheme`；无则默认 + 自洽检查） |
| `header` | 各节页眉横线（有标准判有无、无标准报各节是否一致）、页眉文字（有标准比对、无标准报各节是否一致） | warn | `header-rule-missing` 在能定位页眉段落时给动作 `para.border`（补 bottom 边框）；其余仅提示 | 模板（`header_content`/`header_rule`；无则默认） |
| `fonts` | 同一样式内出现 2 种以上字号 / 中文字体 → warn；有标准时逐段比对字号 | warn | 仅提示 | 模板（`body_size_pt`/`body_font_cn`；无则默认） |
| `toc` | 目录域：完全无目录→info；TOC 域缺失/失效（「错误！未定义书签」）→error；目录条目数 < 正文标题数→warn | info / error / warn | 仅提示 | 默认（目录域应存在/不失效/条目数 = 标题数） |

> 另有内部规则 `audit-internal-error`（error，不可压制）：当某条规则执行抛异常时升格为 error 级 Issue，提示「本次结果不可信，报给维护者」。正常审计不会出现。

---

## 命令行用法

所有脚本用 venv 解释器绝对路径调用：

```bash
VPY="python"
SK="F:/Coding/Project/paper-detail-doctor/skills/format-audit"
```

真实参数（`audit.py --help` 原文）：

```
usage: audit.py [-h] -d DOCX [-s SPEC] [-i IGNORE] [-o OUT] [--print] [--json]

L1 格式与版式审计（只读）

options:
  -h, --help           show this help message and exit
  -d, --docx DOCX
  -s, --spec SPEC      不传则自动在文档旁找 spec.json
  -i, --ignore IGNORE  .thesisignore 路径
  -o, --out OUT        输出目录
  --print              把报告打到终端
  --json               打印 JSON
```

| 参数 | 含义 |
|---|---|
| `-d/--docx`（必填） | 目标 docx 路径 |
| `-s/--spec` | 模板/范文 spec.json；不传则自动在文档旁找（`shared/lib/standards.py` 的 `SPEC_FILENAMES`） |
| `-i/--ignore` | `.thesisignore` 路径（按 id 或 rule 裁定保留）；不传则自动在文档旁找 |
| `-o/--out` | 输出目录；不传默认 `<文档旁>/.thesis-doctor/audit/format-audit` |
| `--print` | 把 Markdown 报告打到终端 |
| `--json` | 把 audit.json 的 payload 打到终端 |

---

## 产物

写到文档旁的 `.thesis-doctor/audit/format-audit/` 下（源码 `main()` 里的真实文件名）：

| 文件 | 内容 |
|---|---|
| `audit.json` | 机器用：符合 `shared/contracts/issue.schema.json` 的 Issue 列表 + suppressed + summary + meta |
| `audit-report.md` | 人看：每条带定位、期望/实际对照、标准来源标签（模板✅/配置🔧/默认⚠️） |

这两个文件正是 `workflow/audit_all.py` 合成总报告时读取的入口。

---

## 典型调用示例

```bash
# ① 没有模板标准，纯默认惯例体检（报告顶部会标注「未提供模板标准」）
"$VPY" "$SK/scripts/audit.py" -d "F:/论文/初稿.docx"

# ② 带模板 spec，产物落默认目录，并直接打印报告
"$VPY" "$SK/scripts/audit.py" -d 初稿.docx -s spec.json --print

# ③ 指定输出目录，并打出 JSON（供 workflow 聚合）
"$VPY" "$SK/scripts/audit.py" -d 初稿.docx -s spec.json -o out/ --json
```

预期要点：
- 终端打印 `[format-audit] <路径>` + 段落数/分节数/引注数/参考文献数 + 标准来源 + 问题条数。
- 无 spec 时：`page-setup` 只报纸张（不报页边距）、`header` 只报各节横线/文字是否一致（不报「该不该有横线」）、`fonts`/`style-compliance` 落默认；报告 `meta.notes` 写明「因缺标准未判」的项。
- 有 spec 且页眉缺横线：`header` 报 `warn` 并带动作 `para.border`（可被 apply 自动补）。

---

## 局限与注意（均来自源码注释与 meta.notes）

1. **页边距**：`page_margins_cm` 为 None（无 spec）时不判页边距，只在 `meta.notes` 写「因缺标准未判」。页边距是「范文有值 ≠ 规范有规定」项，宁可少报不可瞎报。
2. **页眉横线**：`header_rule` 为 None 时不判「该不该有横线」，只报各节是否一致（None 视为「无横线」）。
3. **字体字号**：前 80 段实测 run 级若完全没有 `w:sz`，字号全部来自命名样式——必须沿样式链解析（`paragraph.style.font` + 递归 `base_style`），段落直接格式优先于样式。同一样式内出现 2 种以上字号/字体才报 warn。
4. **判不出来的如实说**：「目录是否已更新」「页眉与封面是否一致」这类需渲染后才能对比的项，一律写进 `meta.notes`（「需渲染确认，docx 层面判不出来」），绝不瞎报「目录是最新的」「页眉与封面一致」。
5. **不自动改分节页面设置**：`page-setup` 的纸张/页边距问题一律仅提示，避免误伤其他节。
6. **规则异常不静默**：单条规则炸了会被升格为 `audit-internal-error`（error，不可压制），不会伪装成「全部通过」。

---

## 与其他子 skill 的关系

- **只通过 `shared/contracts/issue.schema.json` 的 Issue 结构交接，不互相 import。** 各子 skill 只 `import shared.lib.*`，不直接引用别的子 skill。
- `workflow/audit_all.py` 按 `workflow/_common.py` 的 `ALL_SKILLS` 依次读取 `.thesis-doctor/audit/<skill>/audit.json` 合成一份总报告。
- 修问题的动作（如 `run.clear_size`、`para.border`）由 `workflow/apply.py` 按 `plan.yaml` 里回 `audit.json` 取到的 `action` 执行；本 skill 只产出 Issue + 可选 `action` 字段，不负责落盘。
- 标准来源统一由 `shared/lib/standards.py` 的三级优先解析，保证报告里每条都标 `standard_source`（template/config/default）。
