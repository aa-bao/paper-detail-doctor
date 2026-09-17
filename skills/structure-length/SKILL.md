---
name: structure-length
description: 审计论文的 L4 结构与篇幅问题——章节编号连续性（同级连续/不跳级/编号层级与标题样式匹配）、图表编号与题注（编号自洽/题注位置/正文引用存在）、摘要与 ABSTRACT 字数上限与关键词、双口径字数与分章配额。当用户问「字数超了/要压到多少字」「摘要超一页/摘要太长」「每章都要有图表/我这章没图」「图表编号断了/图题表题位置不对」「章节编号乱了」「正文一共多少字」时路由到本 skill。只读，绝不写盘。
metadata:
  parent_package: paper-detail-doctor
  version: "1.0.0"
  last_updated: "2026-09-17"
  status: active
  related_skills:
    - template-extract
    - format-audit
    - cite-doctor
    - text-style
---

# 结构与篇幅审计（structure-length，L4）

> **这是 `paper-detail-doctor` 包的子 skill（L4 · 结构与篇幅）。** 它是**只读**审计：扫描 docx 的章节、图表、摘要与全文文本，按 4 条规则产出问题清单（每条带定位 + 期望/实际对照 + 标准来源），**全部 `action` 为 None（仅提示，不自动改）**，绝不写盘。

判定基准来自 `shared/lib/standards.py` 的三级优先：**模板 spec ＞ 配置 ＞ 默认**。

---

## 这个 skill 管什么 / 不管什么

**管（4 条规则，覆盖 L4 的 22–25 项）**

- 章节编号：同级连续、不跳级、编号层级与标题样式匹配（heading-sequence）
- 图表编号与题注：编号自洽（不缺失/不重复）、题注位置（上/下）、正文引用存在（figure-table）
- 摘要/ABSTRACT：中英文摘要字数上限、关键词个数与分隔符一致、摘要内不应有引注/图表引用（abstract-layout）
- 篇幅：双口径字数 + 分章配额（length）

**不管（反例，防误命中）**

- 开题报告与正文对得上吗 → `proposal-consistency`（L4·26，单独成 skill）
- 引注上标、参考文献顺序 → `cite-doctor`
- 页眉横线、页码、目录、字体字号 → `format-audit`
- 引号不统一、破折号太多、标题太虚 → `text-style`

---

## 检查项清单

下表 `rule id` 与源码 `ctx.add('...')` 的第一参数**逐字一致**。源码 `_RULE_SOURCE` 把 4 条规则都映射到 `default` 键，故多数项标准来源为 `default`；但 `abstract-layout` 字数上限、`figure-table` 题注位置、`length` 字数下限在 spec 提供对应字段时会标 `template`（通过显式 `source=...` 传入）。

| rule id | 判什么 | 严重度 | 能否自动修复 | 标准来源 |
|---|---|---|---|---|
| `heading-sequence` | 章节编号：① 编号层级（段数）与标题样式层级不匹配（error）；② 跳级——某编号的上级编号缺失（warn）；③ 同级序号不连续/重复（error）。"三级标题数上限"**不在本规则报** | error / warn | 仅提示（改章节编号属全局手术，工具不自动改） | 默认（通用） |
| `figure-table` | 图表题注位置（按 `caption_table_position`/`caption_figure_position`，默认表上图下；warn）；编号自洽：同章内图/表题编号缺失/重复（error）；正文引用「如图 1-1 所示」但无对应题注（warn）。**全篇无图时跳过图相关检查、不报缺图** | warn / error | 仅提示（移动题注/改编号属全局手术） | 模板（题注位置 `caption_*_position`；无则默认 below/above） |
| `abstract-layout` | 中英文摘要字数上限（`abstract_cn_max_chars` 默认 1000 / `abstract_en_max_words` 默认 500，超限 error、未超限 info）；关键词个数与分隔符一致（warn/info）；摘要内不应有引注 `[n]` / 图表引用（warn） | error / warn / info | 仅提示（不擅自删改摘要内容） | 模板（摘要上限字段；无则默认 1000/500，标 default） |
| `length` | 双口径字数：① 汉字数 = `[\u4e00-\u9fff]` 计数；② 计空格的全角字符数（CJK+全角标点+全角 forms，含 U+3000）；③ 不含空白的字符数。分「含参考文献与附录」和「仅正文」两套 + 按一级标题拆章占比。`body_min_chars` 设了才判下限 | info / warn | 仅提示（只读统计，供你把握篇幅） | 配置/默认（`body_min_chars`；默认 None 不判下限） |

> 另有内部规则 `audit-internal-error`（error，不可压制）：某条规则执行抛异常时升格为 error 级 Issue。

---

## 命令行用法

```bash
VPY="C:/Users/Tian/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
SK="F:/Coding/Project/paper-detail-doctor/skills/structure-length"
```

真实参数（`audit.py --help` 原文）：

```
usage: audit.py [-h] -d DOCX [-s SPEC] [-i IGNORE] [-o OUT] [--print] [--json]

L4 结构与篇幅审计（只读）

options:
  -h, --help           show this help message and exit
  -d, --docx DOCX
  -s, --spec SPEC      不传则自动在文档旁找 spec.json
  -i, --ignore IGNORE  .thesisignore 路径
  -o, --out OUT        输出目录
  --print              把报告打到终端
  --json               打印 JSON
```

> 本 skill **没有** `--optional` 开关；CLI 与 format-audit / cite-doctor 完全一致。

---

## 产物

写到文档旁的 `.thesis-doctor/audit/structure-length/` 下（源码 `main()` 里的真实文件名）：

| 文件 | 内容 |
|---|---|
| `audit.json` | 机器用：Issue 列表（所有 `action` 为 None）+ suppressed + summary + meta（含 `length` 统计块：双口径字数、按章拆分、下限来源） |
| `audit-report.md` | 人看：每条带定位、期望/实际对照、标准来源标签；字数统计表 |

---

## 典型调用示例

```bash
# ① 默认惯例体检（无 spec，摘要上限走默认 1000/500，字数下限不判）
"$VPY" "$SK/scripts/audit.py" -d "F:/论文/初稿.docx"

# ② 带模板 spec（摘要上限/题注位置由 spec 决定），打印报告
"$VPY" "$SK/scripts/audit.py" -d 初稿.docx -s spec.json --print

# ③ 落地指定目录
"$VPY" "$SK/scripts/audit.py" -d 初稿.docx -s spec.json -o out/
```

预期要点：
- 终端打印 `[structure-length] <路径>` + 段落数/表格数/图片数 + 标准来源 + 问题条数。
- 中文摘要汉字数 > 上限 → `abstract-layout` 报 error；未超限也给一条 info（体现「已统计」）。
- 章节编号断号（如 1.2 之后直接 1.4）→ `heading-sequence` 报 error；某二级下挂了 5 个三级 → 不在此报（归 text-style 的 `heading-style`）。
- `length` 一定产出一条 info 字数摘要（双口径 + 分章占比），供你判断篇幅。

---

## 局限与注意（均来自源码注释与 meta）

1. **三级标题数上限不在此报**：`std[max_h3_per_h2]` 由 `text-style` 的 `heading-style` 负责；本 skill 只在 `meta.notes` 写一句说明，避免同一件事报两遍。
2. **改章节编号 / 重排图表编号是全局手术**：`action` 一律 None，只在 `suggestion` 说明原因，交给人工。
3. **全篇无图时不报缺图**：`figure-table` 在图数为 0 时跳过图相关检查（不报「缺图」、不报「图编号断号」），只在 `meta.notes` 说明「全文无图」。
4. **摘要是否单独成页判不出**：docx 层面除非有分页符 `w:br w:type="page"`，否则判不出；`meta.notes` 写明「需导出 PDF 后人工核对」，不猜。
5. **摘要字数上限的诚实纪律**：上限**只认 spec 的 `abstract_cn_max_chars` / `abstract_en_max_words`**；没有就走默认 1000/500 并标 `standard_source='default'`。**绝不可用 template-extract 测出的范文实测字数当阈值**（否则所有比范文长一点的摘要都被误报）。
6. **body_min_chars 默认 None**：正文字数下限**不判**，需 spec 显式设置 `length.body_min_chars` 才会在 `length` 里报 warn。

---

## 与其他子 skill 的关系

- **只通过 `shared/contracts/issue.schema.json` 的 Issue 结构交接，不互相 import。** 本 skill 只 `import shared.lib.*`。
- 因所有 `action` 为 None，本 skill 的 Issue 在 `workflow/apply.py` 阶段不会被改；用户只能在报告里看、或用 `.thesisignore` 裁定保留。
- `workflow/audit_all.py` 读取 `.thesis-doctor/audit/structure-length/audit.json` 合成总报告。
- 与 `text-style` 的分工：三级标题数上限归 `text-style.heading-style`，章节编号连续性/跳级归本 skill 的 `heading-sequence`，两者通过 `meta.notes` 互相说明、不重复报。
- 标准来源统一由 `shared/lib/standards.py` 三级优先解析。
