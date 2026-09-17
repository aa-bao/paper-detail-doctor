---
name: cite-doctor
description: 审计论文的 L2 引注与文献问题——引注应放在句读符前、应为上标方括号形态、Ctrl+单击能跳到参考文献表、编号连续且与正文首次出现顺序一致、参考文献正反双向覆盖（孤儿引用/未被引条目）、GB/T 7714 著录要素、重复条目。当用户问「引用序号放右上角/上标」「Ctrl 单击能跳转到参考文献」「参考文献顺序不对/要按正文出现顺序排」「有没有文献没被引/引了表里没有」「一条文献被引了两次」「著录格式不对/GB-T 7714」「引注要放在句号前」时路由到本 skill。只读，绝不写盘。
metadata:
  parent_package: paper-detail-doctor
  version: "1.0.0"
  last_updated: "2026-09-17"
  status: active
  related_skills:
    - template-extract
    - format-audit
    - text-style
    - structure-length
---

# 引注与文献审计（cite-doctor，L2）

> **这是 `paper-detail-doctor` 包的子 skill（L2 · 引注与文献）。** 它是**只读**审计：扫描 docx 的引注与参考文献表，按 7 条规则产出问题清单（每条带定位 + 期望/实际对照 + 标准来源），**绝不写盘**。部分问题带可执行的修复动作（上标、补书签、挂链接），由 `workflow/apply.py` 在用户勾选后落盘。

判定基准来自 `shared/lib/standards.py` 的三级优先：**模板 spec ＞ 配置 ＞ 默认**。没给 spec 时全部落 `default`，报告顶部会明写「未提供模板标准」。

---

## 这个 skill 管什么 / 不管什么

**管（7 条规则，覆盖 L2 的 7–13 项）**

- 引注位置：应落在句读符之前（cite-position）
- 引注形态：上标 + 方括号；不应显式写 `w:sz`（cite-form）
- 引注跳转：Ctrl+单击两个方向都要通（正文锚点 ↔ 文献表书签）（cite-jump）
- 编号体系：从 1 连续；顺序编码制还要求与正文首现顺序一致（cite-numbering）
- 文献覆盖：表里有没被引的、正文引了表里没有的（ref-coverage）
- 著录格式：GB/T 7714 纯文本可判定项（类型标识/年/期页码）（ref-format）
- 重复条目：同一条文献被著录多次（ref-duplicate）

**不管（反例，防误命中）**

- 「这条引语是不是受访者原话」→ `quote-verify`（引语逐字校验，质性研究）
- 摘要超页、字数超了 → `structure-length`
- 页眉横线、页码、目录 → `format-audit`
- 引号不统一、破折号太多、标题太虚 → `text-style`

---

## 检查项清单

下表 `rule id` 与源码 `ctx.add('...')` 的第一参数**逐字一致**。标准来源通过 `ctx.add` 里的映射表解析（`cite-position→cite_position`、`cite-form→cite_style`、`cite-jump→cite_style`、`cite-numbering→bib_system`、`ref-coverage→bib_system`、`ref-format→bib_standard`、`ref-duplicate→bib_standard`，均 template ＞ config ＞ default）。

| rule id | 判什么 | 严重度 | 能否自动修复 | 标准来源 |
|---|---|---|---|---|
| `cite-position` | 引注应落在句读符**之前**（……问题[1]。）；`after-punct`（紧跟标点后）要改，`end-of-paragraph`（段末且无句末标点）仅提示 | error / warn | `after-punct` 给动作 `cite.move_before_punct`；段末缺标点仅提示 | 模板（`cite_position`；若 spec 非 `before-punct`，本规则直接跳过不报） |
| `cite-form` | 形态：应为上标（`w:vertAlign=superscript`）+ 方括号；不支持的著录制（如 author-year）给 info 提示；上标不该再显式写 `w:sz`（渲染会缩两次） | info / error / warn | 上标与否给 `cite.superscript`；方括号给 `cite.rewrite`；显式字号给 `run.clear_size` | 模板（`cite_style`；著录制不被支持时 `source='template'` 显式标注） |
| `cite-jump` | 两个方向跳转完整性：① 文献表条目缺 `refN` 书签（error）；② 正文引注锚点悬空（指向不存在）/错指/纯文字无链接（error，部分给动作）；③ 一处多编号合并（warn，不修） | error / warn | 条目缺书签给 `ref.bookmark`；错指/无链接给 `cite.link`；悬空与多编号仅提示 | 模板/默认（`cite_style`；跳转完整性属通用） |
| `cite-numbering` | 编号从 [1] 起连续、不重复；顺序编码制还要求「编号顺序 = 正文首次出现顺序」；缺编号/断号 error，首现顺序不符 warn | error / warn | 仅提示（重编号属全局手术，工具不自动做） | 模板（`bib_system`） |
| `ref-coverage` | 正反覆盖：表里有、正文从未引（warn）；正文引了、表里没有（error，硬伤） | warn / error | 仅提示 | 模板（`bib_system`） |
| `ref-format` | GB/T 7714 纯文本可判定项：文献类型标识（[M]/[J]/[D]…）、出版年、[J] 的页码/文章号 | warn | 仅提示（只报不修——机器擅自补页码会造假） | 默认/模板（`bib_standard`） |
| `ref-duplicate` | 重复条目：两条开头 30 字完全相同 | error | 仅提示（删一条要连带把它后面所有引注重编号 + 重挂链接，属全局手术，不自动） | 默认/模板（`bib_standard`） |

> 另有内部规则 `audit-internal-error`（error，不可压制）：某条规则执行抛异常时升格为 error 级 Issue，提示「本次结果不可信，报给维护者」。

---

## 命令行用法

```bash
VPY="python"
SK="F:/Coding/Project/paper-detail-doctor/skills/cite-doctor"
```

真实参数（`audit.py --help` 原文）：

```
usage: audit.py [-h] -d DOCX [-s SPEC] [-i IGNORE] [-o OUT] [--print] [--json]

L2 引注与文献审计（只读）

options:
  -h, --help           show this help message and exit
  -d, --docx DOCX
  -s, --spec SPEC      不传则自动在文档旁找 spec.json
  -i, --ignore IGNORE  .thesisignore 路径
  -o, --out OUT        输出目录
  --print              把报告打到终端
  --json               打印 JSON
```

> 注意：cite-doctor **没有** `--optional` 这类开关；CLI 与 format-audit / structure-length 完全一致。

---

## 产物

写到文档旁的 `.thesis-doctor/audit/cite-doctor/` 下（源码 `main()` 里的真实文件名）：

| 文件 | 内容 |
|---|---|
| `audit.json` | 机器用：符合 `shared/contracts/issue.schema.json` 的 Issue 列表（部分带 `action` 字段，如 `cite.superscript`/`ref.bookmark`/`cite.link`）+ suppressed + summary + meta |
| `audit-report.md` | 人看：每条带定位、期望/实际对照、标准来源标签；有可修动作时标注「可自动修复：<动作名>」 |

---

## 典型调用示例

```bash
# ① 默认惯例体检（无 spec，顶部标注「未提供模板标准」）
"$VPY" "$SK/scripts/audit.py" -d "F:/论文/初稿.docx"

# ② 带模板 spec，打印报告（可看到上标/跳转/编号问题及对应动作名）
"$VPY" "$SK/scripts/audit.py" -d 初稿.docx -s spec.json --print

# ③ 落地到指定目录并打 JSON（供 workflow 聚合）
"$VPY" "$SK/scripts/audit.py" -d 初稿.docx -s spec.json -o out/ --json
```

预期要点：
- 终端打印 `[cite-doctor] <路径>` + 引注数/参考文献条数/书签数 + 标准来源 + 问题条数。
- 引注不是上标 → `cite-form` 报 error 并带动作 `cite.superscript`；文献表某条没挂书签 → `cite-jump` 报 error 带动作 `ref.bookmark`；正文引了 [15] 但表里没有 → `ref-coverage` 报 error（不可压制，答辩必被问）。
- 顺序编码制下编号与首现顺序不符 → `cite-numbering` 报 warn（不自动重编号）。

---

## 局限与注意（均来自源码注释与 meta）

1. **ref-duplicate 不给自动删除**：删一条要连带把它后面所有引注重编号 + 重挂链接，属全局手术，自动做等于替你改写一整套引用体系。
2. **顺序编码制「重编号」不给自动动作**：同上，编号与正文引注、链接三者必须同步改，留给人工。
3. **ref-format 只报不修**：著录改写要看原文献，机器擅自补页码/期号会造假。
4. **著者-出版年制（author-year）**：只给 info 提示「请人工核对」，不自动核对形态（代码只自动处理上标/方括号类）。
5. **多编号合并引用**：一处引注含多个编号时给 warn 且**不自动修复**（需人工拆成多个上标或确认分别跳转）。
6. **缺 spec 全部落 default**：顶部标注「未提供模板标准」，不会让人误以为「上标方括号」是学校规定。

---

## 与其他子 skill 的关系

- **只通过 `shared/contracts/issue.schema.json` 的 Issue 结构交接，不互相 import。** 本 skill 只 `import shared.lib.*`。
- 带 `action` 字段的 Issue（上标/补书签/挂链接）由 `workflow/apply.py` 按 `plan.yaml` 回 `audit.json` 取 `action.name` + `params` 执行；本 skill 只产出 Issue，不落盘。
- `workflow/audit_all.py` 读取 `.thesis-doctor/audit/cite-doctor/audit.json` 合成总报告。
- 标准来源统一由 `shared/lib/standards.py` 三级优先解析，报告每条标 `standard_source`（template/config/default）。
