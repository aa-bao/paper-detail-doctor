---
name: text-style
description: 审计论文的 L3 文字体例问题——引号体例（正文用直引号 " '，不用中文弯引号）、标点异常（连续/半角夹围/省略号误写/全半角括号混用）、破折号密度（每千字阈值 + 单段聚类）、标题体例（过长/冒号问号感叹号/虚词/末尾标点/二级下三级数上限）、过程痕迹残留（TODO/待补/（略）/【】等）；另有 3 项可选规则默认关闭。当用户问「引号不统一/中英文引号混用」「标点有问题/标点连用」「破折号太多」「标题太虚/标题不要冒号/标题平实一点」「有没有写作痕迹没删干净」「全角半角/中英文之间要不要空格」时路由到本 skill。只读，绝不写盘。
metadata:
  parent_package: paper-detail-doctor
  version: "1.0.0"
  last_updated: "2026-09-17"
  status: active
  related_skills:
    - template-extract
    - format-audit
    - cite-doctor
    - structure-length
---

# 文字体例审计（text-style，L3）

> **这是 `paper-detail-doctor` 包的子 skill（L3 · 文字体例）。** 它是**只读**审计：扫描 docx 正文（自动排除参考文献表/目录/页眉），按 5 条默认规则 + 3 条可选规则产出问题清单，**所有文字体例问题 `action` 一律为 None（仅提示，不自动改）**，绝不写盘。

判定基准来自 `shared/lib/standards.py` 的三级优先：**模板 spec ＞ 配置 ＞ 默认**。

---

## 这个 skill 管什么 / 不管什么

**管（5 条默认规则，覆盖 L3 的 14–21 项；另有 3 条可选）**

- 引号体例：正文用直引号，不用中文弯引号（quote-style）
- 标点异常：连续标点/半角标点被中文夹围/省略号误写/全半角括号混用（punctuation）
- 破折号密度：每千汉字出现次数 + 单段 ≥3 处（dash-density）
- 标题体例：过长/含冒号问号感叹号/含虚词/末尾带标点 + 每个二级下三级数上限（heading-style）
- 过程痕迹：TODO/TBD/FIXME/XXX/待补/待定/（略）/【】/？？？ 等（process-trace）
- 可选（默认关闭，`--optional` 才开）：cjk-spacing / fullwidth-halfwidth / term-consistency

**不管（反例，防误命中）**

- 章节编号断了、图表编号不对 → `structure-length`
- 引注上标、参考文献顺序 → `cite-doctor`
- 页眉横线、页码、目录、字体字号 → `format-audit`

---

## 检查项清单

下表 `rule id` 与源码 `ctx.add('...')` 的第一参数**逐字一致**。标准来源通过 `RULE_SRC` 映射解析（`quote-style→quote_style`、`dash-density→dash_per_1000_max`、`heading-style→heading_max_chars`/`max_h3_per_h2`），其余规则无 `source_key` → 落 `default`。

| rule id | 判什么 | 严重度 | 能否自动修复 | 标准来源 |
|---|---|---|---|---|
| `quote-style` | 正文用直引号 `"` `'`（U+0022/U+0027），不用中文弯引号 “ ” ‘ ’；**按段落聚合**（一段只报 1 条） | warn | 仅提示（弯引号改直引号需结合成对引号/嵌套上下文，机器盲目替换会改错） | 模板（`quote_style`；无则默认） |
| `punctuation` | 连续标点（。。 ，， ！！、跨类 ，。）、半角标点被中文夹围、省略号误写（`...` 或 `。。。）、全半角括号混用；一段合并 1 条 | warn | 仅提示（盲目替换可能破坏数字/英文里的合法半角标点） | 默认（通用惯例） |
| `dash-density` | 破折号密度：整体 > `dash_per_1000_max`（默认 6.0/千汉字）报 warn；单段 ≥3 处报 info | info / warn | 仅提示（删改会动语义） | 配置/默认（`dash_per_1000_max`；默认 6.0） |
| `heading-style` | 标题过长（> `heading_max_chars` 默认 25）、含冒号/问号/感叹号、含虚词（浅析/初探/试论/刍议/略论）、末尾带标点（逐标题 1 条）；每个二级标题下三级数 > `max_h3_per_h2`（默认 3） | warn | 仅提示（标题改写涉及作者意图） | 模板（`heading_max_chars` / `max_h3_per_h2`；无则默认） |
| `process-trace` | 过程痕迹残留：TODO/TBD/FIXME/XXX/待补/待定/待核实/此处省略/（略）/【】/？？？/@ @/<< />> ；一段合并 1 条 | warn | 仅提示（需先确认是否为占位待补内容） | 配置/默认 |
| `cjk-spacing`（可选） | 中英文 / 中文与数字之间空格是否统一（只报「不统一」，不报「该不该加」） | info | 仅提示 | 默认 |
| `fullwidth-halfwidth`（可选） | 全角数字/字母（０-９ａ-ｚＡ-Ｚ）与半角混用 | info | 仅提示 | 默认 |
| `term-consistency`（可选） | 术语一致性：同一概念多种写法混用；读 `assets/rules/term-map.yaml`，不存在则跳过并说明 | warn | 仅提示 | 配置（term-map） |

> 另有内部规则 `audit-internal-error`（error，不可压制）：某条规则执行抛异常时升格为 error 级 Issue。

---

## 命令行用法

```bash
VPY="python"
SK="F:/Coding/Project/paper-detail-doctor/skills/text-style"
```

真实参数（`audit.py --help` 原文）：

```
usage: audit.py [-h] -d DOCX [-s SPEC] [-i IGNORE] [-o OUT] [--optional]
                [--print] [--json]

L3 文字体例审计（只读）

options:
  -h, --help           show this help message and exit
  -d, --docx DOCX
  -s, --spec SPEC      不传则自动在文档旁找 spec.json
  -i, --ignore IGNORE  .thesisignore 路径
  -o, --out OUT        输出目录
  --optional           连同三条可选规则（cjk-spacing/fullwidth-halfwidth/term-
                       consistency）一起跑
  --print              把报告打到终端
  --json               打印 JSON
```

| 参数 | 含义 |
|---|---|
| `-d/--docx`（必填） | 目标 docx 路径 |
| `-s/--spec` | 模板/范文 spec.json；不传则自动在文档旁找 |
| `-i/--ignore` | `.thesisignore` 路径 |
| `-o/--out` | 输出目录；不传默认 `<文档旁>/.thesis-doctor/audit/text-style` |
| `--optional` | **本 skill 独有**：连同 3 条可选规则（cjk-spacing / fullwidth-halfwidth / term-consistency）一起跑；默认只跑 5 条 |
| `--print` / `--json` | 同其他子 skill |

---

## 产物

写到文档旁的 `.thesis-doctor/audit/text-style/` 下（源码 `main()` 里的真实文件名）：

| 文件 | 内容 |
|---|---|
| `audit.json` | 机器用：Issue 列表（所有 `action` 均为 None）+ suppressed + summary + meta（含 `optional` 开关状态） |
| `audit-report.md` | 人看：每条带定位、期望/实际对照、标准来源标签 |

---

## 典型调用示例

```bash
# ① 只跑 5 条默认规则（推荐，信噪比高）
"$VPY" "$SK/scripts/audit.py" -d "F:/论文/初稿.docx"

# ② 默认 + 3 条可选规则一起跑
"$VPY" "$SK/scripts/audit.py" -d 初稿.docx --optional --print

# ③ 带 spec，落地指定目录
"$VPY" "$SK/scripts/audit.py" -d 初稿.docx -s spec.json -o out/
```

预期要点：
- 终端打印 `[text-style] <路径>` + 段落数/排除范围（参考文献/目录/页眉）+ 标准来源 + 问题条数（含可选时标注「含可选规则」）。
- 一段里多处中文弯引号 → `quote-style` 只报 1 条 warn（按段聚合，不刷屏）。
- 不加 `--optional` 时 `cjk-spacing`/`fullwidth-halfwidth`/`term-consistency` **完全不跑**；`term-consistency` 若 `assets/rules/term-map.yaml` 不存在，会在 `meta.notes` 写「未找到映射表，跳过」。

---

## 局限与注意（均来自源码注释与 meta）

1. **全部文字体例不给自动修复**：`action` 一律 None。弯引号改直引号、标点修正、标题平实化都需结合上下文（成对引号、嵌套层级、作者意图），机器盲目替换会改错，故只在 `suggestion` 说明原因，`workflow/apply.py` 不会碰它们。
2. **quote-style 按段落聚合**：一段里 N 处弯引号只报 1 条，不每处一条（信噪比第一）。
3. **排除范围**：基于公开方法 `reference_section()` 与 `style_names()` 自行推出——参考文献表、目录（toc 样式）、页眉页脚天然排除，不碰 `shared/` 私有方法。
4. **三条可选规则默认关闭**：信噪比低（长文动辄数百条告警），需时用 `--optional` 打开；`term-consistency` 还需术语映射表 `assets/rules/term-map.yaml`，不存在则跳过并说明。
5. **来源去歧义**：引号/破折号/省略号一律 unicode 转义写死（`\u201c \u201d \u2018 \u2019 \u2014 \u2026`），不依赖肉眼区分 U+0022 与 U+201D；汉字计数用 `[\u4e00-\u9fff]`。

---

## 与其他子 skill 的关系

- **只通过 `shared/contracts/issue.schema.json` 的 Issue 结构交接，不互相 import。** 本 skill 只 `import shared.lib.*`。
- 因所有 `action` 为 None，本 skill 的 Issue 在 `workflow/apply.py` 阶段一律不会被改；用户只能在报告里看、或用 `.thesisignore` 裁定保留。
- `workflow/audit_all.py` 读取 `.thesis-doctor/audit/text-style/audit.json` 合成总报告。
- `heading-style` 与 `structure-length` 的 `heading-sequence` **不重复**：三级标题数上限由本 skill 的 `heading-style` 负责，`structure-length` 只在 `meta.notes` 写一句说明，避免同一件事报两遍。
- 标准来源统一由 `shared/lib/standards.py` 三级优先解析。
