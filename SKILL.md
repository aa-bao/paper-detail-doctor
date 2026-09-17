---
name: paper-detail-doctor
description: "论文格式与细节优化的 skill 包（索引 + 工作流编排）。面向已经有一份 Word 论文的人，诊断并修好格式、版式、引注、文献、文字体例、结构篇幅、开题一致性等问题，且不破坏人工修改过的内容。内含 7 个专职子 skill：template-extract（模板/规范→Word 模板与格式 spec）、format-audit（L1 格式与版式）、cite-doctor（L2 引注与文献）、text-style（L3 文字体例）、structure-length（L4 结构与篇幅）、proposal-consistency（开题—正文一致性核对）、quote-verify（引语逐字校验，质性研究可选）。Triggers: 论文格式检查, 论文排版, 引注上标, 文献引用跳转, 参考文献顺序, 论文细节, 格式规范, 论文模板, 开题报告和正文一致吗, 论文体检, 帮我改论文格式, 目录页码, 页眉横线, 字数超出, 摘要超页, 破折号太多, 标题太虚, paper format check, thesis formatting, citation check, thesis template."
metadata:
  version: "0.1.0"
  last_updated: "2026-09-17"
  status: skeleton
  data_access_level: raw
  task_type: open-ended
  related_skills:
    - template-extract
    - format-audit
    - cite-doctor
    - text-style
    - structure-length
    - proposal-consistency
    - quote-verify
---

# paper-detail-doctor — 论文格式与细节优化（skill 包）

本文件是**包入口**：做两件事 —— **索引**（把请求路由到专职子 skill）和**工作流执行**（把子 skill 串成一次完整体检）。

它自己**不实现任何检查**，检查逻辑全在 `skills/` 下的子 skill 里。

---

## 0. 先读：本包的核心约定

> **docx 是唯一真值源。本包不会重新生成你的文档，只对现有 docx 做定点手术。**

| 通道 | 用途 | 频率 | 入口 |
|---|---|---|---|
| **建通道**（Markdown → Word） | 从 0 搭初稿 | **一次性** | 见 `skills/template-extract/` |
| **改通道**（docx → docx） | 返修、改稿、格式修正 | **长期反复** | 本包的 `workflow/` |

> ⚠️ **硬性提示**：一旦稿子有过人工修改，**不要再跑建通道**。它重建文档，会覆盖你的修改。

改通道永远是三段式，**诊断与修改强制分离**：

```
audit（只读，绝不写盘） → plan（可编辑 yaml，你勾选） → apply（先快照 · 幂等 · 可回滚）
```

三条铁律贯穿所有子 skill：**幂等** · **最小侵入** · **可回滚**。

---

## 1. 索引：7 个专职子 skill

| 子 skill | 职责 | 覆盖检查项 | 主要输入 | 状态 |
|---|---|---|---|---|
| **`template-extract`** | 从模板/格式规范/范文提取真实排版与**体例规则**，产出可套用的 Word 模板 + 格式 spec | 标准来源（所有检查的"期望值"） | PDF / docx / 截图 | 已有实现待并入 |
| **`format-audit`** | 格式与版式：段落样式合规、页面设置、页码方案、页眉、字体字号行距、目录域 | L1 · 1–6 | docx + spec | 待实现 |
| **`cite-doctor`** | 引注与文献：位置、形态（上标）、Ctrl+点击跳转、编号制、正反覆盖、GB/T 7714 著录、重复编号 | L2 · 7–13 | docx + spec | 待实现 |
| **`text-style`** | 文字体例：引号体例、标点异常、破折号密度、标题体例、过程痕迹（另有 3 项可选：中英空格/全半角/术语一致性） | L3 · 14–21 | docx / md | 待实现 |
| **`structure-length`** | 结构与篇幅：章节编号连续性、图表覆盖与编号与题注、摘要/ABSTRACT 版式、双口径字数与分章配额 | L4 · 22–25 | docx + spec | 待实现 |
| **`proposal-consistency`** | 开题/提纲 ↔ 正文**一致性核对**：锚点法 + 三层判定 + 证据卡 + 金标准跑分 | L4 · 26 | 开题 + 正文 | 待实现 |
| **`quote-verify`** | 引语**逐字**校验（不得拼接、跨句须用「……」） | 可选（质性研究） | 语料库 + 正文 | 待实现 |

---

## 2. 路由：用户说什么 → 走哪个子 skill

**索引模式**（单点请求，直接进对应子 skill）：

| 用户的话 | 路由到 |
|---|---|
| "帮我提取这个模板 / 按这份规范排" / "论文格式规范" | `template-extract` |
| "页码不对 / 页眉没横线 / 目录没更新 / 样式不统一" | `format-audit` |
| "引用序号要上标 / 要能 Ctrl+点击跳转 / 参考文献顺序不对 / 有文献没被引用" | `cite-doctor` |
| "引号不统一 / 破折号太多 / 标题太虚 / 标点有问题 / 有写作痕迹没删" | `text-style` |
| "字数超了 / 摘要超一页 / 每章没图表 / 图表编号断了" | `structure-length` |
| "开题报告和正文对得上吗 / 有没有跑题" | `proposal-consistency` |
| "这条引语是不是受访者原话" | `quote-verify` |

**工作流模式**（"给我做一次完整体检" / "定稿前帮我查一遍"）→ 见下一节。

> 路由不确定时，默认进**工作流模式**的 `audit` 段（只读），报告出来后再决定改什么。**只读永远安全。**

---

## 3. 工作流：一次完整体检怎么跑

```
template-extract          ← ① 先拿标准（可选，没有就降级到默认规范并显著标注）
      │  template-spec.json
      ▼
audit_all（只读）          ← ② 依次调用 5 个审计型子 skill，汇总成一份报告
      │  audit-report.md + audit.json
      ▼
plan                       ← ③ 问题 → 待执行操作列表（可编辑 yaml），你勾选/排除
      │  plan.yaml
      ▼
apply（写盘）              ← ④ 先快照，逐条手术，幂等；输出 apply-log.json
      │
      ▼
  同一份 docx（你的修改完好无损）
```

**编排顺序是有原因的**：
- `template-extract` **必须最先**——它是所有检查项"期望值"的来源。缺了它，检查标准会退化成通用惯例（报告会标注 `标准来源: 默认`）。
- `audit` 段**纯只读**，可以在任何时刻安全重跑。
- `apply` 段**按 plan 执行**，你没有勾的项绝不会被改。

**分段执行**（推荐，安全）：

```
# ① 标准
python workflow/extract_standard.py --template 学校模板.pdf --out spec/template-spec.json
# ② 只读诊断
python workflow/audit_all.py --docx 我的论文.docx --spec spec/template-spec.json
# ③ 生成可编辑计划 → 打开 plan.yaml 勾选
python workflow/plan.py --audit _output/audit.json --out plan.yaml
# ④ 执行（先快照）
python workflow/apply.py --docx 我的论文.docx --plan plan.yaml
```

---

## 4. 共享契约（子 skill 之间交接用的数据结构）

放在 `shared/contracts/`，所有子 skill 必须遵循，这样 `workflow/audit_all.py` 才能把 7 份报告合成一份。

**Issue（一条问题）**

```yaml
id: cite-form-0042             # 稳定 ID，供 .thesisignore 引用
rule: cite-form                # 规则名
layer: L2                      # L1/L2/L3/L4
skill: cite-doctor             # 产出它的子 skill
severity: error                # error | warn | info
location:
  kind: paragraph              # paragraph | run | table | header | footer | toc
  index: 137                   # 段落序号
  excerpt: "……分配问题[1]。"    # 便于人工定位的片段
expected: "上标方括号，字号 0.65 倍，句读符前"
actual:   "行内方括号，句读符后"
standard_source: template      # template | config | default ← 必须标注判定依据来自哪
suggestion: "改为上标，前置到句号前"
suppressible: true             # 是否允许人工裁定保留
```

**PlanItem（一条待执行操作）**

```yaml
issue_id: cite-form-0042
action: render-superscript-citation
params: { ref: 1, target: ref1 }
enabled: true                  # 你在 plan.yaml 里改这个
reason: null                   # 关掉时建议写原因
```

**Report**：`audit-report.md`（给人看，每条带定位与期望/实际对照）+ `audit.json`（给机器用，符合上面的 Issue 结构）。

**人工裁定保留**：`.thesisignore`（放论文同目录或 `output/` 下），按 `id` 或 `rule` 级别标记"刻意保留"，`apply` 一律跳过。

---

## 5. 包结构

```
paper-detail-doctor/
├── SKILL.md                    # ← 本文件：索引 + 工作流编排
├── skills/                     # 7 个专职子 skill，每个自带 SKILL.md
│   ├── template-extract/
│   ├── format-audit/
│   ├── cite-doctor/
│   ├── text-style/
│   ├── structure-length/
│   ├── proposal-consistency/
│   └── quote-verify/
├── shared/
│   ├── lib/                    # docx_ops / docx_scan / ooxml_guard / report
│   ├── contracts/              # Issue / PlanItem / Report 的 schema
│   └── references/             # 通用惯例兜底清单、GB/T 7714 要点
├── workflow/                   # 编排入口：audit_all / plan / apply / extract_standard
├── references/                 # 跨 skill 的 SOP
├── assets/                     # config.example.yaml + rules/
├── tests/golden/               # 金标准样本 + 跑分
└── docs/                       # 设计方案 + 路线图
```

**为什么共享层不放在各个子 skill 里**：`docx_ops`（书签/超链接/上标/段落边框）、`ooxml_guard`（幂等守卫 + 快照回滚）这些是 7 个子 skill 都要用的。放 `shared/` 一份，避免复制漂移。

---

## 6. 标准来源：三级优先

```
① template-spec.json      ← 由 template-extract 产出，最高优先（"学校说了算"）
② config.yaml 的 overrides ← 你手动指定（"我导师另有要求"）
③ GB/T 7714 + 通用学术惯例 ← 兜底，报告中必须标注"标准来源：默认"
```

每一条 Issue 都必须写 `standard_source`。**不许静默用默认值假装合规。**

---

## 7. 反模式（不要这样做）

| 反模式 | 为什么错 |
|---|---|
| 稿子有人工修改后，为了"统一格式"重跑建通道 | 会整体覆盖人工修改。改通道只做定点手术 |
| 直接把 audit 结果套用到文档 | 必须先出 plan 让作者勾选。作者的意图 > 工具的规范 |
| 把阈值/体例判断硬编码进子 skill | 标准必须来自模板；硬编码换个学校就全错 |
| 只报"不符合规范"不给定位和期望/实际对照 | 无法复核、无法定位的报告等于没报 |
| 用默认规范但报告里不说明 | 会让人误以为"学校就是这么要求的" |
| 子 skill 各自复制一份 docx 操作库 | 必然漂移。统一走 `shared/lib/` |

---

## 8. 状态

**当前：骨架阶段（v0.1.0）。** 目录结构与契约已定，`skills/` 下 7 个子 skill 的代码与 SKILL.md 待逐个实现。

进度见 `docs/路线图.md`；设计依据见 `docs/设计方案-v2.md`。
