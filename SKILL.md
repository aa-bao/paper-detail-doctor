---
name: paper-detail-doctor
description: "论文格式与细节优化的 skill 包（索引 + 工作流编排）。面向已经有一份 Word 论文的人，诊断并修好格式、版式、引注、文献、文字体例、结构篇幅、开题一致性等问题，且不破坏人工修改过的内容。内含 7 个专职子 skill：template-extract（模板/规范→Word 模板与格式 spec）、format-audit（L1 格式与版式）、cite-doctor（L2 引注与文献）、text-style（L3 文字体例）、structure-length（L4 结构与篇幅）、proposal-consistency（开题—正文一致性核对）、quote-verify（引语逐字校验，质性研究可选）。Triggers: 论文格式检查, 论文排版, 引注上标, 文献引用跳转, 参考文献顺序, 论文细节, 格式规范, 论文模板, 开题报告和正文一致吗, 论文体检, 帮我改论文格式, 目录页码, 页眉横线, 字数超出, 摘要超页, 破折号太多, 标题太虚, paper format check, thesis formatting, citation check, thesis template."
metadata:
  version: "0.3.0"
  last_updated: "2026-09-17"
  status: usable
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

**索引的唯一真源是 [`assets/skills-index.yaml`](assets/skills-index.yaml)。** 下面这张表只是给人看的摘要 —— 新增/改名/下线子 skill 时改 yaml，不要改这里。

| 子 skill | 职责 | 覆盖检查项 | 是否写盘 | 状态 |
|---|---|---|---|---|
| **`template-extract`** | 从模板/规范/范文提取真实排版与**体例规则**，产出 Word 模板 + Format Spec | 标准来源（所有检查的"期望值"） | 写 | ✅ 已实现 |
| **`format-audit`** | 段落样式、页面设置、页码方案、页眉文字与横线、字体字号行距、目录域 | L1 · 1–6（6 条规则） | 写 | ✅ 已实现 |
| **`cite-doctor`** | 引注位置与形态、Ctrl+点击跳转、编号制、正反覆盖、GB/T 7714 著录、去重 | L2 · 7–13（7 条规则） | 写 | ✅ 已实现 |
| **`text-style`** | 引号体例、标点异常、破折号密度、标题体例、过程痕迹（+3 项可选） | L3 · 14–21（5 + 3 条规则） | 写 | ✅ 已实现 |
| **`structure-length`** | 章节编号连续性、图表覆盖/编号/题注、摘要版式、双口径字数与配额 | L4 · 22–25（4 条规则） | 写 | ✅ 已实现 |
| **`proposal-consistency`** | 开题 ↔ 正文一致性核对（锚点法 + 三层判定 + 证据卡 + 金标准跑分） | L4 · 26（7 条规则） | **只读** | ✅ 已实现 |
| **`quote-verify`** | 引语逐字校验（不得拼接、跨句须用「……」），需自备语料库 | 可选（质性研究） | **只读** | ⬜ 未实现（可选模块） |

---

## 2. 路由：拿到 prompt 后怎么定位到子 skill

用户会用 `/paper-detail-doctor <prompt>` 触发。因为包内 skill 多、名字记不住，**路由由本文件负责**，用户只需要用自己的话描述需求。

### 第 0 步：先看有没有语义

| prompt 形态 | 行为 |
|---|---|
| 空 / `help` / `有哪些功能` / `你能干什么` | **打印目录**：逐个列出子 skill 的「中文名 — 一句话职责 — 你可以这样说：<触发语>」。不执行任何检查。 |
| 有具体请求 | 进入第 1 步 |

> 打印目录时按「标准来源 → L1 → L2 → L3 → L4 → 可选」的顺序，不要按字母序 —— 这个顺序本身就是使用顺序。

### 第 1 步：读索引

读 `assets/skills-index.yaml`。**不要凭子 skill 名字猜职责** —— 名字可能误导（例如 `cite-doctor` 也管参考文献著录格式，不只是引注）。

### 第 2 步：抽意图三元组

从 prompt 里抽出三样东西，缺哪样就只按有的匹配：

| 要素 | 问的问题 | 例 |
|---|---|---|
| **对象** | 论文的哪一部分？ | 引注 / 参考文献 / 页眉 / 目录 / 摘要 / 图表 / 字数 / 开题 |
| **动作** | 要我干什么？查（只读）/ 改（写盘）/ 答（解释规范） | "看看有没有问题" = 查 |
| **产物** | 要交什么？报告 / 改好的 docx / 一句话答案 | "给我个清单" = 报告 |

### 第 3 步：匹配 + 排除

拿三元组去比每个 skill 的 `triggers` 与 `covers` 打分，**再用 `not_for` 把误命中剔掉**。
`not_for` 是关键 —— 例如"这条引语是不是受访者原话"字面上像引注问题，实际属于 `quote-verify`。

### 第 4 步：决策

| 情况 | 怎么办 |
|---|---|
| **单一命中** | 直接执行。 |
| **命中多个、且有 `depends_on` 顺序** | 按序执行（`template-extract` 永远最先，它提供标准）。 |
| **命中多个、互不依赖** | 全部执行，但**先列出你要跑哪几个**；用户没反对就跑。 |
| **命中多个、但请求本身有歧义** | **用 AskUserQuestion 让用户选**，不要自己拍板。歧义典型：只说"帮我看看格式"——是查格式（`format-audit`）还是查全套（工作流）？ |
| **全部低分 / 无命中** | **明确说"本包不覆盖这个"**，并列出最接近的 1–2 个候选供选择。**不要挑一个勉强相关的来执行。** |
| **请求是整体性的**（"定稿前帮我查一遍" / "全身体检"） | 进 **工作流模式**（下一节）。 |

### 第 5 步：先声明，再执行

路由结果必须**明文写出来**，让用户能当场纠正：

```
→ 路由到 cite-doctor（L2 引注与文献）—— 你要的是引注上标＋可跳转，属引注形态与跳转检查
```

**不要静默路由。** 用户看到路由声明才能发现"我其实想说的是另一件事"。

### 第 6 步：执行前确认写盘边界

- `writes: false` 的 skill（`proposal-consistency`、`quote-verify`）→ 只出报告，**不改文档**。
- `writes: true` 的 skill → 走上图的**改通道三段式**：先 `audit`（只读）出报告，
  用户确认后再 `plan` → `apply`。**永远不要跳过 audit 直接改。**

### 典型 prompt 的路由结果（自检用）

| prompt | 路由 | 理由 |
|---|---|---|
| "引用序号放右上角，能 Ctrl 点着跳过去" | `cite-doctor` | 引注形态 + 跳转 |
| "参考文献的顺序好像不是按正文出现的" | `cite-doctor` | 编号制 |
| "我这章的破折号是不是太多了" | `text-style` | 破折号密度 |
| "第二章标题太玄乎了，改平实点" | `text-style` | 标题体例 |
| "页眉是不是该有根横线" | `format-audit` | 页眉横线属 L1 版式 |
| "字数超出要求了，帮我压到 3 万 5" | `structure-length` | 字数与配额 |
| "开题说要访谈 30 人，我最后只访了 15 人，会不会被问" | `proposal-consistency` | 方法要素一致性 |
| "这条引语是我受访者的原话吗" | `quote-verify` | 引语逐字（**不是** cite-doctor） |
| "学校给了份格式规范 PDF，帮我做成模板" | `template-extract` | 标准来源 |
| "定稿前帮我从头查一遍" | **工作流模式** | 整体请求 |
| "帮我写个摘要" | **不覆盖** | 本包只处理格式与细节，不生成内容（用 `academic-paper`） |
| "帮我查重" | **不覆盖** | 无此能力 |

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

**分段执行**（推荐，安全）。下面是与实现一致的命令（`<文档旁>/.thesis-doctor/` 是产物落点）：

```
# ① 标准（可选）。从"已排好版的范文 docx"反推体例规则，产出/并入 spec
python skills/template-extract/scripts/extract_conventions.py -d 范文.docx -o spec.json
#   也可以并进已有的 spec：  ... -d 范文.docx --merge-into spec/template-spec.json

# ② 只读诊断（五个审计型子 skill 汇总成一份报告）
python workflow/audit_all.py -d 我的论文.docx -s spec.json
#   要跑开题↔正文一致性核对，再给一份开题报告（不给就跳过并写进 notes）：
python workflow/audit_all.py -d 我的论文.docx -p 开题报告.docx
#   只跑某一个：             ... --only cite-doctor

# ③ 生成可编辑计划（产物写在文档旁的 .thesis-doctor/plan.yaml）→ 打开勾选
python workflow/plan.py -d 我的论文.docx
#   想让 warn/info 也默认勾上：  ... --all

# ④ 先预览（不写盘），确认后再执行
python workflow/apply.py -p "<文档旁>/.thesis-doctor/plan.yaml" --dry-run
python workflow/apply.py -p "<文档旁>/.thesis-doctor/plan.yaml" --yes

# 出问题就回滚（写入前会自动快照，保留最近 12 份）
python workflow/apply.py -d 我的论文.docx --list-snapshots
python workflow/apply.py -d 我的论文.docx --rollback
```

**产物落点约定**：审计报告固定在文档旁的 `.thesis-doctor/audit/`，计划与快照在同一 `.thesis-doctor/` 下。
这个目录是**项目数据**（含快照与幂等账本），不要当缓存删掉 —— 删了就没法回滚。

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
│   ├── lib/                    # docx_scan（读）/ docx_ops（写·13 个具名动作）/ ooxml_guard（三道闸门）/ report / standards（标准来源三级优先）
│   ├── contracts/              # Issue / PlanItem 的 JSON Schema（子 skill 之间只靠它交接）
│   └── references/             # 通用惯例兜底清单（预留，尚未填）
├── workflow/                   # 编排入口：audit_all / plan / apply（+ _common 共用工具）
├── references/                 # 跨 skill 的三份 SOP（见下）+ README 索引
├── assets/                     # config.example.yaml + skills-index.yaml（路由唯一真源）+ rules/（术语表等，尚未填）
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

**当前：6/7 个子 skill 已实现并通过验收（v0.3.0）。**

| 项 | 状态 |
|---|---|
| 包结构 / 索引 / 路由 / 共享契约 | ✅ 已定 |
| `template-extract` | ✅ 已并入；`extract_conventions.py` 已在真实论文上验证（8 项真值命中 7 项） |
| `format-audit` / `cite-doctor` / `text-style` / `structure-length` | ✅ 已实现，各带注入式自检 |
| `proposal-consistency` | ✅ 已实现；金标准跑分 **召回 4/4、误报 0/4、定位 4/4** |
| 验收 | ✅ **8 套测试 227 项全通过**（`python tests/run_all.py`） |
| `quote-verify` | ⬜ 未实现（可选模块，需自备语料库；→ 见下"命中未实现 skill 怎么办"） |
| `references/` 的 SOP | ✅ **三份已成文**：`SOP-引注上标与跳转.md` / `SOP-模板标准.md` / `SOP-开题一致性核对.md`（+ `references/README.md` 索引） |
| 安装到 `~/.workbuddy/skills/` | ✅ 只注册包入口（子 skill 留在包内，不单独注册） |

**命中未实现的 skill 怎么办**：索引里若保留了尚未实现的条目（当前是 `quote-verify`），
路由到它时必须**明说"该模块尚未实现"**并给出替代方案（引语逐字校验需自备语料库；
做文献研究的人用不上），**绝不假装执行、绝不吐空报告当"没问题"**。
索引里这类条目的 `status` 一律标 `planned`。

**一条已验证的纪律（源自踩坑）**：审计规则里任何异常都必须升格为 `audit-internal-error` 的 error 级 Issue，
**不许只写进 notes**。否则"检查器坏了"会伪装成"全部合格"——曾经就有一个规则因为
参数名撞名抛异常被吞掉，报出"0 条问题"而看起来一切正常。

进度见 `docs/路线图.md`；设计依据见 `docs/设计方案-v2.md`。
