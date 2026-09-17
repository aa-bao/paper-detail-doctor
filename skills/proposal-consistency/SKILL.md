---
name: proposal-consistency
description: 开题报告 ↔ 正文（初稿）一致性核对（只读，绝不写盘）。从开题报告抽有限个可指定锚点（题目/概念/理论/方法要素/大纲/预期结论/文献承诺），回正文找落点，按三层判定给出「缺失/改动/偏离/无法判定/一致」五类结论与可复核证据卡，并自带金标准跑分（召回率/误报率/定位准确率）。当用户说「帮我看看开题和正文对不对得上」「开题和论文一致吗」「初稿有没有偏离开题」「开题承诺的 X 正文做到了没」「答辩前核对开题」时使用。注意：这是双文档核对，必须同时给正文 docx 与开题报告 docx，单给正文会报错退出。
agent_created: true
metadata:
  parent_package: paper-detail-doctor
  version: "1.0.0"
  last_updated: "2026-09-17"
  status: active
  related_skills:
    - format-audit
    - cite-doctor
    - text-style
    - structure-length
---

# 开题报告 ↔ 正文 一致性核对（proposal-consistency）

> **这是 `paper-detail-doctor` 包的子 skill。** 它在包内的角色是**开题/提纲 ↔ 正文一致性核对**（设计方案 §5，对应 L4·26）。可由包根 `SKILL.md` 路由调用，也可独立用 CLI 跑，或由 `workflow/audit_all.py -p <开题.docx>` 串进全身体检。

**核心承诺（铁律，改不了）**：

1. **docx 是唯一真值源**；本 skill 在 audit 阶段**只读**，绝不写盘改文档。所有 issue 的 `action` 一律为 `null`——一致性问题只能人改，工具不代劳。
2. 与别的子 skill **不互相 import**，只通过 `shared/contracts/issue.schema.json` 的 issue 结构交接。
3. 共用代码只在 `shared/`（Scanner / report / ooxml_guard），不复制到本目录。

**方法论一句话**：不做"全文语义比对"（不可复核也说不清）。做法是**锚点法**——从开题抽有限个可指定锚点，回正文找落点，逐条给「两侧原文 + 判定 + 建议」的证据卡，让你自己能判。

---

## 管什么 / 不管什么

**管（本 skill 的职责）**

- 开题报告承诺的**可结构化项** ↔ 正文实际落点是否对得上：题目、核心概念/关键词、理论基础与学者、研究方法要素（方法/对象/样本量/地点）、章节大纲、预期结论/研究假设、文献承诺数量。
- 输出五类结论（缺失/改动/偏离/无法判定/一致）+ 定位 + 两侧原文 + 建议，交你终裁。
- 对 L2 语义等价项，**只摆证据、不下结论**，产出 `verdict=待判` 的证据卡交宿主 agent 判。

**不管（not_for，防误命中）**

- ❌ 不判**论证质量 / 观点对错**（只能查"开题说了、正文没做到"这类结构性问题，§5.6 第 1 条）。
- ❌ 不查**版式/字号/引注形态/标点/字数**（那是 format-audit / cite-doctor / text-style / structure-length 的活）。
- ❌ 不查**正文内部的逻辑/事实错误**，只查"开题 ↔ 正文"这一对关系。
- ❌ **不自动改文档**：所有 issue `action=null`，没有"一键修复"。需要改，走 `workflow/plan.py` → `apply.py`（但本 skill 的条目 action 为 null，plan/apply 不会动它们）。
- ❌ 不开题报告本身不完整（只有一段笼统描述）时强求：锚点抽不出来会退化（§5.6 第 3 条），会如实标 `无法判定`。

---

## 七类锚点（A1–A7）

每类都预先定义"怎么抽、怎么找、怎么判"。下表是唯一权威定义，脚本照此实现。
**表里说的是"哪一类锚点"；每篇论文的具体值（题眼、概念、理论家、方法要素、大纲锚点）
在锚点档案 `anchors.yaml` 里**，见下一节。

| 类 | 锚点 | 在开题哪里抽 | 在正文哪里找 | 判定层 | 落到结论 |
|---|---|---|---|---|---|
| A1 | 论文题目 | 封面题眼 | 封面题眼 | L1 字面 | 双向命中题眼 → 一致；否则 缺失/偏离 |
| A2 | 核心概念 / 关键词 | 概念界定、关键词 | 第 1 章概念界定、第 2 章框架 | L1 + **L2 语义** | 概念词双向命中作一致证据；**语义等价交宿主 agent（待判）** |
| A3 | 理论基础与学者 | 理论基础节 | 理论/框架章 | L1 + **L2 语义** | 理论清单摆齐；**语义等价交宿主 agent（待判）** |
| A4 | 研究方法要素（方法、对象、样本量、地点） | 研究方法节 | 研究方法节 | **L1 数字为主** | 方法/对象/地点一致 → 通过；样本量数字不等 → 改动；抽不出 → 无法判定 |
| A5 | 章节大纲 | 论文框架节（五、论文大纲） | 实际标题树（目录） | L3 结构 | 独立章并入 → 改动；正文新增开题没有的节 → 偏离；开题有正文无 → 缺失 |
| A6 | 预期结论 / 研究假设 | 研究意义、预期结论 | 结论章 | **L2 语义** | 摆齐两侧核心论断；**语义等价交宿主 agent（待判）** |
| A7 | 文献承诺（数量） | 参考文献节（"不少于 N 条"） | 文献表 | L1 数字 | 正文条数 ≥ 承诺 → 一致；否则 缺失 |

**三层判定（§5.2）**

- **L1 字面**：归一化（去空白/标点）后子串或数字严格匹配。全自动、结果确定可复核。数字不等 → **改动**。
- **L2 语义**：同一命题两处表述是否等价。**本脚本不含 LLM**——它只负责把两侧原文、判定问题、置信度槽位摆齐，输出 `verdict=待判` + `confidence=null`，**终裁交宿主 agent**（见下节协议）。（§5.2 原稿设想用 LLM 判，但只读脚本不内嵌模型，改为人-in-the-loop，意图与 §5.6 第 4 条一致：裁定权在人和导师。）
- **L3 结构**：大纲树 ↔ 实际标题树做覆盖矩阵。全自动。开题某节在正文无对应 → **缺失**。

---

## 锚点从哪来 —— 锚点档案 `anchors.yaml`（重要）

上表的**检查逻辑是通用的**，但"题眼是哪个词、点了哪些核心概念、引了哪些理论家、
方法要素有哪些、大纲里哪一章被并了节"——**这些是某一篇论文的数据，不是这个工具的知识**。
所以它们外置在**锚点档案**里，由使用者（或宿主 agent 读一遍开题后）填写，**不写死在代码里**。

档案位置按顺序查找，先命中者胜（模板见 `assets/anchors.example.yaml`）：

1. 命令行 `--anchors <file>`
2. 环境变量 `PDD_ANCHORS`
3. `<正文 docx 同目录>/.thesis-doctor/anchors.yaml`
4. `<包根>/assets/anchors.local.yaml`（`*.local.yaml` 已被 `.gitignore` 忽略）

档案里对应上表的七段：`A1_topic` / `A2_concept` / `A3_theory` / `A4_method` /
`A5_outline` / `A6_conclusion` / `A7_refs`。填法见那份模板，每段都有注释。

**未配置时怎么办 —— 写死的规矩**：

- **一律产出「无法判定」并写明原因**，交人工/宿主 agent。
- **绝不静默跳过、绝不判成「一致」**。静默跳过会被读成"这一项没问题"，
  正是本包最警惕的静默失效（"检查器没跑"伪装成"全部合格"）。
- 唯一的例外是 A7（文献承诺）与 A4 的样本量数字：它们是通用正则，不需要配置。
- A1 在 `A1_topic.tokens` 空缺时会退化为"从开题全文里找《…》"，仍找不到才报无法判定。

> ⚠️ 档案里的值必须与文档**逐字一致**（中文标点也要一致）。多数所谓"误报"其实是这里填错了词。

---

## 五类结论 → severity 映射（写死）

| 结论 | 定义 | severity | 是否产 issue |
|---|---|---|---|
| 一致 | L1/L2/L3 均通过 | （不产 issue，进通过清单 `meta.passed`） | 否 |
| 缺失 | 开题有、正文没有 | **error** | 是 |
| 改动 | 两边都有但不一致（六章→五章、30–40 人→15 人） | **warn** | 是 |
| 偏离 | 正文有、开题没有（新增节） | **warn** | 是 |
| 无法判定 | 定位失败或 L2 置信度低 | **info** | 是（但 `verdict=待判`，交人） |

⚠️ **所有 issue 的 `action` 一律为 `null`**——一致性问题绝不自动改文档。apply 不会碰它们。

---

## L2 语义判定协议（交回宿主 agent）

脚本对 A2 / A3 / A6 这三类 L2 项**不下结论**，而是产出一张"待判证据卡"，字段（见 `shared/contracts/issue.schema.json`）：

- `verdict: "待判"`
- `confidence: null`（脚本不估置信度，留给宿主 agent）
- `proposal_excerpt` / `thesis_excerpt`：两侧原文摘录（供你自判）
- `judgment_question`：交给你的判定问题，固定三选一格式——

> 等价 / 不等价 / 无法判定（置信度 < 0.7 一律归「无法判定」）

**宿主 agent（你）的裁定规则，写死不许含糊**：

1. 只许给 `等价` / `不等价` / `无法判定` 三种之一。
2. **置信度 < 0.7 一律归 `无法判定`**，转人工（或转导师）。**不许给"差不多一致""基本对得上"这种模糊结论**——那是 §5.2 明令禁止的。
3. 你判完把结论回写 issue 的 `verdict`（auto 由脚本定 / 待判 交你 / 待人工确认 已给结论终裁在导师），并填 `confidence`。

---

## 真实 CLI 用法

```bash
VPY="$HOME/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
SK="F:/Coding/Project/paper-detail-doctor/skills/proposal-consistency"

# 双文档核对（必需 -p）
"$VPY" "$SK/scripts/audit.py" \
  -d "F:/路径/初稿.docx" \
  -p "F:/路径/开题报告.docx" \
  [--anchors <锚点档案.yaml>] \
  [-o <产物目录>] [--json] [--print]

# 不给 --anchors 时依次找 PDD_ANCHORS、<正文目录>/.thesis-doctor/anchors.yaml、
# <包根>/assets/anchors.local.yaml；都没有则各锚点报「无法判定」交人工。
cp assets/anchors.example.yaml assets/anchors.local.yaml   # 起步方式

# 串进全身体检（audit_all 会把开题接进来并合并 issue）
"$VPY" "F:/Coding/Project/paper-detail-doctor/workflow/audit_all.py" \
  -d "F:/路径/初稿.docx" -p "F:/路径/开题报告.docx"
```

**缺 `-p` 的行为（重要）**：这是双文档核对，必须给开题。脚本在 `main()` 里检测 `not a.proposal` 会**立即写错误到 stderr 并以退出码 2 退出**，打印：

```
错误：开题报告—正文一致性核对是双文档核对，必须提供 -p/--proposal <开题报告.docx>。
用法：python audit.py -d <正文.docx> -p <开题报告.docx>
```

**绝不假装跑成功**（不会退回单文档模式、不会吐空报告当 0 问题）。

产物（写盘，但只读不改文档；路径用 `workflow/_common.audit_dir` 约定，落在文档旁边的 `.thesis-doctor/audit/proposal-consistency/`）：

- `audit.json`：机器用，keys = `meta / counts / issues / suppressed`。每条 issue 带 `anchor`(A1–A7)、`conclusion`、`rule`、`severity`、`proposal_excerpt`、`thesis_excerpt`、`verdict`、`confidence`、`judgment_question`。
- `audit-report.md`：人读证据卡（每张卡给两侧原文 + 期望/实际 + 判定问题 + 建议 + 裁定位）。

`meta.notes` 固定写明：① 只读未改文档、`action` 全 `null`；② **锚点档案来源**（或"未找到 → 各锚点将报无法判定"）；
③ L2 待判交宿主；④ 五类结论→severity 映射；⑤ 锚点未配置一律产「无法判定」、不静默跳过。
另在 `meta.anchors` 里给出档案的绝对路径（或 `null`）。

---

## 金标准跑分（自带评分，不靠感觉）

金标准固化在 `tests/golden/proposal-consistency.yaml`（人工逐条核对得来的 4 处真差异 + 8 条对齐锚点 + `locate` 定位关键词，**这是事实不是程序输出反推**）。
差异 id 形如 `A5-01`（前缀即锚点），`(anchor, type) → id` 的映射与定位关键词都由 yaml 推导，**不写死在测试代码里**——换论文只改 yaml。本 skill 的验收脚本 `tests/selftest_proposal_consistency.py` 跑真实双文档并当次从磁盘读回结果，报三个真实指标：

| 指标 | 含义 | 期望 |
|---|---|---|
| **召回率** | 4 处真差异被检出几个 | 越高越好（漏报最致命） |
| **误报率** | 报出的"差异"（缺失/改动/偏离）里金标准没有的占多少 | 越低越好（误报让你白费功夫） |
| **定位准确率** | 给出的正文位置（location.index）确实是对的那一处占多少 | 越高越好（给错位置等于没用） |

> 注意：`无法判定`（L2 待判）**不计入"差异"**，因此不算误报。误报率只统计 `缺失/改动/偏离` 三类结论里金标准没有的。

跑法：

```bash
"$VPY" "F:/Coding/Project/paper-detail-doctor/tests/selftest_proposal_consistency.py" \
  "F:/路径/初稿.docx" "F:/路径/开题报告.docx"
```

验收范式是**注入式反向验证**（证明检查器真在读内容、不是硬编码）：复制文档到临时目录，把开题样本量改成与正文一致的 15 人 → A4-03 差异应**消失**；改回 30–40 → 应**重现**。测试全程只在临时副本上操作，绝不改真实论文文件。

---

## §5.6 本方法的四条局限（先说清）

1. **只能查"开题说了、正文没做到"这类结构性问题，不能判断论证质量。**
2. **L2 语义判定有主观性**——所以强制带原文引用和置信度槽位，且低置信度（<0.7）一律转人工，不允许它替你下结论。（本脚本进一步把"LLM 判定"改为"宿主 agent 判定"，终裁权始终在人。）
3. **开题报告本身不完整时（只有一段笼统描述）**，锚点抽不出来，能力会退化。你这篇开题内容很完整（单表格、要素齐全），属有利情况。
4. **一致性的最终裁定权在你和导师**，工具只负责"把证据摆齐、把该问的问题列出来"。

---

## 文件

| 文件 | 作用 |
|---|---|
| `scripts/audit.py` | 只读核对主脚本（锚点抽取 + 三层判定 + 证据卡 + 金标准跑分逻辑） |
| `../../assets/anchors.example.yaml` | **锚点档案模板**：每篇论文不同的题眼/概念/理论/方法要素/大纲锚点填这里 |
| `../../assets/anchors.local.yaml` | 本机私有档案（已被 `.gitignore` 忽略，不入库） |
| `SKILL.md` | 本文件（路由/编排用） |

## 输出约定

产物落 `文档同级/.thesis-doctor/audit/proposal-consistency/`（与 `workflow/_common.audit_dir` 一致），不污染包目录。只读，不改任何 docx。
