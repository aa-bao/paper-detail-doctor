# paper-detail-doctor · 论文格式与细节优化 skill 包

> 面向**已经有一份 Word 论文**的人：把格式、排版、引注、文献、文字体例、结构篇幅、开题一致性上的问题诊断出来并修好，**且不破坏你已经手改过的内容**。

这是一个**包**，里面装 7 个专职子 skill；包根做**索引**与**工作流编排**。
包入口：`SKILL.md`　｜　设计方案：`docs/设计方案-v2.md`　｜　进度：`docs/路线图.md`
当前阶段：**骨架（v0.1.0）** —— 结构与契约已定，子 skill 逐个实现。

---

## 核心约定（先读这一条，它决定整套设计）

**docx 是唯一真值源。本包不会重新生成你的文档，只对现有 docx 做定点手术。**

| 通道 | 用途 | 频率 |
|---|---|---|
| **建通道**（Markdown → Word） | 从 0 搭初稿 | **一次性** |
| **改通道**（docx → docx） | 返修、改稿、格式修正 | **长期反复** |

> ⚠️ 一旦稿子有过人工修改，**不要再跑建通道** —— 它重建文档，会覆盖你的修改。

改通道三段式，**诊断与修改强制分离**：

```
audit（只读，绝不写盘） → plan（可编辑 yaml，你勾选） → apply（先快照 · 幂等 · 可回滚）
```

三条铁律：**幂等** · **最小侵入** · **可回滚**。

---

## 7 个专职子 skill

| 子 skill | 职责 | 覆盖 |
|---|---|---|
| `template-extract` | 模板/规范/范文 → Word 模板 + 格式 spec（体例规则） | 标准来源 |
| `format-audit` | 段落样式、页面设置、页码、页眉、字体字号行距、目录域 | L1 · 6 项 |
| `cite-doctor` | 引注位置/形态/跳转、编号制、正反覆盖、著录、去重 | L2 · 7 项 |
| `text-style` | 引号体例、标点异常、破折号密度、标题体例、过程痕迹 | L3 · 8 项 |
| `structure-length` | 章节编号、图表覆盖/编号/题注、摘要版式、双口径字数 | L4 · 4 项 |
| `proposal-consistency` | 开题/提纲 ↔ 正文一致性核对（锚点法 + 三层判定 + 跑分） | L4 · 1 项 |
| `quote-verify` | 引语逐字校验（不得拼接），质性研究可选 | 可选 |

**默认启用 23 项检查**（L3 里的中英空格 / 全半角 / 术语一致性 3 项默认关闭，信噪比低，按需打开）。

路由表、工作流与共享契约见 `SKILL.md`。

---

## 目录结构

```
paper-detail-doctor/
├── SKILL.md                    # 包入口：索引 + 工作流编排
├── skills/                     # 7 个专职子 skill，各自带 SKILL.md
├── shared/
│   ├── lib/                    # docx_ops / docx_scan / ooxml_guard / report
│   ├── contracts/              # Issue / PlanItem / Report schema
│   └── references/             # 通用惯例兜底 + GB/T 7714 要点
├── workflow/                   # audit_all / plan / apply / extract_standard
├── references/                 # 跨 skill 的 SOP
├── assets/                     # config.example.yaml + rules/
├── tests/golden/               # 金标准样本 + 跑分
└── docs/                       # 设计方案 + 路线图
```

---

## 环境

- Python 3.13（本机：`E:\Miniconda3\python.exe`；子 skill 各自的 venv 见其 SKILL.md）
- 依赖：`python-docx`、`lxml`、`pymupdf`；（PDF 导出 / 目录更新的域更新需 Windows + WPS，走 `win32com`）
- 路径、阈值、规则文件全部外置到 `assets/config.yaml`（样例见 `assets/config.example.yaml`）
