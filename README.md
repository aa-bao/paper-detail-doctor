# paper-detail-doctor · 论文格式与细节优化

> 面向**已经有一份 Word 论文**的人：把格式、排版、引注、体例、结构上的问题诊断出来并修好，**且不破坏你已经手改过的内容**。

当前阶段：**A1 骨架已建，功能待实现**（见 `docs/路线图.md`）
设计方案：`docs/设计方案-v2.md`（v1 见 `docs/设计方案-v1.md`，仅作对照）

---

## 核心约定（先读这一条，它决定整套设计）

**docx 是唯一真值源。** 本工具**不会重新生成你的文档**，只对现有 docx 做**定点手术**。

| 通道 | 用途 | 使用频率 |
|---|---|---|
| **建通道**（Markdown → Word，`scripts/build_from_md.py`） | 从 0 搭初稿 | **一次性** |
| **改通道**（docx → docx，`scripts/audit.py` / `plan.py` / `apply.py`） | 返修、改稿、格式修正 | **长期反复** |

> ⚠️ 一旦你的稿子有了人工修改痕迹，**不要再跑建通道** —— 它重建文档，会覆盖你的修改。

改通道三段式，**诊断与修改强制分离**：

```
audit（只读，绝不写盘） → plan（可编辑 yaml，你勾选） → apply（先快照 · 幂等 · 可回滚）
```

三条铁律：**幂等**（跑 N 次结果相同）· **最小侵入**（只动目标节点）· **可回滚**（写入前快照到 `_backup/`）。

---

## 目录结构

```
paper-detail-doctor/
├── SKILL.md                  # WorkBuddy skill 入口（待实现）
├── docs/                     # 设计方案与路线图
├── scripts/
│   ├── lib/                  # docx 底层：docx_ops / docx_scan / ooxml_guard / report
│   ├── checks/               # 检查项模块（L1 格式 / L2 引注 / L3 文字 / L4 结构 / 一致性）
│   ├── template/             # 模板提取（并入 paper-template-extractor）
│   ├── audit.py              # 只读总审计（待实现）
│   ├── plan.py               # 生成可编辑修复计划（待实现）
│   ├── apply.py              # 定点手术（待实现）
│   ├── build_from_md.py      # 建通道（待实现）
│   └── export.py             # WPS COM：更新目录域 / 导出 PDF（待实现）
├── references/               # SOP 与体例说明
├── assets/                   # config.example.yaml + rules/
├── tests/golden/             # 金标准样本 + 跑分
└── optional/quote_verify/    # 引语逐字校验（质性研究可选模块）
```

---

## 检查清单

**默认启用 23 项**，分四层：L1 格式与版式（6）· L2 引注与文献（7）· L3 文字体例（8）· L4 结构与篇幅（5，含开题—正文一致性核对）。

另有 **3 项可选**（中英空格 / 全半角 / 术语一致性），默认关闭 —— 这类检查在长文里动辄数百条告警，信噪比低，需要时按需打开。

完整清单与判定口径见 `docs/设计方案-v2.md` §4。

---

## 环境

- Python 3.13（本机：`E:\Miniconda3\python.exe`）
- 依赖：`python-docx`、`lxml`；（PDF 导出/目录更新需 Windows + WPS，走 `win32com`）
- 路径、阈值、规则文件全部外置到 `assets/config.yaml`（样例见 `assets/config.example.yaml`）
