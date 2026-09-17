# paper-detail-doctor · 论文格式与细节优化 skill 包

> 面向**已经有一份 Word 论文**的人：把格式、排版、引注、文献、文字体例、结构篇幅、开题一致性上的问题诊断出来并修好，**且不破坏你已经手改过的内容**。

这是一个**包**，里面装 7 个专职子 skill；包根做**索引**与**工作流编排**，自己**不实现任何检查**。

包入口：`SKILL.md`　｜　设计方案：`docs/设计方案-v2.md`　｜　进度：`docs/路线图.md`

**当前状态（v0.3.0）：6/7 个子 skill 已实现，8 套验收 219 项全通过。**

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

## 安装

```bash
python tools/install.py            # 安装 / 更新
python tools/install.py --check    # 只体检
python tools/install.py --uninstall
```

装完之后在任意项目里用 `/paper-detail-doctor <用你自己的话描述需求>` 触发，
包根会按 `assets/skills-index.yaml` 把请求路由到对应子 skill，**你不需要记子 skill 的名字**。

**为什么 skills 目录下只有一个"入口存根"**：宿主会*递归发现任意深度的 `SKILL.md`*。
如果整包拷进 `~/.workbuddy/skills/`，会同时注册 **1 个入口 + 7 个子 skill = 8 个条目**，
与"只记一个名字"的设计初衷相反。所以子 skill 的 `SKILL.md` 留在包内（包不在 skills 目录下
→ 宿主扫不到），skills 目录里只放一个**薄的入口存根**（记录包路径、先去读什么、三条硬约束、
几个入口命令），避免出现第二份真源。

---

## 快速开始

```bash
# ① 标准（可选）。从"已排好版的范文 docx"反推体例规则
python skills/template-extract/scripts/extract_conventions.py -d 范文.docx -o spec.json

# ② 只读诊断：五个审计型子 skill 汇总成一份报告
python workflow/audit_all.py -d 我的论文.docx -s spec.json
python workflow/audit_all.py -d 我的论文.docx -p 开题报告.docx   # 加 -p 才跑开题一致性核对
python workflow/audit_all.py -d 我的论文.docx --only cite-doctor  # 只跑一个

# ③ 生成可编辑计划（落在文档旁的 .thesis-doctor/plan.yaml）→ 打开勾选
python workflow/plan.py -d 我的论文.docx

# ④ 先预览不写盘，确认后执行
python workflow/apply.py -p "<文档旁>/.thesis-doctor/plan.yaml" --dry-run
python workflow/apply.py -p "<文档旁>/.thesis-doctor/plan.yaml" --yes

# 出问题就回滚（写入前自动快照，保留最近 12 份）
python workflow/apply.py -d 我的论文.docx --list-snapshots
python workflow/apply.py -d 我的论文.docx --rollback
```

产物固定落在**文档旁**的 `.thesis-doctor/`（审计报告在 `audit/`，计划与快照在同级）。
该目录含快照与幂等账本，**不要当缓存删除** —— 删了就无法回滚。

---

## 7 个专职子 skill

| 子 skill | 职责 | 覆盖 | 状态 |
|---|---|---|---|
| `template-extract` | 模板/规范/范文 → Word 模板 + 格式 spec（体例规则） | 标准来源 | ✅ |
| `format-audit` | 段落样式、页面设置、页码、页眉、字体字号行距、目录域 | L1 · 6 条规则 | ✅ |
| `cite-doctor` | 引注位置/形态/跳转、编号制、正反覆盖、著录、去重 | L2 · 7 条规则 | ✅ |
| `text-style` | 引号体例、标点异常、破折号密度、标题体例、过程痕迹 | L3 · 5 + 3 条可选 | ✅ |
| `structure-length` | 章节编号、图表覆盖/编号/题注、摘要版式、双口径字数 | L4 · 4 条规则 | ✅ |
| `proposal-consistency` | 开题 ↔ 正文一致性核对（锚点法 + 三层判定 + 证据卡 + 跑分） | L4 · 26 | ✅ |
| `quote-verify` | 引语逐字校验（不得拼接、跨句须用「……」），需自备语料库 | 可选（质性研究） | ⬜ |

设计清单 26 项（**23 项默认 + 3 项可选**：中英空格 / 全半角 / 术语一致性，信噪比低故默认关闭）；
落到代码里是 **29 条默认规则 + 3 条可选规则**（一个检查项可能由多条规则覆盖）。

路由表（含 `not_for` 反例与 `out_of_scope` 不覆盖清单）见 `assets/skills-index.yaml`；
工作流与共享契约见 `SKILL.md`。

---

## 目录结构

```
paper-detail-doctor/
├── SKILL.md                    # 包入口：索引 + 路由程序 + 工作流编排
├── tools/install.py            # 安装到 ~/.workbuddy/skills/（只注册包入口）
├── skills/                     # 7 个专职子 skill，各自带 SKILL.md
├── shared/
│   ├── lib/                    # docx_scan（读）/ docx_ops（写）/ ooxml_guard / report / standards
│   └── contracts/              # Issue / PlanItem 的 JSON Schema
├── workflow/                   # audit_all / plan / apply（+ _common）
├── references/                 # 跨 skill 的 SOP（尚未成文）
├── assets/                     # config.example.yaml + skills-index.yaml（路由唯一真源）+ rules/
├── tests/                      # 8 套验收（入口：tests/run_all.py）+ tests/golden/
└── docs/                       # 设计方案 + 路线图
```

---

## 验收

```bash
python tests/run_all.py
```

八套测试各有分工：

| 套 | 验什么 |
|---|---|
| `selftest_shared_lib` | 读侧/写侧/闸门的**机制**：定位三元组指纹、幂等账本、快照、回滚 |
| `selftest_cite_doctor` 等四套 | 各子 skill 的规则**真的能查出问题**：注入缺陷 → 检出 → 修复 → 复检 → 幂等 |
| `selftest_proposal_consistency` | 一致性核对的金标准跑分（召回/误报/定位三指标）+ 改开题内容看结论是否跟着变 |
| `selftest_routing` | 路由可区分性：示例 prompt 是否各归其位、`not_for` 与 `out_of_scope` 是否生效 |
| `e2e_cli` | 命令行**串起来**能不能用：yaml 往返、路径约定、哈希门禁、回滚 |

> **验收范式**：注入式。在干净稿上人为造出缺陷 → 确认能抓到 → 修完确认消失。
> 只验"干净稿不报错"是没有意义的 —— 一个永远返回空列表的检查器也能通过那种测试。

---

## 环境

- Python 3.11+（本机可用：`C:\Users\Tian\.workbuddy\binaries\python\envs\default\Scripts\python.exe`；
  `tools/install.py` 会自动探测一个真能 `import lxml, docx, yaml` 的解释器写进入口存根）
- 依赖：`python-docx`、`lxml`、`pyyaml`；PDF 相关脚本另需 `pymupdf`；
  PDF 导出 / WPS 目录域更新需 Windows + WPS（走 `win32com`）
- 路径、阈值、规则全部外置到 `assets/config.yaml`（样例见 `assets/config.example.yaml`）
- ⚠️ 本机 Git Bash 的 `dirname` / `ls` / `head` / `tail` 不可用，**管道到 `tail` 会直接失败**；
  所有脚本与测试直接用 Python 跑，不要用管道截断

---

## 两条踩出来的硬纪律

1. **规则内异常必须升格为 `audit-internal-error` 的 error 级 Issue，不许只写进 notes。**
   否则"检查器坏了"会伪装成"全部合格" —— 曾经就有一条规则因参数名撞名抛异常被 `except` 吞掉，
   报出"0 条问题"而看起来一切正常。
2. **`apply` 的顺序固定为「先哈希门禁、后回填参数」。**
   否则文档在生成 plan 之后被人工改过时，过期 plan 会解析出 0 条可执行项，然后以"什么都没做"正常退出 —— 又是一个静默失效。
