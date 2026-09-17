---
name: template-extract
description: 从论文模板/格式规范/范文文档中提取真实排版与体例规则，产出可直接套用的 Word 模板（.docx）与 Format Spec（spec.json）。输入可以是 PDF、已有 Word 模板，或模板的截图/扫描图片；输出是一份所有排版都落在**命名样式**上的 .docx——一级标题、正文文本、表格文本等样式都已定义好，在 WPS/Word 的「样式」面板里改一次即可全篇生效。同时产出 `conventions` 体例规则块（引注形态与位置、参考文献编码制、页眉与横线、分节页码方案、摘要上限、图表题注格式），供 paper-detail-doctor 包内其他子 skill 当作「期望值」使用。当用户提到「提取论文模板」「论文格式规范」「按这个模板排版」「从 PDF/Word 提取格式」「论文排版样式」「毕业论文格式」「模板提取」「把这份规范做成 Word 模板」「参考文献用什么编码制」「引注是上标还是行内」时使用。
agent_created: true
metadata:
  parent_package: paper-detail-doctor
  version: "1.1.0"
  last_updated: "2026-09-17"
  status: active
  related_skills:
    - format-audit
    - cite-doctor
    - text-style
    - structure-length
---

# 论文模板提取器（template-extract）

> **这是 `paper-detail-doctor` 包的子 skill。** 它在包内的角色是**标准来源**——其余 5 个审计型子 skill 的「期望值」全部由它产出的 `spec.json` 提供。可由包根 `SKILL.md` 路由调用，也可独立使用。

把「一份论文格式规范 / 模板 / 样张」变成「一个可用的 Word 模板」+「一份机器可读的体例规则」。

**核心承诺**：输出的 .docx 里，所有排版都落在**命名段落样式**上，而不是逐段硬排版。
用户在 WPS 里打开「样式」面板，改一次「一级标题」的字号，全篇一级标题同步变化 —— 这就是「一键调整」。

**在包内的第二重承诺**：spec 里的 `conventions` 块是后面所有检查项的判定依据。
没有它，检查只能退化成"通用学术惯例"，报告里会标注 `标准来源: 默认`。

---

## 环境

| 项 | 值 |
|---|---|
| Python 解释器 | `C:\Users\Tian\.workbuddy\binaries\python\envs\default\Scripts\python.exe` |
| 依赖 | `python-docx`、`pymupdf`、`lxml`（已装在上述 venv） |
| 视觉预览 | LibreOffice（`soffice --headless`），本机在 `F:\Dev\LibreOffice-26.2.3.2\program\soffice.exe` |

**运行方式**：所有脚本都用 venv 的解释器绝对路径调用。

```bash
VPY="$HOME/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
# 安装后：
SK="$HOME/.workbuddy/skills/paper-detail-doctor/skills/template-extract"
# 开发期（本仓库）：
# SK="F:/Coding/Project/paper-detail-doctor/skills/template-extract"
```

若 venv 不存在，用 `venv.EnvBuilder` 重建（注意：本机 `python -m venv` 会静默失败）：

```bash
"$HOME/.workbuddy/binaries/python/versions/3.13.12/python.exe" - <<'PY'
import venv
venv.EnvBuilder(with_pip=True, symlinks=False).create(
    r"C:\Users\Tian\.workbuddy\binaries\python\envs\default")
PY
"$VPY" -m pip install --proxy http://127.0.0.1:7897 python-docx pymupdf
```

> 本机 shell 是 Git Bash，长驻 shell 里 `unset PYTHONPATH` 会打坏 coreutils
> （`ls`/`grep` 全变 command not found）。跑外部程序前用
> `unset NODE_OPTIONS PYTHONPATH; export PATH="/usr/bin:/bin:$PATH"` 兜一下。

**不依赖 COM**：本 skill 纯文件生成，核心链路**不需要** WPS/Word 被自动化驱动。
（WPS COM 已于 2026-09-16 修复可用，见 `harness-anything` SKILL.md 与
`~/.workbuddy/tools/harness-anything/fix_wps_com_hklm.py`；仅「导出 PDF / 批量套模板」
等增强场景才需要它。）

---

## 工作流

```
接入输入 → 判定类型 → 提取 Format Spec → 提取体例规则(conventions) → 人工确认低置信项 → 生成 docx → 回读校验 → 视觉验收 → 落位 .template/
```

### 第 1 步：判定输入类型

```bash
"$VPY" "$SK/scripts/probe_input.py" -i <输入路径> -o probe.json
```

| `kind` | 含义 | 走哪条路 |
|---|---|---|
| `docx` | 已有 Word 模板 | **无损路径**（首选，置信度 high） |
| `pdf` | PDF 规范 / 模板 | 半精确路径 + 视觉复核 |
| `images` | 图片目录（截图/扫描） | **agent 视觉分析**（Read 图片逐页看） |
| `legacy-doc` | 旧版 .doc | 先转 .docx 再走无损路径 |

**同时拿到多种输入时，优先用 `.docx`。** 一份 .docx 顶十张截图。

### 第 2 步：提取 Format Spec

**DOCX（无损）**

```bash
"$VPY" "$SK/scripts/extract_from_docx.py" -d 模板.docx -o spec.json --used-only
```

直接读 `styles.xml` / `sectPr`，拿到的是真值。读回后仍是**草稿**：
需要 agent 把 `suggested_role` 映射到 spec 的标准样式 key，并补齐 `table` / `figure` /
`equation` / `references` / `demo` 等段落（脚本不生成这些）。

**PDF（半精确）**

```bash
"$VPY" "$SK/scripts/extract_from_pdf.py" -p 模板.pdf -o spec.json --max-pages 30
```

产出字号直方图与推断出的 `h1/h2/h3/body`。**必须**再用 Read 工具看几页 PDF
（或渲染出的图）做视觉复核；段前段后、缩进、页边距、表格线宽要**问用户或读文字规范**。

**图片（视觉）**

脚本帮不上忙 —— **agent 自己 Read 每一页图片**，按 `references/extraction_heuristics.md`
的规则推断。关键动作：

1. 先通读全部图片，分出「文字规范页」和「排版样张页」。
2. **文字规范页优先**：段前段后磅值、行距、缩进字符、表格线宽、页码方案，
   这些精确值只在文字里，样张看不出来。
3. 样张页用来确认字体、字号档位、加粗、对齐、编号形式（`1.1` 还是 `一、`）。
4. 两者冲突时**以文字规范为准**，并把冲突记进报告。
5. 样张里字号只能定**档位**（比正文大一号），无法定磅值 —— 标 `medium`。

> 若输入里同时有文字规范页和样张页，在 `meta.source.note` 里写清楚这个互补关系。

### 第 2b 步：提取体例规则（`conventions`）—— 包内其他子 skill 靠它吃饭

> 只在**有"已排好版的范文 / 样张 docx"**时才能拿到高置信度结果。
> 范文比"格式规范文字"更可靠：文字规范可能没写完，范文是实际排出来的。

```bash
"$VPY" "$SK/scripts/extract_conventions.py" -d 范文.docx --merge-into spec.json
```

脚本从范文的**实际排版**里反推六组规则 —— 这些正是包内其他子 skill 的判定基准：

| 组 | 推断什么 | 怎么推 |
|---|---|---|
| `citation` | 引注形态（上标方括号 / 行内方括号 / 圆括号 / 上标数字）、位置（句读前 / 后）、上标字号比、多引用写法 | 扫正文 run 的 `vertAlign=superscript`；看引注 run 后一个字符是不是句读符 |
| `bibliography` | **编码制**（顺序编码制 / 著者-出版年制）、著录标准、编号格式 | 看参考文献条目开头是 `[n]` 还是「作者 年份」 |
| `header` | 页眉文字（论文题目 / 学校名 / 空）、**是否要横线**、线型与粗细 | 读页眉段落的 `pPr/pBdr/bottom` |
| `page_numbering` | 分节页码方案（封面无 / 摘要罗马 / 正文阿拉伯）与起始页 | 遍历各节 `sectPr/pgNumType` 的 `fmt` 与 `start` |
| `abstract` | 中英文摘要字数上限、是否单独成页、关键词个数与分隔符 | 量摘要正文长度、数关键词个数 |
| `caption` | 图题/表题位置与编号格式（图 1-1 / 图1.1） | 扫图题表题段落的相对位置与文本模式 |

产出的 `conventions` 块**并进 spec.json 顶层**，键名与 `references/format_spec_schema.md` 的「conventions」一节一致。

**低置信度怎么办**：推断不出来的（比如范文本身格式就不统一）标 `confidence: low` 并进 `open_questions`，
**不要猜得像真的** —— 下游会当硬标准用。宁可标 low 让用户确认。
`citation.style` 与 `bibliography.system` 这两项如果没推出来，会**显著影响**后面 L2 引注检查的准确性，
所以它们 low 时必须问用户。

### 第 3 步：人工确认低置信项

把 `confidence` 为 `low` 的字段收集进 `open_questions`，**逐条向用户确认**。
参考 `references/extraction_heuristics.md` §1 —— 规范没写的，宁可标 low 也别猜得像真的。

### 第 4 步：生成 Word 模板

```bash
"$VPY" "$SK/scripts/build_docx.py" -s spec.json -o 模板.docx
```

生成内容：页面设置 → 命名样式（标题1-3/正文/表格文本/表题/图题/摘要/关键词/参考文献/页眉/页码…）
→ 标题多级自动编号 → 分节页码方案 → 演示骨架页 → 样式总览页 → **自动清理多余样式**。

可选参数：`--no-numbering`（不启用自动编号）、`--no-demo`（只出样式，不要演示内容）、
`--no-clean`（**不建议**：保留 python-docx 默认模板自带的 164 个未使用内置样式）。

#### 自动清理（默认开启，别关掉）

python-docx 的默认模板自带整套 Word 内置样式 —— **164 个**，其中 100 个是表格样式
（Light Shading Accent 1 … Colorful Grid Accent 6）。不清理的话，用户在 WPS 里打开
「样式」面板会看到一大片与本文档毫无关系的样式，非常碍事。

`build_docx.py` 保存后会自动调 `clean_styles.py`，实测 **191 → 32**，只剩：

- 我们的 27 个自定义样式
- 4 个规范必需的默认样式（Normal / Default Paragraph Font / Normal Table / No List，不可删）
- Table Grid（三线表实际在用）

清理是安全的：只动"没有任何内容部件引用、且 spec 里也没定义"的死样式。
为了做到这一点，生成阶段还有两处配套处理：

- **标题样式自带大纲级别**（`w:outlineLvl`），不再基于内置 `Heading 1/2/3`
  → 内置标题样式可以彻底删掉，目录域照样抓得到（见「关键约束」4）。
- **页眉/页脚段落改写成本模板自己的 `页眉`/`页码` 样式**
  → 内置 `Header`/`Footer` 不再被引用而无法清理。
  空白页眉（规范里"摘要、目录部分无页眉"）**不挂样式**，否则页眉样式的下划线会在空白页眉上画线。

单独清理任意 docx（也可用于别人给的模板）：

```bash
"$VPY" "$SK/scripts/clean_styles.py" -d 模板.docx [-o 输出.docx] \
       -s spec.json [--keep "样式名1,样式名2"] [--report clean_report.json]
```

> 目标文件被 WPS 打开时重命名会失败，脚本会自动退化成**原地重写**，并打印 warning。
> 此时让用户关掉再重新打开文档即可。

### 第 5 步：回读校验 + 视觉验收

```bash
"$VPY" "$SK/scripts/verify_docx.py" -d 模板.docx -s spec.json     # 逐字段比对 + 样式体检
"$VPY" "$SK/scripts/preview_docx.py" -d 模板.docx -o preview/     # 渲染 PNG
```

`verify_docx.py` 检查两组东西，**两组都要通过**：

| 字段 | 必须 | 含义 |
|---|---|---|
| `styles_with_diff` / `total_field_diffs` | 0 | 27 个样式的字体/字号/间距/缩进与 spec 完全一致 |
| `cleanliness.clean` | `true` | 没有多余样式（`styles_unused` 为空） |

`cleanliness` 会把样式分成三类解释：被引用 / **预留**（spec 里定义了但还没用上，
如四级标题、目录一二三级）/ 未使用（既没引用、spec 也没定义 —— 这才是要被清掉的）。

然后 Read `preview/preview_pXX.png` 亲眼确认排版，再交付。

### 第 6 步：落位 `.template/`（必做，保证模板统一性）

提取只发生一次，但"按模板写作 / 排版"会发生无数次。所以最终产物**必须固化到项目根目录的
`.template/`**，让后续会话有唯一可寻址的格式真值：

```bash
"$VPY" "$SK/scripts/init_template_dir.py" \
  -p <项目根> \
  --spec spec.json --docx 模板.docx --report 提取报告.md \
  --source <输入文件（证据存档，可选）> \
  --name "<模板名，如：本科毕业论文>"
```

脚本做四件事：建目录结构、复制 spec/docx/报告、**从 spec 派生 `constraints.md`**、
维护 `manifest.json` 状态机。幂等：已存在的 spec/docx 不会被覆盖（除非 `--force`）；
spec 有更新时重跑一次，`constraints.md` / manifest 会同步刷新。

**状态机（manifest.json）**：`draft → confirmed → applied`。
- `open_questions > 0` 时强制 `draft` —— **draft 状态的模板不得被下游写作环节消费**；
- 用户逐条确认完待确认项后：更新 spec → 重跑本脚本 → 手动把 status 改为 `confirmed`；
- 下游（写作 / 排版 / apply）首次成功套用后记 `applied`。

> spec 变更的正确姿势：改 `.template/spec.json` → 重跑 `build_docx.py` 重新生成 docx
> → `verify_docx.py` 校验 → 重跑 `init_template_dir.py`（**不要 `--force`**，
> 它只刷新派生物）。直接手改 `template.docx` 会在下次生成时被覆盖。

#### `.template/` 目录契约

| 文件 | 角色 | 下游谁在消费 |
|---|---|---|
| `manifest.json` | 状态清单（状态机 / 来源 / schema 版本 / 多模板登记） | agent 判断可用性 |
| `spec.json` | ★ 唯一格式真值（Format Spec） | `build_docx.py`、未来的 apply |
| `template.docx` | 派生：命名样式模板（WPS/Word 用） | 用户直接打开写作 |
| `constraints.md` | ★ 写作约束摘要（人+机器双读） | 写作 agent（如 academic-paper）读它获取格式硬约束 |
| `extraction_report.md` | 提取报告（规则来源 + 置信度 + 待确认项） | 人 |
| `preview/` | 验收渲染图 | 人 |
| `source/` | 输入证据存档 | 复查 / 重新提取 |

> 多模板：同一项目需要"毕业论文版 + 期刊投稿版"时，用 `--name` 分别登记，
> manifest 的 `templates` 映射 + `active` 指针管理；spec/docx 放 `.template/<profile>/` 子目录。

---

## 关键约束（改代码前必读）

1. **排版必须落到命名样式**，不要用 `run.font.size = ...` 逐段硬排版 ——
   那样用户就没法「一键调整」了。所有格式化都写在 `doc.styles` 里。
2. **中文字体要写 `w:eastAsia` 槽**。python-docx 的 `font.name` 只写 ascii/hAnsi，
   汉字会跟默认字体渲染。统一走 `common.set_east_asian_font()`。
3. **首行缩进用 `w:firstLineChars`**（字符语义），不要用 `w:firstLine`（磅值语义）。
4. **标题样式自带大纲级别**：spec 里写 `outline_level`（一级=0、二级=1…），由
   `common.set_outline_level()` 落成 `w:outlineLvl`。**不要**再基于内置 `Heading 1/2/3`
   —— 一旦被引用，这些内置样式就被判定为"已使用"而清理不掉，WPS 样式面板里会一直
   多出"标题 1 / 标题 2 / 标题 3"。目录域（`TOC \o`）和导航窗格认的是**段落大纲级别**，
   不认样式名，所以自带级别一样抓得到。
5. **不该进目录的标题**（封面题目、摘要标题、目录标题）**不写** `outline_level`。
   参考文献/致谢/附录标题写 `outline_level: 0`（模板的目录样张里有它们）。
6. **标题自动编号**用 `common.ensure_heading_numbering()` + `bind_style_to_numbering()`；
   失败时自动降级为 `fallback_number` 手工编号。辅助页（如样式总览）用
   `disable_paragraph_numbering()` 关掉编号。
7. **三线表** = 上下线 1.5 磅 + 表头下中线 0.5 磅 + 无竖线，见
   `common.apply_three_line_table()`。磅值在 XML 里是 1/8 磅（1.5pt → `sz=12`）。
8. **节的页码格式**（罗马/阿拉伯）用 `common.set_section_page_numbering()`；
   `pgNumType` 在 `sectPr` 里位置有讲究，不能随便 append。
9. **交付的 docx 里不允许有死样式**。生成后必须跑 `clean_styles.py`（`build_docx.py`
   已默认调用）。新增任何"基于内置样式"的写法前，先想清楚它会不会把内置样式重新
   拉进文件里。改完重新生成，看 `verify_docx.py` 的 `cleanliness.clean` 是否为 `true`。

---

## WPS / Word 的坑（实测）

| 现象 | 原因 | 怎么办 |
|---|---|---|
| 样式面板列出一大片 Light Shading / Colorful Grid… | python-docx 默认模板自带 164 个内置样式 | 交给 `clean_styles.py`（默认已开） |
| 样式名后多了个 `1`（`正文文本1` / `页眉1` / `页码1`） | WPS 内置样式集里就有「正文文本」「页眉」「页码」，保存时与自定义样式撞名，WPS 给自定义的加后缀 | 要么接受；要么把 spec 里这 3 个 `display_name` 改成不撞名的写法（如 `论文正文`/`论文页眉`/`论文页码`）。`clean_styles.py` 的 `--keep` 匹配已容忍这个后缀 |
| 保存后 `styleId` 全变成纯数字、还多了 `toc 1/2/3` | WPS 会重写整份 `styles.xml`（内置样式"物化"） | 无法阻止，属 WPS 行为；我们的清理保证**交付时**是干净的 |
| 目录页是空的、只有一句提示 | 目录是**域**，未展开 | 让用户在 WPS 里点目录 → 右键「更新域 → 更新整个目录」 |
| 脚本报 `PermissionError: ...docx.tmp -> ...docx` | 文档正被 WPS/预览面板占用，`os.replace` 拿不到 DELETE 权限 | 已自动退化为原地重写；提醒用户关闭后重新打开 |

> 附带发现：目录域展开后用的是 Word 内置的 `toc 1/2/3` 样式，**不认**我们的
> `目录一级/二级/三级`。所以那 3 个样式是给"手工目录"用的；若要自动目录完全符合格式，
> 只能改内置 `toc 1-3`，或手工排目录。

## 文件

| 文件 | 作用 |
|---|---|
| `scripts/probe_input.py` | 判定输入类型，给推荐路径 |
| `scripts/extract_from_docx.py` | DOCX → spec（无损，含继承链解析） |
| `scripts/extract_from_pdf.py` | PDF → spec（字号直方图 + 角色推断） |
| `scripts/extract_conventions.py` | ★ **范文 → `conventions` 体例规则**（引注形态/位置、编码制、页眉与横线、分节页码、摘要上限、题注格式），可 `--merge-into spec.json` |
| `scripts/build_docx.py` | spec → 带命名样式的 .docx（保存后自动清理样式） |
| `scripts/clean_styles.py` | 清掉未使用样式 + 未引用编号定义（可单独对任意 docx 使用） |
| `scripts/verify_docx.py` | 回读 .docx 逐字段校验 + 样式体检（`cleanliness`） |
| `scripts/preview_docx.py` | .docx → PNG（视觉自检） |
| `scripts/init_template_dir.py` | 在项目根建立/刷新 `.template/` 资产目录 + 派生 constraints.md + manifest 状态机 |
| `scripts/common.py` | 字号换算 / OOXML 工具 / 三线表 / 编号 |
| `references/format_spec_schema.md` | Format Spec 字段定义 |
| `references/cn_font_sizes.md` | 号↔磅对照、字体槽、PDF 字体反查 |
| `references/extraction_heuristics.md` | 置信度判定与推断规则 |
| `assets/spec_example.json` | 完整示例（中文本科毕业论文格式） |

## 输出约定

工作目录的中间产物放 `<输入同级>/模板提取输出/`（spec、docx、报告、preview），
但**最终必须落位到项目根 `.template/`**（见第 6 步）——下游只认 `.template/`：

- `.template/spec.json`（★ 唯一真值，可改可重新生成）
- `.template/template.docx`（派生物）
- `.template/constraints.md`（★ 写作约束摘要，下游 agent 的消费入口）
- `.template/extraction_report.md` + `manifest.json`（报告与状态）
- `.template/preview/`、`.template/source/`（验收图与证据存档）

### 路线图（未实现，勿在对话里承诺）

- `apply_spec.py`：把 spec 套到**已有内容**的 docx（真正的"按模板重排已有论文"）——
  下一个大功能，实现前先读 spec 的样式 key 映射与 `common.py` 的样式构造函数。
