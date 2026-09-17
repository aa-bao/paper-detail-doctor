# SOP｜模板标准（把"学校要求"变成可判定的基准）

> **跨 skill 的 SOP。** 本文讲的是"判定基准从哪来、为什么不能内置、怎么从模板/范文反推、
> 哪些值不能当标准用"。它不是 `template-extract` 的使用说明——那部分见
> `skills/template-extract/SKILL.md`。
>
> 涉及：`template-extract`（产出 spec）+ `shared/lib/standards.py`（三级优先解析）
> + 全部 5 个审计型子 skill（消费标准）。

---

## 一、为什么标准必须外置

**这个包不内置"我的排版偏好"。** 一切判定的第一优先依据是**你学校给的模板或范文**。

理由很简单：工具作者觉得"页边距 2.5cm 好看"毫无意义——判你论文的是你学校的格式规范。
把偏好写死进代码，等于拿别人的标准判你的稿子，报告里也就再也说不清"到底谁规定的"。

于是本包的做法是：**把标准抽成一个可替换的档案**（`spec.json`），
判定逻辑与判定依据彻底分离。换学校 → 换 spec，代码一行不改。

---

## 二、三级优先（`shared/lib/standards.py` 是唯一解析入口）

| 优先级 | 来源 | 标在报告里的值 | 什么时候用 |
|---|---|---|---|
| ① 最高 | 模板/范文 spec | `template` | 跑了 `template-extract`，从学校模板/范文量出来的 |
| ② | 用户配置覆盖 | `config` | 你有明确的规矩但模板里没有（如"我们院要求摘要 ≤800 字"） |
| ③ 最低 | 通用学术惯例 | `default` | 没 spec 可用时的兜底；**报告里必须显著标出** |

**为什么必须抽到 shared**：每个子 skill 都要回答"这条判定依据的是学校模板、用户配置，
还是通用惯例"。各写各的，同一个字段就会出现不同默认值与不同来源口径，
报告里再也说不清责任。所以这是**唯一**的解析入口。

一条纪律：**报告里每条 issue 都带 `standard_source`**（`template` / `config` / `default`）。
看到 `default` 就说明"这是通用惯例，不代表你学校的要求"——别把它当成学校规定去改稿。

没有 spec 时，报告顶部会明写：

> **未提供模板 spec**：本次全部判定依据为通用学术惯例（default），不代表你学校的格式要求。
> 要按学校标准判，请先跑 template-extract。

---

## 三、★ 核心纪律：实测值 ≠ 上限

这是本 SOP 最值钱的一条，也是最容易做错的一条。

`extract_conventions` 只能**测量范文实际长什么样**。它**不知道**学校的上限是多少。举例：

| 情况 | 能不能当标准用 |
|---|---|
| 范文摘要 506 字 | ❌ **不能**说"超过 506 字就报警"——所有比范文长一点的摘要都会被误报 |
| 学校规范文档写着"摘要不超过 800 字" | ✅ 这是硬指标，可以判 |
| 范文正文 35,000 字 | ❌ 不能当年限；字数配额要靠用户显式给 |
| 范文页眉有横线 | ❌ 不能推出"学校要求有横线"——可能只是这位作者的个人习惯 |

**本包的处理方式**（写死在 `standards.py` 里）：

1. 纯测量值**不进** `CONV_MAP`（那张"spec 路径 → 扁平键"的映射表只收录**确实能当标准用**的项），
   避免被误当上限。
2. 真要有上限，必须在 spec 里**显式写**上限字段（如 `abstract.cn.max_chars`），
   表示"这是从格式规范文档抄来的硬指标"。
3. 对"范文有值但规范未必规定"的字段（页边距、字数配额、页码方案等）：
   **默认值兜底 + 把实测值作为上下文写进 Issue**，并在 `standard_source` 里**如实标 `default`**。

> 换句话说：**实测值可以进报告当参考，不可以进判定当门槛。**

---

## 四、宁可少报，不可瞎报

`standards.py` 的 `DEFAULTS` 只放"不写就没法判"的项。凡是"没标准就不该判"的，
**一律留 `None`**，由报告明说"因缺标准未判"：

| 字段 | 默认 | 理由 |
|---|---|---|
| `page_margins_cm` | `None` | 每个学校不同，猜必错 |
| `body_size_pt` / `body_font_cn` / `body_line_spacing` | `None` | 同上 |
| `page_numbering_scheme` | `None` | 范文的分节方案是"范文长这样"，不是"学校要求这样" |
| `header_content` / `header_rule` | `None` | 页眉文字/横线尤其危险（见第三节例子） |
| `body_min_chars` | `None` | 字数配额必须用户显式给 |

解析后如果有 `None`，`resolve_standard()` 会往 `notes` 里塞一句：

> 以下项**因未提供标准而未作判定**（不是"没问题"，是"没法判"）：……

**这条与 A6 的"未配置锚点一律报无法判定"是同一条纪律的两次应用：
"没法判"必须说出来，不许伪装成"没问题"。**

---

## 五、怎么从模板/范文反推标准（`template-extract` 流水线）

```bash
VPY="python"
SK="F:/Coding/Project/paper-detail-doctor/skills/template-extract"

# ① 判定输入类型（docx / pdf / 规范文档），给出推荐路径
"$VPY" "$SK/scripts/probe_input.py" -i <输入路径> -o probe.json

# ② 提取 Format Spec（二选一，取决于你手上是什么）
"$VPY" "$SK/scripts/extract_from_docx.py" -d 模板.docx -o spec.json --used-only
"$VPY" "$SK/scripts/extract_from_pdf.py"  -p 模板.pdf  -o spec.json --max-pages 30

# ②b ★ 从"已排好版的范文"反推体例规则，并进 spec 顶层 —— 包内其他子 skill 靠它吃饭
"$VPY" "$SK/scripts/extract_conventions.py" -d 范文.docx --merge-into spec.json

# ③ 人工确认低置信项（见第六节）

# ④⑤⑥ 生成 Word 模板 → 回读校验 + 视觉自检 → 落位 .template/
"$VPY" "$SK/scripts/build_docx.py"  -s spec.json -o 模板.docx
"$VPY" "$SK/scripts/verify_docx.py" -d 模板.docx -s spec.json
"$VPY" "$SK/scripts/preview_docx.py" -d 模板.docx -o preview/
"$VPY" "$SK/scripts/init_template_dir.py" ...
```

**第 ②b 步是整个包的支点**：`conventions` 块是 L1–L4 五个审计型子 skill 的判定基准。
`extract_conventions.py` 的真实参数（`--help` 原文）：

```
usage: extract_conventions.py [-h] -d DOCX [-o OUT] [--merge-into MERGE_INTO] [--strict]

  -d, --docx DOCX       范文 / 样张 docx（已排好版的）
  -o, --out OUT         输出 JSON 路径（默认打印到 stdout）
  --merge-into MERGE_INTO   把 conventions 块并进指定 spec.json 的顶层
  --strict              有命门项缺失时非零退出
```

它提取六组规则：`citation`（形态/位置/字号比/多引用）、`bibliography`（编码制/著录标准/编号格式）、
`header`（文字 + 是否要横线 + 线型粗细）、`page_numbering`（分节页码方案）、
`abstract`（字数/关键词/是否单页）、`caption`（图题表题位置与编号格式）。

---

## 六、提取环节的坑位（都真实踩过）

### 坑 1｜引注位置："句读符"不是"句末标点"

一度把 `；` `，` 当成"标点后"，于是一大批 `……问题[1]。` 被判成"标点后"。
实际上 `；，` 属**句读符**，引注落在它们**之前**正是 `before-punct`（句读前）体例。

判据必须写死三条（错一条就会批量误判）：

1. 后一字符 ∈ 句读符 → `before-punct`
2. 前一字符 ∈ 句末标点 **且** 后一字符非句读符 → `after-punct`
3. 否则 → `mid-sentence`

### 坑 2｜标题含全角空格未归一化 → 参考文献扫描越界

范文标题写作「摘　　要」「目　　录」「致　　谢」（中间是 U+3000 全角空格）。
不归一化 → "致谢"匹配不上 → 参考文献扫描**越界**，把 7 段致谢正文当成条目（46 → 53 条）。

修法：加 `_norm()`（去掉所有空白**含 U+3000** + 小写）后再比对。

### 坑 3｜条目长度阈值会误杀英文长条目

`len(t) > 250` 丢掉了一条 262 字符的英文条目（英文作者名多）。
改为：**有 `[n]` 编号的一律保留**，只丢"长且无编号"的段落。

### 坑 4｜页眉"是不是论文题目"不能硬猜

原本的判据是"页眉文字是否出现在正文前 40 段"。但该 docx 正文里根本没有封面
（题目只写在 `header2.xml`），于是判成 `custom`。

**正确做法是老实认输**：报 `content_confidence: low`，写进 `open_questions` 转人工，
并在 evidence 里说明"该文档正文可能没有封面"。
**猜错一个标准，比不判一个标准危害大得多**——它会连带污染所有依赖该标准的规则。

### 坑 5｜上标字号比测不出来是正常的

`superscript_size_ratio` 为 `null` 时不要慌：上标字号由 `w:vertAlign` **自动缩放约 65%**，
范文里根本没有显式的 `w:sz`。所以 `null` 本身就是要的答案——期望状态是"不设字号"。
（对应 `cite-form` 的 `cite_allow_explicit_size = False`。）

### 坑 6｜判"能否自动清理"要看得到证据

`own_page`（摘要是否独占一页）这类字段，判不出来时应返回 `null` 并说明"须看渲染结果"，
而不是给个 best-guess。同理 `page_numbering` 的分节方案，判不出就报"多分节/未判"。

---

## 七、spec 放哪、谁来读

`shared/lib/standards.py` 的 `find_spec()` 按下列顺序在**文档同级目录**里找，先命中者胜：

1. `paper-spec.json`
2. `spec.json`
3. `.thesis-doctor/spec.json`
4. `.paper-doctor/spec.json`

找不到时**不报错**（这是正常情况——用户还没跑 template-extract），
各子 skill 据此把 `standard_source` 全标 `default` 并说明。

子 skill 也可以显式指定（`cite-doctor -s spec.json`、`audit_all -s spec.json`）。

---

## 八、迁移到另一篇论文的改动清单

1. **换模板**：跑第三节流水线，得到新的 `spec.json`。**代码不改**。
2. **只覆盖差异项**：模板里没有、但你有明确规矩的（如院系补充规定），
   写进 config 覆盖（第 ② 级），不要改模板值。
3. **确认低置信项**：把 extract 报的 `content_confidence: low` / `open_questions` 全部人工过一遍。
4. **别把实测值当上限**：看到 `conventions.abstract.cn.chars = 506` 这类字段，
   记住它是"范文长这样"；要判上限，去 spec 里显式写 `max_chars`。
5. **重跑体检**，看报告里 `standard_source` 是 `template` 的项是否变多了——
   变多说明标准真的生效了；仍是 `default` 说明那一项没接上。

---

## 与其他文档的分工

| 想了解 | 去哪看 |
|---|---|
| template-extract 的完整流水线、`.template/` 契约、样式清理 | `skills/template-extract/SKILL.md` |
| spec 的字段 schema（含 `conventions` 块） | `skills/template-extract/references/format_spec_schema.md` |
| 中文字号对照 | `skills/template-extract/references/cn_font_sizes.md` |
| 三来源解析的代码与默认值全表 | `shared/lib/standards.py` |
| 引注相关标准怎么用 | `SOP-引注上标与跳转.md` |
| 开题与正文对不对得上 | `SOP-开题一致性核对.md` |
