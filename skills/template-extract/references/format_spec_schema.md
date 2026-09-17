# Format Spec 字段定义

Format Spec 是整个 skill 的**中间产物**：提取的结果先落成它，再由它生成 Word。
好处是可人工修正、可版本化、**同一份规范能反复套用到多篇论文**。

顶层结构：

```jsonc
{
  "meta":     { ... },   // 来源、时间、总体置信度
  "fonts":    { ... },   // 默认中英文字体
  "page":     { ... },   // 页面与版面
  "numbering":{ ... },   // 章节编号体系
  "styles":   { ... },   // 各段落样式（核心）
  "table":    { ... },
  "figure":   { ... },
  "equation": { ... },
  "references": { ... },
  "sections": [ ... ],   // 分节与页码方案
  "demo":     [ ... ],   // 生成模板时的骨架内容
  "open_questions": [ ... ]  // 低置信项，待人工确认
}
```

## meta

| 字段 | 说明 |
|---|---|
| `template_name` | 模板名称 |
| `source.kind` | `docx` / `pdf` / `images` |
| `overall_confidence` | `high` / `medium-high` / `medium` / `low` |
| `confidence_legend` | 三档置信度的判定说明 |

## page

| 字段 | 单位 | 说明 |
|---|---|---|
| `width_cm` / `height_cm` | cm | 纸张尺寸（A4 = 21.0 × 29.7） |
| `orientation` | — | `portrait` / `landscape` |
| `margin_top_cm` / `margin_bottom_cm` / `margin_left_cm` / `margin_right_cm` | cm | 页边距 |
| `gutter_cm` / `gutter_position` | cm / `left\|right\|top` | 装订线 |
| `header_distance_cm` / `footer_distance_cm` | cm | 页眉/页脚距边界 |
| `doc_grid` | — | `none` 或 `lines` 等 |
| `duplex` | bool | 是否双面印刷 |

## numbering

| 字段 | 说明 |
|---|---|
| `enabled` | 是否启用自动编号 |
| `scheme` | `numeric`（1.1）/ `chinese`（一、）/ `mixed` |
| `levels[]` | 每级的 `lvlText`（`%1`=一级号，`%2`=二级号…）与 `suff`（编号与文字间分隔：`space` / `tab` / `nothing`） |

示例（数字型 + 一级用「第X章」）：

```json
"levels": [
  { "lvlText": "第%1章", "suff": "space" },
  { "lvlText": "%1.%2",  "suff": "space" },
  { "lvlText": "%1.%2.%3", "suff": "space" }
]
```

## styles（核心）

每个样式对象：

| 字段 | 取值 | 说明 |
|---|---|---|
| `display_name` | 字符串 | **WPS 样式面板里显示的名字**（用户就是靠它一键调样式） |
| `role` | `heading1..4` / `body` / `caption` / `table` / `toc` / `header` / `pagenumber` / … | 语义角色 |
| `cn_font` / `en_font` | 字体名 | 写入 eastAsia / ascii 槽 |
| `size_pt` | 数字 | 磅值（**不是**「三号」，由换算表转） |
| `size_cn_label` | 字符串 | 仅用于报告可读性 |
| `bold` / `italic` | bool | |
| `align` | `left` / `center` / `right` / `justify` | |
| `space_before_pt` / `space_after_pt` | 数字 | 段前 / 段后（磅） |
| `line_spacing_type` | `multiple` / `exact` / `atLeast` | |
| `line_spacing` | 数字 | 倍数时为 1.5，固定值时为磅数 |
| `first_line_indent_chars` | 数字 | 首行缩进字符数（用 `firstLineChars`） |
| `left_indent_chars` | 数字 | 左缩进字符数 |
| `base_style` | `Normal` / `Heading 1..9` / null | 父样式。**一般填 `null`**（见下方「大纲级别」） |
| `outline_level` | 0-8 / 省略 | **大纲级别**：一级标题 0、二级 1、三级 2…。写进 `w:outlineLvl`，目录域与导航窗格靠它识别。封面题目/摘要标题/目录标题**不写**（不进目录） |
| `keep_with_next` | bool | 标题是否与下段同页（`w:keepNext`） |
| `keep_lines` | bool | 段中不分页（`w:keepLines`），标题建议开 |
| `confidence` / `evidence` | — | 该样式的置信度与依据 |

### 大纲级别：用 `outline_level`，不要用 `base_style: Heading N`

早期写法是把标题样式挂在 Word 内置的 `Heading 1/2/3` 上借大纲级别。**现在不要这么做**：

- 一旦有样式引用 `Heading 1`，这个内置样式就被判定为"已使用"，`clean_styles.py`
  删不掉它 —— 用户在 WPS 样式面板里就会看到多余的「标题 1 / 标题 2 / 标题 3」。
- 目录域（`TOC \o "1-3"`）和导航窗格认的是**段落的 outlineLvl**，跟样式叫什么名字无关，
  所以自带级别完全等效，还能省掉对内置样式的依赖。

所以：标题类样式统一 `base_style: null` + `outline_level: N`。

### 推荐样式清单

| key | display_name | 典型来源 |
|---|---|---|
| `cover_title` | 封面题目 | 封面 |
| `h1` / `h2` / `h3` / `h4` | 一级标题 / 二级标题 / 三级标题 / 四级标题 | 正文章标题 |
| `body` | 正文文本 | 正文 |
| `abstract_title_cn` / `abstract_body_cn` / `keywords_cn` | 摘要标题 / 摘要正文 / 关键词 | 中文摘要 |
| `abstract_title_en` / `abstract_body_en` / `keywords_en` | 英文摘要标题 / 英文摘要正文 / 英文关键词 | 外文摘要 |
| `toc_title` / `toc_h1` / `toc_h2` / `toc_h3` | 目录标题 / 目录一级 / 目录二级 / 目录三级 | 目录 |
| `table_caption` / `table_text` | 表题 / 表格文本 | 表 |
| `figure_caption` / `figure_source` | 图题 / 图片来源 | 图 |
| `equation` | 公式 | 公式 |
| `ref_title` / `ref_entry` | 参考文献标题 / 参考文献条目 | 参考文献 |
| `ack_title` / `appendix_title` | 致谢标题 / 附录标题 | 致谢、附录 |
| `header` / `page_number` | 页眉 / 页码 | 页眉页脚 |

## sections

描述分节与页码方案（摘要目录用罗马数字、正文用阿拉伯数字是常见要求）：

```json
[ { "name": "封面", "header": null, "page_number": null },
  { "name": "摘要 / ABSTRACT / 目录", "header": null,
    "page_number": { "fmt": "lowerRoman", "start": 1 } },
  { "name": "正文 / 参考文献 / 致谢 / 附录",
    "header": "本科生毕业论文（设计）题目",
    "page_number": { "fmt": "decimal", "start": 1 } } ]
```

`fmt` 取值：`decimal` / `lowerRoman` / `upperRoman` / `lowerLetter` / `upperLetter`。

## demo

生成模板时的骨架内容，一个块一个对象：

| `block` | 含义 | 关键字段 |
|---|---|---|
| （无） | 普通段落 | `style`（样式 key）、`text`、`runs`（run 级加粗）、`fallback_number` |
| `cover` | 封面页 | `text`、`lines` |
| `new_section` | 新分节 | `header`、`page_number`、`no_header`、`page_number_footer` |
| `pagebreak` | 分页 | — |
| `toc` | 目录域 | `text` |
| `table` | 三线表 | `caption`、`header`、`rows`、`source_note` |
| `figure` | 图占位 | `text`、`caption`、`source_note` |
| `equation` | 公式 | `text`、`note` |
| `style_overview` | 样式总览页 | — |

> `fallback_number`：当无法启用自动编号时，用它给标题手工加序号（如 `"第1章"`）。

## conventions（体例规则 · 包内其他子 skill 的判定基准）

> 由 `scripts/extract_conventions.py` 从**已排好版的范文 docx** 反推，并进 spec 顶层。
> `styles` / `page` / `table` 描述的是"排版长什么样"，`conventions` 描述的是"规则是什么"。
> **paper-detail-doctor 包内 5 个审计型子 skill 全部读这一块当期望值** —— 没有它，
> 检查只能退化成通用学术惯例，报告里会标 `标准来源: 默认`。

```jsonc
"conventions": {
  "citation": {
    "style": "superscript-bracket",   // superscript-bracket | superscript-number
                                      // | inline-bracket | inline-paren
    "position": "before-punct",       // before-punct（句读前，规范写法）| after-punct
    "multiple": "single",             // single | separate（[1][2]，多被规范禁止）| merged（[1-2]）
    "superscript_size_ratio": null,   // 上标字号 / 正文字号；null 表示未显式设 w:sz
    "superscript_size_ratio_note": "字号由 w:vertAlign 自动缩放（约 65%）",
    "cited_count": 46,
    "form_tally": {...},              // 各形态出现次数，用于判断体例是否统一
    "position_tally": {"before-punct": 41, "mid-sentence": 5},
    "confidence": "high", "evidence": [...]
  },
  "bibliography": {
    "system": "sequence",             // ★ sequence（顺序编码制）| author-year（著者-出版年制）
    "standard": "GB/T 7714",
    "number_format": "[n]",
    "entry_count": 46,
    "stopped_at": "致　　谢",          // 靠哪个标题截住条目扫描
    "tally": {"sequence_like": 46, "author_year_like": 0, "has_type_marker": 46},
    "confidence": "high", "evidence": [...]
  },
  "header": {
    "content": "paper-title",         // paper-title | custom | none
    "content_confidence": "high",     // ★ 判纸页眉写什么：low 时必须问用户
    "text": "（论文题目）",
    "found_in_body": true,            // 页眉文字是否在正文里出现（用于判 content）
    "rule": true,                     // ★ 页眉要不要横线
    "rule_detail": {"style": "single", "size_pt": 0.5, "space_pt": 1.0},
    "confidence": "high", "evidence": [...]
  },
  "page_numbering": {
    "scheme": "multi-section",        // single | multi-section（前置罗马+正文阿拉伯）| unknown
    "sections": [{"index": 1, "format": "upperRoman", "start": 1},
                 {"index": 2, "format": "arabic", "start": 1}],
    "confidence": "high", "evidence": [...]
  },
  "abstract": {
    "cn": {"found": true, "chars": 506, "keywords_count": 5, "keywords_separator": "；",
           "confidence": "high"},
    "en": {"found": true, "words": 306, "confidence": "high"},
    "own_page": true,                 // null = docx 层面判不出来，须看渲染结果
    "own_page_basis": "摘要标题段落设了 pageBreakBefore",
    "confidence": "high", "evidence": [...]
  },
  "caption": {
    "figure": {"position": "below", "number_format": "图 {chapter}-{index}",
               "confidence": "medium"},
    "table":  {"position": "above", "number_format": "表 {chapter}-{index}",
               "confidence": "medium"},
    "confidence": "medium", "evidence": [...]
  }
}
```

### 三条使用纪律

1. **`confidence: low` 的字段一律不得当硬标准用**，必须进 `open_questions` 让用户确认。
   `citation.style` 与 `bibliography.system` 是 L2 引注检查的命门，为 low 时**必须问**。
2. **报告里每条判定都要回指来源**：`standard_source` 取 `template`（来自本块）/
   `config`（用户覆盖）/ `default`（通用惯例兜底）。
3. **`null` 不等于"没有要求"，只等于"没推出来"**。两种情形要分开表达：
   未设置（如 `superscript_size_ratio: null` + note 说明由 vertAlign 自动缩放）
   ≠ 没查到。

### 已知判不出来的项（须看渲染结果，别在 docx 层面硬猜）

| 项 | 为什么 | 怎么办 |
|---|---|---|
| 摘要/ABSTRACT 是否**真的**各占一页 | docx 里没有分页标记时，分页由内容长度和 Word 排版决定 | 导出 PDF 后看首页与页数 |
| 图题在图的**上方还是下方** | 图若是文本框/ASCII 框图，题注与图的从属关系在 XML 里看不出来 | 看渲染结果或问用户 |
| 页码是否**真的**从 1 开始 | `pgNumType/@start` 只写起始值，实际分节链接关系要看 `headerReference` | 看 PDF 页脚 |

## 消费契约（.template/ 目录）

spec.json 落位到项目根 `.template/` 后成为**唯一格式真值**：

- `build_docx.py` 读它生成 `template.docx`（派生物，可随时重建）；
- `init_template_dir.py` 读它派生 `constraints.md`（写作约束摘要，下游写作 agent 的消费入口）
  并维护 `manifest.json` 状态机（`draft → confirmed → applied`，`open_questions > 0` 时强制 draft）；
- 消费方（写作 / 排版环节）**只认 `.template/`**，且 manifest 状态必须非 draft。
- 修改排版：改 spec → 重新生成 → 校验 → 重跑 init_template_dir.py（不加 --force，只刷新派生物）。
