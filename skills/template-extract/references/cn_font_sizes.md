# 字号 / 字体 / 间距 对照参考

## 1. 中文字号 ↔ 磅值（pt）

中文论文规范用「号」，Word 内部用「磅」。换算必须准，否则样式会差一档。

| 字号 | 磅值 | | 字号 | 磅值 |
|---|---|---|---|---|
| 初号 | 42 | | 小三 | 15 |
| 小初 | 36 | | 四号 | 14 |
| 一号 | 26 | | 小四 | 12 |
| 小一 | 24 | | 五号 | 10.5 |
| 二号 | 22 | | 小五 | 9 |
| 小二 | 18 | | 六号 | 7.5 |
| 三号 | 16 | | 小六 | 6.5 |

> 换算已封装在 `scripts/common.py`：`cn_size_to_pt()` / `pt_to_cn_size()`。
> 样板：`一级标题 → 三号 → 16pt`，`正文 → 小四 → 12pt`，`表格 → 五号 → 10.5pt`。

## 2. 字体槽（这是中文渲染的关键）

Word/WPS 的一个 run 有 4 个字体槽，缺一不可：

| 槽位 | XML 属性 | 作用 |
|---|---|---|
| 西文 | `w:rFonts/@w:ascii` | 字母、数字 |
| 西文（高 ANSI） | `w:rFonts/@w:hAnsi` | 扩展拉丁字符 |
| **东亚** | `w:rFonts/@w:eastAsia` | **汉字、标点** |
| 复杂文种 | `w:rFonts/@w:cs` | 阿拉伯文等 |

⚠️ **python-docx 的 `font.name` 只写 ascii/hAnsi，不写 eastAsia**。
中文论文里「汉字宋体 + 字母数字 Times New Roman」这种混排，必须显式写 eastAsia，
否则汉字会跟随默认字体渲染。本项目统一由 `common.set_east_asian_font()` 处理。

## 3. PDF 内嵌中文字体名反查

PDF 里的字体名常带子集前缀（`ABCDEF+SimSun`），且不同发行版命名不一。
`scripts/extract_from_pdf.py` 内置了以下别名表，遇到未收录的需人工补：

| PDF 中可能出现 | 实际字体 |
|---|---|
| `SimSun` / `NSimSun` / `STSong` / `Songti` | 宋体 |
| `SimHei` / `STHei` / `Heiti` | 黑体 |
| `SimKai` / `KaiTi` / `STKaiti` | 楷体 |
| `FangSong` | 仿宋 |
| `MicrosoftYaHei` / `MSYH` | 微软雅黑 |

**判断技巧**：若字体名是纯 6 位大写字母 + 短名（如 `ABCDEF+TT1234`），说明子集化严重，
字体不可靠；此时**以模板文字规范（若有）为准**，或向用户确认。

## 4. 间距语义

| 概念 | Word 表示 | 常见取值 |
|---|---|---|
| 倍数行距 | `w:spacing/@w:line` = 240 × n，`@w:lineRule="auto"` | 1.5 倍 = 360 |
| 固定值行距 | `@w:line` = pt × 20，`@w:lineRule="exact"` | 20pt = 400 |
| 最小值行距 | `@w:lineRule="atLeast"` | — |
| 段前 / 段后 | `@w:before` / `@w:after`，单位 1/20 pt | 12 磅 = 240 |
| 首行缩进 N 字符 | `w:ind/@w:firstLineChars` = N × 100 | 2 字符 = 200 |
| 左缩进 N 字符 | `w:ind/@w:leftChars` = N × 100 | 1 字符 = 100 |

⚠️ **首行缩进要用 `firstLineChars` 而不是 `firstLine`**。
`firstLine` 是固定磅值，换字号后缩进会跟着变；`firstLineChars` 才是「2 个字符」的语义。

## 5. 长度换算

- 1 磅（pt） = 1/72 英寸
- 1 英寸 = 2.54 厘米
- A4 = 21.0 × 29.7 cm；Letter = 21.59 × 27.94 cm
- 1 twip = 1/20 pt（OOXML 里大量使用 twip）

## 6. 表格线宽

三线表的磅值在 XML 里用 `w:sz`，单位是 **1/8 磅**：

| 目标线宽 | `w:sz` |
|---|---|
| 0.5 磅（中线 / 细线） | 4 |
| 1.0 磅 | 8 |
| 1.5 磅（上下线 / 粗线） | 12 |
| 2.25 磅 | 18 |
