# SOP｜引注上标与跳转（上标形态 · Ctrl+单击跳到参考文献表）

> **跨 skill 的 SOP。** 本文讲的是"这件事怎么做对、会踩哪些坑、换一篇论文要改什么"，
> 不是"本包里有哪些参数"——那部分见 `skills/cite-doctor/SKILL.md`（规则与 CLI）与
> `skills/template-extract/references/format_spec_schema.md`（spec 字段）。
>
> 适用对象：正文用**顺序编码制**（`[1]`、`[2]`…）的中文论文，要求序号"上标"显示，
> 且在 Word/WPS 中 Ctrl+单击正文序号能跳到文末参考文献表对应条目。
>
> 涉及：`cite-doctor`（判定）+ `shared/lib/docx_ops.py`（6 个写动作）+ `workflow/apply.py`（落盘）。

---

## 一、先把体例量化，再动手

不要凭印象判断"上标对不对"。**从学校范文里取基准**，用同一把尺子验收。

被引用项目（本包的来源项目）的基准取自两份已通过答辩的范文，实测：

| 项目 | 范文1 | 范文2 | 体例结论 |
|---|---|---|---|
| 标注形式 | `[1]` 保留方括号 | 同 | **上标 + 保留方括号** |
| 字号比（上标/正文） | 0.65 | 0.67 | ≈ Word 默认上标（0.65） |
| 基线抬高 | 1.87pt | 1.40pt | 跟随字号自动计算，不必手设 |
| 上标字体 | TimesNewRomanPSMT | 同 | 数字走西文字体、正文走宋体，**这是正常体例** |
| 颜色／下划线 | 黑色、无下划线 | 同 | 不能出现超链接默认的蓝字下划线 |

**取基准的脚本**（可复用于任何论文）：

```python
import fitz, re
d = fitz.open("范文.pdf")
for pi in range(d.page_count):
    for b in d[pi].get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            sp = ln["spans"]
            for i, s in enumerate(sp):
                if re.fullmatch(r"\[\d{1,2}\]", s["text"].strip()) and i > 0:
                    prev = sp[i - 1]
                    print(s["size"] / prev["size"],        # 字号比
                          prev["bbox"][1] - s["bbox"][1],  # 基线抬高
                          s["font"], s["color"], s["flags"])
```

> **量出来的东西才是标准。** 这正是本包"标准来源三级优先"里 `template` 那一级的由来，
> 详见 `SOP-模板标准.md`。

---

## 二、两条实现路径：本例走哪条

| 路径 | 何时用 | 做法 |
|---|---|---|
| **A. 排版期生成** | 你在"从 0 搭初稿"（md → docx 建通道），只有一次性成本 | md 里**保留 `[n]` 作为语义标记**，排版脚本渲染成上标 + 内部超链接 |
| **B. docx 后处理**（本包默认） | 你手上只有一份 docx（绝大多数人的情况），且人工还会继续改它 | `audit` 找出不合规处 → `plan` 勾选 → `apply` 定点手术 |

**为什么 B 是本包的默认**：本包的铁律是 **docx 是唯一真值源**。人工微调改的是 docx，
如果靠"重跑排版脚本"来修上标，就会把人工改动整个覆盖掉。A 只适合建通道那一次。

> 路径 A 的三点好处（当时选它的理由）值得记住，它是本包"建通道"阶段的推荐做法：
> ① md 是底本，语义不丢，体检门禁照旧；② 每次重排自动重建，**不可能漂移**；
> ③ 参考文献表重排序号时，正文标注跟着一起走，无需手工对齐。
> 核心原则：**让 md 只表达"这是第 n 条引用"，把"长什么样、能不能点"全部交给排版层。**

---

## 三、路径 B（后处理）怎么做 —— 三个动作覆盖全部问题

`cite-doctor` 把问题分成三类，各自挂**真实存在**的动作（动作名取自 `shared/lib/docx_ops.py` 的 `ACTIONS`）：

| 问题 | rule | 动作 | 落到哪 |
|---|---|---|---|
| 形态不是上标 | `cite-form` | `cite.superscript` | 该引注 run 加 `vertAlign=superscript` |
| 缺方括号 / 括号形态不对 | `cite-form` | `cite.rewrite` | 重写 `[n]` 文本 |
| 上标上还显式写了 `w:sz` | `cite-form` | `run.clear_size` | 删掉显式字号（否则渲染器缩两次） |
| 文献表条目没挂书签 | `cite-jump` | `ref.bookmark` | 给条目加书签名 `refN` |
| 正文锚点悬空 / 错指 / 纯文字无链接 | `cite-jump` | `cite.link` | 补/改 `w:hyperlink w:anchor="refN"` |
| 引注落在句读符之后 | `cite-position` | `cite.move_before_punct` | 把引注移到句读符前 |

**不给自动动作的（刻意的，不是没做）**：

- **重编号**（`cite-numbering`）、**删重复条目**（`ref-duplicate`）、**著录格式**（`ref-format`）：
  这三件事都是"一处改动牵动全篇"的全局手术——删一条要连带把它后面所有引注重编号并重挂链接；
  而补页码/期号要看原文献，机器擅自补等于**造假**。所以只报不修。
- **多编号合并引用**（一处含多个编号）：给 warn 且不自动拆。

**落盘流程**（顺序不可颠倒：先哈希门禁，后回填参数）：

```bash
VPY="python"
PDD="F:/Coding/Project/paper-detail-doctor"

"$VPY" "$PDD/workflow/audit_all.py" -d 我的论文.docx -s spec.json          # ① 只读诊断
"$VPY" "$PDD/workflow/plan.py"      -d 我的论文.docx                      # ② 生成 plan.yaml，手工勾选
"$VPY" "$PDD/workflow/apply.py"     -p plan.yaml --dry-run                # ③ 先预览
"$VPY" "$PDD/workflow/apply.py"     -p plan.yaml --yes                    # ④ 执行（先快照，可 --rollback）
```

Apply 的三条硬约束（都在 `shared/lib/ooxml_guard.py` 里实现）：**幂等**（已做过的跳过）、
**最小侵入**（只动目标 OOXML 节点，且**空操作不重写文件**）、**可回滚**（写入前快照到 `_backup/<时间戳>/`）。
三者缺一，长时间反复使用（每次编稿后重跑一遍）就必然出问题。

---

## 四、坑位（都真实踩过）

### 坑 1｜参考文献表自身的 `[1]` 不是引注

条目行首的 `[1]` 是**列表编号**。若把它也一并转上标并挂链接，会生成**指向自身的循环超链接**，
超链接数变成 2×N（N=46 → 92）。定位方法：`Counter(链接锚点)` 显示每个编号出现 2 次。

- **路径 A 的解法**：给统一写入函数加 `cites` 开关，参考文献条目一律 `cites=False`；
  书签也要单独加（`add_bookmark(p, "ref" + n)`）。
- **路径 B 的解法**：靠**位置判定**——引注在正文、条目在文献表区段。`cite-doctor` 先扫描出
  参考文献表起止，表内段落不参与 `cite-form` 判定。

### 坑 2｜WPS/Word 保存后会把 `w:hyperlink` 改写成域代码 —— 校验必须两种形态都认

写完 docx 后只要用 WPS 打开并保存过（比如"更新目录"那一步），WPS 会重写 `document.xml`，把

```xml
<w:hyperlink w:anchor="ref1"><w:r>…<w:t>[1]</w:t></w:r></w:hyperlink>
```

改写成等价但形式完全不同的 HYPERLINK 域代码：

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText> HYPERLINK \l "ref1" </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:rPr>…<w:vertAlign w:val="superscript"/></w:rPr><w:t>[1]</w:t></w:r>
<w:r><w:rPr>…<w:vertAlign w:val="superscript"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
```

三个连带后果，一个比一个坑：

1. **只按 `<w:hyperlink w:anchor=...>` 匹配的校验会误报"超链接 0 个"**，其实功能完好、能点。
2. **`vertAlign=superscript` 的出现次数翻倍**（46 → 92）——WPS 在 `fldCharType="end"` 的 run 上
   复制了同一份 rPr。数上标时必须**过滤掉没有 `<w:t>` 文字的 run**。
3. WPS 会顺手给目录加 `_Toc*` 书签（本项目 59 个）和 1 个 `_GoBack`，所以
   `书签数` 与 `Hyperlinks.Count` 都会**大于**文献条数。**不要按等号断言**，要按 `ref` 前缀筛。

> **结论：功能不会被破坏，但校验必须兼容两种序列化形态。** 本包的 `docx_scan.py` 读侧
> 同时识别 `w:hyperlink/@anchor` 与 `HYPERLINK \l` 域代码，所以流程中无论有没有被 WPS 保存过都能跑。

### 坑 3｜`xml:space` 不能用命名空间快捷写法

lxml 要求完整命名空间 URI，`t.set(q("xml:space"), "preserve")` 会抛
`ValueError: Invalid attribute name`。正确写法：

```python
t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
```

### 坑 4｜`<w:pBdr>`（页眉横线）必须排在 `<w:pStyle>` 之后

OOXML 的 `pPr` 子元素有严格顺序（`pStyle → … → pBdr → shd → …`）。
直接 `s.replace("<w:pPr>", "<w:pPr><w:pBdr>…")` 插到开头会破坏顺序——WPS 可能容忍但不可靠。
先 `re.search(r"<w:pStyle\b[^>]*/>\s*")` 再在其后插入。

### 坑 5｜写回 docx 时文件可能被预览面板占用

`os.replace` 会抛 `PermissionError`（即便 `os.rename` 能成功）。本包的处理是
「重试 5 次 → 旧文件改名挪走 → 新文件落位 → 删旧文件」的组合写法，**不可去掉这个回退**，
否则流水线会在最没道理的地方中断。

### 坑 6｜上标不要显式设字号

`w:vertAlign="superscript"` 已让渲染器把字号缩到约 65%，**再写 `w:sz` 会被缩两次**。
所以 `extract_conventions` 实测出"范文没有显式 w:sz"时，那本身就是正确答案，
期望状态就是"不设"——`standard_source` 可以标 `template`。

---

## 五、验收（三步，缺一不可）

```bash
# 1) 结构：一一对应、无孤儿、编号连续、外观干净
python skills/cite-doctor/scripts/audit.py -d 我的论文.docx --print
#    期望：引注数 == 书签数 == 上标数，无孤儿，1..N 连续

# 2) 几何：与范文基准比对（取数脚本见第一节）
#    期望：字号比 ≈0.65、基线抬高 1.4–1.9pt、颜色为黑、无下划线

# 3) 交互：Word/WPS 打开 → 按住 Ctrl 点正文某个 [n]
#    期望：光标跳到文末参考文献表第 n 条
```

**为什么第 3 步不能省**：XML 写对了 ≠ 能点，也不等于渲染正确。
只有 Word/WPS 的对象模型承认它，Ctrl+单击才真的可用。会被 WPS 保存过的文档，
必须用 COM 读 `Hyperlinks.Count` / `Bookmarks.Count` / 逐条 `SubAddress` 才能确认。

`cite-doctor` 的对应规则与严重度：

| 规则 | 查什么 | 失败含义 |
|---|---|---|
| `cite-position` | 引注在句读符前 | `after-punct` 要改；段末无句末标点仅提示 |
| `cite-form` | 上标 + 方括号 + 不显式设字号 | 有引注不是上标形态 / 括号丢了 / 字号会被缩两次 |
| `cite-jump` | 正文锚点 ↔ 文献书签两个方向 | 条目没书签 / 锚点悬空 / 错指 / 纯文字无链接 |
| `cite-numbering` | 1..N 连续、首现顺序 | 断号、重复、顺序与正文首现不符 |

---

## 六、迁移到另一篇论文的改动清单

1. **取基准**：用第一节的脚本跑目标学校 1–2 份官方范文，确定"是否保留方括号、字号比"，
   把结果写进 `spec.json`（或交给 `template-extract` 自动提取，见 `SOP-模板标准.md`）。
2. **定正则**：多数情况是 `\[(\d+)\]`；若是 `[1-3]` 区间或多组并列，需要扩展模式。
   注意本包对**多编号合并引用**只给 warn、不自动拆。
3. **确认引注位置体例**：`before-punct`（句读符前）还是 `after-punct`。若 spec 里写了非
   `before-punct`，`cite-position` 规则**直接跳过不报**（不拿别的学校的体例套你）。
4. **参考文献表**：确认每条以 `[n]` 开头。书签前缀（`ref`，即 `ref1…refN`）必须与正文锚点前缀
   一致——两处改动必须同改，否则链接全部悬空。
5. **跑一遍验收三步**，尤其第 3 步（Ctrl+单击）。

---

## 七、本包与它的来源项目

这套做法源自本包的**来源项目**（一篇以 md → Word 流水线排版的学位论文）。该项目自己留了一份
同题 SOP，是本文的直系前身，其执行记录如下：两份范文均为「上标 + 保留方括号」，
实现是在统一的文本写入函数上新增 `add_bookmark` / `add_citation_hyperlink`，并给写入函数
加 `cites` 开关；踩到自链接 92→46 与 `xml:space` 命名空间两个坑；四项校验全 PASS，
WPS 读到 `Hyperlinks=46 / Bookmarks=46`，几何 0.65／1.85pt 与范文1 一致。

**本包把它泛化成了通用能力**：该项目走的是"路径 A（排版期生成）"，本包把同一套体例判定
与同一批坑位做成 `cite-doctor` 的 7 条只读规则 + 6 个可执行动作，因此**手上只有 docx 的人也能用**。

---

## 与其他文档的分工

| 想了解 | 去哪看 |
|---|---|
| 本 skill 有哪些 rule id、CLI 怎么调、产物在哪 | `skills/cite-doctor/SKILL.md` |
| spec 里引注相关字段怎么写 | `skills/template-extract/references/format_spec_schema.md` |
| 标准从哪来、实测值为什么不能当上限 | `SOP-模板标准.md` |
| 开题与正文对不对得上 | `SOP-开题一致性核对.md` |
