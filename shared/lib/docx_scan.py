#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
shared/lib/docx_scan.py —— 文档只读扫描器（包内所有审计型子 skill 的地基）

设计原则
    1. **只读**。本模块不做任何写操作，绝不改用户文档。
    2. **解析一次，多处复用**。Scanner(docx) 只建一次，各子 skill 共享同一实例，
       避免每个检查各自重复解析 docx（docx 解析很贵）。
    3. **文本偏移必须与"肉眼看到的文本"一致**。python-docx 的 `p.text` 只拼
       `p.runs`（`w:p` 的直接 `w:r` 子元素），会漏掉包在 `w:hyperlink` 里、或落在
       域代码结果里的 run —— 而引注恰恰经常在里面。所以这里自己按文档顺序遍历，
       hyperlink 与域结果都算进去。

两种引注序列化形态（都要认，别只认一种）
    A. `<w:hyperlink w:anchor="ref15"><w:r>…[15]…</w:r></w:hyperlink>`
       —— Word 自己生成的内部链接
    B. `HYPERLINK \l "ref15"` 域代码
       —— **WPS 保存时会把它从 A 改写成 B**，所以只看 A 会漏

用法
    import sys, os
    PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    sys.path.insert(0, PKG)
    from shared.lib.docx_scan import Scanner, norm, issue_id

    scan = Scanner(r"D:/论文.docx")
    for c in scan.citations():
        print(c.para_index, c.text, c.position, c.anchor)
"""
import hashlib
import os
import re
from dataclasses import dataclass, field
from typing import Optional

try:
    from docx import Document
except ImportError:
    raise SystemExit(
        "需要 python-docx。用模板提取那个 venv 跑：\n"
        '  "C:/Users/Tian/.workbuddy/binaries/python/envs/default/Scripts/python.exe"'
    )

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

# ---------------------------------------------------------------- 常量

REF_HEADINGS = {'参考文献', 'references', 'bibliography', '主要参考文献'}
ACK_HEADINGS = {'致谢', '致谢：', 'acknowledgement', 'acknowledgments'}
APPENDIX_PREFIX = '附录'
ABS_HEADINGS = {'摘要', '中文摘要', 'abstract', '英文摘要'}
TOC_HEADINGS = {'目录', '目次', 'contents'}

# 句末标点 / 句读符 —— 引注位置的判定基准
SENTENCE_END = '。！？!?'
CLAUSE_PUNCT = '；，、：,;:）)】」”'
# 引注**紧跟其后**即视为位置违规的分隔符。
# 刻意不含 `）】」”` 这类收尾符：`……"墙里墙外"[1]` 里引注跟在引号后完全正常，
# 把它们算进来会造成一片误报。只收真正的分句/分列符。
PUNCT_BEFORE = '。！？；，、：!?,;:'

# 一个引注号：[1] / ［1］ / (1) / （1） / 1 / [1,2] / [1-3]
CITE_RE = re.compile(r'^\s*([\[［【（(])?\s*(\d+(?:\s*[-–—~,，、]\s*\d+)*)\s*([\]］】）)])?\s*$')
# 域代码里的内部链接目标
HYPERLINK_REF_RE = re.compile(r'HYPERLINK\s+\\l\s+"?([^"\s\\]+)"?', re.I)
# 参考文献条目开头的编号
REF_NUM_RE = re.compile(r'^\s*[\[［]\s*(\d+)\s*[\]］]')
YEAR_RE = re.compile(r'[（(]\s*(19|20)\d{2}\s*[）)]')


# ---------------------------------------------------------------- 工具

def norm(s: Optional[str]) -> str:
    """标题/文本归一化：去所有空白（含全角空格 U+3000）后小写。

    必须做 —— 范文与论文里的标题常写作「摘　　要」「目　　录」「致　　谢」，
    不归一化就匹配不上（曾因此让参考文献扫描越界到致谢正文）。
    """
    return re.sub(r'[\s\u3000]+', '', s or '').lower()


def issue_id(rule: str, *fingerprint) -> str:
    """由 规则名 + 定位指纹 生成稳定 ID（形如 cite-form-0042）。

    不能用自增序号 —— 那样重跑一次 ID 全变，`.thesisignore` 里的人工裁定
    和 plan.yaml 里的勾选就全失效了。
    """
    h = hashlib.sha1(('|'.join(str(x) for x in fingerprint)).encode('utf-8')).hexdigest()
    return '%s-%04d' % (rule, int(h[:8], 16) % 10000)


def excerpt(text: str, start: int, end: int, pad: int = 18) -> str:
    """取包含 [start,end) 的一段上下文，两端用 … 表示截断。"""
    a, b = max(0, start - pad), min(len(text), end + pad)
    return ('…' if a > 0 else '') + text[a:b] + ('…' if b < len(text) else '')


def _run_text(r_el) -> str:
    return ''.join(t.text or '' for t in r_el.findall(W + 't'))


def _is_superscript(r_el) -> bool:
    rPr = r_el.find(W + 'rPr')
    if rPr is None:
        return False
    va = rPr.find(W + 'vertAlign')
    return va is not None and va.get(W + 'val') == 'superscript'


def _run_size_pt(r_el) -> Optional[float]:
    rPr = r_el.find(W + 'rPr')
    if rPr is None:
        return None
    sz = rPr.find(W + 'sz')
    if sz is None:
        return None
    v = sz.get(W + 'val')
    return (int(v) / 2.0) if v and v.isdigit() else None


# ---------------------------------------------------------------- 数据结构

@dataclass
class RunInfo:
    """段落内的一个可见 run，附带它在段落文本里的偏移与链接锚点。"""
    el: object
    text: str
    start: int
    end: int
    superscript: bool
    anchor: Optional[str] = None      # 内部链接目标：ref15 / _Toc123
    link_kind: Optional[str] = None   # 'hyperlink' | 'field' | None
    size_pt: Optional[float] = None


@dataclass
class Citation:
    """正文里的一处文献引注。"""
    para_index: int          # 1-based，与 doc.paragraphs 对齐，便于人工定位
    para_text: str
    text: str                # '[1]' / '1' / '（1）'
    refs: list               # [1] / [1,2] → 展开后的编号列表
    start: int
    end: int
    superscript: bool
    bracket: str             # 'square' | 'paren' | 'none'
    position: str            # before-punct | after-punct | mid-sentence | end-of-paragraph
    prev_char: str
    next_char: str
    anchor: Optional[str]    # 指向的书签名（ref15），None 表示没有链接
    link_kind: Optional[str]
    run_el: object

    @property
    def excerpt(self) -> str:
        return excerpt(self.para_text, self.start, self.end)


@dataclass
class RefEntry:
    """参考文献表里的一条。"""
    order: int               # 在表里的出现次序（1-based）
    number: Optional[int]    # 条目开头 [n] 里的 n；无编号为 None
    text: str
    para_index: int
    style_name: str
    bookmarks: list = field(default_factory=list)


@dataclass
class SectionInfo:
    index: int               # 1-based
    page_fmt: Optional[str]
    page_start: Optional[int]
    header_text: str
    header_has_rule: Optional[bool]
    rule_detail: Optional[dict] = None


# ---------------------------------------------------------------- 扫描器

class Scanner:
    def __init__(self, docx_path: str):
        self.path = os.path.abspath(docx_path)
        if not os.path.exists(self.path):
            raise FileNotFoundError(self.path)
        try:
            self.doc = Document(self.path)
        except Exception as e:
            raise RuntimeError(
                '无法打开 %s（%s）。若文件正被 WPS/Word 打开，请先关闭。' % (self.path, e))
        self._paras = list(self.doc.paragraphs)
        self._walk_cache = {}
        self._cites = None
        self._refsec = None
        self._sections = None
        self._bookmarks = None
        self._anchors_used = None
        self._body_text = None

    # -------------------------------------------------- 段落遍历

    def _walk(self, p_el):
        """按文档顺序遍历段落的可见 run，并跟踪域代码状态。

        返回 list[RunInfo]，其 start/end 累加出的文本**等于肉眼所见的段落文本**
        （包含 hyperlink 与域结果里的文字）。
        """
        key = id(p_el)
        if key in self._walk_cache:
            return self._walk_cache[key]

        out = []
        offset = 0
        # 域状态栈：每层 {'instr': str|None, 'in_result': bool}
        stack = []

        def push(el, anchor, kind):
            nonlocal offset
            t = _run_text(el)
            if not t:
                return
            out.append(RunInfo(el=el, text=t, start=offset, end=offset + len(t),
                               superscript=_is_superscript(el), anchor=anchor,
                               link_kind=kind, size_pt=_run_size_pt(el)))
            offset += len(t)

        def handle_run(el, anchor=None, kind=None):
            fld = el.find(W + 'fldChar')
            instr = el.find(W + 'instrText')
            if fld is not None:
                t = fld.get(W + 'fldCharType')
                if t == 'begin':
                    stack.append({'instr': None, 'in_result': False})
                elif t == 'separate':
                    if stack:
                        stack[-1]['in_result'] = True
                elif t == 'end':
                    if stack:
                        stack.pop()
                return                       # fldChar run 不显示文字
            if instr is not None:
                if stack and stack[-1]['instr'] is None:
                    stack[-1]['instr'] = instr.text or ''
                return                       # instrText 也不显示
            if stack and stack[-1]['in_result']:
                m = HYPERLINK_REF_RE.search(stack[-1]['instr'] or '')
                if m:
                    anchor, kind = m.group(1), 'field'
            push(el, anchor, kind)

        for child in p_el:
            tag = child.tag
            if tag == W + 'r':
                handle_run(child)
            elif tag == W + 'hyperlink':
                a = child.get(W + 'anchor')
                for r in child.findall(W + 'r'):
                    handle_run(r, a, 'hyperlink')
            elif tag in (W + 'ins', W + 'smartTag', W + 'sdt', W + 'sdtContent'):
                # 修订痕迹/内容控件：下钻一层，顺序不变
                for r in child.iter(W + 'r'):
                    handle_run(r)

        self._walk_cache[key] = out
        return out

    def runs(self, para_index: int):
        """按 1-based 段落序号取 RunInfo 列表。"""
        return self._walk(self._paras[para_index - 1]._element)

    def para_text(self, para_index: int) -> str:
        return ''.join(r.text for r in self.runs(para_index))

    def para_el(self, para_index: int):
        """取段落的 XML 元素。

        子 skill 造 Locator 时需要它。刻意不暴露 Paragraph 对象 ——
        审计阶段只该"读 + 造定位"，不该拿到能改写的句柄。
        """
        return self._paras[para_index - 1]._element

    # -------------------------------------------------- 引注

    @staticmethod
    def _classify(text: str):
        m = CITE_RE.match(text)
        if not m:
            return None, None
        op, body, cl = m.group(1), m.group(2), m.group(3)
        if (op and op in '[［【') or (cl and cl in ']］】'):
            bracket = 'square'
        elif (op and op in '（(') or (cl and cl in '）)'):
            bracket = 'paren'
        else:
            bracket = 'none'
        refs = [int(x) for x in re.split(r'[-–—~,，、]', body) if x.strip().isdigit()]
        return bracket, refs

    @staticmethod
    def _position_of(ptext: str, start: int, end: int) -> tuple:
        """判定引注相对标点的位置。返回 (position, prev_char, next_char)。

        判据（写死，不要凭感觉）：
          后一字符 ∈ 句读符（。！？；，、） → 引注落在标点**之前** → before-punct（规范写法）
          前一字符 ∈ 分隔符（。！？；，、：） → 引注**紧跟标点之后** → after-punct
          两侧都是正文 → mid-sentence（如「杨善华[1]指出」「路径[27]与…」）

        注意 after-punct 的判定要用 PUNCT_BEFORE 而不是 SENTENCE_END：
        `……统一；[4]吴帆认为……` 里的引注跟在分号后，同属位置违规（应为
        `……统一[4]；吴帆认为……`）。只认句末标点会把这些放过去。
        """
        nxt = ptext[end:end + 1]
        prev = ptext[max(0, start - 1):start]
        if nxt == '':
            # 段末引注：前面若是分隔符，本质就是"标点之后"，不能算另一种情况
            if prev in PUNCT_BEFORE:
                return 'after-punct', prev, nxt
            return 'end-of-paragraph', prev, nxt
        if nxt in SENTENCE_END or nxt in CLAUSE_PUNCT:
            return 'before-punct', prev, nxt
        if prev in PUNCT_BEFORE:
            return 'after-punct', prev, nxt
        return 'mid-sentence', prev, nxt

    def citations(self):
        """扫描全文正文里的引注。

        只扫「正文段落」—— 参考文献表、目录、页眉页脚里的 [n] 不算正文引注。
        """
        if self._cites is not None:
            return self._cites

        ref_head, _ = self.reference_section()
        ref_range = self._reference_para_range()
        toc_range = self._toc_para_range()

        found = []
        for pi, p in enumerate(self._paras, 1):
            if pi in ref_range or pi in toc_range:
                continue
            runs = self._walk(p._element)
            if not runs:
                continue
            ptext = ''.join(r.text for r in runs)
            for r in runs:
                tok = r.text.strip()
                if not tok or not re.search(r'\d', tok):
                    continue
                bracket, refs = self._classify(tok)
                if bracket is None or not refs:
                    continue
                # 非上标、又无括号、又只是一串数字 → 大概率是正文里的普通数字
                if bracket == 'none' and not r.superscript:
                    continue
                pos, prev, nxt = self._position_of(ptext, r.start, r.end)
                found.append(Citation(
                    para_index=pi, para_text=ptext, text=tok, refs=refs,
                    start=r.start, end=r.end, superscript=r.superscript,
                    bracket=bracket, position=pos, prev_char=prev, next_char=nxt,
                    anchor=r.anchor, link_kind=r.link_kind, run_el=r.el))
        self._cites = found
        return found

    # -------------------------------------------------- 参考文献表

    def _find_heading(self, names, limit=None):
        ns = {norm(x) for x in names}
        for i, p in enumerate(self._paras, 1):
            if limit and i > limit:
                break
            # 标题段不能太长，否则正文里提到「参考文献」四个字也会命中
            t = p.text.strip()
            if len(t) > 20:
                continue
            if norm(t) in ns:
                return i
        return None

    def reference_section(self):
        """返回 (标题段落号, [RefEntry])。找不到标题返回 (None, [])。"""
        if self._refsec is not None:
            return self._refsec
        h = self._find_heading(REF_HEADINGS)
        if h is None:
            self._refsec = (None, [])
            return self._refsec

        entries, order = [], 0
        for pi in range(h + 1, len(self._paras) + 1):
            p = self._paras[pi - 1]
            t = p.text.strip()
            if not t:
                continue
            nt = norm(t)
            if nt in {norm(x) for x in ACK_HEADINGS} or nt.startswith(APPENDIX_PREFIX):
                break
            # 长段落且开头没有 [n] —— 几乎不可能是条目（防误收正文）
            if len(t) > 250 and not REF_NUM_RE.match(t):
                continue
            order += 1
            m = REF_NUM_RE.match(t)
            entries.append(RefEntry(
                order=order,
                number=int(m.group(1)) if m else None,
                text=t,
                para_index=pi,
                style_name=(p.style.name if p.style is not None else ''),
                bookmarks=self.para_bookmarks(pi)))
        self._refsec = (h, entries)
        return self._refsec

    def _reference_para_range(self):
        h, entries = self.reference_section()
        if h is None:
            return set()
        end = entries[-1].para_index if entries else h
        return set(range(h, end + 1))

    def _toc_para_range(self):
        """目录区段（从「目录」标题到最后一个含制表位的条目）。粗略但够用。"""
        h = self._find_heading(TOC_HEADINGS)
        if h is None:
            return set()
        out, started = set(), False
        for pi in range(h, min(h + 400, len(self._paras) + 1)):
            p = self._paras[pi - 1]
            xml = p._element.xml or ''
            if 'w:leader="dot"' in xml or 'w:instrText' in xml or '\t' in (p.text or ''):
                started = True
                out.add(pi)
            elif started and not (p.text or '').strip():
                out.add(pi)
            elif started:
                break
        return out

    # -------------------------------------------------- 书签 / 链接

    def para_bookmarks(self, para_index: int):
        el = self._paras[para_index - 1]._element
        return [b.get(W + 'name') for b in el.findall('.//' + W + 'bookmarkStart')
                if b.get(W + 'name')]

    def bookmarks(self):
        """全文字符→书签名映射（按文档顺序）。"""
        if self._bookmarks is not None:
            return self._bookmarks
        names = []
        for b in self.doc.element.body.iter(W + 'bookmarkStart'):
            n = b.get(W + 'name')
            if n:
                names.append(n)
        self._bookmarks = names
        return names

    def anchors_used(self):
        """正文里所有内部链接指向的锚点及其出现次数。

        两种形态都收：w:hyperlink/@w:anchor 与 HYPERLINK \\l 域代码。
        """
        if self._anchors_used is not None:
            return self._anchors_used
        counts = {}
        for c in self.citations():
            if c.anchor:
                counts[c.anchor] = counts.get(c.anchor, 0) + 1
        self._anchors_used = counts
        return counts

    # -------------------------------------------------- 分节 / 页眉

    def sections(self):
        if self._sections is not None:
            return self._sections
        fmt_map = {'decimal': 'arabic', 'upperRoman': 'upperRoman',
                   'lowerRoman': 'lowerRoman', 'upperLetter': 'upperLetter',
                   'lowerLetter': 'lowerLetter', 'none': 'none'}
        out = []
        for i, s in enumerate(self.doc.sections, 1):
            sectPr = s._sectPr
            pg = sectPr.find(W + 'pgNumType') if sectPr is not None else None
            start = pg.get(W + 'start') if pg is not None else None
            txt, has_rule, detail = '', None, None
            try:
                for hp in s.header.paragraphs:
                    if hp.text.strip():
                        txt = hp.text.strip()
                    pPr = hp._element.find(W + 'pPr')
                    bdr = pPr.find(W + 'pBdr') if pPr is not None else None
                    bottom = bdr.find(W + 'bottom') if bdr is not None else None
                    if bottom is not None and bottom.get(W + 'val') not in (None, 'none', 'nil'):
                        has_rule = True
                        if detail is None:
                            sz = bottom.get(W + 'sz')
                            detail = {'style': bottom.get(W + 'val'),
                                      'size_pt': (int(sz) / 8.0) if sz and sz.isdigit() else None}
                    elif hp.text.strip():
                        has_rule = False
            except Exception:
                pass
            out.append(SectionInfo(
                index=i,
                page_fmt=fmt_map.get(pg.get(W + 'fmt')) if pg is not None else None,
                page_start=int(start) if start and start.isdigit() else None,
                header_text=txt, header_has_rule=has_rule, rule_detail=detail))
        self._sections = out
        return out

    # -------------------------------------------------- 其他

    def image_count(self):
        body = self.doc.element.body
        return len(list(body.iter(W + 'drawing'))) + len(list(body.iter(W + 'pict')))

    def body_text(self):
        if self._body_text is None:
            self._body_text = ''.join(self.doc.element.body.itertext())
        return self._body_text

    def style_names(self):
        """按文档顺序返回每个段落的样式名（供 L1 检查用）。"""
        return [p.style.name if p.style is not None else '' for p in self._paras]

    def summary(self):
        h, entries = self.reference_section()
        cites = self.citations()
        return {
            'path': self.path,
            'paragraphs': len(self._paras),
            'tables': len(self.doc.tables),
            'images': self.image_count(),
            'sections': len(self.doc.sections),
            'citations': len(cites),
            'citation_position_tally': _tally(c.position for c in cites),
            'reference_heading_para': h,
            'reference_entries': len(entries),
            'bookmarks': len(self.bookmarks()),
        }


def _tally(it):
    out = {}
    for x in it:
        out[x] = out.get(x, 0) + 1
    return out


# ---------------------------------------------------------------- CLI（自检用）

def main():
    import argparse
    import json
    ap = argparse.ArgumentParser(description='docx 只读扫描（自检/调试用）')
    ap.add_argument('-d', '--docx', required=True)
    ap.add_argument('--citations', action='store_true', help='列出每处引注')
    ap.add_argument('--refs', action='store_true', help='列出参考文献条目')
    a = ap.parse_args()

    scan = Scanner(a.docx)
    print(json.dumps(scan.summary(), ensure_ascii=False, indent=2))

    if a.citations:
        print('\n=== 引注（前 40 条）===')
        for c in scan.citations()[:40]:
            print('  段%-5d %-6s sup=%-5s %-16s anchor=%-8s %s'
                  % (c.para_index, c.text, c.superscript, c.position,
                     c.anchor or '-', c.excerpt))
    if a.refs:
        h, entries = scan.reference_section()
        print('\n=== 参考文献（标题在段 %s，共 %d 条）===' % (h, len(entries)))
        for e in entries[:10]:
            print('  #%-3d [%-3s] BM=%-8s %s'
                  % (e.order, e.number if e.number else '-',
                     ','.join(e.bookmarks) or '-', e.text[:60]))


if __name__ == '__main__':
    main()
