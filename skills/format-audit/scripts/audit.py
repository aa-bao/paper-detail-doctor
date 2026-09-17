#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
skills/format-audit/scripts/audit.py —— L1 格式与版式审计（**只读，绝不写盘**）

六条规则（与 assets/skills-index.yaml 的 covers 一一对应）
    style-compliance  命名样式合规 + 直接格式硬改
    page-setup        纸张与页边距
    page-numbering    分节页码（pgNumType 格式/起始、是否缺页码）
    header            页眉文字与横线
    fonts             正文字体/字号/行距
    toc               目录域（是否存在/失效、条目数 vs 正文标题数）

设计上的保守点（不是没做，是故意不做 / 故意降级）
    ① page_margins_cm 为 None 时不判页边距，只在 meta.notes 写"因缺标准未判"。
       页边距是典型的"范文有值 ≠ 规范有规定"项，宁可少报不可瞎报。
    ② header_rule 为 None 时不判"该不该有横线"，只报各节**不一致**。
    ③ 字体字号：前 80 段实测 run 级**完全没有 w:sz**，字号全部来自命名样式
       —— 所以必须沿样式链解析（paragraph.style.font + 递归 base_style），
       段落直接格式优先于样式。凡同一层级出现 2 种以上字号/字体 → warn。
    ④ 判不出来的必须如实说：目录"是否已更新"、页眉与封面是否一致这类需渲染
       后才能对比的项，一律写进 meta.notes「需渲染确认，docx 层面判不出来」，
       绝不瞎报"目录是最新的""页眉与封面一致"。

标准来源三级优先（每条 Issue 都带 standard_source，报告里显著标出）
    template（模板/范文 spec） > config（用户配置） > default（通用学术惯例）
    没给 spec 时全部落 default，并在报告顶部明写"未提供模板标准"—— 不能让人
    误以为"正文小四"是学校的规定。

用法
    python audit.py -d "F:/论文/初稿.docx"
    python audit.py -d 初稿.docx -s spec.json -o out/
    python audit.py -d 初稿.docx --print            # 直接把报告打到终端
"""
import argparse
import json
import os
import re
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
if PKG not in sys.path:
    sys.path.insert(0, PKG)

from shared.lib.docx_ops import make_locator                     # noqa: E402
from shared.lib.docx_scan import (Scanner, W, excerpt, issue_id, norm)  # noqa: E402
from shared.lib.standards import (                                # noqa: E402
    describe, find_spec, load_spec, resolve_standard,
)
from shared.lib import report as R                                # noqa: E402

SKILL = 'format-audit'
LAYER = 'L1'

# 纸张名 → (宽, 高) cm。EMU→cm：1cm = 360000 EMU（python-docx 的
# section.page_width/page_height/margins 返回的就是 EMU）。
PAGE_SIZES = {
    'A4': (21.0, 29.7), 'A3': (29.7, 42.0), 'A5': (14.8, 21.0),
    'B5': (17.6, 25.0), 'Letter': (21.59, 27.94), '16K': (18.4, 26.0),
}
EMU_PER_CM = 360000.0

# pgNumType/@fmt → 人类可读
FMT_MAP = {'decimal': 'arabic', 'upperRoman': 'upperRoman',
           'lowerRoman': 'lowerRoman', 'upperLetter': 'upperLetter',
           'lowerLetter': 'lowerLetter', 'none': 'none', None: None}

# 前置部分标题/正文 hint（用于把"含摘要/目录的节"判为前置）
FRONT_STYLE_HINT = ('摘要', '目录', 'abstract')
TOC_HEADINGS = {'目录', '目次', 'contents'}


# ================================================================ 上下文

class Ctx:
    """一次审计的全部共享状态。规则函数只读它，不互相传参。"""

    def __init__(self, docx, spec_path=None):
        self.docx = os.path.abspath(docx)
        self.scan = Scanner(self.docx)
        self.spec, self.spec_path = load_spec(spec_path)
        self.std, self.src, self.notes = resolve_standard(self.spec)
        self.issues = []
        self._ranges = _section_ranges(self.scan.doc)
        self._is_front = _section_kinds(self.scan.doc, self._ranges)

    # -------------------------------------------------- 定位辅助

    def sec_first(self, sec_index):
        """某节第一个段落的 1-based 序号（用于把问题定位到具体位置）。"""
        rng = self._ranges.get(sec_index) or []
        return rng[0] if rng else None

    def para_index_of(self, p_el):
        for i, p in enumerate(self.scan.doc.paragraphs, 1):
            if p._element is p_el:
                return i
        return None

    def act(self, action_name, **params):
        """造一个动作。

        ★ 第一个参数**不能**叫 `name`（同 cite-doctor 的坑）：
        动作里恰好也有个 `name` 键，参数撞名会让 Python 抛
        `got multiple values for argument 'name'`，被单规则 try/except 吞掉后
        整条规则静默失效。
        """
        return {'name': action_name, 'params': params}

    def add(self, rule, severity, *, key, expected, actual, suggestion,
            source_key, suppressible=True, action=None,
            para_index=None, text=None, start=None, end=None, kind='paragraph'):
        if text is not None and start is not None:
            ex = excerpt(text, start, end)
        elif text:
            ex = text[:46] + ('…' if len(text) > 46 else '')
        else:
            ex = None
        loc = {'kind': kind, 'index': para_index, 'excerpt': ex,
               'offset': start if kind == 'run' else None}
        src = self.src.get(source_key)
        if src is None:                       # 防御：理论上 source_key 必在 src 里
            src = 'default'
        self.issues.append({
            'id': issue_id(rule, key),
            'rule': rule, 'layer': LAYER, 'skill': SKILL,
            'severity': severity, 'location': loc,
            'expected': expected, 'actual': actual,
            'standard_source': src,
            'suggestion': suggestion, 'suppressible': suppressible,
            'action': action,
        })


# ================================================================ 样式/字体解析

def _rpr_sz(rpr):
    """rPr 里的字号（半磅 → pt）。返回 None 表示未直接设置。"""
    if rpr is None:
        return None
    for tag in (W + 'sz', W + 'szCs'):
        el = rpr.find(tag)
        if el is not None:
            v = el.get(W + 'val')
            if v and v.isdigit():
                return int(v) / 2.0
    return None


def _rpr_fonts(rpr):
    """rPr 里的字体（ascii, eastAsia）。"""
    if rpr is None:
        return (None, None)
    rf = rpr.find(W + 'rFonts')
    if rf is None:
        return (None, None)
    return (rf.get(W + 'ascii') or rf.get(W + 'cs'), rf.get(W + 'eastAsia'))


def _para_direct_rpr(p_el):
    pPr = p_el.find(W + 'pPr')
    if pPr is None:
        return None
    return pPr.find(W + 'rPr')


def _resolve_style_size(style, doc):
    seen = set()
    s = style
    while s is not None and id(s) not in seen:
        seen.add(id(s))
        rpr = s._element.find(W + 'rPr')
        h = _rpr_sz(rpr)
        if h is not None:
            return h
        b = s._element.find(W + 'basedOn')
        s = doc.styles[b.get(W + 'val')] if b is not None else None
    return None


def _resolve_style_font(style, doc):
    seen = set()
    s = style
    while s is not None and id(s) not in seen:
        seen.add(id(s))
        rpr = s._element.find(W + 'rPr')
        a, e = _rpr_fonts(rpr)
        if a or e:
            return (a, e)
        b = s._element.find(W + 'basedOn')
        s = doc.styles[b.get(W + 'val')] if b is not None else None
    return (None, None)


def para_eff_size(para, doc):
    """段落有效字号（pt）。优先级：段直接 rPr > 任意 run 直接 rPr > 样式链。"""
    p_el = para._element
    sz = _rpr_sz(_para_direct_rpr(p_el))
    if sz is not None:
        return sz
    for r in para.runs:
        rsz = _rpr_sz(r._r.find(W + 'rPr'))
        if rsz is not None:
            return rsz
    if para.style is not None:
        st = _resolve_style_size(para.style, doc)
        if st is not None:
            return st
    return None


def para_eff_font(para, doc):
    """段落有效中文字体名。优先级同上。"""
    p_el = para._element
    a, e = _rpr_fonts(_para_direct_rpr(p_el))
    if a or e:
        return (a, e)
    for r in para.runs:
        ra, re_ = _rpr_fonts(r._r.find(W + 'rPr'))
        if ra or re_:
            return (ra, re_)
    if para.style is not None:
        st = _resolve_style_font(para.style, doc)
        if st is not None:
            return st
    return (None, None)


# 哪些样式算"正文段"（参与字体/字号一致性检查）
_NON_BODY = {'toc 1', 'toc 2', 'toc 3', 'toc 4', 'toc 5', 'toc 6', 'toc 7', 'toc 8', 'toc 9'}
_SPECIAL = {'参考文献条目', '参考文献标题', '表题', '图片来源', '关键词',
            '英文关键词', '摘要标题', '英文摘要标题', '目录标题', '致谢标题'}


def is_body_style(name):
    if not name:
        return True
    if name.lower().startswith('toc'):
        return False
    if name in _NON_BODY:
        return False
    if '标题' in name and name not in {'摘要标题', '英文摘要标题', '目录标题',
                                       '参考文献标题', '致谢标题'}:
        # 章节标题（一级/二级/三级标题）不算正文
        return False
    if name in _SPECIAL:
        return False
    return True


def is_chapter_heading(name):
    """章节标题样式（用于目录条目数比对、误用标题样式检测）。"""
    if not name:
        return False
    if name.lower().startswith('toc'):
        return False
    return bool(re.match(r'^[一二三四五六]级标题$', name or '')) or name in {
        'h1', 'h2', 'h3', 'h4', 'h5', 'h6'}


# ================================================================ 分节 / 页眉 / 目录

def _section_ranges(doc):
    """1-based 段落序号 → 节序号。靠段落 pPr 里的 sectPr（节分隔）切分。"""
    n = len(doc.sections)
    breaks = []
    for i, p in enumerate(doc.paragraphs, 1):
        pPr = p._element.find(W + 'pPr')
        if pPr is not None and pPr.find(W + 'sectPr') is not None:
            breaks.append(i)                       # 该段结束一节，下一段起新节
    bounds = [0] + breaks + [len(doc.paragraphs)]
    ranges = {}
    for s in range(1, n + 1):
        ranges[s] = list(range(bounds[s - 1] + 1, bounds[s] + 1))
    return ranges


def _is_front_para(p):
    nm = (p.style.name if p.style else '') or ''
    if any(k in nm for k in FRONT_STYLE_HINT):
        return True
    t = norm(p.text or '')
    if t.startswith('摘要') or t.startswith('目录') or t.startswith('abstract'):
        return True
    return False


def _section_kinds(doc, ranges):
    """每节是否为前置部分（含摘要/目录的节为前置，最后一节为正文）。"""
    n = len(doc.sections)
    is_front = {}
    for s in range(1, n + 1):
        front = any(_is_front_para(doc.paragraphs[pi - 1])
                    for pi in ranges.get(s, []))
        is_front[s] = front
    return is_front


def _footer_has_page(sec):
    """页脚里是否有 PAGE 域（w:instrText 含 PAGE）。"""
    try:
        for p in sec.footer.paragraphs:
            if 'PAGE' in (p._element.xml or ''):
                return True
    except Exception:
        pass
    return False


def _header_text_of(sec):
    try:
        for hp in sec.header.paragraphs:
            if (hp.text or '').strip():
                return hp.text.strip()
    except Exception:
        pass
    return ''


def _toc_region(doc):
    """返回 (目录标题段号, 目录区段集合)。没有目录标题返回 (None, set())。"""
    h = None
    for i, p in enumerate(doc.paragraphs, 1):
        if norm(p.text or '') in {norm(x) for x in TOC_HEADINGS} and len(p.text or '') <= 20:
            h = i
            break
    if h is None:
        return None, set()
    out, started = set(), False
    for pi in range(h, min(h + 400, len(doc.paragraphs) + 1)):
        p = doc.paragraphs[pi - 1]
        xml = p._element.xml or ''
        if 'w:leader="dot"' in xml or 'w:instrText' in xml or '\t' in (p.text or ''):
            started = True
            out.add(pi)
        elif started and not (p.text or '').strip():
            out.add(pi)
        elif started:
            break
    return h, out


def _run_offset(p_el, r_el):
    """run 在段落文本里的 [start, end) 字符偏移（用于造 run 级定位）。"""
    off = 0
    for child in p_el:
        if child.tag == W + 'r':
            t = ''.join(x.text or '' for x in child.findall(W + 't'))
            if child is r_el:
                return off, off + len(t)
            off += len(t)
        elif child.tag == W + 'hyperlink':
            for rr in child.findall(W + 'r'):
                t = ''.join(x.text or '' for x in rr.findall(W + 't'))
                if rr is r_el:
                    return off, off + len(t)
                off += len(t)
    return 0, 0


# ================================================================ 规则

def rule_page_setup(ctx):
    """纸张与页边距。std['page_margins_cm'] 为 None 时不判页边距。"""
    doc = ctx.scan.doc
    want = ctx.std['page_size']
    expect_wh = PAGE_SIZES.get(want)
    for i, sec in enumerate(doc.sections, 1):
        # ---- 纸张
        if expect_wh is not None:
            w_cm = (sec.page_width or 0) / EMU_PER_CM
            h_cm = (sec.page_height or 0) / EMU_PER_CM
            ew, eh = expect_wh
            if abs(w_cm - ew) > 0.05 or abs(h_cm - eh) > 0.05:
                ctx.add('page-setup', 'error', key='page-size-%d' % i,
                        expected='纸张应为 %s（%.1f×%.1fcm）' % (want, ew, eh),
                        actual='第 %d 节纸张为 %.2f×%.2fcm' % (i, w_cm, h_cm),
                        suggestion='在「页面布局→纸张大小」中选 %s。本工具不自动改分节'
                                   '页面设置（避免误伤其他节）。' % want,
                        source_key='page_size', kind='document',
                        para_index=ctx.sec_first(i), action=None)
        # ---- 页边距（仅在给了标准时判）
        mc = ctx.std['page_margins_cm']
        if mc:
            top, right, bottom, left = (list(mc) + [None] * 4)[:4]
            actual = (sec.top_margin, sec.right_margin, sec.bottom_margin, sec.left_margin)
            names = ['上', '右', '下', '左']
            for j, (exp, emu) in enumerate(zip([top, right, bottom, left], actual)):
                if exp is None or emu is None:
                    continue
                act_cm = emu / EMU_PER_CM
                if abs(act_cm - exp) > 0.05:
                    ctx.add('page-setup', 'error', key='page-margin-%d-%d' % (i, j),
                            expected='第 %d 节页边距%s边=%.2fcm（来自模板）' % (i, names[j], exp),
                            actual='第 %d 节页边距%s边为 %.2fcm' % (i, names[j], act_cm),
                            suggestion='在「页面布局→页边距」中调整%s边为 %.2fcm。'
                                       '本工具不自动改分节页面设置。' % (names[j], exp),
                            source_key='page_margins_cm', kind='document',
                            para_index=ctx.sec_first(i), action=None)


def rule_page_numbering(ctx):
    """逐节页码格式/起始；正文节是否缺页码；无标准时只做自洽检查。"""
    doc = ctx.scan.doc
    n = len(doc.sections)
    secs = []
    for i in range(1, n + 1):
        sec = doc.sections[i - 1]
        sectPr = sec._sectPr
        pg = sectPr.find(W + 'pgNumType') if sectPr is not None else None
        has_pgnt = pg is not None
        fmt = FMT_MAP.get(pg.get(W + 'fmt')) if pg is not None else None
        start = pg.get(W + 'start') if pg is not None else None
        has_page_field = _footer_has_page(sec)
        secs.append((i, has_pgnt, fmt, start, has_page_field))
    is_front = ctx._is_front
    scheme = ctx.std['page_numbering_scheme']

    # 1) 每节是否"有页码"（pgNumType 或页脚 PAGE 域）
    for (i, has_pgnt, fmt, start, has_page_field) in secs:
        has_pg = has_pgnt or has_page_field
        if not has_pg:
            if is_front.get(i):
                ctx.add('page-numbering', 'warn', key='no-pagenum-front-%d' % i,
                        expected='前置部分（第 %d 节）应明确设置页码（pgNumType 或页脚 PAGE 域）' % i,
                        actual='第 %d 节既无 pgNumType 也无页脚 PAGE 域，页码方案不明确' % i,
                        suggestion='如需前置部分用罗马数字，请在节属性设置 pgNumType '
                                   'fmt=lowerRoman start=1；如不要页码也请确认符合学校要求。',
                        source_key='page_numbering_scheme', kind='footer',
                        para_index=ctx.sec_first(i), action=None)
            else:
                ctx.add('page-numbering', 'error', key='no-pagenum-body-%d' % i,
                        expected='正文（第 %d 节）应有页码' % i,
                        actual='第 %d 节既无 pgNumType 也无页脚 PAGE 域，正文没有页码' % i,
                        suggestion='在页脚插入 PAGE 域，或在节属性设置 pgNumType。',
                        source_key='page_numbering_scheme', kind='footer',
                        para_index=ctx.sec_first(i), action=None, suppressible=False)
        # 2) 有标准时逐节比对格式
        if scheme and fmt is not None and fmt != scheme:
            ctx.add('page-numbering', 'error', key='pgfmt-%d' % i,
                    expected='第 %d 节页码格式应为 %s（来自模板）' % (i, scheme),
                    actual='第 %d 节页码格式为 %s' % (i, fmt),
                    suggestion='把第 %d 节的 pgNumType/@fmt 改成 %s。' % (i, scheme),
                    source_key='page_numbering_scheme', kind='footer',
                    para_index=ctx.sec_first(i), action=None)

    # 3) 无标准时的自洽检查（多节文档）
    if scheme is None and n >= 2:
        any_pgnt = any(o[1] for o in secs)
        for (i, has_pgnt, fmt, start, has_page_field) in secs:
            # 前置部分用了阿拉伯数字 → 提示"通常用罗马"
            if is_front.get(i) and fmt == 'arabic':
                ctx.add('page-numbering', 'warn', key='front-arabic-%d' % i,
                        expected='前置部分通常用罗马数字页码（如 i, ii, iii）',
                        actual='第 %d 节（前置部分）使用了阿拉伯数字页码' % i,
                        suggestion='若学校要求前置部分用罗马数字，请改第 %d 节的 '
                                   'pgNumType 为 lowerRoman。' % i,
                        source_key='page_numbering_scheme', kind='footer',
                        para_index=ctx.sec_first(i), action=None)
                break
            # 别节都设了 pgNumType，本节却没设 → 格式不自洽
            if not has_pgnt and any_pgnt:
                ctx.add('page-numbering', 'warn', key='pgnt-missing-%d' % i,
                        expected='多节文档各节页码格式应自洽；其他节已显式设置页码格式',
                        actual='第 %d 节未设置 pgNumType，页码起始/格式不明确'
                               '（可能落到默认连续编号）' % i,
                        suggestion='明确第 %d 节的页码格式（前置用 lowerRoman、正文用 '
                                   'decimal），节与节之间用"下一页"分节符断开。' % i,
                        source_key='page_numbering_scheme', kind='footer',
                        para_index=ctx.sec_first(i), action=None)


def rule_header(ctx):
    """逐节页眉文字与横线。header_rule 为 None 时不判"该不该有横线"，只报不一致。"""
    doc = ctx.scan.doc
    secs = ctx.scan.sections()           # SectionInfo: header_text / header_has_rule
    rule_std = ctx.std['header_rule']
    content_std = ctx.std['header_content']

    # 1) 有标准时：逐节判定横线是否存在
    if rule_std is not None:
        want_rule = rule_std not in (None, False, 'none', 'no', '无')
        for s in secs:
            if not s.header_text:
                continue
            if want_rule and not s.header_has_rule:
                loc = _header_loc(ctx, s.index)
                ctx.add('header', 'warn', key='header-rule-missing-%d' % s.index,
                        expected='页眉应有横线（header_rule=%s）' % rule_std,
                        actual='第 %d 节页眉无横线' % s.index,
                        suggestion='在页眉段落加 bottom 边框（single）。本工具可自动补：见下方动作。',
                        source_key='header_rule', kind='header',
                        para_index=ctx.sec_first(s.index),
                        action=ctx.act('para.border', locator=loc, edge='bottom',
                                       val='single', sz=6) if loc else None)
            elif (not want_rule) and s.header_has_rule:
                ctx.add('header', 'warn', key='header-rule-extra-%d' % s.index,
                        expected='页眉不应有横线',
                        actual='第 %d 节页眉有横线' % s.index,
                        suggestion='若学校要求无横线，请删掉页眉段落的 bottom 边框。',
                        source_key='header_rule', kind='header',
                        para_index=ctx.sec_first(s.index), action=None)

    # 2) 各节横线是否一致（无论有无标准都查；None 视为"无横线"）
    rule_flags = [bool(s.header_has_rule) for s in secs]
    if True in rule_flags and False in rule_flags:
        ctx.add('header', 'warn', key='header-rule-inconsistent',
                expected='各节页眉横线应一致（要么都有、要么都没有）',
                actual='各节页眉横线不一致：%s' % ', '.join(
                    '第%d节%s' % (s.index,
                                 '有横线' if s.header_has_rule else
                                 ('无横线' if s.header_has_rule is False else '无页眉'))
                    for s in secs),
                suggestion='前置部分（摘要/目录）常无页眉；若学校要求全文统一页眉样式，'
                           '请给缺失的节补/去横线。',
                source_key='header_rule', kind='header',
                para_index=ctx.sec_first(secs[0].index), action=None)

    # 3) 页眉文字：与标准比对 / 各节是否一致
    texts = [(s.index, s.header_text) for s in secs if s.header_text]
    if content_std is not None:
        for (i, t) in texts:
            if norm(t) != norm(content_std):
                ctx.add('header', 'warn', key='header-text-%d' % i,
                        expected='页眉文字应为「%s」（来自模板）' % content_std,
                        actual='第 %d 节页眉为「%s」' % (i, t),
                        suggestion='把页眉文字改成规定的题目。',
                        source_key='header_content', kind='header',
                        para_index=ctx.sec_first(i), action=None)
    if len({norm(t) for (_, t) in texts}) > 1:
        ctx.add('header', 'warn', key='header-text-inconsistent',
                expected='各节页眉文字应一致',
                actual='各节页眉文字不一致：%s' % ', '.join(
                    '第%d节「%s」' % (i, t) for (i, t) in texts),
                suggestion='页眉通常应全文统一为论文题目；请确认是否为有意为之。',
                source_key='header_content', kind='header',
                para_index=ctx.sec_first(texts[0][0]), action=None)


def rule_fonts(ctx):
    """正文字体/字号/行距。同一层级出现 2 种以上字号/字体 → warn。"""
    doc = ctx.scan.doc
    size_std = ctx.std['body_size_pt']
    font_std = ctx.std['body_font_cn']
    # 按样式分组：同一样式内应字号/字体统一
    groups = {}
    for i, p in enumerate(doc.paragraphs, 1):
        if p.style is None:
            continue
        nm = p.style.name
        if not is_body_style(nm):
            continue
        if not (p.text or '').strip():
            continue
        sz = para_eff_size(p, doc)
        fo = para_eff_font(p, doc)
        groups.setdefault(nm, []).append((i, sz, fo))

    for nm, items in groups.items():
        sizes = [s for (_, s, _) in items if s is not None]
        if sizes and len({round(x, 3) for x in sizes}) > 1:
            dist = {}
            for (i, s, _) in items:
                if s is not None:
                    dist.setdefault(round(s, 2), []).append(i)
            ctx.add('fonts', 'warn', key='fonts-size-%s' % norm(nm),
                    expected='「%s」样式的段落应字号统一' % nm,
                    actual='「%s」样式出现 %d 种字号：%s' % (
                        nm, len(dist),
                        ', '.join('%gpt→段%s' % (k, ','.join(map(str, v[:5])))
                                 for k, v in sorted(dist.items())),),
                    suggestion='把不同字号的段落改回用「%s」样式（清除直接字号设置，'
                               '改样式才跟着变）。' % nm,
                    source_key='body_size_pt', kind='paragraph',
                    para_index=items[0][0], action=None)
        fonts = [f for (_, _, f) in items if f and f[1]]
        if fonts:
            eastasia = {f[1] for f in fonts}
            if len(eastasia) > 1:
                ctx.add('fonts', 'warn', key='fonts-font-%s' % norm(nm),
                        expected='「%s」样式的段落应中文字体统一' % nm,
                        actual='「%s」样式出现 %d 种中文字体：%s' % (
                            nm, len(eastasia), ', '.join(sorted(eastasia))),
                        suggestion='把不同字体的段落改回用「%s」样式。' % nm,
                        source_key='body_font_cn', kind='paragraph',
                        para_index=items[0][0], action=None)
        # 有标准时逐段比对字号
        if size_std is not None:
            for (i, s, _) in items:
                if s is not None and abs(s - size_std) > 0.5:
                    ctx.add('fonts', 'warn', key='fonts-std-%s-%d' % (norm(nm), i),
                            expected='「%s」段落字号应为 %gpt（来自模板）' % (nm, size_std),
                            actual='第 %d 段（%s 样式）字号为 %gpt' % (i, nm, s),
                            suggestion='把该段改回 %gpt（用命名样式，不要直接设字号）。' % size_std,
                            source_key='body_size_pt', kind='paragraph',
                            para_index=i, action=None)
    # 行距：docx 里行距解析成本高且易误报，标准缺失时整体不判，仅在有标准时提示
    if ctx.std['body_line_spacing'] is not None:
        ctx.notes.append('行距检查：需解析每段 pPr/spacing，docx 层面可粗略比对；'
                         '本次未逐段展开，若需请单独开启。')


def rule_style_compliance(ctx):
    """命名样式合规 + 直接格式硬改。"""
    doc = ctx.scan.doc
    # 1) 样式名是否在 spec 声明样式集（仅 spec 提供时）
    declared = _declared_styles(ctx.spec)
    if declared:
        used = set(ctx.scan.style_names())
        for nm in used:
            if nm and nm not in declared:
                ctx.add('style-compliance', 'warn', key='style-unknown-%s' % norm(nm),
                        expected='段落样式应在模板声明的样式集内',
                        actual='使用了模板未声明的样式「%s」' % nm,
                        suggestion='确认「%s」是否为预期样式；否则改回命名样式。' % nm,
                        source_key='body_size_pt', kind='paragraph', action=None)

    # 2) 正文段误用标题样式（长正文套了章节标题样式）
    for i, p in enumerate(doc.paragraphs, 1):
        if p.style is None:
            continue
        nm = p.style.name
        if is_chapter_heading(nm):
            txt = (p.text or '').strip()
            if len(txt) > 40:
                ctx.add('style-compliance', 'warn', key='heading-misuse-%d' % i,
                        expected='标题样式只用于章节标题（短标题）',
                        actual='第 %d 段用了「%s」样式但正文长达 %d 字，'
                               '疑似正文误用标题样式' % (i, nm, len(txt)),
                        suggestion='若该段是正文，请改回正文样式（如「正文文本」）；'
                                   '若是长标题请确认。',
                        source_key='body_size_pt', kind='paragraph',
                        para_index=i, action=None)

    # 3) 直接格式硬改：body 段落的 run 带了直接 w:sz / w:rFonts
    for i, p in enumerate(doc.paragraphs, 1):
        if p.style is None:
            continue
        nm = p.style.name
        if not is_body_style(nm):
            continue
        p_el = p._element
        for r in p.runs:
            rpr = r._r.find(W + 'rPr')
            if rpr is None:
                continue
            has_sz = rpr.find(W + 'sz') is not None
            has_font = rpr.find(W + 'rFonts') is not None
            if has_sz or has_font:
                what = '字号' if has_sz else '字体'
                action = None
                if has_sz:
                    s, e = _run_offset(p_el, r._r)
                    t = ''.join(x.text or '' for x in r._r.findall(W + 't'))
                    action = ctx.act('run.clear_size',
                                     locator=make_locator(p_el, i, s, e,
                                                         t if t else None))
                ctx.add('style-compliance', 'warn', key='direct-fmt-%d' % i,
                        expected='该段应走命名样式「%s」，排版由样式统一控制' % nm,
                        actual='第 %d 段内有直接%s设置（w:%s），这段的排版脱离样式，'
                               '改样式它不会跟着变' % (
                                   i, what, 'sz' if has_sz else 'rFonts'),
                        suggestion=('删掉这段的直接字号，改回随「%s」样式继承；'
                                    '这样改样式时它才跟着变。' % nm)
                        if has_sz else
                        ('这段直接指定了字体，脱离样式；请清除直接字体设置、'
                         '改用命名样式「%s」。' % nm),
                        source_key='body_size_pt', kind='run',
                        para_index=i, text=(p.text or '')[:46],
                        start=_run_offset(p_el, r._r)[0], end=_run_offset(p_el, r._r)[1],
                        action=action)
                break                              # 每段报一处即可


def rule_toc(ctx):
    """目录：TOC 域是否存在/失效、条目数 vs 正文标题数。"""
    doc = ctx.scan.doc
    h, region = _toc_region(doc)
    if h is None:
        ctx.add('toc', 'info', key='no-toc',
                expected='论文应有目录（TOC 域）',
                actual='未找到「目录/目次」标题，可能没有插入目录',
                suggestion='在正文前插入目录（引用→目录），并设为显示到三级标题。',
                source_key='body_size_pt', kind='document', action=None)
        return

    # TOC 域代码（instrText 含 TOC）
    toc_instr = None
    for pi in region:
        for it in doc.paragraphs[pi - 1]._element.iter(W + 'instrText'):
            if it.text and 'TOC' in (it.text or '').upper():
                toc_instr = (it.text or '').strip()
    if toc_instr is None:
        ctx.add('toc', 'error', key='toc-field-missing',
                expected='目录区应有 TOC 域（w:instrText 含 TOC 指令），用于自动生成目录',
                actual='目录区（从第 %d 段起）未找到 TOC 域代码，目录不会生成或已失效' % h,
                suggestion='删除该目录，重新插入目录域（引用→目录→自动目录）；插入后按 F9 更新。',
                source_key='body_size_pt', kind='toc', para_index=h,
                action=None, suppressible=False)
        return

    # 域已失效（未定义书签）
    for pi in region:
        t = doc.paragraphs[pi - 1].text or ''
        if '错误！未定义书签' in t or 'Error! Bookmark not defined' in t:
            ctx.add('toc', 'error', key='toc-broken',
                    expected='目录域应正常解析，不出现"错误！未定义书签"',
                    actual='目录区第 %d 段出现「错误！未定义书签 / Error! Bookmark not defined」' % pi,
                    suggestion='目录域已失效：选中目录按 F9 更新域；若仍报错，删除目录重新插入。',
                    source_key='body_size_pt', kind='toc', para_index=pi,
                    action=None, suppressible=False)
            break

    # 条目数 vs 正文标题数
    entries = sum(1 for pi in region
                  if '_Toc' in (doc.paragraphs[pi - 1]._element.xml or ''))
    heads = sum(1 for p in doc.paragraphs
                if p.style and is_chapter_heading(p.style.name) and (p.text or '').strip())
    if entries == 0 and heads > 0:
        ctx.add('toc', 'error', key='toc-empty',
                expected='目录应有条目（与正文标题对应）',
                actual='目录区无条目（可能未更新域）',
                suggestion='选中目录按 F9 更新域。',
                source_key='body_size_pt', kind='toc', para_index=h, action=None)
    elif entries < heads:
        ctx.add('toc', 'warn', key='toc-mismatch',
                expected='目录条目数应与正文标题数一致',
                actual='目录条目 %d 条 < 正文标题 %d 个，可能有章节未进目录' % (entries, heads),
                suggestion='选中目录按 F9 更新域；确认缺失的章节是否应出现在目录中。',
                source_key='body_size_pt', kind='toc', para_index=h, action=None)
    # 注：目录"是否已更新"需渲染确认，见 meta.notes


def _declared_styles(spec):
    """spec 声明的样式集（仅当 spec 提供时）。格式不定则宽松返回 None。"""
    if not spec:
        return None
    styles = spec.get('styles')              # 如模板提取产出的 styles 块
    if isinstance(styles, dict):
        return set(styles.keys())
    return None


def _header_loc(ctx, index):
    """造页眉段落的定位（用于自动补横线动作）。"""
    try:
        sec = ctx.scan.doc.sections[index - 1]
        for hp in sec.header.paragraphs:
            if (hp.text or '').strip():
                return make_locator(hp._element, ctx.para_index_of(hp._element))
    except Exception:
        pass
    return None


RULES = [rule_page_setup, rule_page_numbering, rule_header,
         rule_fonts, rule_style_compliance, rule_toc]


# ================================================================ 主流程

def audit(docx_path, spec_path=None, ignore_path=None, verbose=True):
    ctx = Ctx(docx_path, spec_path)
    for fn in RULES:
        try:
            fn(ctx)
        except Exception as e:
            # 单条规则炸了不拖垮整份审计 —— 但**绝不能只写进 notes 就算了**
            # （同 cite-doctor：早期把异常塞 notes，CLI --print 不显示，于是整条
            # 规则失效却对外表现"全部通过"）。升格成 error 级 Issue。
            import traceback
            ctx.notes.append('规则 `%s` 执行异常：%s: %s' % (fn.__name__, type(e).__name__, e))
            ctx.issues.append({
                'id': issue_id('audit-internal', fn.__name__),
                'rule': 'audit-internal-error', 'layer': LAYER, 'skill': SKILL,
                'severity': 'error',
                'location': {'kind': 'document', 'index': None, 'excerpt': None,
                             'offset': None},
                'expected': '每条规则都应正常跑完',
                'actual': '规则 `%s` 执行异常：%s: %s' % (fn.__name__, type(e).__name__, e),
                'standard_source': 'default',
                'suggestion': '这是工具缺陷，不是文档问题。本次结果**不可信**，'
                              '请把这条报给维护者；排查线索：%s'
                              % traceback.format_exc().splitlines()[-1],
                'suppressible': False, 'action': None,
            })

    ign_ids, ign_rules = R.load_ignore(
        ignore_path,
        os.path.join(os.path.dirname(ctx.docx), '.thesisignore'))
    kept, sup = R.apply_ignore(ctx.issues, ign_ids, ign_rules)
    kept.sort(key=lambda x: (R.SEV_ORDER.get(x['severity'], 9),
                             x['location'].get('index') or 0))

    summary = ctx.scan.summary()
    # 把"判不出来"的项如实写进 notes（绝不瞎报）
    ctx.notes.append('页眉文字是否与封面/论文题目一致、目录是否已更新'
                     '（条目与正文标题是否一一对应）：需渲染确认，docx 层面判不出来'
                     '—— 目录域存在与否、是否失效可判，但"是否已更新"不可判。')
    ctx.notes.append('正文/前置节划分依据：含摘要/目录的节判为前置，最后一节判为正文；'
                     '复杂多节文档请人工确认。')
    meta = {
        'skill': SKILL, 'layer': LAYER,
        'docx_sha256': _sha256(ctx.docx),
        'generated_at': time.strftime('%Y-%m-%dT%H:%M:%S'),
        'spec_path': ctx.spec_path,
        'standard': {
            'page_size': ctx.std['page_size'], 'source': ctx.src['page_size'],
            'page_margins_cm': ctx.std['page_margins_cm'],
            'margins_source': ctx.src['page_margins_cm'],
            'page_numbering_scheme': ctx.std['page_numbering_scheme'],
            'scheme_source': ctx.src['page_numbering_scheme'],
            'header_content': ctx.std['header_content'],
            'header_content_source': ctx.src['header_content'],
            'header_rule': ctx.std['header_rule'],
            'header_rule_source': ctx.src['header_rule'],
            'body_size_pt': ctx.std['body_size_pt'],
            'body_size_source': ctx.src['body_size_pt'],
            'body_font_cn': ctx.std['body_font_cn'],
            'body_font_source': ctx.src['body_font_cn'],
            'body_line_spacing': ctx.std['body_line_spacing'],
            'line_source': ctx.src['body_line_spacing'],
        },
        'standard_source_note': _std_note(ctx),
        'notes': ctx.notes,
    }
    if verbose:
        print('[format-audit] %s' % ctx.docx)
        print('  段落 %d，分节 %d，正文引注 %d，参考文献 %d'
              % (summary['paragraphs'], summary['sections'],
                 summary['citations'], summary['reference_entries']))
        print('  标准来源：%s' % _std_note(ctx))
        print('  问题 %d 条（按裁定保留 %d 条已单列）' % (len(kept), len(sup)))
    return {'issues': kept, 'suppressed': sup, 'summary': summary, 'meta': meta}


def _std_note(ctx):
    return describe(ctx.std, ctx.src,
                    ['page_size', 'page_margins_cm', 'page_numbering_scheme',
                     'header_content', 'header_rule', 'body_size_pt',
                     'body_font_cn', 'body_line_spacing'])


def _sha256(path):
    from shared.lib.ooxml_guard import sha256_file
    return sha256_file(path)


def main():
    ap = argparse.ArgumentParser(description='L1 格式与版式审计（只读）')
    ap.add_argument('-d', '--docx', required=True)
    ap.add_argument('-s', '--spec', default=None, help='不传则自动在文档旁找 spec.json')
    ap.add_argument('-i', '--ignore', default=None, help='.thesisignore 路径')
    ap.add_argument('-o', '--out', default=None, help='输出目录')
    ap.add_argument('--print', action='store_true', help='把报告打到终端')
    ap.add_argument('--json', action='store_true', help='打印 JSON')
    a = ap.parse_args()

    spec = a.spec or find_spec(a.docx)
    res = audit(a.docx, spec, a.ignore)
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.docx)),
                                '.thesis-doctor', 'audit', SKILL)
    os.makedirs(out, exist_ok=True)
    payload = R.render_json(res['issues'], res['suppressed'], res['summary'], res['meta'])
    with open(os.path.join(out, 'audit.json'), 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    md = R.render_markdown(res['issues'], res['suppressed'], res['summary'], res['meta'])
    with open(os.path.join(out, 'audit-report.md'), 'w', encoding='utf-8') as f:
        f.write(md)
    print('输出：%s' % out)
    if a.print:
        print('\n' + md)
    if a.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    sys.exit(main())
