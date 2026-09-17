#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
skills/cite-doctor/scripts/audit.py —— L2 引注与文献审计（**只读，绝不写盘**）

七条规则（与 assets/skills-index.yaml 的 covers 一一对应）
    cite-position   引注该在句读符之前（……问题[1]。），而不是之后（……问题。[1]）
    cite-form       形态：上标 + 方括号；不该显式写 w:sz（会被渲染器再缩一次）
    cite-jump       Ctrl+单击要能跳。两个方向都查：正文锚点缺失/悬空、
                    文献表条目缺 refN 书签
    cite-numbering  条目编号须连续且从 1 起；顺序编码制还要求与正文首现顺序一致
    ref-coverage    正反覆盖：表里有没被引的；正文引了表里没有的
    ref-format      GB/T 7714 逐条可判定项（文献类型标识、年份、[J] 的期号页码）
    ref-duplicate   重复条目

设计上刻意"保守"的三处（不是没做，是故意不做）
    ① ref-duplicate 不给自动删除动作。删一条要把它后面所有引注重编号 +
       重挂链接，属全局手术；自动做等于替你改写了一整套引用体系。
    ② 顺序编码制的"重编号"不给自动动作，同上。
    ③ ref-format 只报不修。著录改写要看原文献，机器擅自补页码会造假。

标准来源三级优先（每条 Issue 都带 standard_source，报告里会显著标出）
    template（模板/范文 spec） > config（用户配置） > default（通用学术惯例）
    没给 spec 时全部落 default，并在报告顶部明写"未提供模板标准"—— 不能让人
    误以为"上标方括号"是学校的规定。

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
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
if PKG not in sys.path:
    sys.path.insert(0, PKG)

from shared.lib.docx_ops import make_locator                     # noqa: E402
from shared.lib.docx_scan import Scanner, excerpt, issue_id, norm  # noqa: E402
from shared.lib.standards import (                                # noqa: E402
    describe, find_spec, load_spec, resolve_standard,
)
from shared.lib import report as R                                # noqa: E402

SKILL = 'cite-doctor'
LAYER = 'L2'

TYPE_MARKER_RE = re.compile(r'\[(M|J|D|C|N|R|S|P|A|Z|G|EB/OL|DB/OL|J/OL|M/CD|C/OL|N/OL)\]')
# 出版年：**不要**要求前后有括号或逗号。
# GB/T 7714 的实际写法是「北京: 高等教育出版社, 2011.」—— 年后跟的是句点。
# 第一版写成 `[（(\[,，]\s*(19|20)\d{2}\s*[）)\],，]`，结果 46 条里 33 条被误报"缺年份"，
# 全是噪声。改用"独立四位年份"判定：前后不接数字即可。
YEAR_RE = re.compile(r'(?<!\d)(?:19|20)\d{2}(?!\d)')
# [J] 的页码/文章号判据：**年份之后**要出现「冒号 + 数字」。
#
# 为什么不用"年,卷(期): 起止页"这种严格式：中文期刊与英文期刊写法差别很大，
# 严格式把下面这些合法写法全判成缺页码（实测误报 5 条第 4 条都是这两种）：
#     …, 2026, 7: 100196.          ← 年,卷: 文章号（无期号）
#     …, 2026, 16(5): 767.         ← 单页 / 文章号而非起止页
#     …, 2026: 1-29.               ← 网络首发，无卷期只有页
#     …, 2026: 122049.             ← 只有文章号
# 只查"年后有没有 :数字"，上述全部放行，而真正缺页码的
#     …, 2025(11).                 ← 有年有期、后面直接句点
# 会被准确抓住。宁可判不出，不可误报 —— 报告噪声一起来就没人看了。
PAGES_AFTER_YEAR_RE = re.compile(r'[:：]\s*\d')


# ================================================================ 上下文

class Ctx:
    """一次审计的全部共享状态。规则函数只读它，不互相传参。"""

    def __init__(self, docx, spec_path=None):
        self.docx = os.path.abspath(docx)
        self.scan = Scanner(self.docx)
        self.spec, self.spec_path = load_spec(spec_path)
        self.std, self.src, self.notes = resolve_standard(self.spec)
        self.cites = self.scan.citations()
        self.ref_head, self.refs = self.scan.reference_section()
        self.bookmarks = set(self.scan.bookmarks())
        self.issues = []
        # 文档里现存的链接形态（决定补链接时用哪种序列化）
        kinds = Counter(c.link_kind for c in self.cites if c.link_kind)
        self.dominant_link_kind = kinds.most_common(1)[0][0] if kinds else 'hyperlink'
        # 段落内引注序号（构造稳定 ID 用）
        self._para_counter = Counter()

    # -------------------------------------------------- 造 Issue

    def cite_key(self, c):
        """给一处引注生成内容稳定的唯一键（不依赖段号，段号会漂）。"""
        k = norm(c.para_text)
        n = self._para_counter[k]
        self._para_counter[k] += 1
        return '%s|%s|%d' % (k, c.text, n)

    def locator_for_cite(self, c):
        return make_locator(self.scan.para_el(c.para_index), c.para_index,
                            c.start, c.end, c.text)

    def locator_for_ref(self, e):
        return make_locator(self.scan.para_el(e.para_index), e.para_index)

    def add(self, rule, severity, *, key, expected, actual, suggestion,
            source=None, suppressible=True, action=None,
            para_index=None, text=None, start=None, end=None, kind='paragraph'):
        if text is not None and start is not None:
            ex = excerpt(text, start, end)
        elif text:
            ex = text[:46] + ('…' if len(text) > 46 else '')
        else:
            ex = None
        loc = {'kind': kind, 'index': para_index, 'excerpt': ex,
               'offset': start if kind == 'run' else None}
        self.issues.append({
            'id': issue_id(rule, key),
            'rule': rule, 'layer': LAYER, 'skill': SKILL,
            'severity': severity, 'location': loc,
            'expected': expected, 'actual': actual,
            'standard_source': source or self.src.get(
                {'cite-position': 'cite_position', 'cite-form': 'cite_style',
                 'cite-jump': 'cite_style', 'cite-numbering': 'bib_system',
                 'ref-coverage': 'bib_system', 'ref-format': 'bib_standard',
                 'ref-duplicate': 'bib_standard'}.get(rule, 'bib_standard'), 'default'),
            'suggestion': suggestion, 'suppressible': suppressible,
            'action': action,
        })

    def act(self, action_name, **params):
        """造一个动作。

        ★ 第一个参数**不能**叫 `name` ★
        动作里恰好也有个 `name` 键（如 ref.bookmark 的 name='ref7'），
        参数撞名会让 Python 抛 `got multiple values for argument 'name'`。
        这个异常当时被"单规则 try/except"吞掉，于是 cite-jump 整条规则
        **一条问题都不报**，看起来像"文档没问题"。测试是注入缺陷才揪出来的。
        """
        return {'name': action_name, 'params': params}


# ================================================================ 规则

def rule_cite_position(ctx):
    """引注必须落在句读符之前。after-punct 是要改的；mid-sentence 是合法的。"""
    want = ctx.std['cite_position']
    if want != 'before-punct':
        return                                    # 模板要求别的（如页下注）→ 本规则不适用
    punct_re = re.compile(r'[。！？；，、：]+')
    for c in ctx.cites:
        if c.position == 'after-punct':
            m = punct_re.search(c.prev_char or '')
            punct = m.group(0) if m else (c.prev_char or '。')
            ctx.add('cite-position', 'error', key=ctx.cite_key(c),
                    expected='引注落在句读符**之前**：……问题[1]。'
                             '（紧跟标点之后会让引注归属含混——读的人分不清它注的是上一句还是下一句）',
                    actual='引注「%s」紧跟「%s」之后：……%s%s' % (c.text, punct, punct, c.text),
                    suggestion='把引注挪到「%s」前面。' % punct,
                    para_index=c.para_index, text=c.para_text,
                    start=c.start, end=c.end, kind='run',
                    action=ctx.act('cite.move_before_punct',
                                   locator=ctx.locator_for_cite(c), punct=punct))
        elif c.position == 'end-of-paragraph':
            ctx.add('cite-position', 'warn', key=ctx.cite_key(c),
                    expected='段末引注仍应有句末标点收尾：……问题[1]。',
                    actual='引注位于段末且前面无句末标点：%s' % c.excerpt,
                    suggestion='人工确认句末标点是否漏写（工具不擅自补标点）。',
                    para_index=c.para_index, text=c.para_text,
                    start=c.start, end=c.end, kind='run', action=None)


def rule_cite_form(ctx):
    """形态：上标 / 方括号 / 不得显式设字号。"""
    style = ctx.std['cite_style']
    if style not in ('superscript-bracket', 'superscript', 'bracket'):
        ctx.add('cite-form', 'info', key='style-unsupported',
                expected='—',
                actual='spec 声明的引注形态为「%s」，本工具只自动核对上标/方括号类形态。'
                       % style,
                suggestion='著者-出版年制请人工核对，或改用顺序编码制。',
                source='template', suppressible=True, action=None)
        return
    want_sup = style in ('superscript-bracket', 'superscript')
    want_bracket = 'square' if style in ('superscript-bracket', 'bracket') else None

    for c in ctx.cites:
        loc = ctx.locator_for_cite(c)
        if want_sup and not c.superscript:
            ctx.add('cite-form', 'error', key=ctx.cite_key(c),
                    expected='引注为上标（`w:vertAlign="superscript"`）',
                    actual='引注「%s」不是上标：%s' % (c.text, c.excerpt),
                    suggestion='设为上标。字号不要另外指定，渲染器会自动缩到约 65%。',
                    para_index=c.para_index, text=c.para_text,
                    start=c.start, end=c.end, kind='run',
                    action=ctx.act('cite.superscript', locator=loc, value=True))
        if want_sup is False and c.superscript:
            ctx.add('cite-form', 'error', key=ctx.cite_key(c),
                    expected='引注为行内（非上标）',
                    actual='引注「%s」被设成了上标' % c.text,
                    suggestion='取消上标。',
                    para_index=c.para_index, text=c.para_text,
                    start=c.start, end=c.end, kind='run',
                    action=ctx.act('cite.superscript', locator=loc, value=False))
        if want_bracket == 'square' and c.bracket != 'square':
            ctx.add('cite-form', 'error', key=ctx.cite_key(c),
                    expected='引注用方括号包裹：[1]',
                    actual='引注用的是「%s」括号：%s'
                           % ({'paren': '圆括号', 'none': '无括号'}.get(c.bracket, c.bracket),
                              c.text),
                    suggestion='改成方括号形态 [%s]。'
                               % ','.join(str(x) for x in c.refs),
                    para_index=c.para_index, text=c.para_text,
                    start=c.start, end=c.end, kind='run',
                    action=ctx.act('cite.rewrite', locator=loc,
                                   text='[%s]' % ','.join(str(x) for x in c.refs)))

    # 显式字号：上标自己会缩，再写死一个字号会被"缩两次"
    if not ctx.std['cite_allow_explicit_size']:
        for c in ctx.cites:
            for r in ctx.scan.runs(c.para_index):
                if r.start == c.start and r.end == c.end:
                    if r.size_pt is not None:
                        ctx.add('cite-form', 'warn', key=ctx.cite_key(c) + '|size',
                                expected='上标引注不显式指定字号（由渲染器自动缩放）',
                                actual='引注「%s」被显式设为 %s pt —— 渲染时会在这个基础上再缩约 65%%'
                                       % (c.text, r.size_pt),
                                suggestion='删掉这一处的显式字号，改回随样式继承。',
                                para_index=c.para_index, text=c.para_text,
                                start=c.start, end=c.end, kind='run',
                                action=ctx.act('run.clear_size',
                                               locator=ctx.locator_for_cite(c)))
                    break


def rule_cite_jump(ctx):
    """两个方向的跳转完整性。"""
    ref_by_num = {e.number: e for e in ctx.refs if e.number is not None}
    entry_bm = {}                      # 编号 → 该条目段上挂的 refN 书签
    for e in ctx.refs:
        if e.number is None:
            continue
        want = 'ref%d' % e.number
        if want in [b for b in e.bookmarks]:
            entry_bm[e.number] = want

    # ---- 方向 A：文献表条目缺 refN 书签
    for e in ctx.refs:
        if e.number is None:
            continue
        if e.number not in entry_bm:
            ctx.add('cite-jump', 'error', key='refbm|%s' % norm(e.text),
                    expected='条目 [%d] 上应挂书签 `ref%d`，供正文引注跳转' % (e.number, e.number),
                    actual='该条目没有 `ref%d` 书签' % e.number,
                    suggestion='补上书签。正文里所有指向 [%d] 的引注都要靠它跳。'
                               % e.number,
                    suppressible=False,
                    para_index=e.para_index, text=e.text,
                    action=ctx.act('ref.bookmark', locator=ctx.locator_for_ref(e),
                                   name='ref%d' % e.number))

    # ---- 方向 B：正文引注的锚点
    for c in ctx.cites:
        targets = ['ref%d' % n for n in c.refs]
        if len(c.refs) > 1:
            ctx.add('cite-jump', 'warn', key=ctx.cite_key(c) + '|multi',
                    expected='一处引注对应一个可跳转目标（合并引用需逐个挂链接）',
                    actual='这条引注含多个编号 %s，本工具不给自动修复' % c.refs,
                    suggestion='人工拆成多个上标，或确认是否需要分别跳转。',
                    para_index=c.para_index, text=c.para_text,
                    start=c.start, end=c.end, kind='run', action=None)
            continue
        tgt = targets[0]
        if c.anchor == tgt and tgt in ctx.bookmarks:
            continue                                   # 正常
        if c.anchor == tgt:
            # 链接文字对得上，但目标书签不存在 —— 这种"看着有链接、点了没反应"
            # 最容易被漏掉，必须显式报出来（旧版三个分支全不覆盖，静默通过）
            ctx.add('cite-jump', 'error', key=ctx.cite_key(c) + '|dangling',
                    expected='引注目标 `%s` 应存在于参考文献表' % tgt,
                    actual='引注「%s」指向 `%s`，但该锚点不存在（Ctrl+单击会无反应）'
                           % (c.text, tgt),
                    suggestion='补上 `%s` 书签，或确认该编号是否已失效。' % tgt,
                    suppressible=False,
                    para_index=c.para_index, text=c.para_text,
                    start=c.start, end=c.end, kind='run', action=None)
        elif c.anchor and c.anchor != tgt:
            ctx.add('cite-jump', 'error', key=ctx.cite_key(c) + '|wrong',
                    expected='引注「%s」应指向 `%s`' % (c.text, tgt),
                    actual='实际指向 `%s`（编号对不上）' % c.anchor,
                    suggestion='把链接目标改回 `%s`。' % tgt,
                    suppressible=False,
                    para_index=c.para_index, text=c.para_text,
                    start=c.start, end=c.end, kind='run',
                    action=ctx.act('cite.link', locator=ctx.locator_for_cite(c),
                                   anchor=tgt, form=ctx.dominant_link_kind))
        elif not c.anchor:
            if tgt in ctx.bookmarks:
                ctx.add('cite-jump', 'error', key=ctx.cite_key(c) + '|nolink',
                        expected='引注要实现 Ctrl+单击跳转到参考文献对应条目',
                        actual='引注「%s」没有链接（纯文字）' % c.text,
                        suggestion='补上指向 `%s` 的内部链接。' % tgt,
                        suppressible=False,
                        para_index=c.para_index, text=c.para_text,
                        start=c.start, end=c.end, kind='run',
                        action=ctx.act('cite.link', locator=ctx.locator_for_cite(c),
                                       anchor=tgt, form=ctx.dominant_link_kind))
            else:
                ctx.add('cite-jump', 'error', key=ctx.cite_key(c) + '|dangling',
                        expected='引注目标 `%s` 应存在于参考文献表' % tgt,
                        actual='引注「%s」指向的 `%s` 不存在' % (c.text, tgt),
                        suggestion='先确认第 %d 条是否存在；若不存在，先补条目再挂链接。'
                                   % c.refs[0],
                        suppressible=False,
                        para_index=c.para_index, text=c.para_text,
                        start=c.start, end=c.end, kind='run', action=None)


def rule_cite_numbering(ctx):
    """条目编号连续性 + 顺序编码制的"首现顺序=编号顺序"。"""
    nums = [e.number for e in ctx.refs]
    if not nums:
        return
    if any(n is None for n in nums):
        bad = [e.order for e in ctx.refs if e.number is None]
        ctx.add('cite-numbering', 'error', key='no-number',
                expected='每条参考文献开头都有 [n] 编号',
                actual='有 %d 条没有编号（表内第 %s 条）' % (len(bad), '、'.join(map(str, bad))),
                suggestion='补编号。编号是引注跳转与顺序编码制的基础。',
                suppressible=False, kind='document', action=None)
    valid = [n for n in nums if n is not None]
    if valid and valid != list(range(1, len(valid) + 1)):
        ctx.add('cite-numbering', 'error', key='discontinuous',
                expected='编号从 [1] 起连续递增',
                actual='编号实际为 %s' % ('、'.join(str(n) for n in valid[:12]) +
                                          ('…' if len(valid) > 12 else '')),
                suggestion='重排编号。注意：改编号要同步改正文所有引注与其链接，'
                           '工具不自动做，请人工逐条核对。',
                suppressible=False, kind='document', action=None)

    # 正文首现顺序
    first_seen, order = {}, []
    for c in ctx.cites:
        for n in c.refs:
            if n not in first_seen:
                first_seen[n] = c.para_index
                order.append(n)
    if not order:
        return
    if ctx.std['bib_system'] == 'sequence' and order != sorted(order):
        inv = [(i + 1, n) for i, n in enumerate(order) if i + 1 != n][:8]
        ctx.add('cite-numbering', 'warn', key='first-cite-order',
                expected='顺序编码制：编号顺序 = 正文首次出现顺序',
                actual='有 %d 处不一致，例如第 %s 个被引用的却是 [%s]'
                       % (len(inv), inv[0][0] if inv else '?', inv[0][1] if inv else '?'),
                suggestion='要么按首现顺序重编号，要么在 spec 里声明使用的不是顺序编码制。'
                           '（重编号属全局手术，请人工处理）',
                kind='document', action=None)
    unused = [n for n in sorted({e.number for e in ctx.refs if e.number}) if n not in first_seen]
    if unused:
        ctx.add('ref-coverage', 'warn', key='ref-unused',
                expected='参考文献表里的每条都应在正文被引用',
                actual='有 %d 条从未被正文引用：[%s]'
                       % (len(unused), '、'.join(str(n) for n in unused[:12])),
                suggestion='逐条确认：确实是多余文献就删（删完要重编号），'
                           '只是漏引就在该引的地方补上引注。',
                kind='document', action=None)
    missing = [n for n in first_seen if n not in {e.number for e in ctx.refs}]
    if missing:
        ctx.add('ref-coverage', 'error', key='cite-missing-entry',
                expected='正文引用的每个编号都能在参考文献表里找到',
                actual='正文引用了 [%s]，但表里没有' % '、'.join(str(n) for n in sorted(missing)),
                suggestion='补上缺失条目。这类是硬伤，答辩必被问。',
                suppressible=False, kind='document', action=None)


def rule_ref_format(ctx):
    """GB/T 7714 里可以纯文本判定的几项。只报不修。"""
    std_name = ctx.std['bib_standard']
    for e in ctx.refs:
        t = e.text
        tag = '[%s]' % (e.number if e.number else '?')
        if not TYPE_MARKER_RE.search(t):
            ctx.add('ref-format', 'warn', key='type|%s' % norm(t),
                    expected='%s 要求标注文献类型标识，如 [M] 专著 / [J] 期刊 / [D] 学位论文' % std_name,
                    actual='%s 条目未见类型标识' % tag,
                    suggestion='按实际文献类型补标识（不确定时查原文献，别猜）。',
                    para_index=e.para_index, text=t, action=None)
        if not YEAR_RE.search(t):
            ctx.add('ref-format', 'warn', key='year|%s' % norm(t),
                    expected='%s 要求著录出版年' % std_name,
                    actual='%s 条目未见出版年份' % tag,
                    suggestion='补出版年份。',
                    para_index=e.para_index, text=t, action=None)
        if '[J]' in t:
            m = YEAR_RE.search(t)
            if m and not PAGES_AFTER_YEAR_RE.search(t[m.end():]):
                ctx.add('ref-format', 'warn', key='jpages|%s' % norm(t),
                        expected='期刊论文 [J] 须著录页码或文章号',
                        actual='%s 条目在年份之后没有页码/文章号：%s'
                               % (tag, t[max(0, m.end() - 20):m.end() + 14]),
                        suggestion='补起止页码或文章号（形如 2024(5): 102-108 '
                                   '或 2026, 7: 100196）。',
                        para_index=e.para_index, text=t, action=None)


def rule_ref_duplicate(ctx):
    """重复条目：只报不修。删条目要连带重编号，机器不该替你做决定。"""
    seen = {}
    for e in ctx.refs:
        k = norm(re.sub(r'^\[[^\]]*\]', '', e.text))[:30]
        if not k:
            continue
        if k in seen:
            first = seen[k]
            ctx.add('ref-duplicate', 'error', key='dup|%s' % k,
                    expected='同一文献只应著录一次',
                    actual='[%s] 与 [%s] 开头 30 字完全相同：%s'
                           % (first.number or '?', e.number or '?', e.text[:50] + '…'),
                    suggestion='保留其中一条（通常是著录更全的那条），删掉另一条；'
                               '删完必须把它后面所有引注重新编号并重挂链接。'
                               '此步涉及全篇引用体系，工具不自动执行。',
                    suppressible=False,
                    para_index=e.para_index, text=e.text, action=None)
        else:
            seen[k] = e


RULES = [rule_cite_position, rule_cite_form, rule_cite_jump,
         rule_cite_numbering, rule_ref_format, rule_ref_duplicate]


# ================================================================ 主流程

def audit(docx_path, spec_path=None, ignore_path=None, verbose=True):
    ctx = Ctx(docx_path, spec_path)
    for fn in RULES:
        try:
            fn(ctx)
        except Exception as e:
            # 单条规则炸了不拖垮整份审计 —— 但**绝不能只写进 notes 就算了**。
            # 早期的做法就是把异常塞进 meta.notes，而 CLI --print 不显示 notes，
            # 于是 cite-jump 整条规则失效却对外表现为"全部通过"。
            # 现在把它变成一条 error 级 Issue：审计器自身坏了，本身就是最严重的发现。
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
    meta = {
        'skill': SKILL, 'layer': LAYER,
        'docx_sha256': _sha256(ctx.docx),
        'generated_at': time.strftime('%Y-%m-%dT%H:%M:%S'),
        'spec_path': ctx.spec_path,
        'standard': {
            'cite_style': ctx.std['cite_style'], 'source': ctx.src['cite_style'],
            'cite_position': ctx.std['cite_position'], 'position_source': ctx.src['cite_position'],
            'bib_system': ctx.std['bib_system'], 'system_source': ctx.src['bib_system'],
            'bib_standard': ctx.std['bib_standard'], 'standard_source': ctx.src['bib_standard'],
            'allow_explicit_superscript_size': ctx.std['cite_allow_explicit_size'],
        },
        'standard_source_note': _std_note(ctx),
        'notes': ctx.notes,
        'dominant_link_kind': ctx.dominant_link_kind,
    }
    if verbose:
        print('[cite-doctor] %s' % ctx.docx)
        print('  引注 %d 处，参考文献 %d 条，书签 %d 个'
              % (len(ctx.cites), len(ctx.refs), len(ctx.bookmarks)))
        print('  标准来源：%s' % _std_note(ctx))
        print('  问题 %d 条（按裁定保留 %d 条已单列）' % (len(kept), len(sup)))
    return {'issues': kept, 'suppressed': sup, 'summary': summary, 'meta': meta}


def _std_note(ctx):
    return describe(ctx.std, ctx.src,
                    ['cite_style', 'cite_position', 'bib_system', 'bib_standard'])


def _sha256(path):
    from shared.lib.ooxml_guard import sha256_file
    return sha256_file(path)


def main():
    ap = argparse.ArgumentParser(description='L2 引注与文献审计（只读）')
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
