#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
skills/structure-length/scripts/audit.py —— L4 结构与篇幅审计（**只读，绝不写盘**）

四条规则（与 assets/skills-index.yaml 的 covers 一一对应）
    heading-sequence   章节编号：同级连续 / 不跳级 / 编号层级与样式匹配
    figure-table      图表编号与题注：编号自洽、题注位置、正文引用存在
    abstract-layout   中英文摘要字数上限、关键词、摘要内不应有引注/图表引用
    length            双口径字数 + 分章配额（按章拆分，与 body_min_chars 对比）

设计上刻意"保守"的两处（不是没做，是故意不做）
    ① heading-sequence 的"三级标题数量上限"**不在这里报**——归 text-style 的
       heading-style 管，避免同一件事报两遍；只在 meta.notes 里写一句说明。
    ② 改章节编号 / 重排图表编号是全局手术，本 skill 不给自动动作（action 一律 None），
       只在 suggestion 里说明原因，交给人工。

标准来源三级优先（每条 Issue 都带 standard_source，报告里会显著标出）
    template（模板/范文 spec） > config（用户配置） > default（通用学术惯例）
    没给 spec 时全部落 default，并在报告顶部明写"未提供模板标准"。

关于"摘要字数上限"的诚实纪律（详见 shared/lib/standards.py 顶部注释）
    template-extract 只能"测量"范文实际是多少字（本样本中文摘要实测 506 字），
    它**不知道**学校上限是多少。所以上限**只认 spec 里的 abstract.cn.max_chars**；
    没有就走 default 1000 并标 standard_source='default'。
    **绝不可**把范文的实测字数当阈值，否则所有比范文长一点的摘要都被误报。

用法（与 cite-doctor 完全一致）
    python audit.py -d "F:/论文/初稿.docx"
    python audit.py -d 初稿.docx -s spec.json -o out/
    python audit.py -d 初稿.docx --print            # 把报告打到终端
"""
import argparse
import json
import os
import re
import sys
import time
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

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

SKILL = 'structure-length'
LAYER = 'L4'

# ---------------------------------------------------------------- 常量 / 正则

# 标题样式名 → 层级
HEADING_STYLE_LEVEL = {'一级标题': 1, '二级标题': 2, '三级标题': 3}

# 章节编号：第N章（阿拉伯或中文数字）
CHAPTER_RE = re.compile(r'^第\s*([0-9]+|[一二三四五六七八九十百零〇]+)\s*章')
# 点分编号：1.2 / 1.2.3（至少两段）
DOTTED_RE = re.compile(r'^(\d+(?:\.\d+){1,3})\b')
# 中文小节编号：（一）（二）
CN_SUB_RE = re.compile(r'^[（(]\s*([一二三四五六七八九十]+)\s*[）)]\s')

# 题注：图 1-1 / 表 1-1（允许全/半角短横与句点分隔）
CAP_RE = re.compile(r'^(图|表)\s*([0-9]+)[-—.]?([0-9]+)')
# 正文引用：如图 1-1 所示 / 见表 2-3
FIGREF_RE = re.compile(r'(图|表)\s*([0-9]+)[-—.]([0-9]+)')

# 全角字符（含全角空格 U+3000）：CJK + CJK 标点 + 全角 forms
FULLWIDTH_RE = re.compile(r'[\u2E80-\u9FFF\u3000-\u303F\uFF00-\uFFEF]')
# 汉字（用于"汉字数"口径）
HANZI_RE = re.compile(r'[\u4e00-\u9fff]')

# 中文数字 → 阿拉伯
_CN_DIGIT = {'零': 0, '〇': 0, '一': 1, '二': 2, '两': 2, '三': 3, '四': 4,
             '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}
_SECTION_WORDS = ['十', '百', '千']


def cn_int(s):
    """把「二十三」「一百零五」之类的中文数字转成 int。失败返回 None。"""
    s = (s or '').strip()
    if not s:
        return None
    if s.isdigit():
        return int(s)
    total, cur, prev_unit = 0, 0, 1
    try:
        for ch in s:
            if ch in _CN_DIGIT:
                cur = _CN_DIGIT[ch]
            elif ch in _SECTION_WORDS:
                unit = 10 ** _SECTION_WORDS.index(ch)
                if cur == 0:
                    cur = 1
                total += cur * unit
                cur, prev_unit = 0, unit
            else:
                return None
        return total + cur
    except Exception:
        return None


def _tbl_text(tbl_el):
    """从 w:tbl 元素取全部单元格文本（按文档顺序）。"""
    return '\n'.join(''.join(t.text or '' for t in tc.iter(W + 't'))
                     for tc in tbl_el.iter(W + 'tc'))


# ================================================================ 上下文

class Ctx:
    """一次审计的全部共享状态。规则函数只读它，不互相传参。"""

    def __init__(self, docx, spec_path=None):
        self.docx = os.path.abspath(docx)
        self.scan = Scanner(self.docx)
        self.spec, self.spec_path = load_spec(spec_path)
        self.std, self.src, self.notes = resolve_standard(self.spec)
        self.paras = self.scan.doc.paragraphs
        self._el_to_idx = {p._element: i for i, p in enumerate(self.paras, 1)}
        self.regions = self._compute_regions()
        self.blocks = self._walk_blocks()
        self.issues = []

    # -------------------------------------------------- 区域划分
    def _compute_regions(self):
        """返回 (first_h1, toc, ref, appendix) 四个 1-based 段号（找不到为 None）。"""
        first_h1 = toc = ref = app = None
        for i, p in enumerate(self.paras, 1):
            t = p.text.strip()
            nt = norm(t)
            if first_h1 is None and p.style.name == '一级标题':
                first_h1 = i
            if toc is None and nt == '目录':
                toc = i
            if ref is None and nt in {'参考文献', 'references', 'bibliography', '主要参考文献'}:
                ref = i
            if app is None and t.startswith('附录'):
                app = i
        return {'first_h1': first_h1, 'toc': toc, 'ref': ref, 'appendix': app}

    def _region_of_idx(self, idx):
        r = self.regions
        if r['first_h1'] and idx < r['first_h1']:
            return 'front'
        if r['ref'] and idx >= r['ref']:
            if r['appendix'] and idx >= r['appendix']:
                return 'appendix'
            return 'ref'
        return 'body'

    def _walk_blocks(self):
        """按文档顺序遍历 body 子元素，返回
        [(kind, text, para_index|None, region, order)]。
        kind ∈ {'p','tbl'}；para_index 仅 p 有；region 由位置推断；
        order 为单调递增序号，用于"题注相对表格的上下位置"判断。
        """
        out = []
        order = 0
        cur_region = 'front'
        for child in self.scan.doc.element.body:
            tag = child.tag
            if tag == W + 'p':
                idx = self._el_to_idx.get(child)
                txt = ''.join(n.text or '' for n in child.iter(W + 't'))
                if idx is not None:
                    region = self._region_of_idx(idx)
                else:
                    region = cur_region
                order += 1
                out.append(('p', txt, idx, region, order))
                cur_region = region
            elif tag == W + 'tbl':
                order += 1
                out.append(('tbl', _tbl_text(child), None, cur_region, order))
        return out

    # -------------------------------------------------- 造 Issue
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
        if source is None:
            source = self.src.get(_RULE_SOURCE.get(rule, 'default'), 'default')
        self.issues.append({
            'id': issue_id(rule, key),
            'rule': rule, 'layer': LAYER, 'skill': SKILL,
            'severity': severity, 'location': loc,
            'expected': expected, 'actual': actual,
            'standard_source': source,
            'suggestion': suggestion, 'suppressible': suppressible,
            'action': action,
        })

    def locator_for_idx(self, idx):
        return make_locator(self.scan.para_el(idx), idx)


_RULE_SOURCE = {
    'heading-sequence': 'default',
    'figure-table': 'default',
    'abstract-layout': 'default',
    'length': 'default',
}


# ================================================================ 规则 1：heading-sequence

def _parse_heading_number(text):
    """从标题文本提取编号。返回 (num_tuple, level) 或 (None, None)。

    num_tuple 形如 (1,) / (1,2) / (1,2,1)。level = len(num_tuple)。
    认三种形态：
      第1章 / 第一章         → (N,)  level=1
      1.2 / 1.2.3           → (1,2) / (1,2,1)  level=2/3
      （一）（二）           → (n,)  level=2（中文小节，通常作为二级）
    """
    t = text.strip()
    m = CHAPTER_RE.match(t)
    if m:
        n = cn_int(m.group(1))
        if n is not None:
            return (n,), 1
    m = DOTTED_RE.match(t)
    if m:
        try:
            parts = [int(x) for x in m.group(1).split('.')]
        except ValueError:
            return None, None
        if 2 <= len(parts) <= 3:
            return tuple(parts), len(parts)
    m = CN_SUB_RE.match(t)
    if m:
        n = cn_int(m.group(1))
        if n is not None:
            return (n,), 2
    return None, None


def rule_heading_sequence(ctx):
    """章节编号：同级连续 / 不跳级 / 编号层级与标题样式匹配。"""
    headings = []  # (idx, style_level, num_tuple, level, text)
    for b in ctx.blocks:
        if b[0] != 'p' or b[2] is None:
            continue
        idx = b[2]
        style_level = HEADING_STYLE_LEVEL.get(ctx.paras[idx - 1].style.name)
        if style_level is None:
            continue
        num_tuple, level = _parse_heading_number(b[1])
        if num_tuple is None:
            # 有标题样式却提取不出编号 —— 可能是"本章小结"之类没编号的二级标题，
            # 也可能是样式套错。这里保守地只报"样式为标题却无编号"的提示，不报错误。
            continue
        headings.append((idx, style_level, num_tuple, level, b[1]))

    if not headings:
        return

    by_key = {}  # num_tuple -> (idx, style_level)

    # ---- 1) 编号层级与标题样式/大纲级别是否匹配
    for idx, style_level, num_tuple, level, text in headings:
        if style_level != level:
            ctx.add('heading-sequence', 'error',
                    key='level-mismatch|%s|%d' % ('|'.join(map(str, num_tuple)), idx),
                    expected='标题样式层级（%d 级）应与编号段数一致：'
                             '「%s」提取出的编号有 %d 段，却用了 %d 级标题样式'
                             % (level, text[:24], level, style_level),
                    actual='样式=%d 级，编号=%s（%d 段）'
                           % (style_level, '.'.join(map(str, num_tuple)), level),
                    suggestion='改编号或改标题样式，使层级对齐（改章节编号属全局手术，'
                               '工具不自动改，请人工处理）。',
                    para_index=idx, text=text, action=None)

    # ---- 2) 跳级：每个编号的每个真前缀都必须作为某条标题存在
    seen = {h[2] for h in headings}
    for idx, style_level, num_tuple, level, text in headings:
        for k in range(1, level):
            prefix = num_tuple[:k]
            if prefix not in seen:
                ctx.add('heading-sequence', 'warn',
                        key='jump|%s|%d' % ('|'.join(map(str, num_tuple)), idx),
                        expected='出现编号「%s」之前，应先有上一级标题「%s」'
                                 % ('.'.join(map(str, num_tuple)),
                                    '.'.join(map(str, prefix))),
                        actual='标题「%s」的上级编号 %s 缺失（跳级）'
                               % (text[:24], '.'.join(map(str, prefix))),
                        suggestion='补上缺失的上级标题，或调整编号使其连续。'
                                   '（改编号属全局手术，请人工处理）',
                        para_index=idx, text=text, action=None)

    # ---- 3) 同级序号连续：按 (父前缀, 层级) 分组，段号须从 1 起连续
    groups = {}  # (prefix, level) -> list of (seg, idx, text)
    for idx, style_level, num_tuple, level, text in headings:
        prefix = num_tuple[:-1]
        seg = num_tuple[-1]
        groups.setdefault((prefix, level), []).append((seg, idx, text))

    for (prefix, level), items in groups.items():
        segs = sorted(s[0] for s in items)
        seg_set = set(segs)
        if segs and segs != list(range(segs[0], segs[-1] + 1)):
            missing = [n for n in range(segs[0], segs[-1] + 1) if n not in seg_set]
            # 重复段号
            dup = [n for n in seg_set if segs.count(n) > 1]
            scope = ('.'.join(map(str, prefix)) + ('.' if prefix else '')) or '全文'
            for n in missing:
                full = prefix + (n,)
                ctx.add('heading-sequence', 'error',
                        key='gap|%s' % ('|'.join(map(str, full))),
                        expected='同级编号应连续：在「%s」下，%s 之后应出现 %s'
                                 % (scope, '.'.join(map(str, prefix + (segs[0],)))
                                    if prefix else str(segs[0]), str(n)),
                        actual='编号在 %s 处断号，缺失 %s（现有：%s）'
                               % (scope, '、'.join(map(str, missing)),
                                  '、'.join(map(str, sorted(seg_set)))),
                        suggestion='补齐缺失的章节编号（改章节编号属全局手术，'
                                   '工具不自动改，请人工处理）。',
                        para_index=items[0][1], text=items[0][2], action=None)
            for n in dup:
                full = prefix + (n,)
                ctx.add('heading-sequence', 'error',
                        key='dup|%s' % ('|'.join(map(str, full))),
                        expected='每个章节编号应唯一',
                        actual='编号 %s 出现 %d 次（重复）'
                               % ('.'.join(map(str, full)), segs.count(n)),
                        suggestion='修正重复的章节编号（请人工处理）。',
                        para_index=items[0][1], text=items[0][2], action=None)


# ================================================================ 规则 2：figure-table

def rule_figure_table(ctx):
    """图表编号与题注：编号自洽、题注位置、正文引用存在。

    ★ 关键：表格在 w:body 里、不在 doc.paragraphs；必须按 body 元素顺序扫描。
    ★ 本样本 0 图 1 表 → 必须正确处理"整篇没图"：figure 相关检查在 figure 数为 0 时
      全部跳过（不报缺图、不报图编号断号），只在 meta.notes 里说明。
    ★ spec 里没有"每章都要有图表"的要求 → 只做编号自洽，不报"某章无图/无表"。
    """
    caps = []        # (type, ch, seq, idx, text, order)
    for b in ctx.blocks:
        if b[0] != 'p' or b[2] is None:
            continue
        m = CAP_RE.match(b[1].strip())
        if m:
            t = m.group(1)
            ch = int(m.group(2))
            seq = int(m.group(3)) if m.group(3) else 1
            caps.append((t, ch, seq, b[2], b[1].strip(), b[4]))

    tbl_blocks = [b for b in ctx.blocks if b[0] == 'tbl']

    fig_caps = [c for c in caps if c[0] == '图']
    tab_caps = [c for c in caps if c[0] == '表']

    if not fig_caps:
        ctx.notes.append('全文无图（0 张）。图表编号自洽检查仅针对表格；'
                         '不因此报"缺图"或"图编号断号"。')

    # ---- 题注位置：把表格与相邻题注配对，核对在上方还是下方
    fig_pos = ctx.std['caption_figure_position']   # default 'below'
    tab_pos = ctx.std['caption_table_position']    # default 'above'
    fig_pos_src = ctx.src['caption_figure_position']
    tab_pos_src = ctx.src['caption_table_position']

    # 给每个表格找最近的同类型题注（窗口 ±4 个 order）
    def nearest_cap(tbl_order, ctype):
        best = None
        for c in caps:
            if c[0] != ctype:
                continue
            d = abs(c[5] - tbl_order)
            if d <= 4 and (best is None or d < best[0]):
                best = (d, c)
        return best[1] if best else None

    for b in tbl_blocks:
        c = nearest_cap(b[4], '表')
        if c is None:
            continue
        above = c[5] < b[4]
        actual_pos = 'above' if above else 'below'
        if actual_pos != tab_pos:
            ctx.add('figure-table', 'warn',
                    key='pos|表|%d' % c[3],
                    expected='表题应在表格%s（依据 caption_table_position=%s）'
                             % ('上方' if tab_pos == 'above' else '下方', tab_pos),
                    actual='表「%s」的题注在其%s' % (c[4], '上方' if above else '下方'),
                    suggestion='把表题移到表格%s（移动题注属版式手术，工具不自动改，'
                               '请人工处理）。' % ('上方' if tab_pos == 'above' else '下方'),
                    source=tab_pos_src, para_index=c[3], text=c[4], action=None)

    # 图题位置（仅在存在图时才检查；本样本 0 图，下面图编号检查会直接跳过）

    # ---- 编号自洽：按 (类型, 章) 分组，段号须连续、不重复
    for ctype, cap_list, default_pos in (('表', tab_caps, tab_pos), ('图', fig_caps, fig_pos)):
        if not cap_list:
            continue
        by_ch = {}
        for c in cap_list:
            by_ch.setdefault(c[1], []).append(c)
        for ch, items in by_ch.items():
            segs = sorted(i[2] for i in items)
            seg_set = set(segs)
            if segs and segs != list(range(segs[0], segs[-1] + 1)):
                missing = [n for n in range(segs[0], segs[-1] + 1) if n not in seg_set]
                for n in missing:
                    ctx.add('figure-table', 'error',
                            key='gap|%s|%d|%d' % (ctype, ch, n),
                            expected='%s %d 章内，%s题编号应连续' % (ctype, ch, ctype),
                            actual='%s %d 章的%s题编号缺失 %d（现有：%s）'
                                   % (ctype, ch, ctype, n,
                                      '、'.join(map(str, sorted(seg_set)))),
                            suggestion='补齐缺失的%s题编号（改编号属全局手术，'
                                       '工具不自动改，请人工处理）。' % ctype,
                            para_index=items[0][3], text=items[0][4], action=None)
            for n in seg_set:
                if segs.count(n) > 1:
                    ctx.add('figure-table', 'error',
                            key='dup|%s|%d|%d' % (ctype, ch, n),
                            expected='%s题编号应唯一' % ctype,
                            actual='%s %d-%d 出现 %d 次（重复）'
                                   % (ctype, ch, n, segs.count(n)),
                            suggestion='修正重复的%s题编号（请人工处理）。' % ctype,
                            para_index=items[0][3], text=items[0][4], action=None)

    # ---- 正文引用存在性：如图 x-y 所示 / 见表 x-y 必须真有该题注
    if fig_caps or tab_caps:
        have = {(c[0], c[1], c[2]) for c in caps}
        body_text = ''.join(b[1] for b in ctx.blocks if b[3] == 'body')
        for m in FIGREF_RE.finditer(body_text):
            ctype = m.group(1)
            ch = int(m.group(2))
            seq = int(m.group(3))
            # 题注段落本身也含"表 1-1"，应排除；用"如图/见表/参见图"等引用词才算引用
            pre = body_text[max(0, m.start() - 3):m.start()]
            if pre.rstrip().endswith(('图', '表')):
                continue
            if (ctype, ch, seq) not in have:
                ctx.add('figure-table', 'warn',
                        key='ref|%s|%d|%d' % (ctype, ch, seq),
                        expected='正文引用的%s %d-%d 应在文中真实存在对应的%s题'
                                 % (ctype, ch, seq, ctype),
                        actual='正文出现「%s %d-%d」的引用，但文中没有该题注'
                               % (ctype, ch, seq),
                        suggestion='检查是编号写错，还是漏放了%s题。' % ctype,
                        action=None)


# ================================================================ 规则 3：abstract-layout

def _abstract_regions(ctx):
    """返回 (cn_head_idx, en_head_idx, kw_idx, toc_idx)。找不到对应为 None。"""
    cn = en = kw = None
    for i, p in enumerate(ctx.paras, 1):
        t = p.text.strip()
        if cn is None and p.style.name == '摘要标题':
            cn = i
        if en is None and p.style.name == '英文摘要标题':
            en = i
        if kw is None and t.startswith('关键词'):
            kw = i
    toc = ctx.regions['toc']
    return cn, en, kw, toc


def rule_abstract_layout(ctx):
    """中英文摘要：字数上限、关键词、摘要内不应有引注/图表引用。

    ★ "摘要是否单独成页"在 docx 层面判不出来（除非有分页符），判不出就写进
      meta.notes 说明"需导出 PDF 后确认"，不猜。
    ★ 上限只认 spec 的 abstract.cn.max_chars / abstract.en.max_words；
      没有就走 default（1000 / 500）并标 standard_source='default'。
      **绝不**用 template-extract 测出的范文实测字数当阈值。
    """
    cn, en, kw, toc = _abstract_regions(ctx)
    if cn is None:
        ctx.notes.append('未找到中文摘要标题，跳过 abstract-layout 检查。')
        return

    # 中文摘要正文：cn_head 之后、关键词行与 en_head 之前
    cn_end = (kw if kw and kw < en else en) or (en if en else toc)
    cn_text = ''
    if cn_end:
        for j in range(cn + 1, cn_end):
            cn_text += ctx.paras[j - 1].text
    else:
        cn_text = ''

    # 英文摘要正文：en_head 之后、toc/参考文献/一级标题 之前
    en_end = None
    for j in range((en or 0) + 1, len(ctx.paras) + 1):
        nt = norm(ctx.paras[j - 1].text)
        if nt in {'目录', '参考文献', 'references', 'bibliography'} or \
           ctx.paras[j - 1].style.name == '一级标题':
            en_end = j
            break
    en_text = ''
    if en:
        for j in range(en + 1, en_end or (len(ctx.paras) + 1)):
            en_text += ctx.paras[j - 1].text

    max_cn = ctx.std['abstract_cn_max_chars']
    max_cn_src = ctx.src['abstract_cn_max_chars']
    max_en = ctx.std['abstract_en_max_words']
    max_en_src = ctx.src['abstract_en_max_words']

    cn_chars = len(HANZI_RE.findall(cn_text))
    en_words = len([w for w in re.split(r'\s+', en_text.strip()) if re.search(r'[A-Za-z]', w)])

    if cn_chars > max_cn:
        ctx.add('abstract-layout', 'error',
                key='cn-over|%d' % cn,
                expected='中文摘要汉字数 ≤ %d（上限来自 abstract_cn_max_chars=%s，'
                         'standard_source=%s）' % (max_cn, max_cn, max_cn_src),
                actual='中文摘要汉字数 = %d，超出上限 %d' % (cn_chars, max_cn),
                suggestion='压缩中文摘要至 %d 字以内（工具不擅自删改摘要内容）。' % max_cn,
                source=max_cn_src, para_index=cn, text=cn_text[:60], action=None)
    else:
        ctx.add('abstract-layout', 'info',
                key='cn-ok|%d' % cn,
                expected='中文摘要汉字数 ≤ %d（上限来源=%s）' % (max_cn, max_cn_src),
                actual='中文摘要汉字数 = %d（未超限）' % cn_chars,
                suggestion='中文摘要字数正常。',
                source=max_cn_src, para_index=cn, text=cn_text[:40], action=None)

    if en_words > max_en:
        ctx.add('abstract-layout', 'error',
                key='en-over|%d' % (en or 0),
                expected='英文摘要词数 ≤ %d（上限来自 abstract_en_max_words=%s，'
                         'standard_source=%s）' % (max_en, max_en, max_en_src),
                actual='英文摘要词数 = %d，超出上限 %d' % (en_words, max_en),
                suggestion='压缩英文摘要至 %d 词以内。' % max_en,
                source=max_en_src, para_index=en, text=en_text[:60], action=None)
    else:
        ctx.add('abstract-layout', 'info',
                key='en-ok|%d' % (en or 0),
                expected='英文摘要词数 ≤ %d（上限来源=%s）' % (max_en, max_en_src),
                actual='英文摘要词数 = %d（未超限）' % en_words,
                suggestion='英文摘要词数正常。',
                source=max_en_src, para_index=en, text=en_text[:40], action=None)

    # 关键词：个数与分隔符一致性
    if kw is not None:
        line = ctx.paras[kw - 1].text
        body = re.sub(r'^关键词[：:]?', '', line).strip()
        # 候选分隔符：； ; ， ,
        seps = re.findall(r'[；;，,]', body)
        if seps:
            dominant = max(set(seps), key=seps.count)
            inconsistent = any(s != dominant for s in seps)
            parts = re.split(r'[；;，,]', body)
            count = len([x for x in parts if x.strip()])
        else:
            count, inconsistent = (1 if body.strip() else 0), False
        if count == 0:
            ctx.add('abstract-layout', 'warn',
                    key='kw-empty|%d' % kw,
                    expected='摘要应有中文关键词',
                    actual='未解析到关键词（关键词行：%s）' % line[:40],
                    suggestion='补全关键词。', para_index=kw, text=line, action=None)
        elif inconsistent:
            ctx.add('abstract-layout', 'warn',
                    key='kw-sep|%d' % kw,
                    expected='关键词应使用统一的分隔符（；或 , 等）',
                    actual='关键词行混用了多种分隔符：%s' % line[:50],
                    suggestion='统一关键词分隔符（本样本用"；"）。',
                    para_index=kw, text=line, action=None)
        else:
            ctx.add('abstract-layout', 'info',
                    key='kw-ok|%d' % kw,
                    expected='关键词应有合理个数并使用一致分隔符',
                    actual='关键词 %d 个，分隔符统一（%s）' % (count, dominant),
                    suggestion='关键词正常。', para_index=kw, text=line, action=None)

    # 摘要内不应有引注
    if re.search(r'\[[0-9][0-9,，、-]*\]', cn_text) or \
       re.search(r'\[[0-9][0-9,，、-]*\]', en_text):
        ctx.add('abstract-layout', 'warn',
                key='cite-in-abs',
                expected='摘要（中英文）内不应出现引注（[n]）',
                actual='摘要正文检出疑似引注编号',
                suggestion='摘要一般不放引注；若确需引用，确认学校规范允许。',
                action=None)

    # 摘要内不应有图表引用
    if re.search(r'[图图表表]\s*[0-9一二三四五六七八九十]+[-—][0-9]+', cn_text + en_text):
        ctx.add('abstract-layout', 'warn',
                key='figref-in-abs',
                expected='摘要内不应引用图表（如图 1-1 / 表 2-3）',
                actual='摘要正文检出疑似图表引用',
                suggestion='摘要一般用文字概括，移出图表引用。',
                action=None)

    # "是否单独成页"判不出
    ctx.notes.append('摘要是否单独成页在 docx 层面判不出（除非有分页符 '
                     'w:br w:type="page"）；如需确认，请导出 PDF 后人工核对。')


# ================================================================ 规则 4：length

def _count(text):
    """返回三个口径：(汉字数, 计空格的全角字符数, 不含空白的字符数)。"""
    hanzi = len(HANZI_RE.findall(text))
    full = len(FULLWIDTH_RE.findall(text))
    nows = len(re.sub(r'[\s\u3000]', '', text))
    return hanzi, full, nows


def rule_length(ctx):
    """双口径字数 + 分章配额。

    三个口径（定义写死，报告里明示）：
      ① 汉字数            = 匹配 [\\u4e00-\\u9fff] 的字符数
      ② 计空格的全角字符数 = 全角字符（CJK + 全角标点 + 全角 forms，含全角空格 U+3000）
      ③ 不含空白的字符数  = 总字符数 − 所有空白（\\s 与全角空格 U+3000）

    每套都给"含参考文献与附录"和"不含"两套：
      - 含 = 全文档（front + body + ref + appendix 全部段落与表格）
      - 不含 = 仅正文（body：第一章到参考文献之前的章节，排除摘要/目录/参考文献/致谢/附录）

    再按一级标题切段，列出各章三个口径字数与占比。
    与 std['body_min_chars']（default None → 不判）对比。
    """
    # 全文档文本
    full_text = ''.join(b[1] for b in ctx.blocks)
    # 正文文本（不含 参考文献与附录，亦排除前部摘要/目录）
    body_text = ''.join(b[1] for b in ctx.blocks if b[3] == 'body')

    full = _count(full_text)
    body = _count(body_text)

    # 分章：按一级标题切段，把每段/表格归到当前章
    chapters = []  # (title, hanzi, full, nows)
    cur = None
    for b in ctx.blocks:
        if b[0] == 'p' and b[2] is not None:
            p = ctx.paras[b[2] - 1]
            if p.style.name == '一级标题':
                if cur is not None:
                    chapters.append(cur)
                cur = [p.text.strip(), 0, 0, 0]
        if cur is not None and b[3] == 'body':
            h, f, n = _count(b[1])
            cur[1] += h
            cur[2] += f
            cur[3] += n
    if cur is not None:
        chapters.append(cur)

    body_total = body[0]
    by_chapter = []
    for title, h, f, n in chapters:
        pct = (h / body_total * 100.0) if body_total else 0.0
        by_chapter.append({'chapter': title, 'hanzi': h, 'fullwidth': f,
                           'no_whitespace': n, 'pct': round(pct, 1)})

    body_min = ctx.std['body_min_chars']
    body_min_src = ctx.src['body_min_chars']

    # 结构化数据（供测试交叉验证 & 报告表头）
    ctx.length_stats = {
        'with_ref_appendix': {'hanzi': full[0], 'fullwidth': full[1], 'no_whitespace': full[2]},
        'without_ref_appendix': {'hanzi': body[0], 'fullwidth': body[1], 'no_whitespace': body[2]},
        'by_chapter': by_chapter,
        'body_min_chars': body_min,
        'body_min_chars_source': body_min_src,
    }

    def fmt(d):
        return '汉字数=%d，计空格的全角字符数=%d，不含空白的字符数=%d' % (
            d['hanzi'], d['fullwidth'], d['no_whitespace'])

    expected = ('三个口径定义：①汉字数=[\\u4e00-\\u9fff]计数；'
                '②计空格的全角字符数=CJK+全角标点+全角forms（含U+3000全角空格）；'
                '③不含空白的字符数=总字符−\\s与全角空格。')

    actual = ('含参考文献与附录：%s。\n不含参考文献与附录（仅正文）：%s。\n'
              '按章拆分（汉字数 / 占比）：%s'
              % (fmt(ctx.length_stats['with_ref_appendix']),
                 fmt(ctx.length_stats['without_ref_appendix']),
                 '；'.join('%s：%d字(%.1f%%)' % (c['chapter'], c['hanzi'], c['pct'])
                           for c in by_chapter)))

    ctx.add('length', 'info', key='length-summary',
            expected=expected, actual=actual,
            suggestion='字数为只读统计，供你把握篇幅；如需"正文字数下限"判定，'
                       '请在 spec 里设置 length.body_min_chars。',
            action=None)

    if body_min is not None and body[0] < body_min:
        ctx.add('length', 'warn', key='body-too-short',
                expected='正文字数（汉字数）≥ %d（来自 body_min_chars，source=%s）'
                         % (body_min, body_min_src),
                actual='正文字数（汉字数）= %d，低于下限 %d' % (body[0], body_min),
                suggestion='扩充正文或确认下限设置是否合理。',
                source=body_min_src, action=None)


RULES = [rule_heading_sequence, rule_figure_table, rule_abstract_layout, rule_length]


# ================================================================ 主流程

def audit(docx_path, spec_path=None, ignore_path=None, verbose=True):
    ctx = Ctx(docx_path, spec_path)
    for fn in RULES:
        try:
            fn(ctx)
        except Exception as e:
            import traceback
            ctx.notes.append('规则 `%s` 执行异常：%s: %s' % (fn.__name__, type(e).__name__, e))
            ctx.issues.append({
                'id': issue_id('audit-internal', fn.__name__),
                'rule': 'audit-internal-error', 'layer': LAYER, 'skill': SKILL,
                'severity': 'error',
                'location': {'kind': 'document', 'index': None, 'excerpt': None, 'offset': None},
                'expected': '每条规则都应正常跑完',
                'actual': '规则 `%s` 执行异常：%s: %s' % (fn.__name__, type(e).__name__, e),
                'standard_source': 'default',
                'suggestion': '这是工具缺陷，不是文档问题。本次结果**不可信**，'
                              '请把这条报给维护者；排查线索：%s'
                              % traceback.format_exc().splitlines()[-1],
                'suppressible': False, 'action': None,
            })

    # heading-sequence 的"三级标题数量上限"归 text-style 的 heading-style 管，
    # 这里只写一句说明，避免同一件事报两遍。
    ctx.notes.insert(0, '三级标题数量上限（std[max_h3_per_h2]）由 text-style 的 '
                        'heading-style 规则负责，本 skill 不重复报。')

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
            'abstract_cn_max_chars': ctx.std['abstract_cn_max_chars'],
            'abstract_cn_max_source': ctx.src['abstract_cn_max_chars'],
            'abstract_en_max_words': ctx.std['abstract_en_max_words'],
            'abstract_en_max_source': ctx.src['abstract_en_max_words'],
            'caption_figure_position': ctx.std['caption_figure_position'],
            'caption_figure_position_source': ctx.src['caption_figure_position'],
            'caption_table_position': ctx.std['caption_table_position'],
            'caption_table_position_source': ctx.src['caption_table_position'],
            'body_min_chars': ctx.std['body_min_chars'],
            'body_min_chars_source': ctx.src['body_min_chars'],
        },
        'standard_source_note': _std_note(ctx),
        'notes': ctx.notes,
        'length': getattr(ctx, 'length_stats', None),
    }
    if verbose:
        print('[structure-length] %s' % ctx.docx)
        print('  段落 %d / 表格 %d / 图片 %d'
              % (summary['paragraphs'], summary['tables'], summary['images']))
        print('  标准来源：%s' % _std_note(ctx))
        print('  问题 %d 条（按裁定保留 %d 条已单列）' % (len(kept), len(sup)))
    return {'issues': kept, 'suppressed': sup, 'summary': summary, 'meta': meta}


def _std_note(ctx):
    return describe(ctx.std, ctx.src,
                    ['abstract_cn_max_chars', 'abstract_en_max_words',
                     'caption_figure_position', 'caption_table_position',
                     'body_min_chars'])


def _sha256(path):
    from shared.lib.ooxml_guard import sha256_file
    return sha256_file(path)


def main():
    ap = argparse.ArgumentParser(description='L4 结构与篇幅审计（只读）')
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
