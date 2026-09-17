#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
skills/text-style/scripts/audit.py —— L3 文字体例审计（**只读，绝不写盘**）

五条默认规则（与 assets/skills-index.yaml 的 covers 一一对应）
    quote-style      引号体例：正文用直引号 " '（U+0022/U+0027），不用中文弯引号
    punctuation      标点异常：连续标点 / 半角标点被中文夹围 / 省略号误写 / 全半角括号混用
    dash-density     破折号密度：每千汉字出现次数 + 单段 ≥3 处
    heading-style    标题体例：过长 / 含冒号问号感叹号 / 含虚词 / 末尾带标点 / 二级下三级数
    process-trace    过程痕迹：TODO/TBD/FIXME/XXX/待补/待定/（略）/【】/？？？ 等

三条可选规则（默认关闭，--optional 才开；信噪比低，见 docs 设计方案 v2 §4）
    cjk-spacing        中英文/中文数字间空格是否统一（只报"不统一"）
    fullwidth-halfwidth 全角数字字母与半角混用
    term-consistency   术语一致性（读 assets/rules/term-map.yaml；不存在则跳过并说明）

设计上刻意"保守"的两处（不是没做，是故意不做）
    ① 全部文字体例问题都不给自动修复动作。弯引号改直引号、标点修正、标题平实化，
       都需要结合上下文（成对引号、嵌套、作者意图），机器盲目替换会改错，故
       action 一律为 None，并在 suggestion 里说明原因。apply 不会碰它们。
    ② quote-style 按段落聚合：一段里 N 处弯引号只报 1 条，不每处一条（信噪比第一）。
       一份稿子里有引号是正常的，翻成一堆噪声等于没用。

标准来源三级优先（每条 Issue 都带 standard_source，报告里会显著标出）
    template（模板/范文 spec） > config（用户配置） > default（通用学术惯例）

★ 关键坑（已按约束处理）
    · 引号/破折号/省略号一律 unicode 转义写死（\u201c \u201d \u2018 \u2019 \u2014 \u2026），
      绝不在源码里直接敲中文引号——肉眼分不清 U+0022 与 U+201D。
    · 必须排除参考文献表 / 目录 / 页眉范围：本模块基于公开方法 reference_section()
      与 style_names() 自行推出范围，没有改 shared/。
    · 汉字计数用 [\u4e00-\u9fff]。

用法
    python audit.py -d "F:/论文/初稿.docx"
    python audit.py -d 初稿.docx -s spec.json -o out/
    python audit.py -d 初稿.docx --optional        # 连同三条可选规则一起跑
    python audit.py -d 初稿.docx --print            # 报告打到终端
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
from shared.lib.docx_scan import Scanner, excerpt, issue_id, norm  # noqa: E402
from shared.lib.standards import (                                # noqa: E402
    describe, find_spec, load_spec, resolve_standard,
)
from shared.lib import report as R                                # noqa: E402

SKILL = 'text-style'
LAYER = 'L3'

# ---- 文字常量（unicode 转义写死，不出现裸中文引号）----
CURVED = '\u201c\u201d\u2018\u2019'        # “ ” ‘ ’  中文弯引号
EM_DASH = '\u2014'                          # —  破折号
ELLIPSIS = '\u2026'                          # …  正确省略号
FW_OPEN, FW_CLOSE = '\uff08', '\uff09'       # （ ）
HW_OPEN, HW_CLOSE = '(', ')'

# 连续标点：重复（。。 ，， ！！）与跨类（，。 。，）
REPEAT_PUNCT_RE = re.compile(r'([。，！？；：、])\1+')
CROSS_PUNCT_RE = re.compile(r'(?:，。|。，)')
# 半角逗号/句号被中文夹围（约束里给的原样正则）
HW_BETWEEN_CJK_RE = re.compile(r'[\u4e00-\u9fff][,.;?!][\u4e00-\u9fff]')
# 省略号误写：三个及以上 ASCII 点，或三个及以上全角句号
ELLIPSIS_DOTS_RE = re.compile(r'\.\.\.+')
ELLIPSIS_PERIOD_RE = re.compile(r'。{3,}')
HANZI_RE = re.compile(r'[\u4e00-\u9fff]')

# 标题虚词（约束里列的，不多不少）
HEADING_VIRTUAL = ['浅析', '初探', '试论', '刍议', '略论']
HEADING_PUNCT_RE = re.compile(r'[：:？?！!]')          # 标题含冒号/问号/感叹号
HEADING_TRAIL_RE = re.compile(r'[。！？；，、：]\s*$')  # 标题末尾带标点

# 过程痕迹（unicode 转义写死，避免裸敲）
TRACE_TOKENS = [
    'TODO', 'TBD', 'FIXME', 'XXX',
    '待补', '待定', '待核实', '此处省略',
    '\uff08\u7565\uff09',   # （略）
    '\u3010\u3011',         # 【】
    '\uff1f\uff1f\uff1f',   # ？？？
    '@@', '<<', '>>',
]

# 术语表默认位置（测试可临时覆盖 AUD.TERM_MAP_PATH）
TERM_MAP_PATH = os.path.join(PKG, 'assets', 'rules', 'term-map.yaml')

# rule -> 标准来源键（用于 ctx.src 查 source；None 表示默认走 default）
RULE_SRC = {
    'quote-style': 'quote_style',
    'dash-density': 'dash_per_1000_max',
    'heading-style': 'heading_max_chars',
}


# ================================================================ 上下文

class Ctx:
    """一次审计的全部共享状态。规则函数只读它，不互相传参。"""

    def __init__(self, docx, spec_path=None):
        self.docx = os.path.abspath(docx)
        self.scan = Scanner(self.docx)
        self.spec, self.spec_path = load_spec(spec_path)
        self.std, self.src, self.notes = resolve_standard(self.spec)
        self.issues = []
        self.excluded = self._compute_excluded()

    # -------------------------------------------------- 排除范围
    # 只基于公开方法：reference_section()（标题段号 + 各条目段号）与
    # style_names()（按文档顺序的样式名）。不碰 shared/ 的私有方法。

    def _compute_excluded(self):
        excl = set()
        # ① 参考文献表：从「参考文献」标题段到最后一个条目段
        h, entries = self.scan.reference_section()
        if h is not None:
            end = entries[-1].para_index if entries else h
            excl |= set(range(h, end + 1))
        # ② 目录：样式名含 'toc' 的段落（toc 1/2/3）+ 目录标题段
        for i, st in enumerate(self.scan.style_names(), 1):
            if st and 'toc' in st.lower():
                excl.add(i)
            if st == '目录标题':
                excl.add(i)
        # ③ 页眉/页脚：python-docx 的 document.paragraphs 不含页眉页脚段落，
        #    它们落在独立的 header/footer part，scan 直接不扫，天然已排除。
        return excl

    def in_body(self, i):
        return 1 <= i <= len(self.scan._paras) and i not in self.excluded

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
        if source is None:
            sk = RULE_SRC.get(rule)
            source = self.src.get(sk, 'default') if sk else 'default'
        loc = {'kind': kind, 'index': para_index, 'excerpt': ex,
               'offset': start if kind == 'run' else None}
        self.issues.append({
            'id': issue_id(rule, key),
            'rule': rule, 'layer': LAYER, 'skill': SKILL,
            'severity': severity, 'location': loc,
            'expected': expected, 'actual': actual,
            'standard_source': source,
            'suggestion': suggestion, 'suppressible': suppressible,
            'action': action,
        })

    def act(self, action_name, **params):
        return {'name': action_name, 'params': params}


# ================================================================ 规则

def _heading_level(style):
    """从样式名推断标题层级（1/2/3）；非章节标题返回 0。"""
    if style is None:
        return 0
    st = style
    low = style.lower()
    if '一级标题' in st or re.search(r'heading\s*1\b', low):
        return 1
    if '二级标题' in st or re.search(r'heading\s*2\b', low):
        return 2
    if '三级标题' in st or re.search(r'heading\s*3\b', low):
        return 3
    return 0


def rule_quote_style(ctx):
    """引号体例：正文用直引号，不用中文弯引号。按段落聚合（一段只报 1 条）。"""
    for i in range(1, len(ctx.scan._paras) + 1):
        if not ctx.in_body(i):
            continue
        t = ctx.scan.para_text(i)
        if not t:
            continue
        found = [c for c in t if c in CURVED]
        if not found:
            continue
        n = len(found)
        idx = next(j for j, c in enumerate(t) if c in CURVED)
        ctx.add('quote-style', 'warn',
                key='p%d' % i,
                expected='正文使用直引号 " 与 \' （U+0022 / U+0027），不使用中文弯引号 “ ” ‘ ’',
                actual='本段出现 %d 处中文弯引号（如 %s）' % (n, found[0]),
                suggestion='将弯引号改为直引号。本报告不自动替换——中文引号改直引号需结合上下文'
                           '（成对引号、嵌套层级），机器盲目替换可能改错，建议人工核对后统一。',
                para_index=i, text=t, start=idx, end=idx + 1,
                action=None)


def rule_punctuation(ctx):
    """标点异常：连续标点 / 半角标点被中文夹围 / 省略号误写 / 全半角括号混用。
    一段里多种异常合并成 1 条。"""
    for i in range(1, len(ctx.scan._paras) + 1):
        if not ctx.in_body(i):
            continue
        t = ctx.scan.para_text(i)
        if not t:
            continue
        anoms = []  # (label, example, pos)
        m = REPEAT_PUNCT_RE.search(t)
        if m:
            anoms.append(('连续标点', m.group(0), m.start()))
        m = CROSS_PUNCT_RE.search(t)
        if m:
            anoms.append(('连续标点（跨类）', m.group(0), m.start()))
        m = HW_BETWEEN_CJK_RE.search(t)
        if m:
            anoms.append(('半角标点被中文夹围', m.group(0), m.start()))
        m = ELLIPSIS_DOTS_RE.search(t)
        if m:
            anoms.append(('英文省略号 ...', m.group(0), m.start()))
        m = ELLIPSIS_PERIOD_RE.search(t)
        if m:
            anoms.append(('全角句号当省略号 。。。', m.group(0), m.start()))
        if (FW_OPEN in t or FW_CLOSE in t) and (HW_OPEN in t or HW_CLOSE in t):
            anoms.append(('全角半角括号混用', '（）()', -1))
        if not anoms:
            continue
        first_pos = min((p for _, _, p in anoms if p >= 0), default=None)
        ctx.add('punctuation', 'warn',
                key='p%d' % i,
                expected='标点规范：不连用、半角标点不被中文夹围、省略号用 %s×2、'
                         '同一段内括号全/半角统一' % ELLIPSIS,
                actual='本段存在：%s' % '；'.join('%s（%s）' % (lab, ex)
                                                 for lab, ex, _ in anoms),
                suggestion='人工修正标点。机器不自动改——盲目替换可能破坏数字/英文里的合法半角标点。',
                para_index=i, text=t,
                start=first_pos if first_pos is not None else None,
                end=(first_pos + 1) if first_pos is not None else None,
                action=None)


def rule_dash_density(ctx):
    """破折号密度：每千汉字出现次数（超阈值 warn）+ 单段 ≥3 处（info）。"""
    maxd = ctx.std['dash_per_1000_max']
    total_em, total_hanzi = 0, 0
    per = []  # (para_index, em_count, hanzi)
    for i in range(1, len(ctx.scan._paras) + 1):
        if not ctx.in_body(i):
            continue
        t = ctx.scan.para_text(i)
        e = t.count(EM_DASH)
        hz = len(HANZI_RE.findall(t))
        total_em += e
        total_hanzi += hz
        if e:
            per.append((i, e, hz))

    density = (total_em / (total_hanzi / 1000.0)) if total_hanzi else 0.0

    # 单段 ≥3 处：独立信号（即便整体密度未超也报，帮助定位聚类）
    for i, e, hz in per:
        if e >= 3:
            ctx.add('dash-density', 'info',
                    key='p%d' % i,
                    expected='单段破折号建议 < 3 处',
                    actual='段 %d 出现 %d 处破折号' % (i, e),
                    suggestion='本段破折号较密集，建议拆分句子或改用其他连接方式。',
                    para_index=i, text=ctx.scan.para_text(i),
                    action=None)

    # 整体密度超阈值：文档级 warn，并列出最密集的 3 段
    if total_hanzi > 0 and density > maxd:
        def _ratio(rec):
            _, e, hz = rec
            return e / (max(hz, 1) / 1000.0)
        top = sorted(per, key=_ratio, reverse=True)[:3]
        top_desc = '；'.join('段%d（%d处/%d字）' % (i, e, hz) for i, e, hz in top)
        ctx.add('dash-density', 'warn',
                key='doc',
                expected='破折号密度 ≤ %s 处/千汉字（当前 %.2f）' % (maxd, density),
                actual='全文破折号 %d 处，汉字 %d 字，密度 %.2f 处/千汉字；'
                       '最密集段落：%s' % (total_em, total_hanzi, density, top_desc),
                suggestion='整体破折号偏多，建议精简。机器不自动删——删改会动语义。',
                kind='document', action=None)


def rule_heading_style(ctx):
    """标题体例：过长 / 含冒号问号感叹号 / 含虚词 / 末尾带标点（逐标题聚合 1 条），
    另查每个二级标题下的三级标题数（std['max_h3_per_h2']）。"""
    max_chars = ctx.std['heading_max_chars']
    # ---- 逐标题检查
    for i in range(1, len(ctx.scan._paras) + 1):
        if not ctx.in_body(i):
            continue
        st = (ctx.scan._paras[i - 1].style.name
              if ctx.scan._paras[i - 1].style is not None else '') or ''
        L = _heading_level(st)
        if not L:
            continue
        t = ctx.scan.para_text(i).strip()
        if not t:
            continue
        vios = []
        if len(t) > max_chars:
            vios.append('标题过长（%d字>%d）' % (len(t), max_chars))
        if HEADING_PUNCT_RE.search(t):
            vios.append('含冒号/问号/感叹号')
        hit_v = [v for v in HEADING_VIRTUAL if v in t]
        if hit_v:
            vios.append('含虚词（%s）' % '、'.join(hit_v))
        if HEADING_TRAIL_RE.search(t):
            vios.append('标题末尾带标点')
        if vios:
            ctx.add('heading-style', 'warn',
                    key='h%d' % i,
                    expected='标题平实：长度≤%d字、不含冒号/问号/感叹号、不用虚词、末尾无标点'
                             % max_chars,
                    actual='「%s」：%s' % (t[:30], '；'.join(vios)),
                    suggestion='按体例平实化标题（去冒号/虚词/末尾标点，缩短过长标题）。'
                               '机器不自动改——标题改写涉及作者意图。',
                    para_index=i, text=t, action=None)

    # ---- 每个二级标题下的三级标题数
    max_h3 = ctx.std['max_h3_per_h2']
    cur_h2 = None
    counts = {}   # h2_index -> [count, text]
    for i in range(1, len(ctx.scan._paras) + 1):
        if not ctx.in_body(i):
            continue
        st = (ctx.scan._paras[i - 1].style.name
              if ctx.scan._paras[i - 1].style is not None else '') or ''
        L = _heading_level(st)
        if L == 2:
            if cur_h2 is not None and counts[cur_h2][0] > max_h3:
                _emit_h3(ctx, cur_h2, counts[cur_h2][0], counts[cur_h2][1], max_h3)
            cur_h2 = i
            counts[cur_h2] = [0, ctx.scan.para_text(i).strip()]
        elif L == 3 and cur_h2 is not None:
            counts[cur_h2][0] += 1
    if cur_h2 is not None and counts[cur_h2][0] > max_h3:
        _emit_h3(ctx, cur_h2, counts[cur_h2][0], counts[cur_h2][1], max_h3)


def _emit_h3(ctx, h2_index, count, h2_text, max_h3):
    ctx.add('heading-style', 'warn',
            key='h2-%d' % h2_index,
            expected='每个二级标题下的三级标题数 ≤ %d 个（当前 %d）' % (max_h3, count),
            actual='「%s」下挂了 %d 个三级标题' % (h2_text[:30], count),
            suggestion='按体例控制三级标题数量（合并或升级其中若干）。',
            para_index=h2_index, text=h2_text,
            source=ctx.src.get('max_h3_per_h2', 'default'),
            action=None)


def rule_process_trace(ctx):
    """过程痕迹残留：TODO/TBD/FIXME/XXX/待补/待定/（略）/【】/？？？ 等。
    一段里多种痕迹合并成 1 条。"""
    for i in range(1, len(ctx.scan._paras) + 1):
        if not ctx.in_body(i):
            continue
        t = ctx.scan.para_text(i)
        if not t:
            continue
        hits = []
        for tk in TRACE_TOKENS:
            pos = t.find(tk)
            if pos >= 0:
                hits.append((tk, pos))
        if not hits:
            continue
        hits.sort(key=lambda x: x[1])
        first_pos = hits[0][1]
        ctx.add('process-trace', 'warn',
                key='p%d' % i,
                expected='正文不应残留写作过程痕迹（TODO/TBD/FIXME/待补/待定/（略）/【】 等）',
                actual='本段残留：%s' % '、'.join('%s' % tk for tk, _ in hits),
                suggestion='确认后删除这些写作痕迹。机器不自动删——需先确认是否为占位待补内容。',
                para_index=i, text=t, start=first_pos, end=first_pos + len(hits[0][0]),
                action=None)


# ================================================================ 可选规则

def rule_cjk_spacing(ctx):
    """中英文 / 中文与数字之间空格是否统一（只报"不统一"，不报"该不该加"）。"""
    # 直接相邻（无空格）的 CJK↔ASCII 边界
    adj_re = re.compile(r'[\u4e00-\u9fff][A-Za-z0-9]|[A-Za-z0-9][\u4e00-\u9fff]')
    # 中间有空格的 CJK↔ASCII 边界
    sp_re = re.compile(r'[\u4e00-\u9fff]\s+[A-Za-z0-9]|[A-Za-z0-9]\s+[\u4e00-\u9fff]')
    for i in range(1, len(ctx.scan._paras) + 1):
        if not ctx.in_body(i):
            continue
        t = ctx.scan.para_text(i)
        if not t:
            continue
        n_adj = len(adj_re.findall(t))
        n_sp = len(sp_re.findall(t))
        if n_adj and n_sp and (n_adj + n_sp) >= 2:
            ctx.add('cjk-spacing', 'info',
                    key='p%d' % i,
                    expected='中英文/中文与数字之间空格应统一（要么都加、要么都不加）',
                    actual='本段既有"不加空格"边界 %d 处、又有"加空格"边界 %d 处，不统一'
                           % (n_adj, n_sp),
                    suggestion='统一空格写法（加或不加，二选一并保持全文一致）。机器不自动改。',
                    para_index=i, text=t, action=None)


def rule_fullwidth_halfwidth(ctx):
    """全角数字/字母（０-９ａ-ｚＡ-Ｚ）与半角混用。"""
    fw_re = re.compile(r'[\uff10-\uff19\uff21-\uff3a\uff41-\uff5a]')  # ０-９Ａ-Ｚａ-ｚ
    hw_re = re.compile(r'[A-Za-z0-9]')
    for i in range(1, len(ctx.scan._paras) + 1):
        if not ctx.in_body(i):
            continue
        t = ctx.scan.para_text(i)
        if not t:
            continue
        if fw_re.search(t) and hw_re.search(t):
            ctx.add('fullwidth-halfwidth', 'info',
                    key='p%d' % i,
                    expected='数字/字母全角与半角不混用（统一用半角）',
                    actual='本段同时出现全角与半角数字/字母',
                    suggestion='统一为半角数字字母。机器不自动改——需逐字核对。',
                    para_index=i, text=t, action=None)


def rule_term_consistency(ctx):
    """术语一致性：同一概念多种写法混用（读 TERM_MAP_PATH）。
    文件不存在则跳过并在 meta.notes 说明。"""
    path = TERM_MAP_PATH
    if not os.path.exists(path):
        ctx.notes.append('term-consistency：未找到 %s，跳过术语一致性检查。'
                         '（如需此项，请把术语映射表放到该路径）' % path)
        return
    try:
        import yaml
        with open(path, encoding='utf-8') as f:
            mapping = yaml.safe_load(f) or {}
    except Exception as e:
        ctx.notes.append('term-consistency：读取术语表失败（%s: %s），跳过。'
                         % (type(e).__name__, e))
        return
    if not isinstance(mapping, dict):
        ctx.notes.append('term-consistency：术语表格式应为 {概念: [写法1, 写法2, ...]}，跳过。')
        return

    for concept, variants in mapping.items():
        if not isinstance(variants, list) or len(variants) < 2:
            continue
        seen = {}  # variant -> [para_index, ...]
        for i in range(1, len(ctx.scan._paras) + 1):
            if not ctx.in_body(i):
                continue
            t = ctx.scan.para_text(i)
            if not t:
                continue
            for v in variants:
                if v and v in t:
                    seen.setdefault(v, []).append(i)
        used = list(seen.keys())
        if len(used) > 1:
            detail = '、'.join('%s(段%s)' % (v, ','.join(map(str, seen[v][:3])))
                               for v in used)
            ctx.add('term-consistency', 'warn',
                    key='term-%s' % norm(str(concept)),
                    expected='术语「%s」全文统一为一种写法' % concept,
                    actual='出现多种写法：%s' % detail,
                    suggestion='统一为其中一种写法（保留最通用/最准确者）。机器不自动改。',
                    kind='document', action=None)


RULES = [rule_quote_style, rule_punctuation, rule_dash_density,
         rule_heading_style, rule_process_trace]
OPTIONAL_RULES = [rule_cjk_spacing, rule_fullwidth_halfwidth, rule_term_consistency]


# ================================================================ 主流程

def audit(docx_path, spec_path=None, ignore_path=None, verbose=True, optional=False):
    ctx = Ctx(docx_path, spec_path)
    rules = list(RULES)
    if optional:
        rules += list(OPTIONAL_RULES)
    for fn in rules:
        try:
            fn(ctx)
        except Exception as e:
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
        'optional': optional,
        'standard': {
            'quote_style': ctx.std['quote_style'],
            'quote_source': ctx.src['quote_style'],
            'dash_per_1000_max': ctx.std['dash_per_1000_max'],
            'dash_source': ctx.src['dash_per_1000_max'],
            'heading_max_chars': ctx.std['heading_max_chars'],
            'heading_source': ctx.src['heading_max_chars'],
            'max_h3_per_h2': ctx.std['max_h3_per_h2'],
            'h3_source': ctx.src['max_h3_per_h2'],
        },
        'standard_source_note': _std_note(ctx),
        'notes': ctx.notes,
    }
    if verbose:
        print('[text-style] %s' % ctx.docx)
        print('  段落 %d，排除范围 %d 段（参考文献/目录/页眉）'
              % (summary['paragraphs'], len(ctx.excluded)))
        print('  标准来源：%s' % _std_note(ctx))
        print('  问题 %d 条（按裁定保留 %d 条已单列）%s'
              % (len(kept), len(sup),
                 '（含可选规则）' if optional else ''))
    return {'issues': kept, 'suppressed': sup, 'summary': summary, 'meta': meta}


def _std_note(ctx):
    return describe(ctx.std, ctx.src,
                   ['quote_style', 'dash_per_1000_max',
                    'heading_max_chars', 'max_h3_per_h2'])


def _sha256(path):
    from shared.lib.ooxml_guard import sha256_file
    return sha256_file(path)


def main():
    ap = argparse.ArgumentParser(description='L3 文字体例审计（只读）')
    ap.add_argument('-d', '--docx', required=True)
    ap.add_argument('-s', '--spec', default=None, help='不传则自动在文档旁找 spec.json')
    ap.add_argument('-i', '--ignore', default=None, help='.thesisignore 路径')
    ap.add_argument('-o', '--out', default=None, help='输出目录')
    ap.add_argument('--optional', action='store_true',
                    help='连同三条可选规则（cjk-spacing/fullwidth-halfwidth/term-consistency）一起跑')
    ap.add_argument('--print', action='store_true', help='把报告打到终端')
    ap.add_argument('--json', action='store_true', help='打印 JSON')
    a = ap.parse_args()

    spec = a.spec or find_spec(a.docx)
    res = audit(a.docx, spec, a.ignore, optional=a.optional)
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
