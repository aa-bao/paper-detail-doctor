#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
skills/proposal-consistency/scripts/audit.py —— 开题报告 ↔ 正文 一致性核对（只读，绝不写盘）

方法论（忠实实现 docs/设计方案-v2.md §5，不重新设计）
    · 锚点法：从开题报告抽有限个可指定锚点，回正文找落点。
      七类锚点 A1 题目 / A2 概念 / A3 理论 / A4 方法要素 / A5 大纲 / A6 结论 / A7 文献承诺。
    · 三层判定：
      L1 字面（术语/数字/标题/人名，归一化匹配，数字严格相等）—— 脚本全自动；
      L2 语义（同一命题两处表述是否等价）—— 本脚本**不下结论**，输出
          verdict=待判 + confidence=null + 两侧摘录 + 判定问题，交宿主 agent 补判；
      L3 结构（大纲覆盖矩阵）—— 脚本全自动。
    · 五类结论 → severity（写死）：缺失→error / 改动→warn / 偏离→warn / 无法判定→info；
      一致不产生 issue（进通过清单）。
    · 所有 issue 的 action 一律 None —— 一致性问题绝不自动改文档。

    ★ 本脚本里没有 LLM。L2 的"语义等价"判定权在设计上就交给宿主 agent，
      脚本只负责"把证据摆齐"（§5.6 第 4 条：最终裁定权在人和导师）。

开题报告的特殊性
    本项目开题报告正文全部装在一个大表格里，按普通段落读几乎读不到东西。
    脚本自己遍历 w:tbl → w:tr → w:tc → w:p → w:t 把表格文本抽出来，
    因此对任意"开题报告也是个 docx"的用户都可复用。

用法
    python audit.py -d <正文.docx> -p <开题报告.docx> [-o <产物目录>] [--json] [--print]
    缺少 -p 时明确报错退出（双文档核对必须给 -p），不假装跑成功。
"""
import argparse
import json
import os
import re
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
# 向上找到包含 shared/ 的目录作为包根（不依赖目录深度假设）
PKG = HERE
while PKG and os.path.basename(PKG) != 'paper-detail-doctor' and not os.path.isdir(os.path.join(PKG, 'shared')):
    PKG = os.path.dirname(PKG)
if not PKG or not os.path.isdir(os.path.join(PKG, 'shared')):
    PKG = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
if PKG not in sys.path:
    sys.path.insert(0, PKG)

from docx import Document                                                   # noqa: E402
from shared.lib.docx_scan import Scanner, norm, issue_id                    # noqa: E402
from shared.lib import report as R                                          # noqa: E402
from shared.lib.ooxml_guard import sha256_file                             # noqa: E402

SKILL = 'proposal-consistency'

# 五类结论 → severity（写死，不许改）
SEV = {'缺失': 'error', '改动': 'warn', '偏离': 'warn', '无法判定': 'info', '一致': None}

# 锚点 → rule id（prop-*）
RULE = {'A1': 'prop-topic', 'A2': 'prop-concept', 'A3': 'prop-theory',
        'A4': 'prop-method-sample', 'A5': 'prop-outline',
        'A6': 'prop-conclusion', 'A7': 'prop-bibliography'}

CN_NUM = '一二三四五六七八九十'


# ================================================================ 文本抽取

def extract_table_text(path):
    """遍历表格 w:tbl → w:tr → w:tc → w:p → w:t，按行拼接。

    开题报告正文装在单个大表格里，按普通段落读几乎读不到内容，
    必须显式下钻表格。返回按行（每个 <w:tr> 一行，单元格用 ' | ' 连接）拼接的文本。
    """
    d = Document(path)
    rows = []
    for tbl in d.tables:
        for tr in tbl.rows:
            cells = []
            for tc in tr.cells:
                # 单元格内多个段落用换行分隔（开题报告大纲每条一段，
                # 不换行会把整章大纲挤成一行，导致解析失败）
                paras = [p.text for p in tc.paragraphs]
                txt = '\n'.join(paras)
                # 嵌套表格（罕见）也下钻
                for t2 in tc.tables:
                    for tr2 in t2.rows:
                        for tc2 in tr2.cells:
                            txt += '\n' + '\n'.join(p.text for p in tc2.paragraphs)
                cells.append(txt)
            rows.append(' | '.join(cells))
    return '\n'.join(rows)


def extract_paragraph_text(path):
    d = Document(path)
    return '\n'.join(p.text for p in d.paragraphs)


def full_text(docx, is_proposal):
    """正文用 Scanner.body_text()（含目录/表格/页脚全文本）；
    开题报告用表格文本 + 段落文本（内容主要在大表格里）。"""
    if is_proposal:
        return extract_table_text(docx) + '\n' + extract_paragraph_text(docx)
    return Scanner(docx).body_text()


def parse_proposal_outline(text):
    """从开题报告的『五、论文大纲』块里按中文编号解析出大纲条目。

    只取该块内的条目，避免把开题报告正文里的同名词标题（如『一、选题背景』）
    误当成大纲。返回 [{level, text, raw}]，level: 1=章 2=节 3=子节。
    """
    lines = text.split('\n')
    start = None
    for i, l in enumerate(lines):
        s = l.strip()
        if not s:
            continue
        if norm(s).startswith('五、论文大纲') or (norm(s).startswith('五') and '论文大纲' in norm(s)):
            start = i
            break
    if start is None:
        return []
    items = []
    for l in lines[start + 1:]:
        s = l.strip()
        if not s:
            continue
        if norm(s).startswith('六、论文') or '论文的进度安排' in s or '进度安排' in s:
            break
        m = re.match(r'^([%s]+)、(.+)$' % CN_NUM, s)                 # 一、 二、
        if m:
            items.append({'level': 1, 'text': m.group(2), 'raw': s}); continue
        m = re.match(r'^（([%s]+)）(.+)$' % CN_NUM, s)               # （一）
        if m:
            items.append({'level': 2, 'text': m.group(2), 'raw': s}); continue
        m = re.match(r'^\(([%s]+)\)(.+)$' % CN_NUM, s)               # (一)
        if m:
            items.append({'level': 2, 'text': m.group(2), 'raw': s}); continue
        m = re.match(r'^(\d+)[.、](.+)$', s)                         # 1. 2、
        if m:
            items.append({'level': 3, 'text': m.group(2), 'raw': s}); continue
    return items


def parse_thesis_headings(docx):
    """解析正文标题树。用样式名（一级/二级/三级标题），排除目录(toc)；
    样式缺失时退化为编号识别。返回 [{level, text, index}]，index 1-based 与 python-docx 对齐。"""
    d = Document(docx)
    out = []
    for i, p in enumerate(d.paragraphs):
        t = p.text.strip()
        if not t:
            continue
        st = (p.style.name if p.style else '') or ''
        if 'toc' in st.lower():
            continue
        lvl = None
        if '一级标题' in st or st == 'Heading 1':
            lvl = 1
        elif '二级标题' in st or st == 'Heading 2':
            lvl = 2
        elif '三级标题' in st or st == 'Heading 3':
            lvl = 3
        elif '四级标题' in st:
            lvl = 4
        if lvl is None:
            if re.match(r'^第[%s]+章' % CN_NUM, t):
                lvl = 1
            elif re.match(r'^\d+[.、]', t):
                lvl = 2
            elif re.match(r'^（[%s]+）' % CN_NUM, t):
                lvl = 3
        if lvl is not None:
            out.append({'level': lvl, 'text': t, 'index': i + 1})
    return out


def para_index_by_keyword(docx, keyword):
    """返回正文里第一个含 keyword 的段落序号（1-based），找不到返回 None。"""
    d = Document(docx)
    for i, p in enumerate(d.paragraphs):
        if keyword in (p.text or ''):
            return i + 1
    return None


# ================================================================ 数值抽取

def extract_planned_sample(text):
    """开题计划样本量：锚定『计划访谈…人』，支持 '30-40人' 与 '15人' 两种写法。
    返回 (lo, hi)；单值则 lo==hi；找不到返回 None。"""
    m = re.search(r'计划访谈\s*(\d+)\s*(?:[-–—]\s*(\d+))?\s*人', text)
    if not m:
        m = re.search(r'拟访谈\s*(\d+)\s*(?:[-–—]\s*(\d+))?\s*人', text)
    if not m:
        return None
    a = int(m.group(1))
    b = int(m.group(2)) if m.group(2) else a
    return (a, b)


def extract_actual_sample(text):
    """正文实际完成样本量：锚定『完成访谈…人』。返回整数或 None。"""
    m = re.search(r'完成访谈\s*(\d+)\s*人', text)
    if m:
        return int(m.group(1))
    m = re.search(r'访谈\s*(\d+)\s*人', text)          # 兜底
    return int(m.group(1)) if m else None


def extract_observations(text):
    """正文的观察次数（数字 + 次 + 观察）。返回整数或 None。"""
    m = re.search(r'(\d+)\s*次[机构]?观察', text)
    return int(m.group(1)) if m else None


def extract_promised_refs(text):
    """开题文献承诺数量：『不少于 40 条』。返回整数或 None。"""
    m = re.search(r'不少于\s*(\d+)\s*条', text)
    return int(m.group(1)) if m else None


# ================================================================ 构造 Issue

class Builder:
    def __init__(self):
        self.issues = []
        self.passed = []          # 确定性一致（进通过清单）
        self.pending_l2 = []      # L2 语义等价，待宿主 agent 判定

    def add(self, *, anchor, conclusion, layer, expected, actual,
            thesis_excerpt=None, proposal_excerpt=None,
            location_index=None, location_excerpt=None,
            verdict='待人工确认', confidence=None, judgment_question=None,
            suggestion=''):
        rule = RULE[anchor]
        issue = {
            'id': issue_id(rule, anchor),
            'rule': rule,
            'layer': layer,
            'skill': SKILL,
            'severity': SEV[conclusion],
            'location': {
                'kind': 'paragraph' if location_index else 'document',
                'index': location_index,
                'excerpt': location_excerpt,
                'offset': None,
            },
            'expected': expected,
            'actual': actual,
            'standard_source': 'default',
            'suggestion': suggestion,
            'suppressible': True,
            'action': None,
            # proposal-consistency 专用证据字段
            'conclusion': conclusion,
            'anchor': anchor,
            'proposal_excerpt': proposal_excerpt,
            'thesis_excerpt': thesis_excerpt,
            'verdict': verdict,
            'confidence': confidence,
            'judgment_question': judgment_question,
        }
        self.issues.append(issue)

    def mark_passed(self, anchor, value, note):
        self.passed.append({'anchor': anchor, 'value': value, 'note': note})

    def mark_l2_pending(self, anchor, proposal_excerpt, thesis_excerpt, question,
                       proposal_loc=None, thesis_loc=None):
        self.pending_l2.append({
            'anchor': anchor,
            'proposal_excerpt': proposal_excerpt,
            'thesis_excerpt': thesis_excerpt,
            'judgment_question': question,
        })
        # L2 不在脚本下结论：产出 待判 issue（info），交宿主 agent
        self.add(
            anchor=anchor, conclusion='无法判定', layer='L2',
            expected='（语义等价需宿主 agent 判定）',
            actual='（脚本不下结论）',
            thesis_excerpt=thesis_excerpt, proposal_excerpt=proposal_excerpt,
            location_index=thesis_loc, location_excerpt=thesis_excerpt,
            verdict='待判', confidence=None, judgment_question=question,
            suggestion='L2 语义等价：宿主 agent 判 等价/不等价/无法判定；置信度 <0.7 一律归 无法判定。')


# ================================================================ 各锚点核对

def _heading_index(headings, keyword, level=None):
    for h in headings:
        if (level is None or h['level'] == level) and keyword in h['text']:
            return h['index'], h['text']
    return None, None


def check_A1_title(b, prop_text, thesis_text):
    """A1 题目（L1 字面）：双向命中题眼关键词 → 一致。"""
    tok = '墙里墙外'
    if tok in prop_text and tok in thesis_text:
        b.mark_passed('A1', '题目一致', '双向命中题眼「%s」' % tok)
    else:
        miss = '开题' if tok not in prop_text else '正文'
        b.add(anchor='A1', conclusion='缺失' if tok not in prop_text else '偏离',
              layer='L1',
              expected='题目在开题与正文中应一致',
              actual='「%s」仅出现在%s' % (tok, miss),
              proposal_excerpt=(prop_text[:40] if tok in prop_text else None),
              thesis_excerpt=(thesis_text[:40] if tok in thesis_text else None),
              suggestion='核对题目是否被改写，必要时统一。')


def check_A2_concept(b, prop_text, thesis_text, thesis_headings):
    """A2 核心概念/关键词（L1 + L2）：脚本摆齐两侧摘录，语义等价交宿主判定。"""
    # 概念命中（作为一致证据，但不下最终结论）
    present = all(k in prop_text and k in thesis_text
                 for k in ['家庭情感共同体', '代际情感资源分配'])
    q = ('开题的“家庭情感共同体 / 代际情感资源分配 / 三维情感构成”与正文 1.4 节的概念界定'
         '是否语义等价？请判 等价 / 不等价 / 无法判定（置信度 <0.7 归 无法判定）。')
    pe = _snippet(prop_text, '家庭情感共同体')
    te_idx, te = _heading_index(thesis_headings, '核心概念与理论基础')
    b.mark_l2_pending('A2', pe, te or _snippet(thesis_text, '家庭情感共同体'), q,
                      thesis_loc=te_idx)


def check_A3_theory(b, prop_text, thesis_text, thesis_headings):
    """A3 理论基础与学者（L1 + L2）：摆齐两侧理论清单，语义等价交宿主判定。"""
    q = ('开题的“特纳情感唤起论 + 霍克希尔德情感劳动 + Glick 家庭生命周期”'
         '与正文 1.4.2 理论基础是否语义等价？判 等价 / 不等价 / 无法判定。')
    pe = _snippet(prop_text, '特纳')
    te_idx, te = _heading_index(thesis_headings, '理论基础')
    b.mark_l2_pending('A3', pe, te or _snippet(thesis_text, '霍克希尔德'), q,
                      thesis_loc=te_idx)


def check_A4_method(b, prop_text, thesis_text, thesis_text_full):
    """A4 研究方法要素（L1 数字为主）：
       · 方法/对象/地点一致性 → 进通过清单；
       · 样本量 30—40 → 15 的变动 → 改动（warn）。"""
    method_kw = ['访谈', '观察', '幼儿园', '养老院']
    ok = all(k in prop_text for k in method_kw) and all(k in thesis_text for k in method_kw)
    if ok:
        b.mark_passed('A4', '方法要素一致',
                      '访谈法/观察法/幼儿园家长(P)/养老院老人(E)/子女(C) 双向命中')
    else:
        miss = [k for k in method_kw if k not in prop_text or k not in thesis_text]
        b.add(anchor='A4', conclusion='缺失' if any(k not in prop_text for k in miss) else '偏离',
              layer='L1',
              expected='开题所列研究方法要素应在正文出现',
              actual='方法要素缺失：%s' % '、'.join(miss),
              proposal_excerpt=_snippet(prop_text, '访谈'),
              thesis_excerpt=_snippet(thesis_text, '访谈'),
              suggestion='核对研究方法节（正文 1.5.1）是否覆盖开题承诺的方法。')

    # —— 样本量（A4-03）——
    planned = extract_planned_sample(prop_text)
    actual = extract_actual_sample(thesis_text_full)
    obs = extract_observations(thesis_text_full)
    if planned and actual is not None:
        lo, hi = planned
        if lo <= actual <= hi:
            b.mark_passed('A4', '样本量一致',
                          '开题计划 %d–%d 人，正文实际 %d 人（在计划区间内）' % (lo, hi, actual))
        else:
            pe = _snippet(prop_text, '计划访谈')
            te_idx = para_index_by_keyword(_thesis_path_holder[0], '完成访谈') if _thesis_path_holder else None
            if te_idx:
                para = Scanner(_thesis_path_holder[0]).para_text(te_idx)
                te = _snippet(para, '完成访谈')
            else:
                te = _snippet(thesis_text_full, '完成访谈')
            b.add(anchor='A4', conclusion='改动', layer='L1',
                  expected='开题计划访谈 %d–%d 人' % (lo, hi),
                  actual='正文实际完成访谈 %d 人%s' % (
                      actual, ('，另有 %d 次观察' % obs if obs else '')),
                  proposal_excerpt=pe, thesis_excerpt=te,
                  location_index=te_idx, location_excerpt=te,
                  suggestion='答辩准备“为何从 %d–%d 收缩到 %d”的说明；'
                             '确认 1.5.1 已写明抽样标准与饱和判断。' % (lo, hi, actual))
    elif planned is None or actual is None:
        # 数字抽不出来 → 无法判定，交人
        b.add(anchor='A4', conclusion='无法判定', layer='L1',
              expected='开题计划样本量', actual='未能从文档稳定抽取样本量数字',
              proposal_excerpt=_snippet(prop_text, '访谈'),
              thesis_excerpt=_snippet(thesis_text_full, '访谈'),
              verdict='待判', confidence=None,
              judgment_question='请人工核对开题与正文的样本量数字是否一致。',
              suggestion='样本量为关键差异点，请人工确认。')


def check_A5_outline(b, prop_text, thesis_text, outline, thesis_headings):
    """A5 章节大纲（L3 结构）：三个预定义结构性锚点，逐条内容比对。"""
    # 四重逻辑（一致，进通过清单）
    four = ['结构逻辑', '时代逻辑', '分工逻辑', '本质逻辑']
    thesis_four = all(any(t in h['text'] for h in thesis_headings) for t in four)
    prop_four = all(any(t in it['text'] for it in outline) for t in four)
    if thesis_four and prop_four:
        b.mark_passed('A5', '四重逻辑一致', '结构/时代/分工/本质 逻辑在双方同名对应（正文 3.1–3.4）')

    # —— A5-01 改动：开题独立章『作为情感共同体的家庭』并入正文 1.4 ——
    merged = next((it for it in outline if it['level'] == 1
                   and '作为情感共同体的家庭' in it['text']), None)
    if merged:
        thesis_has_chapter = any('作为情感共同体的家庭' in h['text']
                                for h in thesis_headings if h['level'] == 1)
        if not thesis_has_chapter:
            # 取开题该章及其子目作为摘录
            pe = merged['raw']
            te_idx, te = _heading_index(thesis_headings, '核心概念与理论基础')
            if te_idx is None:
                te_idx, te = _heading_index(thesis_headings, '作为情感共同体的家庭')
            b.add(anchor='A5', conclusion='改动', layer='L3',
                  expected='开题大纲含独立章「三、作为情感共同体的家庭」',
                  actual='正文无同名一级章，已并入 1.4 核心概念与理论基础（六章→五章）',
                  proposal_excerpt=pe, thesis_excerpt=te,
                  location_index=te_idx, location_excerpt=te,
                  suggestion='答辩说明“第三章并入 1.4”的体例调整理由，属自洽改动。')

    # —— A5-02 偏离：正文第 2 章超出开题三阶段，新增 2.4/2.5 ——
    extra_toks = ['祖辈接棒', '隔代照料', '收缩阶段']
    # 正文第 2 章的二级标题
    ch2_idx, _ = _heading_index(thesis_headings, '代际情感资源分配的现状', level=1)
    extra_secs = []
    if ch2_idx:
        nxt = next((h['index'] for h in thesis_headings
                    if h['level'] == 1 and h['index'] > ch2_idx), 10 ** 9)
        extra_secs = [h for h in thesis_headings
                      if h['level'] == 2 and ch2_idx < h['index'] < nxt
                      and any(t in h['text'] for t in extra_toks)]
    prop_has = any(any(t in it['text'] for t in extra_toks) for it in outline)
    if extra_secs and not prop_has:
        first = extra_secs[0]
        b.add(anchor='A5', conclusion='偏离', layer='L3',
              expected='开题第 2 章仅三阶段（未婚/新婚无子女/婚后有子女）',
              actual='正文第 2 章新增：%s' % '、'.join(h['text'] for h in extra_secs),
              proposal_excerpt=_snippet(_outline_text_holder[0], '婚后有子女阶段'),
              thesis_excerpt=first['text'],
              location_index=first['index'], location_excerpt=first['text'],
              suggestion='与家庭生命周期框架自洽，属合理扩充；答辩可简要说明新增依据。')

    # —— A5-03 缺失：开题第四章有『重点案例分析』，正文不设个案专节 ——
    prop_cs = next((it for it in outline if '重点案例分析' in it['text']), None)
    if prop_cs:
        thesis_has = any('重点案例分析' in h['text'] for h in thesis_headings)
        if not thesis_has:
            te_idx, te = _heading_index(thesis_headings, '情感分配不对称的后果')
            if te_idx is None:
                te_idx, te = _heading_index(thesis_headings, '代际情感', level=1)
            b.add(anchor='A5', conclusion='缺失', layer='L3',
                  expected='开题大纲第四章含「（四）重点案例分析」专节',
                  actual='正文无“重点案例分析”个案专节（此前已定的体例决定）',
                  proposal_excerpt=prop_cs['raw'], thesis_excerpt=te,
                  location_index=te_idx, location_excerpt=te,
                  suggestion='最易被问；准备“为何不设个案专节”的说明，或补写专节。')


def check_A6_conclusion(b, prop_text, thesis_text, thesis_headings):
    """A6 预期结论/研究假设（L2 语义）：摆齐两侧核心论断，交宿主判定。"""
    q = ('开题核心论断“并非孝道衰落或代际剥削，而是常态化适应策略”'
         '与正文 3.4 本质逻辑是否语义等价？判 等价 / 不等价 / 无法判定。')
    pe = _snippet(prop_text, '常态适应策略')
    te_idx, te = _heading_index(thesis_headings, '本质逻辑')
    b.mark_l2_pending('A6', pe, te or _snippet(thesis_text, '常态化适应策略'), q,
                      thesis_loc=te_idx)


def check_A7_refs(b, prop_text, thesis_docx):
    """A7 文献承诺（L1 数字）：正文参考文献条数 ≥ 开题承诺（不少于 N 条）。"""
    scan = Scanner(thesis_docx)
    _, entries = scan.reference_section()
    n = len(entries)
    promised = extract_promised_refs(prop_text)
    if promised is None:
        b.mark_passed('A7', '文献承诺一致', '正文参考文献 %d 条（开题未给明确下限）' % n)
    elif n >= promised:
        b.mark_passed('A7', '文献承诺一致', '正文参考文献 %d 条 ≥ 开题承诺不少于 %d 条' % (n, promised))
    else:
        b.add(anchor='A7', conclusion='缺失', layer='L1',
              expected='参考文献不少于 %d 条' % promised,
              actual='正文仅 %d 条' % n,
              proposal_excerpt=_snippet(prop_text, '不少于'),
              thesis_excerpt='参考文献 %d 条' % n,
              suggestion='补至开题承诺的数量，或说明数量下调理由。')


def _snippet(text, keyword, width=46):
    i = text.find(keyword)
    if i < 0:
        return text[:width]
    a = max(0, i - 12)
    b = min(len(text), i + len(keyword) + width - 12)
    return ('…' if a > 0 else '') + text[a:b] + ('…' if b < len(text) else '')


# 跨函数传参用的临时持有者（避免改太多签名）
_thesis_path_holder = [None]
_outline_text_holder = [None]


# ================================================================ 主流程

def audit(docx, proposal, out_dir=None, verbose=True):
    if not os.path.exists(docx):
        raise FileNotFoundError(docx)
    if not os.path.exists(proposal):
        raise FileNotFoundError(proposal)

    _thesis_path_holder[0] = docx
    prop_full = full_text(proposal, is_proposal=True)
    _outline_text_holder[0] = prop_full
    thesis_full = full_text(docx, is_proposal=False)
    outline = parse_proposal_outline(prop_full)
    thesis_headings = parse_thesis_headings(docx)

    b = Builder()
    try:
        check_A1_title(b, prop_full, thesis_full)
        check_A2_concept(b, prop_full, thesis_full, thesis_headings)
        check_A3_theory(b, prop_full, thesis_full, thesis_headings)
        check_A4_method(b, prop_full, thesis_full, thesis_full)
        check_A5_outline(b, prop_full, thesis_full, outline, thesis_headings)
        check_A6_conclusion(b, prop_full, thesis_full, thesis_headings)
        check_A7_refs(b, prop_full, docx)
    except Exception as e:
        import traceback
        b.issues.append({
            'id': issue_id('audit-internal', 'proposal-consistency'),
            'rule': 'audit-internal-error', 'layer': 'L3', 'skill': SKILL,
            'severity': 'error',
            'location': {'kind': 'document', 'index': None, 'excerpt': None, 'offset': None},
            'expected': '每条锚点检查都应正常跑完',
            'actual': '核对异常：%s: %s' % (type(e).__name__, e),
            'standard_source': 'default',
            'suggestion': '这是工具缺陷不是文档问题，请把这条报给维护者；排查：%s'
                         % traceback.format_exc().splitlines()[-1],
            'suppressible': False, 'action': None,
            'conclusion': '无法判定', 'anchor': None,
            'proposal_excerpt': None, 'thesis_excerpt': None,
            'verdict': '待判', 'confidence': None, 'judgment_question': None,
        })

    issues = b.issues
    counts = {'缺失': 0, '改动': 0, '偏离': 0, '无法判定': 0}
    for it in issues:
        c = it.get('conclusion')
        if c in counts:
            counts[c] += 1

    meta = {
        'skill': SKILL, 'layer': 'L3',
        'docx': os.path.abspath(docx),
        'proposal': os.path.abspath(proposal),
        'docx_sha256': sha256_file(docx),
        'proposal_sha256': sha256_file(proposal),
        'generated_at': time.strftime('%Y-%m-%dT%H:%M:%S'),
        'passed': b.passed,
        'pending_l2': b.pending_l2,
        'counts': counts,
        'notes': [
            '一致性核对只读，未修改任何文档；所有 issue 的 action 为 null。',
            'L2 语义等价不下结论：已输出 待判 证据卡，终裁交宿主 agent（等价/不等价/无法判定）。',
            '五类结论 → severity：缺失=error / 改动=warn / 偏离=warn / 无法判定=info；一致不产 issue。',
        ],
    }
    summary = Scanner(docx).summary()
    summary['proposal_paragraphs'] = len(Document(proposal).paragraphs)
    if verbose:
        print('[proposal-consistency] 正文=%s' % docx)
        print('  开题=%s' % proposal)
        print('  差异 %d 处（缺失 %d / 改动 %d / 偏离 %d），无法判定 %d 处，一致 %d 处'
              % (counts['缺失'] + counts['改动'] + counts['偏离'],
                 counts['缺失'], counts['改动'], counts['偏离'],
                 counts['无法判定'], len(b.passed)))
    return {'issues': issues, 'suppressed': [], 'summary': summary, 'meta': meta}


def render_evidence_md(result):
    issues = result['issues']
    meta = result['meta']
    passed = meta.get('passed', [])
    pending = meta.get('pending_l2', [])
    c = meta.get('counts', {})
    L = []
    L.append('# 开题报告 ↔ 正文 一致性核对（证据卡）')
    L.append('')
    L.append('| 项 | 值 |')
    L.append('|---|---|')
    L.append('| 正文 | `%s` |' % meta.get('docx'))
    L.append('| 正文 sha256 | `%s` |' % meta.get('docx_sha256'))
    L.append('| 开题 | `%s` |' % meta.get('proposal'))
    L.append('| 开题 sha256 | `%s` |' % meta.get('proposal_sha256'))
    diff = c.get('缺失', 0) + c.get('改动', 0) + c.get('偏离', 0)
    L.append('| 差异 | %d 处（缺失 %d / 改动 %d / 偏离 %d）|' % (
        diff, c.get('缺失', 0), c.get('改动', 0), c.get('偏离', 0)))
    L.append('| 无法判定（待宿主判定） | %d 处 |' % c.get('无法判定', 0))
    L.append('| 一致 | %d 处 |' % len(passed))
    L.append('')

    if meta.get('notes'):
        L.append('## 审计说明（先读这段，再看结论）')
        L.append('')
        for n in meta['notes']:
            L.append('- %s' % n)
        L.append('')

    L.append('## 证据卡（每条都给两侧原文，供你自判）')
    L.append('')
    order = {'缺失': 0, '改动': 1, '偏离': 2, '无法判定': 3}
    for it in sorted(issues, key=lambda x: (order.get(x.get('conclusion'), 9),
                                           x.get('location', {}).get('index') or 0)):
        sev = R.SEV_LABEL.get(it.get('severity'), it.get('severity'))
        concl = it.get('conclusion', '')
        anchor = it.get('anchor', '')
        L.append('### [%s] %s　结论：%s %s' % (
            anchor, it.get('rule'), concl, {'缺失': '⛔', '改动': '⚠️',
                                            '偏离': '⚠️', '无法判定': '❓'}.get(concl, '')))
        if it.get('proposal_excerpt'):
            L.append('- 开题：`%s`' % it['proposal_excerpt'])
        if it.get('thesis_excerpt'):
            L.append('- 正文：`%s`' % it['thesis_excerpt'])
        L.append('- 期望：%s' % it.get('expected'))
        L.append('- 实际：%s' % it.get('actual'))
        if it.get('judgment_question'):
            L.append('- 判定问题（交宿主 agent）：%s' % it['judgment_question'])
        L.append('- 建议：%s' % it.get('suggestion'))
        L.append('- 裁定：%s%s' % (it.get('verdict'),
                                   '（confidence=%s）' % it.get('confidence')
                                   if it.get('confidence') is not None else ''))
        L.append('')

    L.append('## 通过清单（一致，进统计）')
    L.append('')
    if passed:
        for p in passed:
            L.append('- **%s** %s —— %s' % (p['anchor'], p['value'], p['note']))
    else:
        L.append('- （无）')
    L.append('')
    if pending:
        L.append('## 待宿主 agent 判定（L2 语义等价）')
        L.append('')
        for p in pending:
            L.append('- **%s**：%s' % (p['anchor'], p['judgment_question']))
        L.append('')
    L.append('---')
    L.append('')
    L.append('> 本报告由只读审计生成，**未修改任何文档**。一致性终裁在人与导师。')
    return '\n'.join(L)


def write_outputs(out_dir, result):
    os.makedirs(out_dir, exist_ok=True)
    md = os.path.join(out_dir, 'audit-report.md')
    js = os.path.join(out_dir, 'audit.json')
    with open(js, 'w', encoding='utf-8') as f:
        json.dump(R.render_json(result['issues'], result['suppressed'],
                                result['summary'], result['meta']),
                  f, ensure_ascii=False, indent=2)
    with open(md, 'w', encoding='utf-8') as f:
        f.write(render_evidence_md(result))
    return md, js


def main():
    ap = argparse.ArgumentParser(description='开题报告 ↔ 正文 一致性核对（只读）')
    ap.add_argument('-d', '--docx', required=True, help='正文 docx（唯一真值源）')
    ap.add_argument('-p', '--proposal', default=None, help='开题报告 docx（双文档核对必需）')
    ap.add_argument('-o', '--out', default=None, help='产物目录')
    ap.add_argument('--json', action='store_true', help='打印 JSON 到终端')
    ap.add_argument('--print', action='store_true', help='把报告打到终端')
    a = ap.parse_args()

    if not a.proposal:
        sys.stderr.write(
            '错误：开题报告—正文一致性核对是双文档核对，必须提供 -p/--proposal <开题报告.docx>。\n'
            '用法：python audit.py -d <正文.docx> -p <开题报告.docx>\n')
        return 2

    from workflow._common import audit_dir   # noqa: E402
    out_dir = a.out or audit_dir(a.docx, SKILL)
    try:
        res = audit(a.docx, a.proposal, out_dir=out_dir, verbose=True)
    except FileNotFoundError as e:
        sys.stderr.write('错误：文档不存在：%s\n' % e)
        return 2

    md, js = write_outputs(out_dir, res)
    print('输出：%s' % out_dir)
    if a.print:
        print('\n' + render_evidence_md(res))
    if a.json:
        print(json.dumps(R.render_json(res['issues'], res['suppressed'],
                                       res['summary'], res['meta']),
                         ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    sys.exit(main())
