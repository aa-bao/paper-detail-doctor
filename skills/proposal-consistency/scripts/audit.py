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

锚点档案（为什么必须外置）
    七类锚点 A1–A7 的**检查逻辑是通用的**；但"题眼是哪个词、点了哪些核心概念、
    引了哪些理论家、方法要素有哪些、大纲里哪一章被并了节"——这些是**某一篇论文
    的数据**，不是这个工具的知识。所以它们放在锚点档案 YAML 里
    （模板 assets/anchors.example.yaml），由使用者或宿主 agent 填。
    写死在代码里就等于只对一篇论文有效。
    未配置的锚点**一律产出「无法判定」交人工**，绝不静默判成一致。

开题报告的特殊性
    开题报告正文常常整个装在一个大表格里，按普通段落读几乎读不到东西。
    脚本自己遍历 w:tbl → w:tr → w:tc → w:p → w:t 把表格文本抽出来，
    因此对任意"开题报告也是个 docx"的用户都可复用。

用法
    python audit.py -d <正文.docx> -p <开题报告.docx> \\
        [-o <产物目录>] [--anchors <锚点档案.yaml>] [--json] [--print]
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

def extract_planned_sample(text, anchor=None):
    """开题计划样本量：锚定『计划访谈…人』，支持 '30-40人' 与 '15人' 两种写法。
    anchor 可由锚点档案覆盖（不同课题的措辞不同）。返回 (lo, hi)；单值则 lo==hi；找不到返回 None。"""
    if anchor:
        m = re.search(re.escape(anchor) + r'\s*(\d+)\s*(?:[-–—]\s*(\d+))?\s*人', text)
        if m:
            a = int(m.group(1))
            return (a, int(m.group(2)) if m.group(2) else a)
    m = re.search(r'计划访谈\s*(\d+)\s*(?:[-–—]\s*(\d+))?\s*人', text)
    if not m:
        m = re.search(r'拟访谈\s*(\d+)\s*(?:[-–—]\s*(\d+))?\s*人', text)
    if not m:
        return None
    a = int(m.group(1))
    b = int(m.group(2)) if m.group(2) else a
    return (a, b)


def extract_actual_sample(text, anchor=None):
    """正文实际完成样本量：锚定『完成访谈…人』（anchor 可由锚点档案覆盖）。返回整数或 None。"""
    if anchor:
        m = re.search(re.escape(anchor) + r'\s*(\d+)\s*人', text)
        if m:
            return int(m.group(1))
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

    def mark_unconfigured(self, anchor, what):
        """锚点未配置 → 报「无法判定」交人工。

        **绝不静默跳过** —— 跳过会被读成"这一项没问题"，正是本包最警惕的静默失效
        （"检查器没跑"伪装成"全部合格"）。未配置 ≠ 一致。
        """
        self.add(anchor=anchor, conclusion='无法判定', layer='L1',
                 expected='该锚点应在锚点档案（anchors.yaml）中配置',
                 actual=what,
                 verdict='待判', confidence=None,
                 judgment_question='请人工核对这一项；或按 assets/anchors.example.yaml '
                                   '补配置后重跑。',
                 suggestion='未配置 ≠ 一致。补上 %s 的锚点配置后重跑。' % anchor)


# ================================================================ 锚点档案

class Anchors:
    r"""锚点档案：把"从开题里抽哪些锚点、回正文哪里找"从代码里搬出来。

    本 skill 的**检查逻辑是通用的**（七类锚点 A1–A7 × 三层判定），但每篇论文的
    具体锚点值不同：题眼是什么、点了哪些核心概念、引了哪些理论家、方法要素有
    哪些、大纲里哪一章被并了节……这些**属于那篇论文的数据，不属于这个工具**。
    所以它们放在 YAML 里（见 assets/anchors.example.yaml），而不是写死在代码里。

    解析顺序（先命中者胜）：
      1. 命令行 --anchors <file>
      2. 环境变量 PDD_ANCHORS
      3. <正文 docx 同目录>/.thesis-doctor/anchors.yaml
      4. <包根>/assets/anchors.local.yaml

    一个都没命中时 self.data 为空 —— 各锚点会走"未配置"分支，**一律产出
    「无法判定」并说明原因**，交人工/宿主 agent，**绝不静默判成一致**
    （否则就是"检查器坏了伪装成全部合格"）。
    """

    SECTIONS = ('A1_topic', 'A2_concept', 'A3_theory', 'A4_method',
                'A5_outline', 'A6_conclusion', 'A7_refs')

    def __init__(self, path=None, docx=None):
        self.data = {}
        self.source = None
        cands = []
        if path:
            cands.append(path)
        if os.environ.get('PDD_ANCHORS'):
            cands.append(os.environ['PDD_ANCHORS'])
        if docx:
            cands.append(os.path.join(os.path.dirname(os.path.abspath(docx)),
                                      '.thesis-doctor', 'anchors.yaml'))
        cands.append(os.path.join(PKG, 'assets', 'anchors.local.yaml'))
        for c in cands:
            if c and os.path.exists(c):
                try:
                    import yaml
                    with open(c, encoding='utf-8') as f:
                        d = yaml.safe_load(f) or {}
                    if isinstance(d, dict):
                        self.data = d
                        self.source = os.path.abspath(c)
                        break
                except Exception:
                    continue

    # ---------------------------------------------------------------- 取用
    def sec(self, key):
        v = self.data.get(key)
        return v if isinstance(v, dict) else {}

    def toks(self, key):
        v = self.sec(key).get('tokens')
        if isinstance(v, str):
            v = [v]
        return [str(t).strip() for t in (v or []) if str(t).strip()]

    def val(self, key, field, default=''):
        v = self.sec(key).get(field)
        return v if isinstance(v, str) and v.strip() else default

    def lst(self, key, field):
        v = self.sec(key).get(field)
        if isinstance(v, str):
            v = [v]
        return [str(t).strip() for t in (v or []) if str(t).strip()]

    def configured(self, key):
        return bool(self.sec(key))

    def note(self):
        if self.source:
            return '锚点档案：%s' % self.source
        return ('未找到锚点档案（anchors.yaml）—— 仅 A7 文献承诺与样本量数字可自动判定，'
                '其余锚点一律报「无法判定」交人工。配置方法见 assets/anchors.example.yaml。')

    # -------------------------------------------------- A1 题眼（无配置时自动猜）
    def a1_tokens(self, prop_text):
        t = self.toks('A1_topic')
        if t:
            return t, 'configured'
        m = re.search(r'《([^》]{4,60})》', prop_text)
        if m:
            return [m.group(1).strip()], 'auto'
        return [], 'none'


# ================================================================ 各锚点核对

def _heading_index(headings, keyword, level=None):
    for h in headings:
        if (level is None or h['level'] == level) and keyword in h['text']:
            return h['index'], h['text']
    return None, None


def check_A1_title(b, prop_text, thesis_text, A):
    """A1 题目（L1 字面）：双向命中题眼关键词 → 一致。

    题眼是"这篇论文独有"的数据，必须外置到锚点档案 —— 写死在代码里就只对
    一篇论文有效。未配置时退化为"从开题里找《…》"，再找不到就报无法判定。
    """
    toks, how = A.a1_tokens(prop_text)
    if not toks:
        b.add(anchor='A1', conclusion='无法判定', layer='L1',
              expected='题目在开题与正文中应一致',
              actual='锚点档案未配置 A1_topic.tokens，且开题里找不到《…》形式的题目',
              proposal_excerpt=prop_text[:60], thesis_excerpt=None,
              verdict='待判', confidence=None,
              judgment_question='请人工核对开题与正文的题目是否一致，或补 A1_topic.tokens。',
              suggestion='在锚点档案中补 A1_topic.tokens（题眼关键词，2–6 字的独特词）。')
        return
    hit = [t for t in toks if t in prop_text and t in thesis_text]
    if len(hit) == len(toks):
        b.mark_passed('A1', '题目一致',
                      '双向命中题眼「%s」%s'
                      % ('、'.join(toks), '（自动取自开题《…》）' if how == 'auto' else ''))
    else:
        lost_prop = [t for t in toks if t not in prop_text]
        lost = lost_prop or [t for t in toks if t not in thesis_text]
        b.add(anchor='A1', conclusion='缺失' if lost_prop else '偏离',
              layer='L1',
              expected='题目在开题与正文中应一致',
              actual='题眼「%s」未在%s出现' % ('、'.join(lost), '开题' if lost_prop else '正文'),
              proposal_excerpt=_snippet(prop_text, toks[0]),
              thesis_excerpt=_snippet(thesis_text, toks[0]),
              suggestion='核对题目是否被改写，必要时统一。')


def check_A2_concept(b, prop_text, thesis_text, thesis_headings, A):
    """A2 核心概念/关键词（L1 + L2）：脚本摆齐两侧摘录，语义等价交宿主判定。"""
    toks = A.toks('A2_concept')
    head = A.val('A2_concept', 'thesis_heading')
    if not toks and not head:
        b.add(anchor='A2', conclusion='无法判定', layer='L2',
              expected='开题点名的核心概念应在正文有对应界定',
              actual='锚点档案未配置 A2_concept，抽不出概念锚点',
              proposal_excerpt=prop_text[:60], thesis_excerpt=None,
              verdict='待判', confidence=None,
              judgment_question='请人工核对开题的核心概念在正文哪一节界定、是否语义等价。',
              suggestion='在锚点档案中补 A2_concept.tokens 与 thesis_heading。')
        return
    both = [t for t in toks if t in prop_text and t in thesis_text]
    anchor_tok = (both or toks or [''])[0]
    q = ('开题的核心概念%s与正文「%s」一节的概念界定是否语义等价？'
         '请判 等价 / 不等价 / 无法判定（置信度 <0.7 归 无法判定）。'
         % ('（%s）' % ' / '.join(toks) if toks else '', head or '未指定'))
    te_idx, te = _heading_index(thesis_headings, head) if head else (None, None)
    b.mark_l2_pending('A2', _snippet(prop_text, anchor_tok),
                      te or _snippet(thesis_text, anchor_tok), q, thesis_loc=te_idx)


def check_A3_theory(b, prop_text, thesis_text, thesis_headings, A):
    """A3 理论基础与学者（L1 + L2）：摆齐两侧理论清单，语义等价交宿主判定。"""
    toks = A.toks('A3_theory')
    head = A.val('A3_theory', 'thesis_heading')
    if not toks and not head:
        b.add(anchor='A3', conclusion='无法判定', layer='L2',
              expected='开题所列理论基础应在正文有对应论述',
              actual='锚点档案未配置 A3_theory，抽不出理论锚点',
              proposal_excerpt=prop_text[:60], thesis_excerpt=None,
              verdict='待判', confidence=None,
              judgment_question='请人工核对开题的理论基础与正文对应章节是否语义等价。',
              suggestion='在锚点档案中补 A3_theory.tokens（理论/学者名）与 thesis_heading。')
        return
    both = [t for t in toks if t in prop_text and t in thesis_text]
    anchor_tok = (both or toks or [''])[0]
    q = ('开题的理论基础%s与正文「%s」是否语义等价？'
         '判 等价 / 不等价 / 无法判定（置信度 <0.7 归 无法判定）。'
         % ('（%s）' % ' + '.join(toks) if toks else '', head or '未指定'))
    te_idx, te = _heading_index(thesis_headings, head) if head else (None, None)
    b.mark_l2_pending('A3', _snippet(prop_text, anchor_tok),
                      te or _snippet(thesis_text, anchor_tok), q, thesis_loc=te_idx)


def check_A4_method(b, prop_text, thesis_text, thesis_text_full, A):
    """A4 研究方法要素（L1 数字为主）：
       · 方法/对象/地点一致性 → 进通过清单（锚词来自锚点档案）；
       · 样本量「开题计划 N–M 人 → 正文实际 K 人」→ 改动（warn）。
    未配置 A4_method 时**不跳过、也不判一致**，而是报「无法判定」交人工。"""
    method_kw = A.toks('A4_method')
    if method_kw:
        ok = all(k in prop_text for k in method_kw) and all(k in thesis_text for k in method_kw)
        if ok:
            b.mark_passed('A4', '方法要素一致', '方法要素双向命中：%s' % '、'.join(method_kw))
        else:
            miss = [k for k in method_kw if k not in prop_text or k not in thesis_text]
            b.add(anchor='A4', conclusion='缺失' if any(k not in prop_text for k in miss) else '偏离',
                  layer='L1',
                  expected='开题所列研究方法要素应在正文出现',
                  actual='方法要素缺失：%s' % '、'.join(miss),
                  proposal_excerpt=_snippet(prop_text, method_kw[0]),
                  thesis_excerpt=_snippet(thesis_text, method_kw[0]),
                  suggestion='核对研究方法节是否覆盖开题承诺的方法。')
    else:
        b.mark_unconfigured('A4', '方法要素一致性未检查：锚点档案未配置 A4_method.tokens')

    plan_word = A.val('A4_method', 'plan_anchor') or None
    actual_word = A.val('A4_method', 'actual_anchor') or None
    pe_word = plan_word or '计划访谈'
    te_word = actual_word or '完成访谈'

    # —— 样本量（A4 数字）——
    planned = extract_planned_sample(prop_text, plan_word)
    actual = extract_actual_sample(thesis_text_full, actual_word)
    obs = extract_observations(thesis_text_full)
    if planned and actual is not None:
        lo, hi = planned
        if lo <= actual <= hi:
            b.mark_passed('A4', '样本量一致',
                          '开题计划 %d–%d 人，正文实际 %d 人（在计划区间内）' % (lo, hi, actual))
        else:
            pe = _snippet(prop_text, pe_word)
            te_idx = para_index_by_keyword(_thesis_path_holder[0], te_word) if _thesis_path_holder else None
            if te_idx:
                para = Scanner(_thesis_path_holder[0]).para_text(te_idx)
                te = _snippet(para, te_word)
            else:
                te = _snippet(thesis_text_full, te_word)
            b.add(anchor='A4', conclusion='改动', layer='L1',
                  expected='开题计划访谈 %d–%d 人' % (lo, hi),
                  actual='正文实际完成访谈 %d 人%s' % (
                      actual, ('，另有 %d 次观察' % obs if obs else '')),
                  proposal_excerpt=pe, thesis_excerpt=te,
                  location_index=te_idx, location_excerpt=te,
                  suggestion='答辩准备“为何从 %d–%d 收缩到 %d”的说明；'
                             '确认研究方法节已写明抽样标准与饱和判断。' % (lo, hi, actual))
    elif planned is None or actual is None:
        # 数字抽不出来 → 无法判定，交人（不许当成"一致"）
        b.add(anchor='A4', conclusion='无法判定', layer='L1',
              expected='开题计划样本量',
              actual='未能稳定抽取样本量数字（开题锚词「%s」/ 正文锚词「%s」）' % (pe_word, te_word),
              proposal_excerpt=_snippet(prop_text, pe_word),
              thesis_excerpt=_snippet(thesis_text_full, te_word),
              verdict='待判', confidence=None,
              judgment_question='请人工核对开题与正文的样本量数字是否一致。'
                                '若措辞不同，请在锚点档案里改 A4_method.plan_anchor / actual_anchor。',
              suggestion='样本量为关键差异点，请人工确认。')


def check_A5_outline(b, prop_text, thesis_text, outline, thesis_headings, A):
    """A5 章节大纲（L3 结构）：结构性锚点逐条比对。

    三类结构性差异各自需要锚点档案说明"看哪里"（见 assets/anchors.example.yaml）：
      · 并章（改动）  merged_chapter 在开题是一级章、正文无同名一级章
      · 新增小节（偏离）extra_tokens 出现在正文 chapter_heading 章下，但开题大纲没有
      · 大纲条目缺失（缺失）missing_section 在开题大纲里、正文无对应标题
    一项都没配置时报「无法判定」—— 未配置不等于一致。
    """
    n_prop = len([it for it in outline if it.get('level') == 1])
    n_thesis = len([h for h in thesis_headings if h.get('level') == 1])

    # —— 双方应同名出现的结构性标签（一致证据）——
    tags = A.lst('A5_outline', 'expect_passed')
    if tags:
        thesis_ok = all(any(t in h['text'] for h in thesis_headings) for t in tags)
        prop_ok = all(any(t in it['text'] for it in outline) for t in tags)
        if thesis_ok and prop_ok:
            label = A.val('A5_outline', 'expect_passed_label') or '大纲同名标签一致'
            b.mark_passed('A5', label,
                          '「%s」在开题与正文同名对应' % '、'.join(tags))

    handled = 0

    # —— 改动：开题独立章在正文被并成节 ——
    merged_ch = A.val('A5_outline', 'merged_chapter')
    if merged_ch:
        handled += 1
        m = next((it for it in outline
                  if it.get('level') == 1 and merged_ch in it['text']), None)
        is_chapter_in_thesis = any(merged_ch in h['text']
                                   for h in thesis_headings if h.get('level') == 1)
        if m and not is_chapter_in_thesis:
            target = A.val('A5_outline', 'merged_target')
            te_idx, te = (_heading_index(thesis_headings, target) if target
                          else (None, None))
            if te_idx is None:
                te_idx, te = _heading_index(thesis_headings, merged_ch)
            b.add(anchor='A5', conclusion='改动', layer='L3',
                  expected='开题大纲有独立章「%s」（开题一级章共 %d 个）'
                           % (merged_ch, n_prop),
                  actual='正文无同名一级章，已并入%s（正文一级章共 %d 个）'
                         % (('「%s」' % target) if target else '正文某一节', n_thesis),
                  proposal_excerpt=m.get('raw'), thesis_excerpt=te,
                  location_index=te_idx, location_excerpt=te,
                  suggestion='答辩说明“该章并入”的体例调整理由，属自洽改动。')

    # —— 偏离：正文新增小节，开题大纲里没有 ——
    extra_toks = A.lst('A5_outline', 'extra_tokens')
    ch_head = A.val('A5_outline', 'chapter_heading')
    if extra_toks and ch_head:
        handled += 1
        ch_idx, _ = _heading_index(thesis_headings, ch_head, level=1)
        extra_secs = []
        if ch_idx:
            nxt = next((h['index'] for h in thesis_headings
                        if h.get('level') == 1 and h['index'] > ch_idx), 10 ** 9)
            extra_secs = [h for h in thesis_headings
                          if h.get('level') == 2 and ch_idx < h['index'] < nxt
                          and any(t in h['text'] for t in extra_toks)]
        prop_has = any(any(t in it['text'] for t in extra_toks) for it in outline)
        if extra_secs and not prop_has:
            first = extra_secs[0]
            ot = extra_toks[0]
            b.add(anchor='A5', conclusion='偏离', layer='L3',
                  expected='开题大纲「%s」章下不含这些小节' % ch_head,
                  actual='正文「%s」章新增：%s'
                         % (ch_head, '、'.join(h['text'] for h in extra_secs)),
                  proposal_excerpt=(_snippet(_outline_text_holder[0], ot)
                                    if ot in _outline_text_holder[0] else None),
                  thesis_excerpt=first['text'],
                  location_index=first['index'], location_excerpt=first['text'],
                  suggestion='与正文框架自洽的合理扩充；答辩可简要说明新增依据。')

    # —— 缺失：开题大纲条目在正文无对应标题 ——
    miss_sec = A.val('A5_outline', 'missing_section')
    if miss_sec:
        handled += 1
        item = next((it for it in outline if miss_sec in it['text']), None)
        if item and not any(miss_sec in h['text'] for h in thesis_headings):
            te_idx, te = None, None
            for fb in A.lst('A5_outline', 'missing_fallback'):
                te_idx, te = _heading_index(thesis_headings, fb)
                if te_idx:
                    break
            b.add(anchor='A5', conclusion='缺失', layer='L3',
                  expected='开题大纲含「%s」' % miss_sec,
                  actual='正文无同名标题（该分析环节未落地）',
                  proposal_excerpt=item.get('raw'), thesis_excerpt=te,
                  location_index=te_idx, location_excerpt=te,
                  suggestion='最易被问；准备“为何不设该环节”的说明，或补写。')

    if handled == 0:
        b.mark_unconfigured('A5', '大纲结构性差异未检查：锚点档案未配置 A5_outline')


def check_A6_conclusion(b, prop_text, thesis_text, thesis_headings, A):
    """A6 预期结论/研究假设（L2 语义）：摆齐两侧核心论断，交宿主判定。"""
    tok = A.val('A6_conclusion', 'proposal_token')
    head = A.val('A6_conclusion', 'thesis_heading')
    if not tok and not head:
        b.add(anchor='A6', conclusion='无法判定', layer='L2',
              expected='开题的核心论断应在正文结论处得到呼应',
              actual='锚点档案未配置 A6_conclusion，抽不出结论锚点',
              proposal_excerpt=prop_text[:60], thesis_excerpt=None,
              verdict='待判', confidence=None,
              judgment_question='请人工核对开题的核心论断与正文结论章是否语义等价。',
              suggestion='在锚点档案中补 A6_conclusion.proposal_token 与 thesis_heading。')
        return
    q = ('开题的核心论断（锚词「%s」处）与正文「%s」是否语义等价？'
         '判 等价 / 不等价 / 无法判定（置信度 <0.7 归 无法判定）。'
         % (tok or '未指定', head or '未指定'))
    pe = _snippet(prop_text, tok) if tok else prop_text[:60]
    te_idx, te = _heading_index(thesis_headings, head) if head else (None, None)
    b.mark_l2_pending('A6', pe, te or (tok and _snippet(thesis_text, tok)) or thesis_text[:60],
                      q, thesis_loc=te_idx)


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

def audit(docx, proposal, out_dir=None, verbose=True, anchors_path=None):
    if not os.path.exists(docx):
        raise FileNotFoundError(docx)
    if not os.path.exists(proposal):
        raise FileNotFoundError(proposal)

    A = Anchors(path=anchors_path, docx=docx)
    _thesis_path_holder[0] = docx
    prop_full = full_text(proposal, is_proposal=True)
    _outline_text_holder[0] = prop_full
    thesis_full = full_text(docx, is_proposal=False)
    outline = parse_proposal_outline(prop_full)
    thesis_headings = parse_thesis_headings(docx)

    b = Builder()
    try:
        check_A1_title(b, prop_full, thesis_full, A)
        check_A2_concept(b, prop_full, thesis_full, thesis_headings, A)
        check_A3_theory(b, prop_full, thesis_full, thesis_headings, A)
        check_A4_method(b, prop_full, thesis_full, thesis_full, A)
        check_A5_outline(b, prop_full, thesis_full, outline, thesis_headings, A)
        check_A6_conclusion(b, prop_full, thesis_full, thesis_headings, A)
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
        'anchors': A.source or None,
        'notes': [
            '一致性核对只读，未修改任何文档；所有 issue 的 action 为 null。',
            A.note(),
            'L2 语义等价不下结论：已输出 待判 证据卡，终裁交宿主 agent（等价/不等价/无法判定）。',
            '五类结论 → severity：缺失=error / 改动=warn / 偏离=warn / 无法判定=info；一致不产 issue。',
            '锚点未配置一律产出「无法判定」交人工 —— 静默跳过会被误读成"这一项没问题"。',
        ],
    }
    summary = Scanner(docx).summary()
    summary['proposal_paragraphs'] = len(Document(proposal).paragraphs)
    if verbose:
        print('[proposal-consistency] 正文=%s' % docx)
        print('  开题=%s' % proposal)
        print('  锚点档案=%s' % (A.source or '（未找到 → 各锚点将报「无法判定」交人工）'))
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
    ap.add_argument('--anchors', default=None,
                    help='锚点档案 YAML：每篇论文不同的题眼/概念/理论/方法要素/大纲锚点。'
                         '不给则依次找 环境变量 PDD_ANCHORS、'
                         '<正文目录>/.thesis-doctor/anchors.yaml、assets/anchors.local.yaml；'
                         '都没有则各锚点报「无法判定」交人工（模板见 assets/anchors.example.yaml）')
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
        res = audit(a.docx, a.proposal, out_dir=out_dir, verbose=True,
                    anchors_path=a.anchors)
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
