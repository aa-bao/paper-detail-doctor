#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
extract_conventions.py —— 从「已排好版的范文 / 样张 docx」反推体例规则（conventions）

为什么需要它
    paper-detail-doctor 包里其余 5 个审计型子 skill 都要回答「你这样排对不对」，
    而「对不对」必须有基准。基准不该是工具的偏好，应该来自学校给的模板/范文。
    本脚本把范文的**实际排版**翻译成机器可读的六组规则，并进 spec.json 的
    `conventions` 块。

    没有它，检查只能退化成「通用学术惯例」，报告里会标注 `标准来源: 默认`。

六组规则
    citation        引注：形态（上标/行内、方括号/圆括号/无）、位置（句读前/后）、
                    上标字号比、多引用写法
    bibliography    ★ 编码制（顺序编码制 / 著者-出版年制）、著录标准、编号格式
    header          页眉：文字内容、是否带横线、线型与粗细
    page_numbering  分节页码方案（封面无 / 摘要罗马 / 正文阿拉伯）与起始页
    abstract        中英文摘要字数上限、关键词个数与分隔符
    caption         图题/表题位置与编号格式

用法
    python extract_conventions.py -d 范文.docx                          # 打印 JSON
    python extract_conventions.py -d 范文.docx -o conventions.json      # 写文件
    python extract_conventions.py -d 范文.docx --merge-into spec.json   # 并进 spec 顶层
    python extract_conventions.py -d 范文.docx --strict                 # 有 low 就非零退出

约定
    - 推断不出来的标 `confidence: low` 并写进 `open_questions`，**绝不猜得像真的**。
    - `citation.style` 与 `bibliography.system` 是下游 L2 引注检查的命门，
      这两项 low 时必须向用户确认。
"""
import argparse
import collections
import json
import os
import re
import sys

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:
    sys.stderr.write("需要 python-docx：\n"
                     '  "<python 虚拟环境>'
                     '/Scripts/python.exe" -m pip install python-docx\n')
    raise

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# 一个「引注」的文本形态：[1] / ［1］ / (1) / （1） / 1 / [1,2] / [1-3]
CITE_RE = re.compile(r'^\s*([\[［【（(])?(\d+(?:\s*[-–—~,，、]\s*\d+)*)([\]］】）)])?\s*$')
# 句末标点：引注若出现在它们**之前**，是「句读前」体例；若紧跟其后，是「句读后」体例
SENTENCE_END = '。！？!?'
# 句读符（含分句/停顿）：引注落在它们之前，同样属于「句读前」
CLAUSE_PUNCT = '；，、：,;:）)】」”'
BRACKET_OPEN = '[［【（('
BRACKET_CLOSE = ']］】）)'


def _norm(s):
    """归一化标题文本：去掉所有空白（含全角空格 \\u3000）并小写。

    必须做 —— 范文里的标题常写成「摘　　要」「目　　录」「致　　谢」，
    不归一化就会漏掉它们（曾因此把「致谢」正文误当成参考文献条目）。
    """
    return re.sub(r'[\s\u3000]+', '', s or '').lower()

REF_HEADINGS = {'参考文献', '参考文献：', '参考文献:', 'references', 'bibliography'}
ACK_HEADINGS = {'致谢', '致谢：', 'acknowledgement', 'acknowledgments'}
ABS_HEADINGS = {'摘要', '中文摘要', '摘  要', 'abstract', '英文摘要'}
KW_PREFIXES = ('关键词', '关键字', 'keywords', 'key words')

CAP_FIG = re.compile(r'^\s*图\s*\d+\s*[-－—.．]\s*\d+')
CAP_TAB = re.compile(r'^\s*表\s*\d+\s*[-－—.．]\s*\d+')


# ---------------------------------------------------------------- 基础工具

def _iter_paragraphs(doc):
    """按文档流顺序产出 ('p', paragraph)。跳过表格内部段落（它们不是正文体例证据）。"""
    for p in doc.paragraphs:
        yield p


def _run_size_pt(run):
    """run 的字号（磅）。取 rPr/sz，看不到就返回 None（继承自样式）。"""
    rPr = run._element.find(W + 'rPr')
    if rPr is None:
        return None
    sz = rPr.find(W + 'sz')
    if sz is None:
        return None
    try:
        return int(sz.get(W + 'val')) / 2.0
    except (TypeError, ValueError):
        return None


def _run_is_superscript(run):
    rPr = run._element.find(W + 'rPr')
    if rPr is None:
        return False
    va = rPr.find(W + 'vertAlign')
    return va is not None and va.get(W + 'val') == 'superscript'


def _classify_cite_text(text):
    """返回 (form, bracket)：form ∈ superscript-number / bracket-inline / paren-inline …"""
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
    multi = 'merged' if re.search(r'[-–—~]', body) else (
        'separate' if re.search(r'[,，、]', body) else 'single')
    return bracket, multi


# ---------------------------------------------------------------- 1. citation

def extract_citation(doc):
    """扫正文 run，判定引注形态与位置。"""
    forms = collections.Counter()
    positions = collections.Counter()
    ratios = []
    multi = collections.Counter()
    evidence = []
    total = 0

    for pi, p in enumerate(_iter_paragraphs(doc)):
        if not p.text.strip():
            continue
        # 累积 run 文本边界，用于定位「引注后面那一个字符」
        runs = list(p.runs)
        offset = 0
        spans = []
        for r in runs:
            t = r.text or ''
            if t:
                spans.append((offset, offset + len(t), r))
            offset += len(t)
        ptext = p.text

        for (s, e, r) in spans:
            tok = (r.text or '').strip()
            if not tok or not re.search(r'\d', tok):
                continue
            bracket, m = _classify_cite_text(tok)
            if bracket is None:
                continue
            # 兜底：非上标、又没括号、又只是一串数字 —— 大概率是正文里的普通数字，不算引注
            if bracket == 'none' and not _run_is_superscript(r):
                continue

            total += 1
            sup = _run_is_superscript(r)
            if sup:
                forms['superscript-%s' % ('bracket' if bracket == 'square' else
                                          ('paren' if bracket == 'paren' else 'number'))] += 1
            else:
                forms['inline-%s' % ('bracket' if bracket == 'square' else 'paren')] += 1
            multi[m] += 1

            # 位置上：看引注紧邻的字符
            #   后一个字符是句读符     → 引注落在标点**之前** → 句读前（规范写法）
            #   前一个字符是句末标点   → 引注落在标点**之后** → 句读后
            #   两边都是正文           → 夹在正文中间（如「杨善华[1]指出」），另计
            nxt = ptext[e:e + 1]
            prev = ptext[max(0, s - 1):s]
            if nxt == '':
                positions['end-of-paragraph'] += 1
            elif nxt in SENTENCE_END or nxt in CLAUSE_PUNCT:
                positions['before-punct'] += 1
                if len(evidence) < 3:
                    evidence.append('段%d：…%s%s ← 引注在标点前（句读前）' % (pi + 1, tok, nxt))
            elif prev in SENTENCE_END:
                positions['after-punct'] += 1
                if len(evidence) < 3:
                    evidence.append('段%d：…%s%s ← 引注在句末标点后（句读后）' % (pi + 1, prev, tok))
            else:
                positions['mid-sentence'] += 1

            # 上标字号比
            if sup:
                rs = _run_size_pt(r)
                # 段落里最常见的字号当作正文号
                sizes = [x for x in (_run_size_pt(x) for x in runs) if x]
                if rs and sizes:
                    base = collections.Counter(sizes).most_common(1)[0][0]
                    if base:
                        ratios.append(rs / base)

    conf = 'high' if total >= 10 else ('medium' if total >= 3 else 'low')
    style = forms.most_common(1)[0][0] if forms else None
    pos = positions.most_common(1)[0][0] if positions else None

    if ratios:
        ratio = round(sum(ratios) / len(ratios), 2)
        ratio_note = '由 run 级 w:sz 与正文字号计算'
    elif total and any(k.startswith('superscript') for k in forms):
        ratio = None
        ratio_note = ('上标未显式设置 w:sz —— 字号由 w:vertAlign="superscript" 自动缩放'
                      '（Word/WPS 默认约为正文的 65%），无需也不应另行指定')
        evidence.append(ratio_note)
    else:
        ratio, ratio_note = None, '无上标引注'

    return {
        'style': style,
        'position': pos if pos else None,
        'superscript_size_ratio': ratio,
        'superscript_size_ratio_note': ratio_note,
        'multiple': multi.most_common(1)[0][0] if multi else None,
        'cited_count': total,
        'form_tally': dict(forms),
        'position_tally': dict(positions),
        'confidence': conf,
        'evidence': evidence or ['未在正文中识别到引注'],
    }


# ---------------------------------------------------------------- 2. bibliography

def _find_heading_index(paras, names):
    """按归一化后的文本找标题段落。names 传原始写法即可（内部会归一化）。"""
    ns = {_norm(x) for x in names}
    for i, p in enumerate(paras):
        if _norm(p.text) in ns:
            return i
    return -1


def extract_bibliography(doc):
    paras = [p for p in _iter_paragraphs(doc)]
    idx = _find_heading_index(paras, REF_HEADINGS)
    if idx < 0:
        return {'system': None, 'standard': None, 'number_format': None,
                'entry_count': 0, 'confidence': 'low',
                'evidence': ['未找到「参考文献」标题段落']}

    entries = []
    stopped_at = None
    for p in paras[idx + 1:]:
        t = p.text.strip()
        if not t:
            continue
        # 遇到「致谢」「附录」等后置章节就停 —— 否则会把致谢正文当成条目
        if _norm(t) in {_norm(x) for x in ACK_HEADINGS} or _norm(t).startswith('附录'):
            stopped_at = t
            break
        # 长段落且开头没有 [n] 编号 —— 几乎不可能是条目（防误收正文）
        # 有编号的一律保留：英文文献作者多，长到 200+ 字符是正常的（曾因阈值 250 误杀 1 条）
        if len(t) > 250 and not re.match(r'^\s*[\[［]\s*\d+\s*[\]］]', t):
            continue
        entries.append(t)

    seq = sum(1 for e in entries if re.match(r'^\s*[\[［]\s*\d+\s*[\]］]', e))
    ay = sum(1 for e in entries
             if re.search(r'[（(]\s*(19|20)\d{2}\s*[）)]', e) and not re.match(r'^\s*[\[［]', e))
    gbt = sum(1 for e in entries if re.search(r'\[[MJDCANRGSPZ]\]', e))

    if seq == 0 and ay == 0:
        system, conf = None, 'low'
    elif seq >= ay:
        system = 'sequence'
        conf = 'high' if seq >= 5 and ay == 0 else 'medium'
    else:
        system = 'author-year'
        conf = 'high' if ay >= 5 and seq == 0 else 'medium'

    if gbt >= max(1, len(entries) // 2):
        standard = 'GB/T 7714'
    else:
        standard = None

    return {
        'system': system,
        'standard': standard,
        'number_format': '[n]' if seq else None,
        'entry_count': len(entries),
        'stopped_at': stopped_at,
        'tally': {'sequence_like': seq, 'author_year_like': ay, 'has_type_marker': gbt},
        'confidence': conf,
        'evidence': [e[:60] + ('…' if len(e) > 60 else '') for e in entries[:2]],
    }


# ---------------------------------------------------------------- 3. header

def extract_header(doc):
    # 全文纯文本（含表格/文本框，封面题目常在文本框里，不在 doc.paragraphs 中）
    try:
        blob_raw = ''.join(doc.element.body.itertext())
    except Exception:
        blob_raw = '\n'.join(p.text for p in doc.paragraphs)
    blob = _norm(blob_raw)

    rules, texts, ev = collections.Counter(), [], []
    rule_detail = None
    for si, section in enumerate(doc.sections):
        try:
            hdr = section.header
        except Exception:
            continue
        for p in hdr.paragraphs:
            t = p.text.strip()
            if t:
                texts.append(t)
            pPr = p._element.find(W + 'pPr')
            bdr = pPr.find(W + 'pBdr') if pPr is not None else None
            bottom = bdr.find(W + 'bottom') if bdr is not None else None
            if bottom is not None and bottom.get(W + 'val') not in (None, 'none', 'nil'):
                rules['has-rule'] += 1
                if rule_detail is None:
                    sz = bottom.get(W + 'sz')
                    rule_detail = {
                        'style': bottom.get(W + 'val'),
                        'size_pt': (int(sz) / 8.0) if sz and sz.isdigit() else None,
                        'space_pt': (int(bottom.get(W + 'space')) / 20.0
                                     if bottom.get(W + 'space', '').isdigit() else None),
                    }
                    ev.append('第%d节页眉段落挂有 pBdr/bottom（%s）' % (si + 1, rule_detail))
            elif t:
                rules['no-rule'] += 1

    text = texts[0] if texts else ''
    found_in_body = False
    if not text:
        content, content_conf = 'none', 'high'
        ev.append('页眉为空')
    else:
        # 判据：页眉文字是否在文档正文里出现过（页眉写「论文题目」是常见体例）。
        # 注意：正文里可能**根本没有封面**（题目只写在页眉里），此时无法判定，
        # 必须老实报 low 并转人工 —— 不要猜成 paper-title。
        probe = _norm(text)[:10]
        found_in_body = bool(probe) and probe in blob
        if found_in_body:
            content, content_conf = 'paper-title', 'high'
            ev.append('页眉文字「%s」在正文中出现，判定为论文题目' % text)
        else:
            content, content_conf = 'custom', 'low'
            ev.append('页眉文字「%s」**未在正文中找到对应**，无法判定是论文题目、校名还是其他' % text)
            ev.append('该文档正文里可能没有封面（题目只写在页眉里），须向用户确认页眉该写什么')

    if rules:
        conf_rule = 'high' if (rules['has-rule'] + rules['no-rule']) >= 2 else 'medium'
        rule_flag = rules['has-rule'] >= rules['no-rule']
    else:
        conf_rule, rule_flag = 'low', None

    conf = 'low' if 'low' in (conf_rule, content_conf) else (
        'medium' if 'medium' in (conf_rule, content_conf) else 'high')

    return {
        'content': content,
        'content_confidence': content_conf,
        'text': text or None,
        'found_in_body': found_in_body,
        'rule': rule_flag,
        'rule_detail': rule_detail,
        'confidence': conf,
        'evidence': ev,
    }


# ---------------------------------------------------------------- 4. page_numbering

def extract_page_numbering(doc):
    sections, ev = [], []
    fmt_map = {'decimal': 'arabic', 'upperRoman': 'upperRoman', 'lowerRoman': 'lowerRoman',
               'upperLetter': 'upperLetter', 'lowerLetter': 'lowerLetter', 'none': 'none'}
    for si, section in enumerate(doc.sections):
        sectPr = section._sectPr
        pg = sectPr.find(W + 'pgNumType') if sectPr is not None else None
        fmt = pg.get(W + 'fmt') if pg is not None else None
        start = pg.get(W + 'start') if pg is not None else None
        sections.append({
            'index': si + 1,
            'format': fmt_map.get(fmt, fmt),
            'start': int(start) if start and start.isdigit() else None,
        })
        ev.append('第%d节 pgNumType: fmt=%s start=%s' % (si + 1, fmt or '(未设置)', start or '-'))

    fmts = [s['format'] for s in sections]
    if not any(fmts):
        scheme, conf = 'unknown', 'low'
        ev.append('各节均未显式设置页码格式 —— 页码可能是手打的，或全局用默认阿拉伯数字')
    elif len({x for x in fmts if x}) > 1:
        scheme, conf = 'multi-section', 'high'
        ev.append('各节页码格式不同（如前置部分罗马、正文阿拉伯）')
    else:
        scheme, conf = 'single', 'high'
    ev.append('共 %d 节 —— 节数即分节页码方案的分段数' % len(sections))

    return {'scheme': scheme, 'sections': sections, 'confidence': conf, 'evidence': ev}


# ---------------------------------------------------------------- 5. abstract

def extract_abstract(doc):
    paras = [p for p in _iter_paragraphs(doc)]
    out = {}
    ev = []
    for tag, heads, kwp in (('cn', ABS_HEADINGS, ('关键词', '关键字')),
                            ('en', {'abstract', '英文摘要'}, ('keywords', 'key words'))):
        i = _find_heading_index(paras, heads)
        if i < 0:
            out[tag] = {'found': False, 'confidence': 'low'}
            continue
        body, j = [], i + 1
        while j < len(paras):
            t = paras[j].text.strip()
            nt = _norm(t)
            if nt.startswith(kwp):
                break
            if nt and nt in {_norm(x) for x in REF_HEADINGS}:
                break
            body.append(t)
            j += 1
        blob = ''.join(body)
        if tag == 'cn':
            n = len(re.findall(r'[\u4e00-\u9fff]', blob))
            n_kw, sep = 0, None
            if j < len(paras):
                kt = paras[j].text.strip()
                kt = re.sub(r'^(关键词|关键字)\s*[:：]?', '', kt)
                parts = [x for x in re.split(r'[；;，,、]', kt) if x.strip()]
                n_kw, sep = len(parts), ('；' if '；' in kt else ('，' if '，' in kt else ';'))
            out[tag] = {'found': True, 'chars': n, 'keywords_count': n_kw or None,
                        'keywords_separator': sep, 'confidence': 'high' if n > 50 else 'medium'}
            ev.append('中文摘要 %d 字，关键词 %s 个' % (n, n_kw or '?'))
        else:
            n = len(re.findall(r"[A-Za-z][A-Za-z'-]*", blob))
            out[tag] = {'found': True, 'words': n,
                        'confidence': 'high' if n > 30 else 'medium'}
            ev.append('英文摘要约 %d 词' % n)

    # 是否单独成页：docx 里只有三种可判据的分页标记，都没有就必须看渲染结果
    own_page, own_basis = None, None
    for i, p in enumerate(paras):
        if _norm(p.text) in {_norm(x) for x in ABS_HEADINGS}:
            pPr = p._element.find(W + 'pPr')
            if pPr is not None and pPr.find(W + 'pageBreakBefore') is not None:
                own_page, own_basis = True, '摘要标题段落设了 pageBreakBefore'
            elif i > 0 and '<w:br w:type="page"' in (paras[i - 1]._element.xml or ''):
                own_page, own_basis = True, '摘要前一段落末尾有分页符'
            elif i > 0 and '<w:sectPr' in (paras[i - 1]._element.xml or ''):
                own_page, own_basis = True, '摘要前一段落是节末（分节符自带分页）'
            else:
                own_basis = ('未见任何分页标记；摘要是否单独成页要看渲染结果'
                             '（导出 PDF 后确认），docx 层面判不出来')
            if own_basis:
                ev.append('摘要单独成页：%s' % own_basis)
            break
    out['own_page'] = own_page
    out['own_page_basis'] = own_basis
    out['confidence'] = 'high' if own_page else 'medium'
    out['evidence'] = ev
    return out


# ---------------------------------------------------------------- 6. caption

def extract_caption(doc):
    paras = list(_iter_paragraphs(doc))
    fig_fmt = tab_fmt = None
    fig_pos = tab_pos = None
    ev = []

    for i, p in enumerate(paras):
        t = p.text.strip()
        if CAP_FIG.match(t):
            m = re.match(r'^\s*图\s*(\d+)\s*([-－—.．])\s*(\d+)', t)
            if m and not fig_fmt:
                fig_fmt = '图 {chapter}%s{index}' % ('-' if m.group(2) == '-' else '.')
            if fig_pos is None:
                # 图题若在含图片的段落之后，则题在下
                prev_img = any('<w:drawing' in (paras[j]._element.xml or '')
                               for j in range(max(0, i - 2), i))
                if prev_img:
                    fig_pos = 'below'
                    ev.append('段%d 图题紧跟在含图片段落之后 → 题在下' % (i + 1))
        if CAP_TAB.match(t):
            m = re.match(r'^\s*表\s*(\d+)\s*([-－—.．])\s*(\d+)', t)
            if m and not tab_fmt:
                tab_fmt = '表 {chapter}%s{index}' % ('-' if m.group(2) == '-' else '.')
            if tab_pos is None:
                # 表格在 body 元素流里，不在 paragraphs 里 —— 按 body 顺序判断题注相对位置
                body = list(doc.element.body)
                try:
                    pi = body.index(p._element)
                except ValueError:
                    pi = -1
                if pi >= 0:
                    for k in range(pi + 1, min(pi + 3, len(body))):
                        if body[k].tag == W + 'tbl':
                            tab_pos = 'above'
                            ev.append('段%d 表题后面紧跟表格 → 题在上' % (i + 1))
                            break
                    if tab_pos is None:
                        for k in range(max(0, pi - 2), pi):
                            if body[k].tag == W + 'tbl':
                                tab_pos = 'below'
                                ev.append('段%d 表题前面是表格 → 题在下' % (i + 1))
                                break

    return {
        'figure': {'position': fig_pos, 'number_format': fig_fmt,
                   'confidence': 'medium' if fig_pos else 'low'},
        'table': {'position': tab_pos, 'number_format': tab_fmt,
                  'confidence': 'medium' if tab_pos else 'low'},
        'confidence': 'medium' if (fig_pos or tab_pos) else 'low',
        'evidence': ev or ['未识别到图题/表题段落'],
    }


# ---------------------------------------------------------------- 汇总

def build_conventions(doc, source):
    conv = {
        'citation': extract_citation(doc),
        'bibliography': extract_bibliography(doc),
        'header': extract_header(doc),
        'page_numbering': extract_page_numbering(doc),
        'abstract': extract_abstract(doc),
        'caption': extract_caption(doc),
    }
    lows = []
    for k, v in conv.items():
        if isinstance(v, dict) and v.get('confidence') == 'low':
            lows.append(k)
        if k == 'abstract':
            for sub in ('cn', 'en'):
                if isinstance(v.get(sub), dict) and v[sub].get('confidence') == 'low' \
                        and v[sub].get('found'):
                    lows.append('abstract.%s' % sub)

    # 命门项单独强调
    must = []
    if conv['citation'].get('style') is None:
        must.append('citation.style（引注是上标还是行内、有没有方括号）')
    if conv['bibliography'].get('system') is None:
        must.append('bibliography.system（顺序编码制还是著者-出版年制）')

    # 页眉文字判定不出来时单独问 —— 它直接影响 L1 页眉检查，且无法用默认值兜底
    extra = []
    if conv['header'].get('content_confidence') == 'low' and conv['header'].get('text'):
        extra.append('header.content：页眉文字「%s」未在正文中找到对应，'
                     '请确认页眉该写论文题目还是学校名称' % conv['header']['text'])

    return {
        'meta': {'source': source, 'confidence_legend': {'high': '直接读到的真值',
                                                         'medium': '推断，建议复核',
                                                         'low': '未推出来，须人工确认'}},
        'conventions': conv,
        'open_questions': (
            ['%s：未能从范文推断出可靠取值，须向用户确认' % m for m in must]
            + extra
            + ['%s：置信度 low，建议复核' % x for x in lows if x not in
               ('citation', 'bibliography', 'header')]
        ),
        'overall_confidence': 'low' if (must or extra) else ('medium' if lows else 'high'),
    }


def main():
    ap = argparse.ArgumentParser(description='从范文 docx 反推体例规则（conventions）')
    ap.add_argument('-d', '--docx', required=True, help='范文 / 样张 docx（已排好版的）')
    ap.add_argument('-o', '--out', help='输出 JSON 路径（默认打印到 stdout）')
    ap.add_argument('--merge-into', help='把 conventions 块并进指定 spec.json 的顶层')
    ap.add_argument('--strict', action='store_true', help='有命门项缺失时非零退出')
    a = ap.parse_args()

    doc = Document(a.docx)
    result = build_conventions(doc, os.path.abspath(a.docx))
    payload = result['conventions']
    full = json.dumps(result, ensure_ascii=False, indent=2)

    if a.out:
        # 写完整报告（含 meta / open_questions / overall_confidence），便于人工复核
        with open(a.out, 'w', encoding='utf-8') as f:
            f.write(full)
        print('已写出 %s' % a.out)
    if a.merge_into:
        with open(a.merge_into, encoding='utf-8') as f:
            spec = json.load(f)
        spec['conventions'] = payload
        if result['open_questions']:
            spec.setdefault('open_questions', [])
            for q in result['open_questions']:
                if q not in spec['open_questions']:
                    spec['open_questions'].append(q)
        with open(a.merge_into, 'w', encoding='utf-8') as f:
            json.dump(spec, f, ensure_ascii=False, indent=2)
        print('已并入 %s 的 conventions 块' % a.merge_into)
    if not a.out and not a.merge_into:
        print(full)

    print('\n--- 小结 ---', file=sys.stderr)
    print('引注形态: %s   位置: %s   共识别 %d 处'
          % (payload['citation'].get('style'), payload['citation'].get('position'),
             payload['citation'].get('cited_count') or 0), file=sys.stderr)
    print('编码制:   %s   条目 %s   著录标准: %s'
          % (payload['bibliography'].get('system'),
             payload['bibliography'].get('entry_count'),
             payload['bibliography'].get('standard')), file=sys.stderr)
    print('页眉:     content=%s rule=%s'
          % (payload['header'].get('content'), payload['header'].get('rule')), file=sys.stderr)
    print('页码方案: %s' % payload['page_numbering'].get('scheme'), file=sys.stderr)
    print('总体置信度: %s' % result['overall_confidence'], file=sys.stderr)
    for q in result['open_questions']:
        print('  ! %s' % q, file=sys.stderr)

    if a.strict and result['open_questions']:
        sys.exit(2)


if __name__ == '__main__':
    main()
