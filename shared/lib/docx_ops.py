#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
shared/lib/docx_ops.py —— 文档原子写操作库（所有修复型子 skill 的唯一出口）

为什么要有这一层
    "改通道"的全部安全性都压在这里。所有子 skill 都不许自己 lxml 乱改 XML，
    必须通过本模块的 **具名动作（action）** 来改文档。这样才有三件事：
      1. 幂等 —— 每个动作自带 already_applied()，重跑一次不会叠加第二遍；
      2. 可回滚 —— 动作不改结构性地破坏信息，配合 ooxml_guard 的整档快照；
      3. 可审计 —— plan.yaml 里只出现动作名 + 参数，人能看懂、能勾选。

定位（Locator）—— 这是全库最关键的机制
    绝不按"第 123 段"这种裸序号定位。因为：
      · 你手改过一段，段落序号就整体位移了；
      · 一批动作里有删除段落的，后面的序号全错位；
      · 上一轮 plan 生成后你又编了一稿，序号完全失效。
    所以定位用 **三元组指纹**：
      para_index   —— 只是"提示"，用来在多个候选里挑最近的
      fp           —— 段落文本（去空白后）的 sha1 前 12 位，硬校验
      text         —— 段落全文，指纹对不上时用来模糊重定位
    动作级再叠加 start/end/text_expect，精确到"段内第几个 run"。

    重定位失败时 **绝不猜**：抛 LocateError，由 apply 记成 skip 并写进报告，
    让人工决定。宁可漏改，不可改错。

用法
    import sys, os
    PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    sys.path.insert(0, PKG)
    from shared.lib.docx_ops import Doc, ACTIONS, LocateError, run_action

    doc = Doc(r"D:/论文.docx")
    run_action(doc, "cite.superscript", {"locator": loc})
    doc.save_atomic()
"""
import copy
import hashlib
import os
import re
import shutil
import tempfile
from typing import Optional

try:
    from docx import Document
except ImportError:  # pragma: no cover
    raise SystemExit(
        "需要 python-docx。用那个 venv 跑：\n"
        '  "C:/Users/Tian/.workbuddy/binaries/python/envs/default/Scripts/python.exe"'
    )

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

# 与 docx_scan 保持同一套遍历语义（hyperlink / 域结果都要算进可见文本）
from shared.lib.docx_scan import CLAUSE_PUNCT, SENTENCE_END, norm  # noqa: E402


# ================================================================ 异常

class LocateError(Exception):
    """定位失败：文档状态与 plan 生成时不一致。绝不降级为"就近猜一个"。"""


class OpError(Exception):
    """动作本身不合法（参数缺失、动作名不存在）。"""


# ================================================================ 底层 XML 小工具

def _get_or_add(parent, tag: str, first: bool = False):
    """取子元素，没有就建。first=True 时插到最前（rPr 必须在 w:r 首位）。"""
    el = parent.find(tag)
    if el is not None:
        return el
    el = parent.makeelement(tag, {})
    if first:
        parent.insert(0, el)
    else:
        parent.append(el)
    return el


def _text_of(el) -> str:
    return ''.join(t.text or '' for t in el.findall(W + 't'))


def _set_text(el, s: str):
    """把一个 w:r 的文字整体替换为 s（保留 rPr，重建 w:t）。"""
    for t in el.findall(W + 't'):
        el.remove(t)
    t = el.makeelement(W + 't', {})
    if s != s.strip():
        t.set(XML_SPACE, 'preserve')
    t.text = s
    el.append(t)


def para_fingerprint(p_el) -> str:
    """段落文本指纹：去所有空白（含 U+3000）+ 小写 后取 sha1 前 12 位。

    用全角空格敏感的方式归一化 —— 论文标题常写「摘　　要」，不归一化就认不出。
    """
    t = norm(''.join(t.text or '' for t in p_el.iter(W + 't')))
    return hashlib.sha1(t.encode('utf-8')).hexdigest()[:12]


def make_locator(p_el, para_index: int, start: int = -1, end: int = -1,
                 text_expect: Optional[str] = None) -> dict:
    """给一个段落/run 造定位三元组。审计阶段统一用它，保证前后端一致。"""
    loc = {
        'para_index': para_index,
        'fp': para_fingerprint(p_el),
        'text': ''.join(t.text or '' for t in p_el.iter(W + 't')),
    }
    if start >= 0:
        loc['start'] = start
        loc['end'] = end
    if text_expect is not None:
        loc['text_expect'] = text_expect
    return loc


def fingerprint_of_text(text: str) -> str:
    return hashlib.sha1(norm(text).encode('utf-8')).hexdigest()[:12]


# ================================================================ 段落内 run 遍历
#
# 语义与 docx_scan.Scanner._walk 完全一致：hyperlink 内的 run、域结果 run 都算。
# 区别是本模块需要 **run 元素及其真实父节点**，所以另写一份返回 (el, parent, offset)。

def walk_runs(p_el):
    """返回 [(run_el, parent_el, start, end, text)]，start/end 是段内字符偏移。

    跳过 fldChar / instrText（它们不显示文字）。
    """
    out = []
    offset = 0

    def push(el, parent):
        nonlocal offset
        t = _text_of(el)
        if not t:
            return
        out.append((el, parent, offset, offset + len(t), t))
        offset += len(t)

    def handle(el, parent):
        if el.find(W + 'fldChar') is not None or el.find(W + 'instrText') is not None:
            return
        push(el, parent)

    for child in p_el:
        if child.tag == W + 'r':
            handle(child, p_el)
        elif child.tag == W + 'hyperlink':
            for r in child.findall(W + 'r'):
                handle(r, child)
        elif child.tag in (W + 'ins', W + 'smartTag', W + 'sdt', W + 'sdtContent'):
            for r in child.iter(W + 'r'):
                handle(r, child)
    return out


# ================================================================ 定位器

class Doc:
    """一层薄包装：缓存段落元素、指纹，提供定位与保存。

    刻意不继承 python-docx 的 Document —— 避免子 skill 绕过动作直接改 XML。
    """

    def __init__(self, docx_path: str):
        self.path = os.path.abspath(docx_path)
        if not os.path.exists(self.path):
            raise FileNotFoundError(self.path)
        self.doc = Document(self.path)
        self._paras = list(self.doc.paragraphs)
        self._fps = None
        self.dirty = False

    # -------------------------------------------------- 定位

    def paragraphs(self):
        return self._paras

    def refresh(self):
        """重取段落表并清空指纹缓存。

        结构一变（插段/删段/并段）就**必须**调 —— 否则 _paras 与 _fps 都是旧的，
        后续动作会拿着过期段号去定位。
        """
        self._paras = list(self.doc.paragraphs)
        self._fps = None

    def fingerprints(self):
        if self._fps is None:
            self._fps = [para_fingerprint(p._element) for p in self._paras]
        return self._fps

    def locate_para(self, loc: dict, allow_fuzzy: bool = True):
        """把 locator 解析成 (paragraph, index_1based)。解析不出来抛 LocateError。

        策略（按可信度递减）：
          1. para_index 处的指纹一致          —— 最可信
          2. 全文指纹唯一命中                  —— 段落位移了，但内容没变
          3. text 完全一致且唯一               —— 指纹算法变过，救一下
          4. 首个 30 字完全一致且唯一          —— 段落被轻度编辑过
          都不中就抛错，不做"就近取一个"。
        """
        if not isinstance(loc, dict) or 'fp' not in loc:
            raise LocateError('locator 缺 fp 字段：%r' % (loc,))

        fps = self.fingerprints()
        n = len(self._paras)
        hint = loc.get('para_index') or 0
        fp = loc.get('fp')

        # 1. 提示位直接命中
        if 1 <= hint <= n and fps[hint - 1] == fp:
            return self._paras[hint - 1], hint

        # 2. 全文指纹命中
        hits = [i for i, f in enumerate(fps) if f == fp]
        if len(hits) == 1:
            return self._paras[hits[0]], hits[0] + 1
        if len(hits) > 1:
            # 同文本段落（如多个空标题）——取离提示位最近的
            best = min(hits, key=lambda i: abs((i + 1) - hint))
            return self._paras[best], best + 1

        if not allow_fuzzy:
            raise LocateError('指纹 %s 在文档中找不到（提示位 %d）' % (fp, hint))

        # 3/4. 文本兜底
        txt = loc.get('text') or ''
        if txt:
            same = [i for i, p in enumerate(self._paras)
                    if ''.join(t.text or '' for t in p._element.iter(W + 't')) == txt]
            if len(same) == 1:
                return self._paras[same[0]], same[0] + 1
            head = norm(txt)[:30]
            if head:
                pre = [i for i, p in enumerate(self._paras)
                       if norm(''.join(t.text or '' for t in p._element.iter(W + 't'))).startswith(head)]
                if len(pre) == 1:
                    return self._paras[pre[0]], pre[0] + 1

        raise LocateError(
            '定位失败：指纹 %s / 提示位 %d / 文本「%s…」。'
            '文档很可能已被编辑过，请重新跑 audit 再生成 plan。'
            % (fp, hint, (txt or '')[:24]))

    def locate_run(self, loc: dict):
        """把 locator 解析成 (paragraph, run_el, parent_el, index_1based, start, end)。

        run 级定位在段内做两轮：先按 (start,end) 偏移命中，再按 text_expect 文本命中。
        """
        p, idx = self.locate_para(loc)
        p_el = p._element                      # locate_para 返回 Paragraph，这里要 XML 元素
        runs = walk_runs(p_el)
        if not runs:
            raise LocateError('段 %d 内没有可见 run' % idx)

        exp = loc.get('text_expect')
        s, e = loc.get('start', -1), loc.get('end', -1)

        if s is not None and e is not None and s >= 0:
            for r, par, a, b, t in runs:
                if a == s and b == e and (exp is None or t == exp):
                    return p_el, r, par, idx, a, b

        if exp:
            cand = [(r, par, a, b, t) for r, par, a, b, t in runs if t == exp]
            if len(cand) == 1:
                r, par, a, b, t = cand[0]
                return p_el, r, par, idx, a, b
            if len(cand) > 1 and s >= 0:
                r, par, a, b, t = min(cand, key=lambda c: abs(c[2] - s))
                return p_el, r, par, idx, a, b
            if not cand:
                raise LocateError('段 %d 内找不到文本为 %r 的 run' % (idx, exp))

        raise LocateError('段 %d 内 run 定位失败：start=%s text_expect=%r' % (idx, s, exp))

    # -------------------------------------------------- 保存

    def save_atomic(self, target: Optional[str] = None, backup_dir: Optional[str] = None):
        """原子保存：先写临时文件 → 再落位。返回最终路径。

        ★ 不许改成 doc.save(path) 直写 ★
        目标文件常被 WPS/预览面板占用，直写会半截损坏；os.replace 会被
        PermissionError 挡掉（Windows 上 rename 却能过），所以保留三级回退：
            正常 replace → 重试 5 次 → 把旧文件改名挪走再落位
        这段回退是血泪换来的，删了会复发"改了半天文件没变"。
        """
        import time

        tgt = os.path.abspath(target or self.path)
        d = os.path.dirname(tgt)
        if backup_dir:
            os.makedirs(backup_dir, exist_ok=True)
            cand = os.path.join(backup_dir, os.path.basename(tgt))
            if os.path.exists(cand):
                cand = os.path.join(backup_dir, 'new-' + os.path.basename(tgt))
            tmp = cand
        else:
            fd, tmp = tempfile.mkstemp(suffix='.docx', prefix='~pddoct-', dir=d)
            os.close(fd)
        self.doc.save(tmp)

        for attempt in range(5):
            try:
                os.replace(tmp, tgt)
                self.dirty = False
                self.path = tgt
                return tgt
            except PermissionError:
                time.sleep(0.6 * (attempt + 1))

        # 旧文件挪走（rename 在"被占用"时通常仍可成功）
        stamp = time.strftime('%Y%m%d-%H%M%S')
        moved = '%s.held-%s' % (tgt, stamp)
        try:
            os.rename(tgt, moved)
        except OSError as e:
            raise OpError(
                '目标文件被占用且无法挪走（%s）。请关闭 WPS/Word 与预览面板后重试。' % e)
        os.rename(tmp, tgt)
        self.dirty = False
        self.path = tgt
        return tgt


# ================================================================ 动作实现
#
# 约定：每个动作函数签名 (doc: Doc, params: dict) -> bool（True = 真的改了东西）
# 每个动作都必须可选实现 already_applied(doc, params) -> bool

def set_run_superscript(r_el, on: bool):
    rPr = _get_or_add(r_el, W + 'rPr', first=True)
    va = rPr.find(W + 'vertAlign')
    if on:
        if va is None:
            va = rPr.makeelement(W + 'vertAlign', {})
            rPr.append(va)
        va.set(W + 'val', 'superscript')
    elif va is not None:
        rPr.remove(va)


def is_superscript(r_el) -> bool:
    rPr = r_el.find(W + 'rPr')
    if rPr is None:
        return False
    va = rPr.find(W + 'vertAlign')
    return va is not None and va.get(W + 'val') == 'superscript'


def set_run_size(r_el, pt: float):
    """设置字号。sz / szCs 都是半磅整数（24 = 12pt）。"""
    half = int(round(pt * 2))
    rPr = _get_or_add(r_el, W + 'rPr', first=True)
    for tag in ('sz', 'szCs'):
        el = rPr.find(W + tag)
        if el is None:
            el = rPr.makeelement(W + tag, {})
            rPr.append(el)
        el.set(W + 'val', str(half))


def clear_run_size(r_el) -> bool:
    """抹掉 run 的显式字号，改回"随样式继承"。

    ★ 上标引注**必须**走这一步，不许 set_run_size(6.5) ★
    原因：Word/WPS 对 `w:vertAlign="superscript"` 会自动按比例缩小（约 65%）。
    如果你再显式写一个 6.5pt，渲染出来是 6.5 × 65% ≈ 4.2pt，肉眼几乎看不见 ——
    实测这份初稿的 46 处引注**全部没有显式 w:sz**，缩放完全交给渲染器。
    所以"引注字号偏小"的正确修法是**删掉显式字号**，不是改成别的值。
    """
    rPr = r_el.find(W + 'rPr')
    if rPr is None:
        return False
    hit = False
    for tag in ('sz', 'szCs'):
        el = rPr.find(W + tag)
        if el is not None:
            rPr.remove(el)
            hit = True
    return hit


def set_run_font(r_el, ascii_font=None, eastasia=None, size_pt=None, bold=None):
    rPr = _get_or_add(r_el, W + 'rPr', first=True)
    if ascii_font or eastasia:
        rf = rPr.find(W + 'rFonts')
        if rf is None:
            rf = rPr.makeelement(W + 'rFonts', {})
            rPr.insert(0, rf)
        if ascii_font:
            rf.set(W + 'ascii', ascii_font)
            rf.set(W + 'hAnsi', ascii_font)
        if eastasia:
            rf.set(W + 'eastAsia', eastasia)
    if size_pt is not None:
        set_run_size(r_el, size_pt)
    if bold is not None:
        b = rPr.find(W + 'b')
        if bold and b is None:
            rPr.append(rPr.makeelement(W + 'b', {}))
        elif not bold and b is not None:
            rPr.remove(b)


# ---------------------------------------------------------------- 超链接

def hyperlink_of_run(r_el):
    """返回包住这个 run 的 w:hyperlink（若有）。"""
    par = r_el.getparent()
    return par if par is not None and par.tag == W + 'hyperlink' else None


def unwrap_run(r_el):
    """把 run 从 w:hyperlink 里提出来，放回段落原位。返回是否发生改动。"""
    hl = hyperlink_of_run(r_el)
    if hl is None:
        return False
    grand = hl.getparent()
    at = list(grand).index(hl)
    grand.insert(at, r_el)          # lxml 会自动从 hl 里摘除
    if len(hl.findall(W + 'r')) == 0:
        grand.remove(hl)
    return True


def wrap_run_in_hyperlink(r_el, anchor: str):
    """Word 形态：<w:hyperlink w:anchor="ref15"><w:r>…</w:r></w:hyperlink>"""
    hl = hyperlink_of_run(r_el)
    if hl is not None:
        if hl.get(W + 'anchor') == anchor:
            return False
        hl.set(W + 'anchor', anchor)
        hl.set(W + 'history', '1')
        return True
    par = r_el.getparent()
    at = list(par).index(r_el)
    new = parse_xml('<w:hyperlink %s w:anchor="%s" w:history="1"/>'
                    % (nsdecls('w'), anchor))
    par.remove(r_el)
    par.insert(at, new)
    new.append(r_el)
    return True


def field_code_of_run(r_el):
    """WPS 形态：run 落在 HYPERLINK 域结果里，返回 (begin_run, instrText_el, anchor)。

    ★ 反向扫描的 depth 语义（曾经写反，导致拆不掉链接）：
        separate 是"进入结果区"的标志，**不算嵌套深度**；
        只有 end 才 depth+1（说明撞进了内层域），begin 则 depth-1 或收工。
      若把 separate 也算深度，从结果 run 往回找时永远找不到配对的 begin，
      函数直接返回 None，于是"去链接"变成空操作 —— 表面上没报错，实则没改。
    """
    prev = r_el.getprevious()
    instr, begin, seen_sep = None, None, False
    depth = 0
    while prev is not None:
        if prev.tag == W + 'r':
            fc = prev.find(W + 'fldChar')
            if fc is not None:
                t = fc.get(W + 'fldCharType')
                if t == 'separate':
                    if depth == 0:
                        seen_sep = True
                elif t == 'end':
                    depth += 1
                elif t == 'begin':
                    if depth == 0:
                        begin = prev
                        break
                    depth -= 1
            else:
                it = prev.find(W + 'instrText')
                # 只认同一层级、且已越过 separate 的那个 instrText
                if it is not None and depth == 0 and seen_sep and instr is None:
                    instr = it
        prev = prev.getprevious()
    if begin is None:
        return None, None, None
    from shared.lib.docx_scan import HYPERLINK_REF_RE
    m = HYPERLINK_REF_RE.search(instr.text or '') if instr is not None else None
    return begin, instr, (m.group(1) if m else None)


def strip_field_hyperlink(r_el):
    """拆掉包着 run 的 HYPERLINK 域（begin/instrText/separate/end 四个 run）。

    WPS 会把内部链接重写成域代码；要"去链接"就必须连域一起拆，只删 anchor
    是删不掉的。返回是否改动。
    """
    begin, instr, _ = field_code_of_run(r_el)
    if begin is None:
        return False
    # 从 begin 向前配对到 end。
    # ★★ 只摘「域自身的骨架 run」（fldChar + instrText），**结果 run 一律保留** ★★
    #    结果 run 就是引注文字本体，一起删掉等于把 [12] 从正文里抹了。
    #    只有域配平（找到配对 end）才动手，残缺域宁可不拆。
    doomed, depth, cursor, closed = [], 0, begin, False
    while cursor is not None:
        if cursor.tag == W + 'r':
            fc = cursor.find(W + 'fldChar')
            if fc is not None:
                t = fc.get(W + 'fldCharType')
                if t == 'begin':
                    depth += 1
                    doomed.append((cursor, cursor.getparent()))
                elif t == 'separate':
                    doomed.append((cursor, cursor.getparent()))
                elif t == 'end':
                    depth -= 1
                    doomed.append((cursor, cursor.getparent()))
                    if depth == 0:
                        closed = True
                        break
                cursor = cursor.getnext()
                continue
            if cursor.find(W + 'instrText') is not None:
                doomed.append((cursor, cursor.getparent()))
        cursor = cursor.getnext()
    if not closed:
        return False
    for el, par in doomed:
        if par is not None and el.getparent() is par:
            par.remove(el)
    return True


def make_field_hyperlink(cite_r_el, anchor: str):
    """WPS 形态：把 run 包进 HYPERLINK 域。某些 WPS 版本认这个不认 w:anchor。"""
    par = cite_r_el.getparent()
    at = list(par).index(cite_r_el)
    tmpl = (
        '<w:r %s><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r %s><w:instrText xml:space="preserve"> HYPERLINK \\l "%s" </w:instrText></w:r>'
        '<w:r %s><w:fldChar w:fldCharType="separate"/></w:r>'
    ) % (nsdecls('w'), nsdecls('w'), anchor, nsdecls('w'))
    end = parse_xml('<w:r %s><w:fldChar w:fldCharType="end"/></w:r>' % nsdecls('w'))
    nodes = [parse_xml(m) for m in re.findall(r'<w:r .*?</w:r>', tmpl)]
    for i, nd in enumerate(nodes):
        par.insert(at + i, nd)
    par.insert(at + len(nodes) + 1, end)   # cite run 之后收尾
    return True


# ---------------------------------------------------------------- 书签

def add_bookmark(p_el, name: str) -> bool:
    """给整段套一个书签（refN 锚点就是这样挂上去的）。已存在则不动。"""
    for b in p_el.iter(W + 'bookmarkStart'):
        if b.get(W + 'name') == name:
            return False
    used = [int(b.get(W + 'id')) for b in
            p_el.getroottree().getroot().iter(W + 'bookmarkStart')
            if (b.get(W + 'id') or '').isdigit()]
    bid = str(max(used) + 1) if used else '1'
    start = parse_xml('<w:bookmarkStart %s w:id="%s" w:name="%s"/>'
                      % (nsdecls('w'), bid, name))
    end = parse_xml('<w:bookmarkEnd %s w:id="%s"/>' % (nsdecls('w'), bid))
    pPr = p_el.find(W + 'pPr')
    at = (list(p_el).index(pPr) + 1) if pPr is not None else 0
    p_el.insert(at, start)
    p_el.append(end)
    return True


# ---------------------------------------------------------------- run 拆分 / 移动

def split_run(r_el, at: int):
    """把 run 在段内文本偏移 at 处一分为二，返回后半截 run。

    用于"引注要挪到句号之前"这类操作：句号常和正文挤在同一个 run 里，
    必须先把 run 切开才能把引注插到中间。
    rPr 原样复制，保证切开后字号/字体不变。
    """
    text = _text_of(r_el)
    if not (0 < at < len(text)):
        raise OpError('split_run 偏移越界：at=%d len=%d' % (at, len(text)))
    tail = copy.deepcopy(r_el)
    _set_text(r_el, text[:at])
    _set_text(tail, text[at:])
    r_el.addnext(tail)
    return tail


def citation_block(r_el):
    """返回包含该 run 的"引注块"—— 移动/删除**必须**以块为单位。

    三种形态（少认一种就会把文档改坏）：
      Word 内部链接  → `[w:hyperlink]`（run 包在里面）
      WPS 域代码     → `[begin, instrText, separate, [n], end]` 五个 run
      纯文字         → `[run]`
    只挪承载 `[n]` 那一个 run，会把域拆散（WPS 打开报错）或把空链接留在原地。
    """
    hl = hyperlink_of_run(r_el)
    if hl is not None:
        return [hl]
    begin, _, _ = field_code_of_run(r_el)
    if begin is None:
        return [r_el]
    block, depth, cur = [], 0, begin
    while cur is not None:
        block.append(cur)
        fc = cur.find(W + 'fldChar') if cur.tag == W + 'r' else None
        if fc is not None:
            t = fc.get(W + 'fldCharType')
            if t == 'begin':
                depth += 1
            elif t == 'end':
                depth -= 1
                if depth == 0:
                    break
        cur = cur.getnext()
    return block


def _prev_text_run(first_el):
    """从 first_el 往回找最近的、**有可见文字**的 w:r 兄弟。

    ★ 不能只看 getprevious()：WPS 域形态下，引注 run 的上一个兄弟是
      `fldChar separate`（无文字），只看一格会拿到空串然后直接放弃，
      结果"标点在后"的缺陷修不动 —— 而且**静默返回 False**，不报错。
    """
    prev = first_el.getprevious()
    while prev is not None:
        if prev.tag == W + 'r' and _text_of(prev):
            return prev
        prev = prev.getprevious()
    return None


def move_run_before_punct(r_el, punct: Optional[str] = None) -> bool:
    """把引注挪到它前面那个标点的**前面**（规范写法：句号前）。

    例：`……分配问题。[1]` → `……分配问题[1]。`

    已合规返回 False。整块移动，域/链接一并带走。
    """
    block = citation_block(r_el)
    first = block[0]
    par = first
    while par is not None and par.tag != W + 'p':
        par = par.getparent()
    if par is None:
        return False

    prev = _prev_text_run(first)
    if prev is None:
        return False
    t = _text_of(prev)

    # 判定"标点在引注之前"：紧邻引注左侧那个可见 run 以句读符结尾。
    # 若 plan 里记的标点与文档实际不符（audit 之后你又手动编辑过），
    # 只要实际末尾确实是句读符就按实际的来 —— 修对，比"因为标签不符而放弃"好。
    if punct and t.endswith(punct):
        cut = len(t) - len(punct)
    elif t[-1] in (SENTENCE_END + CLAUSE_PUNCT):
        cut = len(t) - 1
    else:
        return False

    if cut > 0:
        # 标点粘在文字尾巴上 → 切开，让标点独占一个 run
        holder = split_run(prev, cut)
    else:
        holder = prev                      # 整个 run 就是标点

    for el in block:
        if el.getparent() is not None:
            el.getparent().remove(el)
    pt = holder
    for el in block:
        pt.addprevious(el)                 # 依次插到标点之前，保持块内顺序
    return True


# ---------------------------------------------------------------- 段落级

def set_paragraph_style(p_el, style_name: str) -> bool:
    """套命名样式。Word 的样式名与 styleId 不一定同形，这里两种都试。"""
    pPr = _get_or_add(p_el, W + 'pPr', first=True)
    st = pPr.find(W + 'pStyle')
    if st is None:
        st = pPr.makeelement(W + 'pStyle', {})
        pPr.insert(0, st)
    if st.get(W + 'val') == style_name:
        return False
    st.set(W + 'val', style_name)
    return True


def set_paragraph_border(p_el, edge: str = 'bottom', val: str = 'single',
                         sz: int = 6, space: int = 1, color: str = 'auto') -> bool:
    """加段落边框。页眉那条横线就是 bottom 边框（sz 单位 1/8 磅）。"""
    pPr = _get_or_add(p_el, W + 'pPr', first=True)
    bdr = _get_or_add(pPr, W + 'pBdr')
    ex = bdr.find(W + edge)
    if ex is not None and ex.get(W + 'val') == val and ex.get(W + 'sz') == str(sz):
        return False
    if ex is not None:
        bdr.remove(ex)
    el = bdr.makeelement(W + edge, {})
    el.set(W + 'val', val)
    el.set(W + 'sz', str(sz))
    el.set(W + 'space', str(space))
    el.set(W + 'color', color)
    # pBdr 内部有固定顺序：top,left,bottom,right,between,bar —— 乱序 Word 会报错
    order = ['top', 'left', 'bottom', 'right', 'between', 'bar']
    idx = order.index(edge) if edge in order else len(order)
    for pos, child in enumerate(list(bdr)):
        cname = child.tag.split('}')[-1]
        cidx = order.index(cname) if cname in order else len(order)
        if cidx > idx:
            bdr.insert(pos, el)
            return True
    bdr.append(el)
    return True


def delete_paragraph(p_el):
    """彻底删段。只摘元素 —— 别用 p._element.getparent().remove 之外的歪招。"""
    par = p_el.getparent()
    if par is None:
        raise OpError('段落没有父节点，删不了')
    par.remove(p_el)


def merge_paragraph_into(src_p_el, dst_p_el) -> bool:
    """把 src 段并进 dst 段（参考文献去重合并时用），删掉 src。

    ★ 按**文档顺序**拼接：src 在 dst 前面就插到 dst 内容之前。
      参考文献重复条目的典型场景是"相邻两条重复"，若无脑 append，
      合并后读起来是「第 2 条的题名 + 第 1 条的作者」，明显不是想要的。
    """
    if src_p_el is dst_p_el:
        return False
    par = dst_p_el.getparent()
    if par is None:
        raise OpError('目标段落没有父节点')
    siblings = list(par)
    src_before = siblings.index(src_p_el) < siblings.index(dst_p_el)
    moved = [c for c in list(src_p_el) if c.tag != W + 'pPr']

    if src_before:
        first_content = next((c for c in list(dst_p_el) if c.tag != W + 'pPr'), None)
        for c in moved:
            if first_content is None:
                dst_p_el.append(c)
            else:
                first_content.addprevious(c)
    else:
        for c in moved:
            dst_p_el.append(c)
    delete_paragraph(src_p_el)
    return True


def insert_paragraph_after(p_el, text: str = '', style: str = ''):
    """在 p_el 之后插一段。用于拆分过长段落（结构体检用得到）。"""
    new = parse_xml('<w:p %s/>' % nsdecls('w'))
    if style:
        pPr = new.makeelement(W + 'pPr', {})
        st = pPr.makeelement(W + 'pStyle', {})
        st.set(W + 'val', style)
        pPr.append(st)
        new.append(pPr)
    if text:
        r = new.makeelement(W + 'r', {})
        t = r.makeelement(W + 't', {})
        t.set(XML_SPACE, 'preserve')
        t.text = text
        r.append(t)
        new.append(r)
    p_el.addnext(new)
    return new


# ================================================================ 动作表
#
# 名称规范：<域>.<动作>。plan.yaml 的 action 字段直接用它。

def _a_cite_superscript(doc: Doc, params: dict) -> bool:
    _, r, _, _, _, _ = doc.locate_run(params['locator'])
    want = bool(params.get('value', True))
    if is_superscript(r) == want:
        return False
    set_run_superscript(r, want)
    return True


def _a_cite_size(doc: Doc, params: dict) -> bool:
    _, r, _, _, _, _ = doc.locate_run(params['locator'])
    pt = float(params['pt'])
    rPr = r.find(W + 'rPr')
    cur = None
    if rPr is not None:
        sz = rPr.find(W + 'sz')
        if sz is not None and (sz.get(W + 'val') or '').isdigit():
            cur = int(sz.get(W + 'val')) / 2.0
    if cur is not None and abs(cur - pt) < 0.01:
        return False
    set_run_size(r, pt)
    return True


def _a_cite_link(doc: Doc, params: dict) -> bool:
    _, r, _, _, _, _ = doc.locate_run(params['locator'])
    form = params.get('form', 'hyperlink')
    anchor = params.get('anchor')
    changed = False
    if form == 'none' or not anchor:
        changed |= strip_field_hyperlink(r)
        changed |= unwrap_run(r)
        return changed
    if form == 'field':
        changed |= unwrap_run(r)
        begin, instr, cur = field_code_of_run(r)
        if cur == anchor:
            return False
        if begin is not None:
            changed |= strip_field_hyperlink(r)
        make_field_hyperlink(r, anchor)
        return True
    # hyperlink（Word 形态）
    changed |= strip_field_hyperlink(r)
    changed |= wrap_run_in_hyperlink(r, anchor)
    return changed


def _a_cite_rewrite(doc: Doc, params: dict) -> bool:
    """改写引注文字（如 `[12]`→`[12-14]`）与括号形态。"""
    _, r, _, _, _, _ = doc.locate_run(params['locator'])
    new = params['text']
    if _text_of(r) == new:
        return False
    _set_text(r, new)
    return True


def _a_cite_move_before_punct(doc: Doc, params: dict) -> bool:
    _, r, _, _, _, _ = doc.locate_run(params['locator'])
    return move_run_before_punct(r, params.get('punct', '。'))


def _a_run_clear_size(doc: Doc, params: dict) -> bool:
    _, r, _, _, _, _ = doc.locate_run(params['locator'])
    return clear_run_size(r)


def _a_run_font(doc: Doc, params: dict) -> bool:
    _, r, _, _, _, _ = doc.locate_run(params['locator'])
    set_run_font(r, params.get('ascii'), params.get('eastasia'),
                 params.get('size_pt'), params.get('bold'))
    return True


def _a_para_style(doc: Doc, params: dict) -> bool:
    p, _ = doc.locate_para(params['locator'])
    return set_paragraph_style(p._element, params['style'])


def _a_para_border(doc: Doc, params: dict) -> bool:
    p, _ = doc.locate_para(params['locator'])
    return set_paragraph_border(
        p._element, params.get('edge', 'bottom'), params.get('val', 'single'),
        int(params.get('sz', 6)), int(params.get('space', 1)),
        params.get('color', 'auto'))


def _a_para_delete(doc: Doc, params: dict) -> bool:
    p, _ = doc.locate_para(params['locator'])
    delete_paragraph(p._element)
    return True


def _a_para_merge_into(doc: Doc, params: dict) -> bool:
    src, _ = doc.locate_para(params['locator'])
    dst, _ = doc.locate_para(params['target'])
    return merge_paragraph_into(src._element, dst._element)


def _a_para_insert_after(doc: Doc, params: dict) -> bool:
    p, _ = doc.locate_para(params['locator'])
    insert_paragraph_after(p._element, params.get('text', ''), params.get('style', ''))
    return True


def _a_ref_bookmark(doc: Doc, params: dict) -> bool:
    p, _ = doc.locate_para(params['locator'])
    return add_bookmark(p._element, params['name'])


ACTIONS = {
    'cite.superscript': _a_cite_superscript,
    'cite.size': _a_cite_size,
    'cite.link': _a_cite_link,
    'cite.rewrite': _a_cite_rewrite,
    'cite.move_before_punct': _a_cite_move_before_punct,
    'run.clear_size': _a_run_clear_size,
    'run.font': _a_run_font,
    'para.style': _a_para_style,
    'para.border': _a_para_border,
    'para.delete': _a_para_delete,
    'para.merge_into': _a_para_merge_into,
    'para.insert_after': _a_para_insert_after,
    'ref.bookmark': _a_ref_bookmark,
}

# 会改变段落数量的动作 —— apply 时必须按段号**倒序**执行，且执行后要 refresh()
# 漏一个进这个集合，后面的动作就会拿着过期段号去定位（自检里栽过）
STRUCTURAL_ACTIONS = {'para.delete', 'para.merge_into', 'para.insert_after'}


def run_action(doc: Doc, action: str, params: dict) -> bool:
    fn = ACTIONS.get(action)
    if fn is None:
        raise OpError('未知动作 %r。可选：%s' % (action, ', '.join(sorted(ACTIONS))))
    ok = fn(doc, params or {})
    if ok:
        doc.dirty = True
        if action in STRUCTURAL_ACTIONS:
            doc.refresh()      # 段号已位移，必须重取段落表
    return ok
