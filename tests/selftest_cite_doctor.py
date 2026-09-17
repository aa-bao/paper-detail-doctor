#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
tests/selftest_cite_doctor.py —— cite-doctor 的**注入式**验收

为什么不能只跑一遍干净文档就收工
    "在一份本来就干净的稿子上跑出 0 条问题"完全不能说明检查器是好的 ——
    一个永远返回空列表的函数也能做到。所以这里反过来做：
    **人为造出每一种缺陷，确认检查器抓得到；再跑自动修复，确认问题消失。**

    这是本包唯一有意义的验收方式。后续每个子 skill 都要配一份这样的测试。

九个场景（都在临时副本上，不动你的原稿）
    0 基线        干净稿 → 只应剩真问题（[20] 缺页码），且不得有 error
    1 上标丢失    → cite-form
    2 链接丢失    → cite-jump（无链接）
    3 引注跑到标点后 → cite-position，且自动修复后正文变成「…问题[1]。」
    4 条目缺 refN 书签 → cite-jump（缺书签）
    5 编号断号    → cite-numbering + ref-coverage（引了表里没有的）
    6 重复条目    → ref-duplicate，且**必须**没有自动动作
    7 未被引用的条目 → ref-coverage（ref-unused）
    8 引用不存在的编号 → cite-jump（悬空）+ ref-coverage
    另外两项横切属性：问题 ID 的稳定性、自动修复的幂等性。

怎么跑
    cd F:\Coding\Project\paper-detail-doctor
    "python" tests/selftest_cite_doctor.py
"""
import copy
import importlib.util
import os
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(HERE)
sys.path.insert(0, PKG)
sys.path.insert(0, HERE)   # 让 tests/_sample.py 可被 import

from shared.lib.docx_ops import (                                    # noqa: E402
    Doc, W, _text_of, citation_block, delete_paragraph, insert_paragraph_after,
    run_action, split_run, walk_runs,
)
from shared.lib.docx_scan import Scanner, norm                       # noqa: E402
from shared.lib.ooxml_guard import apply_plan, sha256_file           # noqa: E402

from _sample import thesis as _sample_thesis, missing_hint as _missing_hint   # noqa: E402
DEFAULT_SAMPLE = _sample_thesis()
TMP = tempfile.mkdtemp(prefix='pdd-cite-selftest-')

_ok, _fail, _notes = 0, 0, []


def check(name, cond, detail=''):
    global _ok, _fail
    if cond:
        _ok += 1
        print('  [ok]   %s' % name)
    else:
        _fail += 1
        print('  [FAIL] %s   %s' % (name, detail))
        _notes.append('%s :: %s' % (name, detail))


def section(t):
    print('\n=== %s ===' % t)


# ---------------------------------------------------------------- 工具

def load_audit_mod():
    p = os.path.join(PKG, 'skills', 'cite-doctor', 'scripts', 'audit.py')
    spec = importlib.util.spec_from_file_location('cite_audit', p)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


AUD = None


def audit(docx, spec=None):
    """跑一次审计，返回 issues 列表（静默）。"""
    return AUD.audit(docx, spec, verbose=False)['issues']


def rules_of(issues):
    return {i['rule'] for i in issues}


def _jdesc(j):
    """把执行记录压成一行。用函数而不是内联表达式 —— 内联时
    `'...%s' % (a and (x, y))` 会被当成多参数格式化，翻过一次车。"""
    if not j:
        return '（没有可执行动作 → 未运行）'
    return 'changed=%d skipped=%d failed=%s' % (
        j['changed_count'], len(j['skipped']), j['failed'][:1])


def fresh(src, tag):
    p = os.path.join(TMP, tag + '.docx')
    shutil.copy2(src, p)
    return p


def apply_fixes(docx, issues, only_rules=None, dry_run=False):
    items = []
    for it in issues:
        act = it.get('action')
        if not act or not act.get('name'):
            continue
        if only_rules and it['rule'] not in only_rules:
            continue
        items.append({'issue_id': it['id'], 'rule': it['rule'],
                      'action': act['name'], 'params': act.get('params') or {},
                      'title': it.get('suggestion'), 'enabled': True})
    if not items:
        return None
    return apply_plan(docx, {'version': 1, 'docx': docx,
                             'docx_sha256': sha256_file(docx),
                             'items': items}, dry_run=dry_run)


def cite_locator(docx, idx=0, prefer=None):
    """返回第 idx 处引注的 (Doc, locator, Citation)。prefer: 'before-punct' 等过滤。"""
    scan = Scanner(docx)
    cites = scan.citations()
    if prefer:
        cites = [c for c in cites if c.position == prefer] or cites
    c = cites[idx]
    d = Doc(docx)
    from shared.lib.docx_ops import make_locator
    loc = make_locator(d.paragraphs()[c.para_index - 1]._element, c.para_index,
                       c.start, c.end, c.text)
    return d, loc, c


def para_el_of(docx, para_index):
    return Doc(docx).paragraphs()[para_index - 1]._element


# ---------------------------------------------------------------- 注入器

def inject_move_cite_after_punct(docx, idx=3):
    """把一处引注连同它的域块整体挪到句读符之后。

    注意必须挪"整个域块"（begin/instrText/separate/[n]/end 五个 run），
    只挪承载 [n] 的那个 run 会把域拆散，WPS 打开就报错 —— 那样测的就不是
    "位置违规"，而是"文档损坏"，两回事。
    """
    d, loc, c = cite_locator(docx, idx, prefer='before-punct')
    p_el = d.paragraphs()[c.para_index - 1]._element
    runs = walk_runs(p_el)
    r = next(x[0] for x in runs if x[2] == c.start and x[3] == c.end)
    block = citation_block(r)

    # 标点在哪：找到覆盖字符下标 c.end 的那个可见 run（引注之后的第一个字符）
    holder, h_start = None, None
    for el, _par, a, b, _t in runs:
        if a <= c.end < b:
            holder, h_start = el, a
            break
    if holder is None:
        raise RuntimeError('找不到引注之后的字符所在 run')

    # 把标点切成独占一个 run，这样"引注整块挪到标点之后"才干净
    off = c.end - h_start
    if off > 0:
        holder = split_run(holder, off)          # 返回值 = 从标点开始的后半段
    if len(_text_of(holder)) > 1:
        split_run(holder, 1)

    for el in block:
        el.getparent().remove(el)
    pt = holder
    for el in block:
        pt.addnext(el)                           # addnext 依次插，顺序不倒
        pt = el
    d.save_atomic()
    return c.para_index


def inject_drop_bookmark(docx, name='ref7'):
    d = Doc(docx)
    body = d.doc.element.body
    bid = None
    for b in list(body.iter(W + 'bookmarkStart')):
        if b.get(W + 'name') == name:
            bid = b.get(W + 'id')
            b.getparent().remove(b)
            break
    if bid is not None:
        for e in list(body.iter(W + 'bookmarkEnd')):
            if e.get(W + 'id') == bid:
                e.getparent().remove(e)
    d.save_atomic()
    return name


def inject_delete_entry(docx, number=20):
    h, entries = Scanner(docx).reference_section()
    e = next(x for x in entries if x.number == number)
    d = Doc(docx)
    from shared.lib.docx_ops import make_locator
    loc = make_locator(d.paragraphs()[e.para_index - 1]._element, e.para_index)
    p, _ = d.locate_para(loc)
    delete_paragraph(p._element)
    d.save_atomic()
    return e.text


def inject_duplicate_entry(docx, number=1):
    h, entries = Scanner(docx).reference_section()
    e = next(x for x in entries if x.number == number)
    d = Doc(docx)
    p_el = d.paragraphs()[e.para_index - 1]._element
    p_el.addnext(copy.deepcopy(p_el))
    d.save_atomic()
    return e.text


def inject_append_entry(docx, text='[47] 测试文献. 测试题名[M]. 北京: 测试出版社, 2020.'):
    h, entries = Scanner(docx).reference_section()
    last = entries[-1]
    d = Doc(docx)
    insert_paragraph_after(d.paragraphs()[last.para_index - 1]._element, text)
    d.save_atomic()
    return text


def inject_rewrite_cite(docx, idx=0, new_text='[99]'):
    """把一处引注的编号改成不存在的 [99]，并去掉链接。

    第二个动作的 locator 必须用**文本**匹配：改写后 run 的字符长度变了，
    沿用原来的 start/end 偏移会定位失败（LocateError）—— 这也顺带验证了
    "locator 在文本变化后能否靠 text_expect 兜底"。
    """
    d, loc, c = cite_locator(docx, idx)
    run_action(d, 'cite.rewrite', {'locator': loc, 'text': new_text})
    loc2 = {'para_index': loc['para_index'], 'fp': loc['fp'],
            'start': loc['start'], 'end': loc['start'], 'text_expect': new_text}
    run_action(d, 'cite.link', {'locator': loc2, 'form': 'none'})
    d.save_atomic()
    return c.para_index


# ---------------------------------------------------------------- 场景

def main():
    src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SAMPLE
    global AUD
    AUD = load_audit_mod()
    if not src or not os.path.exists(src):
        print('样本不存在：%s' % src)
        print(_missing_hint())
        return 2
    print('样本：%s\n临时目录：%s' % (src, TMP))

    # ---------------------------------------------------- 0 基线
    section('0 基线：干净稿上不该冒出 error')
    base = fresh(src, 'base')
    iss = audit(base)
    check('基线无 error', not [i for i in iss if i['severity'] == 'error'],
          'errors=%s' % [(i['rule'], i['actual'][:40]) for i in iss if i['severity'] == 'error'])
    check('基线只剩已知的真问题（ref-format 缺页码）',
          rules_of(iss) <= {'ref-format'}, 'rules=%s' % rules_of(iss))
    check('真问题定位到 [20] 缺页码',
          any('页码' in (i['expected'] + i['suggestion']) for i in iss),
          'issues=%d' % len(iss))
    h = sha256_file(base)
    iss2 = audit(base)
    check('审计是只读的（哈希不变）', sha256_file(base) == h)
    check('问题 ID 稳定（重跑同 ID）',
          [i['id'] for i in iss] == [i['id'] for i in iss2],
          '%s vs %s' % ([i['id'] for i in iss][:2], [i['id'] for i in iss2][:2]))

    # ---------------------------------------------------- 1 上标丢失
    section('1 注入：引注掉上标 → cite-form')
    f = fresh(src, 'form')
    d, loc, c = cite_locator(f, 5)
    run_action(d, 'cite.superscript', {'locator': loc, 'value': False})
    d.save_atomic()
    iss = audit(f)
    hit = [i for i in iss if i['rule'] == 'cite-form']
    check('检出 cite-form error', bool(hit) and hit[0]['severity'] == 'error',
          'rules=%s' % rules_of(iss))
    check('描述里点明"上标"', bool(hit) and '上标' in (hit[0]['expected'] + hit[0]['actual']))
    j = apply_fixes(f, iss, only_rules={'cite-form'})
    check('自动修复执行成功', j and j['changed_count'] >= 1,
          'journal=%s' % _jdesc(j))
    check('修复后 cite-form 消失', 'cite-form' not in rules_of(audit(f)))
    j2 = apply_fixes(f, audit(f), only_rules={'cite-form'})
    check('幂等：再跑一次不改任何东西', j2 is None or j2['changed_count'] == 0,
          'changed=%s' % (j2 and j2['changed_count']))

    # ---------------------------------------------------- 2 链接丢失
    section('2 注入：引注掉链接 → cite-jump（无链接）')
    f = fresh(src, 'nolink')
    d, loc, c = cite_locator(f, 4)
    run_action(d, 'cite.link', {'locator': loc, 'form': 'none'})
    d.save_atomic()
    iss = audit(f)
    hit = [i for i in iss if i['rule'] == 'cite-jump']
    check('检出 cite-jump', bool(hit), 'rules=%s' % rules_of(iss))
    check('判定为不可抑制（结构性硬伤）',
          bool(hit) and hit[0]['suppressible'] is False)
    check('带自动修复动作', bool(hit) and (hit[0]['action'] or {}).get('name') == 'cite.link')
    j = apply_fixes(f, iss, only_rules={'cite-jump'})
    check('自动补链接成功', j and j['changed_count'] >= 1, 'journal=%s' % (j and j['failed'][:1]))
    check('修复后 cite-jump 消失', 'cite-jump' not in rules_of(audit(f)))
    s = Scanner(f)
    cc = [x for x in s.citations() if x.para_index == c.para_index and x.refs == c.refs]
    check('链接形态与文档原有形态一致',
          bool(cc) and cc[0].anchor and cc[0].link_kind == 'field',
          'anchor=%s kind=%s' % (cc[0].anchor if cc else None,
                                 cc[0].link_kind if cc else None))

    # ---------------------------------------------------- 3 位置违规
    section('3 注入：引注跑到标点之后 → cite-position')
    f = fresh(src, 'pos')
    pi = inject_move_cite_after_punct(f, 3)
    iss = audit(f)
    hit = [i for i in iss if i['rule'] == 'cite-position']
    check('检出 cite-position error', bool(hit) and hit[0]['severity'] == 'error',
          'rules=%s' % rules_of(iss))
    check('动作是移到句读符之前',
          bool(hit) and (hit[0]['action'] or {}).get('name') == 'cite.move_before_punct')
    j = apply_fixes(f, iss, only_rules={'cite-position'})
    check('自动挪位成功', j and j['changed_count'] >= 1, 'journal=%s' % (j and j['failed'][:1]))
    check('修复后 cite-position 消失', 'cite-position' not in rules_of(audit(f)))
    txt = Scanner(f).para_text(pi)
    check('正文顺序恢复为「…[n]。」',
          '[99]' not in txt and txt.count('。') >= 1,
          'para=%s' % txt[:60])

    # ---------------------------------------------------- 4 条目缺书签
    section('4 注入：文献条目缺 refN 书签 → cite-jump（缺书签）')
    f = fresh(src, 'nobm')
    inject_drop_bookmark(f, 'ref7')
    iss = audit(f)
    hit = [i for i in iss if i['rule'] == 'cite-jump']
    check('检出缺书签', any('书签' in i['expected'] + i['actual'] for i in hit),
          'rules=%s' % rules_of(iss))
    check('动作是补书签',
          any((i['action'] or {}).get('name') == 'ref.bookmark' for i in hit))
    check('不可抑制', all(i['suppressible'] is False for i in hit))
    j = apply_fixes(f, iss, only_rules={'cite-jump'})
    check('自动补书签成功', j and j['changed_count'] >= 1, 'journal=%s' % (j and j['failed'][:1]))
    check('ref7 书签已恢复', 'ref7' in set(Scanner(f).bookmarks()))
    check('修复后 cite-jump 消失', 'cite-jump' not in rules_of(audit(f)))

    # ---------------------------------------------------- 5 断号 + 覆盖
    section('5 注入：删掉文献 [20] → cite-numbering + ref-coverage')
    f = fresh(src, 'gap')
    inject_delete_entry(f, 20)
    iss = audit(f)
    check('检出编号断号', any(i['rule'] == 'cite-numbering' for i in iss),
          'rules=%s' % rules_of(iss))
    check('检出"引了表里没有的"', any(i['rule'] == 'ref-coverage' for i in iss))
    nc = [i for i in iss if i['rule'] == 'cite-numbering']
    check('断号不给自动修复（重编号是全局手术）',
          all(not (i.get('action') or {}).get('name') for i in nc))
    check('断号条目标为不可抑制', all(i['suppressible'] is False for i in nc))

    # ---------------------------------------------------- 6 重复条目
    section('6 注入：重复条目 → ref-duplicate')
    f = fresh(src, 'dup')
    inject_duplicate_entry(f, 1)
    iss = audit(f)
    hit = [i for i in iss if i['rule'] == 'ref-duplicate']
    check('检出重复条目', bool(hit), 'rules=%s' % rules_of(iss))
    check('重复条目**不**给自动删除动作（删了要连带重编号）',
          bool(hit) and not (hit[0].get('action') or {}).get('name'))

    # ---------------------------------------------------- 7 未被引用
    section('7 注入：新增一条没人引的文献 → ref-coverage（unused）')
    f = fresh(src, 'unused')
    inject_append_entry(f)
    iss = audit(f)
    cov = [i for i in iss if i['rule'] == 'ref-coverage']
    check('检出未被引用的条目', any('从未被' in i['actual'] for i in cov),
          'actuals=%s' % [i['actual'][:30] for i in cov])

    # ---------------------------------------------------- 8 悬空引用
    section('8 注入：正文引用 [99]（表里没有）→ cite-jump + ref-coverage')
    f = fresh(src, 'dangling')
    inject_rewrite_cite(f, 0, '[99]')
    iss = audit(f)
    check('检出悬空引用（cite-jump）',
          any(i['rule'] == 'cite-jump' for i in iss), 'rules=%s' % rules_of(iss))
    cov = [i for i in iss if i['rule'] == 'ref-coverage']
    check('检出"正文引用但表中缺失"',
          any('但表里没有' in i['actual'] or '正文引用了' in i['actual'] for i in cov),
          'actuals=%s' % [i['actual'][:40] for i in cov])
    check('缺失条目不可抑制',
          all(i['suppressible'] is False for i in cov if '正文引用了' in i['actual']))

    # ---------------------------------------------------- 收尾
    print('\n' + '=' * 62)
    print('通过 %d 项，失败 %d 项' % (_ok, _fail))
    if _notes:
        print('失败明细：')
        for n in _notes:
            print('  - ' + n)
    print('临时副本留在：%s' % TMP)
    return 0 if _fail == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
