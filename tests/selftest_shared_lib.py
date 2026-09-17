#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
tests/selftest_shared_lib.py —— 共享层自检（docx_scan / docx_ops / ooxml_guard）

为什么要有这个
    改 docx 的东西没法靠"看起来对"来验收。XML 写错一个标签，Word 打不开、
    WPS 静默丢格式，肉眼看不出来。所以共享层每加一个动作，都要在这里跑一遍：
      · 动作真的落盘了吗（改完重新打开文档验证）
      · 幂等吗（同一个动作连跑两次，第二次必须返回 False）
      · 定位失效时会怎样（必须抛 LocateError，不许就近乱改）
      · 账本生效吗（同一份文档第二批必须全跳过）

怎么跑（**在文档副本上跑，不动你的原稿**）
    cd F:\Coding\Project\paper-detail-doctor
    "C:/Users/Tian/.workbuddy/binaries/python/envs/default/Scripts/python.exe" tests/selftest_shared_lib.py
    # 想换样本：... tests/selftest_shared_lib.py "D:/其他.docx"
"""
import os
import shutil
import sys
import tempfile
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(HERE)
sys.path.insert(0, PKG)

from shared.lib.docx_scan import Scanner, norm                     # noqa: E402
from shared.lib.docx_ops import (                                  # noqa: E402
    ACTIONS, Doc, LocateError, make_locator, run_action, set_run_size,
)
from shared.lib.ooxml_guard import (                               # noqa: E402
    SnapshotManager, apply_plan, is_locked, sha256_file,
)

DEFAULT_SAMPLE = r'F:\Coding\work\0813-城市家庭\初稿.docx'

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


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SAMPLE
    if not os.path.exists(src):
        print('样本不存在：%s' % src)
        return 2
    if is_locked(src):
        print('样本正被 WPS/Word 独占，先关掉：%s' % src)
        return 2

    work = tempfile.mkdtemp(prefix='pdd-selftest-')
    docx = os.path.join(work, os.path.basename(src))
    shutil.copy2(src, docx)
    print('样本：%s\n副本：%s' % (src, docx))
    h0 = sha256_file(docx)

    # ---------------------------------------------------------- 定位
    section('定位（Locator）')
    scan = Scanner(docx)
    cites = scan.citations()
    check('扫描到引注', len(cites) > 0, 'citations=%d' % len(cites))
    if not cites:
        return 3
    print('      引注 %d 处；位置分布 %s'
          % (len(cites), scan.summary()['citation_position_tally']))

    doc = Doc(docx)
    c0 = cites[0]
    loc = make_locator(doc.paragraphs()[c0.para_index - 1]._element, c0.para_index,
                       c0.start, c0.end, c0.text)
    p, idx = doc.locate_para(loc)
    check('段落定位命中', idx == c0.para_index, 'got=%d want=%d' % (idx, c0.para_index))
    _, r, _, ridx, a, b = doc.locate_run(loc)
    check('run 定位命中', (a, b) == (c0.start, c0.end),
          'got=(%s,%s) want=(%s,%s)' % (a, b, c0.start, c0.end))
    check('段内文本一致',
          ''.join(t.text or '' for t in p._element.iter() if t.tag.endswith('}t')) == c0.para_text)

    try:
        doc.locate_para({'para_index': 3, 'fp': 'deadbeef0000', 'text': '不存在的段落文本'})
        check('脏 locator 必须抛 LocateError', False, '居然定位成功了')
    except LocateError:
        check('脏 locator 抛 LocateError', True)

    # ---------------------------------------------------------- 幂等
    section('幂等（同一动作连跑两次）')
    before = doc.locate_run(loc)[1]
    ch1 = run_action(doc, 'cite.superscript', {'locator': loc, 'value': False})
    ch2 = run_action(doc, 'cite.superscript', {'locator': loc, 'value': False})
    check('第一次改了', ch1 is True)
    check('第二次没改（幂等）', ch2 is False)
    ch3 = run_action(doc, 'cite.superscript', {'locator': loc, 'value': True})
    check('改回上标也有改动', ch3 is True)

    ch = run_action(doc, 'cite.size', {'locator': loc, 'pt': 6.5})
    ch_again = run_action(doc, 'cite.size', {'locator': loc, 'pt': 6.5})
    check('设字号 6.5pt 生效', ch is True)
    check('再设同字号无改动', ch_again is False)

    # ---------------------------------------------------------- 动作
    section('各动作往返')
    run_action(doc, 'ref.bookmark', {'locator': make_locator(doc.paragraphs()[0]._element, 1),
                                     'name': 'pdd_selftest'})
    # 插两段再合并，验证结构性动作不伤原文
    anchor_loc = make_locator(doc.paragraphs()[10]._element, 11)
    run_action(doc, 'para.insert_after',
               {'locator': anchor_loc, 'text': 'PDD-A', 'style': ''})
    n1 = len(doc.paragraphs())
    check('插入后段落表已刷新', n1 == 510 + 1, 'n1=%d' % n1)
    # 插完立刻在内存里找 PDD-A 并再插 PDD-B
    idxA = [i for i, p in enumerate(doc.paragraphs()) if p.text.strip() == 'PDD-A']
    check('插入段落成功', len(idxA) == 1, 'idxA=%s' % idxA)
    if idxA:
        locA = make_locator(doc.paragraphs()[idxA[0]]._element, idxA[0] + 1)
        run_action(doc, 'para.insert_after', {'locator': locA, 'text': 'PDD-B'})
        idxB = [i for i, p in enumerate(doc.paragraphs()) if p.text.strip() == 'PDD-B']
        check('第二次插入成功', len(idxB) == 1)
        if idxB:
            locB = make_locator(doc.paragraphs()[idxB[0]]._element, idxB[0] + 1)
            run_action(doc, 'para.merge_into', {'locator': locA, 'target': locB})
            merged = [p.text.strip() for p in doc.paragraphs() if 'PDD-' in p.text]
            check('合并后只剩一段且文本相接',
                  len(merged) == 1 and merged[0].replace(' ', '') == 'PDD-APDD-B',
                  'merged=%s' % merged)
            # 删掉测试残留
            rest = [i for i, p in enumerate(doc.paragraphs()) if 'PDD-' in p.text]
            for i in rest:
                lm = make_locator(doc.paragraphs()[i]._element, i + 1)
                run_action(doc, 'para.delete', {'locator': lm})
            check('测试残留已清理',
                  not [p for p in doc.paragraphs() if 'PDD-' in p.text])
    check('段落数回到原值', len(doc.paragraphs()) == n1 - 1,
          'now=%d was_after_insert=%d' % (len(doc.paragraphs()), n1))

    # ---------------------------------------------------------- 落盘
    section('落盘与回读（唯一真凭据）')
    doc.save_atomic()
    check('文件确实变了', sha256_file(docx) != h0)

    scan2 = Scanner(docx)
    check('回读：书签已写入',
          'pdd_selftest' in scan2.bookmarks())
    c0b = [c for c in scan2.citations() if c.para_index == c0.para_index]
    if c0b:
        rp = scan2.runs(c0.para_index)
        rr = [x for x in rp if x.text == c0.text]
        check('回读：引注仍是上标', bool(rr) and rr[0].superscript)
        check('回读：字号 = 6.5pt',
              bool(rr) and abs((rr[0].size_pt or 0) - 6.5) < 0.01,
              'size=%s' % (rr[0].size_pt if rr else None))
    else:
        check('回读：该段仍有引注', False)

    # 跳转链接两种形态
    section('引注跳转（Word 形态 / WPS 域形态）')
    c = scan2.citations()[0]
    doc3 = Doc(docx)
    lk = make_locator(doc3.paragraphs()[c.para_index - 1]._element, c.para_index,
                      c.start, c.end, c.text)
    anchor = 'ref%d' % (c.refs[0])
    run_action(doc3, 'cite.link', {'locator': lk, 'anchor': anchor, 'form': 'hyperlink'})
    doc3.save_atomic()
    s3 = Scanner(docx)
    c3 = [x for x in s3.citations() if x.para_index == c.para_index and x.text == c.text]
    check('hyperlink 形态可被回读', bool(c3) and c3[0].anchor == anchor,
          'anchor=%s' % (c3[0].anchor if c3 else None))

    doc3 = Doc(docx)
    c = s3.citations()[0]
    lk = make_locator(doc3.paragraphs()[c.para_index - 1]._element, c.para_index,
                      c.start, c.end, c.text)
    run_action(doc3, 'cite.link', {'locator': lk, 'anchor': anchor, 'form': 'field'})
    doc3.save_atomic()
    s4 = Scanner(docx)
    c4 = [x for x in s4.citations() if x.para_index == c.para_index and x.text == c.text]
    check('WPS 域形态可被回读（link_kind=field）',
          bool(c4) and c4[0].anchor == anchor and c4[0].link_kind == 'field',
          'anchor=%s kind=%s' % (c4[0].anchor if c4 else None,
                                 c4[0].link_kind if c4 else None))

    doc3 = Doc(docx)
    c = s4.citations()[0]
    lk = make_locator(doc3.paragraphs()[c.para_index - 1]._element, c.para_index,
                      c.start, c.end, c.text)
    n_before = len([x for x in doc3.paragraphs()[c.para_index - 1]._element
                    if x.tag.endswith('}r')])
    run_action(doc3, 'cite.link', {'locator': lk, 'anchor': anchor, 'form': 'none'})
    doc3.save_atomic()
    s5 = Scanner(docx)
    c5 = [x for x in s5.citations() if x.para_index == c.para_index and x.text == c.text]
    check('去链接后 anchor 为空（域已拆干净）',
          bool(c5) and c5[0].anchor is None,
          'anchor=%s' % (c5[0].anchor if c5 else None))

    # ---------------------------------------------------------- 段落边框
    section('段落边框（页眉横线用）')
    doc4 = Doc(docx)
    l = make_locator(doc4.paragraphs()[0]._element, 1)
    ch = run_action(doc4, 'para.border', {'locator': l, 'edge': 'bottom', 'sz': 6})
    ch2 = run_action(doc4, 'para.border', {'locator': l, 'edge': 'bottom', 'sz': 6})
    check('加底边框有改动', ch is True)
    check('重复加底边框无改动（幂等）', ch2 is False)
    doc4.save_atomic()
    check('文档仍可被 python-docx 打开', bool(Scanner(docx).summary()))

    # ---------------------------------------------------------- 闸门
    section('快照 / 哈希门禁 / 账本')
    sm = SnapshotManager(docx, keep=3)
    snap = sm.create(tag='selftest')
    check('快照已生成', os.path.exists(snap) and os.path.getsize(snap) > 0)
    check('快照内容与当前文档一致', sha256_file(snap) == sha256_file(docx))

    h_now = sha256_file(docx)
    plan = {
        'version': 1, 'docx': docx, 'docx_sha256': 'ff' * 32,
        'generated_at': '2026-09-17T00:00:00', 'items': [
            {'issue_id': 'cite-size-test', 'action': 'cite.size', 'enabled': True,
             'title': '测试字号', 'params': {'locator': loc, 'pt': 7.5}},
        ]}
    j = apply_plan(docx, plan, dry_run=False)
    check('哈希不符时拒绝落盘', not j.get('saved_to') and bool(j['failed']),
          'failed=%s' % j['failed'][:1])
    check('文档未被改动', sha256_file(docx) == h_now)

    # 用「当前哈希」重建 plan，让第一批真的能跑
    good = dict(plan, docx_sha256=h_now)
    good['items'] = [
        {'issue_id': 'selftest-dim', 'action': 'ref.bookmark', 'enabled': True,
         'title': '测试书签 2', 'params': {
             'locator': make_locator(Doc(docx).paragraphs()[1]._element, 2),
             'name': 'pdd_selftest2'}},
    ]
    j1 = apply_plan(docx, good, dry_run=False)
    check('第一批实际改动 1 条', j1['changed_count'] == 1,
          'changed=%d failed=%s' % (j1['changed_count'], j1['failed'][:1]))
    check('第一批写了快照', bool(j1.get('snapshot')))

    h1 = sha256_file(docx)
    good['docx_sha256'] = h1
    j2 = apply_plan(docx, good, dry_run=False)
    check('第二批全部幂等跳过', j2['changed_count'] == 0 and len(j2['skipped']) == 1,
          'changed=%d skipped=%d' % (j2['changed_count'], len(j2['skipped'])))
    check('幂等跳过也没再改文件', sha256_file(docx) == h1)

    j3 = apply_plan(docx, good, dry_run=True)
    check('dry-run 不写盘且仍能算改动数',
          (not j3.get('saved_to')) and sha256_file(docx) == h1)

    check('回滚可用', sm.restore(snap) and sha256_file(docx) == h_now)

    # ---------------------------------------------------------- 收尾
    print('\n' + '=' * 60)
    print('通过 %d 项，失败 %d 项' % (_ok, _fail))
    if _notes:
        print('失败明细：')
        for n in _notes:
            print('  - ' + n)
    print('副本留在：%s（想复查可以用 WPS 打开）' % work)
    return 0 if _fail == 0 else 1


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(9)
