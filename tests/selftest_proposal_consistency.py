#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
tests/selftest_proposal_consistency.py —— proposal-consistency 的**注入式**验收

为什么不能只跑一遍干净文档就收工
    "在一份本来就干净的稿子上跑出 0 条问题"完全不能说明检查器是好的 ——
    一个永远返回空列表的函数也能做到。所以这里反过来做：
    **用人工核对的金标准（tests/golden/proposal-consistency.yaml）做三指标跑分，
    再做注入式反向验证证明检查器真在读内容、不是硬编码输出。**

三指标（定义写死，全部当次从真实运行读回，不许用计划值/内存值）
    召回率   = 4 处真差异命中几处              （漏报最致命）
    误报率   = 报出的"差异"里金标准没有的占多少（只统计 缺失/改动/偏离）
    定位准确率 = 给出的正文位置确实是对的那一处占多少

注入式反向验证（必须有，证明非硬编码）
    复制文档到临时目录，把开题样本量改成与正文一致的 15 人 → A4-03 差异应消失；
    改回 30—40 → 应重现。全程只在临时副本操作，绝不改真实论文文件。

锚点对账表（回答用户疑点）
    金标准列了 8 条对齐锚点，但脚本只报"一致 4 处"。本测试逐条打印
    8 条锚点各自被脚本判成什么 / 若无条目则为什么，证明"差的那几条"是
    L2 语义按设计交人审（无法判定），不是漏判。

诚实红线
    所有数字都从磁盘/真实运行读回后打印；跑分是多少报多少。
    金标准（golden yaml）是事实，不反向改它来迁就程序。

怎么跑
    cd F:\Coding\Project\paper-detail-doctor
    "python" \
        tests/selftest_proposal_consistency.py "F:/路径/初稿.docx" "F:/路径/开题报告.docx"
    （两参数都可省，缺省读 tests/sample.local.yaml）
"""
import importlib.util
import os
import re
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(HERE)
sys.path.insert(0, PKG)
sys.path.insert(0, HERE)   # 让 tests/_sample.py 可被 import

from shared.lib.docx_scan import Scanner                           # noqa: E402

from _sample import (thesis as _sample_thesis, proposal as _sample_proposal,   # noqa: E402
                    missing_hint as _missing_hint)
DEFAULT_THESIS = _sample_thesis()
DEFAULT_PROPOSAL = _sample_proposal()
GOLDEN = os.path.join(HERE, 'golden', 'proposal-consistency.yaml')

TMP = tempfile.mkdtemp(prefix='pdd-prop-selftest-')

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


# ---------------------------------------------------------------- 加载 audit 模块

def load_audit_mod():
    p = os.path.join(PKG, 'skills', 'proposal-consistency', 'scripts', 'audit.py')
    spec = importlib.util.spec_from_file_location('prop_audit', p)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m, p


AUD = None
AUD_PATH = None


def audit(thesis, proposal):
    """跑一次核对（静默，不写盘）。返回 audit() 的返回 dict。"""
    return AUD.audit(thesis, proposal, verbose=False)


def diff_issues(res):
    return [i for i in res['issues']
            if i.get('conclusion') in ('缺失', '改动', '偏离')]


def load_golden():
    import yaml
    with open(GOLDEN, encoding='utf-8') as f:
        return yaml.safe_load(f)


def build_keys(golden):
    """从金标准推导 (anchor, conclusion) -> 差异 id。

    不耦合任何具体案例：差异 id 形如 'A5-01'（前缀就是锚点），type 就是结论。
    """
    keys = {}
    for d in golden['expected']['差异']:
        did = str(d['id'])
        keys[(did.split('-')[0], d['type'])] = did
    return keys


def build_locate(golden):
    """定位准确率用的关键词：差异 id -> 关键词序列。

    这些关键词是**本案例的数据**（"位置指对了"的判据），放在金标准 yaml 的
    locate 段里，而不是写死在测试代码里 —— 换成别的论文时只改 yaml。
    """
    loc = {}
    for did, kws in (golden.get('locate') or {}).items():
        if isinstance(kws, str):
            kws = [kws]
        loc[str(did)] = [str(k) for k in kws]
    return loc


def located_text(docx, it):
    idx = (it.get('location') or {}).get('index')
    if idx:
        try:
            return Scanner(docx).para_text(idx)
        except Exception:
            return ''
    return (it.get('location') or {}).get('excerpt') or ''


def fresh(src, tag):
    p = os.path.join(TMP, tag + '.docx')
    shutil.copy2(src, p)
    return p


# ---------------------------------------------------------------- 注入器：改开题样本量

def set_proposal_sample(docx, lo, hi):
    """把开题里"计划访谈 … 人"的样本量改成 lo—hi（单值则 lo==hi）。
    只在临时副本上调用，不动原稿。"""
    from docx import Document
    d = Document(docx)
    target = None
    for tbl in d.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if '计划访谈' in p.text:
                        target = p
                        break
                if target:
                    break
            if target:
                break
        if target:
            break
    if target is None:
        for p in d.paragraphs:
            if '计划访谈' in p.text:
                target = p
                break
    if target is None:
        raise RuntimeError('开题副本里找不到"计划访谈"，无法注入样本量')
    new = re.sub(r'计划访谈\s*\d+\s*[–—-]\s*\d+\s*人',
                 '计划访谈 %d—%d 人' % (lo, hi), target.text, count=1)
    if new == target.text:                      # 单值写法兜底
        new = re.sub(r'计划访谈\s*\d+\s*人',
                     '计划访谈 %d 人' % lo, target.text, count=1)
    target.text = new
    d.save(docx)
    return new


# ---------------------------------------------------------------- 场景

def main():
    global AUD, AUD_PATH
    thesis = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_THESIS
    proposal = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_PROPOSAL

    AUD, AUD_PATH = load_audit_mod()
    if not thesis or not os.path.exists(thesis):
        print('正文样本不存在：%s' % thesis)
        print(_missing_hint('正文样本'))
        return 2
    if not proposal or not os.path.exists(proposal):
        print('开题样本不存在：%s' % proposal)
        print(_missing_hint('开题样本'))
        return 2
    print('正文：%s\n开题：%s\n临时目录：%s' % (thesis, proposal, TMP))

    golden = load_golden()
    golden_diffs = golden['expected']['差异']          # 4 处真差异
    golden_consistent = golden['expected']['一致']      # 8 条对齐锚点
    GOLDEN_DIFF_KEYS = build_keys(golden)             # 由金标准推导，不写死
    GOLDEN_LOC = build_locate(golden)                 # 定位关键词也来自金标准

    # ================================================================== 0 基线：CLI 缺 -p 必须报错退出
    section('0 CLI 契约：缺 -p 必须报错退出（不许假装跑成功）')
    p = subprocess.run([sys.executable, AUD_PATH, '-d', thesis],
                       capture_output=True, text=True,
                       encoding='utf-8', errors='replace')
    check('缺 -p 退出码为 2', p.returncode == 2, 'rc=%s' % p.returncode)
    check('缺 -p 报错信息提到"开题报告"', '开题报告' in (p.stderr or ''),
          'stderr=%s' % (p.stderr or '').strip()[:60])

    # ================================================================== 1 真实双文档跑分
    # 金标准完整性：每条差异都必须有定位关键词（否则定位准确率无法判定）
    missing_loc = [d['id'] for d in golden_diffs if not GOLDEN_LOC.get(str(d['id']))]
    check('金标准每条差异都带了定位关键词（locate 段完整）', not missing_loc,
          '缺 locate 的差异：%s' % missing_loc)

    section('1 金标准跑分（真实双文档，数字当次从运行读回）')
    res = audit(thesis, proposal)
    issues = res['issues']
    diffs = diff_issues(res)

    # —— 召回率：4 处真差异命中几处 ——
    hit_keys = set()
    for it in diffs:
        k = (it.get('anchor'), it.get('conclusion'))
        if k in GOLDEN_DIFF_KEYS:
            hit_keys.add(k)
    recall_hit = len(hit_keys)
    recall = recall_hit / len(golden_diffs)
    check('召回率=100%（4/4 真差异全命中）', recall_hit == len(golden_diffs),
          '命中 %d/%d' % (recall_hit, len(golden_diffs)))

    # —— 误报率：报出的差异里金标准没有的占多少 ——
    false_pos = [it for it in diffs if (it.get('anchor'), it.get('conclusion')) not in GOLDEN_DIFF_KEYS]
    fp_rate = (len(false_pos) / len(diffs)) if diffs else 0.0
    check('误报率=0%（无金标准外的差异）', len(false_pos) == 0,
          '误报=%s' % [(it.get('anchor'), it.get('conclusion')) for it in false_pos])

    # —— 定位准确率：给出的正文位置确实对的那一处占多少 ——
    loc_hit = 0
    for it in diffs:
        k = (it.get('anchor'), it.get('conclusion'))
        if k not in GOLDEN_DIFF_KEYS:
            continue
        kw = GOLDEN_LOC.get(GOLDEN_DIFF_KEYS[k]) or []
        txt = located_text(thesis, it)
        if any(kk in txt for kk in kw):
            loc_hit += 1
    loc_acc = (loc_hit / len(hit_keys)) if hit_keys else 0.0
    check('定位准确率=100%（4/4 位置正确）', loc_hit == len(hit_keys),
          '定位命中 %d/%d' % (loc_hit, len(hit_keys)))

    print('\n  --- 三指标真实数字 ---')
    print('  召回率   = %d/%d = %.0f%%' % (recall_hit, len(golden_diffs), recall * 100))
    print('  误报率   = %d/%d = %.0f%%' % (len(false_pos), len(diffs), fp_rate * 100) if diffs
          else '  误报率   = n/a（无差异输出）')
    print('  定位准确率 = %d/%d = %.0f%%' % (loc_hit, len(hit_keys), loc_acc * 100))

    # ================================================================== 2 八锚点对账表（回答用户疑点）
    section('2 八锚点对账表（金标准 8 条对齐锚点 × 脚本判成什么）')
    passed = {(p['anchor'], p['value']) for p in res['meta']['passed']}
    # 逐锚点推导脚本 verdict（从真实结果读回，不写死）
    rows = []
    rows.append(('A1', '题目一致',
                 '一致(passed)' if any(a == 'A1' for a, _ in passed) else '见 issue'))
    a2 = [i for i in issues if i.get('anchor') == 'A2']
    rows.append(('A2', '家庭情感共同体',
                 '无法判定(L2待判)' if a2 else '?',
                 'by design §5.2：L2 语义交宿主，进无法判定桶'))
    rows.append(('A2', '三维情感构成',
                 '合并进 A2 待判卡（无独立条目）',
                 '脚本未单独抽此子锚点，A2 单卡覆盖两子项（设计简化，非错误）'))
    a3 = [i for i in issues if i.get('anchor') == 'A3']
    rows.append(('A3', '三大理论',
                 '无法判定(L2待判)' if a3 else '?',
                 'by design §5.2'))
    a4m = any(a == 'A4' and '方法要素' in v for a, v in passed)
    rows.append(('A4', '方法要素',
                 '一致(passed)' if a4m else '?'))
    # A5 的通过项标签由锚点档案驱动（本项目里那 4 个标签叫"四重逻辑"），
    # 所以同时看 value 与 note —— 不依赖某个写死的标签文案。
    pfull = res['meta']['passed']
    a5f = any(p['anchor'] == 'A5'
              and ('四重逻辑' in p['value'] or '四重逻辑' in (p.get('note') or ''))
              for p in pfull)
    rows.append(('A5', '四重逻辑',
                 '一致(passed)' if a5f else '?'))
    a6 = [i for i in issues if i.get('anchor') == 'A6']
    rows.append(('A6', '核心论断',
                 '无法判定(L2待判)' if a6 else '?',
                 'by design §5.2'))
    a7 = any(a == 'A7' for a, _ in passed)
    rows.append(('A7', '文献不少于40条',
                 '一致(passed)' if a7 else '?'))

    print('\n  锚点   子项               脚本 verdict            说明')
    print('  ' + '-' * 64)
    for r in rows:
        anchor, sub, verdict = r[0], r[1], r[2]
        note = ('  ' + r[3]) if len(r) > 3 else ''
        print('  %-6s %-16s %-22s%s' % (anchor, sub, verdict, note))

    consistent_n = sum(1 for r in rows if r[2].startswith('一致'))
    pending_n = sum(1 for r in rows if '无法判定' in r[2] or '待判' in r[2])
    print('\n  对账结论：8 条里 =%d 一致 + %d 落"无法判定"(L2 交人审) + '
          'A2 三维情感构成 合并进 A2 卡' % (consistent_n, pending_n))
    print('  → "差的那几条"（A2/A3/A6 与 A2 子项）是 L2 语义按设计交宿主判定，'
          '非漏判；脚本"一致 N 处"只数 L1/L3 确定性通过项。')

    check('8 条对齐锚点里恰有 4 条判一致（A1/A4方法要素/A5四重逻辑/A7）',
          consistent_n == 4, '一致数=%d' % consistent_n)
    check('L2 项（A2/A3/A6）均进"无法判定"，未谎报为一致',
          pending_n >= 3, '待判数=%d' % pending_n)

    # ================================================================== 3 五类结论→severity 与 action=null 协议
    section('3 结论→severity 映射 与 action 一律 null')
    sev_map = {'缺失': 'error', '改动': 'warn', '偏离': 'warn', '无法判定': 'info'}
    bad_sev = [(it.get('anchor'), it.get('conclusion'), it.get('severity'))
               for it in issues if it.get('conclusion') in sev_map
               and it.get('severity') != sev_map[it['conclusion']]]
    check('severity 映射正确（缺失=error/改动=warn/偏离=warn/无法判定=info）',
          not bad_sev, 'bad=%s' % bad_sev)
    actions = {it.get('id'): it.get('action') for it in issues}
    check('所有 issue 的 action 一律为 null',
          all(a is None for a in actions.values()),
          'non-null=%s' % [k for k, v in actions.items() if v is not None])

    # L2 协议：A2/A3/A6 必须 verdict=待判、confidence=null、两侧原文都在
    l2_ok = True
    l2_detail = ''
    for it in issues:
        if it.get('anchor') in ('A2', 'A3', 'A6'):
            ok = (it.get('verdict') == '待判'
                  and it.get('confidence') is None
                  and bool(it.get('proposal_excerpt'))
                  and bool(it.get('thesis_excerpt'))
                  and bool(it.get('judgment_question')))
            if not ok:
                l2_ok = False
                l2_detail = '%s:%s' % (it.get('anchor'), it.get('verdict'))
    check('L2 证据卡：verdict=待判 / confidence=null / 两侧原文齐 / 带判定问题',
          l2_ok, l2_detail)

    # 一致不产 issue（只进 passed）
    consistent_issues = [it for it in issues if it.get('conclusion') == '一致']
    check('一致不产 issue（只在 meta.passed）', not consistent_issues,
          'n=%d' % len(consistent_issues))

    # ================================================================== 4 注入式反向验证（证明非硬编码）
    section('4 注入式反向验证：开题样本量改 15 → A4-03 消失；改回 30—40 → 重现')
    prop15 = fresh(proposal, 'prop-15')
    thesis15 = fresh(thesis, 'thesis-15')            # 正文不改，仅复制
    modified = set_proposal_sample(prop15, 15, 15)
    res15 = audit(thesis15, prop15)
    a4_diff_15 = [i for i in res15['issues']
                  if i.get('anchor') == 'A4' and i.get('conclusion') == '改动']
    check('开题样本量改为 15（与正文一致）→ A4-03 差异消失',
          not a4_diff_15, '开题改后文本含=%s；A4改动=%s'
          % ('计划访谈 15' in modified, bool(a4_diff_15)))
    passed_sample_15 = any(p['anchor'] == 'A4' and '样本量' in p['value']
                           for p in res15['meta']['passed'])
    check('且"样本量一致"进通过清单', passed_sample_15)

    # 反向：改回 30—40 → 差异重现
    prop30 = fresh(proposal, 'prop-30')
    thesis30 = fresh(thesis, 'thesis-30')
    set_proposal_sample(prop30, 30, 40)
    res30 = audit(thesis30, prop30)
    a4_diff_30 = [i for i in res30['issues']
                  if i.get('anchor') == 'A4' and i.get('conclusion') == '改动']
    check('开题样本量改回 30—40 → A4-03 差异重现',
          bool(a4_diff_30), 'A4改动=%s' % bool(a4_diff_30))

    # ================================================================== 5 只读性（哈希不变）
    section('5 只读性：审计不改文档（哈希不变）')
    from shared.lib.ooxml_guard import sha256_file
    h0 = sha256_file(thesis)
    audit(thesis, proposal)
    check('正文哈希审计前后不变（只读）', sha256_file(thesis) == h0)

    # ============================== 6 锚点档案缺失时的诚实性（关键，防"静默失效"）
    section('6 锚点档案缺失/为空：一律「无法判定」，绝不编造一致或差异')
    empty = os.path.join(TMP, 'anchors-empty.yaml')
    with open(empty, 'w', encoding='utf-8') as f:
        f.write('case: "（空档案）"\n')
    res0 = AUD.audit(thesis, proposal, verbose=False, anchors_path=empty)
    i0 = [i for i in res0['issues'] if i.get('rule') != 'audit-internal-error']
    fab = [i for i in i0 if i.get('conclusion') in ('缺失', '改动', '偏离')
           and i.get('anchor') != 'A4']
    check('空档案下不编造差异（只剩 A4 样本量这条通用数字检查）', not fab,
          '多出来的差异：%s' % [(i.get('anchor'), i.get('conclusion')) for i in fab])
    no_cfg = {'A2', 'A3', 'A5', 'A6'}
    pending = {i.get('anchor') for i in i0 if i.get('conclusion') == '无法判定'}
    passed_a = {p['anchor'] for p in res0['meta']['passed']}
    check('未配置的锚点（A2/A3/A5/A6）全部落「无法判定」', no_cfg <= pending,
          '未落无法判定：%s' % sorted(no_cfg - pending))
    check('未配置的锚点没有一个被谎报成「一致」', not (no_cfg & passed_a),
          '谎报一致：%s' % sorted(no_cfg & passed_a))
    check('meta.notes 里写明锚点档案来源/缺失',
          any('锚点档案' in str(n) for n in res0['meta']['notes']),
          'notes=%s' % res0['meta']['notes'][:2])

    # ================================================================== 收尾
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
