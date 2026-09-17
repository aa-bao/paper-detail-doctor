#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
workflow/plan.py —— 把 audit 结果变成一张"给人勾选"的改稿方案（仍然不写文档）

这一步存在的全部理由
    audit 和 apply 之间必须有人。因为：
      · 有些"问题"是你的体例选择，不是错误（所以每条都有 enabled 开关）；
      · 有些修法会牵动全篇（重编号、删文献），机器不该替你决定；
      · 你要能在动手前一眼看清"这次到底会改哪几处"。
    所以 plan 不是中间产物，它是**你签字的那张纸**。

默认勾选策略（保守）
    error 且带自动修复动作         → enabled: true   （明确违规、且改动局部、幂等）
    warn/info 且带自动修复动作     → enabled: false  （可疑项要你确认）
    没有自动修复动作               → 不进 plan，只在报告里列着
    不可抑制（suppressible=false） → enabled: true   （结构性硬伤，必须处理）

用法
    python workflow/plan.py -d "F:/论文/初稿.docx"
    python workflow/plan.py -d 初稿.docx -o my-plan.yaml
    python workflow/plan.py -d 初稿.docx --all      # 连 warn 也默认勾上
"""
import argparse
import os
import sys

from _common import (audit_dir, banner, dump_plan_yaml, load_audit, now,
                     out_dir, resolve_params_from_audit, write_json)  # noqa: E402
import os.path as _p
sys.path.insert(0, _p.dirname(_p.dirname(_p.abspath(__file__))))

from shared.lib.ooxml_guard import sha256_file   # noqa: E402


def build_plan(docx, audit_blob, auto_warn=False):
    items = []
    for it in audit_blob.get('issues', []):
        act = it.get('action')
        if not act or not act.get('name'):
            continue                     # 只能人工改的，不进勾选表
        sev = it.get('severity')
        if it.get('suppressible') is False:
            enabled = True               # 结构性硬伤：默认必须处理
        elif sev == 'error':
            enabled = True
        elif auto_warn:
            enabled = True
        else:
            enabled = False
        loc = it.get('location') or {}
        items.append({
            'issue_id': it['id'], 'rule': it.get('rule'), 'skill': it.get('skill'),
            'title': _title(it),
            'location': _loc_text(loc),
            'action': act['name'],
            'params': act.get('params') or {},
            'enabled': enabled,
            'reason': None,
        })
    items.sort(key=lambda x: (not x['enabled'], x.get('rule') or ''))
    n_err = sum(1 for it in audit_blob.get('issues', []) if it['severity'] == 'error')
    n_manual = sum(1 for it in audit_blob.get('issues', [])
                   if not (it.get('action') or {}).get('name'))
    notes = []
    if items:
        notes.append('本批含 %d 条可自动修复，默认勾选 %d 条。'
                     % (len(items), sum(1 for x in items if x['enabled'])))
    if n_manual:
        notes.append('另有 %d 条**只能人工改**（删文献要连带重编号这类），已在 audit 报告里列出。' % n_manual)
    if n_err:
        notes.append('其中 %d 条为错误级。' % n_err)
    plan = {
        'version': 1, 'docx': docx,
        'docx_sha256': audit_blob.get('meta', {}).get('docx_sha256') or sha256_file(docx),
        'generated_at': now(),
        'audit_ref': os.path.join(audit_dir(docx), 'audit.json'),
        'items': items,
        'notes': ' '.join(notes) if notes else None,
    }
    return plan, {'manual_count': n_manual,
                  'total_issues': len(audit_blob.get('issues', []))}


def _title(it):
    return '%s · %s' % (it.get('rule'), it.get('suggestion') or '')


def _loc_text(loc):
    bits = []
    if loc.get('index'):
        bits.append('段 %s' % loc['index'])
    if loc.get('excerpt'):
        bits.append('「%s」' % loc['excerpt'][:40])
    return ' '.join(bits)


def main():
    ap = argparse.ArgumentParser(description='生成改稿方案（不写文档）')
    ap.add_argument('-d', '--docx', required=True)
    ap.add_argument('-o', '--out', default=None)
    ap.add_argument('--all', action='store_true', help='warn/info 也默认勾上')
    a = ap.parse_args()

    docx = os.path.abspath(a.docx)
    blob = load_audit(docx)
    if not blob:
        raise SystemExit('还没跑过 audit。先执行：\n'
                         '  python workflow/audit_all.py -d "%s"' % docx)

    plan, stat = build_plan(docx, blob, a.all)
    out = a.out or os.path.join(out_dir(docx), 'plan.yaml')
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        f.write(dump_plan_yaml(plan))
    write_json(os.path.join(out_dir(docx), 'plan.json'), plan)

    banner('改稿方案已生成')
    print('可自动修复：%d 条（默认勾选 %d 条）｜只能人工改：%d 条｜审计共 %d 条'
          % (len(plan['items']), sum(1 for x in plan['items'] if x['enabled']),
             stat['manual_count'], stat['total_issues']))
    print('')
    print('方案文件：%s' % out)
    print('')
    print('请打开它，把要执行的条目改成 enabled: true（默认已按保守策略勾好），')
    print('然后执行：')
    print('  python workflow/apply.py -p "%s"' % out)
    print('')
    print('★ 想先看会发生什么、不落盘：加 --dry-run。')
    return 0


if __name__ == '__main__':
    sys.exit(main())
