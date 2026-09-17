#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
workflow/audit_all.py —— 跑全部适用子 skill，合成**一份**报告（只读，不写盘改文档）

为什么要有个"全跑一遍"的入口
    单跑某个子 skill 适合"我知道问题在哪"；但更多时候用户想的是
    "帮我看看这份论文有没有问题"。这时候要的是**一份**按严重度排序的报告，
    而不是七个 skill 各吐一份。合并在这一层做，子 skill 只管产 Issue。

目前接入：cite-doctor（L2）。其余子 skill 建好后会自动被发现 —— 见 SKILLS 表。

用法
    python workflow/audit_all.py -d "F:/论文/初稿.docx"
    python workflow/audit_all.py -d 初稿.docx --only cite-doctor
    python workflow/audit_all.py -d 初稿.docx --json
"""
import argparse
import os
import sys

from _common import (CITE, audit_dir, banner, load_audit, now, write_json)  # noqa: E402
import os.path as _p
sys.path.insert(0, _p.dirname(_p.dirname(_p.abspath(__file__))))

from shared.lib import report as R                                        # noqa: E402
from shared.lib.ooxml_guard import is_locked, sha256_file                 # noqa: E402
from shared.lib.docx_scan import Scanner                                  # noqa: E402

# 子 skill 注册表。新增子 skill：加一行，并在 audit_dir 下产出 audit.json。
SKILLS = [
    {'id': CITE, 'layer': 'L2', 'run': 'skills/cite-doctor/scripts/audit.py'},
]


def run_one(docx, skill, spec=None, verbose=True):
    """在**子进程**里跑子 skill 的 audit。

    为什么用子进程而不是 import：子 skill 之间将来会有不兼容的依赖，
    也可能某个 skill 崩了。进程隔离之后，一个 skill 挂掉不影响其余 ——
    体检工具最忌讳"因为第 3 项检查报错，前两项结果也丢了"。
    """
    import subprocess
    script = os.path.join(_p.dirname(_p.dirname(_p.abspath(__file__))), skill['run'])
    cmd = [sys.executable, script, '-d', docx]
    if spec:
        cmd += ['-s', spec]
    p = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if verbose and p.stdout:
        print(p.stdout.strip())
    if p.returncode != 0:
        print('  [%s] 执行失败（退出码 %s）' % (skill['id'], p.returncode))
        if p.stderr:
            print('  ' + (p.stderr.strip().splitlines() or [''])[-1])
        return None
    return load_audit(docx, skill['id'])


def main():
    ap = argparse.ArgumentParser(description='论文全项体检（只读）')
    ap.add_argument('-d', '--docx', required=True)
    ap.add_argument('-s', '--spec', default=None)
    ap.add_argument('--only', default=None, help='只跑指定子 skill（逗号分隔）')
    ap.add_argument('--json', action='store_true')
    a = ap.parse_args()

    docx = os.path.abspath(a.docx)
    if not os.path.exists(docx):
        raise SystemExit('文档不存在：%s' % docx)
    if is_locked(docx):
        raise SystemExit('文档正被 WPS/Word 独占，请先关闭：%s' % docx)

    banner('论文体检 · %s' % os.path.basename(docx))
    print('（只读审计，不会改动文档）')

    wanted = [s.strip() for s in a.only.split(',')] if a.only else None
    issues, suppressed, notes = [], [], []
    ran = []
    for skill in SKILLS:
        if wanted and skill['id'] not in wanted:
            continue
        if not os.path.exists(os.path.join(
                _p.dirname(_p.dirname(_p.abspath(__file__))), skill['run'])):
            continue                                  # 尚未实现的子 skill 安静跳过
        b = run_one(docx, skill, a.spec)
        ran.append(skill['id'])
        if not b:
            notes.append('子 skill `%s` 未产出结果。' % skill['id'])
            continue
        for it in b.get('issues', []):
            it.setdefault('skill', skill['id'])
            issues.append(it)
        suppressed += b.get('suppressed', [])
        for n in (b.get('meta') or {}).get('notes', []):
            notes.append(n)

    # 跨 skill 汇总后重排：先按严重度，再按段号
    issues.sort(key=lambda x: (R.SEV_ORDER.get(x.get('severity'), 9),
                              x.get('location', {}).get('index') or 0))
    summary = Scanner(docx).summary()
    meta = {
        'skills_run': ran, 'generated_at': now(),
        'docx_sha256': sha256_file(docx),
        'notes': list(dict.fromkeys(notes)),          # 去重保序
    }
    meta['standard_source_note'] = '；'.join(
        dict.fromkeys(_std_line(it) for it in issues)) or '（本次未触碰标准相关判定）'

    # ★ 产物位置必须与 _common.audit_dir 一致（.thesis-doctor/audit/）——
    #   第一版直接写到 .thesis-doctor/ 下，结果 plan.py 按约定路径找不到，
    #   报"还没跑过 audit"，而报告其实已经生成了。约定路径只有一处，别各写各的。
    d = audit_dir(docx)
    md_path = os.path.join(d, 'audit-report.md')
    js_path = os.path.join(d, 'audit.json')
    R.write(d, issues, suppressed, summary, meta)
    write_json(js_path, dict(R.render_json(issues, suppressed, summary, meta),
                             skills_run=ran, generated_at=meta['generated_at'],
                             docx_sha256=meta['docx_sha256']))

    print('')
    print('跑了：%s' % ('、'.join(ran) or '（无）'))
    print('问题：%d 条（错误 %d / 警告 %d / 提示 %d）'
          % (len(issues),
             sum(1 for i in issues if i['severity'] == 'error'),
             sum(1 for i in issues if i['severity'] == 'warn'),
             sum(1 for i in issues if i['severity'] == 'info')))
    if issues:
        tally = {}
        for i in issues:
            tally[i['rule']] = tally.get(i['rule'], 0) + 1
        print('规则分布：%s' % '　'.join(
            '%s×%d' % (k, v) for k, v in sorted(tally.items(), key=lambda x: -x[1])))
    if suppressed:
        print('另有 %d 条已按 .thesisignore 裁定保留（报告末尾单列）' % len(suppressed))
    print('报告：%s' % md_path)
    print('数据：%s' % js_path)
    if issues:
        print('')
        print('下一步：python workflow/plan.py -d "%s"  # 生成勾选表，再 apply' % docx)

    if a.json:
        import json
        print(json.dumps(R.render_json(issues, suppressed, summary, meta),
                         ensure_ascii=False, indent=2))
    return 0 if not [i for i in issues if i['severity'] == 'error'] else 1


def _std_line(it):
    return '%s=%s(%s)' % (it.get('rule'), it.get('standard_source'),
                          R.SRC_LABEL.get(it.get('standard_source'), ''))


if __name__ == '__main__':
    sys.exit(main())
