#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
tests/run_all.py —— 跑齐全部验收，末尾给一张总表

改完共享层或任何子 skill 之后跑这一个就够了。三份测试各有分工，缺一不可：
    selftest_shared_lib   读侧/写侧/闸门的**机制**是否成立（定位、幂等、快照、回滚）
    selftest_cite_doctor  规则是否**真的能查出问题**（注入缺陷 → 检出 → 修复 → 复检）
    e2e_cli               命令行**串起来**能不能用（yaml 往返、路径约定、门禁）

用法
    cd F:\Coding\Project\paper-detail-doctor
    "C:/Users/Tian/.workbuddy/binaries/python/envs/default/Scripts/python.exe" tests/run_all.py
    ... tests/run_all.py "D:/别的样本.docx"
"""
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(HERE)
PY = sys.executable

SUITES = [
    ('共享层机制', 'selftest_shared_lib.py'),
    ('cite-doctor 规则（注入式）', 'selftest_cite_doctor.py'),
    ('命令行全链路', 'e2e_cli.py'),
]


def main():
    args = sys.argv[1:]
    rows, total_ok, total_fail = [], 0, 0
    for label, script in SUITES:
        p = subprocess.run([PY, os.path.join(HERE, script)] + args,
                           cwd=PKG, capture_output=True, text=True,
                           encoding='utf-8', errors='replace')
        out = (p.stdout or '') + (p.stderr or '')
        out = '\n'.join(l for l in out.splitlines()
                        if 'shell-runtime-bash-env' not in l and 'command not found' not in l)
        m = re.search(r'通过 (\d+) 项，失败 (\d+) 项', out)
        ok, bad = (int(m.group(1)), int(m.group(2))) if m else (0, -1)
        total_ok += max(ok, 0)
        total_fail += max(bad, 0) or (0 if m else 1)
        rows.append((label, script, ok, bad if m else -1, out))

    print('\n' + '=' * 70)
    print('%-28s %-26s %6s %6s' % ('测试', '脚本', '通过', '失败'))
    print('-' * 70)
    for label, script, ok, bad, _ in rows:
        print('%-28s %-26s %6d %6s' % (label, script, ok, bad if bad >= 0 else '崩溃'))
    print('-' * 70)
    print('%-28s %-26s %6d %6d' % ('合计', '', total_ok, total_fail))

    if total_fail:
        print('\n失败详情：')
        for label, script, ok, bad, out in rows:
            if bad:
                print('\n---- %s ----' % label)
                started = False
                for line in out.splitlines():
                    if line.startswith('==='):
                        started = False
                    if '[FAIL]' in line or '失败明细' in line:
                        started = True
                    if started:
                        print('  ' + line)
    print('\n结论：%s' % ('全部通过' if not total_fail else '存在失败，见上'))
    return 0 if not total_fail else 1


if __name__ == '__main__':
    sys.exit(main())
