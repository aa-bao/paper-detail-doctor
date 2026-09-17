#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
tests/run_all.py —— 跑齐全部验收，末尾给一张总表

改完共享层或任何子 skill 之后跑这一个就够了。七份测试各有分工，缺一不可：
    selftest_shared_lib       读侧/写侧/闸门的**机制**是否成立（定位、幂等、快照、回滚）
    selftest_cite_doctor      L2 规则是否**真的能查出问题**（注入缺陷 → 检出 → 修复 → 复检）
    selftest_format_audit     L1 规则（注入页边距/页码/页眉/字号/目录域）
    selftest_text_style       L3 规则（注入引号/标点/破折号/标题/TODO，含 3 项可选）
    selftest_structure_length L4 规则（注入断号/重复题注/摘要超限 + 三口径字数交叉验证）
    selftest_proposal_consistency  开题↔正文一致性：金标准跑分（召回/误报/定位三指标，
                             真实双文档）+ 改开题内容看结论是否跟着变（反硬编码验证）
    e2e_cli                   命令行**串起来**能不能用（yaml 往返、路径约定、门禁、回滚）

五个 audit 型子 skill 的自检都遵循同一个范式：**在干净稿上人为注入缺陷 →
确认能检出 → 再确认改完就消失**。只跑"干净稿不报错"是没有意义的 ——
一个永远返回空列表的检查器也能通过那种测试。

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
    ('format-audit 规则（注入式）', 'selftest_format_audit.py'),
    ('text-style 规则（注入式）', 'selftest_text_style.py'),
    ('structure-length 规则（注入式）', 'selftest_structure_length.py'),
    ('proposal-consistency 金标准跑分', 'selftest_proposal_consistency.py'),
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
