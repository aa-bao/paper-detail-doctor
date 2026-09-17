#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
tests/e2e_cli.py —— 命令行全链路验收（audit_all → plan → apply → 回滚）

和 selftest_* 的分工
    selftest_cite_doctor.py 验的是"规则对不对"（注入缺陷 → 能不能查出来）；
    e2e_cli.py 验的是"命令串起来能不能用"（yaml 读写、哈希门禁、快照、回滚、
    幂等账本）。两者都要有：规则对了但 --dry-run 会把文档改掉，照样是废的。

它覆盖的手写风险点
    · plan.yaml 的**写出去 / 读回来**（一个是手写模板、一个是 pyyaml 解析，
      不对称最容易露馅：字段名对不上、中文冒号被当分隔符、布尔变成字符串）
    · audit 与 plan 的产物路径约定是否一致（曾经一个写 .thesis-doctor/，
      一个读 .thesis-doctor/audit/，导致"报告明明生成了却说没跑过 audit"）
    · --dry-run 是否真的不落盘
    · 重复 apply 是否幂等
    · --rollback 能否把文档还原到字节级一致

怎么跑
    cd F:\Coding\Project\paper-detail-doctor
    "python" tests/e2e_cli.py
"""
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(HERE)
sys.path.insert(0, PKG)
sys.path.insert(0, HERE)

import selftest_cite_doctor as T                                    # noqa: E402
from shared.lib.docx_scan import Scanner                            # noqa: E402
from shared.lib.ooxml_guard import sha256_file                      # noqa: E402

from _sample import thesis as _sample_thesis, missing_hint as _missing_hint   # noqa: E402
SAMPLE = _sample_thesis()
PY = sys.executable
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


def run(args, expect_rc=(0,)):
    p = subprocess.run([PY] + args, cwd=PKG, capture_output=True,
                       text=True, encoding='utf-8', errors='replace')
    out = (p.stdout or '') + (p.stderr or '')
    # 过滤掉 Git Bash 的噪声行
    out = '\n'.join(l for l in out.splitlines() if 'shell-runtime-bash-env' not in l
                    and 'command not found' not in l)
    return p.returncode, out


def main():
    if not SAMPLE or not os.path.exists(SAMPLE):
        print('样本不存在：%s' % SAMPLE)
        print(_missing_hint())
        return 2
    tmp = tempfile.mkdtemp(prefix='pdd-e2e-')
    docx = os.path.join(tmp, '初稿.docx')
    shutil.copy2(SAMPLE, docx)
    os.chmod(tmp, 0o777)
    print('样本副本：%s' % docx)

    T.AUD = T.load_audit_mod()

    # 造两个可自动修复的缺陷：ref7 书签丢失 + 一处引注掉上标
    T.inject_drop_bookmark(docx, 'ref7')
    d, loc, c = T.cite_locator(docx, 5)
    T.run_action(d, 'cite.superscript', {'locator': loc, 'value': False})
    d.save_atomic()
    h_defect = sha256_file(docx)

    # ------------------------------------------------------ audit_all
    print('\n=== 1 audit_all（只读）===')
    rc, out = run(['workflow/audit_all.py', '-d', docx])
    check('退出码非零（存在 error 级问题）', rc != 0, 'rc=%s' % rc)
    check('报告落在 .thesis-doctor/audit/',
          os.path.exists(os.path.join(tmp, '.thesis-doctor', 'audit', 'audit-report.md')))
    check('文档未被改动', sha256_file(docx) == h_defect)
    check('规则分布里出现 cite-form 与 cite-jump',
          'cite-form×' in out and 'cite-jump×' in out, out[-300:])
    j = os.path.join(tmp, '.thesis-doctor', 'audit', 'audit.json')
    blob = open(j, encoding='utf-8').read() if os.path.exists(j) else ''
    check('audit.json 存在且含 issues', '"issues"' in blob)
    import json as _json
    data = _json.loads(blob) if blob else {'issues': []}
    got = {i['rule'] for i in data['issues']}
    check('cite-form 与 cite-jump 均被检出', {'cite-form', 'cite-jump'} <= got,
          'rules=%s' % got)
    check('两条缺陷都可自动修复',
          sum(1 for i in data['issues'] if (i.get('action') or {}).get('name')) >= 2)

    # ------------------------------------------------------ plan
    print('\n=== 2 plan（生成勾选表）===')
    rc, out = run(['workflow/plan.py', '-d', docx])
    check('plan 成功', rc == 0, out[-300:])
    plan_yaml = os.path.join(tmp, '.thesis-doctor', 'plan.yaml')
    check('plan.yaml 已生成', os.path.exists(plan_yaml))
    txt = open(plan_yaml, encoding='utf-8').read()
    good_txt = txt          # 留一份：第 6 步会把 plan.yaml 覆盖成"空计划"
    check('plan.yaml 含 enabled 开关', 'enabled: true' in txt)
    check('plan.yaml 带人话注释', '# ' in txt and '段' in txt)
    n_on = txt.count('enabled: true')
    check('默认勾选了可自动修复项（>=2）', n_on >= 2, 'n_on=%d' % n_on)

    # 读回来（pyyaml）—— 验"写出去的手写模板"能被"真解析器"读懂
    rc, out = run(['-c',
                   'import sys;sys.path.insert(0,"workflow");'
                   'from _common import parse_plan_yaml;'
                   'd=parse_plan_yaml(r"%s");'
                   'print(len(d["items"]), d["items"][0]["issue_id"], '
                   'type(d["items"][0]["enabled"]).__name__)' % plan_yaml])
    check('plan.yaml 可被 pyyaml 读回', rc == 0 and 'str' not in out.split()[-1],
          out[-200:])

    # ------------------------------------------------------ dry-run
    print('\n=== 3 apply --dry-run（不得落盘）===')
    rc, out = run(['workflow/apply.py', '-p', plan_yaml, '--dry-run'])
    check('dry-run 成功', rc == 0, out[-300:])
    check('dry-run 未改文档', sha256_file(docx) == h_defect)
    check('dry-run 报出改动条数', '实际改动' in out)
    check('dry-run 明确声明未写盘', '未写盘' in out or '试运行' in out)

    # ------------------------------------------------------ apply
    print('\n=== 4 apply（真改）===')
    rc, out = run(['workflow/apply.py', '-p', plan_yaml, '--yes'])
    check('apply 成功', rc == 0, out[-400:])
    check('文档已改变', sha256_file(docx) != h_defect)
    check('写了快照', os.path.isdir(os.path.join(tmp, '.thesis-doctor', 'backups')))
    h_fixed = sha256_file(docx)

    s = Scanner(docx)
    check('ref7 书签已补回', 'ref7' in set(s.bookmarks()))
    cc = [x for x in s.citations() if x.para_index == c.para_index and x.text == c.text]
    check('引注已恢复上标', bool(cc) and cc[0].superscript)

    # ------------------------------------------------------ 门禁 + 幂等
    print('\n=== 5 拿旧 plan 再跑（哈希门禁必须拦下）===')
    rc, out = run(['workflow/apply.py', '-p', plan_yaml, '--yes'])
    check('旧 plan 被拒绝（rc=2）', rc == 2, 'rc=%s' % rc)
    check('拒绝原因是哈希不符', '哈希' in out, out[-300:])
    check('文档未被二次改动', sha256_file(docx) == h_fixed)

    print('\n=== 6 重新 audit + plan（缺陷已修完 → 计划应为空）===')
    rc, out = run(['workflow/audit_all.py', '-d', docx])
    check('cite-form / cite-jump 均已消失',
          'cite-form' not in out and 'cite-jump' not in out, out[-300:])
    # 不要断言"全稿只剩 1 条问题"：audit_all 跑的是注册表里的全部子 skill，
    # format-audit / text-style / structure-length 在样本上本来就有正常告警，
    # 总数会随接入的 skill 数变化。这里只对该 e2e 关心的引注/文献类规则下断言：
    # 注入的两类缺陷已消失，只剩样本自带的、本就不可自动修复的 ref-format 一条。
    _d2 = _json.loads(open(j, encoding='utf-8').read())
    cite_ish = [i['rule'] for i in _d2['issues']
                if i['rule'].startswith('cite-') or i['rule'].startswith('ref-')]
    check('引注/文献类只剩已知的 ref-format 一条', cite_ish == ['ref-format'],
          'cite-ish rules=%s' % cite_ish)
    rc, out = run(['workflow/plan.py', '-d', docx])
    check('新计划里没有可自动修复项', '可自动修复：0 条' in out, out[-300:])
    rc, out = run(['workflow/apply.py', '-p', os.path.join(tmp, '.thesis-doctor', 'plan.yaml'),
                   '--yes'])
    check('无可执行条目时正常退出', rc == 0, 'rc=%s' % rc)
    check('空计划也没碰文档', sha256_file(docx) == h_fixed)

    # ------------------------------------------------------ 回滚
    print('\n=== 7 rollback（回到"动手之前"）===')
    rc, out = run(['workflow/apply.py', '-d', docx, '--list-snapshots'])
    check('能列出快照', '.docx' in out, out[-200:])
    rc, out = run(['workflow/apply.py', '-d', docx, '--rollback'])
    check('回滚成功', rc == 0, out[-300:])
    # 唯一的一份 pre-apply 快照产生在 apply#1 之前，也就是"缺陷还在"的状态；
    # 所以正确的期望是回到 h_defect，而不是回到修好之后。
    check('回滚到动手之前的字节级状态', sha256_file(docx) == h_defect,
          '当前 sha8=%s 期望 %s' % (sha256_file(docx)[:8], h_defect[:8]))
    s2 = Scanner(docx)
    check('缺陷确实回来了（ref7 书签消失）', 'ref7' not in set(s2.bookmarks()))
    # 回滚本身也会先存一份"回滚前"，所以还能再回滚回去 —— 这是有意的
    rc, out = run(['workflow/apply.py', '-d', docx, '--rollback'])
    check('可以对回滚再回滚（撤销误回滚）', rc == 0 and sha256_file(docx) == h_fixed,
          out[-300:])

    # ------------------------------------------------------ 坏 plan
    print('\n=== 8 防呆：哈希不符的 plan 必须拒绝 ===')
    # 拿真 plan 改掉哈希 —— 模拟"文档在生成 plan 之后被编辑过"。
    # 不要自己另写一份小 plan：那样 issue_id 必然对不上，会被"plan 已过期"
    # 这条更早的错误拦住，测不到哈希门禁本身。
    bad = os.path.join(tmp, 'bad-plan.yaml')
    import re as _re
    bad_txt = _re.sub(r'docx_sha256: "[0-9a-f]*"', 'docx_sha256: "%s"' % ('ff' * 32), good_txt)
    open(bad, 'w', encoding='utf-8').write(bad_txt)
    h_before = sha256_file(docx)
    rc, out = run(['workflow/apply.py', '-p', bad, '--yes'])
    check('哈希不符被拒绝（退出码 2）', rc == 2, 'rc=%s' % rc)
    check('拒绝时未改文档', sha256_file(docx) == h_before)

    print('\n' + '=' * 62)
    print('通过 %d 项，失败 %d 项' % (_ok, _fail))
    for n in _notes:
        print('  - ' + n)
    print('临时目录：%s' % tmp)
    return 0 if _fail == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
