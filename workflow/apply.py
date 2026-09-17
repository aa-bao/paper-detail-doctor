#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
workflow/apply.py —— 唯一真正落盘改文档的入口（薄壳，重活全在 ooxml_guard）

这里只做四件事，一件都不多做：
    1. 读勾选表，回 audit.json 取动作参数；
    2. 打印"将要发生什么"的预览（段号 + 动作 + 人话说明）；
    3. 交给 ooxml_guard.apply_plan —— 它负责哈希门禁、快照、幂等、倒序、原子保存；
    4. 把执行结果打出来，并告诉用户怎么回滚。

为什么把改动都压在 ooxml_guard 里，而不是写在这里
    因为"改文档"这件事只有一个正确实现，不允许有两个入口各写一套。
    apply.py 是命令行外壳，谁想从代码里调，直接调 ooxml_guard.apply_plan。

用法
    python workflow/apply.py -p "F:/论文/.thesis-doctor/plan.yaml" --dry-run   # 先看
    python workflow/apply.py -p "F:/论文/.thesis-doctor/plan.yaml"             # 真改
    python workflow/apply.py -p plan.yaml --rollback                            # 回滚到上一份
"""
import argparse
import os
import sys

from _common import (banner, out_dir, parse_plan_yaml, resolve_params_from_audit,
                     write_json)                                            # noqa: E402
import os.path as _p
sys.path.insert(0, _p.dirname(_p.dirname(_p.abspath(__file__))))

from shared.lib.ooxml_guard import (                                         # noqa: E402
    SnapshotManager, apply_plan, guard_plan, sha256_file,
)


def main():
    ap = argparse.ArgumentParser(description='执行改稿方案')
    ap.add_argument('-p', '--plan', required=False, help='plan.yaml 路径')
    ap.add_argument('-d', '--docx', default=None, help='直接指定文档（配合 --plan 缺省路径）')
    ap.add_argument('--dry-run', action='store_true', help='只预览，不写盘')
    ap.add_argument('--yes', action='store_true', help='跳过确认（脚本化时用）')
    ap.add_argument('--rollback', action='store_true', help='回滚到最近一份快照')
    ap.add_argument('--list-snapshots', action='store_true')
    a = ap.parse_args()

    # ---------------------------------------------------------- 回滚
    if a.list_snapshots or a.rollback:
        docx = a.docx or _docx_from_plan(a.plan)
        sm = SnapshotManager(docx)
        snaps = sm.list()
        if a.list_snapshots:
            print('快照（新 → 旧）：')
            for s in snaps:
                print('  %s' % s)
            return 0
        if not snaps:
            raise SystemExit('没有可用快照。')
        print('将把文档回滚到：%s' % snaps[0])
        print('（回滚前会自动再存一份当前状态，防止二次误操作）')
        sm.restore(snaps[0])
        print('已回滚：%s' % docx)
        return 0

    if not a.plan:
        raise SystemExit('需要 -p plan.yaml（或用 --rollback / --list-snapshots）')
    plan_doc = parse_plan_yaml(a.plan)
    docx = a.docx or plan_doc.get('docx')
    if not docx:
        raise SystemExit('plan.yaml 里没有 docx 字段，请用 -d 指定文档。')
    docx = os.path.abspath(docx)

    wanted = [it for it in plan_doc['items'] if it.get('enabled')]
    if not wanted:
        banner('没有需要执行的条目')
        print('勾选表里没有任何 enabled: true 的条目，什么都没做。')
        print('（打开 plan.yaml 勾选后再跑）')
        return 0

    # ---------------------------------------------------------- 前置体检
    # ★ 顺序很重要：**先**验哈希与目标，**再**回填参数。
    #   反过来的话，一份过期的 plan（issue_id 已失效）会在回填阶段被清空成
    #   "0 条要执行"，于是打印一句"什么都没做"就退出 0 —— 用户以为成功了，
    #   实际上他手里的 plan 早就该重跑 audit 了。过期必须报错，不能报"没事做"。
    checks = guard_plan(docx, {'docx': docx, 'docx_sha256': plan_doc.get('docx_sha256'),
                               'items': wanted})
    for lv, msg in checks:
        print('  %s %s' % ({'error': '✗', 'warn': '!', 'info': '·'}[lv], msg))
    if any(lv == 'error' for lv, _ in checks) and not a.dry_run:
        print('')
        print('前置体检未通过，已中止（没有改动任何文件）。')
        print('若提示哈希不符：说明文档在生成 plan 之后被改过，请重跑 audit 再生成 plan。')
        return 2

    # ---------------------------------------------------------- 回填参数
    items, missing = resolve_params_from_audit(docx, wanted)
    if missing and not items:
        print('')
        print('plan 里勾选的 %d 条在 audit.json 中**一条都找不到** —— plan 已过期。'
              % len(missing))
        print('请重新执行：audit_all.py → plan.py，再 apply。')
        return 2
    if missing:
        print('注意：有 %d 条在 audit.json 里找不到对应记录，已跳过：%s'
              % (len(missing), '、'.join(missing[:5])))
    enabled = items

    banner('执行前预览 · %s' % os.path.basename(docx))
    for it in enabled:
        print('  [即将执行] 段%s  %-24s  %s'
              % (it.get('params', {}).get('locator', {}).get('para_index') or '-',
                 it.get('action'), it.get('title') or ''))
    print('')
    print('共 %d 条。' % len(enabled))

    if not a.yes and not a.dry_run:
        print('')
        try:
            ans = input('确认执行？输入 y 继续：').strip().lower()
        except EOFError:
            ans = 'n'
        if ans not in ('y', 'yes'):
            print('已取消。')
            return 0

    # ---------------------------------------------------------- 执行
    journal = apply_plan(docx, {'docx': docx,
                                'docx_sha256': plan_doc.get('docx_sha256'),
                                'items': enabled, 'version': 1},
                         dry_run=a.dry_run)
    write_json(os.path.join(out_dir(docx), 'last-apply.json'), journal)

    banner('执行结果' + ('（试运行，未写盘）' if a.dry_run else ''))
    print('实际改动：%d 条' % journal['changed_count'])
    print('幂等跳过：%d 条' % len(journal['skipped']))
    print('未成功：  %d 条' % len(journal['failed']))
    for r in journal['applied'][:40]:
        print('  ✓ 段%s %s' % (r.get('para_index') or '-', r.get('title') or r.get('action')))
    for r in journal['failed']:
        print('  ✗ 段%s %s —— %s'
              % (r.get('para_index') or '-', r.get('title') or r.get('action'),
                 r.get('reason')))
    print('')
    if not a.dry_run:
        if journal.get('snapshot'):
            print('改前快照：%s' % journal['snapshot'])
            print('不满意就回滚：python workflow/apply.py -d "%s" --rollback' % docx)
        print('详情：%s' % (journal.get('journal_md') or ''))
    else:
        print('这是试运行。确认无误后去掉 --dry-run 再跑一次。')
    return 0


def _docx_from_plan(plan_path):
    if not plan_path:
        raise SystemExit('需要 -d 指定文档，或 -p 指定 plan.yaml。')
    return parse_plan_yaml(plan_path).get('docx')


if __name__ == '__main__':
    sys.exit(main())
