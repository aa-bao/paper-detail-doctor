#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
workflow/_common.py —— 三个 workflow 脚本的共用零件

包根 SKILL.md 只负责"路由到哪个子 skill"，真正把子 skill 串起来的执行编排在 workflow/：
    audit_all.py  只读，跑全部适用子 skill，合成一份报告
    plan.py       把 audit.json 转成 plan.yaml，给用户勾选
    apply.py      读勾选后的 plan.yaml，受控落盘

这里放的是三者都要用的东西：包路径、子 skill 发现、输出目录约定、终端提示。
"""
import json
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(HERE)
if PKG not in sys.path:
    sys.path.insert(0, PKG)

STATE_DIR = '.thesis-doctor'
AUDIT_DIR = 'audit'
CITE = 'cite-doctor'


def out_dir(docx_path):
    """审计产物一律落在文档旁边的 .thesis-doctor/ 里，跟着文档走，不污染包目录。"""
    return os.path.join(os.path.dirname(os.path.abspath(docx_path)), STATE_DIR)


def audit_dir(docx_path, skill=None):
    d = os.path.join(out_dir(docx_path), AUDIT_DIR)
    return os.path.join(d, skill) if skill else d


def load_audit(docx_path, skill=None):
    p = os.path.join(audit_dir(docx_path, skill), 'audit.json')
    if not os.path.exists(p):
        return None
    with open(p, encoding='utf-8') as f:
        return json.load(f)


def write_json(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    return path


def read_plan_yaml(path):
    """读 plan.yaml。

    **写出去用手写模板（带注释、形状可控），读回来用真 YAML 解析器。**
    为什么不对称：plan.yaml 的定位就是"给人改的勾选表"，用户会顺手加注释、
    调缩进、删整块。自己写个 YAML 子集解析器去读人改过的文件，迟早会在某个
    缩进上翻车，然后用户拿到一个莫名其妙的"格式错误"。
    """
    try:
        import yaml
    except ImportError:
        raise SystemExit(
            'plan.yaml 需要 pyyaml 解析。装一下：\n'
            '  "C:/Users/Tian/.workbuddy/binaries/python/envs/default/Scripts/pip.exe" install pyyaml')
    with open(path, encoding='utf-8') as f:
        return yaml.safe_load(f) or {}


def dump_plan_yaml(plan):
    """把 plan 写成**给人勾选**的 yaml 字符串。

    写盘形状的全部意义在于"人可以安全地手改"：
      · 每条独立成块，`enabled` 永远单独一行，改 true/false 一眼可见；
      · 每条都带 `title`，不用回头查 audit 报告就知道这条在改什么；
      · 段号写在注释里（# 段 476），改完 yaml 自己也能读。
    """
    q = lambda s: json.dumps('' if s is None else str(s), ensure_ascii=False)
    L = []
    L.append('# 改稿方案（由 audit 结果生成，**勾选后交给 apply 执行**）')
    L.append('#')
    L.append('# 用法：把要执行的条目的 enabled 改成 true，不执行的改 false（或整块删掉）。')
    L.append('# 每条只靠 issue_id 识别，不要改它。改完保存，再跑 apply.py。')
    L.append('# 改写细节（动作参数）不在本文件里 —— apply 会回 audit.json 取，')
    L.append('# 所以这里改坏了也伤不到执行逻辑，放心当一张勾选表用。')
    L.append('')
    L.append('version: 1')
    L.append('docx: %s' % q(plan['docx'].replace('\\', '/')))
    L.append('docx_sha256: %s' % q(plan.get('docx_sha256') or ''))
    L.append('generated_at: %s' % q(plan['generated_at']))
    L.append('audit_ref: %s' % q((plan.get('audit_ref') or '').replace('\\', '/')))
    if plan.get('notes'):
        L.append('notes: %s' % q(plan['notes']))
    L.append('')
    if not plan['items']:
        L.append('items: []')
        L.append('')
        return '\n'.join(L)
    L.append('items:')
    for it in plan['items']:
        L.append('')
        L.append('  # %s' % (it.get('title') or ''))
        if it.get('location'):
            L.append('  # 位置：%s' % it['location'])
        L.append('  - issue_id: %s' % it['issue_id'])
        L.append('    rule: %s' % (it.get('rule') or ''))
        L.append('    skill: %s' % (it.get('skill') or ''))
        L.append('    action: %s' % it['action'])
        L.append('    title: %s' % q(it.get('title')))
        L.append('    enabled: %s' % ('true' if it.get('enabled') else 'false'))
        if it.get('reason'):
            L.append('    reason: %s' % q(it['reason']))
    L.append('')
    return '\n'.join(L)


def parse_plan_yaml(path):
    """读回 plan.yaml，只取"勾选表"部分。

    items 里的 params 故意不落盘（太长，人也无需看），apply 一律回到
    audit.json 按 issue_id 取 params。这样 yaml 永远只是一张勾选表，
    用户把它改坏了也伤不到执行细节 —— 边界清晰，容错率高。
    """
    doc = read_plan_yaml(path)
    items = doc.get('items') or []
    if not isinstance(items, list):
        raise SystemExit('plan.yaml 的 items 不是列表，文件可能被改坏了：%s' % path)
    return {
        'version': doc.get('version') or 1,
        'docx': doc.get('docx') or '',
        'docx_sha256': doc.get('docx_sha256') or None,
        'generated_at': doc.get('generated_at') or '',
        'items': [it for it in items if isinstance(it, dict)],
    }


def banner(t):
    print('')
    print('=' * 62)
    print(t)
    print('=' * 62)


def now():
    return time.strftime('%Y-%m-%dT%H:%M:%S')


# 所有会产出 Issue 的子 skill 目录名。新增子 skill 时**必须**加进这里，
# 否则它的修复动作会因为"在 audit.json 里找不到 issue_id"而被当成过期条目跳过 ——
# 表面现象是"plan 里勾了但 apply 说找不到"，很难查。
ALL_SKILLS = ('format-audit', CITE, 'text-style', 'structure-length')


def resolve_params_from_audit(docx_path, items):
    """按 issue_id 从 audit.json 回填 action 的 params。

    plan.yaml 里故意不写 params：那是机器的事，写进去人也不会看，
    反而多一处可能被改坏的地方。
    """
    blobs = []
    for skill in ALL_SKILLS:
        b = load_audit(docx_path, skill)
        if b:
            blobs.append(b)
    idx = {}
    for b in blobs:
        for it in b.get('issues', []):
            if it.get('action'):
                idx[it['id']] = (it['action'].get('name'), it['action'].get('params') or {})
    out, missing = [], []
    for it in items:
        iid = it.get('issue_id')
        if iid not in idx:
            missing.append(iid)
            continue
        name, params = idx[iid]
        merged = dict(it)
        merged['action'] = name
        merged['params'] = params
        out.append(merged)
    return out, missing
