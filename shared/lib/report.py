#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
shared/lib/report.py —— 统一的问题渲染（子 skill 不自己产报告）

分工
    子 skill 只返回 list[Issue]（结构见 shared/contracts/issue.schema.json）。
    本模块负责把 7 份 Issue 列表合成 **一份** 给人看的 Markdown + 给机器用的 JSON。

为什么要统一
    否则 audit_all 要写 7 套解析逻辑，报告里也会出现 7 种排版风格。

三个必须做到的事
    1. **每条都带定位**（段落号 + 上下文片段）。没有定位的报告没法复核。
    2. **每条都标标准来源**（模板 / 配置 / 默认）。让人一眼分清
       「学校就这么要求的」和「工具的建议」—— 这是最容易变成噪声的地方。
    3. **被人工裁定保留的条目单独列出**，不要静默丢弃。用户要能看到
       "有 3 条我按 .thesisignore 跳过了"。
"""
import json
import os
import re

SEV_ORDER = {'error': 0, 'warn': 1, 'info': 2}
SEV_LABEL = {'error': '错误', 'warn': '警告', 'info': '提示'}
SRC_LABEL = {'template': '模板', 'config': '配置', 'default': '默认'}
SRC_FLAG = {'template': '✅', 'config': '🔧', 'default': '⚠️'}


# ---------------------------------------------------------------- .thesisignore

def load_ignore(*paths):
    """读取 .thesisignore（按 id 或 rule 级别裁定保留）。返回 (ids:set, rules:set)。

    格式（YAML 子集，容错解析，不强依赖 pyyaml）：
        - id: cite-position-0042
          reason: 导师要求此处用文中注
        - rule: dash-density
          reason: 本篇是文艺学论文
    """
    ids, rules = set(), set()
    for p in paths:
        if not p or not os.path.exists(p):
            continue
        cur = None
        for line in open(p, encoding='utf-8'):
            s = line.strip()
            if not s or s.startswith('#'):
                continue
            m = re.match(r'^-?\s*id:\s*["\']?([\w-]+)', s)
            if m:
                ids.add(m.group(1))
                cur = None
                continue
            m = re.match(r'^-?\s*rule:\s*["\']?([\w-]+)', s)
            if m:
                rules.add(m.group(1))
                cur = None
                continue
            m = re.match(r'^reason:', s)
            if m:
                cur = 'reason'
                continue
    return ids, rules


def apply_ignore(issues, ignore_ids, ignore_rules):
    """把被裁定保留的条目标记出来（**不删除**）。返回 (kept, suppressed)。"""
    kept, sup = [], []
    for it in issues:
        if it.get('suppressible') and (it.get('id') in ignore_ids or it.get('rule') in ignore_rules):
            it = dict(it)
            it['_suppressed'] = True
            sup.append(it)
        else:
            kept.append(it)
    return kept, sup


# ---------------------------------------------------------------- 渲染

def _loc(it):
    loc = it.get('location') or {}
    kind = loc.get('kind') or 'document'
    idx = loc.get('index')
    ex = loc.get('excerpt')
    bits = []
    if kind == 'paragraph' and idx:
        bits.append('段 %s' % idx)
    elif kind != 'document':
        bits.append(kind)
    if ex:
        bits.append('「%s」' % ex)
    return ' · '.join(bits) or '全文'


def render_markdown(issues, suppressed, summary, meta=None):
    meta = meta or {}
    lines = []
    lines.append('# 论文体检报告')
    lines.append('')
    lines.append('| 项 | 值 |')
    lines.append('|---|---|')
    lines.append('| 目标文档 | `%s` |' % (summary.get('path') or '-'))
    lines.append('| 段落 / 表格 / 图片 | %s / %s / %s |' % (
        summary.get('paragraphs'), summary.get('tables'), summary.get('images')))
    lines.append('| 分节 | %s |' % summary.get('sections'))
    lines.append('| 正文引注 | %s 处 |' % summary.get('citations'))
    lines.append('| 参考文献 | %s 条 |' % summary.get('reference_entries'))
    if meta.get('standard_source_note'):
        lines.append('| 标准来源 | %s |' % meta['standard_source_note'])
    lines.append('')

    if not issues:
        lines.append('## 结果：全部通过')
        lines.append('')
        lines.append('本次审计未发现问题。')
    else:
        by_sev = {}
        for it in issues:
            by_sev.setdefault(it.get('severity', 'info'), []).append(it)
        lines.append('## 结果：%d 条问题' % len(issues))
        lines.append('')
        lines.append('| 级别 | 条数 |')
        lines.append('|---|---|')
        for sev in ('error', 'warn', 'info'):
            if by_sev.get(sev):
                lines.append('| %s | %d |' % (SEV_LABEL[sev], len(by_sev[sev])))
        lines.append('')

        # 按 子skill → 规则 分组，便于逐项处理
        for skill in sorted({it.get('skill', '-') for it in issues}):
            group = [it for it in issues if it.get('skill', '-') == skill]
            group.sort(key=lambda x: (SEV_ORDER.get(x.get('severity'), 9),
                                      x.get('location', {}).get('index') or 0))
            lines.append('### %s（%d 条）' % (skill, len(group)))
            lines.append('')
            for it in group:
                lines.append('**[%s] %s**　`%s`　%s' % (
                    SEV_LABEL.get(it.get('severity'), it.get('severity')),
                    it.get('rule'), it.get('id'), _loc(it)))
                lines.append('')
                lines.append('- 期望：%s' % it.get('expected'))
                lines.append('- 实际：%s' % it.get('actual'))
                lines.append('- 依据：%s %s' % (
                    SRC_FLAG.get(it.get('standard_source'), ''),
                    SRC_LABEL.get(it.get('standard_source'), it.get('standard_source'))))
                lines.append('- 建议：%s' % it.get('suggestion'))
                if it.get('action'):
                    lines.append('- 可自动修复：%s' % it['action'].get('name'))
                else:
                    lines.append('- 需人工修改（无可执行动作）')
                lines.append('')

    if suppressed:
        lines.append('---')
        lines.append('')
        lines.append('## 已按 `.thesisignore` 裁定保留（%d 条，未作处理）' % len(suppressed))
        lines.append('')
        for it in suppressed:
            lines.append('- `%s` %s（%s）' % (it.get('id'), it.get('rule'), _loc(it)))
        lines.append('')

    lines.append('---')
    lines.append('')
    lines.append('> 本报告由只读审计生成，**未修改任何文档**。要修，请走 '
                 '`workflow/plan.py` → 勾选 → `workflow/apply.py`。')
    return '\n'.join(lines)


def render_json(issues, suppressed, summary, meta=None):
    return {
        'meta': dict(meta or {}, summary=summary),
        'counts': {
            'total': len(issues),
            'suppressed': len(suppressed),
            'by_severity': _tally(it.get('severity') for it in issues),
            'by_skill': _tally(it.get('skill') for it in issues),
            'by_standard_source': _tally(it.get('standard_source') for it in issues),
        },
        'issues': issues,
        'suppressed': suppressed,
    }


def write(out_dir, issues, suppressed, summary, meta=None):
    os.makedirs(out_dir, exist_ok=True)
    md = os.path.join(out_dir, 'audit-report.md')
    js = os.path.join(out_dir, 'audit.json')
    with open(md, 'w', encoding='utf-8') as f:
        f.write(render_markdown(issues, suppressed, summary, meta))
    with open(js, 'w', encoding='utf-8') as f:
        json.dump(render_json(issues, suppressed, summary, meta),
                  f, ensure_ascii=False, indent=2)
    return md, js


def validate_issue(it):
    """基本字段完整性检查（schema 的轻量替代，不引第三方依赖）。"""
    need = ['id', 'rule', 'layer', 'skill', 'severity', 'location',
            'expected', 'actual', 'standard_source', 'suggestion', 'suppressible']
    missing = [k for k in need if k not in it]
    if it.get('standard_source') not in ('template', 'config', 'default'):
        missing.append('standard_source(取值非法)')
    if it.get('severity') not in SEV_ORDER:
        missing.append('severity(取值非法)')
    return missing


def _tally(it):
    out = {}
    for x in it:
        out[x] = out.get(x, 0) + 1
    return out
