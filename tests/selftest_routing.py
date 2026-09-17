#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
tests/selftest_routing.py —— 路由可区分性自检

为什么需要它
    包根 SKILL.md 声称"用户只记一个名字，路由由包负责"。这句话能不能成立，
    取决于 `assets/skills-index.yaml`：**每个 prompt 是否只打中它该打中的那个 skill**。
    索引是手写的，加一条 trigger 就可能把别人的 prompt 抢走 —— 这种回归
    肉眼看不出来，必须机器验。

它验什么、不验什么（说清楚，别夸大）
    ✓ 验：包根 §2 的 12 条示例 prompt，预期 skill 是否得分最高；
      `not_for` 能否挡掉误命中；`out_of_scope` 能否把"本包不干的活"挡在外面；
      以及"两个 skill 咬得很紧"时是否判为歧义（设计上歧义必须交给用户选）。
    ✗ 不验：真实路由是**模型**读索引后做的语义判断。本脚本用字符 bigram 相似度
      近似，只是一道**可区分性下界**。它过了不代表模型一定路由对；
      它不过，则说明索引本身有歧义或缺"不覆盖"条目 —— 那是真问题。

首跑揪出的两个真问题（已修，留作记录）
    ① `template-extract` 的 trigger「页眉要不要横线」与 `format-audit` 的
       「页眉没有横线」语义重叠 —— 两者其实是**问规范** vs **查我的文档** 两种意图，
       索引里没写清语境，于是互相抢 prompt。已在两侧 trigger 补上语境并把反例写进 not_for。
    ② 索引里**没有"不覆盖"清单**，导致"帮我写个摘要""帮我查重"这类请求
       也会被硬塞给某个 skill（短 prompt 跟任何 trigger 都有零点几的相似度）。
       已补 `out_of_scope` 段。

用法
    "python" tests/selftest_routing.py
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(HERE)
sys.path.insert(0, PKG)

import yaml                                                          # noqa: E402

INDEX = os.path.join(PKG, 'assets', 'skills-index.yaml')
NOT_COVERED = '(不覆盖)'
AMBIGUOUS = '(歧义→问用户)'

# 分数低于这个值就当"没命中任何 skill"。取值偏低是刻意的：
# 本匹配器只是下界，宁可少判"不覆盖"，也不要把它变成假阳性来源。
LOW_SCORE = 0.12
# 头两名差距小于这个值 → 判为歧义（设计上必须交用户选，不许自己拍板）。
TIE_MARGIN = 0.08

# 包根 SKILL.md §2「典型 prompt 的路由结果」表里的 12 条，原样抄过来。
# 期望值来自**人工设计意图**，不是从匹配器输出反推的。
CASES = [
    ('引用序号放右上角，能 Ctrl 点着跳过去', 'cite-doctor'),
    ('参考文献的顺序好像不是按正文出现的', 'cite-doctor'),
    ('我这章的破折号是不是太多了', 'text-style'),
    ('第二章标题太玄乎了，改平实点', 'text-style'),
    ('页眉是不是该有根横线', 'format-audit'),
    ('字数超出要求了，帮我压到 3 万 5', 'structure-length'),
    ('开题说要访谈 30 人，我最后只访了 15 人，会不会被问', 'proposal-consistency'),
    ('这条引语是我受访者的原话吗', 'quote-verify'),
    ('学校给了份格式规范 PDF，帮我做成模板', 'template-extract'),
    ('帮我写个摘要', NOT_COVERED),
    ('帮我查重', NOT_COVERED),
]

# 允许判为"歧义"的用例：这两个意图本来就咬得很紧，判歧义反而是**正确**行为。
# 但只允许在给定的候选集合内并列，落到第三个 skill 上仍算失败。
AMBIGUOUS_OK = {
    '页眉是不是该有根横线': {'format-audit', 'template-extract'},
}


# 纯功能词 bigram —— 在**语料上**极常见、在**索引里**却罕见，于是 IDF 反而给它们
# 高权重，造成"是不是""帮我"这种片段成为强证据。必须显式降为零。
# 这份清单是通用的中文功能词，**不是**针对某条测试样本挑的。
STOP = {
    '是不', '不是', '我这', '我的', '帮我', '我想', '我要', '有没', '没有',
    '怎么', '什么', '为什', '可以', '需要', '要是', '不要', '一下', '一点',
    '这个', '那个', '这些', '那些', '还有', '或者', '但是', '如果', '因为',
    '所以', '就是', '已经', '应该', '可能', '知道', '看看', '弄弄', '改改',
    '得对', '对不对', '吗', '呢', '吧', '啊',
}


def bigrams(s):
    s = re.sub(r'\s+', '', s)
    return {s[i:i + 2] for i in range(len(s) - 1)} or {s}


def dice(a, b):
    """字符 bigram 的 Dice 相似度（未加权，保留作对照）。"""
    A, B = bigrams(a), bigrams(b)
    if not A or not B:
        return 0.0
    return 2.0 * len(A & B) / (len(A) + len(B))


class Matcher:
    def __init__(self, index):
        self.skills = [s for s in index.get('skills', []) if s.get('id')]
        self.oos = index.get('out_of_scope') or []
        self.idf = self._build_idf(index)

    # ---------------------------------------------------------------- 权重
    def _build_idf(self, index):
        """索引内的 IDF：在**本索引的 trigger 集合**里统计每个 bigram 的文档频率。

        为什么需要：纯 bigram 会把「是不是」「我的」这种通用片段算成强证据 ——
        首跑时就出现了「页眉是不是该有根横线」被判成与 `quote-verify` 咬紧，
        唯一原因是两者都含「是不是」，而两者的实体词（页眉/横线 vs 引语/原话）毫无关系。
        通用片段降权、实体词升权，这是标准做法；**不是**为了凑测试结果 —— 
        它同时会削弱所有"靠虚词蹭分"的误命中。
        """
        import math
        docs = []
        for s in index.get('skills', []) or []:
            docs += [re.sub(r'#.*$', '', str(t)).strip() for t in (s.get('triggers') or [])]
            docs += [re.sub(r'#.*$', '', str(t)).strip() for t in (s.get('not_for') or [])]
        for e in index.get('out_of_scope') or []:
            docs += [v.strip() for v in re.split(r'[/／]', str(e.get('prompt', '')))]
        docs = [d for d in docs if d]
        df = {}
        for d in docs:
            for g in bigrams(d):
                df[g] = df.get(g, 0) + 1
        n = max(len(docs), 1)
        self._max_idf = math.log((n + 1) / 0.5)
        return {g: math.log((n + 1) / (c + 0.5)) for g, c in df.items()}

    def w(self, g):
        if g in STOP:           # 功能词一律零权重
            return 0.0
        return self.idf.get(g, self._max_idf)

    def wdice(self, a, b):
        """加权 Dice：通用片段权重低，实体词权重高。"""
        A, B = bigrams(a), bigrams(b)
        if not A or not B:
            return 0.0
        inter = sum(self.w(g) for g in A & B)
        denom = sum(self.w(g) for g in A) + sum(self.w(g) for g in B)
        return 2.0 * inter / denom if denom else 0.0

    # ---------------------------------------------------------------- 打分
    def score(self, prompt, skill):
        best, which = 0.0, None
        for t in skill.get('triggers') or []:
            t = re.sub(r'#.*$', '', str(t)).strip()
            d = self.wdice(prompt, t)
            if d > best:
                best, which = d, t
        return best, which

    def blocked_by(self, prompt, skill):
        """not_for 里有没有哪条比本次最佳 trigger 更像这个 prompt。"""
        best_t, _ = self.score(prompt, skill)
        for nf in skill.get('not_for') or []:
            nf = re.sub(r'#.*$', '', str(nf)).strip()
            if nf and self.wdice(prompt, nf) >= max(best_t, 0.30):
                return nf
        return None

    def oos_hit(self, prompt):
        """不覆盖清单里，有没有比任何 skill 都更像这个 prompt 的条目。"""
        best, which = 0.0, None
        for e in self.oos:
            for variant in re.split(r'[/／]', str(e.get('prompt', ''))):
                v = variant.strip()
                if not v:
                    continue
                d = self.wdice(prompt, v)
                if d > best:
                    best, which = d, v
        return best, which

    # ---------------------------------------------------------------- 路由
    def route(self, prompt):
        scored = []
        for s in self.skills:
            v, trig = self.score(prompt, s)
            scored.append({'id': s['id'], 'score': v, 'trigger': trig,
                           'blocked': self.blocked_by(prompt, s)})
        scored.sort(key=lambda x: -x['score'])

        # 1) 先看不覆盖清单
        oos_v, oos_t = self.oos_hit(prompt)
        if oos_v > scored[0]['score']:
            return NOT_COVERED, scored, '命中不覆盖清单「%s」' % oos_t

        # 2) 全被 not_for 挡掉 → 不覆盖
        alive = [x for x in scored if not x['blocked']]
        if not alive:
            return NOT_COVERED, scored, '全部被 not_for 挡掉'

        # 3) 分数太低 → 不覆盖（不许硬凑）
        if alive[0]['score'] < LOW_SCORE:
            return NOT_COVERED, scored, '最高分 %.3f 低于阈值 %.2f' % (alive[0]['score'], LOW_SCORE)

        # 4) 头两名咬得紧 → 歧义，必须问用户
        if len(alive) > 1 and (alive[0]['score'] - alive[1]['score']) < TIE_MARGIN:
            return AMBIGUOUS, scored, '与 %s 咬紧（%.3f vs %.3f）' % (
                alive[1]['id'], alive[0]['score'], alive[1]['score'])
        return alive[0]['id'], scored, None


def main():
    if not os.path.exists(INDEX):
        print('索引不存在：%s' % INDEX)
        return 2
    idx = yaml.safe_load(open(INDEX, encoding='utf-8')) or {}
    m = Matcher(idx)
    print('子 skill（索引登记 %d 个）：' % len(m.skills))
    for s in m.skills:
        print('   %-22s triggers=%-2d not_for=%d'
              % (s['id'], len(s.get('triggers') or []), len(s.get('not_for') or [])))
    print('不覆盖清单（out_of_scope）：%d 条' % len(m.oos))
    print()

    ok, fail, notes = 0, 0, []
    print('%-6s %-30s %-20s %-18s %s' % ('结果', 'prompt', '预期', '实得', '依据'))
    print('-' * 108)
    for prompt, want in CASES:
        got, scored, why = m.route(prompt)
        top = scored[0]
        detail = 'top=%s %.3f' % (top['id'], top['score'])
        if got == want:
            verdict = '[ok]'
            ok += 1
        elif got == AMBIGUOUS and want in AMBIGUOUS_OK.get(prompt, set()):
            # 设计上"歧义必须问用户"，所以判歧义是允许的结果；但候选必须在允许集合内
            candidates = {x['id'] for x in scored[:2]}
            if candidates <= AMBIGUOUS_OK.get(prompt, set()):
                verdict = '[注]'
                ok += 1
                notes.append('%s :: 判为歧义交用户选（设计允许，候选 %s）' % (prompt, '/'.join(sorted(candidates))))
            else:
                verdict = '[FAIL]'
                fail += 1
                notes.append('%s :: 歧义候选超出允许集合 %s' % (prompt, sorted(candidates)))
        else:
            verdict = '[FAIL]'
            fail += 1
            notes.append('%s :: 预期 %s，实得 %s（%s）' % (prompt, want, got, detail))
        print('%-6s %-30s %-20s %-18s %s'
              % (verdict, prompt[:28], want, got, why or detail))

    # 单独验一条硬要求：not_for 必须能挡住"引语"落到 cite-doctor
    print()
    p = '这条引语是我受访者的原话吗'
    c = next(x for x in m.skills if x['id'] == 'cite-doctor')
    blocked = m.blocked_by(p, c)
    if blocked:
        print('  [ok]   not_for 生效：cite-doctor 被「%s」挡掉' % blocked)
        ok += 1
    else:
        print('  [FAIL] not_for 未能挡住 cite-doctor')
        fail += 1
        notes.append('not_for 未挡住：%s' % p)

    # 整体性请求不计入通过/失败（本匹配器表达不了"进工作流模式"这个约定），
    # 但要把它打出来，免得这条覆盖在重构时被悄悄丢掉。
    print()
    p2 = '定稿前帮我从头查一遍'
    got2, scored2, why2 = m.route(p2)
    print('  [注]   整体性请求「%s」→ 本包约定进**工作流模式**（跑 audit_all 全套），'
          '而非单点路由；匹配器只能给到 %s（%.3f）'
          % (p2, got2, scored2[0]['score']))

    print()
    print('=' * 70)
    print('通过 %d 项，失败 %d 项' % (ok, fail))
    for n in notes:
        print('  - ' + n)
    print()
    print('说明：本脚本用字符 bigram 近似语义，只是**可区分性下界**；')
    print('      真实路由由模型读索引完成，需在装好包的新会话里实测。')
    return 0 if fail == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
