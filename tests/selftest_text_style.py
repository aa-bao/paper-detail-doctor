#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
tests/selftest_text_style.py —— text-style 的**注入式**验收

为什么不能只跑一遍干净文档就收工
    "在一份本来就干净的稿子上跑出 0 条问题"完全不能说明检查器是好的——
    一个永远返回空列表的函数也能做到。所以这里反过来做：
    **人为造出每一种缺陷，确认检查器抓得到；并确认可选规则默认不跑、加了开关才跑。**

场景（都在临时副本上，不动你的原稿；绝不用 shutil.copy2 之外的任何方式碰样本）
    0 基线          干净稿 → 只应有真实的 11 条标题冒号 + 2 条破折号密度 info；无 error
    1 引号体例      → quote-style（注入中文弯引号）
    2 标点异常      → punctuation（注入连续标点 / 跨类标点）
    3 破折号密度    → dash-density（注入一堆破折号，触发单段≥3 与整体密度超阈）
    4 标题体例-过长  → heading-style（注入超长三级标题）
    5 标题体例-三级数 → heading-style（在一个二级标题下塞 5 个三级标题）
    6 过程痕迹      → process-trace（注入 TODO）
    7 可选三件套    → 默认不出现；--optional 才出现（cjk-spacing / fullwidth-halfwidth / term-consistency）
    另外：只读性（哈希不变）、问题 ID 稳定、term-consistency 无 yaml 时跳过并说明。

怎么跑
    cd F:\Coding\Project\paper-detail-doctor
    "python" tests/selftest_text_style.py
"""
import importlib.util
import os
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(HERE)
sys.path.insert(0, PKG)
sys.path.insert(0, HERE)   # 让 tests/_sample.py 可被 import

from shared.lib.docx_ops import (                                    # noqa: E402
    Doc, W, insert_paragraph_after, make_locator,
)
from shared.lib.docx_scan import Scanner, norm                       # noqa: E402
from shared.lib.ooxml_guard import sha256_file                       # noqa: E402

from _sample import thesis as _sample_thesis, missing_hint as _missing_hint   # noqa: E402
DEFAULT_SAMPLE = _sample_thesis()
TMP = tempfile.mkdtemp(prefix='pdd-text-selftest-')

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


def load_audit_mod():
    p = os.path.join(PKG, 'skills', 'text-style', 'scripts', 'audit.py')
    spec = importlib.util.spec_from_file_location('text_audit', p)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


AUD = None


def audit(docx, optional=False):
    return AUD.audit(docx, None, verbose=False, optional=optional)['issues']


def rules_of(issues):
    return {i['rule'] for i in issues}


def fresh(src, tag):
    p = os.path.join(TMP, tag + '.docx')
    shutil.copy2(src, p)
    return p


def find_para(docx, pred):
    s = Scanner(docx)
    for i, p in enumerate(s._paras, 1):
        st = (p.style.name if p.style else '') or ''
        if pred(st, p):
            return i
    return None


def style_id(docx, name):
    """显示名 → style_id（insert_paragraph_after 把字符串直接写进 pStyle@w:val，
    必须给 style_id 而非显示名，否则 python-docx 解析不到样式会回退 Normal）。"""
    from docx import Document
    doc = Document(docx)
    for st in doc.styles:
        if st.name == name:
            return st.style_id
    return name


# ---------------------------------------------------------------- 注入器

def _last_body_idx(docx):
    """参考文献标题之前最后一个正文段（注入普通段落用，避免污染结构）。"""
    h, _ = Scanner(docx).reference_section()
    return (h - 2) if h else (len(Scanner(docx)._paras) - 1)


def inject_curved_quote(docx):
    d = Doc(docx)
    idx = _last_body_idx(docx)
    el = d.paragraphs()[idx - 1]._element
    insert_paragraph_after(el, '\u201c\u8fd9\u662f\u4e2d\u6587\u5f2f\u5f15\u53f7\u201d', '正文文本1')
    d.save_atomic()
    return idx + 1


def inject_bad_punct(docx):
    d = Doc(docx)
    idx = _last_body_idx(docx)
    el = d.paragraphs()[idx - 1]._element
    insert_paragraph_after(el, '这是连续句号测试。 。不好的例子，。再试一次', '正文文本1')
    d.save_atomic()
    return idx + 1


def inject_many_dash(docx, n=100):
    d = Doc(docx)
    idx = _last_body_idx(docx)
    el = d.paragraphs()[idx - 1]._element
    insert_paragraph_after(el, AUD.EM_DASH * n + '测试结尾', '正文文本1')
    d.save_atomic()
    return idx + 1


def inject_long_heading(docx):
    d = Doc(docx)
    h3 = find_para(docx, lambda st, p: '三级标题' in st)
    el = d.paragraphs()[h3 - 1]._element
    sid = style_id(docx, '三级标题')
    long = '这是一个明显超过二十五字长度限制的用于测试标题体例规则的非常冗长的章节标题示例'
    insert_paragraph_after(el, long, sid)
    d.save_atomic()
    return h3 + 1


def inject_h2_many_h3(docx, extra=5):
    d = Doc(docx)
    h2 = find_para(docx, lambda st, p: '二级标题' in st)
    el = d.paragraphs()[h2 - 1]._element
    sid = style_id(docx, '三级标题')
    # 始终挂在同一个 h2 元素后面（insert_paragraph_after 用 addnext，元素引用稳定）；
    # 不中途索引 d.paragraphs()（否则会因 _paras 未刷新而拿到过期节点）。
    for k in range(extra):
        insert_paragraph_after(el, '测试三级标题%d' % (k + 1), sid)
    d.save_atomic()
    return h2


def inject_process_trace(docx):
    d = Doc(docx)
    idx = _last_body_idx(docx)
    el = d.paragraphs()[idx - 1]._element
    insert_paragraph_after(el, 'TODO：这一节还需要补充数据', '正文文本1')
    d.save_atomic()
    return idx + 1


def inject_cjk_spacing(docx):
    d = Doc(docx)
    idx = _last_body_idx(docx)
    el = d.paragraphs()[idx - 1]._element
    insert_paragraph_after(el, '中文 iPhone 手机 Android 测试 混排iPhone直接接 结束', '正文文本1')
    d.save_atomic()
    return idx + 1


def inject_fullwidth(docx):
    d = Doc(docx)
    idx = _last_body_idx(docx)
    el = d.paragraphs()[idx - 1]._element
    insert_paragraph_after(el, '版本０１２３与v2并存 测试', '正文文本1')
    d.save_atomic()
    return idx + 1


def inject_term(docx, variant):
    d = Doc(docx)
    idx = _last_body_idx(docx)
    el = d.paragraphs()[idx - 1]._element
    insert_paragraph_after(el, '%s这一概念很重要' % variant, '正文文本1')
    d.save_atomic()
    return idx + 1


def write_temp_term_map(path, mapping):
    import yaml
    with open(path, 'w', encoding='utf-8') as f:
        yaml.safe_dump(mapping, f, allow_unicode=True)


# ---------------------------------------------------------------- 场景

def main():
    src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SAMPLE
    global AUD
    AUD = load_audit_mod()
    if not src or not os.path.exists(src):
        print('样本不存在：%s' % src)
        print(_missing_hint())
        return 2
    print('样本：%s\n临时目录：%s' % (src, TMP))

    # ---------------------------------------------------- 0 基线
    section('0 基线：干净稿上只应有真实告警，无 error')
    base = fresh(src, 'base')
    res = AUD.audit(base, verbose=False)
    iss = res['issues']
    from collections import Counter
    byrule = Counter(i['rule'] for i in iss)
    bysev = Counter(i['severity'] for i in iss)
    print('  基线规则分布：%s' % dict(byrule))
    print('  基线级别分布：%s' % dict(bysev))
    check('基线无 error', not [i for i in iss if i['severity'] == 'error'],
          'errors=%s' % [i['id'] for i in iss if i['severity'] == 'error'])
    check('基线只有 heading-style + dash-density（无其它规则误报）',
          rules_of(iss) <= {'heading-style', 'dash-density'},
          'rules=%s' % rules_of(iss))
    check('基线 heading-style = 11（均为标题冒号，非虚词/过长误报）',
          byrule.get('heading-style') == 11,
          'heading-style=%s' % byrule.get('heading-style'))
    check('基线 heading-style 全部是"含冒号"，无虚词/过长/末尾标点误报',
          all('含冒号/问号/感叹号' in i['actual']
              for i in iss if i['rule'] == 'heading-style'),
          'actuals=%s' % [i['actual'][:30] for i in iss if i['rule'] == 'heading-style'][:3])
    check('基线 dash-density = 2（均为单段≥3 的 info）',
          byrule.get('dash-density') == 2,
          'dash-density=%s' % byrule.get('dash-density'))
    h = sha256_file(base)
    iss2 = AUD.audit(base, verbose=False)['issues']
    check('审计是只读的（哈希不变）', sha256_file(base) == h)
    check('问题 ID 稳定（重跑同 ID）',
          [i['id'] for i in iss] == [i['id'] for i in iss2])
    # 默认（optional=False）不应出现任何可选规则
    check('默认运行不含可选规则（cjk-spacing/fullwidth-halfwidth/term-consistency）',
          not (rules_of(iss) & {'cjk-spacing', 'fullwidth-halfwidth', 'term-consistency'}),
          'rules=%s' % rules_of(iss))

    # ---------------------------------------------------- 1 引号体例
    section('1 注入：中文弯引号 → quote-style')
    f = fresh(src, 'quote')
    pi = inject_curved_quote(f)
    iss = audit(f)
    hit = [i for i in iss if i['rule'] == 'quote-style']
    check('检出 quote-style', bool(hit), 'rules=%s' % rules_of(iss))
    check('定位到注入段 %d' % pi,
          any(i['location']['index'] == pi for i in hit),
          'indices=%s' % [i['location']['index'] for i in hit])
    check('一处只报 1 条（按段落聚合）', len(hit) == 1, 'hit=%d' % len(hit))
    check('无自动动作（文字体例不给盲改）',
          bool(hit) and (hit[0].get('action') or {}).get('name') is None)
    check('standard_source 为 default',
          bool(hit) and hit[0]['standard_source'] == 'default')

    # ---------------------------------------------------- 2 标点异常
    section('2 注入：连续/跨类标点 → punctuation')
    f = fresh(src, 'punct')
    pi = inject_bad_punct(f)
    iss = audit(f)
    hit = [i for i in iss if i['rule'] == 'punctuation']
    check('检出 punctuation', bool(hit), 'rules=%s' % rules_of(iss))
    check('定位到注入段 %d' % pi, any(i['location']['index'] == pi for i in hit))
    check('实际描述点出连续/跨类标点',
          bool(hit) and ('连续标点' in hit[0]['actual']))
    check('无自动动作', bool(hit) and (hit[0].get('action') or {}).get('name') is None)

    # ---------------------------------------------------- 3 破折号密度
    section('3 注入：一堆破折号 → dash-density（单段≥3 + 整体超阈）')
    f = fresh(src, 'dash')
    pi = inject_many_dash(f, 100)
    iss = audit(f)
    hit = [i for i in iss if i['rule'] == 'dash-density']
    check('检出 dash-density', bool(hit), 'rules=%s' % rules_of(iss))
    check('含单段≥3 的 info 条', any(i['severity'] == 'info' for i in hit))
    check('整体密度超阈 → 含 warn 条（列最密集段落）',
          any(i['severity'] == 'warn' for i in hit))
    check('warn 条定位到注入段 %d' % pi,
          any(i['location']['index'] == pi and i['severity'] == 'info' for i in hit))
    check('无自动动作', all((i.get('action') or {}).get('name') is None for i in hit))

    # ---------------------------------------------------- 4 标题过长
    section('4 注入：超长三级标题 → heading-style（过长）')
    f = fresh(src, 'long')
    pi = inject_long_heading(f)
    iss = audit(f)
    hit = [i for i in iss if i['rule'] == 'heading-style'
           and '过长' in i['actual']]
    check('检出 heading-style 过长', bool(hit), 'actuals=%s' % [
          i['actual'][:30] for i in iss if i['rule'] == 'heading-style'][:3])
    check('定位到注入段 %d' % pi, any(i['location']['index'] == pi for i in hit))

    # ---------------------------------------------------- 5 二级下三级数
    section('5 注入：二级标题下塞 5 个三级标题 → heading-style（三级数超限）')
    f = fresh(src, 'h3')
    h2 = inject_h2_many_h3(f, 5)
    iss = audit(f)
    hit = [i for i in iss if i['rule'] == 'heading-style'
           and '三级标题' in i['actual']]
    check('检出 heading-style 三级数超限', bool(hit), 'actuals=%s' % [
          i['actual'][:40] for i in iss if i['rule'] == 'heading-style'][:3])
    check('定位到该二级标题段 %d' % h2, any(i['location']['index'] == h2 for i in hit))
    check('超限条目的 standard_source 来自 max_h3_per_h2（default）',
          bool(hit) and hit[0]['standard_source'] == 'default')

    # ---------------------------------------------------- 6 过程痕迹
    section('6 注入：TODO → process-trace')
    f = fresh(src, 'trace')
    pi = inject_process_trace(f)
    iss = audit(f)
    hit = [i for i in iss if i['rule'] == 'process-trace']
    check('检出 process-trace', bool(hit), 'rules=%s' % rules_of(iss))
    check('定位到注入段 %d' % pi, any(i['location']['index'] == pi for i in hit))
    check('实际点出 TODO', bool(hit) and 'TODO' in hit[0]['actual'])

    # ---------------------------------------------------- 7 可选三件套
    section('7 可选规则：默认不出现，--optional 才出现')
    f = fresh(src, 'opt')
    inject_cjk_spacing(f)
    inject_fullwidth(f)
    inject_term(f, '代际关系')
    inject_term(f, '代间关系')
    # 默认（optional=False）
    iss_def = audit(f, optional=False)
    check('默认运行不报 cjk-spacing', not any(i['rule'] == 'cjk-spacing' for i in iss_def))
    check('默认运行不报 fullwidth-halfwidth',
          not any(i['rule'] == 'fullwidth-halfwidth' for i in iss_def))
    check('默认运行不报 term-consistency',
          not any(i['rule'] == 'term-consistency' for i in iss_def))
    # 打开可选
    iss_opt = audit(f, optional=True)
    check('optional 报 cjk-spacing', any(i['rule'] == 'cjk-spacing' for i in iss_opt),
          'rules=%s' % rules_of(iss_opt))
    check('optional 报 fullwidth-halfwidth',
          any(i['rule'] == 'fullwidth-halfwidth' for i in iss_opt))
    # term-consistency 需要术语表；先用临时表验证可检测
    tmp_yaml = os.path.join(TMP, 'term-map.yaml')
    write_temp_term_map(tmp_yaml, {'代际关系': ['代际关系', '代间关系']})
    saved = AUD.TERM_MAP_PATH
    AUD.TERM_MAP_PATH = tmp_yaml
    iss_term = audit(f, optional=True)
    hit = [i for i in iss_term if i['rule'] == 'term-consistency']
    check('optional + 术语表 → 检出 term-consistency', bool(hit),
          'rules=%s' % rules_of(iss_term))
    check('term-consistency 点出两种写法混用',
          bool(hit) and ('代际关系' in hit[0]['actual'] and '代间关系' in hit[0]['actual']))
    # 还原：无术语表时跳过并说明
    AUD.TERM_MAP_PATH = saved
    res_skip = AUD.audit(f, None, verbose=False, optional=True)
    skip_note = [n for n in res_skip['meta']['notes'] if 'term-consistency' in n]
    check('无术语表时 term-consistency 跳过并在 meta.notes 说明',
          bool(skip_note) and any('跳过' in n for n in skip_note),
          'notes=%s' % skip_note)

    # ---------------------------------------------------- 收尾
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
