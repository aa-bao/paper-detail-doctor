#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
shared/lib/standards.py —— 标准来源解析（三个来源 → 一个扁平口径）

为什么必须抽到 shared
    每个子 skill 都要回答同一个问题："这条判定依据的是学校模板、用户配置，还是
    通用惯例？" 如果各写各的，同一个字段会出现不同的默认值、不同的来源标注口径，
    报告里就再也说不清"到底谁规定的"。所以这里是**唯一**的解析入口。

三级优先
    template  —— 模板/范文 spec（最高）
    config    —— 用户配置覆盖
    default   —— 通用学术惯例（报告里必须显著标出，别让人误以为学校就这么要求）

一条重要的诚实纪律
    template-extract 只能"**测量**范文实际是什么样"（比如范文摘要 506 字），
    它**不知道**学校的上限是多少。所以：
      · spec.conventions.abstract.cn.chars = 506  → 这是实测值，不是上限，
        不能拿来当"超过 506 字就报警"，否则所有比范文长一点的摘要都被误报；
      · 真要有上限，得在 spec 里显式写 `abstract.cn.max_chars`（从格式规范文档
        抄来的硬指标）。
    同理适用于页边距、字数配额等一切"范文有值但规范未必规定"的字段。
    本模块对这类字段的处理是：**默认值兜底 + 把实测值作为上下文写进 Issue**，
    并在 standard_source 里如实标注 default。

用法
    from shared.lib.standards import find_spec, load_spec, resolve_standard
    spec_path = find_spec(docx)
    spec, spec_path = load_spec(spec_path)
    std, src, notes = resolve_standard(spec)
    std['cite_style']        # 值
    src['cite_style']        # 'template' | 'config' | 'default'
"""
import json
import os

# ---------------------------------------------------------------- 默认口径
#
# 只放"不写就没法判"的项。凡是"没标准就不该判"的（如页边距、正文字号），
# 一律留 None —— 宁可少报，不可瞎报。报告里会说明"因缺标准未判"。
DEFAULTS = {
    # ---- L2 引注与文献 ----
    'cite_style': 'superscript-bracket',   # superscript-bracket | superscript | bracket | author-year
    'cite_position': 'before-punct',       # before-punct | in-text
    'cite_allow_explicit_size': False,     # 上标不该再写 w:sz（会被渲染器缩两次）
    'bib_system': 'sequence',              # sequence（顺序编码制）| author-year
    'bib_standard': 'GB/T 7714',

    # ---- L1 版式 ----
    'page_size': 'A4',
    'page_margins_cm': None,               # 无标准不判
    'page_numbering_scheme': None,         # 无标准不判（实测值在 spec 里，是"范文长这样"）
    'header_content': None,                # 无标准不判
    'header_rule': None,                   # 无标准不判（范文有横线 ≠ 学校要求有横线）
    'body_font_cn': None,
    'body_size_pt': None,
    'body_line_spacing': None,

    # ---- L3 文字体例 ----
    'quote_style': 'straight',             # straight（本项目统一直引号 U+0022/U+0027）
    'dash_per_1000_max': 6.0,
    'heading_max_chars': 25,
    'max_h3_per_h2': 3,

    # ---- L4 结构与篇幅 ----
    'abstract_cn_max_chars': 1000,
    'abstract_en_max_words': 500,
    'caption_figure_position': 'below',    # 图题在下
    'caption_table_position': 'above',     # 表题在上
    'body_min_chars': None,                # 无标准不判（配额要靠用户显式给）
}

# spec.conventions 里的路径 → 扁平键
# 只映射"确实能当标准用"的项；纯测量值（页码起止、摘要实测字数、标题实测值）
# 不映射，避免被误当上限。
CONV_MAP = {
    'citation.style': 'cite_style',
    'citation.position': 'cite_position',
    'bibliography.system': 'bib_system',
    'bibliography.standard': 'bib_standard',
    # ★ 显式上限（从格式规范文档抄来的硬指标），区别于 extract 测出来的实测值
    'abstract.cn.max_chars': 'abstract_cn_max_chars',
    'abstract.cn.max_words': 'abstract_cn_max_words',
    'abstract.en.max_words': 'abstract_en_max_words',
    'caption.figure.position': 'caption_figure_position',
    'caption.table.position': 'caption_table_position',
    'text.dash_per_1000_max': 'dash_per_1000_max',
    'text.heading_max_chars': 'heading_max_chars',
    'text.quote_style': 'quote_style',
    'page.size': 'page_size',
    'page.margins_cm': 'page_margins_cm',
    'header.content': 'header_content',
    'header.rule': 'header_rule',
    'page_numbering.scheme': 'page_numbering_scheme',
    'length.body_min_chars': 'body_min_chars',
}

# 中文名（报告与提示语用）
CN_NAME = {
    'cite_style': '引注形态', 'cite_position': '引注位置',
    'cite_allow_explicit_size': '上标是否显式设字号',
    'bib_system': '参考文献编码制', 'bib_standard': '著录标准',
    'page_size': '纸张', 'page_margins_cm': '页边距',
    'page_numbering_scheme': '页码方案', 'header_content': '页眉内容',
    'header_rule': '页眉横线', 'body_font_cn': '中文正文字体',
    'body_size_pt': '正文字号', 'body_line_spacing': '正文行距',
    'quote_style': '引号体例', 'dash_per_1000_max': '破折号密度上限（每千字）',
    'heading_max_chars': '标题长度上限（字）', 'max_h3_per_h2': '每个二级标题下三级标题数上限',
    'abstract_cn_max_chars': '中文摘要字数上限', 'abstract_en_max_words': '英文摘要词数上限',
    'caption_figure_position': '图题位置', 'caption_table_position': '表题位置',
    'body_min_chars': '正文字数下限',
}

SPEC_FILENAMES = ('paper-spec.json', 'spec.json',
                  os.path.join('.thesis-doctor', 'spec.json'),
                  os.path.join('.paper-doctor', 'spec.json'))


def find_spec(docx_path):
    """在文档旁边按约定位置找 spec.json；找不到返回 None（走通用惯例）。"""
    d = os.path.dirname(os.path.abspath(docx_path))
    for name in SPEC_FILENAMES:
        p = os.path.join(d, name)
        if os.path.exists(p):
            return p
    return None


def load_spec(path):
    """返回 (spec_dict|None, abs_path|None)。路径为空或不存在时返回 (None, None)。

    刻意不 raise：没有 spec 是**正常情况**（用户还没跑 template-extract），
    各子 skill 应当据此把 standard_source 全标 default 并说明，而不是直接失败。
    """
    if not path:
        return None, None
    if not os.path.exists(path):
        return None, None
    with open(path, encoding='utf-8') as f:
        return json.load(f), os.path.abspath(path)


def _dig(d, path):
    cur = d
    for k in path.split('.'):
        if not isinstance(cur, dict):
            return None
        cur = cur.get(k)
        if cur is None:
            return None
    return cur


def resolve_standard(spec, config=None):
    """解析出扁平口径。

    返回 (std: dict, src: dict, notes: list[str])
        std[key]  最终取值
        src[key]  'template' | 'config' | 'default'
        notes     需要在报告里如实说明的话（尤其是"某项因缺标准未判"）
    """
    std, src, notes = dict(DEFAULTS), {k: 'default' for k in DEFAULTS}, []
    conv = ((spec or {}).get('conventions') or {})
    cfg = ((config or {}).get('conventions') or {}) if config else {}

    # ① template
    for path, key in CONV_MAP.items():
        v = _dig(conv, path)
        if v not in (None, '', [], {}):
            std[key] = v
            src[key] = 'template'

    # ② config 覆盖 template
    for path, key in CONV_MAP.items():
        v = _dig(cfg, path)
        if v not in (None, '', [], {}):
            std[key] = v
            src[key] = 'config'

    # ③ 引注显式字号的特殊处理
    #    范文没有显式 w:sz = 作者就没显式设 → 期望状态就是"不设"，这是可靠推断，
    #    所以 source 可以标 template（有 spec 时）。
    cit_style = _dig(conv, 'citation.style')
    if cit_style:
        std['cite_allow_explicit_size'] = bool((_dig(conv, 'citation.superscript_size_ratio')
                                               is not None))
        src['cite_allow_explicit_size'] = 'template'

    if not spec:
        notes.insert(0, '**未提供模板 spec**：本次全部判定依据为通用学术惯例（default），'
                        '不代表你学校的格式要求。要按学校标准判，请先跑 template-extract。')

    # ④ 如实交代"因缺标准未判"的项
    missing = [CN_NAME[k] for k, v in std.items()
               if v is None and k in CN_NAME]
    if missing:
        notes.append('以下项**因未提供标准而未作判定**（不是"没问题"，是"没法判"）：'
                     + '、'.join(missing) + '。')
    return std, src, notes


def describe(std, src, keys=None):
    """一行话概括标准来源，供报告表头用。"""
    label = {'template': '模板', 'config': '配置', 'default': '默认'}
    keys = keys or [k for k in ('cite_style', 'cite_position', 'bib_system',
                                'bib_standard') if k in std]
    return '；'.join('%s=%s(%s)' % (CN_NAME.get(k, k), std[k], label.get(src.get(k), '?'))
                     for k in keys)
