# -*- coding: utf-8 -*-
r"""测试样本（论文正文 / 开题报告 docx）的解析。

这个包是**通用工具，不绑定任何一篇具体的论文**。所以测试不把任何人的本机路径
写死在代码里，而是按下面的顺序解析：

    1. 命令行参数（各测试脚本自己处理，优先级最高）
    2. 环境变量  PDD_SAMPLE / PDD_PROPOSAL
    3. tests/sample.local.yaml —— 本机私有配置，**不入库**
       （模板见同目录 sample.local.example.yaml）
    4. 都没有 -> 返回 None，测试会明确报「样本未提供」并按约定以退出码 2 跳过，
       绝不假装通过

为什么不做成"仓库里放一份样例文档"
    规则类测试（cite-doctor / format-audit / text-style / structure-length）
    都是**注入式**的：在干净稿上人为造缺陷，看能不能查出来、改完能不能消失。
    这需要一个结构完整、能被 ooxml 正常改写的真实 docx，合成一份成本高且容易
    掩盖真实文档的怪癖（合并单元格、域代码、样式继承）。所以本包的选择是：
    测试**依赖本机样本**，但样本路径绝不进仓库 —— 别人 clone 后按
    sample.local.example.yaml 指到自己的文档即可，规则本身完全通用。
"""
import os

HERE = os.path.dirname(os.path.abspath(__file__))
LOCAL = os.path.join(HERE, 'sample.local.yaml')

# 本机私有配置的键名
KEY_THESIS = 'thesis'
KEY_PROPOSAL = 'proposal'

_ENV_THESIS = 'PDD_SAMPLE'
_ENV_PROPOSAL = 'PDD_PROPOSAL'


def _from_local(key):
    """从 tests/sample.local.yaml 读一个路径；读不到就返回 None。"""
    if not os.path.exists(LOCAL):
        return None
    try:
        import yaml
    except ImportError:
        return None
    try:
        with open(LOCAL, encoding='utf-8') as f:
            data = yaml.safe_load(f) or {}
    except Exception:
        return None
    v = data.get(key)
    if isinstance(v, str) and v.strip():
        return v.strip()
    return None


def _resolve(env, key):
    return os.environ.get(env) or _from_local(key)


def thesis():
    """论文正文 docx 的路径；未配置返回 None。"""
    return _resolve(_ENV_THESIS, KEY_THESIS)


def proposal():
    """开题报告 docx 的路径；未配置返回 None。"""
    return _resolve(_ENV_PROPOSAL, KEY_PROPOSAL)


def missing_hint(what='样本'):
    """样本缺失时统一打印的提示，让用户知道怎么配。"""
    return ('%s未提供。请任选其一：\n'
            '  · 设环境变量  %s / %s\n'
            '  · 或复制 tests/sample.local.example.yaml 为 tests/sample.local.yaml 并填写路径\n'
            '  · 或直接作为命令行参数传入' % (what, _ENV_THESIS, _ENV_PROPOSAL))
