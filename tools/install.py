#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
tools/install.py —— 把本包安装到 ~/.workbuddy/skills/（**只注册包入口**）

为什么是"只注册包入口"（方案乙）
    宿主会**递归发现任意深度的 SKILL.md**。已验证：那个结构坏掉的 `academic-paper`
    包把子 skill 递归复制了 5 层，照样一个个被注册出来。
    所以如果整包拷进 skills 目录，会同时注册 **1 个包入口 + 7 个子 skill = 8 个条目**，
    而本包的设计初衷恰恰相反 —— 用户只记一个名字 `/paper-detail-doctor`，
    由包根按 prompt 路由。

    做法：skills 目录下**只放一个入口存根** `paper-detail-doctor/SKILL.md`，
    它把真正的索引与编排指向包目录；7 个子 skill 的 SKILL.md 全留在包内
    （包不在 skills 目录下 → 宿主扫不到 → 不注册）。

为什么存根写得这么"薄"
    薄才不会有第二份真源。存根只记录「包在哪、先去读什么、三条硬约束、几个入口命令」，
    **不复制**路由表与检查清单 —— 那些改了要同步两处，必然漂移。

用法
    python tools/install.py              # 安装（或更新存根）
    python tools/install.py --check      # 只体检：存根在不在、有没有意外的 SKILL.md
    python tools/install.py --uninstall  # 卸载（存根先备份再删）
"""
import argparse
import os
import shutil
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(HERE)
SKILL_NAME = 'paper-detail-doctor'
SKILLS_DIR = os.path.join(os.path.expanduser('~'), '.workbuddy', 'skills')
TARGET_DIR = os.path.join(SKILLS_DIR, SKILL_NAME)
TARGET_SKILL = os.path.join(TARGET_DIR, 'SKILL.md')

# 存根里要写死的版本号（与包根 SKILL.md 的 metadata.version 保持一致）
VERSION = '0.3.0'

STUB = '''---
name: {name}
description: "论文格式与细节优化的 skill 包入口，只做路由与编排，自身不执行任何检查。内含 7 个专职子 skill：template-extract（模板/范文→Word 模板与格式 spec）、format-audit（L1 格式与版式）、cite-doctor（L2 引注与文献）、text-style（L3 文字体例）、structure-length（L4 结构与篇幅）、proposal-consistency（开题—正文一致性核对）、quote-verify（引语逐字校验，可选）。当用户要检查或修正一份 Word 论文的格式、版式、引注上标与跳转、参考文献、页眉页码、目录、摘要、字数、图表题注、或核对开题与正文是否一致时使用。Triggers: 论文格式检查, 论文排版, 引注上标, 文献引用跳转, 参考文献顺序, 论文细节, 论文体检, 目录页码, 页眉横线, 字数超出, 摘要超页, 破折号太多, 标题太虚, 开题报告和正文一致吗, paper format check, thesis formatting, citation check, thesis template."
metadata:
  version: "{version}"
  kind: entry-stub
  package_root: "{pkg}"
  python_exe: "{python_exe}"
---

# paper-detail-doctor —— 入口存根

⚠️ 本文件只是**入口**。真正的索引、路由程序、共享契约、工作流编排都在包目录里：

    {pkg}\\SKILL.md

## 你（宿主 agent）要做的第一件事

**读上面这个文件**，然后按它的 **§2「路由程序」** 执行。它包含：
7 个子 skill 的索引、六步路由、共享契约、工作流编排、反模式清单。

**不要凭本存根直接执行检查** —— 这里没有检查项明细，明细在子 skill 的 SKILL.md 里。

## 三条不可遗忘的硬约束（摘自包内 SKILL.md）

1. **docx 是唯一真值源**：只对现有 docx 做定点手术，**绝不重新生成文档**。
   一旦稿子有过人工修改，就不要再跑建通道（md→docx），否则会覆盖人工修改。
2. **先声明路由再执行**：必须明文写出 `→ 路由到 <子skill>（理由）`，让用户能当场纠正。
   有歧义时用 AskUserQuestion 问用户，不要自己拍板；都不匹配就直说"本包不覆盖"。
3. **诊断与修改强制分离**：`audit`（只读）→ `plan`（可编辑 yaml，用户勾选）→ `apply`（先快照 · 幂等 · 可回滚）。
   **永远不要跳过 audit 直接改文档。**

## 关键路径与入口命令（包内）

用 python 解释器：`{python_exe}`
（脚本依赖 lxml / python-docx / pyyaml，上面这个解释器已确认可跑）

| 用途 | 命令 |
|---|---|
| 路由唯一真源 | 读 `{pkg}\\assets\\skills-index.yaml` |
| 全项体检（只读） | `{python_exe} "{pkg}\\workflow\\audit_all.py" -d "论文.docx" [-s spec.json] [-p "开题报告.docx"]` |
| 生成改稿计划 | `{python_exe} "{pkg}\\workflow\\plan.py" -d "论文.docx"` → 打开文档旁 `.thesis-doctor/plan.yaml` 勾选 |
| 执行（写盘） | 先 `... "workflow\\apply.py" -p "…\\.thesis-doctor\\plan.yaml" --dry-run`，确认后把 `--dry-run` 换成 `--yes` |
| 回滚 | `{python_exe} "{pkg}\\workflow\\apply.py" -d "论文.docx" --rollback` |

产物固定落在**文档旁**的 `.thesis-doctor/`（审计报告在 `audit/`、计划与快照在同级）。
该目录含快照与幂等账本，**不要当缓存删除**，删了就无法回滚。

子 skill 的详细说明在各自的 `SKILL.md`，例如 `{pkg}\\skills\\cite-doctor\\SKILL.md`。

## 若包目录不存在

先告诉用户「paper-detail-doctor 的包目录 `{pkg}` 找不到了」，
**不要凭记忆编造命令或路径**。包已托管在 GitHub（`https://github.com/aa-bao/paper-detail-doctor.git`），
可以重新 clone 后再跑一次 `tools/install.py`。
'''


def find_python():
    """找一个真的能 import lxml / docx / yaml 的解释器，写进存根。

    这一步不是洁癖：本包的脚本依赖 lxml + python-docx + pyyaml，
    而裸的 `python` 在有些机器上没装这些。存根里写死一个验证过的解释器，
    用户就不用去猜"为什么跑不起来"。
    """
    cands = [sys.executable, 'python', 'python3',
             r'E:/Miniconda3/python.exe']
    for c in cands:
        if not c:
            continue
        exe = c if os.path.isabs(c) else (shutil.which(c) or '')
        if not exe or not os.path.exists(exe):
            continue
        try:
            import subprocess
            p = subprocess.run([exe, '-c', 'import lxml, docx, yaml'],
                               capture_output=True, timeout=60)
            if p.returncode == 0:
                return exe
        except Exception:
            continue
    return sys.executable          # 兜底：至少写个存在的


def scan_unexpected_skill_md():
    """列出目标目录下所有 SKILL.md。方案乙下应该**只有 1 个**。

    这个体检存在的意义：一旦有人把子 skill 也拷进来（或整包拷进来），
    宿主就会多注册 N 个 skill —— 正好是本包要避免的情况。
    """
    found = []
    for root, dirs, files in os.walk(TARGET_DIR):
        dirs[:] = [d for d in dirs if d not in ('.git', '__pycache__')]
        for f in files:
            if f == 'SKILL.md':
                found.append(os.path.join(root, f))
    return found


def cmd_check():
    print('包目录        : %s' % PKG)
    print('安装目标      : %s' % TARGET_DIR)
    print('存根是否存在  : %s' % os.path.exists(TARGET_SKILL))
    if not os.path.isdir(TARGET_DIR):
        print('→ 尚未安装。跑 `python tools/install.py` 安装。')
        return 0
    found = scan_unexpected_skill_md()
    print('目标目录下的 SKILL.md：')
    for p in found:
        print('   %s' % os.path.relpath(p, TARGET_DIR))
    if len(found) > 1:
        print('⚠️ 发现 %d 个 SKILL.md —— 宿主会注册 %d 个 skill，'
              '与本包的"只注册入口"设计不符。多余的是：' % (len(found), len(found)))
        for p in found:
            if os.path.abspath(p) != os.path.abspath(TARGET_SKILL):
                print('   ' + p)
        return 1
    print('✓ 只有入口存根 1 个 SKILL.md，符合"只注册包入口"的设计。')
    # 包目录可达性
    real = os.path.join(PKG, 'SKILL.md')
    print('包内真实入口  : %s (%s)' % (real, '可读' if os.path.exists(real) else '⚠️ 不存在'))
    return 0


def cmd_uninstall():
    if not os.path.isdir(TARGET_DIR):
        print('未安装，无需卸载。')
        return 0
    bk = os.path.join(os.path.expanduser('~'), '.workbuddy',
                      '_skill-backups', '%s-uninstalled-%s' % (SKILL_NAME, time.strftime('%Y%m%d-%H%M%S')))
    os.makedirs(os.path.dirname(bk), exist_ok=True)
    shutil.copytree(TARGET_DIR, bk)
    print('已备份到：%s' % bk)
    shutil.rmtree(TARGET_DIR)
    print('已卸载：%s' % TARGET_DIR)
    return 0


def main():
    ap = argparse.ArgumentParser(description='安装 paper-detail-doctor 到 ~/.workbuddy/skills/（只注册包入口）')
    ap.add_argument('--check', action='store_true', help='只体检，不写盘')
    ap.add_argument('--uninstall', action='store_true', help='卸载（先备份）')
    a = ap.parse_args()

    if a.check:
        return cmd_check()
    if a.uninstall:
        return cmd_uninstall()

    if not os.path.exists(os.path.join(PKG, 'SKILL.md')):
        raise SystemExit('包内 SKILL.md 不存在，拒绝安装：%s' % PKG)
    if not os.path.isdir(SKILLS_DIR):
        raise SystemExit('skills 目录不存在：%s' % SKILLS_DIR)

    py = find_python()
    body = STUB.format(name=SKILL_NAME, version=VERSION, pkg=PKG, python_exe=py)

    # 已存在且内容相同 → 什么都不做（幂等）
    if os.path.exists(TARGET_SKILL):
        old = open(TARGET_SKILL, encoding='utf-8').read()
        if old == body:
            print('存根已是最新（内容一致），未改动。')
            return 0

    os.makedirs(TARGET_DIR, exist_ok=True)

    # 若目标里混进了别的东西（比如误拷的子 skill），先挪走并告知，别直接删
    stray = [p for p in scan_unexpected_skill_md()
             if os.path.abspath(p) != os.path.abspath(TARGET_SKILL)]
    if stray:
        bk = os.path.join(os.path.expanduser('~'), '.workbuddy', '_skill-backups',
                          '%s-strays-%s' % (SKILL_NAME, time.strftime('%Y%m%d-%H%M%S')))
        os.makedirs(bk, exist_ok=True)
        for p in stray:
            rel = os.path.relpath(os.path.dirname(p), TARGET_DIR)
            dst = os.path.join(bk, rel.replace(os.sep, '__'))
            shutil.move(os.path.dirname(p), dst)
            print('⚠️ 把意外的 skill 目录挪出安装位置：%s → %s' % (rel, dst))

    if os.path.exists(TARGET_SKILL):
        bk = TARGET_SKILL + '.bak-' + time.strftime('%Y%m%d-%H%M%S')
        shutil.copy2(TARGET_SKILL, bk)
        print('旧存根已备份：%s' % bk)

    with open(TARGET_SKILL, 'w', encoding='utf-8', newline='\n') as f:
        f.write(body)
    print('已写入存根：%s（%d 字节）' % (TARGET_SKILL, len(body.encode('utf-8'))))
    print('python 解释器：%s' % py)
    print()
    return cmd_check()


if __name__ == '__main__':
    sys.exit(main())
