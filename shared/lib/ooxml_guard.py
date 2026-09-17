#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
shared/lib/ooxml_guard.py —— 三道闸门：快照 / 幂等 / 回滚

改稿这件事最怕三样：
    ① 改坏了，且改之前没留底 → 毕不了业
    ② 同一批修改跑了两遍 → 引注变成 [12][12]，字号被反复缩放
    ③ 改到一半报错退出 → 文件处于半改状态，肉眼看不出来

本模块对应地给出三件东西：
    ① SnapshotManager —— 动手前把整档复制到 .thesis-doctor/backups/，保留最近 N 份；
       回滚 = 拿快照覆盖回去，一行命令，不依赖任何"反向动作"。
    ② RunManifest    —— 按「文档哈希 + issue_id」记账。同一份文档、同一条问题
       已经改过就不再改。注意记账键里带文档哈希：你改完又手动编了一稿，
       哈希变了，账本自动失效，下次该改的还会改 —— 这才是正确语义。
    ③ GuardedApply   —— 唯一允许真正落盘的入口：先校验哈希 → 快照 → 按段号倒序
       执行结构性动作 → 原子保存 → 写 journal。任何一条动作失败都不会中断整批，
       会记成 skip/failed 并汇总成报告给人看。

为什么"倒序执行"必须写死
    段落删除/合并之后，其后所有段的段号都会位移。若 plan 里同时有
    "删第 500 段"和"改第 520 段的样式"，正序执行会让第 2 条定位漂到别的段上。
    倒序执行则前面的段号不受影响。非结构性动作（改字号、加链接）顺序无关。

放置位置约定
    全部落在**目标文档所在目录**的 .thesis-doctor/ 下，不污染包目录，也方便
    用户跟文档一起备份/带走：
        .thesis-doctor/backups/<name>.<stamp>.docx
        .thesis-doctor/runs/<name>.<sha8>.json
        .thesis-doctor/last-journal.md
"""
import hashlib
import json
import os
import re
import shutil
import time

from shared.lib.docx_ops import (
    ACTIONS, STRUCTURAL_ACTIONS, Doc, LocateError, OpError, run_action,
)

STATE_DIR = '.thesis-doctor'


# ---------------------------------------------------------------- 小工具

def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def sha8(path_or_hash: str) -> str:
    if os.path.exists(path_or_hash):
        return sha256_file(path_or_hash)[:8]
    return (path_or_hash or '')[:8]


def state_dir(docx_path: str) -> str:
    return os.path.join(os.path.dirname(os.path.abspath(docx_path)), STATE_DIR)


def is_locked(path: str) -> bool:
    """探测文件是否被 WPS/Word 独占。Windows 上独占时 r+b 会 PermissionError。"""
    try:
        with open(path, 'r+b'):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


# ---------------------------------------------------------------- ① 快照

class SnapshotManager:
    def __init__(self, docx_path: str, keep: int = 12):
        self.docx = os.path.abspath(docx_path)
        self.keep = keep
        self.root = os.path.join(state_dir(self.docx), 'backups')

    def create(self, tag: str = 'pre-apply') -> str:
        os.makedirs(self.root, exist_ok=True)
        stamp = time.strftime('%Y%m%d-%H%M%S')
        base = '%s.%s.%s' % (os.path.splitext(os.path.basename(self.docx))[0], tag, stamp)
        dst = os.path.join(self.root, base + '.docx')
        # ★ 同一秒内连做两次备份（比如连着跑两遍 apply）会重名，
        #   不加重号就直接被后一次覆盖掉 —— 第一次的回滚点就没了。
        n = 2
        while os.path.exists(dst):
            dst = os.path.join(self.root, '%s-%d.docx' % (base, n))
            n += 1
        shutil.copy2(self.docx, dst)
        self._prune()
        return dst

    def _prune(self):
        files = self.list()
        for old in files[self.keep:]:
            try:
                os.remove(old)
            except OSError:
                pass

    @staticmethod
    def _sort_key(path):
        """按文件名里的时间戳排序，**不要**用 mtime。

        同一秒内产生的几个快照 mtime 完全相同，靠 mtime 排序时谁排第一是不确定的
        （取决于 os.listdir 的顺序），"回滚到最近一份"就会随机回到某一份 ——
        自检里正是因此回滚到了更早的那份快照。文件名带秒级时间戳 + 重号，可完全排序。
        """
        m = re.search(r'\.(\d{8}-\d{6})(?:-(\d+))?\.docx$', os.path.basename(path))
        if m:
            return (m.group(1), int(m.group(2) or 1))
        return ('00000000-000000', 0)          # 认不出名字的老文件排到最后

    def list(self):
        if not os.path.isdir(self.root):
            return []
        fs = [os.path.join(self.root, f) for f in os.listdir(self.root)
              if f.lower().endswith('.docx') and not f.startswith('.')]
        return sorted(fs, key=self._sort_key, reverse=True)

    def restore(self, snapshot: str):
        """回滚：把快照覆盖回目标文档。

        ★ 顺序不能反：**先把快照内容读到临时文件**，再生成"回滚前"备份。
          因为 create() 里会 _prune()（只保留最近 keep 份），若先备份后读源，
          很可能把你要恢复的那个快照本身给裁掉了 —— 自检里就撞见
          "系统找不到指定的文件"，回滚直接失效，这属于最不该坏的地方。
        """
        snapshot = os.path.abspath(snapshot)
        if not os.path.exists(snapshot):
            raise OpError('快照不存在：%s' % snapshot)
        os.makedirs(self.root, exist_ok=True)
        tmp = os.path.join(self.root, '.restore-tmp.docx')
        shutil.copy2(snapshot, tmp)              # ① 先把内容物理搬出来
        try:
            if os.path.exists(self.docx):
                self.create(tag='pre-restore')   # ② 再备份当前（可能触发裁剪）
        except OSError:
            pass
        os.replace(tmp, self.docx)               # ③ 原子落位
        return self.docx


# ---------------------------------------------------------------- ② 幂等账本

class RunManifest:
    """按「文档 sha8」分文件的已应用记录。

    键 = issue_id。值 = {at, action, note}。
    audit 生成的 issue_id 是"规则名 + 定位指纹"的稳定哈希，所以只要文档那段
    文字没变，重跑时 ID 不变 → 账本命中 → 跳过。这正是"状态驱动、不叠加改动"
    的落点。
    """

    def __init__(self, docx_path: str):
        self.docx = os.path.abspath(docx_path)
        self.root = os.path.join(state_dir(self.docx), 'runs')
        self.digest = sha8(self.docx)
        self.path = os.path.join(
            self.root, '%s.%s.json' % (os.path.splitext(os.path.basename(self.docx))[0], self.digest))
        self.data = {'docx': self.docx, 'sha8': self.digest, 'applied': {}}
        if os.path.exists(self.path):
            try:
                with open(self.path, 'r', encoding='utf-8') as f:
                    old = json.load(f)
                if old.get('sha8') == self.digest:
                    self.data = old
            except (OSError, ValueError):
                pass

    def is_applied(self, issue_id: str) -> bool:
        return issue_id in self.data.get('applied', {})

    def mark(self, issue_id: str, action: str, note: str = ''):
        self.data.setdefault('applied', {})[issue_id] = {
            'at': time.strftime('%Y-%m-%dT%H:%M:%S'),
            'action': action, 'note': note}

    def save(self):
        os.makedirs(self.root, exist_ok=True)
        tmp = self.path + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, self.path)

    def forget(self, issue_id: str):
        self.data.get('applied', {}).pop(issue_id, None)


# ---------------------------------------------------------------- ③ 受控执行

def guard_plan(docx_path: str, plan: dict, strict_hash: bool = True):
    """apply 之前的静态体检。返回 [(level, message)]，level ∈ info/warn/error。"""
    out = []
    if not os.path.exists(docx_path):
        out.append(('error', '目标文档不存在：%s' % docx_path))
        return out
    p_docx = os.path.abspath(plan.get('docx') or '')
    if p_docx and p_docx != os.path.abspath(docx_path):
        out.append(('error',
                    'plan 是为「%s」生成的，当前目标是「%s」—— 拒绝套用。'
                    % (p_docx, os.path.abspath(docx_path))))
        return out
    want = plan.get('docx_sha256')
    if want:
        got = sha256_file(docx_path)
        if got != want:
            msg = ('文档哈希已变（plan 生成于改动前）。很可能你已经手改或编过一稿，'
                   'plan 里的段落定位可能失效。')
            out.append(('error' if strict_hash else 'warn', msg))
    else:
        out.append(('warn', 'plan 未记录文档哈希，无法判断文档是否被改过。'))
    if is_locked(docx_path):
        out.append(('error', '文档正被 WPS/Word 打开独占，请先关闭再 apply。'))
    unknown = sorted({it.get('action') for it in plan.get('items', [])
                      if it.get('action') and it['action'] not in ACTIONS})
    if unknown:
        out.append(('error', 'plan 含未知动作：%s' % ', '.join(unknown)))
    n_on = sum(1 for it in plan.get('items', []) if it.get('enabled'))
    out.append(('info', '已勾选 %d / %d 条。' % (n_on, len(plan.get('items', [])))))
    return out


def apply_plan(docx_path: str, plan: dict, dry_run: bool = False,
               use_manifest: bool = True, keep_snapshots: int = 12) -> dict:
    """唯一允许落盘的入口。

    返回 journal：
      {docx, dry_run, snapshot, applied:[…], skipped:[…], failed:[…], changed_count}
    每条记录都带 issue_id / action / title / reason / 定位段落号 —— 报告直接可用。
    """
    journal = {
        'docx': os.path.abspath(docx_path), 'dry_run': bool(dry_run),
        'started_at': time.strftime('%Y-%m-%dT%H:%M:%S'),
        'snapshot': None, 'applied': [], 'skipped': [], 'failed': [],
        'changed_count': 0,
    }
    checks = guard_plan(docx_path, plan)
    errors = [m for lv, m in checks if lv == 'error']
    if errors and not dry_run:
        journal['failed'].append({'issue_id': '-', 'action': '-',
                                  'title': '前置体检未通过', 'reason': '；'.join(errors)})
        return journal

    items = [it for it in plan.get('items', []) if it.get('enabled')]
    if not items:
        return journal

    # ★ 执行分三段，顺序不能乱（写死，别改成"按 plan 里的顺序"）：
    #   ① ref.bookmark  —— 先把书签立起来，后面的引注链接才有目标可指
    #   ② 原地动作      —— 改字号 / 加超链接 / 挪位置，互不干扰
    #   ③ 结构性动作    —— 删段/并段/插段，按段号**倒序**，避免段号位移影响前面的动作
    def _order(it):
        act = it.get('action')
        loc = (it.get('params') or {}).get('locator') or {}
        hint = loc.get('para_index') or 0
        if act == 'ref.bookmark':
            return 0, -hint
        if act in STRUCTURAL_ACTIONS:
            return 2, -hint
        return 1, -hint

    items = sorted(items, key=_order)

    doc = Doc(docx_path)
    man = RunManifest(docx_path) if use_manifest else None

    if not dry_run:
        journal['snapshot'] = SnapshotManager(docx_path, keep_snapshots).create('pre-apply')

    for it in items:
        iid, action = it.get('issue_id'), it.get('action')
        rec = {'issue_id': iid, 'action': action,
               'rule': it.get('rule'), 'title': it.get('title'),
               'para_index': ((it.get('params') or {}).get('locator') or {}).get('para_index')}
        if action not in ACTIONS:
            rec['reason'] = '未知动作'
            journal['failed'].append(rec)
            continue
        if man is not None and man.is_applied(iid):
            rec['reason'] = '账本记录已应用（幂等跳过）'
            journal['skipped'].append(rec)
            continue
        try:
            changed = run_action(doc, action, it.get('params') or {})
            if changed:
                journal['applied'].append(rec)
                journal['changed_count'] += 1
                if man is not None:
                    man.mark(iid, action)
            else:
                rec['reason'] = '已是目标状态（幂等跳过）'
                journal['skipped'].append(rec)
                if man is not None:
                    man.mark(iid, action)
        except (LocateError, OpError) as e:
            rec['reason'] = '%s: %s' % (type(e).__name__, e)
            journal['failed'].append(rec)
        except Exception as e:  # 不让单条异常打断整批
            rec['reason'] = '意外错误 %s: %s' % (type(e).__name__, e)
            journal['failed'].append(rec)

    if not dry_run and (journal['changed_count'] or journal['skipped']):
        journal['saved_to'] = doc.save_atomic()
        if man is not None:
            man.save()

    journal['finished_at'] = time.strftime('%Y-%m-%dT%H:%M:%S')
    _write_journal(docx_path, journal)
    return journal


def _write_journal(docx_path: str, journal: dict):
    root = state_dir(docx_path)
    os.makedirs(root, exist_ok=True)
    p = os.path.join(root, 'last-journal.json')
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(journal, f, ensure_ascii=False, indent=2)
    md = os.path.join(root, 'last-journal.md')
    with open(md, 'w', encoding='utf-8') as f:
        f.write(render_journal(journal))
    journal['journal_json'] = p
    journal['journal_md'] = md


def render_journal(j: dict) -> str:
    L = []
    L.append('# 改稿执行记录')
    L.append('')
    L.append('- 文档：`%s`' % j['docx'])
    L.append('- 时间：%s → %s' % (j.get('started_at'), j.get('finished_at', '')))
    if j.get('snapshot'):
        L.append('- 动手前的快照：`%s`（改坏了一句话就能还原）' % j['snapshot'])
    if j.get('dry_run'):
        L.append('- **试运行**，没有写盘。')
    L.append('- 实际改动 **%d** 条；幂等跳过 %d 条；失败 %d 条。'
             % (j.get('changed_count', 0), len(j.get('skipped', [])), len(j.get('failed', []))))
    L.append('')
    if j.get('applied'):
        L.append('## 已改（%d）' % len(j['applied']))
        L.append('')
        for r in j['applied']:
            L.append('- 段%s `%s` %s' % (r.get('para_index') or '-', r.get('action'),
                                         r.get('title') or ''))
        L.append('')
    if j.get('skipped'):
        L.append('## 跳过（%d）' % len(j['skipped']))
        L.append('')
        for r in j['skipped']:
            L.append('- 段%s `%s` %s —— %s'
                     % (r.get('para_index') or '-', r.get('action'),
                        r.get('title') or '', r.get('reason') or ''))
        L.append('')
    if j.get('failed'):
        L.append('## 未改成功（%d）—— **请人工处理**' % len(j['failed']))
        L.append('')
        for r in j['failed']:
            L.append('- 段%s `%s` %s' % (r.get('para_index') or '-', r.get('action'),
                                         r.get('title') or ''))
            L.append('  - 原因：%s' % (r.get('reason') or ''))
        L.append('')
    return '\n'.join(L)
