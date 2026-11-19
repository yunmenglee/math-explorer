#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""知识库数据体检: 结构硬校验 + 质量软提示。CI 里跑, 结构问题返回非 0。"""
import json, os, re, sys
from collections import Counter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
IMG = os.path.join(ROOT, 'images')

errors, warns = [], []
def err(m): errors.append(m)
def warn(m): warns.append(m)

def load(name):
    p = os.path.join(DATA, name)
    if not os.path.exists(p):
        err(f'缺少数据文件 data/{name}')
        return None
    try:
        with open(p, encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        err(f'data/{name} 不是合法 JSON: {e}')
        return None

entries = load('entries.json')
textbooks = load('textbooks.json')
summary = load('summary.json')
categories = load('categories.json')
meta = load('meta.json')
if errors:
    print('\n'.join('❌ ' + e for e in errors)); sys.exit(1)

STAGES = {'小学', '初中', '中考数学'}
TPS = {'basic', 'extend', 'exam'}
SEC_NUM = re.compile(r'^\d+(\.\d+)*$')

# ---------- 条目 ----------
for i, e in enumerate(entries):
    for k in ('m', 't', 'd', 'c1', 'c2', 'c3', 'x', 'src', 'tp'):
        if k not in e:
            err(f'entries[{i}] 缺字段 {k} (t={e.get("t")})'); break
    if e.get('m') not in STAGES:
        err(f'entries[{i}] 学段异常: {e.get("m")}')
    if e.get('tp') not in TPS:
        err(f'entries[{i}] tp 异常: {e.get("tp")}')
    for nm in (e.get('img') or '').split(','):
        nm = nm.strip()
        if nm and not os.path.exists(os.path.join(IMG, nm)):
            err(f'entries[{i}] 图片不存在: images/{nm}')

if 'auth_password' in meta:
    err('data/meta.json 仍包含明文密码字段 auth_password')

# ---------- 教材 ----------
ids = [b.get('id') for b in textbooks]
dup = [k for k, v in Counter(ids).items() if v > 1]
if dup:
    err(f'教材 id 重复: {dup[:5]}')
grades = {'一年级上册','一年级下册','二年级上册','二年级下册','三年级上册','三年级下册',
          '四年级上册','四年级下册','五年级上册','五年级下册','六年级上册','六年级下册',
          '七年级上册','七年级下册','八年级上册','八年级下册','九年级上册','九年级下册'}
for b in textbooks:
    bid = b.get('id')
    for k in ('id','publisher','grade','label','iconText','iconBg','chapters'):
        if k not in b:
            err(f'教材 {bid} 缺字段 {k}')
    if b.get('grade') not in grades:
        err(f'教材 {bid} 年级异常: {b.get("grade")}')
    if 'stage' not in b:
        warn(f'教材 {bid} 缺少 stage 字段(将按初中处理)')
    for ch in b.get('chapters') or []:
        if 'num' not in ch or 'name' not in ch:
            err(f'教材 {bid} 章节缺 num/name')
        for sec in ch.get('sections') or []:
            if not SEC_NUM.match(str(sec.get('num', ''))):
                err(f'教材 {bid} 第{ch.get("num")}章节号异常: {sec.get("num")!r}')
            for k in sec.get('kps') or []:
                if isinstance(k, str):
                    continue
                if not isinstance(k, dict) or not k.get('name'):
                    err(f'教材 {bid} KP 格式异常: {k!r}')
            for k in sec.get('des') or []:
                if isinstance(k, str):
                    continue
                if not isinstance(k, dict) or not k.get('name'):
                    err(f'教材 {bid} DE 格式异常: {k!r}')

# ---------- 统计对账 ----------
recalc = {}
recat = {}
for e in entries:
    st = e.get('m')
    if e.get('c1'):
        recalc.setdefault(st, Counter())[e['c1']] += 1
        recat.setdefault(st, {}).setdefault(e['c1'], Counter())[e.get('c2') or ''] += 1
for st in STAGES:
    s = summary.get(st) or {}
    total = sum(recalc.get(st, Counter()).values())
    if s.get('n') != total:
        err(f'summary[{st}].n={s.get("n")} 与条目重算 {total} 不一致')
    cs = s.get('cs') or {}
    for c1, n in recalc.get(st, Counter()).items():
        if cs.get(c1) != n:
            err(f'summary[{st}].cs[{c1}]={cs.get(c1)} 与重算 {n} 不一致')
    for c1 in cs:
        if c1 not in recalc.get(st, {}):
            err(f'summary[{st}].cs 多出分类 {c1}')
    for c1, cnt in (categories.get(st) or {}).items():
        if c1 not in recat.get(st, {}):
            err(f'categories[{st}] 多出一级分类 {c1}')
            continue
        for c2, n in cnt.items():
            real = recat[st][c1].get(c2, 0)
            if real != n:
                err(f'categories[{st}][{c1}][{c2}]={n} 与重算 {real} 不一致')

if meta.get('total_entries') != len(entries):
    err(f'meta.total_entries={meta.get("total_entries")} 与条目数 {len(entries)} 不一致')

# ---------- 质量提示(不阻断) ----------
stubs = [e for e in entries if not (e.get('c1') or '').strip() and not (e.get('d') or '').strip()]
empty_d = [e for e in entries if not (e.get('d') or '').strip()]
dup_t = [k for k, v in Counter(e['t'] for e in entries).items() if v > 1]
used = set()
for e in entries:
    for nm in (e.get('img') or '').split(','):
        if nm.strip(): used.add(nm.strip())
files = {f for f in os.listdir(IMG) if not f.startswith('.')} if os.path.isdir(IMG) else set()
unref = files - used

print(f'条目 {len(entries)} | 教材 {len(textbooks)} | 章 {sum(len(b.get("chapters") or []) for b in textbooks)}'
      f' | 节 {sum(len(c.get("sections") or []) for b in textbooks for c in (b.get("chapters") or []))}')
print(f'图片 {len(files)} 张, 未被引用 {len(unref)} 张')
print(f'质量提示: 空壳条目 {len(stubs)} (已标 todo {sum(1 for e in entries if e.get("todo"))}), '
      f'空描述 {len(empty_d)}, 重复标题 {len(dup_t)}')
if warns:
    print('提醒:')
    for w in warns[:10]:
        print('  ⚠️ ', w)

if errors:
    print(f'\n❌ 体检失败, {len(errors)} 项结构问题:')
    for m in errors[:30]:
        print('  -', m)
    sys.exit(1)
print('\n✅ 数据体检通过')
