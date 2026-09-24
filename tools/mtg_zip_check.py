#!/usr/bin/env python3
"""check_meeting_output.js が書き出したパーツを pptx に固め、中身を確かめる。

    python3 tools/mtg_zip_check.py <出力ディレクトリ> <保存先.pptx>
"""
import json
import os
import re
import sys
import zipfile
from xml.dom import minidom

OUT, DST = sys.argv[1], os.path.abspath(sys.argv[2])
plan = json.load(open(os.path.join(OUT, 'plan.json')))

with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
    for p, spec in plan['plan'].items():
        src = os.path.join(OUT, 'gen', p) if spec['from'] == 'gen' else spec['src']
        z.write(src, p)

z = zipfile.ZipFile(DST)
names = set(z.namelist())
fails, checks = [], 0


def ck(c, m):
    global checks
    checks += 1
    if not c:
        fails.append(m)


ck(z.testzip() is None, 'zipが壊れている')
for n in sorted(names):
    if n.endswith('.xml') or n.endswith('.rels'):
        try:
            minidom.parseString(z.read(n))
        except Exception as e:
            fails.append('XMLとして読めない: %s (%s)' % (n, e))
        checks += 1
for n in sorted(names):
    if not n.endswith('.rels'):
        continue
    base = os.path.dirname(os.path.dirname(n))
    for m in re.finditer(r'<Relationship\b[^>]*>', z.read(n).decode('utf-8')):
        if 'External' in m.group(0):
            continue
        t = re.search(r'Target="([^"]+)"', m.group(0)).group(1)
        tgt = os.path.normpath(os.path.join(base, t)).replace('\\', '/')
        ck(tgt in names, '%s の参照先がない: %s' % (n, tgt))

pres = z.read('ppt/presentation.xml').decode('utf-8')
rels = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"',
                       z.read('ppt/_rels/presentation.xml.rels').decode('utf-8')))
order = [rels[r] for r in re.findall(r'<p:sldId id="\d+" r:id="([^"]+)"/>', pres)]


def text(p):
    s = z.read('ppt/' + p).decode('utf-8')
    return s, ''.join(re.findall(r'<a:t(?=[\s>])[^>]*>([^<]*)</a:t>', s))


left = sum(t.count('{{') for _, t in (text(p) for p in order))
ck(left == 0, '差し込み口が %d 個残っている' % left)

shown, hidden = [], []
for i, p in enumerate(order, 1):
    s, t = text(p)
    (hidden if 'show="0"' in s[:400] else shown).append((i, t[:46]))

m = plan['map']
mp = [m.get('メインプレゼン1氏名', ''), m.get('メインプレゼン2氏名', '')]
found = [p for p in order if 'Main Presenter' in text(p)[1]]
ck(len(found) == 1, 'メインプレゼンのページが %d 枚' % len(found))
if found:
    _, t = text(found[0])
    for v in [m.get('メインプレゼン1氏名'), m.get('メインプレゼン1会社名'), m.get('メインプレゼン1カテゴリー'),
              m.get('メインプレゼン2氏名'), m.get('メインプレゼン2会社名'), m.get('メインプレゼン2カテゴリー')]:
        if v:
            ck(v in t, 'メインプレゼンのページに「%s」が無い' % v)
    print('  メインプレゼンのページ: %s' % t[:110])

ms = [p for p in order if 'チャプターが求める専門分野' in text(p)[1]]
ck(len(ms) == 1, 'メンバーシップ委員会のページが %d 枚' % len(ms))
if ms:
    _, t = text(ms[0])
    for i in range(1, 13):
        v = m.get('求める専門分野%d' % i)
        if v:
            ck(v in t, 'メンバーシップ委員会のページに「%s」が無い' % v)
    if m.get('開放カテゴリー'):
        ck(m['開放カテゴリー'] in t, '開放カテゴリーが入っていない')
    if m.get('審査中カテゴリー'):
        ck(m['審査中カテゴリー'] in t, '審査中カテゴリーが入っていない')
    print('  メンバーシップ委員会のページ: %s' % t[:130])

info = plan['info']
if info.get('policy'):
    ck(info['policy'].get('shown') == 1, '一般規定の表示が %s 枚' % info['policy'].get('shown'))
if info.get('core'):
    ck(info['core'].get('shown') == 1, 'コアバリューの表示が %s 枚' % info['core'].get('shown'))

print('%s  %.1fMB  %d枚（表示 %d／非表示 %d）'
      % (os.path.basename(DST), os.path.getsize(DST) / 1048576, len(order), len(shown), len(hidden)))
print('検査 %d 件' % checks)
if fails:
    print('NG: %d 件' % len(fails))
    for f in fails[:20]:
        print('   ', f)
    sys.exit(1)
print('OK: すべて指示どおりです')
