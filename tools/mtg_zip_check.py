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
# 左右にお2人が並ぶページ（前半＝メインプレゼン／後半＝推薦のことば・抽選コーナー）
for title, prefix in (('Main Presenter', 'メインプレゼン'),
                      ('Words of Recommendation', '推薦のことば'),
                      ('賞品の抽選', '抽選')):
    found = [p for p in order if title in text(p)[1]]
    if not found:
        continue
    ck(len(found) == 1, '%s のページが %d 枚' % (prefix, len(found)))
    _, t = text(found[0])
    for n in (1, 2):
        for f in ('氏名', '会社名', 'カテゴリー'):
            v = m.get('%s%d%s' % (prefix, n, f))
            if v:
                ck(v in t, '%s のページに「%s」が無い' % (prefix, v))
    print('  %s のページ: %s' % (prefix, t[:110]))

ms = [p for p in order if 'チャプターが求める専門分野' in text(p)[1]]
if ms:
    ck(len(ms) == 1, 'メンバーシップ委員会のページが %d 枚' % len(ms))
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
# 音楽・動画（差し替えたものは中身も入れ替わっているか）
for a in info.get('audioList', []):
    n = a['slide'].replace('ppt/slides/', '')
    rp = 'ppt/slides/_rels/%s.rels' % n
    if rp in names:
        r = z.read(rp).decode('utf-8')
        for rid in (a['linkRid'], a.get('embedRid')):
            if not rid:
                continue
            mm = re.search(r'Id="%s"[^>]*Target="([^"]+)"[^>]*/>' % rid, r)
            ck(mm is not None, '%s の音楽の関係 %s が無い' % (n, rid))
            if mm and 'External' not in mm.group(0):
                tgt = os.path.normpath(os.path.join('ppt/slides', mm.group(1))).replace('\\', '/')
                ck(tgt in names, '%s の音楽の参照先がない: %s' % (n, tgt))

# リファーラル発表：人数ぶんのページが、ひな形のあった場所に並んでいるか
plan_rf = plan['info'].get('referralPlan') or []
if plan_rf:
    idx = [i for i, p in enumerate(order) if 'REFERRAL PRESENTATION' in text(p)[1]]
    ck(len(idx) == len(plan_rf),
       'リファーラル発表のページが %d 枚（%d 枚のはず）' % (len(idx), len(plan_rf)))
    ck(idx == list(range(idx[0], idx[0] + len(idx))) if idx else False,
       'リファーラル発表のページが連続していない')
    for k, it in enumerate(plan_rf):
        if k >= len(idx):
            break
        sx, t = text(order[idx[k]])
        ck(it['name'] in t, '%d人目のページに「%s」が無い' % (k + 1, it['name']))
        for line in it['companyLines']:
            if line:
                ck(line in t, '%d人目のページに会社名「%s」が無い' % (k + 1, line))
        if it['nextName']:
            ck(it['nextName'] in t, '%d人目のページにNEXT「%s」が無い' % (k + 1, it['nextName']))
        adv = set(re.findall(r'advTm="(\d+)"', sx))
        if it.get('auto'):
            ck(adv == {str((it['seconds'] + 1) * 1000)},
               '%d人目の自動送りが %s（%d のはず）' % (k + 1, adv or 'なし', (it['seconds'] + 1) * 1000))
            ck('<p:cond delay="indefinite"/>' not in sx, '%d人目のカウントダウンがクリック待ち' % (k + 1))
        else:
            # 自動で進まない：自動送りが無く、カウントダウンはクリックで始まる
            ck(not adv, '%d人目に自動送りが残っている: %s' % (k + 1, adv))
            ck('<p:cond delay="indefinite"/>' in sx, '%d人目のカウントダウンがクリックで始まらない' % (k + 1))
        ck(len(re.findall(r'<p:spTgt spid="\d+"/>', sx)) == it['seconds'],
           '%d人目のカウントダウンの手順が %d 回（%d 回のはず）'
           % (k + 1, len(re.findall(r'<p:spTgt spid="\d+"/>', sx)), it['seconds']))
    # 最後の人だけ NEXT が消えていること
    if idx:
        _, tl = text(order[idx[-1]])
        ck('NEXT' not in tl, '最後の方のページに NEXT が残っている')
    print('  リファーラル発表: %d枚（%s → … → %s）%s'
          % (len(idx), plan_rf[0]['name'], plan_rf[-1]['name'],
             '・自動で次へ' if plan_rf[0].get('auto') else '・クリックで次へ'))
    if not any(x.get('auto') for x in plan_rf):
        # 自動で進むページが無いなら、ファイル全体の設定も元のまま（タイミングを使わない）
        pp = z.read('ppt/presProps.xml').decode('utf-8')
        ck(re.search(r'useTimings="1"', pp) is None,
           '自動で進むページが無いのに「保存済みのタイミングを使用」が入っている')

# 前半の「ウィークリープレゼンテーション」の見出しと「全員終わりましたか？」のページ
def flat(p):
    return re.sub(r'[\s\u3000]', '', text(p)[1])


anchor = next((i for i, p in enumerate(order)
               if 'ウィークリー' in flat(p) and 'プレゼンテーション' in flat(p)
               and '終わりましたか' not in flat(p)), None)
wk_end = next((i for i, p in enumerate(order)
               if anchor is not None and i > anchor and '終わりましたか' in flat(p)), None)
plan_mp = plan['info'].get('memberPresenPlan') or []

# アンバサダー・ディレクターのページ：チェックした方だけ表示で、
# メンバーのページの後ろ・「全員終わりましたか？」の前にあること
guest_pages = info.get('guestPages') or []
guest_res = info.get('guests')
guest_auto = set()
if guest_pages and guest_res is not None:
    picked = set(guest_res.get('shown') or [])
    for g in guest_pages:
        gp = g['path'].replace('ppt/', '')
        ck(gp in order, '%sさんのページが消えている' % g['name'])
        if gp not in order:
            continue
        gi, (sx, t) = order.index(gp), text(gp)
        on = g['name'] in picked
        ck(('show="0"' not in sx[:600]) == on,
           '%sさんのページが%s（%sのはず）' % (g['name'], '非表示' if on else '表示', '表示' if on else '非表示'))
        ck(g['name'] in t, '%sさんのページに氏名が無い' % g['name'])
        if wk_end is not None:
            ck(gi < wk_end, '%sさんのページが「全員終わりましたか？」より後ろにある' % g['name'])
        if anchor is not None:
            ck(gi > anchor + len(plan_mp), '%sさんのページがメンバーのページより前にある' % g['name'])
        adv = set(re.findall(r'advTm="(\d+)"', sx))
        if on:
            guest_auto.add(gp)
            ck(adv == {'31000'}, '%sさんのページの自動送りが %s（31000 のはず）' % (g['name'], adv or 'なし'))
            ck('<p:cond delay="indefinite"/>' not in sx, '%sさんのカウントダウンがクリック待ち' % g['name'])
            ck(len(re.findall(r'<p:spTgt spid="\d+"/>', sx)) == 30,
               '%sさんのカウントダウンの手順が %d 回（30 回のはず）'
               % (g['name'], len(re.findall(r'<p:spTgt spid="\d+"/>', sx))))
        else:
            ck(not adv, '非表示の%sさんのページに自動送りが残っている: %s' % (g['name'], adv))
    print('  アンバサダー・ディレクター: 表示 %s／非表示 %s'
          % ('、'.join(guest_res.get('shown') or []) or 'なし', '、'.join(guest_res.get('hidden') or []) or 'なし'))

# メンバープレゼンを前半に差し込んだとき
if plan_mp:
    # 差し込んだページは「ウィークリープレゼンテーション」の見出しの直後に並んでいるか
    ck(anchor is not None, '「ウィークリープレゼンテーション」の見出しページが無い')
    if anchor is not None:
        block = order[anchor + 1: anchor + 1 + len(plan_mp)]
        ck(len(block) == len(plan_mp), '差し込んだページが %d 枚（%d 枚のはず）' % (len(block), len(plan_mp)))
        for k, it in enumerate(plan_mp):
            if k >= len(block):
                break
            sx, t = text(block[k])
            if it['kind'] == 'overview':
                ck(it['block'] in t, '%d枚目の扉ページに「%s」が無い' % (k + 1, it['block']))
                ck('advTm=' not in sx, '%d枚目の扉ページに自動送りが入っている' % (k + 1))
            else:
                ck(it['name'] in t, '%d枚目に「%s」が無い' % (k + 1, it['name']))
                if it.get('nextName'):
                    ck(it['nextName'] in t, '%d枚目にNEXT「%s」が無い' % (k + 1, it['nextName']))
                sec = it.get('countdownSec') or 30
                adv = set(re.findall(r'advTm="(\d+)"', sx))
                ck(adv == {str((sec + 1) * 1000)},
                   '%d枚目（%s）の自動送りが %s（%d のはず）' % (k + 1, it['name'], adv or 'なし', (sec + 1) * 1000))
                ck('<p:cond delay="indefinite"/>' not in sx, '%d枚目のカウントダウンがクリック待ち' % (k + 1))
        # 差し込んだページの後ろに、2分30秒の下書き（非表示）が残っているか
        rest = [text(p)[1] for p in order[anchor + 1 + len(plan_mp):]]
        ck(any('１分' in t for t in rest), '2分30秒の下書きのページが消えている')
    # 自動送りがあるので「保存済みのタイミングを使用」が入っていること
    pp = z.read('ppt/presProps.xml').decode('utf-8')
    ck(re.search(r'useTimings="1"', pp) is not None, '「保存済みのタイミングを使用」が入っていない')
    # 自分で作っていないページには自動送りが残っていないこと（0.4秒などが急に効くのを防ぐ）
    mine = set(order[anchor + 1: anchor + 1 + len(plan_mp)]) if anchor is not None else set()
    mine |= guest_auto
    stale = [(i + 1, re.findall(r'advTm="(\d+)"', z.read('ppt/' + p).decode('utf-8')))
             for i, p in enumerate(order) if p not in mine and 'advTm=' in z.read('ppt/' + p).decode('utf-8')]
    ck(not stale, '作っていないページに自動送りが残っている: %s' % stale[:5])
    two = [it['name'] for it in plan_mp if it.get('countdownSec')]
    print('  メンバープレゼン: %d枚を差し込み（2分30秒: %s）' % (len(plan_mp), '、'.join(two) or 'なし'))

# 一般規定・コアバリューは前半スライドだけにあるページ。
# そのページが見つかったときにだけ「1枚だけ表示」になっていることを確かめる。
if info.get('policy') and info['policy'].get('found', 0) >= 6:
    ck(info['policy'].get('shown') == 1, '一般規定の表示が %s 枚' % info['policy'].get('shown'))
if info.get('core') and info['core'].get('kinds', 0) >= 3:
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
