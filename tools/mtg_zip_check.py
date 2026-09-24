#!/usr/bin/env python3
"""check_meeting_output.js が書き出したパーツを pptx に固め、中身を確かめる。

    python3 tools/mtg_zip_check.py <出力ディレクトリ> <保存先.pptx>
"""
import hashlib
import html
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
    return s, html.unescape(''.join(re.findall(r'<a:t(?=[\s>])[^>]*>([^<]*)</a:t>', s)))


left = sum(t.count('{{') for _, t in (text(p) for p in order))
ck(left == 0, '差し込み口が %d 個残っている' % left)

shown, hidden = [], []
for i, p in enumerate(order, 1):
    s, t = text(p)
    (hidden if 'show="0"' in s[:400] else shown).append((i, t[:46]))

def check_fit(p, label, vals):
    """お2人のページ：氏名・会社名は1行、カテゴリーは2行に収まる文字の大きさか（全角1em・半角0.55em・空白0.3emで見積もる）"""
    sx = z.read('ppt/' + p).decode('utf-8')
    for sp in re.finditer(r'<p:sp>.*?</p:sp>', sx, re.S):
        seg = sp.group(0)
        st = html.unescape(''.join(re.findall(r'<a:t(?=[\s>])[^>]*>([^<]*)</a:t>', seg))).strip()
        for v, lines in vals:
            if not v or st != v:
                continue
            ext = re.search(r'<a:ext cx="(\d+)"', seg)
            bp = re.search(r'<a:bodyPr\b[^>]*>', seg)

            def ins(k):
                mm = re.search(r'\s%s="(\d+)"' % k, bp.group(0)) if bp else None
                return int(mm.group(1)) if mm else 91440
            w = (int(ext.group(1)) - ins('lIns') - ins('rIns')) / 12700.0 * 0.95  # 字幅のゆとり5%
            sz = [int(x) / 100.0 for x in re.findall(r'<a:rPr\b[^>]*\ssz="(\d+)"', seg)]
            em = sum(0.3 if c == ' ' else (0.55 if ord(c) < 128 else 1.0) for c in v)
            ck(bool(sz) and em * max(sz) <= w * lines + 0.01,
               '%s の「%s」が %d 行に収まらない（%.1fpt）' % (label, v, lines, max(sz) if sz else 0))


m = plan['map']
# 左右にお2人が並ぶページ（前半＝メインプレゼン／後半＝抽選コーナー）。推薦のことばは組ごとなので別に確かめる
for title, prefix in (('Main Presenter', 'メインプレゼン'),
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
        check_fit(found[0], prefix, [(m.get('%s%d氏名' % (prefix, n)), 1), (m.get('%s%d会社名' % (prefix, n)), 1),
                                     (m.get('%s%dカテゴリー' % (prefix, n)), 2)])
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

# 写真の取り違い：ページごとに、入るはずの方の写真が入るはずの枠に入っているか。
# 写真は1人ずつ中身の違う画像にしてあるので、中身（md5）で誰の写真かが分かる。
photo_of = info.get('photoOf') or {}
md5_of_person = {}
for nm, fp in photo_of.items():
    try:
        md5_of_person[hashlib.md5(open(fp, 'rb').read()).hexdigest()] = nm
    except OSError:
        pass
slide_w = int((re.search(r'<p:sldSz\s+cx="(\d+)"', pres) or [0, 12192000])[1])


def pics_of(p):
    """ページの画像 → [(id, x, y, cx, cy, 写っている方の氏名 or '')]"""
    sx = z.read('ppt/' + p).decode('utf-8')
    rp = 'ppt/slides/_rels/' + p.split('/')[-1] + '.rels'
    rels = z.read(rp).decode('utf-8') if rp in names else ''
    tg = {}
    for m in re.finditer(r'<Relationship\b[^>]*>', rels):
        a = m.group(0)
        i = re.search(r'Id="([^"]+)"', a)
        t = re.search(r'Target="([^"]+)"', a)
        if i and t and 'External' not in a:
            tg[i.group(1)] = os.path.normpath(os.path.join('ppt/slides', t.group(1))).replace('\\', '/')
    out = []
    for m in re.finditer(r'<p:pic>.*?</p:pic>', sx, re.S):
        seg = m.group(0)
        pid = re.search(r'<p:cNvPr id="(\d+)"', seg)
        geo = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"\s*/>\s*<a:ext cx="(\d+)" cy="(\d+)"', seg)
        rid = re.search(r'<a:blip[^>]*r:embed="([^"]+)"', seg)
        who = ''
        if rid and tg.get(rid.group(1)) in names:
            who = md5_of_person.get(hashlib.md5(z.read(tg[rid.group(1)])).hexdigest(), '')
        if pid and geo:
            out.append((pid.group(1),) + tuple(int(v) for v in geo.groups()) + (who,))
    return out


def check_two_photos(p, prefix, want):
    """左右にお2人が並ぶページ：左に1人目、右に2人目。ほかの方の写真や、元の写真が残っていないこと"""
    pics = pics_of(p)
    # 写真の枠＝縦長で大きな画像（背景いっぱいの画像と、飾りの小さな画像は除く）
    frames = [q for q in pics if q[4] >= 2500000 and q[3] <= slide_w * 0.4]
    for side, nm in enumerate(want[:2]):
        on_side = [q[5] for q in frames if q[5] and ((q[1] + q[3] / 2) < slide_w / 2) == (side == 0)]
        exp = [nm] if nm in photo_of else []
        ck(on_side == exp, '%s の%sの写真が %s（%s のはず）'
           % (prefix, '左' if side == 0 else '右', on_side or 'なし', exp or '写真なし'))
    # 写真の枠に、入るはずの方以外（元のテンプレートの写真）が残っていないこと
    stale = [q[0] for q in frames if q[5] not in [n for n in want if n]]
    ck(not stale, '%s のページに、元の写真が残っている枠がある（図形ID %s）' % (prefix, stale))
    # 飾りの画像に、誰かの写真が入っていないこと
    deco = [(q[0], q[5]) for q in pics if q not in frames and q[5]]
    ck(not deco, '%s のページで、写真の枠でない画像に写真が入っている: %s' % (prefix, deco))


# 推薦のことば（何組でも）：
#   定例会中の組 … ひな形のページの場所に続けて並ぶ（組が無ければ、ひな形のページは非表示で残る）
#   アフター・定例会後の組 … 抽選コーナーのページのすぐうしろに並ぶ
reco_pages = []
pairs = info.get('recoPairs')
if pairs is not None:
    during = [x for x in pairs if not x.get('after')]
    after = [x for x in pairs if x.get('after')]
    idx = [i for i, p in enumerate(order) if 'Words of Recommendation' in text(p)[1]]
    nd = max(len(during), 1)
    ck(len(idx) == nd + len(after), '推薦のことばのページが %d 枚（%d 枚のはず）' % (len(idx), nd + len(after)))
    ck(idx[:nd] == list(range(idx[0], idx[0] + nd)) if idx else False,
       '定例会中の推薦のことばのページが続いていない: %s' % idx[:nd])
    lot = next((i for i, p in enumerate(order) if '賞品の抽選' in text(p)[1]), None)
    if after:
        ck(lot is not None, '抽選コーナーのページが無い')
        if lot is not None:
            ck(idx[nd:] == list(range(lot + 1, lot + 1 + len(after))),
               'アフター・定例会後の推薦のことばが抽選コーナーのすぐあとに無い: %s（抽選=%d）' % (idx[nd:], lot))
    if lot is not None and idx:
        ck(idx[nd - 1] < lot, '定例会中の推薦のことばが抽選コーナーより後ろにある')
    for k, i in enumerate(idx):
        sx, t = text(order[i])
        hid = 'show="0"' in sx[:600]
        if k < nd:
            pair, label = (during[k], '推薦のことば（定例会中%d組目）' % (k + 1)) if during else (None, '推薦のことば（ひな形）')
        else:
            pair, label = after[k - nd], '推薦のことば（アフター%d組目）' % (k - nd + 1)
        if pair is None:
            ck(hid, '定例会中の組が無いのに、推薦のことばのページが表示のまま')
            reco_pages.append((order[i], {'giver': {'name': ''}, 'receiver': {'name': ''}}, label))
            continue
        ck(not hid, '%s のページが非表示になっている' % label)
        for who in ('giver', 'receiver'):
            for f in ('name', 'company', 'category'):
                v = pair[who].get(f)
                if v:
                    ck(v in t, '%s のページに「%s」が無い' % (label, v))
            check_fit(order[i], label, [(pair[who].get('name'), 1), (pair[who].get('company'), 1),
                                        (pair[who].get('category'), 2)])
        # 左＝推薦する人・右＝推薦される人（氏名の文字の位置で確かめる）
        gx, rx = [], []
        for sp in re.finditer(r'<p:sp>.*?</p:sp>', sx, re.S):
            st = html.unescape(''.join(re.findall(r'<a:t(?=[\s>])[^>]*>([^<]*)</a:t>', sp.group(0))))
            off = re.search(r'<a:off x="(-?\d+)"', sp.group(0))
            if not off:
                continue
            if pair['giver']['name'] and st.strip() == pair['giver']['name']:
                gx.append(int(off.group(1)))
            if pair['receiver']['name'] and st.strip() == pair['receiver']['name']:
                rx.append(int(off.group(1)))
        if gx and rx:
            ck(max(gx) < min(rx), '%s のページで、推薦する人が左・推薦される人が右になっていない' % label)
        reco_pages.append((order[i], pair, label))
    print('  推薦のことば: 定例会中 %d 組（%s）／アフター・定例会後 %d 組%s'
          % (len(during), '、'.join('%s→%s' % (x['giver']['name'] or '空', x['receiver']['name'] or '空') for x in during) or 'ページは非表示',
             len(after), ('（抽選コーナーのあと: %s）' % '、'.join('%s→%s' % (x['giver']['name'] or '空', x['receiver']['name'] or '空') for x in after)) if after else ''))

# 書記兼会計による報告（更新状況一覧）：表の下の段が、画面の一覧と同じで、2行に収まる大きさか
RENEW = (('90日以内', '更新90'), ('60日以内', '更新60'), ('30日以内', '更新30'), ('期限切れ', '更新超過'))
if any(k in m for _, k in RENEW):
    pg = [p for p in order if '更新を迎えるメンバー' in text(p)[1]]
    ck(len(pg) == 1, '書記兼会計による報告（更新状況）のページが %d 枚' % len(pg))
    seen = []
    for p in pg:
        sx = z.read('ppt/' + p).decode('utf-8')
        for fm in re.finditer(r'<p:graphicFrame>.*?</p:graphicFrame>', sx, re.S):
            seg = fm.group(0)
            rows = re.findall(r'<a:tr\b.*?</a:tr>', seg, re.S)
            if len(rows) < 2:
                continue
            head = ''.join(re.findall(r'<a:t(?=[\s>])[^>]*>([^<]*)</a:t>', rows[0]))
            key = next((k for lb, k in RENEW if lb in head), None)
            if not key or key not in m:
                continue
            want = (m[key] or '').strip() or '該当者なし'
            body = html.unescape(''.join(re.findall(r'<a:t(?=[\s>])[^>]*>([^<]*)</a:t>', rows[1])))
            ck(body == want, '更新状況「%s」が「%s」（「%s」のはず）' % (head.strip(), body[:40], want[:40]))
            szs = [int(v) for v in re.findall(r'<a:rPr\b[^>]*\ssz="(\d+)"', rows[1])]
            hsz = [int(v) for v in re.findall(r'<a:rPr\b[^>]*\ssz="(\d+)"', rows[0])]
            base = (hsz[0] / 100.0) if hsz else 20.0          # 元の大きさ（見出しの段と同じ）
            col = int(re.search(r'<a:gridCol w="(\d+)"', seg).group(1))
            tcpr = re.search(r'<a:tcPr\b[^>]*>', rows[1])

            def mar(k):                           # セルの左右の余白（既定は0.1インチ）
                mm = re.search(r'\s%s="(\d+)"' % k, tcpr.group(0)) if tcpr else None
                return int(mm.group(1)) if mm else 91440
            width_pt = (col - mar('marL') - mar('marR')) / 12700.0
            pt = (szs[0] / 100.0) if szs else base
            # 段落（改行）ごとに、何行になるか（切り上げ）
            paras = [html.unescape(''.join(re.findall(r'<a:t(?=[\s>])[^>]*>([^<]*)</a:t>', q)))
                     for q in re.findall(r'<a:p>.*?</a:p>|<a:p\b[^>]*>.*?</a:p>', rows[1], re.S)]
            lines = sum(max(1, -(-sum(0.5 if ord(c) < 128 else 1.0 for c in q) * pt // width_pt)) for q in paras)
            # 元の大きさで2行ぶんの高さに収まるか（小さくすると1行の高さも縮む）
            ck(lines * pt <= base * 2 + 0.01, '更新状況「%s」が枠に収まらない（%d行・%.1fpt）' % (head.strip(), lines, pt))
            ck(pt <= base, '更新状況「%s」の文字が元より大きい（%.1fpt）' % (head.strip(), pt))
            ck(len(set(szs)) <= 1, '更新状況「%s」の文字の大きさがそろっていない: %s' % (head.strip(), szs))
            # 改行は「、」のうしろだけ（お名前の途中で改行しない）
            ck(all(q.endswith('、') for q in paras[:-1]), '更新状況「%s」がお名前の途中で改行されている: %s' % (head.strip(), paras))
            print('  更新状況 %s: %s（%.1fpt・%d段落・%d行）'
                  % (head.strip()[:5], want[:30] + ('…' if len(want) > 30 else ''), pt, len(paras), lines))
            seen.append(key)
    ck(sorted(seen) == sorted(k for _, k in RENEW if k in m),
       '更新状況の表が %s しか見つからない' % seen)

if photo_of:
    # リファーラル発表：どのページにも、その方の写真だけ
    for k, it in enumerate(plan_rf):
        pg = [p for p in order if 'REFERRAL PRESENTATION' in text(p)[1]]
        if k >= len(pg):
            break
        faces = [w for (_i, _x, _y, _cx, _cy, w) in pics_of(pg[k]) if w]
        want = [it['name']] if it['name'] in photo_of else []
        ck(faces == want, 'リファーラル発表 %d枚目（%s）の写真が %s' % (k + 1, it['name'], faces or 'なし'))
    # 左右にお2人が並ぶページ：左に1人目、右に2人目。ほかの方の写真や、元の写真が残っていないこと
    for title, prefix in (('Main Presenter', 'メインプレゼン'), ('賞品の抽選', '抽選')):
        pg = [p for p in order if title in text(p)[1]]
        if not pg:
            continue
        check_two_photos(pg[0], prefix, (info.get('twoPerson') or {}).get(prefix) or [])
    # 推薦のことば：組ごとのページに、左＝推薦する人・右＝推薦される人
    for p, pair, label in reco_pages:
        check_two_photos(p, label, [pair['giver']['name'], pair['receiver']['name']])
    # 前半に差し込んだメンバーのページ：扉ページは先頭の方、個人ページはその方の写真
    if plan_mp and anchor is not None:
        for k, it in enumerate(plan_mp):
            if anchor + 1 + k >= len(order):
                break
            faces = [q[5] for q in pics_of(order[anchor + 1 + k]) if q[5]]
            want = [it['photoName']] if it.get('photoName') in photo_of else []
            ck(faces == want, 'メンバーのページ %d枚目（%s）の写真が %s' % (k + 1, it.get('photoName'), faces or 'なし'))
    print('  写真: リファーラル発表・左右にお2人のページ・メンバーのページを、1人ずつ中身で照合しました')

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
