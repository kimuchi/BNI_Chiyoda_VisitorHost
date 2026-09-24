#!/usr/bin/env python3
"""ハーネスが作った pptx を組み立て、中身を検証する。

PowerPointで開けるかはここでは確かめられないので、代わりに
  ・zipとして正しいか／全パートがXMLとして読めるか
  ・関係ファイル(.rels)の参照先が実在するか（切れたリンクが無いか）
  ・[Content_Types] とスライドの数・順番が合っているか
  ・各スライドの文字・座標・写真の切り抜きが指示どおりか
を機械的に確かめる。
"""
import json
import os
import re
import sys
import zipfile
from xml.dom import minidom

WORK, OUT = sys.argv[1], sys.argv[2]
plan = json.load(open(os.path.join(OUT, 'plan.json')))
items = json.load(open(os.path.join(WORK, 'items.json')))
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

dst = os.path.join(OUT, 'result.pptx')
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
    for p, spec in plan['plan'].items():
        src = os.path.join(OUT, 'gen', p) if spec['from'] == 'gen' else os.path.join(WORK, 'parts', spec['src'])
        z.write(src, p)

fails, checks = [], 0


def ck(cond, msg):
    global checks
    checks += 1
    if not cond:
        fails.append(msg)


WITH_PHOTO = set(json.load(open(os.path.join(WORK, 'photo_names.json'))))

z = zipfile.ZipFile(dst)
names = set(z.namelist())
ck(z.testzip() is None, 'zipが壊れている')

# すべてのXMLが読めるか
for n in sorted(names):
    if n.endswith('.xml') or n.endswith('.rels'):
        try:
            minidom.parseString(z.read(n))
        except Exception as e:
            fails.append('XMLとして読めない: %s (%s)' % (n, e))
        checks += 1

# 関係ファイルの参照先が実在するか
for n in sorted(names):
    if not n.endswith('.rels'):
        continue
    base = os.path.dirname(os.path.dirname(n))
    for m in re.finditer(r'<Relationship\b[^>]*>', z.read(n).decode('utf-8')):
        tag = m.group(0)
        if 'External' in tag:
            continue
        t = re.search(r'Target="([^"]+)"', tag).group(1)
        tgt = os.path.normpath(os.path.join(base, t)).replace('\\', '/')
        ck(tgt in names, '%s の参照先がない: %s' % (n, tgt))

# スライドの数と順番
pres = z.read('ppt/presentation.xml').decode('utf-8')
prels = z.read('ppt/_rels/presentation.xml.rels').decode('utf-8')
rid2tgt = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', prels))
sld = re.findall(r'<p:sldId id="\d+" r:id="([^"]+)"/>', pres)
ck(len(sld) == len(items), 'スライド数が合わない: %d / 指示は %d' % (len(sld), len(items)))
order = [rid2tgt[r] for r in sld]
ck(order == ['slides/slide%d.xml' % (i + 1) for i in range(len(items))], 'スライドの順番が指示どおりでない')
ct = z.read('[Content_Types].xml').decode('utf-8')
ck(ct.count('presentationml.slide+xml') == len(items),
   '[Content_Types]のスライド定義が %d 件（%d 件のはず）' % (ct.count('presentationml.slide+xml'), len(items)))
ck('notesSlide' not in ct, '[Content_Types]にノートの定義が残っている')
ck(not [n for n in names if 'notesSlides/' in n], 'ノートのパートが残っている')

# [Content_Types] の Override が実在するパートを指しているか／重複が無いか
ovr = re.findall(r'<Override PartName="([^"]+)"', ct)
ck(len(ovr) == len(set(ovr)), '[Content_Types]に同じパートの定義が2つある')
for o in ovr:
    ck(o.lstrip('/') in names, '[Content_Types]の定義にパートが無い: %s' % o)
# 拡張子の既定が無いパートは Override が要る（画像が表示されなくなるため）
defs = {e.lower() for e in re.findall(r'<Default Extension="([^"]+)"', ct)}
for n in sorted(names):
    ext = n.rsplit('.', 1)[-1].lower()
    ck(ext in defs or ('/' + n) in ovr, '種類が決まっていないパート: %s' % n)



def img_size(path):
    import struct
    b = open(path, 'rb').read(64 * 1024)
    if b[:4] == b'\x89PNG':
        return struct.unpack('>II', b[16:24])
    if b[:2] == b'\xff\xd8':
        i = 2
        while i + 9 < len(b):
            if b[i] != 0xFF:
                i += 1
                continue
            mk = b[i + 1]
            if mk == 0xFF:
                i += 1
                continue
            if mk == 0x01 or 0xD0 <= mk <= 0xD9:
                i += 2
                continue
            ln = (b[i + 2] << 8) | b[i + 3]
            if 0xC0 <= mk <= 0xCF and mk not in (0xC4, 0xC8, 0xCC):
                return ((b[i + 7] << 8) | b[i + 8], (b[i + 5] << 8) | b[i + 6])
            if mk == 0xDA:
                break
            i += 2 + ln
    return (0, 0)


def cover_crop(sw, sh_, bw, bh):
    """はみ出す分を左右または上下から均等に切る量（10万分率）。ooxml.js の coverCrop_ と同じ式。"""
    if not (sw and sh_ and bw and bh):
        return (0, 0, 0, 0)
    sr, tr = sw / sh_, bw / bh
    if sr > tr:
        f = round((1 - tr / sr) / 2 * 100000)
        return (f, 0, f, 0)
    if sr < tr:
        g = round((1 - sr / tr) / 2 * 100000)
        return (0, g, 0, g)
    return (0, 0, 0, 0)


def shapes(doc):
    tree = doc.getElementsByTagNameNS(NS_P, 'spTree')[0]
    out = {}
    for ch in tree.childNodes:
        if ch.nodeType != 1 or ch.localName not in ('sp', 'pic', 'graphicFrame', 'grpSp'):
            continue
        cn = ch.getElementsByTagNameNS(NS_P, 'cNvPr')
        if cn:
            out[cn[0].getAttribute('id')] = ch
    return out


def paras(el):
    tb = el.getElementsByTagNameNS(NS_P, 'txBody') or el.getElementsByTagNameNS(NS_A, 'txBody')
    if not tb:
        return []
    return [''.join(t.firstChild.nodeValue if t.firstChild else ''
                    for t in p.getElementsByTagNameNS(NS_A, 't'))
            for p in tb[0].getElementsByTagNameNS(NS_A, 'p')]


def geom(el):
    xf = el.getElementsByTagNameNS(NS_A, 'xfrm')
    if not xf:
        return None
    o = xf[0].getElementsByTagNameNS(NS_A, 'off')[0]
    e = xf[0].getElementsByTagNameNS(NS_A, 'ext')[0]
    return (int(o.getAttribute('x')), int(o.getAttribute('y')),
            int(e.getAttribute('cx')), int(e.getAttribute('cy')))


for i, it in enumerate(items):
    n = i + 1
    doc = minidom.parseString(z.read('ppt/slides/slide%d.xml' % n))
    sh = shapes(doc)
    rels = z.read('ppt/slides/_rels/slide%d.xml.rels' % n).decode('utf-8')
    tag = '[扉]' if it['kind'] == 'overview' else '[個]'
    if it['kind'] == 'overview':
        ck(paras(sh['11']) == [it['block']], '%d %s 業種区分名が違う: %s' % (n, tag, paras(sh['11'])))
        ck(paras(sh['31']) == [it['nextName']], '%d %s Next氏名が違う: %s' % (n, tag, paras(sh['31'])))
        tbl = sh['6'].getElementsByTagNameNS(NS_A, 'tbl')[0]
        rows = tbl.getElementsByTagNameNS(NS_A, 'tr')
        ck(len(rows) == 8, '%d %s 表の行数が %d' % (n, tag, len(rows)))
        for r in range(7):
            want = it['rows'][r] if r < len(it['rows']) else {'title': '', 'name': ''}
            cells = rows[r + 1].getElementsByTagNameNS(NS_A, 'tc')
            got = (paras(cells[0]), paras(cells[1]))
            ck(got == ([want['title']], [want['name']]),
               '%d %s 表%d行目が %s（%s のはず）' % (n, tag, r + 1, got, want))
    else:
        ck(paras(sh['56']) == [it['name']], '%d %s 氏名が違う: %s' % (n, tag, paras(sh['56'])))
        ck(paras(sh['2']) == it['companyLines'], '%d %s 会社名が違う: %s' % (n, tag, paras(sh['2'])))
        ck(paras(sh['12']) == it['categoryLines'], '%d %s カテゴリーが違う: %s' % (n, tag, paras(sh['12'])))
        g = geom(sh['2'])
        cg = it['companyGeom']
        want = (cg['x'], cg['y'], cg['cx'], cg['cy'])
        ck(g == want, '%d %s 会社名の枠が %s（%s のはず）' % (n, tag, g, want))
        gc = geom(sh['12'])
        ck(gc[0] == 4830481 and gc[1] == it['categoryTop'] and gc[2] == 7311519,
           '%d %s カテゴリーの枠が %s（x=4830481 y=%d cx=7311519 のはず）' % (n, tag, gc, it['categoryTop']))
        # 会社名の下に【カテゴリー】が来ていること（3行になったときに重なっていた）
        ck(it['categoryTop'] >= g[1] + g[3] - 200000,
           '%d %s 【カテゴリー】が会社名の枠に食い込んでいる（枠の下端 %d / カテゴリー上端 %d）'
           % (n, tag, g[1] + g[3], it['categoryTop']))
        # 30秒カウントダウンの数字（上端 4414085）に重ならないこと
        ck(it['categoryTop'] + 584775 <= 4700000,
           '%d %s 【カテゴリー】がカウントダウンの数字に重なる（上端 %d）' % (n, tag, it['categoryTop']))
        bp = sh['12'].getElementsByTagNameNS(NS_A, 'bodyPr')[0]
        ck(bp.getAttribute('anchor') == 't', '%d %s カテゴリーが上寄せになっていない' % (n, tag))
        ck(len(bp.getElementsByTagNameNS(NS_A, 'noAutofit')) == 1, '%d %s カテゴリーの自動調整が切れていない' % (n, tag))
        ck(not bp.getElementsByTagNameNS(NS_A, 'spAutoFit'), '%d %s カテゴリーに spAutoFit が残っている' % (n, tag))
        if it.get('categoryPt'):
            szs = {r.getAttribute('sz') for r in sh['12'].getElementsByTagNameNS(NS_A, 'rPr')}
            ck(szs == {str(int(it['categoryPt'] * 100))},
               '%d %s カテゴリーの文字サイズが %s（%d のはず）' % (n, tag, szs, it['categoryPt'] * 100))
        if it.get('companyPt'):
            szs = {r.getAttribute('sz') for r in sh['2'].getElementsByTagNameNS(NS_A, 'rPr')}
            ck(szs == {str(int(it['companyPt'] * 100))}, '%d %s 会社名の文字サイズが %s' % (n, tag, szs))
        if it.get('categoryTight'):
            pcts = {p.getAttribute('val') for p in sh['12'].getElementsByTagNameNS(NS_A, 'spcPct')}
            ck(pcts == {'85000'}, '%d %s カテゴリーの行間が %s（85000 のはず）' % (n, tag, pcts))
        if it['nextName']:
            ck(paras(sh['14']) == [it['nextName']], '%d %s NEXT氏名が違う: %s' % (n, tag, paras(sh['14'])))
        else:
            ck('14' not in sh and '6' not in sh, '%d %s 最後の人なのにNEXTが残っている' % (n, tag))
        # カウントダウンが自動で始まり、終わったら次のスライドへ進むこと
        sx = z.read('ppt/slides/slide%d.xml' % n).decode('utf-8')
        sec = it.get('countdownSec') or 30
        ck('<p:cond delay="indefinite"/>' not in sx,
           '%d %s カウントダウンがクリック待ちのまま' % (n, tag))
        adv = set(re.findall(r'advTm="(\d+)"', sx))
        ck(adv == {str((sec + 1) * 1000)},
           '%d %s 自動で次へ進む時間が %s（%d のはず）' % (n, tag, adv or 'なし', (sec + 1) * 1000))
        ck(len(re.findall(r'<p:spTgt spid="\d+"/>', sx)) == sec,
           '%d %s カウントダウンの手順が %d 回（%d 回のはず）'
           % (n, tag, len(re.findall(r'<p:spTgt spid="\d+"/>', sx)), sec))
        # 数字の箱が「残り最大」から「0」まで揃っていること
        want = ['%d:%02d' % (t // 60, t % 60) if sec >= 60 else str(t) for t in range(sec, -1, -1)]
        doc = minidom.parseString(sx)
        nums = []
        for sp2 in doc.getElementsByTagNameNS(NS_P, 'sp'):
            # <a:extLst> の中にも <a:ext> があるので、必ず <a:xfrm> の中を見る
            xf = sp2.getElementsByTagNameNS(NS_A, 'xfrm')
            ex = xf[0].getElementsByTagNameNS(NS_A, 'ext') if xf else []
            if not ex or ex[0].getAttribute('cx') != '2878814':
                continue
            nums.append(''.join(t.firstChild.nodeValue if t.firstChild else ''
                                for t in sp2.getElementsByTagNameNS(NS_A, 't')))
        ck(nums == list(reversed(want)),
           '%d %s カウントダウンの数字が %s…%s（%s…%s のはず）'
           % (n, tag, nums[:2], nums[-2:], list(reversed(want))[:2], list(reversed(want))[-2:]))

    # 写真
    pid = '2' if it['kind'] == 'overview' else '3'
    rid = 'rId3' if it['kind'] == 'overview' else 'rId2'
    if it['photoName'].replace('　', '') in WITH_PHOTO:
        ck(pid in sh, '%d %s 写真の図形が消えている' % (n, tag))
        tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels)
        ck(tgt is not None and 'mpphoto' in tgt.group(1),
           '%d %s 写真の差し替え先が %s' % (n, tag, tgt.group(1) if tgt else 'なし'))
        sr = sh[pid].getElementsByTagNameNS(NS_A, 'srcRect')
        ck(len(sr) == 1, '%d %s srcRectが %d 個' % (n, tag, len(sr)))
        # 切り抜きが「縦横比を変えずに枠を埋める」量になっているか、実物の写真から検算する
        if sr and tgt:
            src = os.path.normpath(os.path.join('ppt/slides', tgt.group(1))).replace('\\', '/')
            open(os.path.join(OUT, '_photo.tmp'), 'wb').write(z.read(src))
            dim = img_size(os.path.join(OUT, '_photo.tmp'))
            g = geom(sh[pid])
            want = cover_crop(dim[0], dim[1], g[2], g[3])
            got = tuple(int(sr[0].getAttribute(k) or 0) for k in 'ltrb')
            ck(all(abs(a - b) <= 1 for a, b in zip(got, want)),
               '%d %s 写真の切り抜きが %s（%s のはず / 写真%sx%s 枠%sx%s）'
               % (n, tag, got, want, dim[0], dim[1], g[2], g[3]))
    else:
        ck(pid not in sh, '%d %s 写真が無い人なのに図形が残っている' % (n, tag))
        ck(('Id="%s"' % rid) not in rels, '%d %s 写真が無い人なのに関係が残っている' % (n, tag))

tmp = os.path.join(OUT, '_photo.tmp')
if os.path.exists(tmp):
    os.remove(tmp)

# 「保存済みのタイミングを使用」が有効か（これが無いと自動送りは効かない）
pp = z.read('ppt/presProps.xml').decode('utf-8')
ck(re.search(r'<p:showPr\b[^>]*\buseTimings="1"', pp) is not None,
   'スライドショーが保存済みのタイミングを使う設定になっていない（advTmが無視される）')

sz = os.path.getsize(dst)
print('出来上がり: %s  %.2f MB  %d パーツ  %d スライド' % (os.path.basename(dst), sz / 1048576, len(names), len(items)))
print('検査 %d 件' % checks)
if fails:
    print('NG: %d 件' % len(fails))
    for f in fails[:40]:
        print('   ', f)
    sys.exit(1)
print('OK: すべて指示どおりです')
