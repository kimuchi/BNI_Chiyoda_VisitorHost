#!/usr/bin/env python3
"""事前MTG（朝イチMTG）の既定のひな形を作る。

見本は「BNI週次役職情報共有 パワポ自動生成ツール」で作った事前MTGのpptx
（まとめ1枚＋役職のページ。役職のページは1人のページと2人のページがある）。
そこから次の3枚だけを残し、文字を {{ }} の差し込み口に、写真を仮の画像に置き換える。
デザイン・座標・配色は見本のまま。書体はメイリオにする（テーマの見出し・本文の書体。
メイリオは1行の高さが大きいので、本文の行間を 130% → 110%・115% → 100% に詰める）。

  1枚目 まとめ     … {{月日}}・{{直近のイベント}}・{{お願い事項}}・{{定例会関連}}
  2枚目 2人のページ … {{アイコン1}} {{役職1}}・{{氏名1}}・{{共有事項1}}（右側は2）
  3枚目 1人のページ … {{アイコン1}} {{役職1}}・{{氏名1}}・{{共有事項1}}

写真の枠は「写真1」「写真2」、役職の帯は「帯1」「帯2」という名前にしておく
（サーバーはこの名前で写真・帯を見つける。名前が無いひな形でも位置から見つける）。
実在の方の氏名・写真・書き込みは残さない。

    python3 tools/build_premtg_template.py <見本.pptx> <出力.pptx> [<埋め込み用.html>]

作ってあるひな形に、書体（メイリオ）と行間だけを当て直すときは（見本が手元に無くてもよい）:

    python3 tools/build_premtg_template.py --meiryo docs/templates/BNI_テンプレート_事前MTG.pptx premtg_template.html

埋め込み用.html を渡すと、出力をbase64にしたものも書く（リポジトリ直下の premtg_template.html。
ひな形を登録していなくても事前MTGのパワポを作れるように、Apps Script に同梱する）。
"""
import base64
import io
import re
import sys
import zipfile

from PIL import Image, ImageDraw

SLIDE_W = 12191695
EMU_PT = 12700

CT = {
    'slide': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    'main': 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    'master': 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
    'layout': 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
    'theme': 'application/vnd.openxmlformats-officedocument.theme+xml',
    'presProps': 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
    'viewProps': 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
    'tableStyles': 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
    'core': 'application/vnd.openxmlformats-package.core-properties+xml',
    'app': 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
}
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
PHOTO = 'ppt/media/premtg_photo.png'
FONT = 'メイリオ'


def meiryo_theme(theme):
    """テーマの見出し・本文の書体を、ラテン文字も日本語もメイリオにする（サーバーの setThemeFonts_ と同じ）"""
    def one(m):
        body = re.sub(r'<a:latin\b[^>]*/>', '<a:latin typeface="%s"/>' % FONT, m.group(2), count=1)
        body = re.sub(r'<a:ea\b[^>]*/>', '<a:ea typeface="%s"/>' % FONT, body, count=1)
        if re.search(r'<a:font\s+script="Jpan"[^>]*/>', body):
            body = re.sub(r'<a:font\s+script="Jpan"[^>]*/>', '<a:font script="Jpan" typeface="%s"/>' % FONT, body, count=1)
        else:
            body = re.sub(r'(<a:cs\b[^>]*/>)', r'\1<a:font script="Jpan" typeface="%s"/>' % FONT, body, count=1)
        return '<a:%s>%s</a:%s>' % (m.group(1), body, m.group(1))
    return re.sub(r'<a:(majorFont|minorFont)>(.*?)</a:\1>', one, theme, flags=re.S)


def meiryo_spacing(xml):
    """メイリオは1行の高さが大きいので、本文の行間を詰める（130% → 110%・115% → 100%。2回当てても同じ）"""
    table = {'130000': '110000', '115000': '100000'}
    return re.sub(r'(<a:lnSpc>\s*<a:spcPct val=")(\d+)(")', lambda m: m.group(1) + table.get(m.group(2), m.group(2)) + m.group(3), xml)


def write_html(data, html):
    b64 = base64.b64encode(data).decode('ascii')
    lines = [b64[i:i + 100] for i in range(0, len(b64), 100)]
    with open(html, 'w', encoding='utf-8') as f:
        f.write('PPTX_BASE64_BEGIN\n' + '\n'.join(lines) + '\nPPTX_BASE64_END\n')
    print('埋め込み用: %s（%d 文字）' % (html, len(b64)))


def apply_meiryo(path, html):
    """作ってあるひな形に、書体と行間だけを当て直す（部品の並び・日付はそのまま）"""
    z = zipfile.ZipFile(path)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as out:
        for info in z.infolist():
            data = z.read(info.filename)
            if re.match(r'ppt/theme/theme\d+\.xml$', info.filename):
                data = meiryo_theme(data.decode('utf-8')).encode('utf-8')
            elif re.match(r'ppt/slides/slide\d+\.xml$', info.filename):
                data = meiryo_spacing(data.decode('utf-8')).encode('utf-8')
            out.writestr(zipfile.ZipInfo(info.filename, info.date_time), data, compress_type=zipfile.ZIP_DEFLATED)
    with open(path, 'wb') as f:
        f.write(buf.getvalue())
    print('書体と行間を当て直しました: %s（%d バイト）' % (path, len(buf.getvalue())))
    if html:
        write_html(buf.getvalue(), html)


def text_of(x):
    return ''.join(re.findall(r'<a:t>([^<]*)</a:t>', x))


def shapes(xml):
    """<p:sp>/<p:pic> を上から順に（位置・大きさ・文字つき）"""
    out = []
    for m in re.finditer(r'<p:(sp|pic)>.*?</p:\1>', xml, re.S):
        seg = m.group(0)
        off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"', seg)
        if not off:
            continue
        x, y, cx, cy = (int(v) for v in off.groups())
        out.append({'kind': m.group(1), 'seg': seg, 'start': m.start(), 'end': m.end(),
                    'x': x, 'y': y, 'cx': cx, 'cy': cy, 'text': text_of(seg),
                    'filled': '<a:solidFill><a:srgbClr' in seg.split('</p:spPr>')[0] and '<p:txBody>' not in seg})
    return out


def one_paragraph(seg, text):
    """文字箱の段落を、1段落目の書式のまま text の1段落にする"""
    body = re.search(r'<p:txBody>.*?</p:txBody>', seg, re.S).group(0)
    ps = re.findall(r'<a:p>.*?</a:p>', body, re.S)
    p = ps[0]
    runs = re.findall(r'<a:r>.*?</a:r>', p, re.S)
    first = re.sub(r'<a:t>[^<]*</a:t>', '<a:t>' + text + '</a:t>', runs[0], count=1)
    p = p.replace(''.join(runs), first) if ''.join(runs) in p else re.sub(r'<a:r>.*</a:r>', first, p, flags=re.S)
    new_body = body[:body.index(ps[0])] + p + body[body.rindex(ps[-1]) + len(ps[-1]):]
    return seg.replace(body, new_body)


def set_geom(seg, x=None, y=None, cx=None, cy=None):
    def off(m):
        return '<a:off x="%s" y="%s"/>' % (m.group(1) if x is None else x, m.group(2) if y is None else y)

    def ext(m):
        return '<a:ext cx="%s" cy="%s"/>' % (m.group(1) if cx is None else cx, m.group(2) if cy is None else cy)
    seg = re.sub(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', off, seg, count=1)
    return re.sub(r'<a:ext cx="(\d+)" cy="(\d+)"/>', ext, seg, count=1)


def set_font(seg, sz):
    return re.sub(r'(<a:(?:rPr|endParaRPr)\b[^>]*\ssz=")\d+(")', r'\g<1>%d\2' % sz, seg)


def rename(seg, name, descr=None):
    def f(m):
        attrs = re.sub(r'\s(name|descr)="[^"]*"', '', m.group(1))
        return '<p:cNvPr' + attrs + ' name="' + name + '"' + (' descr="' + descr + '"' if descr else '') + m.group(2)
    return re.sub(r'<p:cNvPr\b([^>]*?)(/?>)', f, seg, count=1)


def replace_segs(xml, pairs):
    """[(元のseg, 新しいseg)] を後ろから入れ替える"""
    for old, new in pairs:
        xml = xml.replace(old, new, 1)
    return xml


def build_summary(xml):
    sh = shapes(xml)
    title = next(s for s in sh if '朝イチMTG' in s['text'])
    body = [s for s in sh if s['y'] > title['y'] + title['cy'] - 1 and s['kind'] == 'sp']
    body.sort(key=lambda s: s['y'])
    if len(body) != 6:
        raise SystemExit('まとめのページの文字箱が6個ではありません: %d' % len(body))
    labels, contents = body[0::2], body[1::2]
    tokens = ['直近のイベント', 'お願い事項', '定例会関連']
    for lab, tok in zip(labels, tokens):
        if tok not in lab['text']:
            raise SystemExit('見出し「%s」が見つかりません（%s）' % (tok, lab['text']))
    # 既定の配置：見出しの高さ・見出しと本文の間・まとまりの間は見本のまま、
    # 本文の枠は下の余白0.2インチまでを 15:42.5:42.5 で分ける（サーバーが文字の量に合わせて並べ直す）
    gap_lc = contents[0]['y'] - (labels[0]['y'] + labels[0]['cy'])
    gap_blk = labels[1]['y'] - (contents[0]['y'] + contents[0]['cy'])
    bottom = 6858000 - 182880
    room = bottom - labels[0]['y'] - sum(l['cy'] for l in labels) - 3 * gap_lc - 2 * gap_blk
    heights = [int(room * 0.15), int(room * 0.425)]
    heights.append(room - sum(heights))
    pairs, y = [], labels[0]['y']
    pairs.append((title['seg'], one_paragraph(title['seg'], '{{月日}} 定例会　朝イチMTG')))
    for i in range(3):
        pairs.append((labels[i]['seg'], set_geom(labels[i]['seg'], y=y)))
        y += labels[i]['cy'] + gap_lc
        seg = one_paragraph(contents[i]['seg'], '{{' + tokens[i] + '}}')
        seg = set_font(set_geom(seg, y=y, cy=heights[i]), 1600)
        pairs.append((contents[i]['seg'], seg))
        y += heights[i] + gap_blk
    return replace_segs(xml, pairs)


def build_role(xml, n):
    """n=1 … 1人のページ、n=2 … 2人のページ"""
    sh = shapes(xml)
    halves = [[s for s in sh if (s['x'] + s['cx'] / 2 < SLIDE_W / 2 or n == 1)]]
    if n == 2:
        halves = [[s for s in sh if s['x'] + s['cx'] / 2 < SLIDE_W / 2 and s['cx'] > 0],
                  [s for s in sh if s['x'] + s['cx'] / 2 >= SLIDE_W / 2 and s['cx'] > 0]]
    pairs = []
    for k, half in enumerate(halves, start=1):
        pics = [s for s in half if s['kind'] == 'pic']
        bands = [s for s in half if s['filled']]
        texts = sorted([s for s in half if s['kind'] == 'sp' and s['text']], key=lambda s: s['y'])
        if len(pics) != 1 or len(bands) != 1 or len(texts) != 3:
            raise SystemExit('役職のページ（%d人）の%d人目の作りが想定と違います: 写真%d・帯%d・文字%d'
                             % (n, k, len(pics), len(bands), len(texts)))
        role, name, body = texts
        pairs.append((pics[0]['seg'], rename(pics[0]['seg'], '写真%d' % k, '写真%d' % k)))
        pairs.append((bands[0]['seg'], rename(bands[0]['seg'], '帯%d' % k)))
        pairs.append((role['seg'], rename(one_paragraph(role['seg'], '{{アイコン%d}} {{役職%d}}' % (k, k)), '役職%d' % k)))
        pairs.append((name['seg'], rename(one_paragraph(name['seg'], '{{氏名%d}}' % k), '氏名%d' % k)))
        pairs.append((body['seg'], rename(one_paragraph(body['seg'], '{{共有事項%d}}' % k), '共有事項%d' % k)))
    return replace_segs(xml, pairs)


def placeholder_png():
    """写真が無い方の仮の画像（灰色の地に白い人の形）"""
    size = 400
    im = Image.new('RGB', (size, size), (208, 211, 216))
    d = ImageDraw.Draw(im)
    d.ellipse((135, 70, 265, 200), fill=(255, 255, 255))
    d.ellipse((60, 215, 340, 495), fill=(255, 255, 255))
    buf = io.BytesIO()
    im.save(buf, 'PNG', optimize=True)
    return buf.getvalue()


def slide_rels(xml):
    rids = sorted(set(re.findall(r'r:embed="(rId\d+)"', xml)), key=lambda r: int(r[3:]))
    rels = ''.join('<Relationship Id="%s" Type="%simage" Target="../media/premtg_photo.png"/>' % (r, REL) for r in rids)
    lay = 'rId%d' % (max([int(r[3:]) for r in rids] or [0]) + 1)
    rels += '<Relationship Id="%s" Type="%sslideLayout" Target="../slideLayouts/slideLayout1.xml"/>' % (lay, REL)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels + '</Relationships>')


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == '--meiryo':
        apply_meiryo(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else '')
        return
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    src, dst = sys.argv[1], sys.argv[2]
    z = zipfile.ZipFile(src)
    names = z.namelist()
    slides = sorted([n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)],
                    key=lambda n: int(re.search(r'(\d+)\.xml$', n).group(1)))
    xmls = {n: z.read(n).decode('utf-8') for n in slides}
    summary = next(n for n in slides if '朝イチMTG' in text_of(xmls[n]))
    two = next(n for n in slides if xmls[n].count('<p:pic>') == 2)
    one = next(n for n in slides if xmls[n].count('<p:pic>') == 1)
    out_slides = [meiryo_spacing(x) for x in (build_summary(xmls[summary]), build_role(xmls[two], 2), build_role(xmls[one], 1))]
    for i, x in enumerate(out_slides, start=1):
        x = re.sub(r'<p:cSld name="[^"]*">', '<p:cSld name="%s">' % ['まとめ', '2人のページ', '1人のページ'][i - 1], x)
        out_slides[i - 1] = x

    prs = z.read('ppt/presentation.xml').decode('utf-8')
    prs = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>',
                 '<p:sldIdLst>' + ''.join('<p:sldId id="%d" r:id="rId%d"/>' % (255 + i, i + 1) for i in (1, 2, 3))
                 + '</p:sldIdLst>', prs, flags=re.S)
    prs = re.sub(r'<p:notesMasterIdLst>.*?</p:notesMasterIdLst>', '', prs, flags=re.S)
    prs_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="%sslideMaster" Target="slideMasters/slideMaster1.xml"/>' % REL
                + ''.join('<Relationship Id="rId%d" Type="%sslide" Target="slides/slide%d.xml"/>' % (i + 1, REL, i) for i in (1, 2, 3))
                + '<Relationship Id="rId5" Type="%spresProps" Target="presProps.xml"/>' % REL
                + '<Relationship Id="rId6" Type="%sviewProps" Target="viewProps.xml"/>' % REL
                + '<Relationship Id="rId7" Type="%stheme" Target="theme/theme1.xml"/>' % REL
                + '<Relationship Id="rId8" Type="%stableStyles" Target="tableStyles.xml"/>' % REL
                + '</Relationships>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="png" ContentType="image/png"/>'
          '<Default Extension="jpeg" ContentType="image/jpeg"/>'
          '<Default Extension="jpg" ContentType="image/jpeg"/>'
          '<Override PartName="/ppt/presentation.xml" ContentType="%s"/>' % CT['main']
          + '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="%s"/>' % CT['master']
          + '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="%s"/>' % CT['layout']
          + '<Override PartName="/ppt/theme/theme1.xml" ContentType="%s"/>' % CT['theme']
          + ''.join('<Override PartName="/ppt/slides/slide%d.xml" ContentType="%s"/>' % (i, CT['slide']) for i in (1, 2, 3))
          + '<Override PartName="/ppt/presProps.xml" ContentType="%s"/>' % CT['presProps']
          + '<Override PartName="/ppt/viewProps.xml" ContentType="%s"/>' % CT['viewProps']
          + '<Override PartName="/ppt/tableStyles.xml" ContentType="%s"/>' % CT['tableStyles']
          + '<Override PartName="/docProps/core.xml" ContentType="%s"/>' % CT['core']
          + '<Override PartName="/docProps/app.xml" ContentType="%s"/>' % CT['app']
          + '</Types>')
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>事前MTG（朝イチMTG）ひな形</dc:title><cp:revision>1</cp:revision>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">2026-09-27T00:00:00Z</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-27T00:00:00Z</dcterms:modified>'
            '</cp:coreProperties>')
    app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
           'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
           '<Application>Microsoft Office PowerPoint</Application><PresentationFormat>ワイド画面</PresentationFormat>'
           '<Slides>3</Slides></Properties>')

    files = [
        ('[Content_Types].xml', ct),
        ('_rels/.rels', z.read('_rels/.rels')),
        ('docProps/app.xml', app),
        ('docProps/core.xml', core),
        ('ppt/presentation.xml', prs),
        ('ppt/_rels/presentation.xml.rels', prs_rels),
    ]
    for p in ['ppt/slideMasters/slideMaster1.xml', 'ppt/slideMasters/_rels/slideMaster1.xml.rels',
              'ppt/slideLayouts/slideLayout1.xml', 'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
              'ppt/theme/theme1.xml', 'ppt/presProps.xml', 'ppt/viewProps.xml', 'ppt/tableStyles.xml']:
        files.append((p, meiryo_theme(z.read(p).decode('utf-8')) if p == 'ppt/theme/theme1.xml' else z.read(p)))
    for i, x in enumerate(out_slides, start=1):
        files.append(('ppt/slides/slide%d.xml' % i, x))
        files.append(('ppt/slides/_rels/slide%d.xml.rels' % i, slide_rels(x)))
    files.append((PHOTO, placeholder_png()))

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as out:
        for name, data in files:
            out.writestr(zipfile.ZipInfo(name, (2026, 9, 27, 0, 0, 0)),
                         data.encode('utf-8') if isinstance(data, str) else data,
                         compress_type=zipfile.ZIP_DEFLATED)
    with open(dst, 'wb') as f:
        f.write(buf.getvalue())
    print('ひな形: %s（%d バイト・3枚）' % (dst, len(buf.getvalue())))

    if len(sys.argv) > 3:
        write_html(buf.getvalue(), sys.argv[3])


if __name__ == '__main__':
    main()
