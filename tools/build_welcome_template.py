#!/usr/bin/env python3
"""事前MTG（朝イチMTG）に入れる「熱烈歓迎」のページの、既定のひな形を作る。

事前MTGの既定のひな形（docs/templates/BNI_テンプレート_事前MTG.pptx）と同じ土台（マスター・レイアウト・テーマ）に、
熱烈歓迎のページを1枚だけ置いたもの。土台が同じなので、事前MTGのパワポに入れても見た目は変わらない。
実在の方の氏名・写真は入れない（写真は仮の画像。お名前などは {{ }} の差し込み口）。

  見出し（赤い帯）… 🎉 熱烈歓迎 🎉
  写真（丸）       … 名前が「写真」の画像
  {{カテゴリー}} ／ {{氏名}}さん ／ {{会社名}}
  ようこそ、{{チャプター}}へ！ ／ {{開催日}} 定例会

    python3 tools/build_welcome_template.py docs/templates/BNI_テンプレート_事前MTG.pptx \\
        docs/templates/BNI_テンプレート_熱烈歓迎.pptx welcome_template.html

埋め込み用.html を渡すと、出力をbase64にしたものも書く（リポジトリ直下の welcome_template.html。
ひな形を登録していなくても熱烈歓迎のページを作れるように、Apps Script に同梱する）。
"""
import base64
import io
import re
import sys
import zipfile

from PIL import Image, ImageDraw

REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
PHOTO = 'ppt/media/welcome_photo.png'
RED = 'C8102E'


def placeholder_png():
    """写真が無い方の仮の画像（灰色の地に白い人の形。事前MTGのひな形と同じ）"""
    size = 400
    im = Image.new('RGB', (size, size), (208, 211, 216))
    d = ImageDraw.Draw(im)
    d.ellipse((135, 70, 265, 200), fill=(255, 255, 255))
    d.ellipse((60, 215, 340, 495), fill=(255, 255, 255))
    buf = io.BytesIO()
    im.save(buf, 'PNG', optimize=True)
    return buf.getvalue()


def xfrm(x, y, cx, cy):
    return '<a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>' % (x, y, cx, cy)


def text_box(sid, name, x, y, cx, cy, text, sz, color, bold=False, algn='l', anchor='ctr'):
    b = ' b="1"' if bold else ''
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="%s">'
            '<a:noAutofit/></a:bodyPr><a:lstStyle/>'
            '<a:p><a:pPr algn="%s" indent="0" marL="0"><a:buNone/></a:pPr>'
            '<a:r><a:rPr lang="ja-JP" altLang="en-US" sz="%d"%s dirty="0"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:rPr>'
            '<a:t>%s</a:t></a:r><a:endParaRPr lang="ja-JP" altLang="en-US" sz="%d"%s dirty="0"/></a:p></p:txBody></p:sp>'
            % (sid, name, xfrm(x, y, cx, cy), anchor, algn, sz * 100, b, color, text, sz * 100, b))


def welcome_slide(w, h):
    band = 1600200
    shapes = [
        # 赤い帯と見出し
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="帯"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>'
        % (xfrm(0, 0, w, band), RED),
        text_box(3, '見出し', 0, 0, w, band, '🎉 熱烈歓迎 🎉', 54, 'FFFFFF', bold=True, algn='ctr'),
        # 写真（丸）
        '<p:pic><p:nvPicPr><p:cNvPr id="4" name="写真" descr="写真"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        '<p:spPr>%s<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:ln w="57150"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln></p:spPr></p:pic>'
        % (xfrm(1005840, 2011680, 3291840, 3291840), RED),
        # お名前・会社名・カテゴリー
        text_box(5, 'カテゴリー', 4754880, 2148840, 6858000, 640080, '{{カテゴリー}}', 28, '555555'),
        text_box(6, '氏名', 4754880, 2788920, 6858000, 1188720, '{{氏名}}さん', 66, '222222', bold=True),
        text_box(7, '会社名', 4754880, 3977640, 6858000, 640080, '{{会社名}}', 28, '333333'),
        # ようこそ・日付
        text_box(8, 'ようこそ', 0, 5532120, w, 777240, 'ようこそ、{{チャプター}}へ！', 36, RED, bold=True, algn='ctr'),
        text_box(9, '日付', w - 4206240, h - 548640, 4114800, 457200, '{{開催日}} 定例会', 16, '777777', algn='r'),
    ]
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<p:sld %s><p:cSld name="熱烈歓迎"><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '%s</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>' % (NS, ''.join(shapes)))


def main():
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    src, dst = sys.argv[1], sys.argv[2]
    z = zipfile.ZipFile(src)
    prs = z.read('ppt/presentation.xml').decode('utf-8')
    sz = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', prs)
    w, h = int(sz.group(1)), int(sz.group(2))
    prs = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>', prs, flags=re.S)
    prs_rels = z.read('ppt/_rels/presentation.xml.rels').decode('utf-8')
    prs_rels = re.sub(r'<Relationship\b[^>]*Type="[^"]*/slide"[^>]*/>', '', prs_rels)
    if 'Id="rId2"' in prs_rels:
        raise SystemExit('presentation.xml.rels の rId2 がページ以外に使われています')
    prs_rels = prs_rels.replace('</Relationships>', '<Relationship Id="rId2" Type="%sslide" Target="slides/slide1.xml"/></Relationships>' % REL)
    ct = z.read('[Content_Types].xml').decode('utf-8')
    ct = re.sub(r'<Override PartName="/ppt/slides/slide\d+\.xml"[^>]*/>', '', ct)
    ct = ct.replace('</Types>', '<Override PartName="/ppt/slides/slide1.xml" '
                    'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>')
    core = z.read('docProps/core.xml').decode('utf-8')
    core = re.sub(r'<dc:title>.*?</dc:title>', '<dc:title>熱烈歓迎（事前MTG）ひな形</dc:title>', core)
    app = re.sub(r'<Slides>\d+</Slides>', '<Slides>1</Slides>', z.read('docProps/app.xml').decode('utf-8'))
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="%sslideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
            '<Relationship Id="rId2" Type="%simage" Target="../media/welcome_photo.png"/></Relationships>' % (REL, REL))

    files = [('[Content_Types].xml', ct), ('_rels/.rels', z.read('_rels/.rels')), ('docProps/app.xml', app),
             ('docProps/core.xml', core), ('ppt/presentation.xml', prs), ('ppt/_rels/presentation.xml.rels', prs_rels)]
    keep = [n for n in z.namelist()
            if re.match(r'ppt/(slideMasters|slideLayouts|theme)/', n) or n in ('ppt/presProps.xml', 'ppt/viewProps.xml', 'ppt/tableStyles.xml')]
    for n in keep:
        files.append((n, z.read(n)))
    files += [('ppt/slides/slide1.xml', welcome_slide(w, h)), ('ppt/slides/_rels/slide1.xml.rels', rels), (PHOTO, placeholder_png())]

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as out:
        for name, data in files:
            out.writestr(zipfile.ZipInfo(name, (2026, 10, 3, 0, 0, 0)),
                         data.encode('utf-8') if isinstance(data, str) else data,
                         compress_type=zipfile.ZIP_DEFLATED)
    with open(dst, 'wb') as f:
        f.write(buf.getvalue())
    print('ひな形: %s（%d バイト・1枚）' % (dst, len(buf.getvalue())))

    if len(sys.argv) > 3:
        b64 = base64.b64encode(buf.getvalue()).decode('ascii')
        lines = [b64[i:i + 100] for i in range(0, len(b64), 100)]
        with open(sys.argv[3], 'w', encoding='utf-8') as f:
            f.write('PPTX_BASE64_BEGIN\n' + '\n'.join(lines) + '\nPPTX_BASE64_END\n')
        print('埋め込み用: %s（%d 文字）' % (sys.argv[3], len(b64)))


if __name__ == '__main__':
    main()
