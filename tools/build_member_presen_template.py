#!/usr/bin/env python3
"""メンバープレゼンのテンプレートを、必要なものだけに作り直す。

もとのファイルは61枚・14.7MBあるが、生成に使うのは
「業種区分の扉ページ」と「個人ページ」の2枚だけ。
残りのページ・レイアウト・写真は出来上がりに一切出ないのに、
毎回サーバーが読み書きすることになり、ただ遅くなる。

このスクリプトは、
  ・使う2枚とその土台（レイアウト・マスター・テーマ・画像）だけを残す
  ・実在の方の氏名・会社名・写真を、差し込み用の見本に置き換える
  ・30秒カウントダウンを自動で始め、終わったら次のスライドへ進むようにする
を行う。デザイン・座標・フォント・配色には触れない。

    python3 tools/build_member_presen_template.py <元のpptx> <出力先pptx>
"""
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 残す2枚と、その土台
SRC_OVERVIEW = 'ppt/slides/slide1.xml'
SRC_INDIVIDUAL = 'ppt/slides/slide3.xml'
LAYOUT = 'ppt/slideLayouts/slideLayout16.xml'
MASTER = 'ppt/slideMasters/slideMaster2.xml'
THEME = 'ppt/theme/theme2.xml'
EXTRA = ['ppt/presProps.xml', 'ppt/viewProps.xml', 'ppt/tableStyles.xml',
         'docProps/app.xml', 'docProps/core.xml']

CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def rels_path(part):
    d, b = posixpath.split(part)
    return posixpath.join(d, '_rels', b + '.rels')


def rel_targets(z, part):
    rp = rels_path(part)
    if rp not in z.namelist():
        return []
    out = []
    d = posixpath.dirname(part)
    for m in re.findall(r'<Relationship\b[^>]*>', z.read(rp).decode('utf-8')):
        if 'External' in m:
            continue
        t = re.search(r'Target="([^"]+)"', m).group(1)
        ty = re.search(r'Type="[^"]*/(\w+)"', m).group(1)
        out.append((ty, posixpath.normpath(posixpath.join(d, t))))
    return out


def placeholder_png(path, w, h):
    """写真の見本。差し込み時に必ず入れ替わるので、誰の顔も入れない。"""
    from PIL import Image, ImageDraw
    img = Image.new('RGB', (w, h), (237, 237, 237))
    d = ImageDraw.Draw(img)
    cx, r = w / 2, w * 0.20
    d.ellipse([cx - r, h * 0.26 - r, cx + r, h * 0.26 + r], fill=(208, 208, 208))
    bw, bh = w * 0.62, h * 0.42
    d.rounded_rectangle([cx - bw / 2, h * 0.54, cx + bw / 2, h * 0.54 + bh],
                        radius=int(w * 0.14), fill=(208, 208, 208))
    img.save(path, 'PNG', optimize=True)


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    src, dst = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
    z = zipfile.ZipFile(src)
    names = set(z.namelist())

    # 1) 残すパーツを、関係をたどって集める（レイアウトは16だけ）
    keep = set()

    def walk(p):
        if p in keep or p not in names:
            return
        keep.add(p)
        rp = rels_path(p)
        if rp in names:
            keep.add(rp)
        for ty, t in rel_targets(z, p):
            if ty == 'slideLayout' and t != LAYOUT:
                continue                       # マスターがぶら下げている他のレイアウトは捨てる
            walk(t)

    for p in [SRC_OVERVIEW, SRC_INDIVIDUAL, LAYOUT, MASTER, THEME] + EXTRA:
        walk(p)
    # 入れ物そのもの（この4つは関係をたどっても出てこないので、名指しで入れる）
    for p in ['[Content_Types].xml', '_rels/.rels',
              'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels']:
        keep.add(p)

    work = tempfile.mkdtemp(prefix='mptpl_')
    try:
        # 2) 展開する。個人ページは slide2 に詰め直す
        rename = {SRC_INDIVIDUAL: 'ppt/slides/slide2.xml',
                  rels_path(SRC_INDIVIDUAL): 'ppt/slides/_rels/slide2.xml.rels'}
        parts = {}
        for p in sorted(keep):
            out = rename.get(p, p)
            fp = os.path.join(work, out)
            os.makedirs(os.path.dirname(fp), exist_ok=True)
            with open(fp, 'wb') as f:
                f.write(z.read(p))
            parts[out] = fp

        # 3) 写真の見本を差し替える（実在の方の顔を入れたままにしない）
        ph = 'ppt/media/mp_photo_placeholder.png'
        placeholder_png(os.path.join(work, ph), 900, 1100)
        parts[ph] = os.path.join(work, ph)
        for slide, rid in (('ppt/slides/_rels/slide1.xml.rels', 'rId3'),
                           ('ppt/slides/_rels/slide2.xml.rels', 'rId2')):
            x = open(parts[slide], encoding='utf-8').read()
            x = re.sub(r'(<Relationship\b[^>]*\bId="%s"[^>]*\bTarget=")[^"]*(")' % rid,
                       r'\1../media/%s\2' % posixpath.basename(ph), x)
            open(parts[slide], 'w', encoding='utf-8').write(x)
        for old in ('ppt/media/image12.jpeg', 'ppt/media/image16.jpeg'):
            parts.pop(old, None)

        # 4) スライドの文字と動きを書き換える（本番と同じ ooxml.js を使う）
        r = subprocess.run(['node', 'tools/mp_make_template.js', work], cwd=ROOT)
        if r.returncode != 0:
            sys.exit(r.returncode)

        # 5) マスターが持つレイアウトを1つだけにする
        mrels = parts[rels_path(MASTER)]
        x = open(mrels, encoding='utf-8').read()
        keep_rid = re.search(r'<Relationship\b[^>]*\bId="([^"]+)"[^>]*slideLayout16\.xml"[^>]*/>', x).group(1)
        x = re.sub(r'<Relationship\b[^>]*slideLayouts/slideLayout(?!16\.xml")\d+\.xml"[^>]*/>', '', x)
        open(mrels, 'w', encoding='utf-8').write(x)
        mx = parts[MASTER]
        x = open(mx, encoding='utf-8').read()
        one = re.search(r'<p:sldLayoutId\b[^>]*r:id="%s"[^>]*/>' % keep_rid, x).group(0)
        x = re.sub(r'<p:sldLayoutIdLst>.*?</p:sldLayoutIdLst>',
                   '<p:sldLayoutIdLst>%s</p:sldLayoutIdLst>' % one, x, flags=re.S)
        open(mx, 'w', encoding='utf-8').write(x)

        # 6) presentation.xml と、その関係を作り直す
        px = parts['ppt/presentation.xml']
        x = open(px, encoding='utf-8').read()
        x = re.sub(r'<p:sldMasterIdLst>.*?</p:sldMasterIdLst>',
                   '<p:sldMasterIdLst><p:sldMasterId id="2147483665" r:id="rId1"/></p:sldMasterIdLst>',
                   x, flags=re.S)
        x = re.sub(r'<p:notesMasterIdLst>.*?</p:notesMasterIdLst>', '', x, flags=re.S)
        x = re.sub(r'<p:handoutMasterIdLst>.*?</p:handoutMasterIdLst>', '', x, flags=re.S)
        x = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>',
                   '<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst>',
                   x, flags=re.S)
        open(px, 'w', encoding='utf-8').write(x)

        rel = lambda i, ty, tgt: ('<Relationship Id="rId%d" Type="%s/%s" Target="%s"/>' % (i, REL_NS, ty, tgt))
        # 入れ物の関係。サムネイルやカスタムプロパティは残さないので、ここも作り直す
        open(parts['_rels/.rels'], 'w', encoding='utf-8').write(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + rel(1, 'officeDocument', 'ppt/presentation.xml')
            + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/'
              'relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            + rel(3, 'extended-properties', 'docProps/app.xml')
            + '</Relationships>')
        open(parts['ppt/_rels/presentation.xml.rels'], 'w', encoding='utf-8').write(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + rel(1, 'slideMaster', 'slideMasters/slideMaster2.xml')
            + rel(2, 'slide', 'slides/slide1.xml')
            + rel(3, 'slide', 'slides/slide2.xml')
            + rel(4, 'presProps', 'presProps.xml')
            + rel(5, 'viewProps', 'viewProps.xml')
            + rel(6, 'theme', 'theme/theme2.xml')
            + rel(7, 'tableStyles', 'tableStyles.xml')
            + '</Relationships>')

        # 7) [Content_Types].xml を、残したパーツだけで作り直す
        ct = open(parts['[Content_Types].xml'], encoding='utf-8').read()
        defaults = re.findall(r'<Default\b[^>]*/>', ct)
        exts = {re.search(r'Extension="([^"]+)"', d).group(1).lower() for d in defaults}
        if 'png' not in exts:
            defaults.append('<Default Extension="png" ContentType="image/png"/>')
        types = {
            'ppt/presentation.xml': 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
            'ppt/presProps.xml': 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
            'ppt/viewProps.xml': 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
            'ppt/tableStyles.xml': 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
            'docProps/app.xml': 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
            'docProps/core.xml': 'application/vnd.openxmlformats-package.core-properties+xml',
        }
        ov = []
        for p in sorted(parts):
            if p.endswith('.rels') or p.startswith('ppt/media/') or p == '[Content_Types].xml':
                continue
            if p in types:
                t = types[p]
            elif p.startswith('ppt/slides/'):
                t = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
            elif p.startswith('ppt/slideLayouts/'):
                t = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
            elif p.startswith('ppt/slideMasters/'):
                t = 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
            elif p.startswith('ppt/theme/'):
                t = 'application/vnd.openxmlformats-officedocument.theme+xml'
            else:
                raise SystemExit('種類の分からないパートがあります: %s' % p)
            ov.append('<Override PartName="/%s" ContentType="%s"/>' % (p, t))
        open(parts['[Content_Types].xml'], 'w', encoding='utf-8').write(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="%s">%s%s</Types>' % (CT_NS, ''.join(defaults), ''.join(ov)))

        # 8) 固める
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as out:
            for p in sorted(parts):
                out.write(parts[p], p)
    finally:
        shutil.rmtree(work, ignore_errors=True)

    before, after = os.path.getsize(src), os.path.getsize(dst)
    print('%s' % os.path.basename(dst))
    print('  %d パーツ → %d パーツ／ %.1fMB → %.1fMB（%d%%）'
          % (len(names), len(parts), before / 1048576, after / 1048576, round(after / before * 100)))


if __name__ == '__main__':
    main()
