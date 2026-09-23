# -*- coding: utf-8 -*-
"""実物のpptxから1枚だけ取り出して、単独のテンプレートpptxにする。

GAS側は必ず ppt/slides/slide1.xml をひな形として読むため、
取り出したスライドを slide1 に付け替える。
デザイン・画像・動画・フォント・レイアウトは元のまま触らない。
"""
import re, shutil, sys, zipfile

REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def extract_slide(src, n, dst, new_title=None, title_id='15'):
    z = zipfile.ZipFile(src)
    names = z.namelist()
    n_slides = len([x for x in names if re.match(r'ppt/slides/slide\d+\.xml$', x)])

    # 残すスライドに紐づく notesSlide の番号を調べる
    keep_notes = None
    rels_path = 'ppt/slides/_rels/slide%d.xml.rels' % n
    if rels_path in names:
        m = re.search(r'Target="\.\./notesSlides/notesSlide(\d+)\.xml"', z.read(rels_path).decode('utf-8'))
        if m:
            keep_notes = int(m.group(1))

    rename = {
        'ppt/slides/slide%d.xml' % n: 'ppt/slides/slide1.xml',
        'ppt/slides/_rels/slide%d.xml.rels' % n: 'ppt/slides/_rels/slide1.xml.rels',
    }
    if keep_notes:
        rename['ppt/notesSlides/notesSlide%d.xml' % keep_notes] = 'ppt/notesSlides/notesSlide1.xml'
        rename['ppt/notesSlides/_rels/notesSlide%d.xml.rels' % keep_notes] = 'ppt/notesSlides/_rels/notesSlide1.xml.rels'

    def dropped(name):
        if re.match(r'ppt/slides/(_rels/)?slide\d+\.xml(\.rels)?$', name) and name not in rename:
            return True
        if re.match(r'ppt/notesSlides/(_rels/)?notesSlide\d+\.xml(\.rels)?$', name) and name not in rename:
            return True
        return False

    # presentation.xml.rels から、残すスライドの rId を拾う
    prels = z.read('ppt/_rels/presentation.xml.rels').decode('utf-8')
    m = re.search(r'<Relationship Id="(rId\d+)"[^>]*Target="slides/slide%d\.xml"[^>]*/>' % n, prels)
    keep_rid = m.group(1)

    out = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
    for name in names:
        if dropped(name):
            continue
        data = z.read(name)
        target = rename.get(name, name)

        if name == 'ppt/presentation.xml':
            x = data.decode('utf-8')
            lst = re.search(r'<p:sldIdLst>(.*?)</p:sldIdLst>', x, re.S).group(1)
            keep = re.search(r'<p:sldId[^>]*r:id="%s"\s*/>' % keep_rid, lst).group(0)
            x = x.replace(re.search(r'<p:sldIdLst>.*?</p:sldIdLst>', x, re.S).group(0),
                          '<p:sldIdLst>%s</p:sldIdLst>' % keep)
            data = x.encode('utf-8')

        elif name == 'ppt/_rels/presentation.xml.rels':
            x = data.decode('utf-8')
            for mm in re.finditer(r'<Relationship[^>]*Target="slides/slide(\d+)\.xml"[^>]*/>', x):
                if int(mm.group(1)) != n:
                    x = x.replace(mm.group(0), '')
            x = x.replace('Target="slides/slide%d.xml"' % n, 'Target="slides/slide1.xml"')
            data = x.encode('utf-8')

        elif name == '[Content_Types].xml':
            x = data.decode('utf-8')
            for i in range(1, n_slides + 5):
                if i != n:
                    x = x.replace('<Override PartName="/ppt/slides/slide%d.xml" '
                                  'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' % i, '')
                if keep_notes is None or i != keep_notes:
                    x = x.replace('<Override PartName="/ppt/notesSlides/notesSlide%d.xml" '
                                  'ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>' % i, '')
            x = x.replace('PartName="/ppt/slides/slide%d.xml"' % n, 'PartName="/ppt/slides/slide1.xml"')
            if keep_notes:
                x = x.replace('PartName="/ppt/notesSlides/notesSlide%d.xml"' % keep_notes,
                              'PartName="/ppt/notesSlides/notesSlide1.xml"')
            data = x.encode('utf-8')

        elif target == 'ppt/slides/_rels/slide1.xml.rels' and keep_notes:
            data = data.decode('utf-8').replace(
                '../notesSlides/notesSlide%d.xml' % keep_notes, '../notesSlides/notesSlide1.xml').encode('utf-8')

        elif target == 'ppt/notesSlides/_rels/notesSlide1.xml.rels':
            data = data.decode('utf-8').replace(
                '../slides/slide%d.xml' % n, '../slides/slide1.xml').encode('utf-8')

        elif target == 'ppt/slides/slide1.xml' and new_title:
            x = data.decode('utf-8')
            # 見出しシェイプの最初の <a:t> だけ差し替える
            sp = re.search(r'<p:sp>(?:(?!</p:sp>).)*?<p:cNvPr[^>]*\sid="%s"(?:(?!</p:sp>).)*?</p:sp>' % title_id, x, re.S)
            if not sp:
                raise SystemExit('見出しシェイプ id=%s が見つかりません' % title_id)
            block = sp.group(0)
            # 見出しは複数ラン（「歓迎 」＋「本日のビジター」）に分かれている。
            # GAS側の setTextInShape_ と同じく、先頭ランに入れて残りのランは消す。
            runs = re.findall(r'<a:r>.*?</a:r>', block, re.S)
            if not runs:
                raise SystemExit('見出しシェイプにランがありません')
            nb = block
            for r in runs[1:]:
                nb = nb.replace(r, '')
            nb = nb.replace(runs[0], re.sub(r'<a:t>[^<]*</a:t>', '<a:t>%s</a:t>' % new_title, runs[0], count=1))
            x = x.replace(block, nb)
            data = x.encode('utf-8')

        out.writestr(target, data)
    out.close()
    print('  %s ← スライド%d%s' % (dst, n, ('（見出しを「%s」に）' % new_title) if new_title else ''))


if __name__ == '__main__':
    import os
    # 運用中の3枚構成のpptxを real.pptx として、このスクリプトと同じ場所に置いて実行する
    SRC = sys.argv[1] if len(sys.argv) > 1 else 'real.pptx'
    if not os.path.exists(SRC):
        raise SystemExit('元となるpptx（%s）が見つかりません。運用中のファイルを置いてください。' % SRC)
    # 紹介・ゲスト・代理は、いずれもスライド1から作る。
    # 実物のスライド2（代理）は、その週が2名だったため3枠目の見出し（氏名/専門分野/招待者）が
    # 消されている。3名いる週に使うと3枠目だけ見出しの無い表示になるので、
    # 枠が揃っているスライド1を使い、見出しの文字だけ差し替える。
    extract_slide(SRC, 1, 'BNI_テンプレート_ビジター紹介.pptx')
    extract_slide(SRC, 1, 'BNI_テンプレート_ゲスト紹介.pptx', new_title='歓迎 本日のゲスト')
    extract_slide(SRC, 1, 'BNI_テンプレート_代理紹介.pptx', new_title='代理出席の方々')
    extract_slide(SRC, 3, 'BNI_テンプレート_ビジタープレゼン.pptx')
