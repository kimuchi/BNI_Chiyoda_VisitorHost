# -*- coding: utf-8 -*-
"""pptxに残る「ファイルの情報」から、人の名前やアカウントを消す（docs/templates に置く前に使う）。

    python3 tools/clean_pptx_meta.py <入力.pptx> [<出力.pptx>]   … 消す（出力を省くと上書き）
    python3 tools/clean_pptx_meta.py --check <pptx>…             … 残っていないか確かめる（残っていれば終了コード1）

PowerPointは、スライドの文字とは別に、次のところへ名前やアカウントを書き込む。
スライドの文字を見本に置き換えても、ここに残っていると公開されてしまう。

    docProps/core.xml          作成者（dc:creator）・最終更新者（cp:lastModifiedBy）
    ppt/commentAuthors.xml     コメントを書いた方の名前・イニシャル・アカウントのID
    ppt/changesInfos/*.xml     共同編集の変更の記録（編集した方の名前・アカウントのID・日時）
    ppt/revisionInfo.xml       同じく版の記録
    docProps/thumbnail.jpeg    保存したときの1枚目の縮小画像（差し替える前の名前が写っていることがある）

消すだけで、スライド・画像・音楽には触らない。実在の方の氏名・写真・会社名は、別に見本へ置き換えること
（docs/templates/README.md の「見本にするとき」）。
"""
import io, re, sys, zipfile

DROP_RE = re.compile(r'^(ppt/changesInfos/.+|ppt/revisionInfo\.xml)$')


def clean(src, dst=None):
    zin = zipfile.ZipFile(src)
    infos = [i for i in zin.infolist() if not DROP_RE.match(i.filename)]
    dropped = [i.filename for i in zin.infolist() if DROP_RE.match(i.filename)]
    parts = {i.filename: zin.read(i.filename) for i in infos}
    if 'docProps/core.xml' in parts:
        x = parts['docProps/core.xml'].decode('utf-8')
        x = re.sub(r'<dc:creator>[^<]*</dc:creator>', '<dc:creator></dc:creator>', x)
        x = re.sub(r'<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', '<cp:lastModifiedBy></cp:lastModifiedBy>', x)
        parts['docProps/core.xml'] = x.encode('utf-8')
    if 'ppt/commentAuthors.xml' in parts:
        x = parts['ppt/commentAuthors.xml'].decode('utf-8')
        x = re.sub(r'\bname="[^"]*"', 'name="見本"', x)
        x = re.sub(r'\binitials="[^"]*"', 'initials="M"', x)
        x = re.sub(r'\buserId="[^"]*"', 'userId="見本"', x)
        x = re.sub(r'\bproviderId="[^"]*"', 'providerId="None"', x)
        parts['ppt/commentAuthors.xml'] = x.encode('utf-8')
    # 取り除いた部品への参照（関係と型の登録）も消す
    for name in dropped:
        rel = 'ppt/_rels/presentation.xml.rels'
        if rel in parts:
            target = name[len('ppt/'):]
            x = parts[rel].decode('utf-8')
            parts[rel] = re.sub(r'<Relationship\b[^>]*Target="' + re.escape(target) + r'"[^>]*/>', '', x).encode('utf-8')
        x = parts['[Content_Types].xml'].decode('utf-8')
        parts['[Content_Types].xml'] = re.sub(r'<Override\b[^>]*PartName="/' + re.escape(name) + r'"[^>]*/>', '', x).encode('utf-8')
    if 'docProps/thumbnail.jpeg' in parts:
        parts['docProps/thumbnail.jpeg'] = blank_thumbnail(parts['docProps/thumbnail.jpeg'])
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i.filename, parts[i.filename])
    open(dst or src, 'wb').write(out.getvalue())
    return dropped


def blank_thumbnail(data):
    """縮小画像を、同じ大きさの白い画像にする（Pillow が無ければそのまま。--check で分かる）"""
    try:
        from PIL import Image
        w, h = Image.open(io.BytesIO(data)).size
        b = io.BytesIO()
        Image.new('RGB', (w, h), (255, 255, 255)).save(b, 'JPEG', quality=80)
        return b.getvalue()
    except Exception:
        return data


def problems(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    out = []
    if 'docProps/core.xml' in names:
        x = z.read('docProps/core.xml').decode('utf-8')
        for tag in ('dc:creator', 'cp:lastModifiedBy'):
            m = re.search(r'<%s>([^<]+)</%s>' % (tag, tag), x)
            if m:
                out.append('%s に「%s」' % (tag, m.group(1)))
    if 'ppt/commentAuthors.xml' in names:
        x = z.read('ppt/commentAuthors.xml').decode('utf-8')
        for attr in ('name', 'userId'):
            left = [v for v in re.findall(r'\b%s="([^"]*)"' % attr, x) if v not in ('見本', '')]
            if left:
                out.append('コメントの作成者の %s に %s' % (attr, left))
    left = [n for n in names if DROP_RE.match(n)]
    if left:
        out.append('変更の記録が残っている: %s' % left)
    return out


def main(argv):
    if len(argv) >= 2 and argv[0] == '--check':
        bad = 0
        for p in argv[1:]:
            pr = problems(p)
            if pr:
                bad += 1
                print('NG: %s … %s' % (p, ' ／ '.join(pr)))
        if bad:
            print('→ python3 tools/clean_pptx_meta.py <ファイル> で消してください')
            return 1
        print('OK: pptxのファイルの情報に、名前・アカウント・変更の記録は残っていません（%d ファイル）' % (len(argv) - 1))
        return 0
    if not argv:
        print(__doc__)
        return 2
    dropped = clean(argv[0], argv[1] if len(argv) > 1 else None)
    print('消しました: %s%s' % (argv[1] if len(argv) > 1 else argv[0], '（取り除いた部品: %s）' % dropped if dropped else ''))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
