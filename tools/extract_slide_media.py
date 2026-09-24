#!/usr/bin/env python3
"""スライドに埋め込まれている音楽・動画を、曲名のファイル名で取り出す。

差し替え用に手元へ置いておくため、また「いまどの曲が何ページで鳴っているか」を
一覧で確かめるために使う。

    python3 tools/extract_slide_media.py <pptx> <取り出し先ディレクトリ>
"""
import os
import re
import sys
import zipfile

MEDIA_RE = re.compile(r'\.(mp3|m4a|wav|wma|aac|mp4|mov|m4v|wmv)$', re.I)


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    src, out = sys.argv[1], sys.argv[2]
    os.makedirs(out, exist_ok=True)
    z = zipfile.ZipFile(src)

    # スライドごとに、音声・動画の図形名（＝曲名）を拾う
    label = {}
    for n in sorted(x for x in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', x)):
        no = int(re.search(r'(\d+)', n.split('/')[-1]).group(1))
        rp = 'ppt/slides/_rels/%s.rels' % n.split('/')[-1]
        if rp not in z.namelist():
            continue
        rels = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', z.read(rp).decode('utf-8')))
        s = z.read(n).decode('utf-8')
        for m in re.finditer(r'<p:pic>.*?</p:pic>', s, re.S):
            seg = m.group(0)
            f = re.search(r'<a:(?:audio|video)File[^>]*r:link="([^"]+)"', seg)
            if not f:
                continue
            tgt = rels.get(f.group(1), '')
            if not tgt.startswith('../media/'):
                continue
            part = 'ppt/media/' + tgt.split('/')[-1]
            nm = re.search(r'<p:cNvPr id="\d+" name="([^"]*)"', seg)
            vol = re.search(r'spid="%s"' % (re.search(r'<p:cNvPr id="(\d+)"', seg).group(1)), s)
            label.setdefault(part, []).append((no, (nm.group(1) if nm else '')))

    n = 0
    for part in sorted(x for x in z.namelist() if x.startswith('ppt/media/') and MEDIA_RE.search(x)):
        ext = MEDIA_RE.search(part).group(0).lower()
        uses = label.get(part, [])
        name = (uses[0][1] if uses else part.split('/')[-1].rsplit('.', 1)[0])
        pages = '_'.join('P%d' % u[0] for u in uses) or '未使用'
        safe = re.sub(r'[\\/:*?"<>|]', '_', name).strip() or 'media'
        dst = os.path.join(out, '%s_%s%s' % (pages, safe, ext))
        with open(dst, 'wb') as f:
            f.write(z.read(part))
        n += 1
        print('  %-52s %7.2f MB' % (os.path.basename(dst), os.path.getsize(dst) / 1048576))
    print('%d 件を取り出しました → %s' % (n, out))


if __name__ == '__main__':
    main()
