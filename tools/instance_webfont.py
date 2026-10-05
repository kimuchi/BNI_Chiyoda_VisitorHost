#!/usr/bin/env python3
"""Googleフォントを写したディレクトリの可変フォントを、太さごとの普通の書体に切り出す。

    python3 tools/instance_webfont.py <写したディレクトリ> <出力のディレクトリ>

写したディレクトリは tools/check_memberbook.js・tools/build_manual_pdf.js と同じ置き方（css.txt と gstatic/。
gstatic/ には https://fonts.gstatic.com/… を、パスの「/」を「_」にした名前で置く）。出力も同じ置き方で作る。

Noto Sans JP などは太さを変えられる書体（可変フォント）で、Chromium でPDFにすると文字が絵（Type3）として入り、
PDFが大きくなる（検索・書き写しもしにくい）。css.txt の太さ（font-weight）ごとに切り出した書体を使うと、ふつうの書体として入る。

fontTools と brotli（woff2 の読み書き）が要る:
    pip install fonttools brotli
    （入れられないときは pip install --target <ディレクトリ> fonttools brotli として PYTHONPATH=<ディレクトリ> で動かす）
"""
import os
import re
import sys
from concurrent.futures import ProcessPoolExecutor

try:
    from fontTools.ttLib import TTFont
    from fontTools.varLib import instancer
except ImportError:
    sys.exit('fontTools がありません（pip install fonttools brotli）')

G = 'https://fonts.gstatic.com/'


def file_name(url):
    return url.replace(G, '').replace('/', '_')


def make(job):
    src, dst, url, wght = job
    out_url = G + 'inst%d/' % wght + url.replace(G, '')
    out = os.path.join(dst, 'gstatic', file_name(out_url))
    if not os.path.exists(out):
        f = TTFont(os.path.join(src, 'gstatic', file_name(url)))
        if 'fvar' in f:
            f = instancer.instantiateVariableFont(f, {'wght': wght}, updateFontNames=True)
        for rec in f['name'].names:                    # 名前に残る「Thin」（可変フォントのもとの名前）を外す
            if rec.nameID in (1, 3, 4, 6, 16, 17):
                rec.string = rec.toUnicode().replace(' Thin', '').replace('Thin', '')
        f.flavor = 'woff2'
        f.save(out)
    return out_url


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    src, dst = sys.argv[1], sys.argv[2]
    css = open(os.path.join(src, 'css.txt'), encoding='utf-8').read()
    os.makedirs(os.path.join(dst, 'gstatic'), exist_ok=True)
    blocks = re.findall(r'@font-face\s*\{[^}]*\}', css)
    jobs = []
    for b in blocks:
        w = re.search(r'font-weight:\s*(\d+)', b)
        u = re.search(r'url\((https://fonts\.gstatic\.com/[^)]+)\)', b)
        if not (w and u):
            sys.exit('css.txt の読めない @font-face があります:\n' + b)
        jobs.append((src, dst, u.group(1), int(w.group(1))))
    with ProcessPoolExecutor(os.cpu_count() or 4) as ex:
        outs = list(ex.map(make, jobs))
    new = css
    for (_, _, u, _), o, b in zip(jobs, outs, blocks):
        new = new.replace(b, b.replace(u, o), 1)
    open(os.path.join(dst, 'css.txt'), 'w', encoding='utf-8').write(new)
    print('切り出した書体: %d ファイル（もとの可変フォント %d ファイル）→ %s' % (len(set(outs)), len(set(j[2] for j in jobs)), dst))


if __name__ == '__main__':
    main()
