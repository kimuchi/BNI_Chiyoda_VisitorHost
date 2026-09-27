#!/usr/bin/env python3
"""pptx の部品のつながりを確かめる（PowerPoint が「修復」を求めるような壊れ方を、開く前に見つける）。

    python3 tools/pptx_integrity.py <ファイル.pptx>…

確かめること
  ・XML がすべて正しく読める
  ・関係（.rels）の行き先の部品がある。種類の登録（[Content_Types].xml）が全部の部品にある
  ・presentation.xml のページの並びが、関係のページを指している
  ・ページごとに：図形の番号が重ならない。アニメーションが指す図形（spTgt / bldP）がある。
    図形の r:embed / r:link / r:id の関係がある
"""
import posixpath
import re
import sys
import zipfile
from xml.dom import minidom


def check(path):
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    probs = []
    ct = z.read('[Content_Types].xml').decode('utf-8')
    defaults = {m.lower() for m in re.findall(r'<Default Extension="([^"]+)"', ct)}
    overrides = set(re.findall(r'<Override PartName="/([^"]+)"', ct))
    for o in overrides:
        if o not in names:
            probs.append('種類の登録にあって部品が無い: ' + o)
    for n in names:
        if n.endswith('/'):
            continue
        ext = n.rsplit('.', 1)[-1].lower() if '.' in n else ''
        if n not in overrides and ext not in defaults and n != '[Content_Types].xml':
            probs.append('種類の登録が無い部品: ' + n)
        if n.endswith('.xml') or n.endswith('.rels'):
            try:
                minidom.parseString(z.read(n))
            except Exception as e:  # noqa
                probs.append('XMLとして読めない: %s (%s)' % (n, e))
    rels_of = {}
    for n in names:
        if not n.endswith('.rels'):
            continue
        owner_dir = posixpath.dirname(posixpath.dirname(n))
        x = z.read(n).decode('utf-8')
        rels = {}
        for m in re.finditer(r'<Relationship\b[^>]*>', x):
            t = m.group(0)
            rid = re.search(r'\bId="([^"]+)"', t).group(1)
            tgt = re.search(r'Target="([^"]*)"', t).group(1)
            if rid in rels:
                probs.append('関係IDが重なっている: %s %s' % (n, rid))
            rels[rid] = tgt
            if 'TargetMode="External"' in t or tgt == 'NULL':
                continue
            abs_ = tgt[1:] if tgt.startswith('/') else posixpath.normpath(posixpath.join(owner_dir, tgt))
            if abs_ not in names:
                probs.append('関係の行き先が無い: %s %s → %s' % (n, rid, abs_))
        rels_of[n] = rels
    prs = z.read('ppt/presentation.xml').decode('utf-8')
    prels = rels_of.get('ppt/_rels/presentation.xml.rels', {})
    slides = []
    for rid in re.findall(r'<p:sldId [^>]*r:id="([^"]+)"', prs):
        t = prels.get(rid)
        if not t:
            probs.append('ページの並びの関係が無い: ' + rid)
        else:
            slides.append('ppt/' + t)
    ids = re.findall(r'<p:sldId id="(\d+)"', prs)
    if len(set(ids)) != len(ids):
        probs.append('ページの番号（sldId）が重なっている')
    for s in slides:
        x = z.read(s).decode('utf-8')
        sid = re.findall(r'<p:cNvPr\b[^>]*\sid="(\d+)"', x)
        dup = sorted({i for i in sid if sid.count(i) > 1})
        if dup:
            probs.append('%s 図形の番号が重なっている: %s' % (s, dup))
        have = set(sid)
        for ref in set(re.findall(r'\bspid="(\d+)"', x)):
            if ref not in have:
                probs.append('%s アニメーションが無い図形を指している: spid=%s' % (s, ref))
        rels = rels_of.get(posixpath.join(posixpath.dirname(s), '_rels', posixpath.basename(s) + '.rels'), {})
        for ref in set(re.findall(r'\br:(?:embed|link|id|pict)="(rId\d+)"', x)):
            if ref not in rels:
                probs.append('%s 関係の無いIDを使っている: %s' % (s, ref))
        ctn = re.findall(r'<p:cTn id="(\d+)"', x)
        if len(set(ctn)) != len(ctn):
            probs.append('%s アニメーションの番号（cTn）が重なっている' % s)
    return probs, len(slides)


def main(argv):
    bad = 0
    for p in argv:
        probs, n = check(p)
        if probs:
            bad += 1
            print('NG: %s（%d枚）' % (p, n))
            for q in probs[:30]:
                print('   ' + q)
        else:
            print('OK: %s（%d枚）' % (p, n))
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
