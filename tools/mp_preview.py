#!/usr/bin/env python3
"""スライドXMLの図形を、おおよその見た目でHTMLに描く（配置の確認用）。

PowerPointが無い環境では出来上がりを開けないので、位置・大きさ・文字だけでも
目で確かめられるようにする。装飾やグループ図形の中身までは再現しない。

    python3 tools/mp_preview.py <pptx> <スライド番号…> > preview.html
"""
import base64
import os
import re
import sys
import zipfile
from xml.dom import minidom

NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
EMU = 12700.0     # 1pt


def esc(s):
    return (s or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def render(z, n):
    doc = minidom.parseString(z.read('ppt/slides/slide%d.xml' % n))
    rels = z.read('ppt/slides/_rels/slide%d.xml.rels' % n).decode('utf-8')
    rid2tgt = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    pres = z.read('ppt/presentation.xml').decode('utf-8')
    m = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', pres)
    W, H = int(m.group(1)) / EMU, int(m.group(2)) / EMU
    out = ['<div class="slide" style="width:%.0fpt;height:%.0fpt;">' % (W, H),
           '<div class="no">slide%d</div>' % n]

    tree = doc.getElementsByTagNameNS(NS_P, 'spTree')[0]
    for ch in tree.childNodes:
        if ch.nodeType != 1 or ch.localName not in ('sp', 'pic', 'graphicFrame', 'grpSp'):
            continue
        # 表(graphicFrame)の位置は <p:xfrm>、図形は <a:xfrm> に入っている
        xf = ch.getElementsByTagNameNS(NS_A, 'xfrm') or ch.getElementsByTagNameNS(NS_P, 'xfrm')
        if not xf:
            continue
        o = xf[0].getElementsByTagNameNS(NS_A, 'off')[0]
        e = xf[0].getElementsByTagNameNS(NS_A, 'ext')[0]
        x, y = int(o.getAttribute('x')) / EMU, int(o.getAttribute('y')) / EMU
        cx, cy = int(e.getAttribute('cx')) / EMU, int(e.getAttribute('cy')) / EMU
        cn = ch.getElementsByTagNameNS(NS_P, 'cNvPr')
        sid = cn[0].getAttribute('id') if cn else '?'
        box = 'left:%.1fpt;top:%.1fpt;width:%.1fpt;height:%.1fpt;' % (x, y, cx, cy)

        if ch.localName == 'pic':
            blip = ch.getElementsByTagNameNS(NS_A, 'blip')
            src = ''
            if blip:
                rid = blip[0].getAttributeNS(NS_R, 'embed')
                tgt = os.path.normpath(os.path.join('ppt/slides', rid2tgt.get(rid, ''))).replace('\\', '/')
                if tgt in z.namelist():
                    ext = tgt.rsplit('.', 1)[-1]
                    src = 'data:image/%s;base64,%s' % (ext, base64.b64encode(z.read(tgt)).decode())
            sr = ch.getElementsByTagNameNS(NS_A, 'srcRect')
            fit = 'cover' if sr else 'fill'
            out.append('<img class="pic" style="%sobject-fit:%s;" src="%s">' % (box, fit, src))
            continue

        if ch.localName == 'graphicFrame':
            tbl = ch.getElementsByTagNameNS(NS_A, 'tbl')
            if not tbl:
                continue
            rows = ''
            for tr in tbl[0].getElementsByTagNameNS(NS_A, 'tr'):
                cells = ''
                for tc in tr.getElementsByTagNameNS(NS_A, 'tc'):
                    t = ''.join(x.firstChild.nodeValue if x.firstChild else ''
                                for x in tc.getElementsByTagNameNS(NS_A, 't'))
                    cells += '<td>%s</td>' % esc(t)
                rows += '<tr>%s</tr>' % cells
            out.append('<div class="tbl" style="%s"><table>%s</table></div>' % (box, rows))
            continue

        if ch.localName == 'grpSp':
            out.append('<div class="grp" style="%s"><span>グループ %s</span></div>' % (box, sid))
            continue

        tb = ch.getElementsByTagNameNS(NS_P, 'txBody')
        if not tb:
            continue
        bp = tb[0].getElementsByTagNameNS(NS_A, 'bodyPr')
        anchor = bp[0].getAttribute('anchor') if bp else ''
        just = {'t': 'flex-start', 'ctr': 'center', 'b': 'flex-end'}.get(anchor, 'center')
        lines = ''
        for p in tb[0].getElementsByTagNameNS(NS_A, 'p'):
            runs = p.getElementsByTagNameNS(NS_A, 'r')
            if not runs:
                continue
            rpr = runs[0].getElementsByTagNameNS(NS_A, 'rPr')
            sz = int(rpr[0].getAttribute('sz') or 1800) / 100 if rpr else 18
            bold = rpr and rpr[0].getAttribute('b') == '1'
            ppr = p.getElementsByTagNameNS(NS_A, 'pPr')
            algn = ppr[0].getAttribute('algn') if ppr else ''
            spc = p.getElementsByTagNameNS(NS_A, 'spcPct')
            lh = (int(spc[0].getAttribute('val')) / 1000.0) if spc else 100
            txt = ''.join(x.firstChild.nodeValue if x.firstChild else ''
                          for x in p.getElementsByTagNameNS(NS_A, 't'))
            lines += ('<div style="font-size:%.1fpt;%stext-align:%s;line-height:%.0f%%;">%s</div>'
                      % (sz, 'font-weight:bold;' if bold else '',
                         {'ctr': 'center', 'r': 'right'}.get(algn, 'left'), lh, esc(txt) or '&nbsp;'))
        out.append('<div class="tx" style="%sjustify-content:%s;" data-id="%s">%s</div>' % (box, just, sid, lines))
    out.append('</div>')
    return '\n'.join(out)


def main():
    z = zipfile.ZipFile(sys.argv[1])
    body = '\n'.join(render(z, int(n)) for n in sys.argv[2:])
    print('''<!DOCTYPE html><html><head><meta charset="utf-8"><style>
 body{background:#555;margin:0;padding:12px;font-family:"Meiryo UI","Noto Sans CJK JP",sans-serif;}
 .slide{position:relative;background:#fff;margin:0 auto 14px;overflow:hidden;box-shadow:0 2px 10px #0006;}
 .slide>div,.slide>img{position:absolute;}
 .no{left:2pt;top:2pt;font-size:8pt;color:#c00;z-index:99;}
 .tx{display:flex;flex-direction:column;outline:.5pt dashed #09f;}
 .tx>div{width:100%;}
 .pic{outline:.5pt dashed #0a0;}
 .grp{outline:.5pt dashed #aaa;color:#bbb;font-size:9pt;}
 .tbl table{border-collapse:collapse;width:100%;font-size:11pt;}
 .tbl td{border:.5pt solid #999;padding:1pt 4pt;}
</style></head><body>''' + body + '</body></html>')


if __name__ == '__main__':
    main()
