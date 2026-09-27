#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MANUAL.md から、Apps Script のダイアログで表示する manual.html を生成する。

マニュアルの本文は MANUAL.md だけが正本。内容を変えたらこのスクリプトを実行して
manual.html を作り直し、両方をコミットする。

    python3 tools/build_manual.py

画像は MANUAL.md に 1行で ![説明](docs/images/名前.webp) と書く。manual.html には画像を埋め込む
（Apps Script の画面は、別に置いた画像ファイルを読めないため）。画像は tools/make_manual_shots.js で撮る。
"""
import base64
import html
import io
import os
import re
import sys

SRC = "MANUAL.md"
DST = "manual.html"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MIME = {".webp": "image/webp", ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif"}


def image_size(raw):
    """WebP・PNG の幅と高さ（分からなければ None）。"""
    if raw[:4] == b"RIFF" and raw[8:12] == b"WEBP":
        kind = raw[12:16]
        if kind == b"VP8 ":
            return (int.from_bytes(raw[26:28], "little") & 0x3FFF, int.from_bytes(raw[28:30], "little") & 0x3FFF)
        if kind == b"VP8L":
            b = int.from_bytes(raw[21:25], "little")
            return ((b & 0x3FFF) + 1, ((b >> 14) & 0x3FFF) + 1)
        if kind == b"VP8X":
            return (int.from_bytes(raw[24:27], "little") + 1, int.from_bytes(raw[27:30], "little") + 1)
    if raw[:8] == b"\x89PNG\r\n\x1a\n":
        return (int.from_bytes(raw[16:20], "big"), int.from_bytes(raw[20:24], "big"))
    return None


def figure(alt, rel):
    """画像を埋め込んだ <figure>。画像が無ければ止める（マニュアルに穴が空かないように）。
    幅と高さを書いておく（目次から飛んだあとに画像が読み込まれて、見出しが押し出されないように）。"""
    path = os.path.join(ROOT, rel)
    if not os.path.isfile(path):
        sys.exit("NG: MANUAL.md の画像が見つかりません: " + rel)
    raw = io.open(path, "rb").read()
    size = image_size(raw)
    wh = ' width="%d" height="%d"' % size if size else ""
    mime = MIME.get(os.path.splitext(path)[1].lower(), "application/octet-stream")
    return ('<figure class="shot"><img src="data:%s;base64,%s" alt="%s"%s>'
            '<figcaption>%s</figcaption></figure>'
            % (mime, base64.b64encode(raw).decode("ascii"), html.escape(re.sub(r"[`*]", "", alt)), wh, inline(alt)))


def inline(text):
    """行内記法（太字・コード）をHTMLに変換する。"""
    out = html.escape(text)
    out = re.sub(r"`([^`]+)`", r"<code>\1</code>", out)
    out = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", out)
    return out


def slug(n):
    return "sec%d" % n


def convert(md):
    lines = md.split("\n")
    body, toc = [], []
    i, sec = 0, 0
    in_code = False
    list_type = None   # 'ul' / 'ol' / None
    in_table = False
    para = []

    def flush_para():
        if para:
            body.append("<p>%s</p>" % inline(" ".join(para)))
            del para[:]

    def close_list():
        nonlocal list_type
        if list_type:
            body.append("</%s>" % list_type)
            list_type = None

    def close_table():
        nonlocal in_table
        if in_table:
            body.append("</tbody></table>")
            in_table = False

    def close_all():
        flush_para(); close_list(); close_table()

    while i < len(lines):
        line = lines[i].rstrip()

        # コードブロック
        if line.startswith("```"):
            if in_code:
                body.append("</pre>"); in_code = False
            else:
                close_all(); body.append("<pre>"); in_code = True
            i += 1; continue
        if in_code:
            body.append(html.escape(lines[i]))
            i += 1; continue

        # 空行
        if not line.strip():
            close_all(); i += 1; continue

        # 画像（1行だけの ![説明](パス)）
        m = re.match(r"^!\[([^\]]*)\]\(([^)\s]+)\)\s*$", line.strip())
        if m:
            close_all()
            body.append(figure(m.group(1), m.group(2)))
            i += 1; continue

        # 水平線
        if re.match(r"^---+$", line):
            close_all(); body.append("<hr>"); i += 1; continue

        # 見出し
        m = re.match(r"^(#{1,4})\s+(.*)$", line)
        if m:
            close_all()
            level, text = len(m.group(1)), m.group(2)
            if level == 2:
                sec += 1
                toc.append((sec, text))
                body.append('<h2 id="%s">%s</h2>' % (slug(sec), inline(text)))
            else:
                body.append("<h%d>%s</h%d>" % (level, inline(text), level))
            i += 1; continue

        # 表
        if line.startswith("|"):
            cells = [c.strip() for c in line.strip("|").split("|")]
            if re.match(r"^\|[\s:|-]+\|?$", line):   # 区切り行
                i += 1; continue
            if not in_table:
                close_all()
                body.append("<table><thead><tr>" +
                            "".join("<th>%s</th>" % inline(c) for c in cells) +
                            "</tr></thead><tbody>")
                in_table = True
            else:
                body.append("<tr>" + "".join("<td>%s</td>" % inline(c) for c in cells) + "</tr>")
            i += 1; continue
        close_table()

        # 引用
        if line.startswith(">"):
            close_all()
            quote = []
            while i < len(lines) and lines[i].startswith(">"):
                quote.append(lines[i].lstrip(">").strip())
                i += 1
            body.append("<blockquote>%s</blockquote>" % inline(" ".join(q for q in quote if q)))
            continue

        # 箇条書き / 番号リスト
        m = re.match(r"^\s*[-*]\s+(.*)$", line)
        if m:
            flush_para()
            if list_type != "ul":
                close_list(); body.append("<ul>"); list_type = "ul"
            body.append("<li>%s</li>" % inline(m.group(1)))
            i += 1; continue
        m = re.match(r"^\s*\d+\.\s+(.*)$", line)
        if m:
            flush_para()
            if list_type != "ol":
                close_list(); body.append("<ol>"); list_type = "ol"
            body.append("<li>%s</li>" % inline(m.group(1)))
            i += 1; continue
        # 箇条書きの続きの行（字下げされた行）は、直前の項目につなげる。
        # 別の段落にすると番号付きリストがそこで切れ、次の項目がまた「1.」から始まってしまう。
        if list_type and re.match(r"^\s{2,}\S", line) and body and body[-1].endswith("</li>"):
            body[-1] = body[-1][:-len("</li>")] + " " + inline(line.strip()) + "</li>"
            i += 1; continue
        close_list()

        para.append(line.strip())
        i += 1

    close_all()
    if in_code:
        body.append("</pre>")
    return "\n".join(body), toc


TEMPLATE = """<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {{ font-family: sans-serif; margin: 0; padding: 0; color: #222; line-height: 1.75; font-size: 14px; }}
    .wrap {{ padding: 18px 22px 40px; }}
    h1 {{ font-size: 20px; color: #0055ff; margin: 0 0 6px; }}
    h2 {{ font-size: 17px; color: #0055ff; margin: 28px 0 10px; padding-bottom: 6px; border-bottom: 2px solid #e3ecff; scroll-margin-top: 10px; }}
    h3 {{ font-size: 15px; margin: 20px 0 6px; color: #333; }}
    h4 {{ font-size: 14px; margin: 14px 0 4px; color: #555; }}
    p {{ margin: 8px 0; }}
    ul, ol {{ margin: 8px 0 8px 22px; padding: 0; }}
    li {{ margin: 4px 0; }}
    code {{ background: #f1f3f6; padding: 2px 6px; border-radius: 3px; font-size: 12.5px; color: #0f3d91; }}
    pre {{ background: #f7f8fa; border: 1px solid #e2e6eb; border-radius: 5px; padding: 12px; overflow-x: auto;
           font-size: 12.5px; line-height: 1.6; white-space: pre; }}
    blockquote {{ margin: 10px 0; padding: 10px 14px; background: #fffbe8; border-left: 4px solid #ffc107; border-radius: 0 4px 4px 0; }}
    table {{ border-collapse: collapse; width: 100%; margin: 10px 0; font-size: 13px; }}
    th, td {{ border: 1px solid #dde2e8; padding: 7px 9px; text-align: left; vertical-align: top; }}
    th {{ background: #f2f6ff; }}
    hr {{ border: none; border-top: 1px solid #e5e5e5; margin: 22px 0; }}
    figure.shot {{ margin: 12px 0 18px; }}
    figure.shot img {{ display: block; max-width: 100%; height: auto; border: 1px solid #dde3ea; border-radius: 6px;
                       box-shadow: 0 1px 4px rgba(0,0,0,.08); }}
    figure.shot figcaption {{ font-size: 12px; color: #667; margin-top: 5px; }}
    .toc {{ background: #f7f9fc; border: 1px solid #dde3ea; border-radius: 6px; padding: 12px 16px; margin: 14px 0 6px; }}
    .toc b {{ display: block; margin-bottom: 6px; color: #0055ff; }}
    .toc ol {{ margin: 0 0 0 20px; }}
    .toc a {{ color: #0f3d91; text-decoration: none; }}
    .toc a:hover {{ text-decoration: underline; }}
    .top {{ position: fixed; right: 18px; bottom: 16px; background: #0055ff; color: #fff;
            padding: 8px 14px; border-radius: 20px; font-size: 12px; text-decoration: none;
            box-shadow: 0 2px 6px rgba(0,0,0,.25); }}
  </style>
</head>
<body>
  <div class="wrap" id="top">
    <div class="toc">
      <b>目次</b>
      <ol>
{toc}
      </ol>
    </div>
{body}
  </div>
  <a class="top" href="#top" target="_self">▲ 目次へ</a>
  <script>
    // 目次などのページ内のリンクは、この中で見出しまで動かすだけにする。
    // Apps Script の画面は枠（iframe）の中で表示され、<base target="_top"> のままだと
    // 外側の画面ごと移動して「ページが開かない」になるため
    document.addEventListener('click', function (e) {{
      var a = e.target && e.target.closest ? e.target.closest('a[href^="#"]') : null;
      if (!a) return;
      var id = a.getAttribute('href').slice(1), t = id ? document.getElementById(id) : null;
      e.preventDefault();
      if (t) t.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
    }});
  </script>
</body>
</html>
"""


def main():
    src, dst = os.path.join(ROOT, SRC), os.path.join(ROOT, DST)
    md = io.open(src, encoding="utf-8").read()
    body, toc = convert(md)
    toc_html = "\n".join(
        '        <li><a href="#%s" target="_self">%s</a></li>' % (slug(n), html.escape(t)) for n, t in toc
    )
    io.open(dst, "w", encoding="utf-8").write(
        TEMPLATE.format(toc=toc_html, body=body)
    )
    print("generated %s (%d sections, %d bytes)" % (DST, len(toc), os.path.getsize(dst)))


if __name__ == "__main__":
    sys.exit(main())
