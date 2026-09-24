#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MANUAL.md から、Apps Script のダイアログで表示する manual.html を生成する。

マニュアルの本文は MANUAL.md だけが正本。内容を変えたらこのスクリプトを実行して
manual.html を作り直し、両方をコミットする。

    python3 tools/build_manual.py
"""
import html
import io
import os
import re
import sys

SRC = "MANUAL.md"
DST = "manual.html"


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
  <a class="top" href="#top">▲ 目次へ</a>
</body>
</html>
"""


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    src, dst = os.path.join(root, SRC), os.path.join(root, DST)
    md = io.open(src, encoding="utf-8").read()
    body, toc = convert(md)
    toc_html = "\n".join(
        '        <li><a href="#%s">%s</a></li>' % (slug(n), html.escape(t)) for n, t in toc
    )
    io.open(dst, "w", encoding="utf-8").write(
        TEMPLATE.format(toc=toc_html, body=body)
    )
    print("generated %s (%d sections, %d bytes)" % (DST, len(toc), os.path.getsize(dst)))


if __name__ == "__main__":
    sys.exit(main())
