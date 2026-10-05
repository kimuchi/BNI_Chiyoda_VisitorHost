#!/usr/bin/env python3
"""MANUAL.md に書かれたメニュー導線が、実際のメニューに存在するかを確認する。

メニューを整理すると導線がずれたまま残りやすい（実際に6か所ずれていた）。
`名簿システム` > `⚙️ 設定` > `休会日` のような表記を拾い、
コード.js の onOpen() が作る項目名と突き合わせる。

    python3 tools/check_manual_links.py
"""
import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def menu_labels():
    code = io.open(os.path.join(ROOT, 'コード.js'), encoding='utf-8').read()
    m = re.search(r"function onOpen\(\)[\s\S]*?\.addToUi\(\);", code)
    if not m:
        print('NG: コード.js に onOpen() が見つかりません')
        sys.exit(1)
    body = m.group(0)
    return set(re.findall(r"\.addItem\('([^']+)'", body)) | \
           set(re.findall(r"createMenu\('([^']+)'", body))


def main():
    labels = menu_labels()
    man = io.open(os.path.join(ROOT, 'MANUAL.md'), encoding='utf-8').read()
    bad = []
    for no, line in enumerate(man.split('\n'), 1):
        for m in re.finditer(r'`名簿システム`((?: > `[^`]+`)+)', line):
            for part in re.findall(r'`([^`]+)`', m.group(1)):
                if part not in labels:
                    bad.append((no, part, line.strip()[:70]))
    if bad:
        print('NG: マニュアルの導線が実際のメニューにありません')
        for no, part, line in bad:
            print('  MANUAL.md:%d  「%s」  %s' % (no, part, line))
        sys.exit(1)
    # manual.html の目次などのページ内リンク：<base target="_top"> のままだと、Apps Script の枠の外側ごと
    # 移動して「ページが開かない」になる。target="_self" と、見出しまで動かすだけの仕組みがあるか
    page = io.open(os.path.join(ROOT, 'manual.html'), encoding='utf-8').read()
    inner = re.findall(r'<a\b[^>]*href="#[^"]*"[^>]*>', page)
    loose = [a for a in inner if 'target="_self"' not in a]
    if not inner or loose or 'e.preventDefault()' not in page or 'scrollIntoView' not in page:
        print('NG: manual.html のページ内リンクが、画面の外側ごと移動します（目次から飛ぶとページが開かない）')
        for a in loose[:5]:
            print('  ' + a)
        sys.exit(1)
    # リンクの書き方：manual.html に変換できるのは [文字](https://…) だけ。[文字](#…) は記号のまま出てしまう
    anchors = [(i + 1, l.strip()) for i, l in enumerate(man.split('\n')) if re.search(r'\]\(#', l)]
    if anchors or re.search(r'\]\((https?://[^)\s]+)\)', re.sub(r'<[^>]+>', '', page.split('</style>')[-1])):
        print('NG: MANUAL.md のリンクが manual.html で記号のまま出ます（ページ内リンク [文字](#…) は使えません）')
        for no, l in anchors[:5]:
            print('  MANUAL.md:%d  %s' % (no, l[:80]))
        sys.exit(1)
    # 記号の残り：表・太字・コードの欄の書き方が、manual.html でそのまま出ていないか
    # （引用の中の表、行をまたいだ太字、箇条書きの中で字下げしたコードの欄は、以前は記号のまま出ていた）
    text = re.sub(r'<pre>[\s\S]*?</pre>|<script[\s\S]*?</script>|<style[\s\S]*?</style>|<img [^>]*>', '', page)
    text = re.sub(r'<[^>]+>', '\n', text)
    marks = []
    for kind, pat in (('表の区切り', r'\|\s*-{3,}'), ('太字の **', r'\*\*'), ('コードの欄の ```', r'```')):
        marks += [(kind, m.strip()[:60]) for m in re.findall(r'[^\n]*' + pat + r'[^\n]*', text)]
    if marks:
        print('NG: manual.html に、MANUAL.md の書き方の記号がそのまま出ています（tools/build_manual.py が変換できない書き方）')
        for kind, m in marks[:8]:
            print('  %s: %s' % (kind, m))
        sys.exit(1)
    # 画像：MANUAL.md の画像がそろっていて、manual.html に同じ数だけ（幅・高さつきで）埋め込まれているか
    shots = re.findall(r'^!\[[^\]]*\]\(([^)\s]+)\)\s*$', man, re.M)
    missing = [s for s in shots if not os.path.isfile(os.path.join(ROOT, s))]
    figs = re.findall(r'<figure class="shot"><img src="data:image/[a-z]+;base64,[A-Za-z0-9+/=]+"[^>]*>', page)
    sized = [f for f in figs if re.search(r' width="\d+" height="\d+"', f)]
    if missing or len(figs) != len(shots) or len(sized) != len(figs):
        print('NG: マニュアルの画像: 無いファイル %s ／ MANUAL.md %d 枚・manual.html %d 枚（幅・高さつき %d 枚）。'
              'python3 tools/build_manual.py で作り直してください' % (missing, len(shots), len(figs), len(sized)))
        sys.exit(1)
    print('OK: マニュアルの導線はすべて実際のメニュー項目と一致しています（メニュー %d 項目）。'
          '目次などのページ内リンク %d 個は、マニュアルの中だけで動きます。画像 %d 枚' % (len(labels), len(inner), len(shots)))


if __name__ == '__main__':
    main()
