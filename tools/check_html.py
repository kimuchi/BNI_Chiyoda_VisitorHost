#!/usr/bin/env python3
"""ダイアログHTMLの、目で見つけにくい壊れ方を検出する。

1. <div> と </div> の数が合っているか
   閉じ忘れると、後ろの要素が別の要素の中に入ってしまう。
   親が display:none だと「ボタンを押しても何も起きない」ことになり、
   画面を見ても原因が分からない。実際にこれで表紙設定が開かなくなった。

2. <script> の中身が構文として通るか
   GASのスクリプトレット <?= ?> は文字列に置き換えてから見る。

    python3 tools/check_html.py
"""
import glob
import io
import os
import re
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def check_divs(path, text):
    body = text
    m = re.search(r'<body[^>]*>(.*)</body>', text, re.S)
    if m:
        body = m.group(1)
    body = re.sub(r'<script[^>]*>.*?</script>', '', body, flags=re.S)
    opens = len(re.findall(r'<div\b', body))
    closes = len(re.findall(r'</div>', body))
    if opens != closes:
        return '<div>が%d個、</div>が%d個で合っていません（%+d）' % (opens, closes, opens - closes)
    return None


def check_scripts(path, text):
    js = '\n'.join(m.group(1) for m in re.finditer(r'<script[^>]*>(.*?)</script>', text, re.S))
    js = re.sub(r'<\?!?=.*?\?>', '""', js, flags=re.S)   # GASのスクリプトレット
    js = re.sub(r'<\?.*?\?>', '', js, flags=re.S)
    if not js.strip():
        return None
    tmp = tempfile.NamedTemporaryFile('w', suffix='.js', delete=False, encoding='utf-8')
    tmp.write(js)
    tmp.close()
    try:
        r = subprocess.run(['node', '--check', tmp.name], capture_output=True, text=True)
    finally:
        os.unlink(tmp.name)
    if r.returncode != 0:
        first = [l for l in r.stderr.split('\n') if 'Error' in l or 'error' in l]
        return 'スクリプトの構文エラー: ' + (first[0].strip() if first else r.stderr.strip()[:200])
    return None


def main():
    bad = []
    files = sorted(glob.glob(os.path.join(ROOT, '*.html')))
    for f in files:
        text = io.open(f, encoding='utf-8').read()
        for problem in (check_divs(f, text), check_scripts(f, text)):
            if problem:
                bad.append((os.path.basename(f), problem))
    if bad:
        print('NG: HTMLに問題があります')
        for name, problem in bad:
            print('  %s  %s' % (name, problem))
        sys.exit(1)
    print('OK: HTML %d ファイル（divの対応・スクリプトの構文）' % len(files))


if __name__ == '__main__':
    main()
