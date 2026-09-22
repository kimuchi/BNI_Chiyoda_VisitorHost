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
    print('OK: マニュアルの導線はすべて実際のメニュー項目と一致しています（メニュー %d 項目）' % len(labels))


if __name__ == '__main__':
    main()
