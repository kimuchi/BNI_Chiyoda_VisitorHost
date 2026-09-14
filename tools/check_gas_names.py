#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GASへ送るファイル名の衝突を検出する。

Google Apps Script はファイルを拡張子なしの名前で識別するため、
foo.js と foo.html は同じ「foo」として衝突し、clasp push が
「A file with this name already exists in the current project」で失敗する。

    python3 tools/check_gas_names.py
"""
import os, re, sys, fnmatch

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def ignored_patterns():
    p = os.path.join(ROOT, '.claspignore')
    if not os.path.exists(p):
        return []
    out = []
    with open(p, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#'):
                out.append(line)
    return out

def is_ignored(name, pats):
    return any(fnmatch.fnmatch(name, p) or fnmatch.fnmatch(name, p.replace('/**', '/*')) for p in pats)

def main():
    pats = ignored_patterns()
    seen = {}
    for name in sorted(os.listdir(ROOT)):
        if not name.endswith(('.js', '.html')):
            continue
        if is_ignored(name, pats):
            continue
        base = os.path.splitext(name)[0]
        seen.setdefault(base, []).append(name)

    clashes = {k: v for k, v in seen.items() if len(v) > 1}
    if clashes:
        print('NG: GAS上で名前が衝突するファイルがあります（clasp push が失敗します）')
        for k, v in sorted(clashes.items()):
            print('  「%s」 ← %s' % (k, ' / '.join(v)))
        print('\n対処: .js 側の名前を変えてください（HTML名はコードから参照されるため触らないこと）。')
        return 1
    print('OK: 名前の衝突はありません（%d ファイル）' % sum(len(v) for v in seen.values()))

    # ダイアログが参照するHTMLの実在も確認する
    missing = []
    for name in os.listdir(ROOT):
        if not name.endswith('.js') or is_ignored(name, pats):
            continue
        src = open(os.path.join(ROOT, name), encoding='utf-8').read()
        for ref in re.findall(r"create(?:HtmlOutput|Template)FromFile\('([^']+)'\)", src):
            if not os.path.exists(os.path.join(ROOT, ref + '.html')):
                missing.append((name, ref))
    if missing:
        print('NG: 参照先のHTMLが見つかりません')
        for js, ref in missing:
            print('  %s → %s.html' % (js, ref))
        return 1
    print('OK: ダイアログの参照先HTMLはすべて存在します')
    return 0

if __name__ == '__main__':
    sys.exit(main())
