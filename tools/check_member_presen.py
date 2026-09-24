#!/usr/bin/env python3
"""メンバープレゼン生成の通し検査。

実物のテンプレートpptxに対して、本番と同じ member_presen_srv.js / ooxml.js / member_presen.html
を動かし、出来上がりのpptxを機械的に検証する。PowerPointが無い環境でも、
壊れたzip・切れた参照・文字や座標の取り違えは、ここで捕まえられる。

    python3 tools/check_member_presen.py <メンバープレゼンのテンプレート.pptx> [作業用ディレクトリ]

テンプレートはリポジトリに入れていない（15MB近くあるため）。
Driveに置いてあるものをダウンロードして渡すこと。
"""
import os
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def run(cmd):
    print('$ ' + ' '.join(os.path.basename(c) if i == 1 else c for i, c in enumerate(cmd)))
    r = subprocess.run(cmd, cwd=ROOT)
    if r.returncode != 0:
        sys.exit(r.returncode)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    tpl = os.path.abspath(sys.argv[1])
    if not os.path.exists(tpl):
        print('テンプレートが見つかりません: %s' % tpl)
        sys.exit(2)
    work = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else tempfile.mkdtemp(prefix='mpcheck_')
    out = os.path.join(work, 'out')
    os.makedirs(work, exist_ok=True)

    run(['node', 'tools/mp_check_rotation.js'])                   # 業種区分の巡回
    run(['python3', 'tools/mp_harness_prepare.py', tpl, work])   # テンプレートを展開
    run(['node', 'tools/mp_plan.js', work])                      # 画面の組版を通して指示を作る
    run(['node', 'tools/mp_harness.js', work, out])              # サーバー側の組み立てを実行
    run(['python3', 'tools/mp_harness_check.py', work, out])     # 出来上がりを検証
    print('作業用: %s' % work)


if __name__ == '__main__':
    main()
