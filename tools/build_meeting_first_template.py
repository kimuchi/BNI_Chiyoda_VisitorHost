#!/usr/bin/env python3
"""前半スライドの「出力」を、差し込み口付きのテンプレートに作り変える。

毎週手で書き換えている箇所（メインプレゼンのお2人、メンバーシップ委員会による報告）に
{{ }} を入れる。デザイン・座標・フォント・画像には触らない。

    python3 tools/build_meeting_first_template.py <出力.pptx> <テンプレート.pptx>
"""
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    src, dst = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
    work = tempfile.mkdtemp(prefix='mtgtpl_')
    try:
        z = zipfile.ZipFile(src)
        names = [n for n in z.namelist() if not n.endswith('/')]
        for n in names:
            fp = os.path.join(work, n)
            os.makedirs(os.path.dirname(fp), exist_ok=True)
            with open(fp, 'wb') as f:
                f.write(z.read(n))

        r = subprocess.run(['node', 'tools/make_meeting_first_template.js', work], cwd=ROOT)
        if r.returncode != 0:
            sys.exit(r.returncode)

        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as out:
            for n in names:
                out.write(os.path.join(work, n), n)
    finally:
        shutil.rmtree(work, ignore_errors=True)

    print('%s  %.1fMB → %.1fMB'
          % (os.path.basename(dst), os.path.getsize(src) / 1048576, os.path.getsize(dst) / 1048576))


if __name__ == '__main__':
    main()
