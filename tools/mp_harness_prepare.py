#!/usr/bin/env python3
"""メンバープレゼン生成の検証用: テンプレートpptxを展開し、写真の見本を用意する。

実物のテンプレートに対して、実際の member_presen_srv.js / ooxml.js を
Nodeで動かすための下ごしらえ。ここではzipを解くだけで、加工はしない。
"""
import json
import os
import struct
import sys
import zipfile

SRC = sys.argv[1]
OUT = sys.argv[2]


def img_size(path):
    with open(path, 'rb') as f:
        b = f.read(64 * 1024)
    if b[:4] == b'\x89PNG':
        return struct.unpack('>II', b[16:24])
    if b[:2] == b'\xff\xd8':
        i = 2
        while i + 9 < len(b):
            if b[i] != 0xFF:
                i += 1
                continue
            mk = b[i + 1]
            if mk == 0xFF:
                i += 1
                continue
            if mk == 0x01 or 0xD0 <= mk <= 0xD9:
                i += 2
                continue
            ln = (b[i + 2] << 8) | b[i + 3]
            if 0xC0 <= mk <= 0xCF and mk not in (0xC4, 0xC8, 0xCC):
                return ((b[i + 7] << 8) | b[i + 8], (b[i + 5] << 8) | b[i + 6])
            if mk == 0xDA:
                break
            i += 2 + ln
    return None


def main():
    parts = os.path.join(OUT, 'parts')
    os.makedirs(parts, exist_ok=True)
    z = zipfile.ZipFile(SRC)
    names = []
    for n in z.namelist():
        if n.endswith('/'):
            continue
        dst = os.path.join(parts, n)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        with open(dst, 'wb') as f:
            f.write(z.read(n))
        names.append(n)
    # 写真の見本にはテンプレート内の画像を流用する。
    # 大きさが分かっているので、切り抜き量を検算できる。
    # 縦横比のちがう3枚を選ぶ（正方形・横長・縦長で切り取られ方が変わるため）。
    cand = []
    for n in sorted(x for x in names if x.startswith('ppt/media/')):
        d = img_size(os.path.join(parts, n))
        if d and d[0] >= 64 and d[1] >= 64:
            cand.append((n, d))
    cand.sort(key=lambda kv: kv[1][0] / kv[1][1])
    photos = {}
    for n, d in ([cand[0], cand[len(cand) // 2], cand[-1]] if len(cand) >= 3 else cand):
        photos[n] = d
    json.dump({'parts': names, 'photos': photos},
              open(os.path.join(OUT, 'manifest.json'), 'w'), ensure_ascii=False, indent=1)
    print('展開: %d パーツ' % len(names))
    for k, v in photos.items():
        print('  見本写真 %s %s' % (k, v))


if __name__ == '__main__':
    main()
