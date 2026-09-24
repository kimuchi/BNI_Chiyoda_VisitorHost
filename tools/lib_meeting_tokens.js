// 定例会スライドのテンプレート化で共通に使う処理。
// 「左右に写真とお名前が並ぶページ」は、前半のメインプレゼン、
// 後半の推薦のことば・抽選コーナーで同じ作りなので、1か所にまとめてある。

// <a:t[^>]*> は <a:tbl> <a:tc> <a:tabLst/> にも当たってしまうので、直後が記号であることを求める
const textOf = (x) => (x.match(/<a:t(?=[\s>])[^>]*>([\s\S]*?)<\/a:t>/g) || [])
  .map((t) => t.replace(/<\/?a:t(?=[\s>])[^>]*>/g, '')).join('');

// 写真より下に並ぶ文字箱を、左右2人ぶん・上から［氏名／会社名／カテゴリー］に割り当てて
// {{○○1氏名}} のような差し込み口に置き換える。
function tokenizeTwoPerson(F, xml, prefix, opts) {
  const o = opts || {};
  const minY = o.minY || 4400000;
  const slideW = o.slideWidth || 12192000;
  const boxes = [];
  for (const r of F.findTagRanges_(xml, 'p:sp')) {
    const seg = xml.substring(r.start, r.end);
    const id = (seg.match(/<p:cNvPr id="(\d+)"/) || [])[1];
    const off = seg.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>/);
    const ext = seg.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
    const t = textOf(seg).trim();
    if (!id || !off || !ext || !t) continue;
    const y = parseInt(off[2], 10);
    if (y < minY) continue;                       // 見出しや飾りは対象外
    boxes.push({ id, y, center: parseInt(off[1], 10) + parseInt(ext[1], 10) / 2, text: t });
  }
  const mid = slideW / 2;
  const groups = [boxes.filter((b) => b.center < mid).sort((a, b) => a.y - b.y),
                  boxes.filter((b) => b.center >= mid).sort((a, b) => a.y - b.y)];
  const FIELDS = ['氏名', '会社名', 'カテゴリー'];
  const report = [];
  groups.forEach((group, gi) => {
    if (group.length !== 3) {
      throw new Error(`${prefix}${gi + 1}人目の文字箱が${group.length}個です（3個のはず）: `
        + JSON.stringify(group.map((b) => b.text)));
    }
    group.forEach((b, i) => {
      const token = `{{${prefix}${gi + 1}${FIELDS[i]}}}`;
      xml = F.setParagraphsInShape_(xml, b.id, [token]);
      report.push(`id=${b.id} 「${b.text}」 → ${token}`);
    });
  });
  return { xml, report };
}

module.exports = { textOf, tokenizeTwoPerson };
