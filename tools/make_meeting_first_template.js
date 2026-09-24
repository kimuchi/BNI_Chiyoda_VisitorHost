// 前半スライドの出力を、差し込み口（{{ }}）付きのテンプレートに変える。
// 本番と同じ ooxml.js を使うので、生成時とまったく同じ見つけ方・入れ方になる。
//
//   node tools/make_meeting_first_template.js <パーツのディレクトリ>

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { textOf, tokenizeTwoPerson } = require('./lib_meeting_tokens.js');

const DIR = process.argv[2];
const sandbox = { console };
sandbox.global = sandbox;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'ooxml.js'), 'utf8'), sandbox, { filename: 'ooxml.js' });
const F = sandbox;

const slideDir = path.join(DIR, 'ppt/slides');
const slides = fs.readdirSync(slideDir).filter((n) => /^slide\d+\.xml$/.test(n))
  .sort((a, b) => parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10));
const read = (n) => fs.readFileSync(path.join(slideDir, n), 'utf8');
const write = (n, x) => fs.writeFileSync(path.join(slideDir, n), x, 'utf8');


let report = [];

// ===== メインプレゼンのページ =====
const mpName = slides.find((n) => textOf(read(n)).indexOf('Main Presenter of the Week') >= 0);
if (!mpName) throw new Error('「Main Presenter of the Week」のページが見つかりません');
{
  const r = tokenizeTwoPerson(F, read(mpName), 'メインプレゼン');
  write(mpName, r.xml);
  r.report.forEach((x) => report.push(`  ${mpName} ${x}`));
}

// ===== メンバーシップ委員会による報告のページ =====
// 表の見出しで見分ける。位置やスライド番号は当てにしない。
function tables(xml) {
  return F.findTagRanges_(xml, 'p:graphicFrame').map((r) => {
    const seg = xml.substring(r.start, r.end);
    const id = (seg.match(/<p:cNvPr[^>]*\sid="(\d+)"/) || [])[1];
    // 見出しは1つ目のマスの文字。PowerPointは1語を複数のランに割ることがあるので、
    // マスの中を全部つないでから見る。
    const tbl = F.findTagRanges_(seg, 'a:tbl')[0];
    let head = '';
    if (tbl) {
      const tc = F.findTagRanges_(seg.substring(tbl.start, tbl.end), 'a:tc')[0];
      if (tc) head = textOf(seg.substring(tbl.start, tbl.end).substring(tc.start, tc.end));
    }
    return { id, head: head.replace(/[\s　]/g, ''), seg };
  }).filter((t) => t.id);
}
const msName = slides.find((n) => tables(read(n)).some((t) => t.head.indexOf('チャプターが求める') === 0));
if (!msName) throw new Error('「チャプターが求める専門分野」の表が見つかりません');
{
  let xml = read(msName);
  const ts = tables(xml);
  const want = ts.find((t) => t.head.indexOf('チャプターが求める') === 0);
  const open = ts.find((t) => t.head.indexOf('開放カテゴリー') === 0);
  const rev = ts.find((t) => t.head.indexOf('審査中') === 0);
  if (!open || !rev) throw new Error('「開放カテゴリー」または「審査中の申込み」の表が見つかりません');

  // 求める専門分野：見出しの次の行から、左上から順に番号を振る。
  // 空のマスは段落にランが無く文字を入れられないので、埋まっているマスの書式を写す。
  {
    const fr = F.findShapeRange_(xml, want.id);
    let seg = xml.substring(fr.start, fr.end);
    const tb = F.findTagRanges_(seg, 'a:tbl')[0];
    let tbl = seg.substring(tb.start, tb.end);
    const trs = F.findTagRanges_(tbl, 'a:tr');
    // 書式の見本＝2行目1列目の段落
    const firstRow = tbl.substring(trs[1].start, trs[1].end);
    const firstCell = firstRow.substring(...(() => { const c = F.findTagRanges_(firstRow, 'a:tc')[0]; return [c.start, c.end]; })());
    const model = (() => { const p = F.findTagRanges_(firstCell, 'a:p')[0]; return firstCell.substring(p.start, p.end); })();

    let n = 0, out = '', prev = 0;
    for (let ri = 1; ri < trs.length; ri++) {
      let tr = tbl.substring(trs[ri].start, trs[ri].end);
      const tcs = F.findTagRanges_(tr, 'a:tc');
      let trOut = '', tprev = 0;
      for (let ci = 0; ci < tcs.length; ci++) {
        n++;
        let tc = tr.substring(tcs[ci].start, tcs[ci].end);
        const body = F.findTagRanges_(tc, 'a:txBody')[0];
        const token = `{{求める専門分野${n}}}`;
        // <a:t[^>]*> は <a:tabLst/> にも当たってしまう。直後が記号であることを求める。
        const para = model.replace(/<a:t(?=[\s>])[^>]*>[\s\S]*?<\/a:t>/, `<a:t>${token}</a:t>`);
        const ps = F.findTagRanges_(tc.substring(body.start, body.end), 'a:p');
        const inner = tc.substring(body.start, body.end);
        const newBody = inner.substring(0, ps[0].start) + para + inner.substring(ps[ps.length - 1].end);
        tc = tc.substring(0, body.start) + newBody + tc.substring(body.end);
        trOut += tr.substring(tprev, tcs[ci].start) + tc;
        tprev = tcs[ci].end;
      }
      trOut += tr.substring(tprev);
      out += tbl.substring(prev, trs[ri].start) + trOut;
      prev = trs[ri].end;
    }
    out += tbl.substring(prev);
    seg = seg.substring(0, tb.start) + out + seg.substring(tb.end);
    xml = xml.substring(0, fr.start) + seg + xml.substring(fr.end);
    report.push(`  ${msName} 表id=${want.id} チャプターが求める専門分野 → {{求める専門分野1}}〜{{求める専門分野${n}}}`);
  }

  xml = F.setTableCellText_(xml, open.id, 1, 0, '{{開放カテゴリー}}');
  xml = F.setTableCellText_(xml, rev.id, 1, 0, '{{審査中カテゴリー}}');
  report.push(`  ${msName} 表id=${open.id} 開放カテゴリー → {{開放カテゴリー}}`);
  report.push(`  ${msName} 表id=${rev.id} 審査中の申込み → {{審査中カテゴリー}}`);
  write(msName, xml);
}

console.log('差し込み口を入れました:');
report.forEach((r) => console.log(r));
