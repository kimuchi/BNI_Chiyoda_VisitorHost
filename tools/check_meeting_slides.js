// 定例会スライドの書き換えを確かめる。
//   ・{{ }} が無いページの「第○回」「○年○月○日」を形で見つけて直せるか
//   ・コアバリューのページを1枚だけ表示にできるか
// 本番と同じ ooxml.js / meeting_slides_srv.js / routine_srv.js を使う。
//
//   node tools/check_meeting_slides.js

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const sandbox = {
  console,
  SpreadsheetApp: {}, PropertiesService: {}, Utilities: {},
  getSS_: () => { throw new Error('使わない'); },
  getMembersList: () => [],
  normName_: (s) => String(s || '').replace(/[\s　]/g, ''),
  matchInviterToMember: (s) => s,
};
sandbox.global = sandbox;
vm.createContext(sandbox);
for (const f of ['ooxml.js', 'routine_srv.js', 'meeting_slides_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), sandbox, { filename: f });
}
const F = sandbox;

let fails = 0, checks = 0;
function ck(cond, msg) { checks++; if (!cond) { fails++; console.log('  NG ' + msg); } }

// --- 段落を組み立てる小道具（PowerPointはよく文字をランに分ける）---
const para = (...runs) => '<a:p>' + runs.map((t) =>
  `<a:r><a:rPr lang="ja-JP"/><a:t>${t}</a:t></a:r>`).join('') + '</a:p>';
const slide = (...paras) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>'
  + `<p:sp><p:txBody>${paras.join('')}</p:txBody></p:sp>`
  + '</p:spTree></p:cSld></p:sld>';
const textOf = (xml) => (xml.match(/<a:t>([^<]*)<\/a:t>/g) || [])
  .map((t) => t.replace(/<\/?a:t>/g, '')).join('');

// ===== 1. 第○回・日付の書き換え =====
const d = new Date(2026, 8, 30);            // 2026-09-30
const rules = F.meetingPatternRules_('535', d);

const cases = [
  // [入力の段落, 期待する文字]
  [para('第5', '26回 2026年', '07月22日'), '第535回 2026年09月30日'],   // ランをまたぐ
  [para('第526回'), '第535回'],
  [para('2026年7月22日'), '2026年9月30日'],                            // 0埋めなしはそのまま
  [para('2026/07/22'), '2026/09/30'],
  [para('2026-7-22'), '2026-9-30'],
  [para('第 526 回'), '第535回'],
  [para('ビジター募集中です'), 'ビジター募集中です'],                    // 触らない
  [para('第526回のあと 2026年07月22日 に開催'), '第535回のあと 2026年09月30日 に開催'],
];
console.log('■ 第○回・日付の書き換え');
for (const [input, want] of cases) {
  const r = F.replacePatternsInXml_(input, rules);
  const got = textOf(r.xml);
  ck(got === want, `「${textOf(input)}」→「${got}」（「${want}」のはず）`);
  console.log(`  ${got === want ? 'OK' : 'NG'}  ${textOf(input)}  →  ${got}`);
}

// 回数だけ・日付だけの指定でも壊れないこと
ck(F.replacePatternsInXml_(para('第526回 2026年07月22日'), F.meetingPatternRules_('', d)).xml.indexOf('第526回') >= 0,
   '開催回を渡さないときは第○回を触らない');
ck(F.replacePatternsInXml_(para('第526回 2026年07月22日'), F.meetingPatternRules_('535', null)).xml.indexOf('2026年07月22日') >= 0,
   '開催日を渡さないときは日付を触らない');

// ===== 2. コアバリューのページ =====
console.log('\n■ コアバリューのページ');
const VALUES = ['Givers Gain', 'Building Relationships', 'Lifelong Learnign',
                'Traditions+Innovation', 'Positive Atitude', 'Accountability', 'Recognition'];
function deck(extra) {
  const parts = {};
  VALUES.forEach((v, i) => {
    parts[`ppt/slides/slide${i + 1}.xml`] =
      { _x: slide(para('BNIの7つのコアバリュー'), para(v)).replace('<p:sld ', '<p:sld show="0" ') };
  });
  // 紛らわしいページ：ふつうの文に「責任」「伝統」が出てくる
  parts['ppt/slides/slide8.xml'] = { _x: slide(para('一般規定：メンバーの責任について'), para('伝統ある運営')) };
  parts['ppt/slides/slide9.xml'] = { _x: slide(para('本日のリファーラル数')) };
  if (extra) parts['ppt/slides/slide10.xml'] = { _x: extra };
  return parts;
}
// xmlOf_/putXml_ が使える形にする
sandbox.xmlOf_ = (m, p) => (m[p] ? m[p]._x : null);
sandbox.putXml_ = (m, p, x) => { m[p] = { _x: x }; };

const parts = deck();
const res = F.applyCoreValue_(parts, 'Lifelong Learning（生涯学習）');
console.log('  ' + res.message);
ck(res.value === 'Lifelong Learning', 'コアバリューの判別: ' + res.value);
ck(res.shown === 1, '表示にしたページ数: ' + res.shown + '（1のはず）');
ck(res.hidden === 6, '非表示にしたページ数: ' + res.hidden + '（6のはず）');
const shownIdx = VALUES.indexOf('Lifelong Learnign') + 1;
ck(!/show="0"/.test(parts[`ppt/slides/slide${shownIdx}.xml`]._x), '選んだページが表示になっていない');
ck(/show="0"/.test(parts['ppt/slides/slide1.xml']._x), '選ばなかったページが非表示になっていない');
ck(!/show=/.test(parts['ppt/slides/slide8.xml']._x), 'ふつうの文の「責任・伝統」を巻き込んでいる');
ck(!/show=/.test(parts['ppt/slides/slide9.xml']._x), '関係ないページを触っている');

// コアバリューのページが無いテンプレートでは、何もしないこと
const few = { 'ppt/slides/slide1.xml': { _x: slide(para('本日のリファーラル数')) } };
const r2 = F.applyCoreValue_(few, 'Givers Gain');
ck(r2.shown === 0 && r2.hidden === 0, 'コアバリューのページが無いのに触っている');
console.log('  ' + r2.message);

// 判別できない指定は何もしない
ck(F.applyCoreValue_(deck(), '対面BOD') === null, 'コアバリューでない指定で動いている');

console.log(`\n検査 ${checks} 件`);
if (fails) { console.log(`NG: ${fails} 件`); process.exit(1); }
console.log('OK: すべて指示どおりです');
