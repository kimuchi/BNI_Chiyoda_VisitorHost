// 事前MTG（朝イチMTG）のパワポ（premtg_srv.js）を、ルーティンチェックシートの写しで作ってみる。
// 役職ごとの入力と同じ保存（saveRoleInput）で架空の内容を書き込み、同梱の既定のひな形から作る。
//
//   node tools/check_premtg.js <routine.json> <members.json> [出力先のフォルダ]
//
// 出力先のフォルダを渡すと、できたpptx（既定のひな形・並べ替えたひな形の2つ）を書き出す
// （LibreOfficeで画像にして見た目を確かめる用）。
//
// 確かめること
//   ・人数の数え方（「代理・欠席」の欄のお名前。かっこ・遅刻・代理：・時点でなし など）
//   ・まとめ：直近のイベント／お願い事項（三役・【役職】つき）／定例会関連（ツールと同じ並び）
//   ・役職のページ：役職の順・180文字以下同士は2人で1枚・共有事項が無い（なし）役職は作らない
//   ・写真（担当者の写真・丸の枠に合わせた切り抜き）・帯の色（三役・コーディネーター・委員会）
//   ・文字の大きさ（多い週は小さく、まとめの3つのまとまりが重ならない）
//   ・並び（まとめ → 役職のページ）・ひな形のページや {{ }} が残らない・pptxとして壊れていない
//   ・登録したひな形（1人のページが無い・まとめが横並び・ページを足した）でも作れる

const fs = require('fs');
const path = require('path');
const { makeRoleServer, makeBlob } = require('./lib_role_fixture');
const { readZip, writeZip } = require('./lib_zip');

const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

const S = makeRoleServer(process.argv[2], process.argv[3]);
const F = S.F;
const OUT_DIR = process.argv[4] || '';
const DATE = '2026/09/30';

// --- 写真：担当者ごとに縦横比の違う小さなPNG（誰の写真が入ったか見分けられるように）---
const zlib = require('zlib');
function pngOf(w, h, rgb) {
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(zlib.crc32(td));
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const raw = Buffer.alloc((w * 3 + 1) * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) raw.set(rgb, y * (w * 3 + 1) + 1 + x * 3);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr),
                        chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}
const PHOTO = {};                                   // 写真のID → 中身
const PHOTO_OF = {};                                // お名前（空白なし）→ 写真のID
S.MEMBERS.forEach((m, i) => {
  if (/^金井/.test(m.name)) return;                  // 写真が無い方（仮の画像のまま）
  const id = 'photo' + i;
  PHOTO[id] = pngOf(30 + (i % 3) * 10, 40, [(i * 37) % 256, (i * 91) % 256, (i * 53) % 256]);
  PHOTO_OF[m.name.replace(/[\s　]/g, '')] = id;
});
F.findPhotoIdForName_ = (n) => PHOTO_OF[String(n).replace(/[\s　]/g, '')] || '';
F.DriveApp = { getFileById: (id) => ({ getBlob: () => makeBlob(PHOTO[id], 'image/png', id + '.png'), getName: () => id }) };
F.CacheService = undefined;
let SAVED = null;
F.saveOutputFile_ = (blob, name) => { SAVED = { blob, name }; return { id: 'out', url: 'https://example/' + name, downloadUrl: 'https://example/dl/' + name }; };

// ===================== 人数の数え方 =====================
const COUNT = [
  ['小西さん、長尾さん', 2], ['なし', 0], ['村井さん', 1], ['梅田さん　相田さん', 2],
  ['辻さん代理：増井 登志さん', 1], ['船越さん代理：松尾 昂平さん、徳永さん代理：高岡 英明さん', 2],
  ['なし　遅刻：宮崎さん', 0], ['高村さん（当欠）、長尾さん（当欠）', 2],
  ['杉山さん（定例会は参加せず、アフターMTGで挨拶のみ）／遅刻：小西さん、田村浩さん', 1],
  ['細川さん、羽生さん、深井さん(当欠)、※岡林さん遅刻', 3], ['※福元さん遅刻', 0], ['なし／遅刻：岩渕さん', 0],
  ['打越さん（4/1定例会～5/27定例会：6月更新も含めGW明けに本人とコンタクト）\n田村浩美さん（4/8定例会～5/13定例会　その後未定）', 2],
  ['なし（10/1火曜8時時点）', 0], ['24日21時時点でなし', 0], ['伊原さん（見本花代(ミホンハナヨ)さん）', 1],
  ['内野さん→小畑大地さん(8/20火曜8時時点）', 1], ['熊田さん（代理：見本健一 　株式会社熊田建装　外壁修繕工事', 1],
  ['小野塚さん復帰につきプレゼン・リファーラルスライド変更', 1], ['大庭ED', 1], ['坂上アンバサダー、大庭ED', 2],
  ['休会', 0], ['', 0],
];
for (const [text, n] of COUNT) {
  const got = F.premtgCountNames_(text);
  ck(got === n, `人数の数え方「${text.replace(/\n/g, '⏎')}」→ ${got}（${n} のはず）`);
}

// ===================== 役職ごとの入力で保存する =====================
function saveAs(role, values) {
  const ctx = F.getRoleInputContext(DATE, role);
  const entries = [];
  for (const [title, value] of Object.entries(values)) {
    const [parent, t] = title.includes('›') ? title.split('›').map((x) => x.trim()) : ['', title];
    const id = ctx.order.find((k) => ctx.items[k].title === t && (!parent || ctx.items[k].parent === parent)
      && ctx.items[k].roles.includes(role));
    if (!id) throw new Error(`${role} に「${title}」の項目がありません`);
    entries.push({ id, value, orig: ctx.items[id].value });
  }
  const r = F.saveRoleInput(DATE, role, entries);
  if (!r.ok || (r.skipped && r.skipped.length)) throw new Error(`${role} の保存: ${r.message} ${JSON.stringify(r.skipped || [])}`);
}
const LONG = '・来期の役割分担について、各委員会の引継ぎ資料を今週中に共有フォルダへ入れてください。\n'
  + '・引継ぎの打ち合わせは、10/3（土）10:00から、オンラインで行います。参加できない方は、事前に担当の委員長へご連絡ください。\n'
  + '・定例会の進行表は、今週から新しい書式になります。変更点は、入室の時間・ブレイクアウトルームの割り振り・アフターミーティングの進め方の3点です。\n'
  + '・メンバー専用ページの「各種申請フォーム」が新しくなりました。古いリンクは使えませんので、ブックマークを直してください。\n'
  + '・チャプター運営費のお支払いがまだの方は、月末までにお願いします。\n'
  + '・来週の定例会は、会場の都合で入室の受付を6:45から始めます。遅れそうな方は、遅刻・欠席の担当までご連絡ください。\n'
  + '・推薦のことばの原稿は、月曜日の午前までにPDFで書記兼会計へお送りください。修正が終わったものを提出とします。\n'
  + '・メインプレゼンの資料は2週間前までに、景品は業務に関連するもの（消えものでないもの）をご用意ください。\n'
  + '・事前MTGの議事録は毎回Facebookに載せますので、必ずご確認ください。\n'
  + '・会場の予約は来月分まで済んでいます。変更がある場合は、書記兼会計までお知らせください。\n'
  + '・名札とバッジは定例会の前に必ず確認してください。無くした方は予備をお渡しします。\n'
  + '・リファーラルの記入は、定例会の当日中にお願いします。翌日以降の記入は集計に入りません。\n'
  + '・ビジターの方への声かけは、アフターミーティングの前に済ませてください。';
saveAs('president', {
  'お願い事項（プレジデントから）': '・代理の登録は原則として月曜日18時までにお願いします。\n　代理を立てるときは、まずバイスプレジデントへひとことお知らせください。',
  '今週の共有事項（プレジデント）': '・今週の承認コーナー：見本チームの皆さん（ビジター交流会の開催）',
});
saveAs('vice', {
  'お願い事項（バイスプレジデントから）': '・ビジター紹介のときはミュートを外して拍手をお願いします。',
  '人数：ビジター': '3', '人数：ゲスト': '1', '人数：見学': '0', '人数：リージョン参加者': '',
  '代理・欠席 › 代理': '小西さん、長尾さん', '代理・欠席 › 欠席': '村井さん（当欠）／遅刻：高村さん', '代理・欠席 › 医療欠席': 'なし',
  '今週の共有事項（バイスプレジデント）': '・チャプターの人数：目標53名、いま49名です。\n・今週の動きの発表について、1分以内でお願いします。',
});
saveAs('secretary', {
  '直近のイベント': '・10/3　引継ぎの打ち合わせ 10:00～（オンライン）',
  'お願い事項（書記兼会計から）': '・Facebookの投稿を確認したら、いいね！でお知らせください。',
  '卒業コメント': 'なし',
  '今週の共有事項（書記兼会計）': LONG,
});
saveAs('vhc', { '今週の共有事項（ビジターホストコーディネーター）': '・定例会のあと、フォローの連絡をしてスプレディングに記録を残してください。' });
saveAs('mentor', { '今週の共有事項（メンターコーディネーター）': 'なし' });            // ページを作らない
saveAs('ec', { '今週の共有事項（エデュケーションコーディネーター）': '・今週のエデュケーション：見本さん\n・次週：見本二郎さん' });
saveAs('gbc', { '今週の共有事項（グローバルビジネスコーディネーター）': '・海外のメンバーとつながっている方は、情報をお寄せください。' });
saveAs('web', { '今週の共有事項（webマスター）': '・入室したら、お名前の表記を「番号_お名前」に直してください。' });
saveAs('support', { '今週の共有事項（メンバーサポート委員）': '・サポートが必要な方は、委員までご連絡ください。' });
// training・event・bcp は書かない（ページを作らない）
saveAs('spreading', { '今週の共有事項（スプレディング委員）': '・略歴シート・GAINSを最新にしてください。' });

// ===================== 材料（作る前の確かめ）=====================
const pv = F.getPreMeetingPreview(DATE);
ck(pv.ok, '作る前の確かめ: ' + pv.message);
const sm = pv.summary || {};
ck(JSON.stringify(sm['直近のイベント']) === JSON.stringify(['・10/3　引継ぎの打ち合わせ 10:00～（オンライン）']), '直近のイベント: ' + JSON.stringify(sm['直近のイベント']));
ck((sm['お願い事項'] || []).length === 4 && sm['お願い事項'][0].indexOf('【プレジデント】・代理の登録') === 0
   && sm['お願い事項'][1].indexOf('　代理を立てるとき') === 0 && sm['お願い事項'][2].indexOf('【バイスプレジデント】') === 0
   && sm['お願い事項'][3].indexOf('【書記兼会計】') === 0, 'お願い事項: ' + JSON.stringify(sm['お願い事項']));
const ML = sm['定例会関連'] || [];
console.log('定例会関連:\n  ' + ML.join('\n  '));
ck(ML[0] === '・欠席：1名　／　代理：2名　／　医療欠席：0名', '欠席・代理・医療欠席: ' + ML[0]);
ck(ML[1] === '・ビジター：3名', 'ビジター: ' + ML[1]);
ck(/^・ゲスト：1名　／　見学：0名　／　リージョン参加者：—$/.test(ML[2]), 'ゲスト・見学・リージョン: ' + ML[2]);
ck(/^・ウィークリープレゼン：/.test(ML[3]) && /^・スタートアッププレゼン：谷口さん$/.test(ML[4]) && ML[5] === '・メインプレゼン：①山内さん　②金井さん',
   'ウィークリー・スタートアップ・メインプレゼン: ' + ML.slice(3, 6).join(' | '));
ck(/^・新入会：—　／　更新：小西さん（伊豆澤さんは10\/7）　／　退会：村井さん、山岸$/.test(ML[6]), '新入会・更新・退会: ' + ML[6]);
ck(ML[7] === '・卒業コメント：なし', '卒業コメント: ' + ML[7]);
ck(pv.blank.some((b) => b.label === '新入会' && b.role === 'バイスプレジデント')
   && pv.blank.some((b) => b.label === '人数：リージョン参加者'), '空欄の項目: ' + JSON.stringify(pv.blank));
const plan = (pv.pages || []).map((pg) => pg.map((r) => r.label).join('+'));
console.log('役職のページ: ' + plan.join(' ／ '));
ck(plan.join(' / ') === 'プレジデント+バイスプレジデント / 書記兼会計 / ビジターホストコーディネーター+エデュケーションコーディネーター / '
   + 'グローバルビジネスコーディネーター+webマスター / メンバーサポート委員+スプレディング委員',
   '役職のページの割り付け: ' + plan.join(' / '));
ck(JSON.stringify(pv.skipped) === JSON.stringify(['メンターコーディネーター', 'トレーニング委員', 'イベント委員＆1to1促進委員', 'BCP委員']),
   '共有事項が無い役職: ' + JSON.stringify(pv.skipped));
ck(pv.template && !pv.template.registered, 'ひな形は既定のもの: ' + JSON.stringify(pv.template));

// ===================== 作る（既定のひな形）=====================
const r = F.generatePreMeetingSlides(DATE);
ck(r.ok && SAVED && SAVED.name === '20260930_BNI事前MTG.pptx', '作成: ' + r.message);
console.log(r.message);
ck(/写真が見つからない方（仮の画像のまま）: 金井 美里/.test(r.message), '写真が無い方のお知らせ: ' + r.message);
const files = readZip(SAVED.blob._buf);
if (OUT_DIR) fs.writeFileSync(path.join(OUT_DIR, 'premtg_default.pptx'), SAVED.blob._buf);
checkDeck(files, '既定', { pages: 6, plan, summaryStacked: true });

// ===================== 登録したひな形（並べ替えたもの）=====================
// 1人のページを消し、まとめのまとまりを横並びにし、最後に「{{月日}}の連絡」のページを足す
const base = readZip(F.premtgBuiltinBlob_()._buf);
const custom = Object.assign({}, base);
{
  const prs = custom['ppt/presentation.xml'].toString();
  custom['ppt/presentation.xml'] = Buffer.from(prs.replace(/<p:sldId id="\d+" r:id="rId4"\/>/, '<p:sldId id="258" r:id="rId9"/>'));
  const rel = custom['ppt/_rels/presentation.xml.rels'].toString();
  custom['ppt/_rels/presentation.xml.rels'] = Buffer.from(rel.replace('</Relationships>',
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide4.xml"/></Relationships>'));
  // 足すページ（3枚目＝1人のページの代わり。1人のページは並びから外すだけでなく中身も消す）
  custom['ppt/slides/slide4.xml'] = Buffer.from(base['ppt/slides/slide3.xml'].toString()
    .replace(/\{\{アイコン1\}\} \{\{役職1\}\}/, '{{月日}}の連絡（{{開催回}}回）').replace(/\{\{氏名1\}\}|\{\{共有事項1\}\}/g, ''));
  custom['ppt/slides/_rels/slide4.xml.rels'] = base['ppt/slides/_rels/slide3.xml.rels'];
  custom['[Content_Types].xml'] = Buffer.from(custom['[Content_Types].xml'].toString()
    .replace('</Types>', '<Override PartName="/ppt/slides/slide4.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>')
    .replace(/<Override PartName="\/ppt\/slides\/slide3\.xml"[^>]*\/>/, ''));
  delete custom['ppt/slides/slide3.xml'];
  delete custom['ppt/slides/_rels/slide3.xml.rels'];
  custom['ppt/_rels/presentation.xml.rels'] = Buffer.from(custom['ppt/_rels/presentation.xml.rels'].toString()
    .replace(/<Relationship Id="rId4"[^>]*\/>/, ''));
  // まとめ：3つの本文の枠を横に並べる（見出しはそのまま）
  let s1 = custom['ppt/slides/slide1.xml'].toString();
  const W = 3600000;
  ['{{直近のイベント}}', '{{お願い事項}}', '{{定例会関連}}'].forEach((tk, i) => {
    const id = F.premtgShapeWith_(s1, tk);
    s1 = F.setShapeGeomEmu_(s1, id, { x: 365760 + i * (W + 120000), y: 1435608, cx: W, cy: 4800000 });
  });
  custom['ppt/slides/slide1.xml'] = Buffer.from(s1);
}
F.PropertiesService.getScriptProperties().setProperty('BNI_TPL_PREMTG_ID', 'custom_tpl');
const CUSTOM_BLOB = makeBlob(writeZip(custom), 'application/zip', '事前MTG_並べ替え.pptx');
const realDrive = F.DriveApp.getFileById;
F.DriveApp = { getFileById: (id) => (id === 'custom_tpl'
  ? { getBlob: () => makeBlob(CUSTOM_BLOB._buf, 'application/zip', 'x'), getName: () => '事前MTG_並べ替え.pptx' }
  : realDrive(id)) };
SAVED = null;
const r2 = F.generatePreMeetingSlides(DATE);
ck(r2.ok && /登録したもの（事前MTG_並べ替え\.pptx）/.test(r2.message), '登録したひな形で作成: ' + r2.message);
const files2 = readZip(SAVED.blob._buf);
if (OUT_DIR) fs.writeFileSync(path.join(OUT_DIR, 'premtg_custom.pptx'), SAVED.blob._buf);
// 1人のページが無いので、隣と組めない役職（書記兼会計）は2人のページの右を空けて入れる。組み方はツールと同じ
checkDeck(files2, '並べ替え', { pages: 6, plan, summaryStacked: false, extraLast: '9/30の連絡（535回）', noSingle: true });

// ひな形のページを消したもの（n … 消すスライドの番号）
function withoutSlides(nums) {
  const v = Object.assign({}, base);
  for (const n of nums) {
    const rid = 'rId' + (n + 1);                                // 既定のひな形は slideN が rId(N+1)
    v['ppt/presentation.xml'] = Buffer.from(v['ppt/presentation.xml'].toString().replace(new RegExp('<p:sldId id="\\d+" r:id="' + rid + '"/>'), ''));
    v['ppt/_rels/presentation.xml.rels'] = Buffer.from(v['ppt/_rels/presentation.xml.rels'].toString().replace(new RegExp('<Relationship Id="' + rid + '"[^>]*/>'), ''));
    v['[Content_Types].xml'] = Buffer.from(v['[Content_Types].xml'].toString().replace(new RegExp('<Override PartName="/ppt/slides/slide' + n + '\\.xml"[^>]*/>'), ''));
    delete v['ppt/slides/slide' + n + '.xml'];
    delete v['ppt/slides/_rels/slide' + n + '.xml.rels'];
  }
  return makeBlob(writeZip(v), 'application/zip', 'x');
}
let VARIANT = null;
F.DriveApp = { getFileById: (id) => (id === 'custom_tpl'
  ? { getBlob: () => makeBlob(VARIANT._buf, 'application/zip', 'x'), getName: () => '事前MTG_変えたもの.pptx' }
  : realDrive(id)) };
// 2人のページが無い → 1人ずつ
VARIANT = withoutSlides([2]);
const pv5 = F.getPreMeetingPreview(DATE);
ck(pv5.ok && pv5.pages.length === 9 && pv5.pages.every((pg) => pg.length === 1) && !pv5.template.error,
   '2人のページが無いひな形の割り付け: ' + JSON.stringify(pv5.pages && pv5.pages.map((pg) => pg.length)) + ' ' + (pv5.template && pv5.template.error));
SAVED = null;
const r5 = F.generatePreMeetingSlides(DATE);
const f5 = readZip(SAVED.blob._buf);
const n5 = (f5['ppt/presentation.xml'].toString().match(/<p:sldId /g) || []).length;
ck(r5.ok && n5 === 10 && /役職のページ 9枚/.test(r5.message), '2人のページが無いひな形: ' + n5 + '枚 ' + r5.message);
// 帯を別の色にしたひな形 → 塗り替えない（2人のページの帯を紺にする）
{
  const v = Object.assign({}, base);
  v['ppt/slides/slide2.xml'] = Buffer.from(base['ppt/slides/slide2.xml'].toString().replace(/C8102E/g, '1F3864'));
  VARIANT = makeBlob(writeZip(v), 'application/zip', 'x');
  SAVED = null;
  const rb = F.generatePreMeetingSlides(DATE), fb = readZip(SAVED.blob._buf);
  const fills = Object.keys(fb).filter((q) => /slides\/slide\d+\.xml$/.test(q)).map((q) => fb[q].toString())
    .filter((x) => /name="帯2"/.test(x))
    .map((x) => shapesOf(x).filter((sh) => /^帯/.test(sh.name)).map((sh) => sh.fill).join(','));
  ck(rb.ok && fills.length === 4 && fills.every((f) => f === '1F3864,1F3864'), '帯を紺にしたひな形は塗り替えない: ' + JSON.stringify(fills));
}
// 役職のページが無い → 確かめで知らせる・まとめだけ作る
VARIANT = withoutSlides([2, 3]);
const pv6 = F.getPreMeetingPreview(DATE);
ck(pv6.ok && /役職のページ（\{\{共有事項1\}\} のあるページ）がありません/.test(pv6.template.error), '役職のページが無いひな形の確かめ: ' + pv6.template.error);
SAVED = null;
const r6 = F.generatePreMeetingSlides(DATE);
ck(r6.ok && /役職のページは作りませんでした/.test(r6.message), '役職のページが無いひな形: ' + r6.message);
// 開けないひな形
F.DriveApp = { getFileById: (id) => { if (id === 'custom_tpl') throw new Error('見つかりません'); return realDrive(id); } };
const pv7 = F.getPreMeetingPreview(DATE);
ck(pv7.ok && /登録したひな形を開けませんでした/.test(pv7.template.error), '開けないひな形の確かめ: ' + pv7.template.error);
const quiet = console.error; console.error = () => {};          // 失敗の記録（ログ）は出さない
const r7 = F.generatePreMeetingSlides(DATE);
console.error = quiet;
ck(!r7.ok && /登録し直してください/.test(r7.message), '開けないひな形で作ったとき: ' + r7.message);
F.DriveApp = { getFileById: realDrive };
F.PropertiesService.getScriptProperties().deleteProperty('BNI_TPL_PREMTG_ID');

// ===================== とても長い共有事項 =====================
saveAs('gbc', { '今週の共有事項（グローバルビジネスコーディネーター）': Array.from({ length: 40 }, (_, i) => '・' + (i + 1) + '件目のお知らせです。海外のメンバーとのつながりについて、詳しい情報をお寄せください。').join('\n') });
SAVED = null;
const r8 = F.generatePreMeetingSlides(DATE);
ck(r8.ok && /いちばん小さくしても枠に収まらない役職: グローバルビジネスコーディネーター/.test(r8.message), 'とても長い共有事項のお知らせ: ' + r8.message);
{
  const f8 = readZip(SAVED.blob._buf), x8 = Object.keys(f8).filter((q) => /slides\/slide\d+\.xml$/.test(q)).map((q) => f8[q].toString())
    .find((x) => /グローバルビジネスコーディネーター/.test(x));
  const sz8 = [...(x8 || '').matchAll(/name="共有事項1"[\s\S]*?<\/p:sp>/g)].map((m) => (m[0].match(/sz="(\d+)"/) || [])[1]);
  ck(sz8[0] === '900', 'とても長い共有事項は9ptまで小さくする: ' + sz8);
}

// ===================== 共有事項が1つも無い週 =====================
S.sheets.forEach((sh) => {
  if (!/ルーティンチェックシート/.test(sh.getName())) return;
  const col = sh._grid[0].indexOf(DATE);
  if (col < 0) return;
  sh._grid.forEach((row) => { if (/^今週の共有事項/.test(String(row[3] || '').trim()) || /^今週の共有事項/.test(String(row[2] || '').trim())) row[col] = ''; });
});
SAVED = null;
const r3 = F.generatePreMeetingSlides(DATE);
const f3 = r3.ok ? readZip(SAVED.blob._buf) : {};
const n3 = (String(f3['ppt/presentation.xml'] || '').match(/<p:sldId /g) || []).length;
ck(r3.ok && n3 === 1 && /役職のページ 0枚/.test(r3.message), '共有事項が無い週はまとめだけ: ' + n3 + '枚 ' + r3.message);

// ===================== 開催日の列が無い =====================
const r4 = F.generatePreMeetingSlides('2030/01/01');
ck(!r4.ok && /見つかりません/.test(r4.message), '列が無い日: ' + r4.message);

// ---------------------------------------------------------------
function textOf(x) { return (x.match(/<a:t>([^<]*)<\/a:t>/g) || []).map((t) => t.replace(/<\/?a:t>/g, '')).join(''); }
function unesc(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&'); }
function shapesOf(xml) {
  const out = [];
  for (const m of xml.matchAll(/<p:(sp|pic)>([\s\S]*?)<\/p:\1>/g)) {
    const seg = m[0];
    const nv = seg.match(/<p:cNvPr\b[^>]*>/)[0];
    const off = seg.match(/<a:off x="(-?\d+)" y="(-?\d+)"\s*\/>/), ext = seg.match(/<a:ext cx="(\d+)" cy="(\d+)"\s*\/>/);
    out.push({ kind: m[1], seg, name: (nv.match(/name="([^"]*)"/) || [])[1] || '',
               x: +off[1], y: +off[2], cx: +ext[1], cy: +ext[2], text: unesc(textOf(seg)),
               paras: (seg.match(/<a:p>[\s\S]*?<\/a:p>/g) || []).map((p) => unesc(textOf(p))),
               sz: [...seg.matchAll(/<a:rPr\b[^>]*\ssz="(\d+)"/g)].map((q) => +q[1] / 100),
               fill: ((seg.split('</p:spPr>')[0].split('<a:ln')[0].match(/<a:solidFill>\s*<a:srgbClr val="(\w+)"/) || [])[1] || ''),
               embed: (seg.match(/r:embed="(rId\d+)"/) || [])[1] || '', srcRect: /<a:srcRect\b/.test(seg) });
  }
  return out;
}

function checkDeck(files, label, o) {
  const prs = files['ppt/presentation.xml'].toString();
  const prels = files['ppt/_rels/presentation.xml.rels'].toString();
  const ct = files['[Content_Types].xml'].toString();
  const rid2 = {};
  for (const m of prels.matchAll(/<Relationship Id="([^"]+)"[^>]*Target="slides\/(slide\d+\.xml)"/g)) rid2[m[1]] = 'ppt/slides/' + m[2];
  const order = [...prs.matchAll(/<p:sldId id="(\d+)" r:id="([^"]+)"\/>/g)].map((m) => rid2[m[2]]);
  const ids = [...prs.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => m[1]);
  ck(order.length === o.pages + (o.extraLast ? 1 : 0), `${label}: 枚数 ${order.length}`);
  ck(order.every((p) => p && files[p]) && new Set(ids).size === ids.length, `${label}: 並びのページがそろっていない ${JSON.stringify(order)}`);
  // 並びに無いスライド・関係・種類の登録が残っていない
  const slideParts = Object.keys(files).filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p));
  ck(slideParts.length === order.length, `${label}: 並びに無いページが残っている ${slideParts.filter((p) => !order.includes(p))}`);
  ck(slideParts.every((p) => ct.includes('PartName="/' + p + '"')) && (ct.match(/slide\+xml/g) || []).length === slideParts.length,
     `${label}: 種類の登録（[Content_Types]）が合わない`);
  ck(Object.values(rid2).every((p) => files[p]), `${label}: presentation.xml.rels に消したページが残っている`);
  // 画像：関係の先がそろっている・使っていない画像は残さない
  const used = new Set();
  for (const p of Object.keys(files).filter((q) => /\.rels$/.test(q))) {
    for (const m of files[p].toString().matchAll(/Target="\.\.\/media\/([^"]+)"/g)) {
      used.add(m[1]);
      ck(!!files['ppt/media/' + m[1]], `${label}: ${p} の画像 ${m[1]} が無い`);
    }
  }
  ck(Object.keys(files).filter((p) => p.startsWith('ppt/media/')).every((p) => used.has(p.slice(10))), `${label}: 使っていない画像が残っている`);
  ck(/Extension="png"/.test(ct), `${label}: png の種類の登録が無い`);

  const sl = order.map((p) => files[p].toString());
  ck(sl.every((x) => !/\{\{/.test(textOf(x))), `${label}: {{ }} が残っている ` + sl.map((x) => (textOf(x).match(/\{\{[^}]*\}\}/g) || []).join(',')).join(' | '));

  // --- まとめ ---
  const s1 = shapesOf(sl[0]);
  ck(s1.some((s) => s.text === '9/30 定例会　朝イチMTG'), `${label}: まとめの見出し ${s1.map((s) => s.text).slice(0, 2)}`);
  const cont = ['・10/3', '【プレジデント】', '・欠席：'].map((h) => s1.find((s) => s.paras[0] && s.paras[0].indexOf(h) === 0));
  ck(cont.every(Boolean), `${label}: まとめの3つのまとまりが見つからない`);
  if (cont.every(Boolean)) {
    ck(cont[1].paras.length === 4 && cont[2].paras.length === 8, `${label}: お願い事項${cont[1].paras.length}行・定例会関連${cont[2].paras.length}行`);
    const sizes = cont.map((c) => Math.max(...c.sz));
    ck(sizes.every((z) => z === sizes[0]) && sizes[0] >= 9 && sizes[0] <= 16, `${label}: まとめの文字の大きさ ${sizes}`);
    if (o.summaryStacked) {
      const labels = ['直近のイベント', 'お願い事項', '定例会関連'].map((t) => s1.find((s) => s.text.indexOf(t) >= 0 && s.cy < 400000));
      ck(labels.every(Boolean), `${label}: まとめの見出しが見つからない`);
      let prevBottom = 0;
      for (let i = 0; i < 3; i++) {
        ck(labels[i].y >= prevBottom && cont[i].y >= labels[i].y + labels[i].cy, `${label}: まとまり${i + 1}が上と重なっている`);
        prevBottom = cont[i].y + cont[i].cy;
      }
      ck(prevBottom <= 6858000 - 182880 + 10, `${label}: まとめが下にはみ出している ${prevBottom}`);
      console.log(`  ${label}: まとめ ${sizes[0]}pt・枠の高さ ${cont.map((c) => (c.cy / 12700).toFixed(0) + 'pt').join('/')}`);
    } else {
      ck(cont.every((c) => c.cy === 4800000), `${label}: 横並びのまとまりの枠が動いた ${cont.map((c) => c.cy)}`);
    }
  }

  // --- 役職のページ ---
  const COLOR = { プレジデント: 'C8102E', バイスプレジデント: 'C8102E', 書記兼会計: 'C8102E',
                  ビジターホストコーディネーター: '2E5C9A', エデュケーションコーディネーター: '2E5C9A',
                  グローバルビジネスコーディネーター: '2E5C9A', webマスター: '3B8763', メンバーサポート委員: '3B8763',
                  スプレディング委員: '3B8763' };
  const HOLDER = { プレジデント: '熊田 龍平', バイスプレジデント: '船越 雄一', 書記兼会計: '原口 雅樹',
                   ビジターホストコーディネーター: '藤本 礼子', エデュケーションコーディネーター: '伊原 良太',
                   グローバルビジネスコーディネーター: '中西 渉', webマスター: '相田 周作', メンバーサポート委員: '山内 登志夫',
                   スプレディング委員: '金井 美里' };
  const got = [];
  for (let i = 1; i < 1 + o.pages - 1; i++) {
    const xml = sl[i], sh = shapesOf(xml);
    const rels = files[order[i].replace('slides/', 'slides/_rels/') + '.rels'].toString();
    const roleBoxes = sh.filter((s) => /^役職\d$/.test(s.name));
    const people = roleBoxes.map((b) => b.text.trim().replace(/^\S+ /, '')).filter(Boolean);
    got.push(people.join('+'));
    for (const b of roleBoxes) {
      const k = b.name.slice(-1), role = b.text.trim().replace(/^\S+ /, '');
      if (!role) {                                         // 2人のページの右を空けた
        ck(o.noSingle && !sh.some((s) => s.name === '写真' + k || s.name === '帯' + k), `${label}: 空けた側に写真・帯が残っている`);
        continue;
      }
      const band = sh.find((s) => s.name === '帯' + k), pic = sh.find((s) => s.name === '写真' + k);
      const nm = sh.find((s) => s.name === '氏名' + k), body = sh.find((s) => s.name === '共有事項' + k);
      ck(band && band.fill === COLOR[role], `${label}: ${role} の帯の色 ${band && band.fill}`);
      ck(nm && nm.text === HOLDER[role] + 'さん', `${label}: ${role} のお名前 ${nm && nm.text}`);
      ck(b.sz.every((z) => z * 1.0 >= 10), `${label}: ${role} の役職名の大きさ ${b.sz}`);
      const target = pic && (rels.match(new RegExp('Id="' + pic.embed + '"[^>]*Target="\\.\\./media/([^"]+)"')) || [])[1];
      const pid = F.findPhotoIdForName_(HOLDER[role]);
      if (pid) {
        ck(target && files['ppt/media/' + target] && Buffer.compare(files['ppt/media/' + target], PHOTO[pid]) === 0,
           `${label}: ${role} の写真が ${HOLDER[role]} さんのものでない（${target}）`);
        ck(pic.srcRect, `${label}: ${role} の写真の切り抜き（丸の枠に合わせる）が無い`);
      } else {
        ck(target === 'premtg_photo.png', `${label}: 写真が無い ${role} は仮の画像のまま（${target}）`);
      }
      // 本文：書いたとおりの段落。長いもの（書記兼会計）は小さくして収める
      ck(body && body.paras.length >= 1 && body.paras.every((p) => p.length > 0), `${label}: ${role} の本文 ${body && body.paras}`);
      if (role === '書記兼会計') {
        ck(body.paras.length === 13 && Math.max(...body.sz) < 18, `${label}: 書記兼会計の長い本文 ${body.paras.length}段落・${body.sz[0]}pt`);
      } else if (!o.noSingle || roleBoxes.length === 2) {
        ck(Math.max(...body.sz) === (roleBoxes.length === 2 ? 15 : 18), `${label}: ${role} の本文の大きさ ${body.sz[0]}pt`);
      }
    }
  }
  ck(got.join(' / ') === o.plan.join(' / '), `${label}: 役職のページの並び ${got.join(' / ')}`);
  if (o.extraLast) {
    ck(textOf(sl[sl.length - 1]).indexOf(o.extraLast) >= 0, `${label}: 足したページの {{月日}} ${textOf(sl[sl.length - 1])}`);
  }
}

console.log(`\n事前MTGのパワポ: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 人数の数え方・まとめ・役職のページ（割り付け・写真・帯の色・文字の大きさ）・並び・登録したひな形');
