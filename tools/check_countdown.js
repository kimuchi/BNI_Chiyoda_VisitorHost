// カウントダウンの秒数を変えても、音（ベル・動画）の指示が残ることを確かめる（member_presen_srv.js の mpSetCountdown_）。
//
//   node tools/check_countdown.js
//
// 実物の雛形は使わず、同じ作りの見本のページをここで組み立てる。
//   ・ベルの見本 … 20秒のカウントダウン（数字の箱 0〜20）のあと、最後にベルが鳴る（p:audio）
//   ・卵時計の見本 … 最初に卵時計の動画（音つき）が流れ、そのあとで45秒のカウントダウン。
//                    動画をクリックすると一時停止（interactiveSeq）。BNIの公式のスライドと同じ作り
// 確かめること
//   ・数字の箱が新しい秒数ぶんになり、1秒ごとに消える指示も同じ数になる
//   ・ベルは新しいカウントダウンの最後（元と同じ間隔）に鳴る。卵時計の動画は最初に流れ、カウントダウンはそのあと
//   ・動画・音声の登録（p:video・p:audio）とクリックでの一時停止が残る
//   ・指示の番号（p:cTn id）は1からの通し番号で、ほかの指示を指す番号（p:tn val）も合っている
//   ・始め方（元のまま／クリック／すぐ）と自動で次へ（advTm）

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

const sb = { console };
vm.createContext(sb);
for (const f of ['ooxml.js', 'member_presen_srv.js']) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sb, { filename: f });

// --- 見本のページ ---
const box = (id, n) => '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="TextBox ' + id + '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
  + '<p:spPr><a:xfrm><a:off x="7000000" y="4700000"/><a:ext cx="1181100" cy="914400"/></a:xfrm></p:spPr>'
  + '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="ja-JP" sz="4000"><a:solidFill><a:srgbClr val="64666A"/></a:solidFill></a:rPr>'
  + '<a:t>' + n + '</a:t></a:r></a:p></p:txBody></p:sp>';
const media = (id, kind) => '<p:pic><p:nvPicPr><p:cNvPr id="' + id + '" name="' + kind + '"/><p:cNvPicPr/><p:nvPr>'
  + (kind === 'video' ? '<a:videoFile r:link="rId2"/>' : '<a:audioFile r:link="rId2"/>') + '</p:nvPr></p:nvPicPr>'
  + '<p:blipFill><a:blip r:embed="rId3"/></p:blipFill><p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr></p:pic>';
const hide = (idA, idB, idC, delay, spid) => '<p:par><p:cTn id="' + idA + '" fill="hold"><p:stCondLst><p:cond delay="' + delay + '"/></p:stCondLst><p:childTnLst>'
  + '<p:par><p:cTn id="' + idB + '" presetID="1" presetClass="exit" presetSubtype="0" fill="hold" grpId="0" nodeType="afterEffect">'
  + '<p:stCondLst><p:cond delay="1000"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="' + idC + '" dur="1" fill="hold">'
  + '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="' + spid + '"/></p:tgtEl>'
  + '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set>'
  + '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>';
const call = (idA, idB, idC, delay, spid, cmd, node) => '<p:par><p:cTn id="' + idA + '" fill="hold"><p:stCondLst><p:cond delay="' + delay + '"/></p:stCondLst><p:childTnLst>'
  + '<p:par><p:cTn id="' + idB + '" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="' + node + '">'
  + '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="' + cmd + '"><p:cBhvr>'
  + '<p:cTn id="' + idC + '" dur="1247" fill="hold"/><p:tgtEl><p:spTgt spid="' + spid + '"/></p:tgtEl></p:cBhvr></p:cmd>'
  + '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>';
function slide(opts) {
  // 数字の箱：文書順は 0 が奥、最大が手前。id は 10 から
  let shapes = '', ids = [];
  for (let n = 0; n <= opts.sec; n++) { ids[n] = 10 + n; shapes += box(10 + n, n); }
  shapes += media(opts.mediaId, opts.kind);
  let steps = '', id = 4;
  if (opts.kind === 'video') { steps += call(id, id + 1, id + 2, 0, opts.mediaId, 'playFrom(0.0)', 'withEffect'); id += 3; }
  const off = opts.kind === 'video' ? 2090 : 0;
  for (let i = 0; i < opts.sec; i++) { steps += hide(id, id + 1, id + 2, off + i * 1000, ids[opts.sec - i]); id += 3; }
  if (opts.kind === 'audio') { steps += call(id, id + 1, id + 2, off + opts.sec * 1000, opts.mediaId, 'playFrom(0.0)', 'afterEffect'); id += 3; }
  const start = '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>';
  const node = opts.kind === 'video'
    ? '<p:video><p:cMediaNode vol="80000"><p:cTn id="' + (id++) + '" repeatCount="21000" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="' + opts.mediaId + '"/></p:tgtEl></p:cMediaNode></p:video>'
      + '<p:seq concurrent="1" nextAc="seek"><p:cTn id="' + (id++) + '" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="' + opts.mediaId + '"/></p:tgtEl></p:cond></p:stCondLst>'
      + '<p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>' + call(id, id + 1, id + 2, 0, opts.mediaId, 'togglePause', 'withEffect') + '</p:childTnLst></p:cTn>'
      + '<p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="' + opts.mediaId + '"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
    : '<p:audio><p:cMediaNode vol="65854" showWhenStopped="0"><p:cTn id="' + (id++) + '" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn><p:tgtEl><p:spTgt spid="' + opts.mediaId + '"/></p:tgtEl></p:cMediaNode></p:audio>';
  if (opts.kind === 'video') id += 3;
  let bld = '';
  for (let n = 1; n <= opts.sec; n++) bld += '<p:bldP spid="' + ids[n] + '" grpId="0" animBg="1"/>';
  return '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    + '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + shapes + '</p:spTree></p:cSld>'
    + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
    + '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
    + '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
    + '<p:par><p:cTn id="3" fill="hold"><p:stCondLst>' + start + '</p:stCondLst><p:childTnLst>' + steps + '</p:childTnLst></p:cTn></p:par>'
    + '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
    + '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
    + node + '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + bld + '</p:bldLst></p:timing></p:sld>';
}

// --- 読み取り ---
function look(xml) {
  const t = xml.slice(xml.indexOf('<p:timing>'), xml.indexOf('</p:timing>') + 11);
  const ids = [...t.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => +m[1]);
  const steps = [];
  // 1つの手順（外側の par）ごとに、時刻と中身（消す箱・再生する図形）
  const re = /<p:par><p:cTn id="\d+" fill="hold"><p:stCondLst><p:cond delay="(\d+)"\/><\/p:stCondLst><p:childTnLst><p:par><p:cTn id="\d+" presetID="\d+" presetClass="(\w+)"[\s\S]*?<p:spTgt spid="(\d+)"\/>/g;
  let m;
  while ((m = re.exec(t)) !== null) steps.push({ delay: +m[1], kind: m[2], spid: m[3] });
  const main = t.slice(t.indexOf('nodeType="mainSeq"'));
  return {
    t, ids, steps,
    serial: ids.length > 0 && ids.every((v, i) => v === i + 1),
    tn: [...t.matchAll(/<p:tn val="(\d+)"\/>/g)].map((x) => +x[1]),
    mainId: +((t.match(/<p:cTn id="(\d+)" dur="indefinite" nodeType="mainSeq"/) || [])[1] || 0),
    hides: steps.filter((s) => s.kind === 'exit'),
    calls: steps.filter((s) => s.kind === 'mediacall'),
    start: main.slice(main.indexOf('<p:stCondLst>'), main.indexOf('</p:stCondLst>')),
    audio: /<p:audio>[\s\S]*<\/p:audio>/.test(t), video: /<p:video>[\s\S]*<\/p:video>/.test(t),
    toggle: /nodeType="interactiveSeq"[\s\S]*cmd="togglePause"/.test(t),
    bld: (t.match(/<p:bldP /g) || []).length,
    boxes: sb.mpCountdownSeconds_(xml),
    wellFormed: !/id="[A-Z]/.test(t) && !/val="[A-Z]/.test(t) && (t.match(/<p:par>/g) || []).length === (t.match(/<\/p:par>/g) || []).length,
  };
}

// ===== 1. ベル（最後に鳴る音）=====
const bell = slide({ sec: 20, kind: 'audio', mediaId: 49 });
let o = look(bell);
ck(o.serial && o.hides.length === 20 && o.calls.length === 1 && o.calls[0].delay === 20000 && o.audio && o.boxes === 20, '見本（ベル）の作り: ' + JSON.stringify({ h: o.hides.length, c: o.calls }));
for (const sec of [30, 10]) {
  const x = sb.mpSetCountdown_(bell, sec), r = look(x);
  ck(r.boxes === sec && r.hides.length === sec, sec + '秒（ベル）: 数字の箱 ' + r.boxes + '・消す指示 ' + r.hides.length);
  ck(r.calls.length === 1 && r.calls[0].spid === '49' && r.calls[0].delay === sec * 1000, sec + '秒（ベル）: ベルが最後に鳴らない ' + JSON.stringify(r.calls));
  ck(r.hides[0].delay === 0 && r.hides[sec - 1].delay === (sec - 1) * 1000, sec + '秒（ベル）: 消す時刻 ' + r.hides[0].delay + '〜' + r.hides[sec - 1].delay);
  ck(r.audio && r.serial && r.wellFormed && r.bld === sec, sec + '秒（ベル）: 音声の登録・番号・並び ' + JSON.stringify({ a: r.audio, s: r.serial, b: r.bld }));
  ck(r.tn.every((v) => v === r.mainId) && /delay="indefinite"/.test(r.start), sec + '秒（ベル）: 始め方が元のまま（クリック待ち）でない');
}
// クリックで始める／すぐ始める
let x = sb.mpSetCountdown_(bell, 15, false);
o = look(x);
ck(/<p:cond delay="0"\/>/.test(o.start) && !/indefinite/.test(o.start) && o.tn.length === 0, 'すぐ始める: ' + o.start);
x = sb.mpSetCountdown_(bell, 15, true);
o = look(x);
ck(/delay="indefinite"/.test(o.start) && o.tn.length === 1 && o.tn[0] === o.mainId, 'クリックで始める: ' + o.start + ' tn=' + o.tn);

// ===== 2. 卵時計の動画（最初に流れる音）=====
const egg = slide({ sec: 45, kind: 'video', mediaId: 56 });
o = look(egg);
ck(o.serial && o.hides.length === 45 && o.calls[0].delay === 0 && o.hides[0].delay === 2090 && o.video && o.toggle, '見本（卵時計）の作り');
for (const sec of [30, 150]) {
  const y = sb.mpSetCountdown_(egg, sec), r = look(y);
  ck(r.hides.length === sec, sec + '秒（卵時計）: 消す指示 ' + r.hides.length);
  ck(r.calls.length >= 1 && r.calls[0].spid === '56' && r.calls[0].delay === 0, sec + '秒（卵時計）: 最初に動画が流れない ' + JSON.stringify(r.calls));
  ck(r.hides[0].delay === 2090 && r.hides[sec - 1].delay === 2090 + (sec - 1) * 1000, sec + '秒（卵時計）: カウントダウンが動画のあとに始まらない ' + r.hides[0].delay);
  ck(r.video && r.toggle && r.serial && r.wellFormed, sec + '秒（卵時計）: 動画の登録・クリックで一時停止・番号 ' + JSON.stringify({ v: r.video, t: r.toggle, s: r.serial }));
  ck(r.tn.every((v) => v === r.mainId), sec + '秒（卵時計）: ほかの指示を指す番号がずれた ' + r.tn + ' / ' + r.mainId);
}
// 自動で次へ：始め方だけを直し、動画の登録（delay="indefinite"）には触らない
const z = sb.mpAutoAdvance_(sb.mpSetCountdown_(egg, 30, true), 33000), rz = look(z);
ck(/<p:cond delay="0"\/>/.test(rz.start) && /<p:video><p:cMediaNode[^>]*><p:cTn[^>]*><p:stCondLst><p:cond delay="indefinite"\/>/.test(rz.t)
   && /advTm="33000"/.test(z), '自動で次へ: 動画の登録の始め方まで変わった／advTm が無い');
ck(!/advTm/.test(sb.mpNoAutoAdvance_(z)), 'クリックで次へ（advTm を外す）');
// 同じ秒数なら作り直さなくても同じ数（mpCountdownSeconds_）
ck(sb.mpCountdownSeconds_(sb.mpSetCountdown_(egg, 30)) === 30 && sb.mpCountdownSeconds_('<p:sld/>') === 0, 'いまの秒数の読み取り');

console.log(`カウントダウンと音: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 秒数を変えても、ベル（最後）・卵時計の動画（最初）・動画と音声の登録・クリックでの一時停止が残る／始め方／自動で次へ');
