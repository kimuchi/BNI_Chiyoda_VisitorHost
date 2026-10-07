// 後半スライド：リファーラル発表のあと、推薦のことばのページに入ったところで音楽を止める（meeting_slides_srv.js の stopMusicAt_）を、
// 作り物のページで確かめる。名前はすべて架空。
//
//   node tools/check_music_stop.js
//
// 確かめること
//   ・止めるページの画面切り替えに「前のサウンドを停止」（<p:sndAc><p:endSnd/>）が付く。切り替えの効果はそのまま。
//     切り替えの無いページには足す（置き場所は clrMapOvr のあと・timing の前）。2通りの書き方（mc:AlternateContent）の両方に
//   ・それより前から鳴っていて、そのページまで届く音楽（「999枚にわたって再生」など）は、そのページの手前で止まる枚数に直す
//     （非表示のページは数えない）。止める合図（onStopAudio）が無ければ足す
//   ・そのページの1枚だけで鳴る音（ベル）・あとのページの音楽・動画には触らない
//   ・推薦のことばのページが非表示の日（定例会中の組が無い）は、そのあとの最初の表示のページで止める
//   ・作成（editMeetingSlides_）でも止める。前半（推薦のことばのページが無い）では何もしない
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

process.env.TZ = 'Asia/Tokyo';
const { makeEnv } = require('./lib_sheet_fake');
const env = makeEnv({ now: new Date(2026, 9, 5, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
}
env.reset([['メンバー名簿', false, [['No', '業種区分', '氏名'], ['1', '', '見本 一郎']]]], {});
F.findPhotoIdForName_ = () => '';

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
  + 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const sp = (id, text) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="t${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>`
  + `<a:xfrm><a:off x="0" y="${id * 100000}"/><a:ext cx="3000000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
  + `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
// 音楽の図形（p:pic の中の a:audioFile）と、鳴らし方（p:audio の cMediaNode）
const audioPic = (id, rid, name) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr><a:audioFile r:link="${rid}"/>`
  + `<p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media r:embed="${rid}m"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>`
  + `<p:blipFill><a:blip r:embed="rIdIcon"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="400000" cy="400000"/></a:xfrm>`
  + `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
const audioTiming = (spid, numSld, withEnd, selfClose) => '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
  + '<p:audio><p:cMediaNode vol="80000"' + (numSld ? ` numSld="${numSld}"` : '') + '>'
  + (selfClose ? '<p:cTn id="7" repeatCount="indefinite" fill="hold" display="0"/>'
    : '<p:cTn id="7" repeatCount="indefinite" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
      + (withEnd ? '<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>' : '') + '</p:cTn>')
  + `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:audio></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
const slideXml = (inner, transition, timing, hidden) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}${hidden ? ' show="0"' : ''}><p:cSld><p:spTree>`
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + inner
  + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' + (transition || '') + (timing || '') + '</p:sld>';
const ALT = '<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="2000"><p:fade/></p:transition></mc:Choice>'
  + '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>';

// 後半の並び：扉（BGM：999枚）→ リファーラル発表3枚（1枚は非表示）＋ 1枚目にはそのページだけのベル → 推薦のことば → 抽選（別の音楽）
function makeDeck(o) {
  o = o || {};
  const S = [
    ['扉', sp(2, 'リファーラル発表') + audioPic(4, 'rId5', 'BGM'), o.noTransition ? '' : '<p:transition spd="slow"/>', audioTiming(4, o.numSld === undefined ? 999 : o.numSld, o.withEnd, o.selfClose)],
    ['発表1', sp(2, '見本 一郎') + audioPic(9, 'rId5', 'ベル'), '', audioTiming(9, 1, true)],
    ['発表2', sp(2, '試験 花子'), '', '', true],
    ['発表3', sp(2, '架空 三郎'), '', ''],
    ['推薦', sp(2, '推薦のことば') + sp(3, '{{推薦のことば1氏名}}'), o.recoTransition === undefined ? '<p:transition spd="slow"><p:fade/></p:transition>' : o.recoTransition, '', o.recoHidden],
    ['抽選', sp(2, '抽選コーナー') + audioPic(6, 'rId5', '抽選の音楽'), '', audioTiming(6, 999, true)],
  ];
  const parts = {};
  F.putXml_(parts, '[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
  F.putXml_(parts, 'ppt/presentation.xml', `<?xml version="1.0"?><p:presentation ${NS}><p:sldIdLst>`
    + S.map((s, i) => `<p:sldId id="${256 + i}" r:id="rId${100 + i}"/>`).join('') + '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>');
  F.putXml_(parts, 'ppt/_rels/presentation.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + S.map((s, i) => `<Relationship Id="rId${100 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('')
    + '</Relationships>');
  S.forEach(([k, inner, tr, tm, hid], i) => {
    F.putXml_(parts, `ppt/slides/slide${i + 1}.xml`, slideXml(inner, tr, tm, hid));
    F.putXml_(parts, `ppt/slides/_rels/slide${i + 1}.xml.rels`, '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="../media/music1.mp3"/>'
      + '<Relationship Id="rId5m" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/music1.mp3"/></Relationships>');
  });
  return parts;
}
const X = (parts, n) => F.xmlOf_(parts, `ppt/slides/slide${n}.xml`);
const numSldOf = (x, spid) => {
  const a = (x.match(new RegExp('<p:audio>(?:(?!</p:audio>)[\\s\\S])*spid="' + spid + '"[\\s\\S]*?</p:audio>')) || [''])[0];
  return +((a.match(/<p:cMediaNode\b[^>]*\snumSld="(\d+)"/) || [])[1] || 1);
};
const transitions = (x) => x.match(/<p:transition\b[\s\S]*?<\/p:transition>|<p:transition\b[^>]*\/>/g) || [];

// ===== 1) 推薦のことばのページで止める =====
{
  const parts = makeDeck({ withEnd: false });
  const r = F.stopMusicAt_(parts, 'ppt/slides/slide5.xml');
  const reco = X(parts, 5), bgm = X(parts, 1);
  ck(r && r.slide === 5 && /推薦のことばのページ（5枚目）/.test(r.message), '1) 止めるページ・お知らせ: ' + J(r));
  ck(J(transitions(reco)) === J(['<p:transition spd="slow"><p:fade/><p:sndAc><p:endSnd/></p:sndAc></p:transition>']),
     '1) 推薦のことばのページに「前のサウンドを停止」が付かない（切り替えの効果はそのまま）: ' + J(transitions(reco)));
  // 扉から推薦のことばの手前まで、表示のページは 扉・発表1・発表3 の3枚（発表2は非表示）
  ck(numSldOf(bgm, 4) === 3, '1) 999枚にわたって鳴る音楽の枚数が、推薦のことばの手前（表示のページ3枚）にならない: ' + numSldOf(bgm, 4));
  ck(/<p:stCondLst><p:cond delay="indefinite"\/><\/p:stCondLst><p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt\/><\/p:tgtEl><\/p:cond><\/p:endCondLst><\/p:cTn>/.test(bgm),
     '1) 「前のサウンドを停止」で止まる合図（onStopAudio）が足されない・置き場所が違う');
  ck(numSldOf(X(parts, 2), 9) === 1 && X(parts, 2) === slideXml(sp(2, '見本 一郎') + audioPic(9, 'rId5', 'ベル'), '', audioTiming(9, 1, true)),
     '1) そのページだけで鳴るベルを変えた');
  ck(numSldOf(X(parts, 6), 6) === 999 && !/endSnd/.test(X(parts, 6)), '1) あとのページ（抽選）の音楽を変えた');
  ck(J(r.fixed) === J(['BGM（1枚目から3枚）']), '1) 直した音楽の一覧: ' + J(r.fixed));
}
// 止める合図が初めからある・cTn が空の書き方・音楽の設定が「1枚だけ」
{
  const p1 = makeDeck({ withEnd: true });
  F.stopMusicAt_(p1, 'ppt/slides/slide5.xml');
  ck((X(p1, 1).match(/onStopAudio/g) || []).length === 1 && numSldOf(X(p1, 1), 4) === 3, '1) 止める合図があるときに、2つにした');
  const p2 = makeDeck({ selfClose: true });
  F.stopMusicAt_(p2, 'ppt/slides/slide5.xml');
  ck(/<p:cTn id="7" repeatCount="indefinite" fill="hold" display="0"><p:endCondLst><p:cond evt="onStopAudio"[\s\S]*?<\/p:endCondLst><\/p:cTn>/.test(X(p2, 1)),
     '1) 中身の無い cTn に止める合図を足せない: ' + (X(p2, 1).match(/<p:audio>[\s\S]*?<\/p:audio>/) || [''])[0]);
  const p3 = makeDeck({ numSld: 2 });                          // 2枚で止まる設定（推薦のことばまで届かない）
  const r3 = F.stopMusicAt_(p3, 'ppt/slides/slide5.xml');
  ck(numSldOf(X(p3, 1), 4) === 2 && !r3.fixed.length, '1) 推薦のことばまで届かない音楽を変えた');
}
// 推薦のことばのページに切り替えが無い・2通りの書き方
{
  const p1 = makeDeck({ recoTransition: '' });
  F.stopMusicAt_(p1, 'ppt/slides/slide5.xml');
  ck(/<\/p:clrMapOvr><p:transition><p:sndAc><p:endSnd\/><\/p:sndAc><\/p:transition><\/p:sld>/.test(X(p1, 5)), '1) 切り替えの無いページに「前のサウンドを停止」を足せない');
  const p2 = makeDeck({ recoTransition: ALT });
  F.stopMusicAt_(p2, 'ppt/slides/slide5.xml');
  ck((X(p2, 5).match(/<p:fade\/><p:sndAc><p:endSnd\/><\/p:sndAc><\/p:transition>/g) || []).length === 2, '1) 2通りの書き方の両方に付かない: ' + J(transitions(X(p2, 5))));
  const p3 = makeDeck({ recoTransition: '<p:transition><p:sndAc><p:stSnd><p:snd r:embed="rId9" name="x.wav"/></p:stSnd></p:sndAc></p:transition>' });
  F.stopMusicAt_(p3, 'ppt/slides/slide5.xml');
  ck(J(transitions(X(p3, 5))) === J(['<p:transition><p:sndAc><p:endSnd/></p:sndAc></p:transition>']), '1) 別の音を鳴らす設定を「停止」に替えない: ' + J(transitions(X(p3, 5))));
}

// ===== 2) 推薦のことばのページが非表示の日：そのあとの最初の表示のページで止める =====
{
  const parts = makeDeck({ recoHidden: true });
  const r = F.stopMusicAt_(parts, 'ppt/slides/slide5.xml');
  ck(r && r.slide === 6 && /endSnd/.test(X(parts, 6)) && !/endSnd/.test(X(parts, 5)), '2) 推薦のことばが非表示の日に、次の表示のページで止めない: ' + J(r));
  ck(numSldOf(X(parts, 1), 4) === 3, '2) 非表示のページを数えた: ' + numSldOf(X(parts, 1), 4));
}

// ===== 3) 作成（editMeetingSlides_）でも止める。前半（推薦のことばのページが無い）は何もしない =====
{
  const parts = makeDeck({});
  const info = F.editMeetingSlides_(parts, {}, [], { recommendPairs: [] });
  ck(info.musicStop && info.musicStop.slide === 6 && /endSnd/.test(X(parts, 6)) && numSldOf(X(parts, 1), 4) === 3,
     '3) 作成で止めない（定例会中の組が無い日は推薦のことばのページが非表示なので、抽選のページで）: ' + J(info.musicStop));
  const first = makeDeck({});
  F.putXml_(first, 'ppt/slides/slide5.xml', slideXml(sp(2, '倫理規定'), '', ''));
  const before = X(first, 1);
  const i2 = F.editMeetingSlides_(first, {}, [], {});
  ck(!i2.musicStop && X(first, 1) === before, '3) 推薦のことばのページが無いのに音楽を変えた');
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('音楽を止める: 検査 ' + checks + ' 件 OK: 推薦のことばのページで「前のサウンドを停止」・届く音楽の枚数を直す（非表示は数えない）・'
  + '止める合図・ベルとあとの音楽はそのまま・非表示の日は次の表示のページ・作成でも');
