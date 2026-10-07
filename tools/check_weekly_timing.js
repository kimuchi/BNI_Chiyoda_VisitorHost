// ウィークリープレゼン（メンバーのページ）の送り方を、作り物のページで確かめる。名前はすべて架空。
//
//   node tools/check_weekly_timing.js
//
// 確かめること
//   ・業種区分の扉ページは1秒で次へ進む（テンプレートに画面切り替えが無いときは足す・あれば時間だけ1秒に・
//     クリックで始まる動きはすぐ始まるように）。個人ページの「自動で次へ」を選ばなかったときも
//   ・前半に差し込んだとき：扉ページがあれば「保存済みのタイミングを使用」を入れる合図（auto）を返す。お知らせにも出す
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

const F = { console: { log() {}, warn() {}, error() {} } };
vm.createContext(F);
for (const f of ['ooxml.js', 'chapter_srv.js', 'splice_srv.js', 'member_presen_srv.js', 'referral_srv.js', 'meeting_slides_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
}

// ---- 作り物の扉ページ（業種区分名 11・次の方 31・写真 2・表 6）----
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const xf = (x, y, cx, cy) => `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
const sp = (id, text) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
  + `<p:spPr>${xf(id * 100000, 100000, 3000000, 500000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
  + `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP" sz="2400"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
const pic = `<p:pic><p:nvPicPr><p:cNvPr id="2" name="写真"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3"/>`
  + `<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xf(0, 0, 2000000, 2000000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
const tc = (t) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP"/><a:t>${t}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`;
const table = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="表"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>`
  + `<p:xfrm><a:off x="3000000" y="1500000"/><a:ext cx="6000000" cy="3200000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">`
  + `<a:tbl><a:tblGrid><a:gridCol w="3000000"/><a:gridCol w="3000000"/></a:tblGrid>`
  + Array.from({ length: 8 }, (_, i) => `<a:tr h="400000">${tc(i ? '専門分野' + i : '専門分野')}${tc(i ? '氏名' + i : 'メンバー')}</a:tr>`).join('')
  + `</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
// クリックで始まる動き（一覧が出てくる）。1つめのまとまりはクリック待ち（delay="indefinite"）
const TIMING = '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
  + '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold">'
  + '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" presetID="10" presetClass="entr" fill="hold" nodeType="clickEffect">'
  + '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="5" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
  + '<p:tgtEl><p:spTgt spid="6"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
  + '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
  + '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>';
const page = (transition, timing) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>`
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
  + pic + sp(11, '業種区分') + table + sp(31, '次の方')
  + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' + (transition || '') + (timing || '') + '</p:sld>';
const RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>';
const item = { kind: 'overview', block: '見本の業種区分', nextName: '見本 一郎', photoName: '見本 一郎',
               rows: [{ title: '見本の専門分野', name: '見本 一郎' }, { title: '試験の専門分野', name: '試験 花子' }] };
const advOf = (x) => [...x.matchAll(/<p:transition\b[^>]*\sadvTm="(\d+)"/g)].map((m) => +m[1]);

// ===== 1) 扉ページは1秒で次へ =====
{
  // 画面切り替えが無いテンプレート：切り替えを足して1秒（置き場所は clrMapOvr のあと・timing の前）
  const a = F.mpOverviewSlide_(page('', TIMING), RELS, item, null).xml;
  ck(J(advOf(a)) === J([1000]) && /<\/p:clrMapOvr><p:transition advTm="1000"\/><p:timing>/.test(a),
     '1) 画面切り替えの無い扉ページが1秒で次へ進まない: ' + J(advOf(a)));
  // クリックで始まる動きは、すぐ始まる（クリックを待つと時間で進めない）
  const main = a.slice(a.indexOf('nodeType="mainSeq"'));
  ck(!/<p:cond delay="indefinite"\/>/.test(main.slice(0, main.indexOf('</p:stCondLst>') + 20)), '1) 扉ページの動きがクリック待ちのまま');
  // 中身はこれまでどおり
  ck(a.includes('見本の業種区分') && a.includes('試験 花子') && a.includes('見本の専門分野'), '1) 扉ページの中身: ' + a.length);
  // テンプレートに「○秒後に次へ」があれば1秒に、「切り替えだけ」なら時間を足す
  const b = F.mpOverviewSlide_(page('<p:transition spd="slow" advTm="5000"><p:fade/></p:transition>', ''), RELS, item, null).xml;
  ck(J(advOf(b)) === J([1000]) && /<p:fade\/>/.test(b), '1) テンプレートの「5秒後に次へ」が1秒にならない: ' + J(advOf(b)));
  const c = F.mpOverviewSlide_(page('<p:transition spd="med"><p:push dir="u"/></p:transition>', ''), RELS, item, null).xml;
  ck(J(advOf(c)) === J([1000]) && /<p:push dir="u"\/>/.test(c), '1) 切り替えだけの扉ページに1秒が入らない: ' + J(advOf(c)));
}

// ===== 2) 前半に差し込んだとき：扉ページがあれば「保存済みのタイミングを使用」を入れる合図を返す =====
{
  // テンプレートを開くところ・ページを作るところ・差し込むところは作り物に置き換える（ここでは合図とお知らせだけを見る）
  F.weeklyAnchor_ = () => 'ppt/slides/slide3.xml';
  F.getBigTemplateFile_ = () => ({ getBlob: () => null });
  F.unzipToMap_ = () => ({});
  F.buildMemberPresenSlides_ = () => ({ noPhoto: [], gone: [], opened: 1 });
  F.spliceSlides_ = () => ({ paths: ['ppt/slides/slide10.xml', 'ppt/slides/slide11.xml'], missingLayout: [] });
  F.slideOrder_ = () => [];
  const ind = { kind: 'individual', name: '見本 一郎', auto: false };
  const r1 = F.insertMemberPresen_({}, [item, ind]);
  ck(r1.auto === true && /クリックで次へ・業種区分のページは1秒で次へ/.test(r1.message),
     '2) 個人ページがクリックで次へでも、扉ページの1秒が効く合図（auto）を返さない: ' + J(r1));
  const r2 = F.insertMemberPresen_({}, [Object.assign({}, ind, { auto: true })]);
  ck(r2.auto === true && /自動で次へ）/.test(r2.message), '2) 個人ページの自動で次へ: ' + J(r2));
  const r3 = F.insertMemberPresen_({}, [ind]);
  ck(r3.auto === false && !/1秒/.test(r3.message), '2) 扉ページも自動送りも無いのに合図を返した: ' + J(r3));
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('ウィークリープレゼンの送り方: 検査 ' + checks + ' 件 OK: 業種区分の扉ページは1秒で次へ（切り替えが無いときも・クリック待ちの動きも）・'
  + '前半に差し込んだときの「保存済みのタイミングを使用」');
