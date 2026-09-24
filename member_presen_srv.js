// === メンバープレゼン（ウィークリープレゼンのページ）の生成 ===
//
// もとは「メンバープレゼンスライド 自動生成ツール」という単体のHTMLで、
// メンバーリストHTMLを読ませて pptx を作るものだった。それを名簿システムに取り込み、
// 「メンバー名簿」と「メンバー写真」から直接作れるようにしたもの。
// 出来上がりの見た目は元ツールと同じになるようにしてある。
//
// 単独の画面は持たず、定例会スライド（前半）の画面から作る。作ったページは
// 前半スライドの「ウィークリープレゼンテーション」の見出しの直後に差し込む
// （meeting_slides_srv.js の insertMemberPresen_）。
//
// 【役割を2つに分けている理由】
//  ・会社名の折り返しや文字の縮小には文字幅の実測が要るが、GASには文字を測る手段が無い。
//    そこで組版はブラウザ側（canvasで実測）で行い、決まった結果だけを送ってもらう。
//  ・テンプレートは15MB近くある。画面に送ると重いので、サーバーがDriveから直接読む。
//    → 画面とサーバーの間を行き来するのは小さなJSONだけになる。
//
// テンプレートは「⚙️ 設定 ＞ 大きなスライド」の「メンバープレゼン」に登録する。

var MP_TEMPLATE_KIND_ = 'memberPresen';

// テンプレート内の図形ID（<p:cNvPr id="…">）。
// 元ツールは座標の範囲で図形を探していたが、同じテンプレートを使う前提なので
// IDで指定している。テンプレートが差し替わって見つからないときは、その場で止める。
var MP_OVERVIEW_ = {
  name: '業種区分の扉ページ',
  title: 11,        // 業種区分名
  nextName: 31,     // Next Presenter の氏名
  photo: 2,         // 先頭メンバーの写真
  table: 6,         // 専門分野／メンバーの表
  photoRid: 'rId3'
};
var MP_INDIVIDUAL_ = {
  name: '個人ページ',
  photo: 3, nameBox: 56, company: 2, category: 12,
  nextName: 14,     // NEXT➡ の氏名
  nextLabel: 6,     // NEXT➡ の文字
  photoRid: 'rId2'
};
var MP_ROWS_PER_OVERVIEW_ = 7;     // 一覧の表は7行。8人以上いるとページが増える

// 会社名・カテゴリーの枠（EMU）。元ツールが実物のスライドから採った値。
var MP_COMPANY_DEFAULT_ = { x: 4830858, y: 2849608, cx: 7461517, cy: 769441 };
var MP_COMPANY_TALL_    = { x: 5087424, y: 2776058, cx: 6893161, cy: 1446550 };
var MP_CATEGORY_X_ = 4830481, MP_CATEGORY_CX_ = 7311519;
var MP_CATEGORY_TOP_ = 3456634;        // 会社名が1行のとき
var MP_CATEGORY_TOP_LOW_ = 4071101;    // 会社名が2行のとき（そのぶん下げる）

// 巡回の基準。この週が「プロモーション」始まりで、以降1週ごとに1つずつずれる。
var MP_BASE_DATE_ = '2026/08/19';
var MP_BASE_BLOCK_ = 'プロモーション';

// 業種区分の照合キー。「美容と健康」「美容・健康」のような表記ゆれを吸収する
function mpCatKey_(s) {
  return String(s == null ? '' : s).normalize('NFKC').replace(/[\s・･＆&と]/g, '');
}

// 業種区分マスタ → 巡回の並び。
// マスタには昔の区分や表記ゆれの行が残っていることがあるので、
// 「ブロック表示名が同じもの」は1つにまとめる（画面に同じ名前が2つ並ぶのを防ぐ）。
function mpCycle_() {
  var cats = getCategoryMaster(), seen = {}, out = [];
  for (var i = 0; i < cats.length; i++) {
    var c = cats[i];
    if (!c.key) continue;
    var block = c.block || c.label || c.key, g = mpCatKey_(block);
    if (seen[g]) { seen[g].keys.push(c.key); continue; }
    seen[g] = { gkey: g, block: block, order: c.order || (i + 1), keys: [c.key], count: 0, known: true };
    out.push(seen[g]);
  }
  out.sort(function (a, b) { return a.order - b.order; });
  return out;
}

// 巡回の並びに、名簿の人数を乗せる。
// マスタに無い業種区分の人も落とさず、末尾に足して人数を見せる。
function mpBlocks_(members) {
  var cycle = mpCycle_(), byKey = {}, i, j;
  for (i = 0; i < cycle.length; i++) {
    for (j = 0; j < cycle[i].keys.length; j++) byKey[mpCatKey_(cycle[i].keys[j])] = cycle[i];
  }
  var extra = {};
  for (i = 0; i < members.length; i++) {
    var raw = members[i].cat;
    if (!raw) continue;                                  // 業種区分が空の人は数えない
    var nk = mpCatKey_(raw), hit = byKey[nk];
    if (!hit) {
      hit = extra[nk] || (extra[nk] = { gkey: nk, block: raw, order: 9999, keys: [raw], count: 0, known: false });
      if (cycle.indexOf(hit) < 0) cycle.push(hit);
      byKey[nk] = hit;
    }
    hit.count++;
    members[i].blockKey = hit.gkey;
  }
  return cycle;
}

// 指定の開催日に、どの業種区分から始めるか。基準日からの週数で1つずつずらす。
// 人数が0の区分に当たったときは、次の「人がいる区分」まで進める。
function mpStartFor_(cycle, dateStr) {
  var n = cycle.length;
  if (!n) return '';
  var base = new Date(MP_BASE_DATE_ + ' 00:00:00'), target = new Date(dateStr.replace(/-/g, '/') + ' 00:00:00');
  var weeks = Math.round((target.getTime() - base.getTime()) / (7 * 86400000));
  var baseIdx = 0, i;
  for (i = 0; i < n; i++) if (cycle[i].gkey === mpCatKey_(MP_BASE_BLOCK_)) { baseIdx = i; break; }
  var start = ((baseIdx + weeks) % n + n) % n;
  for (i = 0; i < n; i++) {
    var c = cycle[(start + i) % n];
    if (c.count > 0) return c.gkey;
  }
  return cycle[start].gkey;
}

// --- 始まりの業種区分をルーティンチェックシートから決める ---
// 「ウィークリープレゼン」の欄に「建築　住まい　22番　熊谷さん」のように書いてある。
// チャプターの運用は「開催した回ごとに1つ進み、誰もいない区分は飛ばす」なので、
//   1) その日の記載があれば、その区分
//   2) 無ければ、前回までの記載から、開催した回数ぶん進めた区分
//   3) どちらも無ければ、従来の計算（mpStartFor_）
// の順で決める。

// 書いてある文から区分を探す。書き方が揺れる（「建築 住まい」「美容 健康」）ので、
// 空白や「・」を除いて比べ、いちばん前に出てくる区分を採る。
// 区分名が無ければ、書かれている方の氏名から、その方の区分を引く。
function mpBlockFromText_(cycle, members, raw) {
  var t = mpCatKey_(raw), best = null, bestAt = Infinity, i, k;
  if (!t) return '';
  for (i = 0; i < cycle.length; i++) {
    var names = [cycle[i].block].concat(cycle[i].keys || []);
    for (k = 0; k < names.length; k++) {
      var key = mpCatKey_(names[k]);
      var at = key ? t.indexOf(key) : -1;
      if (at >= 0 && at < bestAt) { best = cycle[i]; bestAt = at; }
    }
  }
  if (best) return best.gkey;
  var who = routineMemberName_(raw);
  for (i = 0; who.name && i < members.length; i++) {
    if (members[i].name === who.name && members[i].blockKey) return members[i].blockKey;
  }
  return '';
}

// fromKey の区分から、人がいる区分だけを数えて steps 回進める。
// fromKey に今は誰もいないときは、その次の「人がいる区分」が1回目になる。
function mpAdvance_(cycle, fromKey, steps) {
  var n = cycle.length, at = -1, live = 0, i;
  for (i = 0; i < n; i++) {
    if (cycle[i].gkey === fromKey) at = i;
    if (cycle[i].count > 0) live++;
  }
  if (at < 0 || !live) return '';
  if (cycle[at].count > 0 && steps <= 0) return fromKey;
  var left = steps <= 0 ? 1 : ((steps - 1) % live) + 1;
  for (i = 1; i <= n * 2; i++) {
    var c = cycle[(at + i) % n];
    if (c.count > 0 && --left === 0) return c.gkey;
  }
  return '';
}

// a の次の週から b まで（b を含む）に、休会日を除いて何回開催したか
function mpMeetingsBetween_(a, b, holidays) {
  var cur = new Date(a.getTime()), n = 0;
  for (var i = 0; i < 520; i++) {
    cur.setDate(cur.getDate() + 7);
    if (cur.getTime() > b.getTime()) break;
    if (holidays.indexOf(fmtDate_(cur)) === -1) n++;
  }
  return n;
}

// rows … routineRowValues_(ROUTINE_WEEKLY_LABELS_) の結果
function mpStartFromRoutine_(cycle, members, rows, dateStr, holidays) {
  var target = parseDate_(dateStr);
  if (!target || !rows) return null;
  var key = fmtDate_(target), k;
  if (rows[key]) {
    k = mpBlockFromText_(cycle, members, rows[key]);
    if (k) return { key: mpAdvance_(cycle, k, 0), from: 'routine', raw: rows[key], date: key, steps: 0 };
  }
  var dates = Object.keys(rows).filter(function (d) { return d < key; }).sort().reverse();
  for (var i = 0; i < dates.length && i < 6; i++) {
    k = mpBlockFromText_(cycle, members, rows[dates[i]]);
    if (!k) continue;
    var steps = mpMeetingsBetween_(parseDate_(dates[i]), target, holidays);
    return { key: mpAdvance_(cycle, k, steps), from: 'previous', raw: rows[dates[i]], date: dates[i], steps: steps };
  }
  return null;
}

// 画面が必要とする情報を一度に返す
function getMemberPresenContext() {
  try {
    var master = getMemberMaster(), members = [];
    for (var i = 0; i < (master.members || []).length; i++) {
      var m = master.members[i];
      members.push({ name: m.name, company: m.company, title: m.title,
                     cat: m.cat, blockKey: '', hasPhoto: !!findPhotoIdForName_(m.name) });
    }
    var cycle = mpBlocks_(members);          // members[].blockKey がここで決まる

    var cands = getMeetingCandidates(), weeklyRows = null, holidays = [];
    try { weeklyRows = routineRowValues_(ROUTINE_WEEKLY_LABELS_); holidays = getHolidays(); } catch (e) {
      console.warn('[MPRESEN] ルーティンチェックシートのウィークリープレゼンを読めませんでした: ' + e.message);
    }
    for (var c = 0; c < cands.length; c++) {
      // 始まりの業種区分（ルーティンチェックシート → 前回の記載から数える → 従来の計算）
      var st = null;
      try { st = mpStartFromRoutine_(cycle, members, weeklyRows, cands[c].dateValue, holidays); } catch (e) {}
      cands[c].start = (st && st.key) || mpStartFor_(cycle, cands[c].dateValue);
      cands[c].startFrom = (st && st.key) ? st.from : 'calc';
      cands[c].startRaw = (st && st.key) ? st.raw : '';
      cands[c].startRawDate = (st && st.key) ? st.date : '';
      cands[c].startSteps = (st && st.key) ? st.steps : 0;
      // 2分30秒プレゼンの方は、ルーティンチェックシートに書いてある
      var ri = null;
      try { ri = getRoutineInfo(cands[c].dateValue); } catch (e) {}
      cands[c].longPresenter = (ri && ri.found) ? ri.longPresenter : '';
      cands[c].longPresenterRaw = (ri && ri.found) ? ri.longPresenterRaw : '';
      cands[c].longUnmatched = !!(ri && ri.longPresenterUnmatched);
      cands[c].routineSheet = (ri && ri.found) ? ri.sheetName : '';
    }

    var tpl = null, st = getBigTemplateStatus();
    for (var t = 0; t < st.templates.length; t++) if (st.templates[t].kind === MP_TEMPLATE_KIND_) tpl = st.templates[t];

    return { ok: true, members: members, blocks: cycle, candidates: cands,
             template: tpl, rowsPerPage: MP_ROWS_PER_OVERVIEW_,
             unusedCategories: unusedCategoryRows_(members).map(function (r) { return r.block; }) };
  } catch (e) {
    console.error('[MPRESEN] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 誰も使っていない業種区分マスタの行。
// 昔の区分が残っていると、順番の一覧に人数0の行が並んで分かりにくくなる。
function unusedCategoryRows_(members) {
  var used = {}, i;
  for (i = 0; i < members.length; i++) if (members[i].cat) used[mpCatKey_(members[i].cat)] = true;
  var cats = getCategoryMaster(), out = [];
  for (i = 0; i < cats.length; i++) {
    if (!cats[i].key) continue;
    if (used[mpCatKey_(cats[i].key)]) continue;
    out.push({ key: cats[i].key, block: cats[i].block || cats[i].label || cats[i].key });
  }
  return out;
}

// 誰も使っていない業種区分をマスタから消す。
// apply が false のときは「何を消すか」を返すだけ（画面で確認してもらうため）。
function tidyCategoryMaster(apply) {
  try {
    var master = getMemberMaster(), members = master.members || [];
    var drop = unusedCategoryRows_(members);
    if (!drop.length) return { ok: true, message: '使われていない業種区分はありません。', removed: [] };
    var names = drop.map(function (r) { return r.block + (r.block === r.key ? '' : '（' + r.key + '）'); });
    if (!apply) {
      return { ok: true, preview: true, removed: names,
               message: '次の ' + drop.length + ' 件を消します:\n' + names.join('\n') };
    }
    var dropKey = {};
    for (var d = 0; d < drop.length; d++) dropKey[drop[d].key] = true;
    var cats = getCategoryMaster(), keep = [];
    for (var i = 0; i < cats.length; i++) if (!dropKey[cats[i].key]) keep.push(cats[i]);
    var res = saveCategoryMaster(keep);
    if (!res.ok) return res;
    console.log('[MPRESEN] 業種区分を整理: ' + names.join(' / '));
    return { ok: true, removed: names,
             message: '✅ 使われていない業種区分を ' + drop.length + ' 件消しました:\n' + names.join('\n') };
  } catch (e) {
    console.error('[MPRESEN] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '整理に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 写真をテンプレートの中に入れて、そのパスと大きさを返す。
// 同じ人は扉ページと個人ページの2回出てくるので、1枚だけ入れて使い回す。
function mpAddPhoto_(map, cache, name) {
  var key = normName_(name);
  if (key in cache.by) return cache.by[key];
  var id = key ? findPhotoIdForName_(name) : '';
  if (!id) { cache.by[key] = null; return null; }
  var blob = DriveApp.getFileById(id).getBlob();
  var ext = String(blob.getContentType() || '').toLowerCase().indexOf('png') >= 0 ? 'png' : 'jpeg';

  // 写真の縦横比は、切り抜き（srcRect）を決めるのに要る。
  // 大きさはファイル先頭のヘッダーだけで分かるが、それを読むには一度バイト列に
  // 展開することになる。何十人ぶんも抱えるとメモリが足りなくなるので、
  //  ・読み終えたバイト列はすぐ捨て、zipには元のBlobをそのまま入れる
  //  ・一度測った大きさは覚えておき、次からはバイト列を触らない
  var size = null, ck = 'mpsize_' + id, cs = null;
  try { cs = CacheService.getScriptCache(); } catch (e) {}
  if (cs) { try { var hit = cs.get(ck); if (hit) size = JSON.parse(hit); } catch (e) {} }
  if (!size) {
    try { size = imageSizeOf_(blob.getBytes()); } catch (e) {}
    if (size && cs) { try { cs.put(ck, JSON.stringify(size), 21600); } catch (e) {} }
  }
  // 画像の名前は、すでにある画像と重ならないものにする。
  // 重なると後から入れた写真で上書きされ、別の方のページの写真が入れ替わってしまう。
  cache.seq++;
  var n = cache.seq, path = 'ppt/media/mpphoto' + n + '.' + ext;
  while (map[path]) path = 'ppt/media/mpphoto' + (++n) + '.' + ext;
  cache.seq = n;
  map[path] = blob.setName(path);
  cache.by[key] = { path: path, width: size ? size.width : 0, height: size ? size.height : 0 };
  return cache.by[key];
}

// 関係ファイルから1件消す
function mpRemoveRel_(relsXml, rId) {
  return relsXml.replace(new RegExp('<Relationship\\b[^>]*\\bId="' + rId + '"[^>]*/>'), '');
}

// 写真を差し替える（無い人は写真の枠ごと消す）
function mpApplyPhoto_(xml, rels, def, photo) {
  if (!photo) return { xml: removeShape_(xml, def.photo), rels: mpRemoveRel_(rels, def.photoRid) };
  var box = readShapeGeomEmu_(xml, def.photo);
  if (box && photo.width && photo.height) {
    xml = setSrcRectInPic_(xml, def.photo, coverCrop_(photo.width, photo.height, box.cx, box.cy));
  }
  return { xml: xml, rels: retargetRel_(rels, def.photoRid, '../media/' + photo.path.replace('ppt/media/', '')) };
}

// 業種区分の扉ページ（一覧の表つき）
function mpOverviewSlide_(tplXml, tplRels, item, photo) {
  var xml = tplXml;
  xml = setParagraphsInShape_(xml, MP_OVERVIEW_.title, [item.block]);
  xml = setParagraphsInShape_(xml, MP_OVERVIEW_.nextName, [item.nextName || '']);
  for (var r = 0; r < MP_ROWS_PER_OVERVIEW_; r++) {
    var row = (item.rows || [])[r] || { title: '', name: '' };
    xml = setTableCellText_(xml, MP_OVERVIEW_.table, r + 1, 0, row.title || '');
    xml = setTableCellText_(xml, MP_OVERVIEW_.table, r + 1, 1, row.name || '');
  }
  return mpApplyPhoto_(xml, tplRels, MP_OVERVIEW_, photo);
}

// 個人ページ
function mpIndividualSlide_(tplXml, tplRels, item, photo) {
  var xml = tplXml;
  xml = setParagraphsInShape_(xml, MP_INDIVIDUAL_.nameBox, [item.name || '']);

  // 会社名：行の分け方・文字の大きさ・枠の大きさは、すべて画面側で決めてある
  // （文字幅を実測できるのは画面側だけなので、寸法もそこで出した方が確実）。
  xml = setParagraphsInShape_(xml, MP_INDIVIDUAL_.company, item.companyLines || ['']);
  xml = setShapeGeomEmu_(xml, MP_INDIVIDUAL_.company,
    item.companyGeom || (item.companyTall ? MP_COMPANY_TALL_ : MP_COMPANY_DEFAULT_));
  if (item.companyPt) xml = setFontSizeInShape_(xml, MP_INDIVIDUAL_.company, item.companyPt);

  // カテゴリー：会社名の枠の真下に置く。
  // 会社名が3行になると、2行ぶんの決め打ちでは文字が重なってしまうため、
  // 画面側が実際の行数と文字の大きさから出した位置を使う。
  xml = setShapeGeomEmu_(xml, MP_INDIVIDUAL_.category,
    { x: MP_CATEGORY_X_, cx: MP_CATEGORY_CX_,
      y: item.categoryTop || (item.companyTall ? MP_CATEGORY_TOP_LOW_ : MP_CATEGORY_TOP_) });
  xml = setBodyAnchorInShape_(xml, MP_INDIVIDUAL_.category, 't');
  xml = noAutofitInShape_(xml, MP_INDIVIDUAL_.category);
  xml = setParagraphsInShape_(xml, MP_INDIVIDUAL_.category, item.categoryLines || ['']);
  if (item.categoryPt) xml = setFontSizeInShape_(xml, MP_INDIVIDUAL_.category, item.categoryPt);
  if (item.categoryTight) xml = setLineSpacingInShape_(xml, MP_INDIVIDUAL_.category, 85);

  // 30秒以外にしたい方（2分30秒プレゼンなど）は、カウントダウンごと作り直す
  if (item.countdownSec) xml = mpSetCountdown_(xml, item.countdownSec);
  if (item.autoAdvanceMs) xml = mpAutoAdvance_(xml, item.autoAdvanceMs);

  if (item.nextName) {
    xml = setParagraphsInShape_(xml, MP_INDIVIDUAL_.nextName, [item.nextName]);
  } else {
    xml = removeShape_(xml, MP_INDIVIDUAL_.nextName);
    xml = removeShape_(xml, MP_INDIVIDUAL_.nextLabel);
  }
  return mpApplyPhoto_(xml, tplRels, MP_INDIVIDUAL_, photo);
}

// === カウントダウンと自動送り ===
//
// テンプレートのカウントダウンは、数字を書いた白い箱を重ねておき、
// 1秒ごとに上から1枚ずつ消して下の数字を見せる、という仕掛け。
//
// 自動で次のスライドへ進ませるには、次の3つが揃っている必要がある。
//   (1) カウントダウンがスライド表示と同時に始まること  … <p:cond delay="0"/>
//   (2) スライドに「○秒で次へ」が入っていること         … <p:transition advTm="…">
//   (3) スライドショーが保存済みのタイミングを使う設定   … <p:showPr useTimings="1">
// 元のテンプレートは3つとも噛み合っておらず（クリック待ち・0.4秒・タイミング無視）、
// とくに(3)が0のままだと(1)(2)を直しても絶対に自動で進まない。

// 「スライドショーの設定 ＞ 保存済みのタイミングを使用」を有効にする。
// プレゼンテーション全体の設定なので、スライドごとではなくここで1回だけ。
function mpUseTimingsOnly_(map, ours) {
  var mine = {}, i, p, n = 0;
  for (i = 0; i < (ours || []).length; i++) mine[ours[i]] = true;
  for (p in map) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(p) || mine[p]) continue;
    var xml = xmlOf_(map, p);
    if (!xml || xml.indexOf('advTm="') < 0) continue;
    putXml_(map, p, mpNoAutoAdvance_(xml));
    n++;
  }
  mpUseTimings_(map);
  return n;
}

function mpUseTimings_(map) {
  var path = 'ppt/presProps.xml', xml = xmlOf_(map, path);
  if (!xml) return false;
  if (/<p:showPr\b[^>]*\buseTimings="1"/.test(xml)) return true;
  if (/<p:showPr\b[^>]*\buseTimings="0"/.test(xml)) {
    putXml_(map, path, xml.replace(/(<p:showPr\b[^>]*\buseTimings=")0(")/, '$11$2'));
    return true;
  }
  if (/<p:showPr\b/.test(xml)) {
    putXml_(map, path, xml.replace(/<p:showPr\b/, '<p:showPr useTimings="1"'));
    return true;
  }
  return false;
}

// 自動送りを外す（クリックで次へ進む）。
// 「保存済みのタイミングを使用」がファイル全体で有効になっていても、このページは止まる。
function mpNoAutoAdvance_(xml) {
  return xml.replace(/\sadvTm="\d+"/g, '');
}

function mpAutoAdvance_(xml, ms) {
  // 開始条件（クリック待ち → すぐ開始）。dur="indefinite" には触らないこと。
  var before = xml;
  xml = xml.replace(/<p:cond delay="indefinite"\/><p:cond evt="onBegin" delay="0"><p:tn val="\d+"\/><\/p:cond>/,
                    '<p:cond delay="0"/>');
  if (xml === before) xml = xml.replace('<p:cond delay="indefinite"/>', '<p:cond delay="0"/>');

  // 自動で次へ進む時間
  if (/advTm="\d+"/.test(xml)) {
    xml = xml.replace(/advTm="\d+"/g, 'advTm="' + ms + '"');
  } else if (/<p:transition\b/.test(xml)) {
    xml = xml.replace(/<p:transition\b([^>]*?)(\/?)>/g, '<p:transition$1 advTm="' + ms + '"$2>');
  } else {
    // 画面切り替えの無いページには、自動で進むだけの切り替えを足す。
    // 置き場所は決まっていて、<p:clrMapOvr>（無ければ <p:cSld>）の直後、<p:timing> より前。
    var tr = '<p:transition advTm="' + ms + '"/>';
    if (xml.indexOf('</p:clrMapOvr>') >= 0) xml = xml.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + tr);
    else xml = xml.replace('</p:cSld>', '</p:cSld>' + tr);
  }
  return xml;
}

// --- 長さの違うカウントダウンを作る（2分30秒プレゼンなど）---
// テンプレートに入っているのは30秒ぶんだけなので、秒数を変えるときは
// 数字の箱とアニメーションを作り直す。書式は見本の箱からそのまま引き継ぐ。

// 表示する文字。60秒以上なら「2:30」のような分:秒にする。
function mpCountdownLabels_(sec) {
  var out = [];
  for (var t = sec; t >= 0; t--) {
    out.push(sec >= 60 ? (Math.floor(t / 60) + ':' + (t % 60 < 10 ? '0' : '') + (t % 60)) : String(t));
  }
  return out;
}

// Arial Black での文字の幅（全角を1とした目安）。箱に収まる大きさを決めるのに使う。
function mpTextEm_(s) {
  var w = 0;
  for (var i = 0; i < s.length; i++) w += (s.charAt(i) === ':') ? 0.3335 : 0.6665;
  return w;
}

// スライド上の「カウントダウンの数字の箱」を集める。
// アニメーションの対象になっている図形と、それと同じ大きさの図形（最後に残る0）が対象。
function mpCountdownShapes_(xml) {
  var anim = {}, m, re = /<p:spTgt spid="(\d+)"\/>/g;
  while ((m = re.exec(xml)) !== null) anim[m[1]] = true;
  var ranges = findTagRanges_(xml, 'p:sp'), all = [], sizes = {}, i;
  for (i = 0; i < ranges.length; i++) {
    var seg = xml.substring(ranges[i].start, ranges[i].end);
    var id = (seg.match(/<p:cNvPr[^>]*\sid="(\d+)"/) || [])[1];
    var t = (seg.match(/<a:t>([^<]*)<\/a:t>/) || [])[1];
    var ext = seg.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
    if (!id || t === undefined || !/^\d+$/.test(t) || !ext) continue;
    var key = ext[1] + 'x' + ext[2];
    all.push({ id: id, num: parseInt(t, 10), start: ranges[i].start, end: ranges[i].end,
               xml: seg, key: key, cx: parseInt(ext[1], 10), animated: !!anim[id] });
    if (anim[id]) sizes[key] = true;
  }
  var out = [];
  for (i = 0; i < all.length; i++) if (all[i].animated || sizes[all[i].key]) out.push(all[i]);
  return out;
}

// 数字の箱を1枚作る。見本の書式をそのまま使い、文字・色・大きさだけ変える。
function mpNumberBox_(model, id, text, color, sizePt) {
  return model
    .replace(/<p:cNvPr id="\d+" name="[^"]*"/, '<p:cNvPr id="' + id + '" name="Count ' + id + '"')
    .replace(/<a:extLst>[\s\S]*?<\/a:extLst>/, '')                 // 図形固有の識別子は引き継がない
    .replace(/<a:srgbClr val="[0-9A-Fa-f]{6}"\/>/, '<a:srgbClr val="' + color + '"/>')
    .replace(/(<a:rPr\b[^>]*?)\ssz="\d+"/, '$1 sz="' + Math.round(sizePt * 100) + '"')
    .replace(/<a:t>[^<]*<\/a:t>/, '<a:t>' + escapeXml_(text) + '</a:t>');
}

// 1秒ごとに1枚ずつ消していくアニメーション。テンプレートと同じ形で組み立てる。
// clickStart が true なら、カウントダウンはクリックで始まる（元のテンプレートと同じ動き）。
function mpCountdownTiming_(spids, clickStart) {
  var steps = '', id = 4, i;
  for (i = 0; i < spids.length; i++) {
    var a = id++, b = id++, c = id++;
    steps += '<p:par><p:cTn id="' + a + '" fill="hold"><p:stCondLst><p:cond delay="' + (i * 1000)
           + '"/></p:stCondLst><p:childTnLst>'
           + '<p:par><p:cTn id="' + b + '" presetID="1" presetClass="exit" presetSubtype="0"'
           + ' fill="hold" grpId="0" nodeType="afterEffect">'
           + '<p:stCondLst><p:cond delay="1000"/></p:stCondLst><p:childTnLst>'
           + '<p:set><p:cBhvr><p:cTn id="' + c + '" dur="1" fill="hold">'
           + '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
           + '<p:tgtEl><p:spTgt spid="' + spids[i] + '"/></p:tgtEl>'
           + '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
           + '<p:to><p:strVal val="hidden"/></p:to></p:set>'
           + '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>';
  }
  var bld = '';
  for (i = 0; i < spids.length; i++) bld += '<p:bldP spid="' + spids[i] + '" grpId="0" animBg="1"/>';
  return '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
       + '<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
       + '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
       + '<p:par><p:cTn id="3" fill="hold"><p:stCondLst>'
       + (clickStart ? '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>'
                     : '<p:cond delay="0"/>')
       + '</p:stCondLst>'
       + '<p:childTnLst>' + steps + '</p:childTnLst></p:cTn></p:par>'
       + '</p:childTnLst></p:cTn>'
       + '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
       + '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
       + '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
       + '<p:bldLst>' + bld + '</p:bldLst></p:timing>';
}

// カウントダウンを指定の秒数で作り直す
function mpSetCountdown_(xml, sec, clickStart) {
  var boxes = mpCountdownShapes_(xml);
  if (boxes.length < 5) {
    throw new Error('テンプレートにカウントダウンの数字が見つかりません（' + boxes.length + '枚）。');
  }
  boxes.sort(function (a, b) { return a.num - b.num; });
  var model = boxes[boxes.length - 1], base = model.xml;

  var labels = mpCountdownLabels_(sec), n = labels.length;
  // 文字数が増えるぶん、箱に収まる大きさまで小さくする（「2:30」は4文字）
  var pt = 96, m = base.match(/<a:rPr\b[^>]*?\ssz="(\d+)"/);
  if (m) pt = parseInt(m[1], 10) / 100;
  var usable = (model.cx - 91440 * 2) / 12700;                 // 左右の余白を引いた幅(pt)
  var fit = Math.floor(usable / mpTextEm_(labels[0]) * 0.95);
  if (fit < pt) pt = Math.max(fit, 24);

  // 色は見本から拾う。すべて同じ色のテンプレート（リファーラル発表など）では
  // その色をそのまま使い、30秒のように色分けされていれば残り10秒から赤にする。
  var colors = {}, ci;
  for (ci = 0; ci < boxes.length; ci++) {
    var cm = boxes[ci].xml.match(/<a:srgbClr val="([0-9A-Fa-f]{6})"\/>/);
    if (cm) colors[cm[1].toUpperCase()] = true;
  }
  var only = null, cn = 0;
  for (var ck in colors) { only = ck; cn++; }
  var COL_NEAR = (cn === 1) ? only : 'CF2030', COL_FAR = (cn === 1) ? only : '64666A';

  var maxId = 0, mm, reId = /<p:cNvPr[^>]*\sid="(\d+)"/g;
  while ((mm = reId.exec(xml)) !== null) maxId = Math.max(maxId, parseInt(mm[1], 10));

  // 文書順は「残り0秒」が一番奥、「残り最大」が一番手前。手前から順に消していく。
  var ids = [], shapes = '';
  for (var d = 0; d < n; d++) {
    var id = ++maxId;
    ids[d] = id;
    shapes += mpNumberBox_(base, id, labels[n - 1 - d], d <= 10 ? COL_NEAR : COL_FAR, pt);
  }
  var spids = [];
  for (var r = sec; r >= 1; r--) spids.push(ids[r]);

  boxes.sort(function (a, b) { return a.start - b.start; });
  var at = boxes[0].start, out = xml;
  for (var i = boxes.length - 1; i >= 0; i--) out = out.substring(0, boxes[i].start) + out.substring(boxes[i].end);
  out = out.substring(0, at) + shapes + out.substring(at);

  var timing = mpCountdownTiming_(spids, clickStart);
  if (/<p:timing>/.test(out)) out = out.replace(/<p:timing>[\s\S]*?<\/p:timing>/, timing);
  else out = out.replace('</p:sld>', timing + '</p:sld>');
  return out;
}

// テンプレートの中身を、指示どおりのスライドの並びに置き換える
// ひな形にする2枚を、スライド番号ではなく「載っている図形」で見つける。
// テンプレートを作り直すと何枚目かが変わるため、番号を決め打ちにしない。
function mpFindModelSlide_(map, ids) {
  var nums = [], p;
  for (p in map) {
    var m = p.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (m) nums.push(parseInt(m[1], 10));
  }
  nums.sort(function (a, b) { return a - b; });
  for (var i = 0; i < nums.length; i++) {
    var path = 'ppt/slides/slide' + nums[i] + '.xml', xml = xmlOf_(map, path);
    if (!xml || missingShapeIds_(xml, ids).length) continue;
    var rels = xmlOf_(map, 'ppt/slides/_rels/slide' + nums[i] + '.xml.rels');
    if (rels) return { path: path, xml: xml, rels: rels };
  }
  return null;
}

function buildMemberPresenSlides_(map, items) {
  // 個人ページから先に探す。扉ページとIDが一部重なるため、
  // 個人ページにしか無いID（氏名・カテゴリー・NEXT）を先に当てる。
  var iv = mpFindModelSlide_(map, [MP_INDIVIDUAL_.photo, MP_INDIVIDUAL_.nameBox, MP_INDIVIDUAL_.company,
                                   MP_INDIVIDUAL_.category, MP_INDIVIDUAL_.nextName, MP_INDIVIDUAL_.nextLabel]);
  var ov = mpFindModelSlide_(map, [MP_OVERVIEW_.title, MP_OVERVIEW_.nextName,
                                   MP_OVERVIEW_.photo, MP_OVERVIEW_.table]);
  var lack = [];
  if (!ov) lack.push(MP_OVERVIEW_.name);
  if (!iv) lack.push(MP_INDIVIDUAL_.name);
  if (lack.length) {
    throw new Error('テンプレートに「' + lack.join('」「') + '」のひな形が見つかりません。'
      + '「⚙️ 設定 ＞ 大きなスライド」に登録したファイルをご確認ください。'
      + '（図形の並びが変わると見つけられなくなります）');
  }
  var ovXml = ov.xml, ovRels = ov.rels, ivXml = iv.xml, ivRels = iv.rels;

  // 元のスライドとノートを全部外す（ひな形2枚の内容はもう読み終えている）
  var paths = [], p;
  for (p in map) paths.push(p);
  for (var d = 0; d < paths.length; d++) {
    if (/^ppt\/slides\/(_rels\/)?slide\d+\.xml(\.rels)?$/.test(paths[d])
     || /^ppt\/notesSlides\//.test(paths[d])) delete map[paths[d]];
  }

  var prs = xmlOf_(map, 'ppt/presentation.xml');
  var prsRels = xmlOf_(map, 'ppt/_rels/presentation.xml.rels');
  var ct = xmlOf_(map, '[Content_Types].xml');

  // スライドの関係だけ外し、残った中で一番大きい rId の続きから振る
  prsRels = prsRels.replace(/<Relationship\b[^>]*relationships\/slide"[^>]*\/>/g, '');
  var maxRid = 0, m, reR = /Id="rId(\d+)"/g;
  while ((m = reR.exec(prsRels)) !== null) maxRid = Math.max(maxRid, parseInt(m[1], 10));

  var timings = mpUseTimings_(map);      // 自動送りが効くようにしておく

  var cache = { by: {}, seq: 0 }, sldIds = '', relAdd = '', ctAdd = '', noPhoto = [];
  for (var i = 0; i < items.length; i++) {
    var it = items[i], n = i + 1, rid = 'rId' + (maxRid + n);
    var photo = mpAddPhoto_(map, cache, it.photoName);
    if (!photo && it.photoName && noPhoto.indexOf(it.photoName) === -1) noPhoto.push(it.photoName);

    var built = (it.kind === 'overview')
      ? mpOverviewSlide_(ovXml, ovRels, it, photo)
      : mpIndividualSlide_(ivXml, ivRels, it, photo);

    putXml_(map, 'ppt/slides/slide' + n + '.xml', built.xml);
    putXml_(map, 'ppt/slides/_rels/slide' + n + '.xml.rels', built.rels);
    sldIds += '<p:sldId id="' + (255 + n) + '" r:id="' + rid + '"/>';
    relAdd += '<Relationship Id="' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"'
            + ' Target="slides/slide' + n + '.xml"/>';
    ctAdd  += '<Override PartName="/ppt/slides/slide' + n + '.xml"'
            + ' ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
  }

  prs = prs.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst>' + sldIds + '</p:sldIdLst>');
  prsRels = prsRels.replace('</Relationships>', relAdd + '</Relationships>');
  ct = ct.replace(/<Override\b[^>]*PartName="\/ppt\/(slides\/slide|notesSlides\/notesSlide)\d+\.xml"[^>]*\/>/g, '');
  ct = ct.replace('</Types>', ctAdd + '</Types>');
  putXml_(map, 'ppt/presentation.xml', prs);
  putXml_(map, 'ppt/_rels/presentation.xml.rels', prsRels);
  putXml_(map, '[Content_Types].xml', ct);

  var pruned = mpPruneMedia_(map);
  return { slides: items.length, photos: cache.seq, noPhoto: noPhoto,
           prunedMedia: pruned, useTimings: timings };
}

// どこからも使われなくなった画像を捨てる。
// 元の61枚ぶんの写真が残ったままだと、出来上がりが無駄に十数MBになる。
function mpPruneMedia_(map) {
  var used = {}, p, m;
  for (p in map) {
    if (!/\.rels$/.test(p)) continue;
    var re = /Target="[^"]*media\/([^"\/]+)"/g, xml = xmlOf_(map, p);
    while ((m = re.exec(xml)) !== null) used[m[1]] = true;
  }
  var drop = [];
  for (p in map) {
    if (p.indexOf('ppt/media/') !== 0) continue;
    if (!used[p.substring('ppt/media/'.length)]) drop.push(p);
  }
  for (var i = 0; i < drop.length; i++) delete map[drop[i]];
  return drop.length;
}
