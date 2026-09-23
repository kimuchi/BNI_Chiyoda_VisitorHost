// === メンバー名簿マスタ（メンバーブック／メンバープレゼン／Zoom案内の共通データ基盤）===
// 既存の「メンバーリスト」(A=No/B=氏名、OCRで更新) は割り振り表などのキーとして使い続け、
// 詳細情報はこの「メンバー名簿」シートに持つ。氏名で紐付ける。
// 行順＝メンバーブックの掲載順＝メンバープレゼンの発表順。

var MEMBER_SHEET_ = 'メンバー名簿';
// メンバーリスト(OCR)で読むPDFの列（No/氏名/ふりがな/カテゴリー/会社名/役職/メモ）と、
// メンバーブックで使う項目を1枚に統合した。これが全機能の正本になる。
var MEMBER_HEADERS_ = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ',
                       '写真ファイル名', '一言コメント', '紹介してほしい人', '協業したい人',
                       '入会日', '更新日', '更新期限日'];
var COVER_SHEET_ = 'メンバーブック表紙';
var CAT_SHEET_ = '業種区分マスタ';

// 業種区分（キー・表示ラベル・色1・色2・ブロック表示名・巡回順）。
// キーはSpreadingのカテゴリ設定のグループ名と同じにしてある。名簿の「業種区分」列と
// この キー が一致したときに、メンバーブックのカードに色が付く。
// 色1は冊子のカード帯の地色で、実際の配布版PDFから採取した値。
// 巡回順は、Spreadingのグループの並び順に合わせてある。
var DEFAULT_CATEGORIES_ = [
  ['企業サポート',   '企業サポート',   '#FFDE58', '#F5C400', '企業サポート',   1],
  ['研修・教育',     '研修・教育',     '#C1FF72', '#8FD43A', '研修・教育',     2],
  ['不動産関連',     '不動産関連',     '#37B5FF', '#0B8FE0', '不動産関連',     3],
  ['建築・住まい',   '建築・住まい',   '#AAB5D9', '#7A88B8', '建築＆住まい',   4],
  ['プロモーション', 'プロモーション', '#FF66C3', '#E0329B', 'プロモーション', 5],
  ['暮らし・生活',   '暮らし・生活',   '#5CE1E6', '#1FBCC2', '暮らし・生活',   6],
  ['美容と健康',     '美容と健康',     '#7DD957', '#4FAF2A', '美容と健康',     7],
  ['飲食・エンタメ', '飲食・エンタメ', '#CB6BE6', '#A33BC2', '飲食・エンタメ', 8]
];

function openMemberMasterDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('member_master').setWidth(1000).setHeight(720),
    'メンバー名簿の管理');
}

function ensureMemberSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(MEMBER_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(MEMBER_SHEET_);
    sh.appendRow(MEMBER_HEADERS_);
    sh.getRange(1, 1, 1, MEMBER_HEADERS_.length).setFontWeight('bold').setBackground('#f2f6ff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureCategorySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(CAT_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(CAT_SHEET_);
    sh.appendRow(['キー', '表示ラベル', '色1', '色2', 'ブロック表示名', '巡回順']);
    sh.getRange(2, 1, DEFAULT_CATEGORIES_.length, 6).setValues(DEFAULT_CATEGORIES_);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f2f6ff');
    sh.setFrozenRows(1);
    sh.hideSheet();
    return sh;
  }
  // 区分が増えたとき（Spreadingのグループが変わったときなど）に足りない行を補う。
  // 既にある行の色や並びは、手で直されている可能性があるのでそのまま残す。
  var have = {}, data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) have[String(data[i][0]).trim()] = true;
  var add = [];
  for (var j = 0; j < DEFAULT_CATEGORIES_.length; j++) {
    if (!have[DEFAULT_CATEGORIES_[j][0]]) add.push(DEFAULT_CATEGORIES_[j]);
  }
  if (add.length) {
    sh.getRange(sh.getLastRow() + 1, 1, add.length, 6).setValues(add);
    console.log('[CAT] 業種区分を' + add.length + '件追加しました');
  }
  return sh;
}

// 業種区分の色や並びを初期値に戻す（冊子の配色に合わせ直したいとき）
function resetCategoryMaster() {
  try {
    var sh = ensureCategorySheet_();
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
    sh.getRange(2, 1, DEFAULT_CATEGORIES_.length, 6).setValues(DEFAULT_CATEGORIES_);
    return { ok: true, message: '業種区分を初期値に戻しました（' + DEFAULT_CATEGORIES_.length + '区分）。',
             categories: getCategoryMaster() };
  } catch (e) {
    return { ok: false, message: '戻せませんでした: ' + (e && e.message ? e.message : e) };
  }
}

function toDateStr_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  }
  return String(v).trim();
}

function getMemberMaster() {
  try {
    var sh = ensureMemberSheet_(), data = sh.getDataRange().getValues(), members = [];
    var col = function (r, i) { return String(r[i] == null ? '' : r[i]).trim(); };
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!col(r, 2)) continue;                       // 氏名が無い行は飛ばす
      members.push({
        no:        col(r, 0),
        cat:       col(r, 1),
        name:      col(r, 2),
        kana:      col(r, 3),
        title:     col(r, 4),     // カテゴリー（業務内容）
        company:   col(r, 5),
        role:      col(r, 6),     // 役職
        memo:      col(r, 7),
        photoFile: col(r, 8),
        comment:   col(r, 9),
        refer:     col(r, 10),
        collab:    col(r, 11),
        joinDate:   toDateStr_(r[12]),
        renewDate:  toDateStr_(r[13]),
        expireDate: toDateStr_(r[14])
      });
    }
    return { ok: true, members: members, cover: getCoverInfo_(), categories: getCategoryMaster() };
  } catch (e) {
    console.error('[MEMBER] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'メンバー名簿の読み込みに失敗しました: ' + (e && e.message ? e.message : e), members: [] };
  }
}

function saveMemberMaster(members, cover) {
  try {
    if (!members) return { ok: false, message: '保存するデータがありません。' };
    var sh = ensureMemberSheet_();
    sh.clear();
    sh.appendRow(MEMBER_HEADERS_);
    sh.getRange(1, 1, 1, MEMBER_HEADERS_.length).setFontWeight('bold').setBackground('#f2f6ff');
    sh.setFrozenRows(1);
    var rows = [];
    for (var i = 0; i < members.length; i++) {
      var m = members[i] || {};
      rows.push([m.no || '', m.cat || '', m.name || '', m.kana || '', m.title || '',
                 m.company || '', m.role || '', m.memo || '', m.photoFile || '',
                 m.comment || '', m.refer || '', m.collab || '',
                 m.joinDate || '', m.renewDate || '', m.expireDate || '']);
    }
    if (rows.length) sh.getRange(2, 1, rows.length, MEMBER_HEADERS_.length).setValues(rows);
    if (cover) saveCoverInfo_(cover);
    console.log('[MEMBER] saved=' + rows.length);
    return { ok: true, message: 'メンバー名簿を保存しました（' + rows.length + '名）。' };
  } catch (e) {
    console.error('[MEMBER] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}


// 氏名から写真ファイル名を自動解決して写真列を埋める
function autoFillMemberPhotos() {
  try {
    var m = getMemberMaster();
    if (!m.ok) return m;
    var members = m.members, filled = 0, miss = [];
    for (var i = 0; i < members.length; i++) {
      var id = findPhotoIdForName_(members[i].name);
      if (!id) { miss.push(members[i].name); continue; }
      try {
        var nm = DriveApp.getFileById(id).getName();
        if (members[i].photoFile !== nm) { members[i].photoFile = nm; filled++; }
      } catch (e) { miss.push(members[i].name); }
    }
    var res = saveMemberMaster(members, null);
    if (!res.ok) return res;
    return { ok: true, message: filled + '名の写真を紐付けました。未照合 ' + miss.length + '名。', filled: filled, missing: miss };
  } catch (e) {
    return { ok: false, message: '写真の自動照合に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// --- 業種区分マスタ ---
function getCategoryMaster() {
  var sh = ensureCategorySheet_(), data = sh.getDataRange().getValues(), list = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0] || '').trim()) continue;
    list.push({ key: String(data[i][0]).trim(), label: String(data[i][1] || '').trim(),
                bg: String(data[i][2] || '').trim(), bg2: String(data[i][3] || '').trim(),
                block: String(data[i][4] || '').trim(), order: Number(data[i][5]) || 0 });
  }
  list.sort(function (a, b) { return a.order - b.order; });
  return list;
}

function saveCategoryMaster(rows) {
  try {
    var sh = ensureCategorySheet_();
    sh.clear();
    sh.appendRow(['キー', '表示ラベル', '色1', '色2', 'ブロック表示名', '巡回順']);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f2f6ff');
    var vals = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i] || {};
      vals.push([r.key || '', r.label || '', r.bg || '', r.bg2 || '', r.block || '', Number(r.order) || (i + 1)]);
    }
    if (vals.length) sh.getRange(2, 1, vals.length, 6).setValues(vals);
    return { ok: true, message: '業種区分マスタを保存しました。' };
  } catch (e) {
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// --- メンバーブック表紙情報 ---
// === メンバーブックの表紙ページ ===
// 期ごとに書き換わる文章（プレジデント挨拶など）と、めったに変わらない定型文
// （理念・用語の説明など）をまとめて1件のJSONで持つ。
// 項目が増えてもプロパティを増やさずに済む。
var COVER_KEY_ = 'BNI_MB_COVER';

var DEFAULT_COVER_ = {
  title: 'BNI Active chapter Member Book',
  term: '23期',
  philosophyTitle: 'BNIの理念　Givers Gain®（ギバーズゲイン）',
  philosophy: '他の人におしみなくビジネスを提供することで、自分も他の人からビジネスを提供してもらえる、「与えるものは与えられる」という考え方に基づいて運営されています。',
  aboutTitle: 'BNI Activeチャプターとは？',
  about: '2015年9月に発足し、23期を迎えました。\n個性豊かなメンバーが、様々な分野のプロフェッショナルとして、お互いのビジネス発展、売上UP、人脈の拡大のためにサポートし合うビジネスチームです。',
  benefitsTitle: 'BNIに参加するメリット（BNIを活用する５つのベネフィット＋1）',
  benefits: '１.大きなマーケティングチーム　２.競合のいないビジネス環境　３.継続的な新規顧客の紹介\n４.国内外に広がる人脈　５.長く続く有意義な信頼関係　＋１.生涯学習',
  scheduleFrom: '7:15',
  scheduleTo: '9:15',
  schedule: [
    'オープンネットワーキング',
    'ビジターの歓迎と、リーダーシップチーム及びサポートチームの紹介',
    'BNIのコアバリュー、目的と概要',
    'ネットワーキングに関する学習コーナー',
    'BNIネットワーキングリーダーの発表',
    '新メンバー及び更新メンバーの歓迎',
    '全メンバーによるウィークリープレゼンテーション',
    'ビジターを再度歓迎/ビジターによるプレゼンテーション',
    'ビジネスブレイクアウトルーム',
    'バイスプレジデントによる報告',
    'メンバーシップ委員会による報告',
    '書記兼会計によるスピーカーローテーションの発表と、今週のスピーカーの紹介',
    'スピーカーによるメインプレゼンテーション',
    'リファーラルと推薦のことばの発表とビジターから良かった点のシェア',
    'リファーラル真正度の確認',
    '書記兼会計による報告',
    'プレジデントからビジターへの感謝',
    'BNIによる発表、お知らせ、特別レポート',
    '商品抽選（ビジターを同伴、またはリファーラルを提供したメンバーが対象）',
    '閉会'
  ].join('\n'),
  termsTitle: 'BNIの用語の説明',
  terms: 'チャプター\n1つの専門分野に1名で構成されたビジネスチームのことです。\n'
       + 'リファーラル\nメンバー間で交わされるビジネスや、信頼に基づく人脈の紹介です。\n'
       + 'サンキュー\nそのリファーラルによって発生した売上金額のことで、感謝の意を込めて、サンキューと呼んでいます。\n'
       + 'カテゴリー\n専門業種のことで、BNIでは1つの専門分野に対して加入できるのは1名のみであるという規定があります。',
  pname: '',
  prole: 'Activeチャプター\n第23期プレジデント',
  ptext: '',
  photoFile: ''
};

function getCoverInfo_() {
  var props = PropertiesService.getScriptProperties(), cover = {};
  for (var k in DEFAULT_COVER_) cover[k] = DEFAULT_COVER_[k];
  try {
    var raw = props.getProperty(COVER_KEY_);
    if (raw) { var saved = JSON.parse(raw); for (var j in saved) cover[j] = saved[j]; }
  } catch (e) {
    console.warn('[MBOOK] 表紙設定の読み込みに失敗: ' + (e && e.message ? e.message : e));
  }
  // 旧版で個別のプロパティに保存していた分を引き継ぐ
  var old = { term: 'BNI_MB_TERM', pname: 'BNI_MB_PNAME', ptext: 'BNI_MB_PTEXT', photoFile: 'BNI_MB_COVER_PHOTO' };
  for (var o in old) {
    var v = props.getProperty(old[o]);
    if (v && !cover[o]) cover[o] = v;
  }
  return cover;
}

function saveCoverInfo_(cover) {
  if (!cover) return;
  var cur = getCoverInfo_();
  for (var k in cover) if (cover[k] !== undefined) cur[k] = cover[k];
  PropertiesService.getScriptProperties().setProperty(COVER_KEY_, JSON.stringify(cur));
}

// 画面から呼ぶ用
function saveMemberBookCover(cover) {
  try {
    saveCoverInfo_(cover);
    return { ok: true, message: '表紙の設定を保存しました。', cover: getCoverInfo_() };
  } catch (e) {
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 表紙の文章を初期値に戻す（プレジデント関連は消さない）
function resetMemberBookCoverText() {
  try {
    var cur = getCoverInfo_(), keep = ['term', 'pname', 'prole', 'ptext', 'photoFile'];
    var next = {};
    for (var k in DEFAULT_COVER_) next[k] = DEFAULT_COVER_[k];
    for (var i = 0; i < keep.length; i++) next[keep[i]] = cur[keep[i]];
    PropertiesService.getScriptProperties().setProperty(COVER_KEY_, JSON.stringify(next));
    return { ok: true, message: '定型文を初期値に戻しました（プレジデントの設定はそのままです）。', cover: next };
  } catch (e) {
    return { ok: false, message: '戻せませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// === メンバーブックHTML / TSV からの一括取り込み ===

// 現行のメンバーブックHTML（DATA_STORE に JSON が埋まっている）から取り込む
// 取り込んだ内容を名簿へ反映する共通処理。
// 氏名で照合し、**取り込み側に値がある項目だけ**を上書きする。
// 空欄は「消したい」ではなく「その形式に無い／未入力」なので触らない。
// 名簿にしか居ない人も消さない。
// （以前は取り込みのたびに名簿を丸ごと書き換えており、
//   ふりがな・No・写真・入会日など、取り込み元に無い項目が消えていた）
var MEMBER_FIELDS_ = ['no', 'cat', 'kana', 'title', 'company', 'role', 'memo', 'photoFile',
                      'comment', 'refer', 'collab', 'joinDate', 'renewDate', 'expireDate'];

function mergeMembersInto_(incoming) {
  var cur = getMemberMaster().members || [], byName = {}, i, f;
  for (i = 0; i < cur.length; i++) byName[normName_(cur[i].name)] = i;
  var updated = 0, added = 0, addedNames = [];

  for (var k = 0; k < (incoming || []).length; k++) {
    var e = incoming[k] || {};
    if (!e.name) continue;
    var key = normName_(e.name), idx = byName[key];
    if (idx === undefined) {
      var row = { no: '', cat: '', name: e.name, kana: '', title: '', company: '', role: '', memo: '',
                  photoFile: '', comment: '', refer: '', collab: '',
                  joinDate: '', renewDate: '', expireDate: '' };
      for (f = 0; f < MEMBER_FIELDS_.length; f++) {
        if (e[MEMBER_FIELDS_[f]]) row[MEMBER_FIELDS_[f]] = e[MEMBER_FIELDS_[f]];
      }
      cur.push(row); byName[key] = cur.length - 1; added++; addedNames.push(e.name);
      continue;
    }
    var m = cur[idx], hit = 0;
    for (f = 0; f < MEMBER_FIELDS_.length; f++) {
      var fld = MEMBER_FIELDS_[f];
      if (e[fld] && String(e[fld]) !== String(m[fld] || '')) { m[fld] = e[fld]; hit++; }
    }
    if (hit) updated++;
  }
  return { members: cur, updated: updated, added: added, addedNames: addedNames };
}

function importMemberBookHtml(base64) {
  try {
    if (!base64) return { ok: false, message: 'ファイルデータが空です。' };
    var html = Utilities.newBlob(Utilities.base64Decode(base64), 'text/html', 'mb.html').getDataAsString('UTF-8');
    var m = html.match(/<script[^>]*id=["']DATA_STORE["'][^>]*>([\s\S]*?)<\/script>/i);
    if (!m) return { ok: false, message: 'このHTMLに DATA_STORE が見つかりません。メンバーブック管理ツールで保存したHTMLをお選びください。' };
    var json = m[1].trim();
    if (!json) return { ok: false, message: 'DATA_STORE が空でした。メンバーを登録した状態で保存したHTMLをお使いください。' };
    var data = JSON.parse(json);
    var src = data.members || [], members = [];
    for (var i = 0; i < src.length; i++) {
      var s = src[i] || {};
      members.push({
        no: '', cat: s.cat || '', name: s.name || '', kana: '', title: s.title || '',
        company: s.company || '', role: '', memo: '', photoFile: '',
        comment: s.comment || '', refer: s.refer || '', collab: s.collab || '',
        joinDate: '', renewDate: '', expireDate: ''
      });
    }
    if (!members.length) return { ok: false, message: 'メンバーが0件でした。' };

    // 表紙は、HTML側に値が入っている項目だけ反映する（空で上書きしない）
    var cover = null;
    if (data.cover) {
      cover = {};
      if (data.cover.term)  cover.term  = data.cover.term;
      if (data.cover.pname) cover.pname = data.cover.pname;
      if (data.cover.ptext) cover.ptext = data.cover.ptext;
      if (!Object.keys(cover).length) cover = null;
    }

    var merged = mergeMembersInto_(members);
    var res = saveMemberMaster(merged.members, cover);
    if (!res.ok) return res;
    console.log('[MEMBER] html merged updated=' + merged.updated + ' added=' + merged.added);
    var msg = 'メンバーブックHTMLから ' + members.length + '名を読み取りました。'
            + '（更新 ' + merged.updated + '名 / 新規 ' + merged.added + '名 / 名簿は計 ' + merged.members.length + '名）\n'
            + 'HTMLに無い項目（No・ふりがな・役職・メモ・写真ファイル名・入会日・更新日・更新期限日）は'
            + 'そのまま残しています。';
    if (merged.addedNames.length) msg += '\n新しく追加: ' + merged.addedNames.join('、');
    return { ok: true, message: msg, updated: merged.updated, added: merged.added };
  } catch (e) {
    console.error('[MEMBER] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '取り込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// タブ区切りテキストから取り込む（見出し行は任意）
// 列: No / 業種区分 / 氏名 / ふりがな / カテゴリー / 会社名 / 役職 / メモ / 写真 / 一言 / 紹介 / 協業 / 入会日 / 更新日 / 更新期限日
function importMemberTsv(text, replaceAll) {
  try {
    if (!text || !String(text).trim()) return { ok: false, message: '貼り付けられたデータが空です。' };
    var lines = String(text).replace(/\r/g, '').split('\n'), members = [];
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (!line.trim()) continue;
      var c = line.split('\t');
      var name = (c[2] || '').trim();
      if (!name || name === '氏名') continue;       // 見出し行・空行を飛ばす
      members.push({
        no: (c[0] || '').trim(), cat: (c[1] || '').trim(), name: name, kana: (c[3] || '').trim(),
        title: (c[4] || '').trim(), company: (c[5] || '').trim(), role: (c[6] || '').trim(),
        memo: (c[7] || '').trim(), photoFile: (c[8] || '').trim(), comment: (c[9] || '').trim(),
        refer: (c[10] || '').trim(), collab: (c[11] || '').trim(),
        joinDate: (c[12] || '').trim(), renewDate: (c[13] || '').trim(), expireDate: (c[14] || '').trim()
      });
    }
    if (!members.length) return { ok: false, message: '取り込める行がありませんでした。タブ区切りで、3列目が氏名になっているかご確認ください。' };
    var out, note;
    if (replaceAll) {
      out = members;
      note = '名簿を貼り付けた内容だけに置き換えました';
    } else {
      // 空欄で既存の内容を消さないよう、値のある列だけを反映する
      var merged = mergeMembersInto_(members);
      out = merged.members;
      note = '更新 ' + merged.updated + '名 / 新規 ' + merged.added + '名。'
           + '貼り付けた表で空欄だった列は、元の内容を残しています';
    }
    var res = saveMemberMaster(out, null);
    if (!res.ok) return res;
    return { ok: true, message: members.length + '行を取り込みました（' + note + ' / 名簿は計 ' + out.length + '名）。',
             imported: members.length, total: out.length };
  } catch (e) {
    console.error('[MEMBER] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '取り込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// === メンバーリスト(OCR)の読み取り結果をメンバー名簿へ統合する ===
// PDFに載っている項目（No・氏名・ふりがな・業種区分・カテゴリー・会社名・役職・メモ）だけを
// 上書きし、写真・一言コメント・紹介/協業・日付など、PDFに無い項目は既存の値を残す。
function mergeMembersFromOcr_(extracted) {
  try {
    var cur = getMemberMaster().members || [];
    // 既存行を No と氏名の両方から引けるようにする
    var byNo = {}, byName = {};
    for (var i = 0; i < cur.length; i++) {
      if (cur[i].no) byNo[String(cur[i].no)] = i;
      byName[normName_(cur[i].name)] = i;
    }
    var updated = 0, added = 0, used = {};
    for (var k = 0; k < extracted.length; k++) {
      var e = extracted[k];
      var idx = (e.no && byNo[String(e.no)] !== undefined) ? byNo[String(e.no)]
              : (byName[normName_(e.name)] !== undefined ? byName[normName_(e.name)] : -1);
      if (idx >= 0) {
        var m = cur[idx];
        m.no = e.no || m.no;
        m.name = e.name || m.name;
        if (e.kana) m.kana = e.kana;
        if (e.cat) m.cat = e.cat;
        if (e.title) m.title = e.title;
        if (e.company) m.company = e.company;
        if (e.role) m.role = e.role;
        if (e.memo) m.memo = e.memo;
        used[idx] = true; updated++;
      } else {
        cur.push({ no: e.no, cat: e.cat, name: e.name, kana: e.kana, title: e.title,
                   company: e.company, role: e.role, memo: e.memo,
                   photoFile: '', comment: '', refer: '', collab: '',
                   joinDate: '', renewDate: '', expireDate: '' });
        added++;
      }
    }
    // PDFに載っていた順（No順）に並べ替える。掲載順＝発表順の意味を持つため
    cur.sort(function (a, b) {
      var na = parseInt(a.no, 10), nb = parseInt(b.no, 10);
      if (isNaN(na) && isNaN(nb)) return 0;
      if (isNaN(na)) return 1;
      if (isNaN(nb)) return -1;
      return na - nb;
    });
    var res = saveMemberMaster(cur, null);
    if (!res.ok) return res;
    var msg = '読み取り ' + extracted.length + '件を「メンバー名簿」に反映しました。'
            + '（更新 ' + updated + '名 / 新規 ' + added + '名 / 名簿は計 ' + cur.length + '名）\n'
            + '写真・一言コメント・日付など、PDFに無い項目は残しています。';
    console.log('[OCR] merged updated=' + updated + ' added=' + added);
    return { ok: true, message: msg, updated: updated, added: added, total: cur.length };
  } catch (e) {
    console.error('[OCR] merge ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '名簿への反映に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// メモ欄に「ビジターホスト」等が入っているメンバーを、ビジターホスト設定に反映する
function applyVisitorHostsFromMemo() {
  try {
    var members = getMemberMaster().members || [], hosts = [], names = [];
    for (var i = 0; i < members.length; i++) {
      // メモ欄と役職欄の両方を見る。OCR取込はメモ欄に、Spreading取込は役職欄に入るため。
      var memo = (members[i].memo || '') + ' ' + (members[i].role || '');
      // 「ビジターホスト」「ビジホス」などの表記ゆれを拾う
      if (!/ビジ(ター)?ホス(ト)?/.test(memo)) continue;
      if (!members[i].no) continue;
      hosts.push(String(members[i].no)); names.push(members[i].name);
    }
    if (!hosts.length) return { ok: false, message: 'メモ欄に「ビジターホスト」の記載があるメンバーが見つかりませんでした。' };
    saveVisitorHosts(hosts);
    return { ok: true, message: hosts.length + '名をビジターホストに設定しました。\n' + names.join('、'),
             count: hosts.length, names: names };
  } catch (e) {
    return { ok: false, message: '反映に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 旧「メンバーリスト」シートの内容をメンバー名簿へ移し、旧シートを隠す
function migrateMemberListSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), old = ss.getSheetByName('メンバーリスト');
    if (!old) return { ok: false, message: '「メンバーリスト」シートはありません。移行は不要です。' };
    var data = old.getDataRange().getValues(), src = [];
    for (var i = 1; i < data.length; i++) {
      var no = String(data[i][0] == null ? '' : data[i][0]).trim();
      var nm = String(data[i][1] == null ? '' : data[i][1]).trim();
      if (no && nm) src.push({ no: no, name: normalizeSpace(nm), kana: '', cat: '', title: '', company: '', role: '', memo: '' });
    }
    if (!src.length) {
      old.hideSheet();
      return { ok: true, message: '「メンバーリスト」は空だったため、非表示にしました。' };
    }
    var res = mergeMembersFromOcr_(src);
    if (!res.ok) return res;
    old.hideSheet();
    return { ok: true, message: res.message + '\n旧「メンバーリスト」シートは非表示にしました（データは残っています）。' };
  } catch (e) {
    console.error('[MEMBER] migrate ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '移行に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// === BNI公式レポート（Excel）の取り込み ===
// 「メンバーシップ期間レポート」から入会日、「会費レポート」から更新期限日を読む。
// どちらも拡張子は .xls だが中身は SpreadsheetML（XML）なので、そのまま解析できる。

// SpreadsheetML → 行の二次元配列
function parseSpreadsheetMlRows_(xml) {
  var rows = [], rowRe = /<Row[^>]*>([\s\S]*?)<\/Row>/g, m;
  while ((m = rowRe.exec(xml)) !== null) {
    var cells = [], cellRe = /<Cell[^>]*>([\s\S]*?)<\/Cell>|<Cell[^>]*\/>/g, c;
    while ((c = cellRe.exec(m[1])) !== null) {
      var inner = c[1] || '';
      var d = inner.match(/<Data[^>]*>([\s\S]*?)<\/Data>/);
      cells.push(d ? unescapeXml_(d[1].replace(/<[^>]+>/g, '')).trim() : '');
    }
    rows.push(cells);
  }
  return rows;
}

// 'YYYY-MM-DDT00:00:00.000' → 'YYYY/MM/DD'
function isoToDateStr_(v) {
  var m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? (m[1] + '/' + m[2] + '/' + m[3]) : '';
}

function rowHas_(row, text) {
  for (var i = 0; i < row.length; i++) if (String(row[i]).indexOf(text) !== -1) return true;
  return false;
}

// files = [{name, base64}] を1つ以上。レポートの種類は中身の見出しから自動判別する
function importMembershipReports(files) {
  try {
    if (!files || !files.length) return { ok: false, message: 'ファイルが選択されていません。' };
    var join = {}, expire = {}, states = {}, kinds = [];

    for (var f = 0; f < files.length; f++) {
      var xml = Utilities.newBlob(Utilities.base64Decode(files[f].base64), 'text/xml', files[f].name || 'r.xls')
                         .getDataAsString('UTF-8');
      var rows = parseSpreadsheetMlRows_(xml);
      if (!rows.length) continue;

      // --- メンバーシップ期間レポート（姓・名・Recent Start date）---
      var hi = -1;
      for (var i = 0; i < rows.length; i++) if (rowHas_(rows[i], 'Recent Start date')) { hi = i; break; }
      if (hi >= 0) {
        var n1 = 0;
        for (var r = hi + 1; r < rows.length; r++) {
          var row = rows[r];
          if (row.length < 7 || !row[1]) continue;
          var nm = (row[1] + ' ' + row[2]).trim();
          var d = isoToDateStr_(row[6]);          // Recent Start date = 最新の入会日
          if (!nm || !d) continue;
          join[normName_(nm)] = d; n1++;
        }
        kinds.push('メンバーシップ期間レポート（' + n1 + '名）');
        continue;
      }

      // --- 会費レポート（メンバー名・更新日）---
      var hj = -1;
      for (var j = 0; j < rows.length; j++) if (rowHas_(rows[j], 'AutoRenewal')) { hj = j; break; }
      if (hj >= 0) {
        var n2 = 0;
        for (var k = hj + 1; k < rows.length; k++) {
          var rw = rows[k];
          if (rw.length < 6) continue;
          var name2 = String(rw[1] || '').trim();
          var d2 = isoToDateStr_(rw[5]);          // 更新日 = 会費がいつまでか
          if (!name2 || name2 === 'メンバー名' || !d2) continue;
          expire[normName_(name2)] = d2;
          states[normName_(name2)] = String(rw[4] || '').trim();
          n2++;
        }
        kinds.push('会費レポート（' + n2 + '名）');
        continue;
      }
      kinds.push('「' + (files[f].name || '不明') + '」は種類を判別できませんでした');
    }

    if (!Object.keys(join).length && !Object.keys(expire).length) {
      return { ok: false, message: 'レポートの内容を読み取れませんでした。BNI公式サイトから出力した .xls をそのままお選びください。\n' + kinds.join('\n') };
    }

    // 名簿へ反映（氏名で照合。日付以外は触らない）
    var members = getMemberMaster().members || [];
    var setJoin = 0, setExp = 0, pending = [], unmatched = {};
    for (var q in join) unmatched[q] = true;
    for (var q2 in expire) unmatched[q2] = true;

    for (var mi = 0; mi < members.length; mi++) {
      var key = normName_(members[mi].name);
      if (join[key]) { members[mi].joinDate = join[key]; setJoin++; delete unmatched[key]; }
      if (expire[key]) {
        members[mi].expireDate = expire[key]; setExp++; delete unmatched[key];
        if (states[key] && states[key].indexOf('Active') === -1) pending.push(members[mi].name + '（' + states[key] + '）');
      }
    }
    var res = saveMemberMaster(members, null);
    if (!res.ok) return res;

    var left = [];
    for (var u in unmatched) left.push(u);
    var msg = kinds.join(' / ') + ' を読み込みました。\n'
            + '入会日 ' + setJoin + '名 / 更新期限日 ' + setExp + '名 を更新しました。';
    if (left.length) msg += '\n※ 名簿に見つからなかった氏名 ' + left.length + '件: ' + left.slice(0, 10).join('、');
    if (pending.length) msg += '\n※ 更新手続き中の方 ' + pending.length + '名: ' + pending.slice(0, 10).join('、');
    console.log('[REPORT] join=' + setJoin + ' expire=' + setExp + ' unmatched=' + left.length);
    return { ok: true, message: msg, joinSet: setJoin, expireSet: setExp, unmatched: left, pending: pending };
  } catch (e) {
    console.error('[REPORT] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '取り込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}
