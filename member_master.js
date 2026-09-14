// === メンバー名簿マスタ（メンバーブック／メンバープレゼン／Zoom案内の共通データ基盤）===
// 既存の「メンバーリスト」(A=No/B=氏名、OCRで更新) は割り振り表などのキーとして使い続け、
// 詳細情報はこの「メンバー名簿」シートに持つ。氏名で紐付ける。
// 行順＝メンバーブックの掲載順＝メンバープレゼンの発表順。

var MEMBER_SHEET_ = 'メンバー名簿';
var MEMBER_HEADERS_ = ['業種区分', '会社名', 'カテゴリー', '氏名', '写真ファイル名',
                       '一言コメント', '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日'];
var COVER_SHEET_ = 'メンバーブック表紙';
var CAT_SHEET_ = '業種区分マスタ';

// 現行ツールの8業種区分（キー・表示ラベル・色・ブロック表示名・巡回順）
var DEFAULT_CATEGORIES_ = [
  ['企業サポート', '企業サポート', '#1f4e79', '#2e75b6', '企業サポート', 1],
  ['研修教育',     '研修・教育',   '#375623', '#548235', '研修・教育',   2],
  ['建築住まい',   '建築・住まい', '#7f6000', '#bf9000', '建築＆住まい', 3],
  ['プロモーション', 'プロモーション', '#833c00', '#c55a11', 'プロモーション', 4],
  ['美容健康',     '美容・健康',   '#7b2d52', '#c0507f', '美容・健康',   5],
  ['金融保険',     '金融・保険',   '#1f3864', '#2f5597', '金融・保険',   6],
  ['不動産',       '不動産',       '#4d3b63', '#7030a0', '不動産',       7],
  ['暮らしサービス', '暮らしサービス', '#255e5e', '#38859c', '暮らしサービス', 8]
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
  }
  return sh;
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
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!String(r[3] == null ? '' : r[3]).trim()) continue;   // 氏名が無い行は飛ばす
      members.push({
        cat:       String(r[0] == null ? '' : r[0]).trim(),
        company:   String(r[1] == null ? '' : r[1]).trim(),
        title:     String(r[2] == null ? '' : r[2]).trim(),
        name:      String(r[3] == null ? '' : r[3]).trim(),
        photoFile: String(r[4] == null ? '' : r[4]).trim(),
        comment:   String(r[5] == null ? '' : r[5]).trim(),
        refer:     String(r[6] == null ? '' : r[6]).trim(),
        collab:    String(r[7] == null ? '' : r[7]).trim(),
        joinDate:   toDateStr_(r[8]),
        renewDate:  toDateStr_(r[9]),
        expireDate: toDateStr_(r[10])
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
      rows.push([m.cat || '', m.company || '', m.title || '', m.name || '', m.photoFile || '',
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

// 既存「メンバーリスト」から氏名をシードする（重複は作らない）
function seedMemberMaster() {
  try {
    var src = getMembersList();            // 既存関数 [{no,name}]
    if (!src.length) return { ok: false, message: '「メンバーリスト」シートにデータがありません。先にメンバーリスト(OCR)の更新を行ってください。' };
    var cur = getMemberMaster().members || [], have = {};
    for (var i = 0; i < cur.length; i++) have[normName_(cur[i].name)] = true;
    var added = 0;
    for (var j = 0; j < src.length; j++) {
      var nm = src[j].name;
      if (!nm || have[normName_(nm)]) continue;
      cur.push({ cat: '', company: '', title: '', name: nm, photoFile: '', comment: '', refer: '', collab: '', joinDate: '', renewDate: '', expireDate: '' });
      have[normName_(nm)] = true; added++;
    }
    var res = saveMemberMaster(cur, null);
    if (!res.ok) return res;
    return { ok: true, message: added + '名を追加しました（既存 ' + (cur.length - added) + '名はそのまま）。', added: added, total: cur.length };
  } catch (e) {
    console.error('[MEMBER] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'シードに失敗しました: ' + (e && e.message ? e.message : e) };
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
function getCoverInfo_() {
  var props = PropertiesService.getScriptProperties();
  return {
    term:  props.getProperty('BNI_MB_TERM') || '',
    pname: props.getProperty('BNI_MB_PNAME') || '',
    ptext: props.getProperty('BNI_MB_PTEXT') || '',
    photoFile: props.getProperty('BNI_MB_COVER_PHOTO') || ''
  };
}
function saveCoverInfo_(cover) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('BNI_MB_TERM', String(cover.term || ''));
  props.setProperty('BNI_MB_PNAME', String(cover.pname || ''));
  props.setProperty('BNI_MB_PTEXT', String(cover.ptext || ''));
  if (cover.photoFile !== undefined) props.setProperty('BNI_MB_COVER_PHOTO', String(cover.photoFile || ''));
}

// === メンバーブックHTML / TSV からの一括取り込み ===

// 現行のメンバーブックHTML（DATA_STORE に JSON が埋まっている）から取り込む
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
        cat: s.cat || '', company: s.company || '', title: s.title || '', name: s.name || '',
        photoFile: '', comment: s.comment || '', refer: s.refer || '', collab: s.collab || '',
        joinDate: '', renewDate: '', expireDate: ''
      });
    }
    if (!members.length) return { ok: false, message: 'メンバーが0件でした。' };
    var cover = data.cover ? { term: data.cover.term || '', pname: data.cover.pname || '', ptext: data.cover.ptext || '' } : null;
    var res = saveMemberMaster(members, cover);
    if (!res.ok) return res;
    console.log('[MEMBER] imported from html: ' + members.length);
    return { ok: true, message: 'メンバーブックHTMLから ' + members.length + '名を取り込みました。写真は「⚙️ メンバー写真の管理」で紐付けてください。' };
  } catch (e) {
    console.error('[MEMBER] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '取り込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// タブ区切りテキストから取り込む（見出し行は任意）
// 列: 業種区分 / 会社名 / カテゴリー / 氏名 / 写真ファイル名 / 一言 / 紹介してほしい人 / 協業したい人
function importMemberTsv(text, replaceAll) {
  try {
    if (!text || !String(text).trim()) return { ok: false, message: '貼り付けられたデータが空です。' };
    var lines = String(text).replace(/\r/g, '').split('\n'), members = [];
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (!line.trim()) continue;
      var c = line.split('\t');
      var name = (c[3] || '').trim();
      if (!name || name === '氏名') continue;       // 見出し行・空行を飛ばす
      members.push({
        cat: (c[0] || '').trim(), company: (c[1] || '').trim(), title: (c[2] || '').trim(),
        name: name, photoFile: (c[4] || '').trim(), comment: (c[5] || '').trim(),
        refer: (c[6] || '').trim(), collab: (c[7] || '').trim(),
        joinDate: (c[8] || '').trim(), renewDate: (c[9] || '').trim(), expireDate: (c[10] || '').trim()
      });
    }
    if (!members.length) return { ok: false, message: '取り込める行がありませんでした。タブ区切りで、4列目が氏名になっているかご確認ください。' };
    var out = members;
    if (!replaceAll) {
      var cur = getMemberMaster().members || [], have = {};
      for (var k = 0; k < cur.length; k++) have[normName_(cur[k].name)] = k;
      for (var j = 0; j < members.length; j++) {
        var key = normName_(members[j].name);
        if (have[key] !== undefined) cur[have[key]] = members[j];   // 同じ氏名は上書き
        else cur.push(members[j]);
      }
      out = cur;
    }
    var res = saveMemberMaster(out, null);
    if (!res.ok) return res;
    return { ok: true, message: members.length + '行を取り込みました（名簿は合計 ' + out.length + '名）。', imported: members.length, total: out.length };
  } catch (e) {
    console.error('[MEMBER] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '取り込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}
