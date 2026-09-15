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
        no: '', cat: s.cat || '', name: s.name || '', kana: '', title: s.title || '',
        company: s.company || '', role: '', memo: '', photoFile: '',
        comment: s.comment || '', refer: s.refer || '', collab: s.collab || '',
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
      var memo = members[i].memo || '';
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
