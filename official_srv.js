// === 公式ファイルから雛形を作る ===
//
// BNIメンバーがダウンロードできる公式のpptx（定例会の進行スライド一式）から、このシステムで使う
// 雛形（ビジター紹介・ゲスト紹介・代理紹介・ビジタープレゼン・メンバープレゼン・定例会の前半と後半）を
// 作って「01_テンプレート」に保存し、登録する。次からは、その雛形を直して使う。
// 公式ファイルはリポジトリに入れない。設置者が共有ドライブに置いて、そのリンクを設定する。
//
// 【大きさ】公式ファイルは約100MBあり、Apps Script が一度に扱える大きさ（50MB）を超える。
//   そこで Drive API の「範囲を指定したダウンロード」（Range: bytes=…）で、zipの目次と
//   使う部品だけを取り出す（offZipOpen_ / offZipRead_）。
//   動画は圧縮せずに入っている（1本10MB）ので、雛形に入れるときに圧縮する（1MBほどになる）。
// 【音】ベル・卵時計・コアバリューの動画など、音の出るものは消さない。
// 【作り】雛形ごとの組み立ては official_build_srv.js。ここは読み書きと登録。

var OFF_FILE_KEY_ = 'BNI_OFFICIAL_FILE_ID';        // 公式ファイル（DriveのファイルID）
var OFF_MADE_KEY_ = 'BNI_OFFICIAL_MADE';           // 公式ファイルから作った雛形の記録 { kind: { id, at } }
var OFF_API_ = 'https://www.googleapis.com/drive/v3/files/';
var OFF_BIG_ = 1048576;          // これより大きい圧縮なしの部品は、配列にせず範囲そのままのBlobで受け取る
var OFF_HEAD_ = 4096;            // 部品の頭（ローカルヘッダ）を読むときの余白（名前と拡張欄のぶん）
var OFF_BATCH_ = 16;             // 一度にまとめて取りに行く数（UrlFetchApp.fetchAll）

// 作れる雛形。type: 'visitor' は「PowerPointテンプレート（ビジター用）」、'big' は「大きなスライド」に登録する
var OFF_KINDS_ = [
  { kind: 'intro',         type: 'visitor', label: 'ビジター紹介（3人1枚）', file: 'ビジター紹介' },
  { kind: 'guest',         type: 'visitor', label: 'ゲスト紹介（3人1枚）',   file: 'ゲスト紹介' },
  { kind: 'dairi',         type: 'visitor', label: '代理紹介（3人1枚）',     file: '代理紹介' },
  { kind: 'presen',        type: 'visitor', label: 'ビジタープレゼン（1人1枚）', file: 'ビジタープレゼン' },
  { kind: 'memberPresen',  type: 'big',     label: 'メンバープレゼン',       file: 'メンバープレゼン' },
  { kind: 'meetingFirst',  type: 'big',     label: '定例会スライド（前半）', file: '定例会前半' },
  { kind: 'meetingSecond', type: 'big',     label: '定例会スライド（後半）', file: '定例会後半' }
];

function openOfficialTemplatesDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('official_templates').setWidth(760).setHeight(680), '公式ファイルから雛形を作る');
}

function offKindDef_(kind) {
  for (var i = 0; i < OFF_KINDS_.length; i++) if (OFF_KINDS_[i].kind === kind) return OFF_KINDS_[i];
  return null;
}
// その種類の雛形として登録してあるファイルID
function offRegisteredId_(def) {
  var props = PropertiesService.getScriptProperties();
  var prop = def.type === 'visitor' ? TEMPLATE_KINDS_[def.kind].prop : BIG_TEMPLATE_KINDS_[def.kind].prop;
  return props.getProperty(prop) || '';
}
function offMade_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(OFF_MADE_KEY_) || '{}') || {}; } catch (e) { return {}; }
}

// --- 画面から呼ぶ ---
function getOfficialTemplateStatus() {
  try {
    var props = PropertiesService.getScriptProperties(), id = props.getProperty(OFF_FILE_KEY_) || '';
    var file = null;
    if (id) {
      try {
        var f = DriveApp.getFileById(id);
        file = { id: id, name: f.getName(), url: f.getUrl(), sizeMB: Math.round(f.getSize() / 104857.6) / 10,
                 mimeType: f.getMimeType() };
      } catch (e) { file = { id: id, error: 'ファイルを開けません（アクセス権・リンクをご確認ください）' }; }
    }
    var made = offMade_(), kinds = OFF_KINDS_.map(function (def) {
      var rid = offRegisteredId_(def), row = { kind: def.kind, label: def.label, registered: false, fileName: '', url: '',
                                              fromOfficial: false };
      if (rid) {
        try { var t = DriveApp.getFileById(rid); row.registered = true; row.fileName = t.getName(); row.url = t.getUrl(); }
        catch (e) {}
        row.fromOfficial = !!(made[def.kind] && made[def.kind].id === rid);
        if (row.fromOfficial) row.madeAt = made[def.kind].at || '';
      }
      return row;
    });
    return { ok: true, file: file, kinds: kinds, seconds: chapterPresenSeconds_(), chapter: chapterLabel_() };
  } catch (e) {
    console.error('[OFFICIAL] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み込めませんでした: ' + (e && e.message ? e.message : e) };
  }
}

function saveOfficialFileLink(linkOrId) {
  try {
    var id = extractDriveId_(linkOrId);
    if (!id) return { ok: false, message: 'ファイルのリンクまたはIDを認識できませんでした。Googleドライブの共有リンクを貼り付けてください。' };
    var f;
    try { f = DriveApp.getFileById(id); }
    catch (e) { return { ok: false, message: 'このリンクのファイルを開けませんでした。共有ドライブのフォルダに置いたか、アクセス権があるかをご確認ください。' }; }
    var mime = f.getMimeType();
    if (mime === SLIDES_MIME_) {
      return { ok: false, message: 'Googleスライドに変換されたファイルです。ダウンロードした .pptx をそのままアップロードしたファイルを指定してください'
        + '（「アップロード時に変換」を切っておく）。' };
    }
    if (mime !== PPTX_MIME_ && !/\.pptx$/i.test(f.getName())) {
      return { ok: false, message: 'PowerPoint（.pptx）のファイルを指定してください（いまのファイル: ' + mime + '）。' };
    }
    PropertiesService.getScriptProperties().setProperty(OFF_FILE_KEY_, id);
    return { ok: true, message: '公式ファイルに「' + f.getName() + '」を設定しました。', status: getOfficialTemplateStatus() };
  } catch (e) {
    console.error('[OFFICIAL] save ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '保存できませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// 雛形を1つ作って保存し、登録する。画面から種類ごとに順番に呼ぶ（1回の実行時間の上限に掛からないように）
function buildOfficialTemplate(kind) {
  var t0 = new Date().getTime();
  try {
    var def = offKindDef_(kind);
    if (!def) return { ok: false, message: '雛形の種類が不正です。' };
    var id = PropertiesService.getScriptProperties().getProperty(OFF_FILE_KEY_);
    if (!id) return { ok: false, message: '公式ファイルのリンクを先に設定してください。' };
    var zip = offZipOpen_(id);
    var built = offBuildTemplate_(zip, kind, offBuildOptions_());
    var name = 'BNI_テンプレート_' + def.file + '（公式から作成）.pptx';
    var blob = zipFromMap_(built.map, name);
    var saved = offSaveTemplate_(blob, name);
    offRegister_(def, saved.id);
    var sizeMB = 0;
    try { sizeMB = Math.round(DriveApp.getFileById(saved.id).getSize() / 104857.6) / 10; } catch (e) {}
    console.log('[OFFICIAL] ' + kind + ' -> ' + saved.id + ' ' + sizeMB + 'MB ' + (new Date().getTime() - t0) + 'ms');
    return { ok: true, kind: kind, fileName: name, url: saved.url, sizeMB: sizeMB,
             message: '「' + def.label + '」を作って登録しました（' + sizeMB + 'MB'
               + (built.note ? '・' + built.note : '') + '）。', status: getOfficialTemplateStatus() };
  } catch (e) {
    console.error('[OFFICIAL] build ' + kind + ' ' + (e && e.stack ? e.stack : e));
    return { ok: false, kind: kind, message: '作れませんでした: ' + driveHelpHint_(e) };
  }
}

// 雛形に入れるチャプターごとの値（チャプター名・次回の開催回と日・秒数）
function offBuildOptions_() {
  var next = null;
  try { next = getMeetingCandidates()[0] || null; } catch (e) { next = null; }
  var d = next ? parseDate_(next.dateValue) : null, no = next ? String(next.display || '').match(/第(\d+)回/) : null;
  return { chapter: chapterInfo_().name, seconds: chapterPresenSeconds_(),
           meetingNo: no ? no[1] : '1', meetingDate: d || new Date() };
}

// 01_テンプレート に保存（同じ名前があれば中身を入れ替えて、リンクを変えない）
function offSaveTemplate_(blob, name) {
  var folder = getAssetFolder_('template'), it = folder.getFilesByName(name), file;
  blob.setName(name).setContentType(PPTX_MIME_);
  if (it.hasNext()) {
    file = it.next();
    try { Drive.Files.update({}, file.getId(), blob); }
    catch (e) { file.setTrashed(true); file = folder.createFile(blob); }
  } else {
    file = folder.createFile(blob);
  }
  return { id: file.getId(), url: file.getUrl() };
}

function offRegister_(def, fileId) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(def.type === 'visitor' ? TEMPLATE_KINDS_[def.kind].prop : BIG_TEMPLATE_KINDS_[def.kind].prop, fileId);
  var made = offMade_(), now = new Date();
  made[def.kind] = { id: fileId, at: now.getFullYear() + '/' + (now.getMonth() + 1) + '/' + now.getDate() };
  props.setProperty(OFF_MADE_KEY_, JSON.stringify(made));
}

// === 公式ファイルを、大きいまま部品だけ読む（zip の目次を読み、範囲を指定して取り出す）===

// zip を開く → { id, size, name, entries: { パス: { name, method, crc, csize, usize, off } }, names: [...] }
function offZipOpen_(fileId) {
  var file = DriveApp.getFileById(fileId), size = file.getSize();
  if (!(size > 22)) throw new Error('公式ファイルが空です。');
  var tailLen = Math.min(size, 65557);
  var tail = offFetchBytes_(fileId, [[size - tailLen, size - 1]])[0], e = -1, i;
  for (i = tail.length - 22; i >= 0; i--) {
    if (offU32_(tail, i) === 0x06054b50) { e = i; break; }
  }
  if (e < 0) throw new Error('pptx（zip）として読めませんでした。ダウンロードした公式ファイルをそのまま置いてください。');
  var count = offU16_(tail, e + 10), cdSize = offU32_(tail, e + 12), cdOff = offU32_(tail, e + 16);
  if (count === 0xffff || cdOff === 0xffffffff) throw new Error('この形式（ZIP64）の pptx には対応していません。');
  var cd = offFetchBytes_(fileId, [[cdOff, cdOff + cdSize - 1]])[0], p = 0, entries = {}, names = [];
  for (i = 0; i < count; i++) {
    if (offU32_(cd, p) !== 0x02014b50) throw new Error('pptx の目次が壊れています。');
    var nlen = offU16_(cd, p + 28), xlen = offU16_(cd, p + 30), clen = offU16_(cd, p + 32);
    var ent = { method: offU16_(cd, p + 10), crc: offU32_(cd, p + 16), csize: offU32_(cd, p + 20),
                usize: offU32_(cd, p + 24), off: offU32_(cd, p + 42), flags: offU16_(cd, p + 8),
                name: offUtf8_(cd, p + 46, nlen) };
    p += 46 + nlen + xlen + clen;
    if (/\/$/.test(ent.name)) continue;
    if (ent.method !== 0 && ent.method !== 8) throw new Error('対応していない圧縮方式の部品があります: ' + ent.name);
    entries[ent.name] = ent;
    names.push(ent.name);
  }
  return { id: fileId, size: size, name: file.getName(), entries: entries, names: names };
}

// 部品を読む → { パス: Blob }。無い部品は入れない
function offZipRead_(zip, names) {
  var small = [], big = [], out = {}, i;
  for (i = 0; i < names.length; i++) {
    var e = zip.entries[names[i]];
    if (!e || out[e.name]) continue;
    (e.method === 0 && e.csize >= OFF_BIG_ ? big : small).push(e);
  }
  // 小さな部品：近くにあるものをまとめて1回で取る
  small.sort(function (a, b) { return a.off - b.off; });
  var groups = [], g = null;
  for (i = 0; i < small.length; i++) {
    var s = small[i], end = Math.min(zip.size - 1, s.off + 30 + OFF_HEAD_ + s.csize);
    if (g && s.off - g.end < 65536 && end - g.start < 8388608) { g.items.push(s); g.end = Math.max(g.end, end); }
    else { g = { start: s.off, end: end, items: [s] }; groups.push(g); }
  }
  var bytes = offFetchBytes_(zip.id, groups.map(function (x) { return [x.start, x.end]; })), retry = [];
  for (i = 0; i < groups.length; i++) {
    for (var k = 0; k < groups[i].items.length; k++) {
      var it = groups[i].items[k], at = it.off - groups[i].start, buf = bytes[i];
      if (offU32_(buf, at) !== 0x04034b50) throw new Error('部品の頭が見つかりません: ' + it.name);
      var from = at + 30 + offU16_(buf, at + 26) + offU16_(buf, at + 28);
      if (from + it.csize > buf.length) { retry.push(it); continue; }
      out[it.name] = offPartBlob_(buf.slice(from, from + it.csize), it);
    }
  }
  // 大きな部品（圧縮なし）と、まとめて取れなかった部品：頭を読んで中身の位置を決め、中身だけを取る
  var rest = big.concat(retry);
  if (rest.length) {
    var heads = offFetchBytes_(zip.id, rest.map(function (x) { return [x.off, Math.min(zip.size - 1, x.off + 30 + OFF_HEAD_)]; }));
    var ranges = rest.map(function (x, j) {
      var h = heads[j];
      if (offU32_(h, 0) !== 0x04034b50) throw new Error('部品の頭が見つかりません: ' + x.name);
      var from = x.off + 30 + offU16_(h, 26) + offU16_(h, 28);
      return [from, from + x.csize - 1];
    });
    var resp = offFetch_(zip.id, ranges);
    for (i = 0; i < rest.length; i++) {
      if (rest[i].method === 0) out[rest[i].name] = resp[i].getBlob().setName(rest[i].name);
      else out[rest[i].name] = offPartBlob_(resp[i].getContent(), rest[i]);
    }
  }
  return out;
}

// 中身のバイト列 → Blob（圧縮されていれば戻す）
function offPartBlob_(data, e) {
  if (e.method === 0) return Utilities.newBlob(offSigned_(data), offMime_(e.name), e.name);
  // 生の deflate は、gzip の頭と尻尾（CRC32・元の大きさ。どちらも目次にある）を付ければ Utilities.ungzip で戻せる
  var head = [0x1f, 0x8b, 8, 0, 0, 0, 0, 0, 0, 0xff], tail = [], c = e.crc, u = e.usize, i;
  for (i = 0; i < 4; i++) { tail.push(c % 256); c = Math.floor(c / 256); }
  for (i = 0; i < 4; i++) { tail.push(u % 256); u = Math.floor(u / 256); }
  var gz = Utilities.newBlob(offSigned_(head.concat(offArray_(data), tail)), 'application/x-gzip', e.name + '.gz');
  return Utilities.ungzip(gz).setName(e.name).setContentType(offMime_(e.name));
}

function offMime_(name) {
  var ext = (String(name).match(/\.([A-Za-z0-9]+)$/) || [])[1];
  ext = ext ? ext.toLowerCase() : '';
  return { xml: 'application/xml', rels: 'application/xml', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
           gif: 'image/gif', svg: 'image/svg+xml', emf: 'image/x-emf', wmf: 'image/x-wmf', mp4: 'video/mp4',
           mp3: 'audio/mpeg', wav: 'audio/wav', m4a: 'audio/mp4' }[ext] || 'application/octet-stream';
}

// 範囲ごとに取りに行く → HTTPResponse の並び（範囲指定が効いていなければ止める）
function offFetch_(fileId, ranges) {
  var token = ScriptApp.getOAuthToken(), out = [], i;
  var url = OFF_API_ + encodeURIComponent(fileId) + '?alt=media&supportsAllDrives=true';
  for (i = 0; i < ranges.length; i += OFF_BATCH_) {
    var reqs = ranges.slice(i, i + OFF_BATCH_).map(function (r) {
      return { url: url, method: 'get', muteHttpExceptions: true,
               headers: { Authorization: 'Bearer ' + token, Range: 'bytes=' + r[0] + '-' + r[1] } };
    });
    var res = UrlFetchApp.fetchAll(reqs);
    for (var k = 0; k < res.length; k++) {
      var code = res[k].getResponseCode();
      if (code !== 206 && !(code === 200 && ranges[i + k][0] === 0)) {
        throw new Error('公式ファイルを読めませんでした（' + code + '）。'
          + (code === 403 || code === 404 ? 'アクセス権をご確認ください。' : '') + String(res[k].getContentText()).slice(0, 200));
      }
      out.push(res[k]);
    }
  }
  return out;
}
function offFetchBytes_(fileId, ranges) {
  return offFetch_(fileId, ranges).map(function (r) { return r.getContent(); });
}

// --- バイト列の小さな道具（Apps Script の getContent() は符号付きなので & 255 して読む）---
function offU16_(b, i) { return (b[i] & 255) | ((b[i + 1] & 255) << 8); }
function offU32_(b, i) { return ((b[i] & 255) | ((b[i + 1] & 255) << 8) | ((b[i + 2] & 255) << 16)) + (b[i + 3] & 255) * 16777216; }
function offSigned_(a) {
  var out = [];
  for (var i = 0; i < a.length; i++) { var v = a[i] & 255; out.push(v > 127 ? v - 256 : v); }
  return out;
}
function offArray_(a) {
  var out = [];
  for (var i = 0; i < a.length; i++) out.push(a[i]);
  return out;
}
// UTF-8 のバイト列 → 文字列（zip の中の名前）
function offUtf8_(b, from, len) {
  var s = '', i = from, end = from + len;
  while (i < end) {
    var c = b[i++] & 255;
    if (c < 0x80) { s += String.fromCharCode(c); continue; }
    var n = c >= 0xf0 ? 3 : c >= 0xe0 ? 2 : 1, cp = c & (n === 3 ? 7 : n === 2 ? 15 : 31);
    for (var k = 0; k < n && i < end; k++) cp = (cp << 6) | (b[i++] & 63);
    if (cp > 0xffff) { cp -= 0x10000; s += String.fromCharCode(0xd800 + (cp >> 10), 0xdc00 + (cp & 1023)); }
    else s += String.fromCharCode(cp);
  }
  return s;
}
