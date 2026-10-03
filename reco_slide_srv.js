// === 推薦のことば：受け取ったスライドを後半スライドに入れる ===
//
// 推薦のことばを話す方から受け取ったスライド（pptx・pdf・画像の1枚）を、後半スライドの画面で組ごとに選ぶと、
// 作成するときに、その組の推薦のことばのページのすぐ後ろに1枚のページとして入れる
// （受け取ったスライドの画像を、縦横比のままページいっぱいに。余ったところは黒）。
//
//   画像 … 画面で長い辺を縮め、PNG（写真はJPEG）にしてから送る
//   PDF  … 画面で1ページ目を画像にしてから送る（pdf.js）
//   pptx … サーバーでGoogleスライドに変換してPDFにし（convertRecommendationPptx）、画面に返す。
//          画面はPDFと同じく1ページ目を画像にしてから送る
// 送られた画像は素材フォルダの「03_生成物／推薦のことばのスライド」に置き、
// 開催日・組（推薦する方→推薦される方）ごとに覚えておく（画面を閉じても、次に開いたとき同じ組に付く）。
// 画面は slides_meeting_second.html（3. 推薦のことば）。ページを作るのは meeting_slides_srv.js の expandRecommendations_。

var RECO_SLIDE_FOLDER_ = '推薦のことばのスライド';
var RECO_SLIDE_KEY_ = 'BNI_RECO_SLIDES';            // { 'yyyy/MM/dd': [{ g, r, id, name, w, h, at }] }
var RECO_SLIDE_KEEP_DAYS_ = 120;                    // これより前の開催日の覚えは消す（ファイルはドライブに残る）
var RECO_SLIDE_MAX_BYTES_ = 15 * 1024 * 1024;       // 送られる画像・pptx の上限
var RECO_PPTX_MIME_ = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

function recoSlideFolder_() { return ensureChildFolder_(getAssetFolder_('output'), RECO_SLIDE_FOLDER_); }

function recoSlideAll_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(RECO_SLIDE_KEY_);
    var all = raw ? JSON.parse(raw) : {};
    return (all && typeof all === 'object') ? all : {};
  } catch (e) { return {}; }
}
function recoSlideSaveAll_(all) {
  var from = new Date(); from.setDate(from.getDate() - RECO_SLIDE_KEEP_DAYS_);
  var cut = fmtDate_(from), out = {};
  Object.keys(all).forEach(function (k) { if (k >= cut && all[k] && all[k].length) out[k] = all[k]; });
  PropertiesService.getScriptProperties().setProperty(RECO_SLIDE_KEY_, JSON.stringify(out));
}
function recoSlideView_(x) { return { giver: x.g, receiver: x.r, id: x.id, name: x.name, w: x.w, h: x.h }; }
function recoTrash_(id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
// ファイル名に使えない字を除く
function recoSafeName_(s) { return String(s == null ? '' : s).replace(/[\s　]+/g, '').replace(/[\\\/:*?"<>|]/g, '') || '未選択'; }

// その開催日に覚えている、受け取ったスライド
function getRecommendationSlides(dateStr) {
  try {
    var d = parseDate_(dateStr);
    if (!d) return { ok: false, message: '開催日が分かりません。', slides: [] };
    return { ok: true, slides: (recoSlideAll_()[fmtDate_(d)] || []).map(recoSlideView_) };
  } catch (e) {
    return { ok: false, message: '受け取ったスライドを読めませんでした: ' + (e && e.message ? e.message : e), slides: [] };
  }
}

// 画面で画像にしたスライドを置く。req … { date, giver, receiver, name（元のファイル名）, data（画像の data URL か base64） }
// 同じ組に前に置いたものは、ゴミ箱に入れて置き換える
function saveRecommendationSlide(req) {
  var lock = LockService.getScriptLock(), locked = false;
  try {
    var r = req || {}, d = parseDate_(r.date);
    if (!d) return { ok: false, message: '開催日が分かりません。' };
    var g = String(r.giver || '').trim(), rv = String(r.receiver || '').trim();
    if (!g && !rv) return { ok: false, message: '先に、推薦する方・推薦される方を選んでください。' };
    var bytes = Utilities.base64Decode(String(r.data || '').replace(/^data:[^,]*,/, ''));
    if (!bytes.length) return { ok: false, message: 'ファイルが空です。' };
    if (bytes.length > RECO_SLIDE_MAX_BYTES_) return { ok: false, message: '画像が大きすぎます（15MBまで）。' };
    var png = (bytes[0] & 0xff) === 0x89 && (bytes[1] & 0xff) === 0x50;
    var jpg = (bytes[0] & 0xff) === 0xff && (bytes[1] & 0xff) === 0xd8;
    var size = (png || jpg) ? imageSizeOf_(bytes) : null;
    if (!size || !size.width || !size.height) return { ok: false, message: '画像として読めませんでした（PNG・JPEG にしてください）。' };
    var fileName = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyyMMdd') + '_推薦のことば_' + recoSafeName_(g) + '→' + recoSafeName_(rv)
      + (png ? '.png' : '.jpg');
    var file;
    try { file = recoSlideFolder_().createFile(Utilities.newBlob(bytes, png ? 'image/png' : 'image/jpeg', fileName)); }
    catch (e) { return { ok: false, message: '受け取ったスライドをドライブに置けませんでした。' + driveHelpHint_(e) }; }
    locked = lock.tryLock(20000);
    var all = recoSlideAll_(), key = fmtDate_(d);
    var list = (all[key] || []).filter(function (x) {
      if (x.g === g && x.r === rv) { if (x.id !== file.getId()) recoTrash_(x.id); return false; }
      return true;
    });
    var item = { g: g, r: rv, id: file.getId(), name: String(r.name || fileName).slice(0, 120), w: size.width, h: size.height,
                 at: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') };
    list.push(item);
    all[key] = list;
    recoSlideSaveAll_(all);
    return { ok: true, slide: recoSlideView_(item),
             message: '「' + item.name + '」を、' + (g || '（推薦する方）') + ' → ' + (rv || '（推薦される方）') + ' のページのあとに入れます。' };
  } catch (e) {
    console.error('[RECO] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '受け取ったスライドを置けませんでした: ' + (e && e.message ? e.message : e) };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

// 受け取ったスライドを外す（ファイルはゴミ箱へ）
function removeRecommendationSlide(dateStr, id) {
  var lock = LockService.getScriptLock(), locked = false;
  try {
    var d = parseDate_(dateStr);
    if (!d) return { ok: false, message: '開催日が分かりません。' };
    locked = lock.tryLock(20000);
    var all = recoSlideAll_(), key = fmtDate_(d), hit = false;
    all[key] = (all[key] || []).filter(function (x) { if (x.id === id) { hit = true; return false; } return true; });
    recoSlideSaveAll_(all);
    recoTrash_(id);
    return { ok: true, message: hit ? '受け取ったスライドを外しました。' : '外すスライドが見つかりませんでした（もう外してあります）。' };
  } catch (e) {
    return { ok: false, message: '外せませんでした: ' + (e && e.message ? e.message : e) };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

// pptx → PDF（Googleスライドに変換して書き出す。変換に使ったファイルはゴミ箱へ）。req … { name, data（base64） }
// 1ページ目を画像にするのは画面（PDFと同じ処理）
function convertRecommendationPptx(req) {
  var tmpId = '';
  try {
    var r = req || {};
    var bytes = Utilities.base64Decode(String(r.data || '').replace(/^data:[^,]*,/, ''));
    if (!bytes.length) return { ok: false, message: 'ファイルが空です。' };
    if (bytes.length > RECO_SLIDE_MAX_BYTES_) return { ok: false, message: 'pptx が大きすぎます（15MBまで）。PDFか画像にして選んでください。' };
    if ((bytes[0] & 0xff) !== 0x50 || (bytes[1] & 0xff) !== 0x4b) return { ok: false, message: 'pptx として読めませんでした。' };
    var name = String(r.name || 'スライド.pptx');
    var meta = Drive.Files.create({ name: '（変換中）' + name, mimeType: 'application/vnd.google-apps.presentation',
                                    parents: [recoSlideFolder_().getId()] },
                                  Utilities.newBlob(bytes, RECO_PPTX_MIME_, name), { supportsAllDrives: true });
    tmpId = meta.id;
    var pdf = DriveApp.getFileById(tmpId).getAs('application/pdf');
    return { ok: true, name: name, pdf: Utilities.base64Encode(pdf.getBytes()) };
  } catch (e) {
    console.error('[RECO] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'pptx を画像にできませんでした（Googleスライドへの変換）。PowerPointで「PDFとして保存」したものか、'
      + '画像を選んでください。（' + (e && e.message ? e.message : e) + '）' };
  } finally {
    if (tmpId) recoTrash_(tmpId);
  }
}

// 作成の前に、組に付いた受け取ったスライド（ファイルのID）を開いておく（expandRecommendations_ で使う）。
// 開けなかったものは外して、お知らせの文を返す
function recoLoadSlides_(pairs) {
  var gone = [];
  (pairs || []).forEach(function (p) {
    var s = p && p.slide;
    if (!s || !s.id) return;
    try {
      var blob = DriveApp.getFileById(s.id).getBlob(), bytes = blob.getBytes(), size = imageSizeOf_(bytes);
      if (!size) throw new Error('画像として読めません');
      var png = (bytes[0] & 0xff) === 0x89;
      p.slideImage = { blob: blob, ext: png ? 'png' : 'jpeg', width: size.width, height: size.height, name: s.name || '' };
    } catch (e) {
      console.warn('[RECO] 受け取ったスライドを開けませんでした: ' + s.id + ' ' + (e && e.message ? e.message : e));
      gone.push(s.name || s.id);
    }
  });
  return gone.length ? '受け取ったスライドを開けなかったので入れませんでした: ' + gone.join('、') : '';
}

// 受け取ったスライドの画像を、1枚のページにする（縦横比のまま、ページいっぱいに真ん中へ。余ったところは黒）。
// レイアウトは推薦のことばのページと同じもの。レイアウト・マスターの飾り（ロゴなど）は出さない（showMasterSp="0"）。
//   img … { blob, ext: 'png'|'jpeg', width, height, name }
// 戻り値は addSlidePart_ と同じ（並びには setSlideEntries_ で入れる）
function recoImageSlide_(parts, layoutRels, img, seq) {
  var prs = xmlOf_(parts, 'ppt/presentation.xml') || '';
  var sz = prs.match(/<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
  var W = sz ? parseInt(sz[1], 10) : 12192000, H = sz ? parseInt(sz[2], 10) : 6858000;
  var scale = Math.min(W / img.width, H / img.height);
  var cx = Math.round(img.width * scale), cy = Math.round(img.height * scale);
  var x = Math.round((W - cx) / 2), y = Math.round((H - cy) / 2);
  var n = seq || 1, media = 'ppt/media/recoslide' + n + '.' + img.ext;
  while (parts[media]) media = 'ppt/media/recoslide' + (++n) + '.' + img.ext;
  parts[media] = img.blob.setName(media);
  var layout = (String(layoutRels || '').match(/<Relationship\b[^>]*Type="[^"]*\/slideLayout"[^>]*\/>/) || [''])[0];
  var lt = (layout.match(/Target="([^"]+)"/) || [])[1] || ('../slideLayouts/' + firstLayout_(parts).replace(/^.*\//, ''));
  var REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="' + REL + 'slideLayout" Target="' + lt + '"/>'
    + '<Relationship Id="rId2" Type="' + REL + 'image" Target="../media/' + media.replace(/^.*\//, '') + '"/></Relationships>';
  var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + REL.replace(/\/$/, '')
    + '" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" showMasterSp="0">'
    + '<p:cSld name="推薦のことば（受け取ったスライド）"><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
    + '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    + '<p:pic><p:nvPicPr><p:cNvPr id="2" name="受け取ったスライド" descr="' + escapeXml_(img.name || '') + '"/>'
    + '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
    + '<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
    + '<p:spPr><a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>'
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  var add = addSlidePart_(parts, xml, rels);
  putXml_(parts, '[Content_Types].xml', ensureDefaultType_(xmlOf_(parts, '[Content_Types].xml'), img.ext));
  return add;
}
