// === 事前MTG（朝イチMTG）のパワポ ===
//
// 役職ごとの入力でルーティンチェックシートに入った内容から、事前MTG（朝イチMTG）のパワポを作る。
// これまで「BNI週次役職情報共有 パワポ自動生成ツール」が、フォームの回答（CSV）から作っていたものと同じ中身。
//
//   まとめ（1枚）… 直近のイベント／お願い事項（三役から）／定例会関連
//                  （欠席・代理・医療欠席の人数、ビジターなどの人数、ウィークリープレゼン・スタートアッププレゼン・
//                    メインプレゼンの担当、新入会・更新・退会、卒業コメント）
//   役職のページ … 今週の共有事項を、役職の順（プレジデント → … → スプレディング委員）に、担当者の写真・お名前つきで。
//                  短いもの（180文字以下）同士は2人で1枚にする（ツールと同じ）。
//                  共有事項が空・「なし」の役職のページは作らない。
//
// ひな形 … 「⚙️ 設定 ＞ 大きなスライド」の「事前MTG（朝イチMTG）」に登録したpptx。
//          登録していなければ、同梱の既定のひな形（premtg_template.html。これまでと同じデザイン）で作る。
//          差し込み口（{{月日}}・{{共有事項1}} など）は docs/templates/README.md に。
//
// 画面は role_input.html（入力状況の一覧の「事前MTG（朝イチMTG）のパワポ」）。

var PREMTG_KIND_ = 'preMeeting';
var PREMTG_PAIR_MAX_ = 180;                  // この文字数以下同士なら2人で1枚（ツールと同じ）
var PREMTG_MIN_PT_ = 9;                      // 文字が多いときに小さくする下限
var PREMTG_COLORS_ = { top: 'C8102E', coord: '2E5C9A', comm: '3B8763' };   // 三役・コーディネーター・委員会

// ツールと同じ並び・区分・アイコン
var PREMTG_ROLES_ = [
  { key: 'president', cat: 'top', icon: '👑' },
  { key: 'vice', cat: 'top', icon: '🎯' },
  { key: 'secretary', cat: 'top', icon: '💰' },
  { key: 'vhc', cat: 'coord', icon: '🤝' },
  { key: 'mentor', cat: 'coord', icon: '🧭' },
  { key: 'ec', cat: 'coord', icon: '📚' },
  { key: 'gbc', cat: 'coord', icon: '🌏' },
  { key: 'web', cat: 'comm', icon: '💻' },
  { key: 'support', cat: 'comm', icon: '🛟' },
  { key: 'training', cat: 'comm', icon: '🏋' },
  { key: 'event', cat: 'comm', icon: '🎉' },
  { key: 'bcp', cat: 'comm', icon: '🛡' },
  { key: 'spreading', cat: 'comm', icon: '📣' }
];
// まとめの「お願い事項」に入れる欄（【役職】を頭に付けて、この順に並べる）
var PREMTG_REQUESTS_ = [
  { label: 'プレジデント', title: 'お願い事項（プレジデントから）' },
  { label: 'バイスプレジデント', title: 'お願い事項（バイスプレジデントから）' },
  { label: '書記兼会計', title: 'お願い事項（書記兼会計から）' }
];
// まとめのページの3つのまとまり（差し込み口の名前）
var PREMTG_BLOCKS_ = ['直近のイベント', 'お願い事項', '定例会関連'];

// --- 画面を開く（役職ごとの入力の一覧を、事前MTGの欄を開いた状態で）---
// メニューは「入力状況の一覧・事前MTGのパワポ」にまとめた。これは前のメニューが残っているとき用
function openPreMeetingDialog() { openRoleInputDialog_('', 'premtg'); }

// --- 材料を集める ---
// 書き込みを整える（改行をそろえ、行末の空白と前後の空行を落とす）
function premtgText_(v) {
  var s = String(v == null ? '' : v).replace(/\r\n?/g, '\n').replace(/[ \t　]+$/gm, '');
  return s.replace(/^\n+|\n+$/g, '');
}
// 「なし」「特になし」「ー」などの、中身の無い書き込み
function premtgNone_(s) {
  var t = String(s == null ? '' : s).normalize('NFKC').replace(/[\s。．.！!]/g, '');
  return /^(なし|無し|特になし|特に無し|ありません|特にありません|ー|-|―|—|‐|0)$/.test(t);
}

// 「代理・欠席」の欄などに書かれたお名前の数。
//   ・かっこの中（「（当欠）」「（代理：○○さん）」など）は数えない
//   ・「遅刻：○○さん」「○○さん遅刻」は欠席ではないので数えない
//   ・「○○さん代理：△△さん」「○○さん→△△さん」は○○さんだけ数える
//   ・「なし」「○日○時時点でなし」は0
// 敬称（さん・様など）が1つでもあれば、敬称の数を人数とする（「梅田さん　相田さん」は2名）。
// 敬称が1つも無ければ、区切り（、・／改行など）で分けた数とする（「大庭ED」は1名）。
// 画面（role_input.html の countNames）も同じ数え方にしてある。
function premtgCountNames_(text) {
  var s = String(text == null ? '' : text).normalize('NFKC'), out = '', depth = 0;
  for (var i = 0; i < s.length; i++) {                     // かっこの中を落とす（入れ子・閉じ忘れも）
    var c = s.charAt(i);
    if (c === '(' || c === '「') { depth++; continue; }
    if (c === ')' || c === '」') { if (depth > 0) depth--; continue; }
    if (!depth) out += c;
  }
  var HONOR = /(さん|様|さま|氏|くん|君|ちゃん)/g, items = [];
  var segs = out.split(/[\/\n]+/);
  for (var g = 0; g < segs.length; g++) {
    var seg = segs[g];
    var late = seg.search(/遅刻\s*:/);                      // 「遅刻：○○さん、△△さん」より後ろは遅刻
    if (late >= 0) seg = seg.substring(0, late);
    var parts = seg.split(/[、,※・]+/);
    for (var p = 0; p < parts.length; p++) {
      var it = parts[p].split(/代理|→/)[0].trim();          // 「○○さん代理：…」「○○さん→…」
      if (!it || /遅刻/.test(it)) continue;                  // 「○○さん遅刻」
      if (/(なし|無し|休会)/.test(it) && !it.match(HONOR)) continue;
      if (premtgNone_(it)) continue;
      items.push(it);
    }
  }
  var honor = 0;
  for (var k = 0; k < items.length; k++) honor += (items[k].match(HONOR) || []).length;
  return honor ? honor : items.length;
}

// 開催日の材料。役職ごとの入力（ルーティンチェックシート）から読む。
//   summary … まとめのページの3つのまとまり（行の配列）
//   roles   … 役職ごとの共有事項（PREMTG_ROLES_ の順）
//   blank   … 空欄の項目（パワポでは「—」になるもの）。{ label, role }
function premtgData_(target) {
  var ctx = roleBuildContext_(target, '');
  var out = { ok: true, date: fmtDate_(target), md: (target.getMonth() + 1) + '/' + target.getDate(),
              display: roleMd_(target), meetingNo: ctx.meetingNo || '', sheetName: ctx.sheetName || '',
              summary: {}, roles: [], blank: [] };
  if (!ctx.found) {
    out.ok = false;
    out.message = ctx.message || ('ルーティンチェックシートに ' + out.date + ' の列が見つかりません。');
    return out;
  }
  var roleLabel = function (it) { var d = it && it.roles && it.roles[0] ? roleDefOf_(it.roles[0]) : null; return d ? d.label : ''; };
  var find = function (re, parentRe) {
    for (var i = 0; i < ctx.order.length; i++) {
      var it = ctx.items[ctx.order[i]];
      if (!re.test(roleMatchText_(it.title))) continue;
      if (parentRe && !parentRe.test(roleMatchText_(it.parent))) continue;
      return it;
    }
    return null;
  };
  // 値を読む。空欄なら blank に入れる（label はパワポでの呼び名）
  var read = function (it, label) {
    var v = it ? premtgText_(it.value) : '';
    if (!v && label) out.blank.push({ label: label, role: roleLabel(it) });
    return v;
  };
  var extra = function (title, label) { return read(ctx.items[roleExtraId_(title)], label); };
  var sheet = function (re, parentRe, label) { return read(find(re, parentRe), label); };
  var oneLine = function (v) { return v ? v.replace(/\n+/g, '　') : '—'; };
  var people = function (v) { return v ? premtgCountNames_(v) + '名' : '—'; };
  var count = function (v) {
    if (!v) return '';
    var s = v.normalize('NFKC').replace(/\s/g, '');
    if (premtgNone_(s)) return '0名';
    if (/^\d+$/.test(s)) return s + '名';
    if (/^\d+名$/.test(s)) return s;
    return v.replace(/\n+/g, '　');
  };

  // 直近のイベント
  var events = extra('直近のイベント', '直近のイベント');
  out.summary['直近のイベント'] = events ? events.split('\n') : [];

  // お願い事項（三役から。1行目の頭に【役職】）
  var req = [];
  for (var r = 0; r < PREMTG_REQUESTS_.length; r++) {
    var rv = extra(PREMTG_REQUESTS_[r].title, '');
    if (!rv || premtgNone_(rv)) continue;
    var ls = rv.split('\n');
    ls[0] = '【' + PREMTG_REQUESTS_[r].label + '】' + ls[0];
    req = req.concat(ls);
  }
  out.summary['お願い事項'] = req;

  // 定例会関連（ツールと同じ並び）
  var ABS = /代理・?欠席/;
  var absent = people(sheet(/^欠席$/, ABS, '代理・欠席 › 欠席'));
  var subs = people(sheet(/^代理$/, ABS, '代理・欠席 › 代理'));
  var medical = people(sheet(/^医療欠席$/, ABS, '代理・欠席 › 医療欠席'));
  var visitors = count(extra('人数：ビジター', '人数：ビジター'));
  var guests = count(extra('人数：ゲスト', '人数：ゲスト'));
  var tours = count(extra('人数：見学', '人数：見学'));
  var region = count(extra('人数：リージョン参加者', '人数：リージョン参加者'));
  var regionNames = sheet(/^リージョン参加者$/, null, '');
  if (regionNames && !premtgNone_(regionNames)) {
    if (!region) region = premtgCountNames_(regionNames) + '名';
    region += '　' + regionNames.replace(/\n+/g, '、');
  }
  out.summary['定例会関連'] = [
    '・欠席：' + absent + '　／　代理：' + subs + '　／　医療欠席：' + medical,
    '・ビジター：' + (visitors || '—'),
    '・ゲスト：' + (guests || '—') + '　／　見学：' + (tours || '—') + '　／　リージョン参加者：' + (region || '—'),
    '・ウィークリープレゼン：' + oneLine(sheet(/^ウィークリープレゼン$/, null, 'ウィークリープレゼン')),
    '・スタートアッププレゼン：' + oneLine(sheet(/^(スタートアッププレゼン|2分30秒プレゼン)$/, null, 'スタートアッププレゼン')),
    '・メインプレゼン：' + oneLine(sheet(/^メインプレゼン$/, null, 'メインプレゼン')),
    '・新入会：' + oneLine(sheet(/^新入会$/, null, '新入会'))
      + '　／　更新：' + oneLine(sheet(/^更新式/, null, '更新式'))
      + '　／　退会：' + oneLine(sheet(/^退会者$/, null, '退会者')),
    '・卒業コメント：' + oneLine(extra('卒業コメント', ''))
  ];

  // 役職ごとの共有事項
  var holders = roleHolders_(target);               // その開催日の期（半期）の担当者
  for (var k = 0; k < PREMTG_ROLES_.length; k++) {
    var def = roleDefOf_(PREMTG_ROLES_[k].key);
    if (!def) continue;
    var text = extra('今週の共有事項（' + def.label + '）', '');
    out.roles.push({ key: def.key, label: def.label, icon: PREMTG_ROLES_[k].icon, cat: PREMTG_ROLES_[k].cat,
                     color: PREMTG_COLORS_[PREMTG_ROLES_[k].cat], holder: holders[def.key] || '',
                     text: premtgNone_(text) ? '' : text });
  }
  return out;
}

// 役職のページの割り付け（ツールと同じ）：共有事項のある役職を順に、
// 隣どうしがどちらも180文字以下なら2人で1枚、そうでなければ1人で1枚
function premtgPages_(roles, canPair) {
  var list = [], pages = [], i;
  for (i = 0; i < roles.length; i++) if (roles[i].text) list.push(roles[i]);
  i = 0;
  while (i < list.length) {
    var a = list[i], b = list[i + 1];
    if (canPair && b && a.text.length <= PREMTG_PAIR_MAX_ && b.text.length <= PREMTG_PAIR_MAX_) {
      pages.push([a, b]); i += 2;
    } else {
      pages.push([a]); i += 1;
    }
  }
  return pages;
}

// --- ひな形 ---
// 同梱の既定のひな形（tools/build_premtg_template.py で作ったpptxをbase64にしたもの）
function premtgBuiltinBlob_() {
  var html = HtmlService.createHtmlOutputFromFile('premtg_template').getContent();
  var m = html.match(/PPTX_BASE64_BEGIN([\s\S]*?)PPTX_BASE64_END/);
  var b64 = (m ? m[1] : html).replace(/[^A-Za-z0-9+\/=]/g, '');
  return Utilities.newBlob(Utilities.base64Decode(b64), PPTX_MIME_, 'BNI_テンプレート_事前MTG.pptx');
}

// 使うひな形：登録したもの、無ければ既定のもの
function premtgTemplateInfo_() {
  var id = PropertiesService.getScriptProperties().getProperty(BIG_TEMPLATE_KINDS_[PREMTG_KIND_].prop) || '';
  if (!id) return { registered: false, id: '', name: '既定のひな形' };
  var f;
  try { f = DriveApp.getFileById(id); }
  catch (e) { return { registered: true, id: id, name: '', error: '登録したひな形を開けませんでした（' + (e && e.message ? e.message : e) + '）' }; }
  // Googleスライドのままでは中身を書き換えられない（PowerPointのファイルとして読む）
  if (f.getMimeType && f.getMimeType() === SLIDES_MIME_) {
    return { registered: true, id: id, name: f.getName(), error: '登録したひな形がGoogleスライドです。PowerPoint（.pptx）のファイルを登録してください' };
  }
  return { registered: true, id: id, name: f.getName() };
}
function premtgTemplateParts_(info) {
  if (info.registered) {
    if (info.error) throw new Error(info.error + '。⚙️ 設定 ＞ 大きなスライド で登録し直してください。');
    return unzipToMap_(DriveApp.getFileById(info.id).getBlob());
  }
  return unzipToMap_(premtgBuiltinBlob_());
}

// 既定のひな形を書き出す（デザインを変えたいときの元にする）
function exportPreMeetingTemplate() {
  try {
    var name = 'BNI_テンプレート_事前MTG.pptx';
    var saved = saveOutputFile_(premtgBuiltinBlob_(), name);
    return { ok: true, url: saved.url, downloadUrl: saved.downloadUrl, fileName: name,
             message: '既定のひな形を「03_生成物」に書き出しました。PowerPointで直してDriveに置き、'
                    + '⚙️ 設定 ＞ 大きなスライド の「事前MTG（朝イチMTG）」にリンクを登録すると、そのデザインで作ります。' };
  } catch (e) {
    return { ok: false, message: '書き出せませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// ひな形の中の、まとめのページ・2人のページ・1人のページ
function premtgModels_(parts) {
  var order = slideOrder_(parts), m = { summary: null, two: null, one: null };
  for (var i = 0; i < order.length; i++) {
    var t = slideText_(xmlOf_(parts, order[i]) || '');
    if (!m.two && t.indexOf('{{共有事項2}}') >= 0) { m.two = order[i]; continue; }
    if (!m.one && t.indexOf('{{共有事項1}}') >= 0) { m.one = order[i]; continue; }
    if (!m.summary) {
      for (var b = 0; b < PREMTG_BLOCKS_.length; b++) {
        if (t.indexOf('{{' + PREMTG_BLOCKS_[b] + '}}') >= 0) { m.summary = order[i]; break; }
      }
    }
  }
  return m;
}

// --- 図形を探す・書き換える小さな道具 ---
// その文字（差し込み口）が入っている文字箱の id
function premtgShapeWith_(xml, text) {
  var sps = findTagRanges_(xml, 'p:sp');
  for (var i = 0; i < sps.length; i++) {
    var seg = xml.substring(sps[i].start, sps[i].end);
    if (slideText_(seg).indexOf(text) < 0) continue;
    var m = seg.match(/<p:cNvPr[^>]*\sid="(\d+)"/);
    if (m) return m[1];
  }
  return '';
}

// 図形の一覧（id・名前・代替テキスト・位置）
function premtgShapes_(xml, tag) {
  var out = [], rs = findTagRanges_(xml, tag);
  for (var i = 0; i < rs.length; i++) {
    var seg = xml.substring(rs[i].start, rs[i].end);
    var nv = seg.match(/<p:cNvPr\b[^>]*>/);
    var off = seg.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>/);
    var ext = seg.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
    if (!nv || !off || !ext) continue;
    var attr = function (k) { var a = nv[0].match(new RegExp('\\s' + k + '="([^"]*)"')); return a ? unescapeXml_(a[1]) : ''; };
    out.push({ id: attr('id'), name: attr('name'), descr: attr('descr'), seg: seg,
               x: parseInt(off[1], 10), y: parseInt(off[2], 10), cx: parseInt(ext[1], 10), cy: parseInt(ext[2], 10) });
  }
  return out;
}

// k人目の写真の枠：名前（または代替テキスト）が「写真k」のもの。無ければ、お名前の文字箱にいちばん近い画像
function premtgPicFor_(xml, k, anchorId) {
  var pics = premtgShapes_(xml, 'p:pic').filter(function (p) { return !/<a:(audio|video)File\b/.test(p.seg); });
  for (var i = 0; i < pics.length; i++) if (pics[i].name === '写真' + k || pics[i].descr === '写真' + k) return pics[i].id;
  var a = anchorId ? readShapeGeomEmu_(xml, anchorId) : null;
  if (!a) return '';
  var ax = a.x + a.cx / 2, ay = a.y + a.cy / 2, best = '', bd = Infinity;
  for (var j = 0; j < pics.length; j++) {
    var d = Math.pow(pics[j].x + pics[j].cx / 2 - ax, 2) + Math.pow(pics[j].y + pics[j].cy / 2 - ay, 2);
    if (d < bd) { bd = d; best = pics[j].id; }
  }
  return best;
}

// k人目の役職の帯：名前が「帯k」のもの。無ければ、役職の文字箱の真ん中を含む図形。
// どちらも、区分の色（赤・青・緑）で塗ってあるものだけ（帯を別の色にしたひな形は、塗り替えない）。
// anyColor … 色を問わない（2人のページの空ける側で、帯を消すとき）
function premtgBandFor_(xml, k, anchorId, anyColor) {
  var sps = premtgShapes_(xml, 'p:sp'), i;
  var colors = [PREMTG_COLORS_.top, PREMTG_COLORS_.coord, PREMTG_COLORS_.comm];
  var ours = function (s) {
    var f = premtgFillOf_(s.seg);
    return !!f && (anyColor || colors.indexOf(f.toUpperCase()) >= 0);
  };
  for (i = 0; i < sps.length; i++) if (sps[i].name === '帯' + k) return ours(sps[i]) ? sps[i].id : '';
  var a = anchorId ? readShapeGeomEmu_(xml, anchorId) : null;
  if (!a) return '';
  var ax = a.x + a.cx / 2, ay = a.y + a.cy / 2, best = '', area = Infinity;
  for (i = 0; i < sps.length; i++) {
    var s = sps[i];
    if (!ours(s)) continue;
    if (ax < s.x || ax > s.x + s.cx || ay < s.y || ay > s.y + s.cy) continue;
    if (s.cx * s.cy < area) { area = s.cx * s.cy; best = s.id; }
  }
  return best;
}
// 図形の塗りの色（線の色ではなく）
function premtgFillOf_(seg) {
  var end = seg.indexOf('</p:spPr>');
  if (end < 0) return '';
  var head = seg.substring(0, end), ln = head.indexOf('<a:ln');
  if (ln >= 0) head = head.substring(0, ln);
  var m = head.match(/<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/);
  return m ? m[1] : '';
}
function premtgRecolor_(xml, id, color) {
  var r = findShapeRange_(xml, id);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), end = seg.indexOf('</p:spPr>');
  if (end < 0) return xml;
  var head = seg.substring(0, end), ln = head.indexOf('<a:ln'), cut = ln >= 0 ? ln : head.length;
  var fill = head.substring(0, cut).replace(/(<a:solidFill>\s*<a:srgbClr val=")[0-9A-Fa-f]{6}(")/, '$1' + color + '$2');
  seg = fill + seg.substring(cut);
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}

// 差し込み口のある段落を、行の数だけの段落にする（書式はその段落のまま）
function premtgExpandToken_(xml, token, lines) {
  var tag = '{{' + token + '}}', paras = findTagRanges_(xml, 'a:p');
  var list = (lines && lines.length) ? lines : [''];
  for (var p = paras.length - 1; p >= 0; p--) {
    var seg = xml.substring(paras[p].start, paras[p].end);
    if (slideText_(seg).indexOf(tag) < 0) continue;
    var out = '';
    for (var i = 0; i < list.length; i++) {
      var m = {}; m[token] = list[i];
      out += replaceTokensInParagraph_(seg, m);
    }
    xml = xml.substring(0, paras[p].start) + out + xml.substring(paras[p].end);
  }
  return xml;
}

// 文字箱の中の文字の大きさをそろえる（sz が無い書式には足す）
function premtgSetSize_(xml, id, pt) {
  var r = findShapeRange_(xml, id);
  if (!r) return xml;
  var sz = String(Math.round(pt * 100));
  var seg = xml.substring(r.start, r.end).replace(/<a:(rPr|endParaRPr)\b([^>]*?)(\/?)>/g, function (all, tag, attrs, close) {
    return '<a:' + tag + (/\ssz="/.test(attrs) ? attrs.replace(/\ssz="\d+"/, ' sz="' + sz + '"') : attrs + ' sz="' + sz + '"') + close + '>';
  }).replace(/<a:r>(\s*)<a:t>/g, '<a:r>$1<a:rPr lang="ja-JP" sz="' + sz + '"/><a:t>');
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}

// 文字箱の寸法（文字の入る幅・高さ[pt]）・行間・もとの文字の大きさ
function premtgBoxMetrics_(xml, id) {
  var r = findShapeRange_(xml, id);
  if (!r) return null;
  var seg = xml.substring(r.start, r.end), g = readShapeGeomEmu_(xml, id);
  if (!g) return null;
  var bp = (seg.match(/<a:bodyPr\b[^>]*>/) || [''])[0];
  var ins = function (k, def) { var m = bp.match(new RegExp('\\s' + k + '="(\\d+)"')); return m ? parseInt(m[1], 10) : def; };
  var ln = seg.match(/<a:lnSpc>\s*<a:spcPct val="(\d+)"/);
  var sz = seg.match(/<a:(?:rPr|endParaRPr)\b[^>]*\ssz="(\d+)"/);
  var paras = findTagRanges_(seg, 'a:p'), lines = [];
  for (var i = 0; i < paras.length; i++) lines.push(slideText_(seg.substring(paras[i].start, paras[i].end)));
  return { g: g, lines: lines,
           widthPt: (g.cx - ins('lIns', 91440) - ins('rIns', 91440)) / 12700,
           heightPt: (g.cy - ins('tIns', 45720) - ins('bIns', 45720)) / 12700,
           insV: ins('tIns', 45720) + ins('bIns', 45720),
           lnSpc: ln ? parseInt(ln[1], 10) / 100000 : 1, basePt: sz ? parseInt(sz[1], 10) / 100 : 18 };
}

// その大きさの文字で、何行になるか（折り返しを含む見積もり。少し余裕を見る）
function premtgLineCount_(lines, pt, widthPt) {
  var n = 0, w = widthPt * 0.95;
  for (var i = 0; i < lines.length; i++) n += Math.max(1, Math.ceil(textBoxEm_(lines[i]) * pt / w));
  return n;
}
// 1行の高さ[pt]（行間の設定込み）。これまでのツールの見積もり（行間115%で文字の大きさの1.2倍）より、
// 少しだけ大きめに見る（游ゴシックなど、1行の高さが少し大きい書体でも下のまとまりに重ならないように）。
// PowerPointは開いただけでは「はみ出したら縮小」をやり直さないので、見積もりが小さすぎると下に重なる。
function premtgPitch_(pt, lnSpc) { return pt * 1.15 * lnSpc; }

// いまの文字の大きさで、枠の高さに収まるか
function premtgFits_(xml, id) {
  var m = id ? premtgBoxMetrics_(xml, id) : null;
  if (!m) return true;
  return premtgLineCount_(m.lines, m.basePt, m.widthPt) * premtgPitch_(m.basePt, m.lnSpc) <= m.heightPt + 0.5;
}

// 1行に収まる大きさにする（役職名・お名前）
function premtgFitLine_(xml, id, minPt) {
  var m = id ? premtgBoxMetrics_(xml, id) : null;
  if (!m) return xml;
  var em = 0;
  for (var i = 0; i < m.lines.length; i++) em = Math.max(em, textBoxEm_(m.lines[i]));
  if (!em || em * m.basePt <= m.widthPt * 0.95) return xml;
  return premtgSetSize_(xml, id, Math.max(minPt, Math.floor(m.widthPt * 0.95 / em * 2) / 2));
}

// 枠の高さに収まる大きさにする（共有事項など）。収まらなければ下限の大きさ
function premtgFitBody_(xml, id, minPt) {
  var m = id ? premtgBoxMetrics_(xml, id) : null;
  if (!m || m.basePt <= minPt) return xml;
  var pt = m.basePt;
  while (pt > minPt && premtgLineCount_(m.lines, pt, m.widthPt) * premtgPitch_(pt, m.lnSpc) > m.heightPt) pt -= 0.5;
  pt = Math.max(minPt, pt);
  return pt === m.basePt ? xml : premtgSetSize_(xml, id, pt);
}

// まとめのページ：3つのまとまり（見出し＋本文）が上から順に並んでいれば、
// 本文の量に合わせて枠の高さを決め、上から詰めて並べ直す（ツールと同じ。量が多い週に下のまとまりと重ならないように）。
// 文字の大きさは3つとも同じにし、全部が収まる大きさまで小さくする。
// 並び方が違うひな形（横に並べたなど）は、枠はそのままで、それぞれの枠に収まる大きさにするだけ。
function premtgLayoutBlocks_(xml, ids, info) {
  var blocks = [], i;
  for (i = 0; i < ids.length; i++) {
    var m = premtgBoxMetrics_(xml, ids[i]);
    if (m) blocks.push({ id: ids[i], m: m, label: null });
  }
  if (!blocks.length) return xml;
  blocks.sort(function (a, b) { return a.m.g.y - b.m.g.y; });
  // 見出し：本文の枠のすぐ上（0.5インチ以内）にある、左右が重なる文字箱
  var mine = {};
  for (i = 0; i < blocks.length; i++) mine[blocks[i].id] = true;
  var sps = premtgShapes_(xml, 'p:sp').filter(function (s) { return !mine[s.id] && slideText_(s.seg).trim(); });
  for (i = 0; i < blocks.length; i++) {
    var g = blocks[i].m.g, best = null, bd = 457200;
    for (var j = 0; j < sps.length; j++) {
      var s = sps[j], bottom = s.y + s.cy, d = g.y - bottom;
      if (d < -12700 || d > bd) continue;
      if (s.x >= g.x + g.cx || s.x + s.cx <= g.x) continue;
      best = s; bd = d;
    }
    blocks[i].label = best;
  }
  var topOf = function (b) { return b.label ? Math.min(b.label.y, b.m.g.y) : b.m.g.y; };
  var stacked = true;
  for (i = 0; i + 1 < blocks.length; i++) {
    var a = blocks[i].m.g, n = blocks[i + 1];
    if (a.y + a.cy > topOf(n) + 12700) stacked = false;                 // 下のまとまりに重なっている
    if (a.x >= n.m.g.x + n.m.g.cx || a.x + a.cx <= n.m.g.x) stacked = false;   // 横に並んでいる
  }
  if (!stacked) {
    for (i = 0; i < blocks.length; i++) {
      xml = premtgFitBody_(xml, blocks[i].id, PREMTG_MIN_PT_);
      if (info && !premtgFits_(xml, blocks[i].id)) info.overflow = true;
    }
    return xml;
  }
  // 使える高さ：いちばん上の見出しから、いちばん下の本文の枠の下まで
  var top = topOf(blocks[0]), last = blocks[blocks.length - 1].m.g, bottom = last.y + last.cy;
  var fixed = 0, gaps = [];
  for (i = 0; i < blocks.length; i++) {
    var b = blocks[i], gl = b.label ? Math.max(0, b.m.g.y - (b.label.y + b.label.cy)) : 0;
    var gn = i + 1 < blocks.length ? Math.max(0, topOf(blocks[i + 1]) - (b.m.g.y + b.m.g.cy)) : 0;
    gaps.push({ label: gl, next: gn });
    fixed += (b.label ? b.label.cy : 0) + gl + gn;
  }
  var room = bottom - top - fixed;
  var need = function (b, pt) {
    var lines = premtgLineCount_(b.m.lines, pt, b.m.widthPt);
    return Math.ceil(lines * premtgPitch_(pt, b.m.lnSpc) * 12700 + b.m.insV);
  };
  var pt = 0;
  for (i = 0; i < blocks.length; i++) pt = Math.max(pt, blocks[i].m.basePt);
  var total;
  for (;;) {
    total = 0;
    for (i = 0; i < blocks.length; i++) total += need(blocks[i], pt);
    if (total <= room || pt <= PREMTG_MIN_PT_) break;
    pt -= 0.5;
  }
  var scale = total > room ? room / total : 1;                          // 下限でも収まらないときは枠を按分
  if (info && scale < 1) info.overflow = true;
  var y = top;
  for (i = 0; i < blocks.length; i++) {
    var bl = blocks[i], h = Math.floor(need(bl, pt) * scale);
    if (bl.label) {
      xml = setShapeGeomEmu_(xml, bl.label.id, { y: y });
      y += bl.label.cy + gaps[i].label;
    }
    xml = setShapeGeomEmu_(xml, bl.id, { y: y, cy: h });
    xml = premtgSetSize_(xml, bl.id, pt);
    y += h + gaps[i].next;
  }
  return xml;
}

// --- パワポを組み立てる（Driveの読み書きから切り離してある。tools/check_premtg.js で確かめる）---
//   parts … ひな形のpptxを展開したもの（書き換える）
//   data  … premtgData_ の結果
//   cache … 写真の控え（mpAddPhoto_ と同じもの { by: {}, seq: 0 }）
function buildPreMeetingDeck_(parts, data, cache) {
  var models = premtgModels_(parts), info = { summary: false, pages: [], noPhoto: [], overflow: [], messages: [] };
  var w = ['日', '月', '火', '水', '木', '金', '土'], d = parseDate_(data.date);
  var common = { '月日': data.md, '開催回': data.meetingNo || '',
                 '開催日': d ? d.getFullYear() + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日（' + w[d.getDay()] + '）' : '' };

  // まとめのページ
  if (models.summary) {
    var sx = xmlOf_(parts, models.summary), ids = [];
    for (var b = 0; b < PREMTG_BLOCKS_.length; b++) {
      var name = PREMTG_BLOCKS_[b], id = premtgShapeWith_(sx, '{{' + name + '}}');
      var lines = (data.summary[name] || []).length ? data.summary[name] : ['（記載なし）'];
      sx = premtgExpandToken_(sx, name, lines);
      if (id) ids.push(id);
    }
    var lay = {};
    sx = premtgLayoutBlocks_(replaceTokensInXml_(sx, common), ids, lay);
    putXml_(parts, models.summary, sx);
    info.summary = true;
    if (lay.overflow) info.messages.push('まとめの文字が多く、いちばん小さくしても枠に収まりません。パワポで確かめてください。');
  } else {
    info.messages.push('まとめのページ（{{定例会関連}} などのあるページ）がひな形に無いので、まとめは作りませんでした。');
  }

  // 役職のページ（ひな形の役職のページの場所に、作ったページを並べる。ひな形のページは消す）
  var pages = premtgPages_(data.roles, !!models.two), added = [];
  if (!models.two && !models.one) {
    if (pages.length) info.messages.push('役職のページのひな形（{{共有事項1}} のあるページ）が無いので、役職のページは作りませんでした。');
  } else {
    var photoTypes = {};
    for (var p = 0; p < pages.length; p++) {
      var model = (pages[p].length === 2 || !models.one) ? models.two : models.one;
      var slots = model === models.two ? 2 : 1;
      var rels = (xmlOf_(parts, relsPathOf_(model)) || '').replace(/<Relationship\b[^>]*notesSlides\/[^>]*\/>/g, '');
      var r = premtgRoleSlide_(parts, xmlOf_(parts, model), rels, pages[p], slots, common, cache);
      info.noPhoto = info.noPhoto.concat(r.noPhoto);
      info.overflow = info.overflow.concat(r.overflow);
      for (var t in r.types) photoTypes[t] = true;
      added.push(addSlidePart_(parts, r.xml, r.rels));
      info.pages.push(pages[p].map(function (x) { return x.label; }));
    }
    var ct = xmlOf_(parts, '[Content_Types].xml');
    for (var ext in photoTypes) ct = ensureDefaultType_(ct, ext);
    putXml_(parts, '[Content_Types].xml', ct);
    var entries = slideEntries_(parts), out = [], placed = false;
    for (var e = 0; e < entries.length; e++) {
      var path = entries[e].path;
      if (path === models.two || path === models.one) {
        if (!placed) { out = out.concat(added); placed = true; }
        continue;
      }
      out.push(entries[e]);
    }
    setSlideEntries_(parts, out);
    if (models.two) premtgDropSlide_(parts, models.two);
    if (models.one) premtgDropSlide_(parts, models.one);
  }

  // ほかのページ（ひな形に足したページ）の {{月日}} なども
  var order = slideOrder_(parts);
  for (var o = 0; o < order.length; o++) {
    var x = xmlOf_(parts, order[o]);
    if (x && x.indexOf('{{') >= 0) putXml_(parts, order[o], replaceTokensInXml_(x, common));
  }
  mpPruneMedia_(parts);
  return info;
}

// 役職のページを1枚作る（slots … ひな形の人数。2人のページに1人だけのときは、右側を空ける）
function premtgRoleSlide_(parts, xml, rels, people, slots, common, cache) {
  var vals = {}, noPhoto = [], types = {}, fit = [], overflow = [], k;
  for (k in common) vals[k] = common[k];
  for (k = 1; k <= slots; k++) {
    var p = people[k - 1] || null;
    var roleId = premtgShapeWith_(xml, '{{役職' + k + '}}') || premtgShapeWith_(xml, '{{アイコン' + k + '}}');
    var nameId = premtgShapeWith_(xml, '{{氏名' + k + '}}');
    var bodyId = premtgShapeWith_(xml, '{{共有事項' + k + '}}');
    var picId = premtgPicFor_(xml, k, nameId || roleId);
    var bandId = premtgBandFor_(xml, k, roleId || nameId, !p);
    if (!p) {                                                     // 空ける側：写真と帯を消し、文字は空に
      if (picId) xml = removeShape_(xml, picId);
      if (bandId) xml = removeShape_(xml, bandId);
      vals['アイコン' + k] = vals['役職' + k] = vals['氏名' + k] = '';
      xml = premtgExpandToken_(xml, '共有事項' + k, ['']);
      continue;
    }
    vals['アイコン' + k] = p.icon;
    vals['役職' + k] = p.label;
    vals['氏名' + k] = p.holder ? p.holder + 'さん' : '';
    xml = premtgExpandToken_(xml, '共有事項' + k, p.text.split('\n'));
    if (bandId) xml = premtgRecolor_(xml, bandId, p.color);
    if (picId) {
      var photo = null;
      try { photo = p.holder ? mpAddPhoto_(parts, cache, p.holder) : null; }
      catch (e) { console.warn('[PREMTG] 写真を読めませんでした: ' + p.holder + ' ' + (e && e.message ? e.message : e)); }
      if (photo) {
        var set = setPicImage_(xml, rels, picId, '../media/' + photo.path.replace('ppt/media/', ''));
        xml = set.xml; rels = set.rels;
        types[photo.path.replace(/^.*\./, '')] = true;
        var box = readShapeGeomEmu_(xml, picId);
        if (box && photo.width && photo.height) xml = setSrcRectInPic_(xml, picId, coverCrop_(photo.width, photo.height, box.cx, box.cy));
      } else if (p.holder) {
        noPhoto.push(p.holder);
      }
    }
    fit.push({ role: roleId, name: nameId, body: bodyId, label: p.label });
  }
  xml = replaceTokensInXml_(xml, vals);
  for (var f = 0; f < fit.length; f++) {
    xml = premtgFitLine_(xml, fit[f].role, 10);
    xml = premtgFitLine_(xml, fit[f].name, 9);
    xml = premtgFitBody_(xml, fit[f].body, PREMTG_MIN_PT_);
    if (!premtgFits_(xml, fit[f].body)) overflow.push(fit[f].label);
  }
  return { xml: xml, rels: rels, noPhoto: noPhoto, types: types, overflow: overflow };
}

// ひな形のページを消す（並びからは setSlideEntries_ で外してあるもの）
function premtgDropSlide_(parts, path) {
  var rels = xmlOf_(parts, relsPathOf_(path)) || '';
  var notes = rels.match(/Target="\.\.\/notesSlides\/(notesSlide\d+\.xml)"/);
  if (notes) {
    var np = 'ppt/notesSlides/' + notes[1];
    delete parts[np];
    delete parts['ppt/notesSlides/_rels/' + notes[1] + '.rels'];
  }
  delete parts[path];
  delete parts[relsPathOf_(path)];
  var file = path.replace(/^.*\//, '');
  var prsRels = xmlOf_(parts, 'ppt/_rels/presentation.xml.rels');
  putXml_(parts, 'ppt/_rels/presentation.xml.rels',
    prsRels.replace(new RegExp('<Relationship\\b[^>]*Target="slides/' + file.replace('.', '\\.') + '"[^>]*/>'), ''));
  var ct = xmlOf_(parts, '[Content_Types].xml');
  ct = ct.replace(new RegExp('<Override PartName="/' + path.replace(/[.\/]/g, '\\$&') + '"[^>]*/>'), '');
  if (notes) ct = ct.replace(new RegExp('<Override PartName="/ppt/notesSlides/' + notes[1].replace('.', '\\.') + '"[^>]*/>'), '');
  putXml_(parts, '[Content_Types].xml', ct);
}

// --- 画面から呼ぶ ---
// 作る前に中身を確かめる（まとめの文・役職のページの割り付け・空欄の項目・使うひな形）
function getPreMeetingPreview(dateStr) {
  try {
    routineResetCache_();
    var target = parseDate_(dateStr);
    if (!target) return { ok: false, message: '開催日が分かりません。' };
    var data = premtgData_(target);
    if (!data.ok) return data;
    var tpl = premtgTemplateInfo_(), models = null, warn = [];
    if (tpl.error) warn.push(tpl.error);
    else {
      try { models = premtgModels_(premtgTemplateParts_(tpl)); }
      catch (e) { warn.push('ひな形を開けませんでした: ' + (e && e.message ? e.message : e)); }
    }
    if (models && !models.summary) warn.push('ひな形にまとめのページ（{{定例会関連}} などのあるページ）がありません。');
    if (models && !models.two && !models.one) warn.push('ひな形に役職のページ（{{共有事項1}} のあるページ）がありません。');
    var pages = premtgPages_(data.roles, models ? !!models.two : true).map(function (pg) {
      return pg.map(function (r) { return { label: r.label, icon: r.icon, holder: r.holder, chars: r.text.length, color: r.color }; });
    });
    return { ok: true, date: data.date, display: data.display, meetingNo: data.meetingNo, summary: data.summary,
             pages: pages, blank: data.blank,
             skipped: data.roles.filter(function (r) { return !r.text; }).map(function (r) { return r.label; }),
             template: { registered: tpl.registered, name: tpl.name, error: warn.join('\n') } };
  } catch (e) {
    console.error('[PREMTG] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 事前MTG（朝イチMTG）のパワポを作って、03_生成物 に保存する
function generatePreMeetingSlides(dateStr) {
  try {
    routineResetCache_();
    var target = parseDate_(dateStr);
    if (!target) return { ok: false, message: '開催日が分かりません。' };
    var data = premtgData_(target);
    if (!data.ok) return data;
    var tpl = premtgTemplateInfo_();
    var parts = premtgTemplateParts_(tpl);
    var info = buildPreMeetingDeck_(parts, data, { by: {}, seq: 0 });
    var outName = Utilities.formatDate(target, 'Asia/Tokyo', 'yyyyMMdd') + '_BNI事前MTG.pptx';
    var saved = saveOutputFile_(zipFromMap_(parts, outName), outName);

    var msg = data.display + ' の事前MTG（朝イチMTG）のパワポを作りました（'
      + (info.summary ? 'まとめ1枚＋' : '') + '役職のページ ' + info.pages.length + '枚）。';
    var skipped = data.roles.filter(function (r) { return !r.text; }).map(function (r) { return r.label; });
    if (skipped.length) msg += '\n今週の共有事項が無いので、ページを作らなかった役職: ' + skipped.join('、');
    if (data.blank.length) msg += '\n空欄の項目（「—」と出ています）: ' + data.blank.map(function (b) { return b.label; }).join('、');
    if (info.noPhoto.length) msg += '\n写真が見つからない方（仮の画像のまま）: ' + info.noPhoto.join('、');
    if (info.overflow.length) msg += '\n共有事項が長く、いちばん小さくしても枠に収まらない役職: ' + info.overflow.join('、') + '（パワポで確かめてください）';
    if (info.messages.length) msg += '\n' + info.messages.join('\n');
    msg += '\nひな形: ' + (tpl.registered ? '登録したもの（' + tpl.name + '）' : '既定のもの');
    return { ok: true, message: msg, url: saved.url, downloadUrl: saved.downloadUrl, fileName: outName,
             pages: info.pages.length + (info.summary ? 1 : 0) };
  } catch (e) {
    console.error('[PREMTG] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '事前MTGのパワポを作れませんでした: ' + (e && e.message ? e.message : e) };
  }
}
