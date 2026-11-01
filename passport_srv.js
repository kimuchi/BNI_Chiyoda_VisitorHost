// === パスポートプログラム卒業コメント（後半スライド）===
//
// 事前MTGの書記兼会計の入力「パスポートプログラム卒業コメント」で選んだメンバー（後半スライドの画面で選び直せる）ごとに、
// 登録したひな形（⚙️ 設定 ＞ 大きなスライド ＞ パスポートプログラム卒業コメント）のページを写して、
// 後半スライドの推薦のことばのあと（無ければ最後）に入れる（meeting_slides_srv.js）。
// 入れたお名前・カテゴリーの文字は、書体をメイリオにし、1行に収まる大きさにして折り返さない
// （ひな形の書体のままだと、メイリオの見積もりより広く、変なところで折り返していた）。
// ひな形の目印：「苗字 名前」（または {{氏名}}）、「【カテゴリー】」（または {{カテゴリー}}）、
// 「メンバー写真」とだけ書いた図形（または「お写真」・名前が「写真」の画像）。
// ページの作り方（写真・字の大きさ・ページの大きさ合わせ）は熱烈歓迎と同じ（welcome_srv.js の premtgWelcomePages_）。

var PASSPORT_KIND_ = 'passport';
var PASSPORT_ITEM_ = 'パスポートプログラム卒業コメント';

// 入力欄の値（「、」区切りのお名前）→ 名簿の方 [{ name, company, category, kana }]（名簿に無いお名前は category 空で入れる）
function passportMembersOf_(value) {
  var v = String(value || '').trim();
  if (!v || /^(なし|無し|ー|－|-)$/.test(v.normalize('NFKC').replace(/\s/g, ''))) return [];
  var roster = [];
  try { roster = getMemberMaster({ membersOnly: true }).members || []; } catch (e) {}
  var out = [], seen = {};
  v.split(/[、,，\n]+/).forEach(function (raw) {
    var name = raw.trim().replace(/(さん|様|さま|氏)$/, '').trim();
    if (!name || seen[normName_(name)]) return;
    seen[normName_(name)] = true;
    var m = roster.filter(function (x) { return normName_(x.name) === normName_(name); })[0] || null;
    out.push({ name: m ? m.name : name, company: m && m.company ? CB_.normalize(m.company) : '',
               category: m ? roleInBrackets_(m.title || '') : '', kana: m ? String(m.kana || '').trim() : '' });
  });
  return out;
}

// 開催日の入力から、卒業コメントのメンバー
function passportMembers_(date) {
  var ctx = roleBuildContext_(date, '');
  if (!ctx.found) return [];
  var it = premtgExtraItem_(ctx, PASSPORT_ITEM_);
  return passportMembersOf_(it ? premtgText_(it.value) : '');
}

// ひな形の目印を、熱烈歓迎と同じ差し込み口・目印にそろえる
function passportPrepare_(src) {
  var rules = [
    { re: /苗字[\s　]*名前/, value: function () { return '{{氏名}}'; } },
    { re: /【カテゴリー】/, value: function () { return '{{カテゴリー}}'; } },
    { re: /^[\s　]*メンバー写真[\s　]*$/, value: function () { return 'お写真'; } }
  ];
  slideOrder_(src).forEach(function (p) {
    var r = replacePatternsInXml_(xmlOf_(src, p) || '', rules);
    if (r.changed) putXml_(src, p, r.xml);
  });
  return src;
}

// 後半スライドに入れる。afterPath … このページのあと（無ければ最後）
//   戻り値 { pages, paths, noPhoto, message }
function passportPages_(parts, src, members, cache, afterPath) {
  if (!members || !members.length) return null;
  var entries = slideEntries_(parts);
  var after = afterPath || (entries.length ? entries[entries.length - 1].path : '');
  var r = premtgWelcomePages_(parts, passportPrepare_(src), members, {}, cache, after);
  for (var i = 0; i < r.paths.length; i++) {
    var m = members.filter(function (x) { return x.name === r.pages[i]; })[0] || {}, xml = xmlOf_(parts, r.paths[i]);
    [m.name, m.category].forEach(function (t) {
      var id = t ? premtgShapeWith_(xml, t) : '';
      if (id) xml = passportOneLine_(xml, id, t);
    });
    putXml_(parts, r.paths[i], xml);
  }
  var msgs = r.messages.map(function (s) {
    return s.replace(/熱烈歓迎のひな形/g, 'パスポートプログラム卒業コメントのひな形');
  });
  if (r.pages.length) msgs.unshift('パスポートプログラム卒業コメントのページを入れました（' + r.pages.join('、') + '）。');
  if (r.noPhoto.length) msgs.push('⚠ 写真が見つからなかった方（写真なしで作りました）: ' + r.noPhoto.join('、'));
  return { pages: r.pages, paths: r.paths, noPhoto: r.noPhoto, message: msgs.join('\n') };
}

// 後半スライドの画面用：その開催日に選んであった方（書記兼会計の入力）と、選べる方（入会日から4か月以内の方・名簿の全員）
function getPassportChoices(dateVal) {
  try {
    var target = parseDate_(dateVal);
    if (!target) return { ok: false, message: '開催日が分かりません。', selected: [], recent: [], all: [] };
    var roster = getMemberMaster({ membersOnly: true }).members || [];
    var selected = [];
    try { selected = passportMembers_(target).map(function (m) { return m.name; }); } catch (e) {}
    return { ok: true, selected: selected, recent: roleRecentMembers_(roster, target),
             all: roster.map(function (m) { return m.name; }) };
  } catch (e) {
    return { ok: false, message: '卒業コメントの方を読めませんでした: ' + (e && e.message ? e.message : e), selected: [], recent: [], all: [] };
  }
}

// 入れた文字の枠：書体をメイリオにし、枠の幅（左右の余白を引いた幅の9割）に1行で入る大きさ（もとの大きさまで・10ptまで）にして、
// 折り返さない（wrap="none"。まん中寄せのまま枠の中に収まる）
function passportOneLine_(xml, id, text) {
  var m = premtgBoxMetrics_(xml, id);
  if (m) {
    var em = textBoxEm_(String(text || '')) || 1;
    xml = premtgSetSize_(xml, id, Math.max(10, Math.min(m.basePt, Math.floor(m.widthPt * 0.9 / em * 2) / 2)));
  }
  var r = findShapeRange_(xml, id);
  if (!r) return xml;
  var font = '<a:latin typeface="' + PREMTG_FONT_ + '"/><a:ea typeface="' + PREMTG_FONT_ + '"/><a:cs typeface="' + PREMTG_FONT_ + '"/>';
  var seg = xml.substring(r.start, r.end)
    .replace(/<a:(latin|ea|cs|sym)\b[^>]*\/>/g, '')   // 記号の書体（sym）も外す（残すと書体の指定がその後ろになり、PowerPoint が修復を求める）
    .replace(/<a:(rPr|endParaRPr)\b([^>]*?)\/>/g, '<a:$1$2>' + font + '</a:$1>')
    .replace(/<a:bodyPr\b([^>]*?)(\/?)>/, function (all, attrs, close) {
      return '<a:bodyPr' + (/\swrap="/.test(attrs) ? attrs.replace(/\swrap="\w+"/, ' wrap="none"') : attrs + ' wrap="none"') + close + '>';
    });
  // 中身のある rPr（<a:rPr …>…</a:rPr>）にも書体を入れる（塗りなどのあと、決まった順の位置に）
  seg = seg.replace(/(<a:(rPr|endParaRPr)\b[^>]*[^\/]>)([\s\S]*?)(<\/a:\2>)/g, function (all, open, tag, body, close) {
    if (body.indexOf('typeface="' + PREMTG_FONT_ + '"') >= 0) return all;
    var at = body.search(/<a:(hlinkClick|hlinkMouseOver|rtl|extLst)\b/);
    return at < 0 ? open + body + font + close : open + body.slice(0, at) + font + body.slice(at) + close;
  });
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}
