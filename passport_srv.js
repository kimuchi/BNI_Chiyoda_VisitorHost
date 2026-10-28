// === パスポートプログラム卒業コメント（後半スライド）===
//
// 事前MTGの書記兼会計の入力「パスポートプログラム卒業コメント」で選んだメンバーごとに、
// 登録したひな形（⚙️ 設定 ＞ 大きなスライド ＞ パスポートプログラム卒業コメント）のページを写して、
// 後半スライドの推薦のことばのあと（無ければ最後）に入れる。
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
  var msgs = r.messages.map(function (s) {
    return s.replace(/熱烈歓迎のひな形/g, 'パスポートプログラム卒業コメントのひな形')
            .replace(/事前MTGのパワポ|事前MTGのひな形/g, '後半のスライド');
  });
  if (r.pages.length) msgs.unshift('パスポートプログラム卒業コメントのページを入れました（' + r.pages.join('、') + '）。');
  if (r.noPhoto.length) msgs.push('⚠ 写真が見つからなかった方（写真なしで作りました）: ' + r.noPhoto.join('、'));
  return { pages: r.pages, paths: r.paths, noPhoto: r.noPhoto, message: msgs.join('\n') };
}
