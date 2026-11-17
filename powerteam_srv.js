// === パワーチームが求める専門分野（前半スライド）===
//
// 前半の「チャプターが求める専門分野」のページのすぐあとに、パワーチームごとに1枚、
// 「〇〇パワーチームが求める専門分野」のページを入れる。
// ひな形は ⚙️ 設定 ＞ 大きなスライド ＞ パワーチームが求める専門分野 に登録する（このシステムには同梱しない）。
// ひな形の目印：「〇〇パワーチーム」（チーム名が入る）と、「カテゴリー」とだけ書いた欄（上から順に専門分野が入る。いくつでも）。
// パワーチーム（チーム名と専門分野）は ⚙️ 設定 ＞ パワーチーム に登録しておき、回ごとにローテーションで入れる
// （1回に何チーム入れるか・どの回からどのチームか、を決めておくと、休会日を飛ばして順に回る）。
// 前半スライドの画面では、ローテーションで選んだチームにチェックが入っていて、その回だけ選び直せる。

var POWERTEAM_KIND_ = 'powerteam';
var POWERTEAM_PROP_ = 'BNI_POWER_TEAMS';         // { teams: [{ name, cats }], perMeeting, anchor: { date, index } }
var POWERTEAM_MAX_ = 20;                         // 登録できるチームの数
var POWERTEAM_CATS_ = 10;                        // 1チームの専門分野の数（ひな形の欄の数までしか入らない）

function powerTeamsClean_(list) {
  var clean = function (s, n) { return String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').trim().slice(0, n); };
  return (list || []).map(function (t) {
    var name = clean(t && t.name, 40).replace(/パワーチーム$/, '').trim();
    var cats = ((t && t.cats) || []).map(function (c) { return clean(c, 60); }).filter(Boolean).slice(0, POWERTEAM_CATS_);
    return { name: name, cats: cats, on: !(t && t.on === false) };
  }).filter(function (t) { return t.name || t.cats.length; }).slice(0, POWERTEAM_MAX_);
}

// 登録の内容（前の形 … チームの並びだけ、も読む）
function ptLoad_() {
  var raw = null;
  try { raw = JSON.parse(PropertiesService.getScriptProperties().getProperty(POWERTEAM_PROP_) || 'null'); } catch (e) { raw = null; }
  var st = Array.isArray(raw) ? { teams: raw } : (raw || {});
  var teams = powerTeamsClean_(st.teams || []).map(function (t) { return { name: t.name, cats: t.cats }; });
  var per = Math.max(1, Math.min(5, parseInt(st.perMeeting, 10) || 1));
  var a = st.anchor || {}, ad = parseDate_(a.date);
  return { teams: teams, perMeeting: per,
           anchor: { date: ad ? fmtDate_(ad) : '', index: Math.max(0, parseInt(a.index, 10) || 0) } };
}
function ptSave_(st) {
  PropertiesService.getScriptProperties().setProperty(POWERTEAM_PROP_, JSON.stringify(st));
}
// 起点の回から、その開催日までの回の数（休会日を除く。その日が起点より前なら負）
function ptMeetingsBetween_(from, to, holidays) {
  var sign = 1, a = new Date(from.getFullYear(), from.getMonth(), from.getDate()), b = new Date(to.getFullYear(), to.getMonth(), to.getDate());
  if (b < a) { var x = a; a = b; b = x; sign = -1; }
  var n = 0;
  for (var d = new Date(a.getTime()); d < b; d.setDate(d.getDate() + 7)) if (holidays.indexOf(fmtDate_(d)) < 0) n++;
  return sign * n;
}
// その開催日にローテーションで入れるチームの番号（休会日は無し）
function ptPicks_(st, d, holidays) {
  var n = st.teams.length;
  if (!n || !d || holidays.indexOf(fmtDate_(d)) >= 0) return [];
  var ad = parseDate_(st.anchor.date) || d;
  var k = ptMeetingsBetween_(ad, d, holidays), per = Math.min(st.perMeeting, n), out = [];
  var start = ((st.anchor.index + k * per) % n + n) % n;
  for (var i = 0; i < per; i++) out.push((start + i) % n);
  return out;
}
function ptHolidays_() { try { return getHolidays(); } catch (e) { return []; } }
// その開催日のチーム一覧（on … ローテーションで入れるチーム）
function powerTeamsForDate_(d) {
  var st = ptLoad_(), picks = ptPicks_(st, d, ptHolidays_());
  return st.teams.map(function (t, i) { return { name: t.name, cats: t.cats, on: picks.indexOf(i) >= 0 }; });
}
// 前半スライドの画面から（開催日を選び直したとき）
function getPowerTeamsForDate(dateVal) {
  try { return { ok: true, teams: powerTeamsForDate_(parseDate_(dateVal)) }; }
  catch (e) { return { ok: false, message: 'パワーチームを読めませんでした: ' + (e && e.message ? e.message : e), teams: [] }; }
}

// ⚙️ 設定 ＞ パワーチーム の画面：登録の内容と、これからの予定（次の開催日から12回）
function getPowerTeamSettings() {
  try {
    requireSheetAccess_();
    var st = ptLoad_(), holidays = ptHolidays_(), next = rotNextMeeting_(holidays);
    if (!st.anchor.date) st.anchor = { date: fmtDate_(next), index: 0 };
    var w = ['日', '月', '火', '水', '木', '金', '土'];
    var dates = rotMeetingDates_(next, 12, holidays);
    return { ok: true, teams: st.teams, perMeeting: st.perMeeting, anchor: st.anchor,
             dates: dates.map(function (d) { return { value: fmtDate_(d), md: (d.getMonth() + 1) + '/' + d.getDate() + '(' + w[d.getDay()] + ')' }; }),
             schedule: dates.map(function (d) {
               return { date: fmtDate_(d), md: (d.getMonth() + 1) + '/' + d.getDate() + '(' + w[d.getDay()] + ')',
                        teams: ptPicks_(st, d, holidays).map(function (i) { return st.teams[i].name; }) };
             }) };
  } catch (e) {
    console.error('[POWERTEAM] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'パワーチームを読み込めませんでした: ' + (e && e.message ? e.message : e) };
  }
}
// data … { teams: [{ name, cats }], perMeeting, anchor: { date, index } }（anchor … その開催日に、その番号のチームから）
function savePowerTeamSettings(data) {
  try {
    requireSheetAccess_();
    var teams = powerTeamsClean_((data && data.teams) || []).map(function (t) { return { name: t.name, cats: t.cats }; });
    var noName = teams.filter(function (t) { return !t.name; }).length;
    if (noName) return { ok: false, message: 'チーム名が空のチームがあります（' + noName + 'チーム）。チーム名を入れてください。' };
    var a = (data && data.anchor) || {}, ad = parseDate_(a.date);
    var st = { teams: teams, perMeeting: Math.max(1, Math.min(5, parseInt(data && data.perMeeting, 10) || 1)),
               anchor: { date: ad ? fmtDate_(ad) : '', index: Math.max(0, Math.min(teams.length - 1, parseInt(a.index, 10) || 0)) } };
    ptSave_(st);
    var r = getPowerTeamSettings();
    r.message = 'パワーチームを保存しました（' + teams.length + 'チーム・1回に' + st.perMeeting + 'チーム）。';
    return r;
  } catch (e) {
    return { ok: false, message: '保存できませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// ひな形の中の、パワーチームのページ（「パワーチーム」の文字があるページ。無ければ最初のページ）
function powerTeamModel_(src) {
  var order = slideOrder_(src);
  for (var i = 0; i < order.length; i++) if (slideText_(xmlOf_(src, order[i]) || '').indexOf('パワーチーム') >= 0) return order[i];
  return order[0] || '';
}
// ひな形の「カテゴリー」の欄の数
function powerTeamSlots_(xml) {
  var n = 0;
  findTagRanges_(xml, 'a:p').forEach(function (r) { if (/^[\s　]*カテゴリー[\s　]*$/.test(slideText_(xml.substring(r.start, r.end)))) n++; });
  return n;
}
// 1枚に、チーム名と専門分野を入れる（専門分野は「カテゴリー」の欄に上から順に。足りない欄は空にする）
function powerTeamFill_(xml, team) {
  var name = team.name ? team.name + 'パワーチーム' : '';
  xml = replacePatternsInXml_(xml, [{ re: /[〇○◯]{1,4}パワーチーム/, value: function () { return name || '〇〇パワーチーム'; } }]).xml;
  // replacePatternsInXml_ は段落を後ろから見るので、後ろの欄から順に入れる
  var total = powerTeamSlots_(xml), k = total;
  return replacePatternsInXml_(xml, [{ re: /^[\s　]*カテゴリー[\s　]*$/, value: function () { k--; return team.cats[k] || ''; } }]).xml;
}

// 前半のテンプレートの「チャプターが求める専門分野」のページ（差し込み口 {{求める専門分野1}}、または「求める専門分野」の文字）
function chapterWantedPage_(parts) {
  var p = findSlideWithText_(parts, '{{求める専門分野1}}');
  if (p) return p;
  var order = slideOrder_(parts);
  for (var i = 0; i < order.length; i++) {
    var t = slideText_(xmlOf_(parts, order[i]) || '').replace(/[\s　]/g, '');
    if (t.indexOf('チャプターが求める') >= 0 && t.indexOf('専門分野') >= 0) return order[i];
  }
  return '';
}

// 前半スライドに入れる。teams … [{ name, cats }]。src … ひな形を展開したもの
//   戻り値 { pages, paths, message }
function insertPowerTeamPages_(parts, src, teams) {
  teams = powerTeamsClean_(teams).filter(function (t) { return t.on; });
  if (!teams.length) return null;
  var after = chapterWantedPage_(parts);
  var model = src ? powerTeamModel_(src) : '';
  if (!model) return { message: '⚠ パワーチームのひな形にページがありません。' };
  var msgs = [], fit = fitSourceToTargetSize_(parts, src);
  if (fit.scaled) msgs.push('パワーチームのひな形のページの大きさを、前半スライドの大きさに合わせました（' + Math.round(fit.factor * 100) + '%）。');
  var slots = powerTeamSlots_(xmlOf_(src, model) || '');
  var imp = importSlideCopies_(parts, src, model, teams.length), over = [];
  for (var i = 0; i < teams.length; i++) {
    var path = imp.slides[i].path;
    putXml_(parts, path, powerTeamFill_(xmlOf_(parts, path), teams[i]));
    if (teams[i].cats.length > slots) over.push((teams[i].name || '（チーム名なし）') + '（' + teams[i].cats.length + '件のうち' + slots + '件）');
  }
  // 並び：「チャプターが求める専門分野」のページのすぐあと（無ければ最後）
  var rest = slideEntries_(parts).filter(function (e) { return !imp.slides.some(function (s) { return s.path === e.path; }); });
  var out = [], placed = false;
  for (var j = 0; j < rest.length; j++) {
    out.push(rest[j]);
    if (after && rest[j].path === after) { out = out.concat(imp.slides); placed = true; }
  }
  if (!placed) out = out.concat(imp.slides);
  setSlideEntries_(parts, out);
  syncSections_(parts);
  msgs.unshift('パワーチームが求める専門分野のページを入れました（' + teams.map(function (t) { return (t.name || '〇〇') + 'パワーチーム'; }).join('、') + '）'
    + (placed ? '。' : '。「チャプターが求める専門分野」のページが見つからないので、最後に入れました。'));
  if (!slots) msgs.push('⚠ パワーチームのひな形に「カテゴリー」とだけ書いた欄が無いので、専門分野は入っていません。');
  if (over.length) msgs.push('⚠ ひな形の欄より専門分野が多いので、入りきらなかったチームがあります: ' + over.join('、'));
  return { pages: teams.length, paths: imp.slides.map(function (s) { return s.path; }), message: msgs.join('\n') };
}
