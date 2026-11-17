// === パワーチームが求める専門分野（前半スライド）===
//
// 前半の「チャプターが求める専門分野」のページのすぐあとに、パワーチームごとに1枚、
// 「〇〇パワーチームが求める専門分野」のページを入れる。
// ひな形は ⚙️ 設定 ＞ 大きなスライド ＞ パワーチームが求める専門分野 に登録する（このシステムには同梱しない）。
// ひな形の目印：「〇〇パワーチーム」（チーム名が入る）と、「カテゴリー」とだけ書いた欄（上から順に専門分野が入る。いくつでも）。
// パワーチームは、チームごとに名前と専門分野を覚えておく（いくつでも）。回ごとに、入れるチームを0個以上選ぶ
// （選んだかどうかも覚えておき、次の回の初めの値にする）。

var POWERTEAM_KIND_ = 'powerteam';
var POWERTEAM_PROP_ = 'BNI_POWER_TEAMS';         // 覚えているチーム [{ name, cats: [], on: 前の回に入れたか }]
var POWERTEAM_MAX_ = 20;                         // 覚えておくチームの数
var POWERTEAM_CATS_ = 10;                        // 1チームの専門分野の数（ひな形の欄の数までしか入らない）

function powerTeamsClean_(list) {
  var clean = function (s, n) { return String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').trim().slice(0, n); };
  return (list || []).map(function (t) {
    var name = clean(t && t.name, 40).replace(/パワーチーム$/, '').trim();
    var cats = ((t && t.cats) || []).map(function (c) { return clean(c, 60); }).filter(Boolean).slice(0, POWERTEAM_CATS_);
    return { name: name, cats: cats, on: !(t && t.on === false) };
  }).filter(function (t) { return t.name || t.cats.length; }).slice(0, POWERTEAM_MAX_);
}
// 画面用：覚えているチーム（選んだかどうかも。初めの値）
function powerTeamsSaved_() {
  try { return powerTeamsClean_(JSON.parse(PropertiesService.getScriptProperties().getProperty(POWERTEAM_PROP_) || '[]')); }
  catch (e) { return []; }
}
function powerTeamsRemember_(list) {
  try { PropertiesService.getScriptProperties().setProperty(POWERTEAM_PROP_, JSON.stringify(powerTeamsClean_(list))); } catch (e) {}
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
