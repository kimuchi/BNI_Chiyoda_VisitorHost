// === チャプターの設定 ===
// 画面・メール・参加者一覧・割り振り表・メンバーブック・トークスクリプトに出すチャプターの名前と、
// チャプターごとに違う数え方（期の番号・定例会の回数と曜日）を、設定で変えられるようにする。
// もとは Activeチャプター用にコードへ書いていた値で、何も保存していなければその値（CHAPTER_DEFAULTS_）で動く。
//
// スクリプトのプロパティ BNI_CHAPTER = { name, region, termBase, meetingBaseDate, meetingBaseCount }
//   name             … チャプター名（「Active」。画面には「Activeチャプター」と出す）
//   region           … リージョン（ウェブアプリの下と、トークスクリプトの {リージョン} に出す。空でもよい）
//   termBase         … 2026年4月〜9月の期の番号。期は 4月〜9月・10月〜3月 の半期ごとに1つ進む
//                      （画面では「いまの期」で入れてもらい、ここに直して持つ）
//   meetingBaseDate  … 回数を数える基準の開催日。この曜日に毎週開催として数える
//   meetingBaseCount … その日の回数（休会日の週は数えない）
var CHAPTER_KEY_ = 'BNI_CHAPTER';
var CHAPTER_DEFAULTS_ = {
  name: 'Active', region: 'BNI東京千代田リージョン',
  termBase: 23,                                          // 2026年4月〜9月が23期
  meetingBaseDate: '2026/03/18', meetingBaseCount: 509   // 2026/3/18(水) が第509回
};
var CHAPTER_TERM_YEAR_ = 2026;                           // termBase は、この年の4月〜9月の期
var CHAPTER_WEEK_ = ['日', '月', '火', '水', '木', '金', '土'];
// 初回の準備（setup_srv.js）の記録。空のスプレッドシートから作ったとき mode: 'fresh'
var CHAPTER_SETUP_KEY_ = 'BNI_SETUP';
var CHAPTER_CACHE_ = null;

function openChapterSettingsDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('chapter_settings').setWidth(560).setHeight(640), 'チャプターの設定');
}

// 今の設定（保存していない項目は既定の値）。1回の実行の中では読み直さない
function chapterInfo_() {
  if (CHAPTER_CACHE_) return CHAPTER_CACHE_;
  var s = null;
  try { s = JSON.parse(PropertiesService.getScriptProperties().getProperty(CHAPTER_KEY_) || 'null'); } catch (e) { s = null; }
  if (!s || typeof s !== 'object') s = {};
  var D = CHAPTER_DEFAULTS_, term = parseInt(s.termBase, 10), count = parseInt(s.meetingBaseCount, 10);
  var date = chapterDate_(s.meetingBaseDate), name = chapterName_(s.name);
  CHAPTER_CACHE_ = {
    name: name || D.name,
    region: typeof s.region === 'string' ? chapterText_(s.region, 60) : D.region,
    termBase: (isFinite(term) && Math.abs(term) < 100000) ? term : D.termBase,
    meetingBaseDate: date ? chapterFmt_(date) : D.meetingBaseDate,
    meetingBaseCount: count > 0 ? count : D.meetingBaseCount,
    saved: !!name
  };
  return CHAPTER_CACHE_;
}

// 「Activeチャプター」
function chapterLabel_() { return chapterInfo_().name + 'チャプター'; }
// 「Activeチャプター 名簿システム」（ウェブアプリの題名・メニュー画面の見出し）
function chapterSystemTitle_() { return chapterLabel_() + ' 名簿システム'; }
// 「BNI東京千代田リージョン ｜ Activeチャプター」（ウェブアプリの下）
function chapterFooter_() {
  var c = chapterInfo_();
  return (c.region ? c.region + ' ｜ ' : '') + chapterLabel_();
}
// 2026年4月〜9月の期の番号
function chapterTermBase_() { return chapterInfo_().termBase; }
// 回数を数える基準 { date: 開催日（Date。呼ぶたびに新しく作る）, count: その日の回数 }
function chapterMeetingBase_() {
  var c = chapterInfo_();
  return { date: chapterDate_(c.meetingBaseDate), count: c.meetingBaseCount };
}
// 定例会の曜日（0=日〜6=土）と、その字（「水」）
function chapterWeekday_() { var d = chapterDate_(chapterInfo_().meetingBaseDate); return d ? d.getDay() : 3; }
function chapterWeekdayLabel_() { return CHAPTER_WEEK_[chapterWeekday_()]; }
// 空のスプレッドシートから始めたチャプターか（初回の準備で作ったとき）。
// そのときは、Activeチャプター用に入れてある初期値（役職の担当者など）を使わない
function chapterFresh_() {
  try {
    var s = JSON.parse(PropertiesService.getScriptProperties().getProperty(CHAPTER_SETUP_KEY_) || 'null');
    return !!(s && s.mode === 'fresh');
  } catch (e) { return false; }
}

// --- 画面から呼ぶ ---
function getChapterSettings() {
  try {
    var c = chapterInfo_(), now = roleTermOf_(new Date()), next = null;
    try { next = getMeetingCandidates()[0] || null; } catch (e) { next = null; }
    return {
      ok: true, name: c.name, region: c.region, label: chapterLabel_(), saved: c.saved, fresh: chapterFresh_(),
      term: now, termLabel: roleTermLabel_(now), nextTermLabel: roleTermLabel_(now + 1),
      meetingBaseDate: c.meetingBaseDate, meetingBaseCount: c.meetingBaseCount, weekday: chapterWeekdayLabel_(),
      next: next ? next.display : '',
      defaults: { name: CHAPTER_DEFAULTS_.name, region: CHAPTER_DEFAULTS_.region }
    };
  } catch (e) {
    console.error('[CHAPTER] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'チャプターの設定を読めませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// s … { name, region, term（いまの期）, meetingBaseDate, meetingBaseCount }
// 期の番号を付け直したときは、期ごとに保存してある役職・チームも同じだけずらす（中身が別の期に見えないように）
function saveChapterSettings(s) {
  try {
    s = s || {};
    var name = chapterName_(s.name);
    if (!name) return { ok: false, message: 'チャプター名を入れてください（例: Active）。' };
    var region = chapterText_(s.region, 60);
    var term = parseInt(chapterDigits_(s.term), 10);
    if (!(term > 0 && term < 1000)) return { ok: false, message: 'いまの期は 1〜999 の数字で入れてください。' };
    var d = chapterDate_(s.meetingBaseDate), count = parseInt(chapterDigits_(s.meetingBaseCount), 10);
    if (!d) return { ok: false, message: '基準の開催日を 2026/10/07 のように入れてください。' };
    if (!(count > 0 && count < 100000)) return { ok: false, message: '基準の開催日の回数は、1 以上の数字で入れてください。' };

    var old = chapterInfo_();
    var passed = roleTermOf_(new Date()) - old.termBase;         // 2026年4月〜9月から、いまの期まで何期進んだか
    var termBase = term - passed, delta = termBase - old.termBase;
    if (delta) roleShiftTerms_(delta);
    PropertiesService.getScriptProperties().setProperty(CHAPTER_KEY_, JSON.stringify({
      name: name, region: region, termBase: termBase,
      meetingBaseDate: chapterFmt_(d), meetingBaseCount: count
    }));
    CHAPTER_CACHE_ = null;
    var msg = 'チャプターの設定を保存しました（' + chapterLabel_() + '・いまの期 ' + term + '期・毎週' + chapterWeekdayLabel_() + '曜日）。';
    if (delta) msg += '\n期の番号を付け直したので、登録してある役職・チームの期も同じだけずらしました。';
    // 空のスプレッドシートから始めたとき（初回の準備。setup_srv.js）は、ルーティンチェックシートもここで作る
    // （期の番号と開催日が、この設定で決まるため）。期を付け直したら、作ったシートの名前の期もずらす
    if (chapterFresh_()) {
      try {
        var renamed = setupShiftRoutineNames_(delta), made = setupEnsureRoutineSheets_().filter(function (r) { return r.made; });
        if (renamed.length) msg += '\nルーティンチェックシートの名前の期もずらしました（' + renamed.join('、') + '）。';
        if (made.length) msg += '\nルーティンチェックシートを作りました: ' + made.map(function (r) { return '「' + r.name + '」'; }).join('、');
      } catch (e) {
        console.error('[CHAPTER] setup ' + (e && e.stack ? e.stack : e));
        msg += '\nルーティンチェックシートを作れませんでした（⚙️ 設定 > 足りないシートを作る で、もう一度お試しください）: '
          + (e && e.message ? e.message : e);
      }
    }
    return { ok: true, message: msg, settings: getChapterSettings() };
  } catch (e) {
    console.error('[CHAPTER] save ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '保存できませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// --- 小さな道具（Apps Script の Utilities を使わない。検査でもそのまま動くように）---
function chapterText_(v, max) {
  return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').replace(/[ 　]+/g, ' ').trim().slice(0, max);
}
// 「BNI Activeチャプター」「Active chapter」と入れられても「Active」にする
function chapterName_(v) {
  return chapterText_(v, 60).replace(/^BNI[\s　]*/i, '').replace(/[\s　]*(チャプター|chapter)$/i, '').trim().slice(0, 40);
}
function chapterDigits_(v) {
  return String(v == null ? '' : v).replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }).replace(/[^\d]/g, '');
}
// Date・「2026/10/07」「2026-10-07」「2026年10月7日」→ Date（その日の0時）
function chapterDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  var s = String(v == null ? '' : v).replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  var m = s.match(/(\d{4})\s*[\/\-年.]\s*(\d{1,2})\s*[\/\-月.]\s*(\d{1,2})/);
  if (!m) return null;
  var d = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  return (d.getMonth() === parseInt(m[2], 10) - 1) ? d : null;
}
function chapterFmt_(d) {
  var p = function (n) { return (n < 10 ? '0' : '') + n; };
  return d.getFullYear() + '/' + p(d.getMonth() + 1) + '/' + p(d.getDate());
}
