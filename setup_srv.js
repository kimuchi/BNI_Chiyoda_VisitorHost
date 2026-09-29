// === 初回の準備（空のスプレッドシートで初めて動かしたとき）===
// 新しくこのシステムを使い始めるチャプター向け。空のスプレッドシートにスクリプトを入れ、
// はじめて開いたとき（onOpen）に、システムが使うシートを作る。
//
//   開いたときに作る   … メンバー名簿・業種区分マスタ（非表示）・休会日（非表示）
//   チャプターの設定を保存したときに作る … 【NN期】ルーティンチェックシート
//       期の番号と開催日・回数は、チャプターの設定（chapter_srv.js）で決まるので、設定を待ってから作る。
//       項目（内容・担当・期日・曜日目安・備考）は、下の SETUP_ROUTINE_ITEMS_ の一般的なもの。
//       曜日目安は、定例会の曜日から数える（水曜開催なら「2日前」は「月曜まで」）。
//
// メニュー「⚙️ 設定 > 足りないシートを作る」でも、いつでも作れる（あるシートには触らない）。
// ルーティンチェックシートは、次の4回の定例会のうち、どのシートにも列が無い回の期のぶんを作る。
// 前の期のシートがあれば、それを写して開催日の列だけ新しくする（チャプターで足した項目や書式を引き継ぐ）。
//
// 初回かどうかは、スクリプトのプロパティ BNI_SETUP（chapter_srv.js の CHAPTER_SETUP_KEY_）で見る。
//   無い … まだ見ていない。シートがどれも空なら、空のスプレッドシートとして準備する（mode: 'fresh'）。
//          何か入っていれば、今まで使われてきたスプレッドシート（mode: 'existing'）として何も作らない。
var SETUP_HOLIDAY_SHEET_ = '休会日';
var SETUP_ROUTINE_DATE_COL_ = 10;       // 新しく作るチェックシートは、J列から開催日（A〜I列は項目名・担当・期日・曜日目安・備考）
var SETUP_ROUTINE_HEAD_ = ['', 'No.', '内容', '', '', '担当', '期日', '曜日目安', '備考'];
// 新しく作るチェックシートの項目。[C列, D列, E列（C＞D＞E の順に細かくなる）, 担当, 何日前（'' は期日なし）, 備考]
// 項目名は、スライド・役職ごとの入力・トークスクリプトが読む名前（routine_srv.js・talk_script_default.js）。
// 担当は role_input_srv.js の ROLE_DEFS_ の書き方。期日のある項目が、役職ごとの入力で必須になる。
var SETUP_ROUTINE_ITEMS_ = [
  ['スピーカーローテーション用スライド画像', '', '', '書記兼会計', 5, ''],
  ['メインプレゼン用スライド画像', '', '', '書記兼会計', 5, ''],
  ['定例会基本情報', 'メンバーシップから報告', '', 'バイス', 2, ''],
  ['', '遅刻・欠席担当', '', 'バイス', 2, ''],
  ['', '名札・バッチの注意', '', 'バイス', 2, ''],
  ['', 'リファーラルの注意', '', 'バイス', 2, ''],
  ['', '代理・欠席', '代理', 'バイス', '', '代理を立てるメンバーのお名前'],
  ['', '', '欠席', 'バイス', '', ''],
  ['', '', '医療欠席', 'バイス', '', ''],
  ['', 'リージョン参加者', '', 'プレジ', 3, '来られるリージョンの方（例: 〇〇ED）'],
  ['', '体験談', '', 'プレジ', 2, 'お名前（なければ「なし」）'],
  ['', 'エデュケーション', '', 'EC', 2, 'お名前'],
  ['', '', 'スライド', 'EC', 2, ''],
  ['', '審査中カテゴリー', '', 'バイス', 2, ''],
  ['', '新入会', '', 'バイス', 2, 'お名前（なければ「なし」）'],
  ['', '更新式(更新メンバー)', '', 'バイス', 2, 'お名前（なければ「なし」）'],
  ['', 'ウィークリープレゼン', '', 'プレジ', 2, '始まりの方（「業種区分　番号　お名前」。空欄なら前回から数えます）'],
  ['', '募集カテゴリー', '', 'バイス', 2, ''],
  ['', '更新対象者　30日前', '', 'バイス', 2, ''],
  ['', '更新対象者　60日前', '', 'バイス', 2, ''],
  ['', '更新対象者　90日前', '', 'バイス', 2, ''],
  ['', '退会者', '', 'バイス', 2, '発表した翌週にスライドから外す'],
  ['', '開放カテゴリー', '', 'バイス', 2, ''],
  ['', 'メインプレゼン', '', '書記兼会計', 5, '①〇〇さん　②〇〇さん'],
  ['', '', 'スライド', '書記兼会計', 5, ''],
  ['', 'ネットワーキングリーダー', '', 'バイス', 2, '月初のみ'],
  ['', '推薦の言葉', '', '書記兼会計', 2, '〇〇さん→〇〇さん（1組1行。アフターMTGで読む組は「アフター：」を前に付ける）'],
  ['', '', 'スライド', '書記兼会計', 2, ''],
  ['', 'スタートアッププレゼン', '', 'プレジ', '', '〇〇さん（1人まで）'],
  ['', 'その他注意事項', '', 'バイス', '', ''],
  ['割振表', '', '', 'VH', 1, 'ルーム・オリエン割り振り表'],
  ['ビジター情報', '', '', 'VH', 1, ''],
  ['ビジターの紹介', '', '', 'VH', 1, ''],
  ['ビジターフォロー交流会情報', '', '', 'VH', '', ''],
  ['朝一&アフターミーティング', 'プレジデントより', '', 'プレジ', 1, ''],
  ['', 'バイスプレジデントより', '', 'バイス', 1, ''],
  ['', '書記兼会計より', '', '書記兼会計', 1, ''],
  ['', '本日の招待者', '', 'WEB', '', 'ビジターを招待したメンバー'],
  ['', '各担当者とポジティブな挨拶', '', 'プレジ', '', ''],
  ['', 'その他のお知らせ', '', 'プレジ', 1, ''],
  ['', 'アフターMTG', '', 'プレジ', 1, ''],
  ['トークスクリプト追加事項', '', '', 'プレジ', 1, 'トークスクリプト（台本）に書き足すこと（どこで読むかも）'],
  ['BNI目的と概要', '', '', 'プレジ', 2, 'その月のコアバリュー（例: Givers Gain®）'],
  ['リマインダー・特別報告', '', '', 'プレジ', 2, 'ビジター交流会・トレーニング・ビッグビジネスミーティングの日程など'],
  ['バイスプレジデントによる報告', '', '', 'バイス', 2, '読み上げる文（全文）'],
  ['一般規定', '', '', 'バイス', 2, '読み上げる番号（例: 2番）'],
  ['真正度確認', '', '', 'バイス', 0, '前々回の外部リファーラルから選ぶ（〇〇さん⇒〇〇さん　〇〇様）']
];

// --- 開いたとき（onOpen から呼ぶ。単純なトリガーの中でも使えるものだけを使う）---
// 作ったシートの名前の一覧を返す（何も作らなければ空）
function setupOnOpen_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(CHAPTER_SETUP_KEY_)) return [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return [];
  var lock = null;
  try { lock = LockService.getScriptLock(); lock.waitLock(10000); } catch (e) { lock = null; }
  try {
    if (props.getProperty(CHAPTER_SETUP_KEY_)) return [];         // 待っている間に、ほかの方が済ませた
    var blank = setupIsBlank_(ss), made = [];
    if (blank) {
      made = setupBaseSheets_(ss);
      setupRemoveEmptySheets_(ss);
    }
    props.setProperty(CHAPTER_SETUP_KEY_, JSON.stringify({ mode: blank ? 'fresh' : 'existing', at: chapterFmt_(new Date()), made: made }));
    if (made.length) {
      try {
        ss.toast('「' + made.join('」「') + '」のシートを作りました。次に、メニュー「名簿システム」の'
          + ' ⚙️ 設定 > チャプター で、チャプター名・いまの期・定例会の回数を保存してください'
          + '（保存すると、ルーティンチェックシートを作ります）。', '初回の準備', 30);
      } catch (e) {}
    }
    return made;
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e) {} }
  }
}

// シートがどれも空か（新しく作ったスプレッドシート）
function setupIsBlank_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getLastRow() > 0 || sheets[i].getLastColumn() > 0) return false;
  }
  return true;
}

// メンバー名簿・業種区分マスタ・休会日のうち、無いものを作る → 作ったシートの名前
function setupBaseSheets_(ss) {
  var made = [];
  if (!ss.getSheetByName(MEMBER_SHEET_)) { ensureMemberSheet_(); made.push(MEMBER_SHEET_); }
  if (!ss.getSheetByName(CAT_SHEET_)) { ensureCategorySheet_(); made.push(CAT_SHEET_); }
  if (!ss.getSheetByName(SETUP_HOLIDAY_SHEET_)) {
    ss.insertSheet(SETUP_HOLIDAY_SHEET_).hideSheet();
    made.push(SETUP_HOLIDAY_SHEET_);
  }
  return made;
}

// 最初からある空のシート（「シート1」）を消す。システムのシートと、中身のあるシートは残す
function setupRemoveEmptySheets_(ss) {
  var keep = {};
  keep[MEMBER_SHEET_] = keep[CAT_SHEET_] = keep[SETUP_HOLIDAY_SHEET_] = true;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (keep[sh.getName()] || ROUTINE_SHEET_RE_.test(sh.getName())) continue;
    if (sh.getLastRow() > 0 || sh.getLastColumn() > 0) continue;
    try { ss.deleteSheet(sh); } catch (e) {}
  }
}

// --- メニュー「⚙️ 設定 > 足りないシートを作る」---
function menuCreateMissingSheets() {
  var r = createMissingSheets();
  try { SpreadsheetApp.getUi().alert('足りないシートを作る', r.message, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
  return r;
}

// 足りないシートを作る → { ok, made: [シートの名前], message }
function createMissingSheets() {
  try {
    // ルーティンチェックシートの期と回数は、チャプターの設定で決まる。設定を保存する前に作ると、Activeチャプターの期・回数で
    // 作ってしまい、あとで設定を保存しても直らない（空でないスプレッドシートから始めたチャプターも、シートが1枚も無ければ待つ）
    var ss = getSS_(), made = setupBaseSheets_(ss), lines = [];
    var noRoutine = !ss.getSheets().some(function (sh) { return ROUTINE_SHEET_RE_.test(sh.getName()); });
    var waiting = !chapterInfo_().saved && (chapterFresh_() || noRoutine);
    var routine = waiting ? [] : setupEnsureRoutineSheets_();
    for (var i = 0; i < routine.length; i++) {
      if (routine[i].made) {
        made.push(routine[i].name);
        lines.push('「' + routine[i].name + '」… ' + routine[i].meetings + '回ぶんの開催日'
          + (routine[i].from ? '（項目は「' + routine[i].from + '」から写しました）' : ''));
      }
    }
    var msg = made.length ? '次のシートを作りました。\n' + made.map(function (n) { return '・' + n; }).join('\n')
                          : '足りないシートはありません。';
    if (lines.length) msg += '\n\n' + lines.join('\n');
    if (waiting) {
      msg += '\n\nルーティンチェックシートは、チャプターの設定（いまの期・定例会の回数）を保存したときに作ります。'
        + '先に ⚙️ 設定 > チャプター を保存してください。';
    }
    // 期のシートはあるのに、次の定例会の列が無い（チャプターの設定で開催日・曜日を直した、など）。黙って「足りない」と言わない
    if (!waiting) {
      var gaps = [];
      getMeetingCandidates().forEach(function (c) {
        var d = parseDate_(c.dateValue);
        if (d && !findRoutineColumn_(d) && ss.getSheetByName(setupRoutineName_(roleTermOf_(d)))) gaps.push(c.dateValue);
      });
      if (gaps.length) {
        msg += '\n\nルーティンチェックシートはありますが、' + gaps.join('・') + ' の列が見つかりません'
          + '（シートの1行目の日付が、チャプターの設定の定例会の曜日・日付と合っていません）。'
          + 'チャプターの設定で開催日・曜日・回数を直したときは、シートの1行目（日付）と2行目（回数）も直してください。';
      }
    }
    return { ok: true, made: made, message: msg };
  } catch (e) {
    console.error('[SETUP] ' + (e && e.stack ? e.stack : e));
    return { ok: false, made: [], message: 'シートを作れませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// --- ルーティンチェックシート ---
function setupRoutineName_(term) { return '【' + term + '期】ルーティンチェックシート'; }

// 次の4回の定例会のうち、どのシートにも列が無い回があれば、その期のシートを作る
// → [{ name, term, made, meetings（開催日の数）, from（写した元のシート） }]
function setupEnsureRoutineSheets_() {
  var terms = [], out = [];
  var cands = getMeetingCandidates();
  for (var i = 0; i < cands.length; i++) {
    var d = parseDate_(cands[i].dateValue);
    if (!d || findRoutineColumn_(d)) continue;                  // どこかのシートに列がある
    var t = roleTermOf_(d);
    if (terms.indexOf(t) < 0) terms.push(t);
  }
  for (var j = 0; j < terms.length; j++) out.push(setupEnsureRoutineSheet_(terms[j]));
  return out;
}

// その期のシートが無ければ作る
function setupEnsureRoutineSheet_(term) {
  var ss = getSS_(), name = setupRoutineName_(term);
  if (ss.getSheetByName(name)) return { name: name, term: term, made: false, meetings: 0, from: '' };
  var src = setupRoutineSource_(ss, term), meetings = setupTermMeetings_(term);
  if (src) setupCopyRoutine_(ss, src, name, meetings);
  else setupNewRoutine_(ss, name, meetings);
  routineResetCache_();
  return { name: name, term: term, made: true, meetings: meetings.length, from: src ? src.getName() : '' };
}

// 写す元：その期より前でいちばん新しい期のシート（無ければ、あとの期でいちばん近いもの）
function setupRoutineSource_(ss, term) {
  var sheets = ss.getSheets(), before = null, after = null;
  for (var i = 0; i < sheets.length; i++) {
    var m = sheets[i].getName().match(ROUTINE_SHEET_RE_);
    if (!m) continue;
    var t = parseInt(m[1], 10);
    if (t < term && (!before || t > before.t)) before = { t: t, sheet: sheets[i] };
    if (t > term && (!after || t < after.t)) after = { t: t, sheet: sheets[i] };
  }
  return before ? before.sheet : (after ? after.sheet : null);
}

// 期（数字）→ その半期の初日と最終日（期は 4月〜9月・10月〜3月）
function setupTermRange_(term) {
  var half = term - chapterTermBase_() + CHAPTER_TERM_YEAR_ * 2, y = Math.floor(half / 2);
  return (half % 2 === 0) ? { from: new Date(y, 3, 1), to: new Date(y, 8, 30) }
                          : { from: new Date(y, 9, 1), to: new Date(y + 1, 2, 31) };
}

// その期の開催日（定例会の曜日の毎週）→ [{ date: 'yyyy/MM/dd', count: 回数（休会日は ''） }]
function setupTermMeetings_(term) {
  var r = setupTermRange_(term), wd = chapterWeekday_(), holidays = getHolidays(), out = [];
  var d = new Date(r.from.getTime());
  while (d.getDay() !== wd) d.setDate(d.getDate() + 1);
  for (; d.getTime() <= r.to.getTime(); d.setDate(d.getDate() + 7)) {
    out.push({ date: chapterFmt_(d), count: setupCountOf_(d, holidays) });
  }
  return out;
}

// 開催日の回数。チャプターの設定の基準の日から、休会日の週を除いて数える（基準より前の日も数える）。
// 休会日・曜日の違う日・1回より前は ''
function setupCountOf_(d, holidays) {
  var base = chapterMeetingBase_(), cur = new Date(base.date.getTime()), count = base.count, t = d.getTime();
  if (holidays.indexOf(chapterFmt_(d)) >= 0) return '';
  for (var i = 0; i < 3000 && cur.getTime() !== t; i++) {
    if (t > cur.getTime()) {
      if (holidays.indexOf(chapterFmt_(cur)) < 0) count++;
      cur.setDate(cur.getDate() + 7);
      if (cur.getTime() > t) return '';
    } else {
      cur.setDate(cur.getDate() - 7);
      if (holidays.indexOf(chapterFmt_(cur)) < 0) count--;
      if (cur.getTime() < t) return '';
    }
  }
  return (cur.getTime() === t && count > 0) ? count : '';
}

// 「何日前」→ 期日の欄（「5日前」「当日」）と、曜日目安の欄（水曜開催なら「金曜まで」「火曜」「水曜」）
function setupDueText_(days) {
  if (days === '' || days == null) return '';
  return days === 0 ? '当日' : days + '日前';
}
function setupWeekdayText_(days, weekday) {
  if (days === '' || days == null) return '';
  var w = CHAPTER_WEEK_[((weekday - days) % 7 + 7) % 7] + '曜';
  return days >= 2 ? w + 'まで' : w;
}

// 項目の一覧（SETUP_ROUTINE_ITEMS_）から新しく作る
function setupNewRoutine_(ss, name, meetings) {
  var sh = ss.insertSheet(name, 0), wd = chapterWeekday_(), W = SETUP_ROUTINE_HEAD_.length;
  var width = SETUP_ROUTINE_DATE_COL_ - 1 + Math.max(meetings.length, 1);
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  var pad = function (a) { a = a.slice(); while (a.length < W) a.push(''); return a; };
  var rows = [pad(['', '定例会開催日']), pad(['', '定例会回数']), pad(['', '担当割り振り']), SETUP_ROUTINE_HEAD_.slice()];
  for (var i = 0; i < SETUP_ROUTINE_ITEMS_.length; i++) {
    var it = SETUP_ROUTINE_ITEMS_[i];
    rows.push(['', i + 1, it[0], it[1], it[2], it[3], setupDueText_(it[4]), setupWeekdayText_(it[4], wd), it[5]]);
  }
  sh.getRange(1, 1, rows.length, W).setValues(rows);
  if (meetings.length) {
    sh.getRange(1, SETUP_ROUTINE_DATE_COL_, 2, meetings.length).setValues([
      meetings.map(function (m) { return m.date; }), meetings.map(function (m) { return m.count; })]);
  }
  // 書式（うまくいかなくても、シートは使える）
  try {
    var widths = [16, 40, 170, 160, 90, 90, 60, 80, 220];
    for (var c = 0; c < widths.length; c++) sh.setColumnWidth(c + 1, widths[c]);
    for (var k = 0; k < meetings.length; k++) sh.setColumnWidth(SETUP_ROUTINE_DATE_COL_ + k, 150);
    sh.getRange(1, 2, 3, 1).setFontWeight('bold');
    sh.getRange(4, 1, 1, W).setFontWeight('bold').setBackground('#eeeeee');
    if (meetings.length) {
      sh.getRange(1, SETUP_ROUTINE_DATE_COL_, 2, meetings.length).setFontWeight('bold').setBackground('#dbe7ff').setHorizontalAlignment('center');
      sh.getRange(1, SETUP_ROUTINE_DATE_COL_, 1, meetings.length).setNumberFormat('yyyy/mm/dd(ddd)');
      for (var h = 0; h < meetings.length; h++) {
        if (meetings[h].count === '') sh.getRange(1, SETUP_ROUTINE_DATE_COL_ + h, 2, 1).setBackground('#e0e0e0');   // 休会日
      }
    }
    sh.getRange(5, 1, rows.length - 4, width).setWrap(true).setVerticalAlignment('top');
    sh.setFrozenRows(4);
    sh.setFrozenColumns(5);
  } catch (e) {
    console.error('[SETUP] 書式: ' + (e && e.message ? e.message : e));
  }
  return sh;
}

// 前の期のシートを写して、開催日の列だけ新しくする（項目・担当・期日・備考と書式はそのまま）
function setupCopyRoutine_(ss, src, name, meetings) {
  var sh = src.copyTo(ss);
  sh.setName(name);
  try { ss.setActiveSheet(sh); ss.moveActiveSheet(src.getIndex()); } catch (e) {}   // 写した元の前（新しい期が前）に置く
  var last = sh.getLastColumn(), row1 = last ? sh.getRange(1, 1, 1, last).getValues()[0] : [], cols = [];
  for (var c = 0; c < row1.length; c++) if (routineDate_(row1[c])) cols.push(c + 1);
  var first = cols.length ? cols[0] : Math.max(SETUP_ROUTINE_DATE_COL_, last + 1);
  var stride = cols.length > 1 ? Math.max(1, cols[1] - cols[0]) : 1;     // 1列おき（開催日と、その横の欄）の表もある
  if (last >= first) sh.getRange(1, first, sh.getMaxRows(), last - first + 1).clearContent();
  if (!meetings.length) return sh;
  var span = stride * meetings.length, need = first + span - 1;
  if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
  var r1 = [], r2 = [];
  for (var i = 0; i < meetings.length; i++) {
    for (var s = 0; s < stride; s++) {
      r1.push(s === 0 ? meetings[i].date : '');
      r2.push(s === 0 ? meetings[i].count : '');
    }
  }
  sh.getRange(1, first, 2, span).setValues([r1, r2]);
  return sh;
}

// 期の番号を付け直したとき（チャプターの設定）、初回の準備で作ったスプレッドシートなら、
// ルーティンチェックシートの名前の期も同じだけずらす（担当者・チームの期と合わせる）。
// 同じ名前のシートがあれば、そのシートはそのままにする
function setupShiftRoutineNames_(delta) {
  if (!delta || !chapterFresh_()) return [];
  var ss = getSS_(), sheets = ss.getSheets(), list = [], done = [];
  for (var i = 0; i < sheets.length; i++) {
    var m = sheets[i].getName().match(ROUTINE_SHEET_RE_);
    if (m) list.push({ sheet: sheets[i], term: parseInt(m[1], 10) });
  }
  list.sort(function (a, b) { return delta > 0 ? b.term - a.term : a.term - b.term; });   // ぶつからない順に
  for (var j = 0; j < list.length; j++) {
    var to = setupRoutineName_(list[j].term + delta);
    if (list[j].term + delta < 1 || ss.getSheetByName(to)) continue;
    list[j].sheet.setName(to);
    done.push(to);
  }
  if (done.length) routineResetCache_();
  return done;
}
