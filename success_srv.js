// === チャプターサクセスミーティング（月に1回の、役職ごとの振り返り）===
//
// これまでは Excel（月ごとのシート）を画面共有して進めていた会議。この画面で月ごとに入力・記録する。
//   ・議題（役職ごとのまとまり）ごとに「前月のアクションアイテム」「前月の結果」「今月のアクションアイテム」。
//     前月のアクションアイテムは、前の月に入れた「今月のアクションアイテム」がそのまま出る（入れ直さない）
//   ・記入の状況（入力済みか・だれがいつ入れたか）が分かる
//   ・システムで分かることは「推定」として出し、ボタンで入れる（ビジター数・新メンバー／退会者・
//     パスポートプログラムの方・7ヶ月レビュー・更新の承認）
//   ・これまでの Excel を取り込める（月ごとのシートを、議題の名前で合わせて読む）
//   ・画面共有用に、Excel と同じ並びのシート「サクセスミーティング」に書き出せる
// 記録はシート「サクセスミーティング記録」（隠す）に、1か月・1議題につき1行で持つ。

var SM_SHEET_ = 'サクセスミーティング記録';
var SM_VIEW_SHEET_ = 'サクセスミーティング';
var SM_HEAD_ = ['月', '議題', '前月の結果', '今月のアクションアイテム', '更新者', '更新日時'];
var SM_META_ = '_meta';                    // その月の書記・出席者（JSON を「前月の結果」の欄に入れる）

// 議題（Excel のフォームの並び）。checks … 議題の下に並ぶ確かめる点（入力欄ではない）
var SM_BLOCKS_ = [
  { key: 'vhc', title: '【ビジターホストコーディネーター】',
    checks: ['□ビジター出席レビュー', '先月のビジター数は？', 'メンバーシップ申し込みを行ったビジターは誰ですか？', 'フォローアップが必要なビジターは誰ですか？'] },
  { key: 'membership', title: '【メンバーシップ】',
    checks: ['□メンバーシップレビュー', '先月の退会者数は？', '先月の新メンバー数は？', 'メンバーシップ数の純増は？（ゴール：毎月純増1）'] },
  { key: 'web', title: '【WEBマスター】', checks: [] },
  { key: 'ec', title: '【エデュケーションコーディネーター】', checks: [] },
  { key: 'one2one', title: '【1 to 1促進委員】', checks: [] },
  { key: 'training', title: '【トレーニング促進委員】', checks: [] },
  { key: 'event', title: '【イベント委員】', checks: [] },
  { key: 'bcp', title: '【BCP委員】', checks: [] },
  { key: 'support', title: '【メンバーサポート委員】', checks: [] },
  { key: 'spreading', title: '【Spreading委員】', checks: [] },
  { key: 'mentor', title: '【メンターコーディネーター】', checks: ['□BNIパスポートプログラム'] },
  { key: 'gbc', title: '【グローバルビジネスコーディネーター】', checks: [] },
  { key: 'admin', title: 'メンターコーディネーターとの連携（運営マニュアルチェック）',
    checks: ['□7ヶ月レビュー', '□更新の承認', '□書面で寄せられた苦情への対応進捗', '□チャプタープレジデントへの要望', '□ディレクターコンサルタントへの要望'] }
];

// 議題の見出しの書き方の違いを吸収する（「【1 to 1促進委員】」「1to1促進委員」など）
function smNorm_(s) {
  return String(s == null ? '' : s).normalize('NFKC').replace(/[\s【】\[\]□■◆・]/g, '').toLowerCase();
}
// 見出しの文字 → 議題のキー（知らない見出しは '' ）
function smBlockKeyOf_(title) {
  var t = smNorm_(String(title || '').split('\n')[0]);
  if (!t) return '';
  for (var i = 0; i < SM_BLOCKS_.length; i++) if (smNorm_(SM_BLOCKS_[i].title) === t) return SM_BLOCKS_[i].key;
  if (/^メンターコーディネーターとの連携/.test(t) || /^(7ヶ月|7か月|7カ月)レビュー/.test(t)) return 'admin';
  if (/^ビジターホスト/.test(t)) return 'vhc';
  if (/^メンバーシップ$/.test(t)) return 'membership';
  if (/^(web|ウェブ)マスター/.test(t)) return 'web';
  if (/^エデュケーション/.test(t)) return 'ec';
  if (/^1to1/.test(t)) return 'one2one';
  if (/^トレーニング/.test(t)) return 'training';
  if (/^イベント/.test(t)) return 'event';
  if (/^bcp/.test(t)) return 'bcp';
  if (/^メンバーサポート/.test(t)) return 'support';
  if (/^(spreading|スプレディング)/.test(t)) return 'spreading';
  if (/^メンターコーディネーター$/.test(t)) return 'mentor';
  if (/^グローバルビジネス/.test(t)) return 'gbc';
  return '';
}
function smBlockDef_(key) {
  for (var i = 0; i < SM_BLOCKS_.length; i++) if (SM_BLOCKS_[i].key === key) return SM_BLOCKS_[i];
  return null;
}

// 'YYYY-MM' ⇔ 年・月
function smMonthKey_(y, m) { return y + '-' + (m < 10 ? '0' : '') + m; }
function smParseMonth_(s) {
  var m = String(s || '').match(/^(\d{4})-(\d{1,2})$/);
  return m ? { y: parseInt(m[1], 10), m: parseInt(m[2], 10) } : null;
}
function smShiftMonth_(key, n) {
  var p = smParseMonth_(key), d = new Date(p.y, p.m - 1 + n, 1);
  return smMonthKey_(d.getFullYear(), d.getMonth() + 1);
}
function smMonthLabel_(key) { var p = smParseMonth_(key); return p ? p.y + '年' + p.m + '月' : String(key || ''); }

// 記録のシートの「月」の欄 → 'YYYY-MM'。スプレッドシートが「2026-10」を日付（10月1日）に変えてしまうことがあるので、
// 日付になっていても読む（前は読めず、取り込んだ月・保存した月が出なかった）
function smCellMonth_(x) {
  if (x && typeof x.getFullYear === 'function' && !isNaN(x.getTime())) return smMonthKey_(x.getFullYear(), x.getMonth() + 1);
  var p = smParseMonth_(String(x == null ? '' : x).trim().replace(/^'/, ''));
  return p ? smMonthKey_(p.y, p.m) : '';
}
// 書き込む「月」：文字のまま入るように、頭に ' を付ける
function smMonthCell_(mk) { return "'" + mk; }

// --- 記録のシート ---
function smSheet_(create) {
  var ss = getSS_(), sh = ss.getSheetByName(SM_SHEET_);
  if (!sh && create) {
    sh = ss.insertSheet(SM_SHEET_);
    sh.getRange(1, 1, 1, SM_HEAD_.length).setValues([SM_HEAD_]);
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}
// → { 'YYYY-MM': { blocks: { key: { result, action, by, at, title? } }, meta: {}, order: [キー] } }
function smLoad_() {
  var sh = smSheet_(false), out = {};
  if (!sh) return out;
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    var mk = smCellMonth_(v[i][0]), key = String(v[i][1] || '').trim();
    if (!mk || !key) continue;
    var mo = out[mk] || (out[mk] = { blocks: {}, meta: {}, order: [] });
    if (key === SM_META_) { try { mo.meta = JSON.parse(String(v[i][2] || '{}')) || {}; } catch (e) {} continue; }
    var at = v[i][5] && typeof v[i][5].getFullYear === 'function' ? Utilities.formatDate(v[i][5], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : String(v[i][5] || '');
    var title = '';
    if (key.indexOf('x:') === 0) title = key.substring(2);           // Excel にあった、ほかの議題
    mo.blocks[key] = { result: String(v[i][2] || ''), action: String(v[i][3] || ''), by: String(v[i][4] || ''), at: at, row: i + 1, title: title };
    mo.order.push(key);
  }
  return out;
}

// その月の議題の並び：決まった議題＋その月（無ければ前の月）にある、ほかの議題
function smBlockList_(all, mk) {
  var list = SM_BLOCKS_.map(function (b) { return { key: b.key, title: b.title, checks: b.checks.slice() }; });
  var have = {};
  list.forEach(function (b) { have[b.key] = true; });
  [mk, smShiftMonth_(mk, -1)].forEach(function (k) {
    var mo = all[k];
    if (!mo) return;
    mo.order.forEach(function (key) {
      if (have[key] || key.indexOf('x:') !== 0) return;
      have[key] = true;
      list.push({ key: key, title: key.substring(2), checks: [] });
    });
  });
  return list;
}

// 書いた方（ウェブアプリはその方のアカウント）
function smUser_() {
  try { return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || ''; } catch (e) { return ''; }
}

// --- 画面 ---
// month … 'YYYY-MM'（空なら今月）
function getSuccessMeeting(month) {
  try {
    requireSheetAccess_();
    var now = new Date(), cur = smMonthKey_(now.getFullYear(), now.getMonth() + 1);
    var mk = smParseMonth_(month) ? smMonthKey_(smParseMonth_(month).y, smParseMonth_(month).m) : cur;
    var all = smLoad_(), mo = all[mk] || { blocks: {}, meta: {} }, prev = all[smShiftMonth_(mk, -1)] || { blocks: {} };
    // 選べる月：記録のあるいちばん前の月（無ければ1年前）から来月まで
    var keys = Object.keys(all).sort(), first = keys.length && keys[0] < smShiftMonth_(cur, -12) ? keys[0] : smShiftMonth_(cur, -12);
    var months = [];
    for (var k = smShiftMonth_(cur, 1); k >= first; k = smShiftMonth_(k, -1)) months.push(k);
    if (months.indexOf(mk) < 0) { months.push(mk); months.sort().reverse(); }
    // 前の月までの記録（新しい順。その議題に何か書いてある月だけ、12か月ぶんまで）。参考に見る・コピーする
    var past = keys.filter(function (k) { return k < mk; }).reverse();
    var auto = {};
    try { auto = smAuto_(mk); } catch (e) { console.warn('[SM] 推定を作れませんでした: ' + (e && e.message ? e.message : e)); }
    var blocks = smBlockList_(all, mk).map(function (b) {
      var x = mo.blocks[b.key] || {}, p = prev.blocks[b.key] || {};
      var history = [];
      past.forEach(function (k) {
        var h = all[k].blocks[b.key];
        if (history.length < 12 && h && (h.result || h.action)) history.push({ month: k, label: smMonthLabel_(k), result: h.result, action: h.action });
      });
      return { key: b.key, title: b.title, checks: b.checks, prevAction: p.action || '', prevResult: p.result || '', history: history,
               result: x.result || '', action: x.action || '', by: x.by || '', at: x.at || '', auto: auto[b.key] || '' };
    });
    var roster = [];
    try { roster = (getMemberMaster({ membersOnly: true }).members || []).map(function (m) { return { no: m.no || '', name: m.name }; }); } catch (e) {}
    var secretary = mo.meta.secretary || '';
    if (!secretary) {
      try {
        var p = smParseMonth_(mk), d = new Date(p.y, p.m - 1, 15);
        secretary = roleHoldersOfTerm_(roleHolderTerms_(), roleTermOf_(d)).holders.secretary || '';
      } catch (e) {}
    }
    return { ok: true, month: mk, label: smMonthLabel_(mk), prevLabel: smMonthLabel_(smShiftMonth_(mk, -1)),
             months: months.map(function (k) { return { value: k, label: smMonthLabel_(k), saved: !!all[k] }; }),
             blocks: blocks, secretary: secretary, attendees: mo.meta.attendees || [], roster: roster };
  } catch (e) {
    console.error('[SM] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'サクセスミーティングを読み込めませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// items … [{ key, result, action, origResult, origAction }]（変えた議題だけ）。meta … { secretary, attendees }
// ほかの方が先に同じ欄を書き換えていたら（origResult・origAction が今の記録と違う）、その欄は上書きせず知らせる
function saveSuccessMeeting(month, items, meta) {
  var lock = LockService.getScriptLock();
  try {
    requireSheetAccess_();
    var p = smParseMonth_(month);
    if (!p) return { ok: false, message: '月が分かりません。' };
    var mk = smMonthKey_(p.y, p.m);
    lock.waitLock(20000);
    var sh = smSheet_(true), all = smLoad_(), mo = all[mk] || { blocks: {}, meta: {} };
    var user = smUser_(), at = new Date(), conflicts = [], saved = 0, clean = function (s) { return String(s == null ? '' : s).replace(/\r\n?/g, '\n').slice(0, 5000); };
    (items || []).forEach(function (it) {
      var key = String(it && it.key || '');
      if (!key || key === SM_META_ || !(smBlockDef_(key) || key.indexOf('x:') === 0)) return;
      var cur = mo.blocks[key] || { result: '', action: '' };
      var r = clean(it.result), a = clean(it.action);
      var okR = it.origResult === undefined || clean(it.origResult) === cur.result || r === cur.result;
      var okA = it.origAction === undefined || clean(it.origAction) === cur.action || a === cur.action;
      var title = (smBlockDef_(key) || { title: key.substring(2) }).title;
      if (!okR) { conflicts.push(title + '（結果）'); r = cur.result; }
      if (!okA) { conflicts.push(title + '（アクションアイテム）'); a = cur.action; }
      if (r === cur.result && a === cur.action) return;
      var row = [smMonthCell_(mk), key, r, a, user, at];
      if (cur.row) sh.getRange(cur.row, 1, 1, row.length).setValues([row]);
      else sh.appendRow(row);
      saved++;
    });
    if (meta) {
      var m = { secretary: clean(meta.secretary).slice(0, 60),
                attendees: (meta.attendees || []).map(function (n) { return clean(n).slice(0, 60); }).filter(Boolean) };
      var json = JSON.stringify(m), metaRow = 0;
      var v = sh.getDataRange().getValues();
      for (var i = 1; i < v.length; i++) if (smCellMonth_(v[i][0]) === mk && String(v[i][1]) === SM_META_) { metaRow = i + 1; break; }
      if (JSON.stringify(mo.meta || {}) !== json) {
        if (metaRow) sh.getRange(metaRow, 1, 1, 6).setValues([[smMonthCell_(mk), SM_META_, json, '', user, at]]);
        else sh.appendRow([smMonthCell_(mk), SM_META_, json, '', user, at]);
      }
    }
    SpreadsheetApp.flush();
    var res = getSuccessMeeting(mk);
    res.message = smMonthLabel_(mk) + 'のサクセスミーティングを保存しました（' + saved + '件）。'
      + (conflicts.length ? '\n⚠ ほかの方が先に書き換えていたので、保存しなかった欄があります（いまの内容を出しました。もう一度入れてください）: ' + conflicts.join('、') : '');
    res.conflicts = conflicts;
    return res;
  } catch (e) {
    console.error('[SM] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '保存できませんでした: ' + (e && e.message ? e.message : e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// --- システムで分かること（「前月の結果」の推定）---
// 前月の開催日（チャプターの開催曜日で、休会日を除く）
function smMeetingDates_(mk) {
  var p = smParseMonth_(mk), out = [], holidays = [];
  try { holidays = getHolidays(); } catch (e) {}
  var wd = chapterMeetingBase_().date.getDay();
  for (var d = new Date(p.y, p.m - 1, 1); d.getMonth() === p.m - 1; d.setDate(d.getDate() + 1)) {
    if (d.getDay() === wd && holidays.indexOf(fmtDate_(d)) < 0) out.push(new Date(d.getTime()));
  }
  return out;
}
// その月の参加者シートのビジター（No. が V）・ゲスト（G）の人数。キャンセルは数えない
function smVisitorCounts_(mk) {
  var p = smParseMonth_(mk), v = 0, g = 0, n = 0;
  getExistingVisitorSheets().forEach(function (name) {
    var key = (name.match(/^(\d{8}|\d{4})参加者$/) || [])[1], d = key ? meetingDateFromKey_(key) : null;
    if (!d || d.getFullYear() !== p.y || d.getMonth() !== p.m - 1) return;
    var data = getSS_().getSheetByName(name).getDataRange().getValues(), h = -1;
    for (var i = 0; i < Math.min(10, data.length); i++) if (data[i].indexOf('No.') !== -1) { h = i; break; }
    if (h < 0) return;
    var hd = data[h].map(function (x) { return String(x).trim(); }), noIdx = hd.indexOf('No.');
    var stIdx = hd.indexOf('ステータス') >= 0 ? hd.indexOf('ステータス') : hd.indexOf('出席ステータス');
    n++;
    for (var r = h + 1; r < data.length; r++) {
      var no = String(data[r][noIdx] == null ? '' : data[r][noIdx]).trim();
      if (stIdx >= 0 && /キャンセル|cancel/i.test(String(data[r][stIdx] || ''))) continue;
      if (/^V/i.test(no)) v++;
      else if (/^G/i.test(no)) g++;
    }
  });
  return { visitors: v, guests: g, meetings: n };
}
// その月のルーティンチェックシートの「退会者」に書いてあるお名前
function smLeavers_(mk) {
  var out = [];
  smMeetingDates_(mk).forEach(function (d) {
    var ctx = null;
    try { ctx = roleBuildContext_(d, ''); } catch (e) { ctx = null; }
    if (!ctx || !ctx.found) return;
    for (var i = 0; i < ctx.order.length; i++) {
      var it = ctx.items[ctx.order[i]];
      if (!/^退会者?$/.test(roleMatchText_(it.title))) continue;
      String(premtgText_(it.value) || '').split(/[、,，\n]+/).forEach(function (s) {
        var n = s.replace(/[（(].*$/, '').replace(/(さん|様)$/, '').trim();
        if (n && !premtgNone_(n) && out.indexOf(n) < 0) out.push(n);
      });
    }
  });
  return out;
}
// 議題のキー → 推定の文（前月の結果に入れる候補）
function smAuto_(mk) {
  var prevMk = smShiftMonth_(mk, -1), pp = smParseMonth_(prevMk), p = smParseMonth_(mk);
  var roster = getMemberMaster({ membersOnly: true }).members || [], out = {};
  var inMonth = function (d, y, m) { return d && d.getFullYear() === y && d.getMonth() === m - 1; };
  // ビジター
  try {
    var vc = smVisitorCounts_(prevMk);
    if (vc.meetings) out.vhc = '◆' + pp.m + '月ビジター数　ビジター' + vc.visitors + '名　ゲスト' + vc.guests + '名（定例会 ' + vc.meetings + '回）';
  } catch (e) {}
  // メンバーシップ：入会日が前月の方・前月のルーティンチェックシートの退会者
  var joined = roster.filter(function (m) { return inMonth(parseDate_(m.joinDate), pp.y, pp.m); }).map(function (m) { return m.name; });
  var left = [];
  try { left = smLeavers_(prevMk); } catch (e) {}
  out.membership = '◆新メンバー　' + joined.length + '名' + (joined.length ? '（' + joined.map(rotHonor_).join('、') + '）' : '')
    + '\n◆退会者　' + left.length + '名' + (left.length ? '（' + left.map(rotHonor_).join('、') + '）' : '')
    + '\n◆純増　' + (joined.length - left.length > 0 ? '+' : '') + (joined.length - left.length) + '名';
  // メンターコーディネーター：パスポートプログラムの方（その月の15日時点で、入会日から4か月以内）
  var mid = new Date(p.y, p.m - 1, 15);
  var recent = roleRecentMembers_(roster, mid);
  if (recent.length) out.mentor = '□BNIパスポートプログラム（入会から4か月以内）：' + recent.map(rotHonor_).join('、');
  // 運営：7ヶ月レビュー（入会からその月で7か月目）・更新の承認（更新期限がその月か翌月）
  var seven = roster.filter(function (m) {
    var j = parseDate_(m.joinDate);
    if (!j) return false;
    var r = new Date(j.getFullYear(), j.getMonth() + 7, 1);
    return r.getFullYear() === p.y && r.getMonth() === p.m - 1;
  }).map(function (m) { return m.name; });
  var renew = roster.filter(function (m) {
    var x = parseDate_(m.expireDate), nx = new Date(p.y, p.m, 1);
    return inMonth(x, p.y, p.m) || inMonth(x, nx.getFullYear(), nx.getMonth() + 1);
  }).map(function (m) { var x = parseDate_(m.expireDate); return rotHonor_(m.name) + '（' + (x.getMonth() + 1) + '/' + x.getDate() + '）'; });
  var adm = [];
  if (seven.length) adm.push('□7ヶ月レビュー：' + seven.map(rotHonor_).join('、'));
  if (renew.length) adm.push('□更新の承認（更新期限が今月・来月）：' + renew.join('、'));
  if (adm.length) out.admin = adm.join('\n');
  return out;
}

// --- Excel の取り込み ---
// 1か月ぶんのシートの値（2次元。結合したマスは左上だけに値）→ { blocks: { key: { result, action, title } }, secretary }
// 見出しの行（「〜の結果」「〜のアクションアイテム」）から、結果・今月のアクションアイテムの列を決める
// （右にある方のアクションアイテムが今月のもの。2025年の書き方は A・C・E、2026年4月からの書き方は A・B・C・D）
function smParseGrid_(grid) {
  var res = { blocks: {}, order: [], secretary: '' }, hr = -1, rc = -1, ac = -1;
  for (var r = 0; r < Math.min(12, grid.length) && hr < 0; r++) {
    var row = grid[r] || [], rr = -1, aa = [];
    for (var c = 0; c < row.length; c++) {
      var t = String(row[c] == null ? '' : row[c]);
      if (/の結果/.test(t)) rr = c;
      if (/アクションアイテム/.test(t) && !/前回のミーティング/.test(t)) aa.push(c);
    }
    if (rr >= 0 && aa.length) { hr = r; rc = rr; ac = aa[aa.length - 1]; }
  }
  for (var s = 0; s < Math.min(4, grid.length); s++) {
    (grid[s] || []).forEach(function (x) {
      var m = String(x == null ? '' : x).match(/書記[:：][\s　]*([^\s　]+)/);
      if (m && !res.secretary) res.secretary = m[1].trim();
    });
  }
  if (hr < 0) return res;
  var cur = '';
  for (var i = hr + 1; i < grid.length; i++) {
    var line = grid[i] || [], head = String(line[0] == null ? '' : line[0]).trim();
    var key = head ? smBlockKeyOf_(head) : '';
    if (!key && /^【[^】]+】/.test(head)) key = 'x:' + head.match(/^【[^】]+】/)[0];
    if (key) cur = key;
    if (!cur) continue;
    var b = res.blocks[cur] || (res.blocks[cur] = { result: [], action: [] });
    if (res.order.indexOf(cur) < 0) res.order.push(cur);
    var rv = String(line[rc] == null ? '' : line[rc]).trim(), av = String(line[ac] == null ? '' : line[ac]).trim();
    if (rv && b.result.indexOf(rv) < 0) b.result.push(rv);
    if (av && b.action.indexOf(av) < 0) b.action.push(av);
  }
  res.order.forEach(function (k) { var b = res.blocks[k]; b.result = b.result.join('\n'); b.action = b.action.join('\n'); });
  return res;
}
// シートの名前 → 'YYYY-MM'（「2026年11月 」など）
function smMonthOfSheetName_(name) {
  var m = String(name || '').normalize('NFKC').match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
  return m ? smMonthKey_(parseInt(m[1], 10), parseInt(m[2], 10)) : '';
}
// 画面から：Excel（base64）を取り込む。overwrite … もう記録のある月も、Excel の内容で書き直す
function importSuccessMeetingXlsx(base64, fileName, overwrite) {
  var tmpId = '';
  try {
    requireSheetAccess_();
    var blob = Utilities.newBlob(Utilities.base64Decode(String(base64 || '')),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName || 'success.xlsx');
    // Googleスプレッドシートにして読む（読み終えたら消す）
    var made = Drive.Files.create({ name: '【取り込み用・自動で消えます】' + (fileName || 'サクセスミーティング'),
                                    mimeType: 'application/vnd.google-apps.spreadsheet' }, blob);
    tmpId = made.id;
    var book = SpreadsheetApp.openById(tmpId), months = [];
    book.getSheets().forEach(function (sh) {
      var mk = smMonthOfSheetName_(sh.getName());
      if (!mk) return;
      var g = smParseGrid_(sh.getDataRange().getValues());
      if (g.order.length) months.push({ month: mk, grid: g });
    });
    if (!months.length) return { ok: false, message: '「2026年11月」のような名前の、サクセスミーティングのシートが見つかりませんでした。' };
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var all = smLoad_(), sh = smSheet_(true), user = smUser_() || 'Excelから取り込み', at = new Date(), done = [], skipped = [], rows = [];
      months.forEach(function (x) {
        if (all[x.month] && !overwrite) { skipped.push(smMonthLabel_(x.month)); return; }
        x.grid.order.forEach(function (k) { rows.push([smMonthCell_(x.month), k, x.grid.blocks[k].result, x.grid.blocks[k].action, user, at]); });
        if (x.grid.secretary) rows.push([smMonthCell_(x.month), SM_META_, JSON.stringify({ secretary: x.grid.secretary, attendees: [] }), '', user, at]);
        done.push(x.month);
      });
      if (done.length) {
        // 書き直す月の、前の記録を消してから入れる（下の行から消す）
        var v = sh.getDataRange().getValues();
        for (var i = v.length - 1; i >= 1; i--) if (done.indexOf(smCellMonth_(v[i][0])) >= 0) sh.deleteRow(i + 1);
        if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
      }
      return { ok: true, imported: done, skipped: skipped,
               message: done.length + 'か月ぶんを取り込みました（' + done.map(smMonthLabel_).join('、') + '）。'
                 + (skipped.length ? '\nもう記録のある月は取り込みませんでした（' + skipped.join('、') + '）。書き直すときは「記録のある月も書き直す」にチェックを入れてください。' : '') };
    } finally { lock.releaseLock(); }
  } catch (e) {
    console.error('[SM] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '取り込めませんでした: ' + (e && e.message ? e.message : e) };
  } finally {
    if (tmpId) { try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) {} }
  }
}

// --- 画面共有用のシート（Excel のフォームと同じ並び）---
function exportSuccessMeetingSheet(month) {
  try {
    requireSheetAccess_();
    var r = getSuccessMeeting(month);
    if (!r.ok) return r;
    var ss = getSS_(), sh = ss.getSheetByName(SM_VIEW_SHEET_) || ss.insertSheet(SM_VIEW_SHEET_);
    sh.clear();
    try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (e) {}
    var rows = [['チャプターサクセスミーティングフォーム（' + r.label + '）', '書記：' + (r.secretary || ''), '', '', '', ''],
                ['', '', '', '', '', ''],
                ['議題', r.prevLabel + 'のアクションアイテム', r.prevLabel + 'の結果', r.label + 'のアクションアイテム', '', '本日の出席者']];
    var merges = [], firstRow = rows.length + 1;
    r.blocks.forEach(function (b) {
      var lines = [b.title].concat(b.checks), top = rows.length + 1;
      lines.forEach(function (t, i) { rows.push([i ? '　' + t : t, i ? '' : b.prevAction, i ? '' : b.result, i ? '' : b.action, '', '']); });
      if (lines.length > 1) merges.push([top, lines.length]);
    });
    r.attendees.forEach(function (n, i) {
      var at = firstRow + i;
      while (rows.length < at) rows.push(['', '', '', '', '', '']);
      rows[at - 1][5] = (i + 1) + '　' + n;
    });
    sh.getRange(1, 1, rows.length, 6).setValues(rows);
    sh.getRange(1, 1, rows.length, 6).setWrap(true).setVerticalAlignment('top').setFontSize(11);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold').setFontSize(14);
    sh.getRange(3, 1, 1, 6).setFontWeight('bold').setBackground('#FCE4D6');
    merges.forEach(function (m) { [2, 3, 4].forEach(function (c) { sh.getRange(m[0], c, m[1], 1).merge(); }); });
    sh.getRange(3, 1, rows.length - 2, 4).setBorder(true, true, true, true, true, true);
    sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 300); sh.setColumnWidth(3, 300); sh.setColumnWidth(4, 300);
    sh.setColumnWidth(5, 20); sh.setColumnWidth(6, 160);
    sh.setFrozenRows(3);
    for (var i = firstRow; i <= rows.length; i++) if (/^【|^メンターコーディネーターとの連携/.test(rows[i - 1][0])) sh.getRange(i, 1).setFontWeight('bold');
    try { sh.showSheet(); } catch (e) {}
    return { ok: true, url: ss.getUrl() + '#gid=' + sh.getSheetId(),
             message: '「' + SM_VIEW_SHEET_ + '」のシートに、' + r.label + 'の内容を書き出しました（画面共有に使えます）。' };
  } catch (e) {
    console.error('[SM] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '書き出せませんでした: ' + (e && e.message ? e.message : e) };
  }
}
