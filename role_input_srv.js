// === 役職ごとの入力（定例会の準備）===
//
// ルーティンチェックシートに直接書き込む代わりに、役職ごとの画面から
// 次回の定例会について入力する。保存先はルーティンチェックシート（その開催日の列）。
//
//   項目   … ルーティンチェックシートの「担当」の列から、役職ごとに集める。
//            シートで担当を書き換えれば、画面の項目もそのとおりに変わる。
//            事前MTGの週次情報共有フォームにあってシートに無い項目
//            （今週の共有事項・お願い事項・人数など。ROLE_EXTRA_ITEMS_）は、
//            初めて保存するときに、シートの項目のいちばん下へ行を足してそこに保存する。
//   必須   … シートの「期日」の欄が入っている項目（「5日前」「2日前」「当日」など）。
//            期日を消せば必須でなくなる。数式で自動的に入る欄は入力しない。
//   初期値 … 前回の内容や、名簿・参加者シート・巡回の順番などから推定したもの。
//            画面で確かめて直し、「保存」で初めてシートに書き込まれる。
//
// 画面は role_input.html。入力状況の一覧（どの役職の必須項目が足りていないか）と、
// 役職ごとの入力を1つの画面で切り替える。

// 役職（24期の事前MTGフォームの並び）。
//   sheet … 足した行の「担当」の欄に書く名前
//   alias … シートの「担当」の欄で、この役職を指す書き方（空白・大文字小文字は無視）
//   holder … 担当者の初期値（24期の担当者。担当者は半期ごとに画面で登録する。下の「担当者（半期ごと）」）
var ROLE_DEFS_ = [
  { key: 'president', label: 'プレジデント', sheet: 'プレジ',
    alias: ['プレジ', 'プレジデント', 'P'], holder: '熊谷 龍威' },
  { key: 'vice', label: 'バイスプレジデント', sheet: 'バイス',
    alias: ['バイス', 'バイスプレジデント', 'VP'], holder: '船木 雄大' },
  { key: 'secretary', label: '書記兼会計', sheet: '書記兼会計',
    alias: ['書記兼会計', '書記', '会計'], holder: '原田 雅人' },
  { key: 'vhc', label: 'ビジターホストコーディネーター', sheet: 'VH',
    alias: ['VH', 'VHC', 'ビジターホスト', 'ビジターホストコーディネーター'], holder: '藤田 礼恵' },
  { key: 'mentor', label: 'メンターコーディネーター', sheet: 'メンターコーディネーター',
    alias: ['メンター', 'メンターコーディネーター', 'MC'], holder: '伊五澤 潤' },
  { key: 'ec', label: 'エデュケーションコーディネーター', sheet: 'EC',
    alias: ['EC', 'エデュケーション', 'エデュケーションコーディネーター'], holder: '伊東 良之' },
  { key: 'web', label: 'webマスター', sheet: 'WEB',
    alias: ['WEB', 'webマスター', 'ウェブマスター'], holder: '合川 周平' },
  { key: 'support', label: 'メンバーサポート委員', sheet: 'メンバーサポート委員',
    alias: ['メンバーサポート', 'メンバーサポート委員'], holder: '山本 登一郎' },
  { key: 'training', label: 'トレーニング委員', sheet: 'トレーニング委員',
    alias: ['トレーニング', 'トレーニング委員'], holder: '溝口 懸' },
  { key: 'event', label: 'イベント委員＆1to1促進委員', sheet: 'イベント委員＆1to1促進委員',
    alias: ['イベント', 'イベント委員', '1to1促進委員', 'イベント委員＆1to1促進委員'], holder: '田中 浩子' },
  { key: 'bcp', label: 'BCP委員', sheet: 'BCP委員',
    alias: ['BCP', 'BCP委員'], holder: '竹中 公基' },
  { key: 'spreading', label: 'スプレディング委員', sheet: 'スプレディング委員',
    alias: ['スプレディング', 'スプレディング委員', 'Spreading委員'], holder: '金子 美緒' },
  { key: 'gbc', label: 'グローバルビジネスコーディネーター', sheet: 'GBC',
    alias: ['GBC', 'グローバルビジネス', 'グローバルビジネスコーディネーター'], holder: '中込 渉' }
];
var ROLE_HOLDERS_KEY_ = 'BNI_ROLE_HOLDERS';              // 期ごとにする前の保存先（24期の担当者として読む）
var ROLE_HOLDERS_TERMS_KEY_ = 'BNI_ROLE_HOLDERS_TERMS';   // 期ごとの担当者 { '24': { president: '…', … }, … }
var ROLE_HOLDERS_BASE_TERM_ = 24;                         // ROLE_DEFS_ の holder と、前の保存先の担当者の期
var ROLE_HOLDERS_KEEP_TERMS_ = 10;                        // 保存しておく期の数（新しい方から）
var ROLE_TERM_BASE_ = { year: 2026, term: 23 };           // 2026年4月〜9月が23期

// 事前MTGのフォームにあって、ルーティンチェックシートに無い項目。
// 初めて保存するときに、シートの項目の下へ「事前MTG 共有事項」のまとまりとして行を足す。
//   due/weekday … 足した行の「期日」「曜日目安」（期日があるものが必須）
//   kind        … 'count' は人数（数字）
//   init        … 'carry' は前回の内容を初期値にする
// フォームの「欠席人数・代理人数・医療欠席人数」は、シートの「代理・欠席」の欄（お名前）から数える。
// 「ウィークリープレゼン担当者」「2分30秒プレゼン担当者」「メインプレゼン担当者」「新入会・更新・退会」
// 「リージョン参加者(お名前)」は、シートに同じ項目があるのでそちらに入れる。
var ROLE_EXTRA_GROUP_ = '事前MTG 共有事項';
var ROLE_EXTRA_NOTE_ = '「役職ごとの入力」の画面で追加した行です';
var ROLE_EXTRA_ITEMS_ = [
  { role: 'president', title: 'お願い事項（プレジデントから）', init: 'carry' },
  { role: 'vice', title: 'お願い事項（バイスプレジデントから）', init: 'carry' },
  { role: 'secretary', title: 'お願い事項（書記兼会計から）', init: 'carry' },
  { role: 'secretary', title: '直近のイベント', init: 'carry' },
  { role: 'secretary', title: '卒業コメント' },
  { role: 'vice', title: '人数：ビジター', due: '2日前', weekday: '月曜まで', kind: 'count' },
  { role: 'vice', title: '人数：ゲスト', due: '2日前', weekday: '月曜まで', kind: 'count' },
  { role: 'vice', title: '人数：見学', due: '2日前', weekday: '月曜まで', kind: 'count' },
  { role: 'vice', title: '人数：リージョン参加者', due: '2日前', weekday: '月曜まで', kind: 'count' }
].concat(ROLE_DEFS_.map(function (r) {
  return { role: r.key, title: '今週の共有事項（' + r.label + '）', due: '2日前', weekday: '月曜まで', init: 'carry' };
}));

// 初期値の推定のしかた（項目名で決める。どれにも当たらなければ前回の内容から判断する）。
// 項目名は全角・半角をそろえ（NFKC）、空白を除いてから当てる（「人数：」→「人数:」）。
//   carry … 前回の内容をそのまま（毎週ほぼ同じもの）
//   ほか  … 名簿・参加者シート・巡回の順番などから推定（roleEstimate_）
var ROLE_INIT_RULES_ = [
  { re: /^一般規定$/, est: 'policy' },
  { re: /^ウィークリープレゼン$/, est: 'weekly' },
  { re: /^更新対象者.*30日/, est: 'renew30' },
  { re: /^更新対象者.*60日/, est: 'renew60' },
  { re: /^更新対象者.*90日/, est: 'renew90' },
  { re: /^新入会$/, est: 'joined' },
  { re: /^更新式/, est: 'renewed' },
  { re: /^退会者$/, est: 'leaving' },
  { re: /^代理$/, parent: /代理・?欠席/, est: 'subs' },
  { re: /^人数:ビジター$/, est: 'visitors' },
  { re: /^人数:ゲスト$/, est: 'guests' },
  { re: /^人数:見学$/, est: 'zero' },
  { re: /^人数:リージョン参加者$/, est: 'regionCount' },
  { re: /^各担当者とポジティブな挨拶$/, est: 'greeting' },
  { re: /^メインプレゼン$/, est: 'mainRotation' },
  { re: /^エデュケーション$/, est: 'ecNext' },
  { re: /^BNI目的と概要$/, est: 'coreNext' },
  // 毎週ほぼ同じもの（「遅刻・欠席担当(7:00開始)」のように後ろに書き足しがあってもよい）
  { re: /^(メンバーシップから報告|遅刻・欠席担当|名札・バッチの注意|リファーラルの注意|審査中カテゴリー|募集カテゴリー|開放カテゴリー|ネットワーキングリーダー|リマインダー・特別報告|バイスプレジデントによる報告)/, est: 'carry' }
];

// 長い文を書く項目（入力欄を複数行にする）
// （これ以外でも、前回までの記載に改行があったり長かったりすれば複数行にする）
var ROLE_MULTILINE_RE_ = /共有事項|お願い事項|による報告|リマインダー|注意事項|お知らせ|イベント|コメント|アフター|推薦の言葉|より$|募集カテゴリー|トークスクリプト/;

// --- 画面を開く ---
function openRoleStatusDialog() { openRoleInputDialog_(''); }
// スピーカーローテーション（書記兼会計）の管理画面
function openSpeakerRotationDialog() { openRoleInputDialog_('secretary', 'rotation'); }
function openRoleInputPresident() { openRoleInputDialog_('president'); }
function openRoleInputVice() { openRoleInputDialog_('vice'); }
function openRoleInputSecretary() { openRoleInputDialog_('secretary'); }
function openRoleInputVhc() { openRoleInputDialog_('vhc'); }
function openRoleInputMentor() { openRoleInputDialog_('mentor'); }
function openRoleInputEc() { openRoleInputDialog_('ec'); }
function openRoleInputWeb() { openRoleInputDialog_('web'); }
function openRoleInputSupport() { openRoleInputDialog_('support'); }
function openRoleInputTraining() { openRoleInputDialog_('training'); }
function openRoleInputEvent() { openRoleInputDialog_('event'); }
function openRoleInputBcp() { openRoleInputDialog_('bcp'); }
function openRoleInputSpreading() { openRoleInputDialog_('spreading'); }
function openRoleInputGbc() { openRoleInputDialog_('gbc'); }

function openRoleInputDialog_(roleKey, view) {
  var def = roleDefOf_(roleKey);
  var t = HtmlService.createTemplateFromFile('role_input');
  t.params = { role: def ? def.key : '', view: view || '' };
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(960).setHeight(780),
    view === 'rotation' ? 'スピーカーローテーション（書記兼会計）'
      : (def ? def.label + 'の入力（次回の定例会）' : '役職ごとの入力・事前MTGのパワポ'));
}

function roleDefOf_(key) {
  for (var i = 0; i < ROLE_DEFS_.length; i++) if (ROLE_DEFS_[i].key === key) return ROLE_DEFS_[i];
  return null;
}

// --- 担当者（半期ごと）---
// 役職の担当者は半期ごとに変わる。期は 4月〜9月・10月〜3月（2026年9月までが23期、10月からが24期）。
// 担当者は期ごとに保存する。まだ登録していない期は、いちばん近い前の期（無ければ次の期）の担当者を使う。
// 期ごとにする前の担当者（ROLE_DEFS_ の holder と、画面で直して保存したもの）は、24期の担当者として読む
// （24期の事前MTGフォームの担当者。24期は 9/23 の定例会から引き継いでいる）。

// その日の期
function roleTermOf_(d) {
  var y = d.getFullYear(), m = d.getMonth() + 1;
  var half = y * 2 + (m >= 10 ? 1 : (m >= 4 ? 0 : -1));   // 1〜3月は前の年の10月からの期
  return half - ROLE_TERM_BASE_.year * 2 + ROLE_TERM_BASE_.term;
}
// 期の月（「2026年10月〜2027年3月」）
function roleTermLabel_(term) {
  var half = term - ROLE_TERM_BASE_.term + ROLE_TERM_BASE_.year * 2, y = Math.floor(half / 2);
  return (half % 2 === 0) ? (y + '年4月〜9月') : (y + '年10月〜' + (y + 1) + '年3月');
}

// 保存してある期ごとの担当者（{ '24': { president: '…', … }, … }）
function roleHolderTerms_() {
  var props = PropertiesService.getScriptProperties(), all = null, old = null, i;
  try { all = JSON.parse(props.getProperty(ROLE_HOLDERS_TERMS_KEY_) || 'null'); } catch (e) { all = null; }
  if (all && typeof all === 'object') return all;
  try { old = JSON.parse(props.getProperty(ROLE_HOLDERS_KEY_) || 'null'); } catch (e) { old = null; }
  var base = {};
  for (i = 0; i < ROLE_DEFS_.length; i++) {
    var k = ROLE_DEFS_[i].key;
    base[k] = (old && typeof old[k] === 'string') ? old[k] : ROLE_DEFS_[i].holder;
  }
  all = {};
  all[ROLE_HOLDERS_BASE_TERM_] = base;
  return all;
}

// ある期の担当者 → { term, label, registered（その期として保存してある）, from（使った期）, holders }
function roleHoldersOfTerm_(all, term) {
  var use = roleTermPick_(all, term), src = (use === null ? null : all[use]) || {}, out = {}, i;
  for (i = 0; i < ROLE_DEFS_.length; i++) {
    var k = ROLE_DEFS_[i].key;
    out[k] = (typeof src[k] === 'string') ? src[k] : '';
  }
  return { term: term, label: roleTermLabel_(term), registered: use === term, from: use, holders: out };
}

// 期ごとのもの（all … { 期: 中身 }）のうち、ある期に使う期：その期、無ければいちばん近い前の期、それも無ければ次の期
function roleTermPick_(all, term) {
  if (all[term]) return term;
  var terms = Object.keys(all).map(function (t) { return parseInt(t, 10); })
    .filter(function (t) { return t > 0 && all[t]; }).sort(function (a, b) { return a - b; }), i;
  for (i = terms.length - 1; i >= 0; i--) if (terms[i] < term) return terms[i];
  for (i = 0; i < terms.length; i++) if (terms[i] > term) return terms[i];
  return null;
}

// その日の担当者 { president: '…', … }（日を渡さなければ今日）
function roleHolders_(date) {
  return roleHoldersOfTerm_(roleHolderTerms_(), roleTermOf_(date || new Date())).holders;
}

// 画面の「役職・チーム」に出す期：開催日の期と、その前後の期（登録してある期も、開催日の期の前後2つまで）。
// 期ごとに、担当者（holders）とチーム（teams）を返す
function roleHolderTermList_(all, term, allTeams) {
  var teams = allTeams || roleTeamTerms_(), list = [term - 1, term, term + 1];
  Object.keys(all).concat(Object.keys(teams)).forEach(function (t) {
    var n = parseInt(t, 10);
    if (n > 0 && Math.abs(n - term) <= 2 && list.indexOf(n) < 0) list.push(n);
  });
  list.sort(function (a, b) { return a - b; });
  return list.map(function (t) { return roleTermEntry_(all, teams, t); });
}
function roleTermEntry_(all, teams, term) {
  var e = roleHoldersOfTerm_(all, term), tm = roleTeamsOfTerm_(teams, term);
  e.teams = tm.teams; e.teamsRegistered = tm.registered; e.teamsFrom = tm.from;
  return e;
}

// --- チーム（半期ごと）---
// 役職ごとに、リーダー（その役職の担当者）とサポートメンバーのチームがある（ビジターホスト・Webチームなど）。
// ほかに、メンバーシップ委員会（リーダーの初期値はバイスプレジデント）と、画面で足したチームを持てる。
// 期が替わると顔ぶれが全部替わるので、期ごとに持つ（スクリプトのプロパティ BNI_ROLE_TEAMS_24 =
// { teams: [{ key, name, leader, members: [{ name, note }] }] }。期ごとに分けるのは、1つのプロパティに入る
// 大きさに限りがあるため）。役職のチームのリーダーは担当者なので、ここには持たない。
// まだ登録していない期は、担当者と同じく、いちばん近い前の期（無ければ次の期）のものを出す
var ROLE_TEAMS_KEY_ = 'BNI_ROLE_TEAMS_';
// 既定のチーム（並びはチーム一覧の画面と同じ）。name は初期値で、画面で直せる
var ROLE_TEAM_DEFS_ = [
  { key: 'membership',    name: 'メンバーシップ委員会' },
  { key: 'role:ec',       role: 'ec',        name: 'エデュケーションコーディネーター' },
  { key: 'role:vhc',      role: 'vhc',       name: 'ビジターホスト' },
  { key: 'role:web',      role: 'web',       name: 'Webチーム' },
  { key: 'role:mentor',   role: 'mentor',    name: 'メンターコーディネーター' },
  { key: 'role:event',    role: 'event',     name: 'イベント委員＆1to1促進委員' },
  { key: 'role:support',  role: 'support',   name: 'メンバーサポート委員' },
  { key: 'role:training', role: 'training',  name: 'トレーニング委員' },
  { key: 'role:bcp',      role: 'bcp',       name: 'BCP委員' },
  { key: 'role:spreading', role: 'spreading', name: 'スプレディング委員会' },
  { key: 'role:gbc',      role: 'gbc',       name: 'グローバルビジネスコーディネーター' }
];
function roleTeamDefOf_(key) {
  for (var i = 0; i < ROLE_TEAM_DEFS_.length; i++) if (ROLE_TEAM_DEFS_[i].key === key) return ROLE_TEAM_DEFS_[i];
  return null;
}

// 保存してある期ごとのチーム（{ 24: [チーム…], … }）
function roleTeamTerms_() {
  var all = {}, props = PropertiesService.getScriptProperties().getProperties() || {};
  for (var k in props) {
    if (k.indexOf(ROLE_TEAMS_KEY_) !== 0) continue;
    var t = parseInt(k.slice(ROLE_TEAMS_KEY_.length), 10), v = null;
    if (!(t > 0)) continue;
    try { v = JSON.parse(props[k]); } catch (e) { v = null; }
    if (v && Object.prototype.toString.call(v.teams) === '[object Array]') all[t] = v.teams;
  }
  return all;
}
// ある期のチーム → { term, registered, from, teams }
function roleTeamsOfTerm_(all, term) {
  var use = roleTermPick_(all, term);
  return { term: term, registered: use === term, from: use, teams: roleTeamsNormalize_(use === null ? [] : all[use]) };
}
// チームをそろえる：既定のチーム（メンバーシップ委員会と役職ごとのチーム）はいつも既定の並びで出し、
// 足したチームはそのうしろ。名前・メンバーの空欄や重なりは除く。
// 前の形（{ name, members: [氏名] }。キーが無い）も読む：名前が既定のチームや役職の名前なら、そのチームにする
function roleTeamsNormalize_(teams) {
  var byKey = {}, custom = [], byName = {}, i, n = 0;
  ROLE_TEAM_DEFS_.forEach(function (d) {
    byName[roleNorm_(d.name)] = d.key;
    var r = d.role ? roleDefOf_(d.role) : null;
    if (r) byName[roleNorm_(r.label)] = d.key;
  });
  var clean = function (s, max) { return String(s == null ? '' : s).replace(/[\r\n]+/g, ' ').trim().slice(0, max); };
  for (i = 0; i < (teams || []).length; i++) {
    var t = teams[i] || {}, name = clean(t.name, 40), key = String(t.key || '');
    if (!roleTeamDefOf_(key)) key = byName[roleNorm_(name)] || (/^メンバーシップ/.test(name) ? 'membership' : '');
    var team = { key: key, name: name, leader: clean(t.leader, 40), members: roleTeamMembers_(t.members) };
    if (key) { if (!byKey[key]) byKey[key] = team; }
    else if (name) custom.push(team);
  }
  var out = ROLE_TEAM_DEFS_.map(function (d) {
    var t = byKey[d.key] || { key: d.key, name: '', leader: '', members: [] };
    if (!t.name) t.name = d.name;
    if (d.role) t.leader = '';                           // 役職のチームのリーダーは、その役職の担当者
    return t;
  });
  var names = {};
  out.forEach(function (t) { names[roleNorm_(t.name)] = true; });
  custom.forEach(function (t) {
    if (names[roleNorm_(t.name)] || out.length >= 40) return;
    names[roleNorm_(t.name)] = true;
    t.key = 'custom:' + (++n);
    out.push(t);
  });
  return out;
}
// メンバー [{ name, note }]（note … サブリーダー・メンターなどの役割。前の形の氏名だけの並びも読む）
function roleTeamMembers_(list) {
  var out = [], seen = {};
  for (var j = 0; j < (list || []).length && out.length < 200; j++) {
    var m = list[j], isObj = m && typeof m === 'object';
    var name = String((isObj ? m.name : m) == null ? '' : (isObj ? m.name : m)).replace(/[\r\n]+/g, ' ').trim();
    var note = String(isObj && m.note != null ? m.note : '').replace(/[\r\n]+/g, ' ').trim().slice(0, 30);
    var k = normName_(name);
    if (!name || seen[k]) continue;
    seen[k] = true;
    out.push({ name: name, note: note });
  }
  return out;
}
// その日の期のチーム
function roleTeams_(date) {
  return roleTeamsOfTerm_(roleTeamTerms_(), roleTermOf_(date || new Date())).teams;
}
// メンバー名簿の「役職」に書く、チームでの役割（例：「ビジターホスト」「ビジターホスト（サブリーダー）」
// 「エデュケーションコーディネーター（サポート）」）。チームの名前が役職の名前と同じなら、役割を（）で付ける
function roleTeamMemberLabel_(team, member) {
  var d = roleTeamDefOf_(team.key), r = d && d.role ? roleDefOf_(d.role) : null;
  if (r && roleNorm_(team.name) === roleNorm_(r.label)) return r.label + '（' + (member.note || 'サポート') + '）';
  return team.name + (member.note ? '（' + member.note + '）' : '');
}

// 割り振り（ビジターホスト・優先順位の設定）の「役職から読み取る」：今の期と次の期の、ビジターホストの顔ぶれ
// （ビジターホストのチームのメンバーと、リーダーのビジターホストコーディネーター）。
// 引継ぎの時期（期の最後の月）は次の期の方がビジターホストをするので、次の期が登録してあれば次の期を選んでおく
function getVisitorHostsFromRoles() {
  try {
    var today = new Date(), now = roleTermOf_(today), allH = roleHolderTerms_(), allT = roleTeamTerms_();
    var terms = [now, now + 1].map(function (t) {
      var h = roleHoldersOfTerm_(allH, t), tm = roleTeamsOfTerm_(allT, t), names = [], seen = {};
      var add = function (n) { var k = normName_(n); if (n && !seen[k]) { seen[k] = true; names.push(n); } };
      if (h.holders.vhc) add(h.holders.vhc);
      tm.teams.forEach(function (x) {
        if (x.key === 'role:vhc' || /ビジ(ター)?ホス/.test(x.name)) x.members.forEach(function (m) { add(m.name); });
      });
      return { term: t, label: roleTermLabel_(t), registered: tm.registered, from: tm.from, names: names };
    });
    var m = today.getMonth() + 1, handover = (m === 9 || m === 3) && terms[1].registered;
    return { ok: true, now: now, pick: handover ? now + 1 : now, handover: handover, terms: terms };
  } catch (e) {
    return { ok: false, message: '役職を読み取れませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// map … { president: '熊谷 龍威', … }。空文字は「未設定」。null なら担当者は変えない
// term … 期（数字。無ければ今日の期）。dateStr … 画面の開催日（返す一覧をその期に合わせる）
// teams … チーム [{ key, name, leader, members: [{ name, note }] }]。渡したときだけ、その期のチームとして保存する
function saveRoleHolders(map, term, dateStr, teams) {
  try {
    var t = parseInt(term, 10), props = PropertiesService.getScriptProperties(), i;
    if (!(t > 0 && t < 1000)) t = roleTermOf_(new Date());
    var all = roleHolderTerms_(), cur = roleHoldersOfTerm_(all, t).holders, out = all;
    if (map) {
      for (i = 0; i < ROLE_DEFS_.length; i++) {
        var k = ROLE_DEFS_[i].key;
        if (typeof map[k] === 'string') cur[k] = map[k].replace(/[\r\n]+/g, ' ').trim();
      }
      all[t] = cur;
      // 古い期は、新しい方から ROLE_HOLDERS_KEEP_TERMS_ 期ぶんだけ残す（保存できる大きさに限りがあるため）
      var keep = Object.keys(all).map(function (x) { return parseInt(x, 10); }).filter(function (x) { return x > 0; })
        .sort(function (a, b) { return b - a; }).slice(0, ROLE_HOLDERS_KEEP_TERMS_);
      out = {};
      for (i = 0; i < keep.length; i++) out[keep[i]] = all[keep[i]];
      props.setProperty(ROLE_HOLDERS_TERMS_KEY_, JSON.stringify(out));
    }
    if (teams && teams.length !== undefined) {
      props.setProperty(ROLE_TEAMS_KEY_ + t, JSON.stringify({ teams: roleTeamsNormalize_(teams) }));
      var tkeep = Object.keys(roleTeamTerms_()).map(function (x) { return parseInt(x, 10); })
        .sort(function (a, b) { return b - a; });
      for (i = ROLE_HOLDERS_KEEP_TERMS_; i < tkeep.length; i++) props.deleteProperty(ROLE_TEAMS_KEY_ + tkeep[i]);
    }
    // メンバー名簿の「役職」：いま名簿に反映してある期を直したときと、今の期の担当者を（期が替わってから）
    // 登録したときは、名簿も直す
    var st = roleRosterState_(), now = roleTermOf_(new Date()), roster = null;
    if (st.applied === t || (t === now && !(st.applied >= now))) {
      try { roster = roleRosterApply_(t, st); st.seen = now; roleRosterSaveState_(st); }
      catch (e) { console.warn('[ROLE] 名簿の役職を直せませんでした: ' + (e && e.message ? e.message : e)); }
    }
    var d = parseDate_(dateStr || ''), mt = d ? roleTermOf_(d) : t, allT = roleTeamTerms_();
    return { ok: true, term: t, holders: cur, current: roleTermEntry_(out, allT, mt), terms: roleHolderTermList_(out, mt, allT),
             roster: roster, rosterState: roleRosterStateView_(st),
             message: t + '期（' + roleTermLabel_(t) + '）の' + (teams ? '役職・チーム' : '担当者') + 'を保存しました。'
               + (roster && roster.changes.length ? 'メンバー名簿の「役職」も直しました（' + roster.changes.length + '名）。' : '') };
  } catch (e) {
    return { ok: false, message: '担当者の保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// --- 役職・チームを、メンバー名簿の「役職」に反映する ---
// 期が替わったら、その期の役職・チームを名簿の「役職」の列に書く。期が替わったあと、名簿・メンバーブック・
// 役職ごとの入力などを最初に開いたときに行う（roleRosterAutoSync_）。
// 「役職・チーム（半期ごと）」の「名簿の役職に反映」で、いつでも（次の期を前もってでも）反映できる。
// その期のチームを登録してあれば、役職の欄を全部その期の内容にする（期が替わると顔ぶれは全部替わるため）：
//   ・13の役職の名前と、チームでの役割（「ビジターホスト」「エデュケーションコーディネーター（サポート）」など）を
//     「・」でつなぐ（どれにも入っていない方は空欄）
// チームが未登録の期は、13の役職だけを直す：
//   ・担当者の行     … 役職を、担当している役職の名前にする
//   ・担当者でない行 … 役職が13の役職の名前だけ（前の期の担当者など）なら空欄にする
//   ・それ以外の行   … 「ビジターホスト」「WEB」など、13の役職の名前以外が入っている行は触らない
var ROLE_ROSTER_STATE_KEY_ = 'BNI_ROLE_ROSTER_STATE';   // { applied: 反映した期, at: 反映した日, seen: 確かめた期 }
var ROLE_ROSTER_CHECKED_ = false;                        // 1回の実行の中では1回だけ確かめる

function roleRosterState_() {
  var st = null;
  try { st = JSON.parse(PropertiesService.getScriptProperties().getProperty(ROLE_ROSTER_STATE_KEY_) || 'null'); }
  catch (e) { st = null; }
  return (st && typeof st === 'object') ? st : {};
}
function roleRosterSaveState_(st) {
  PropertiesService.getScriptProperties().setProperty(ROLE_ROSTER_STATE_KEY_, JSON.stringify(st));
}
// 画面に出す反映の記録
function roleRosterStateView_(st) {
  return { applied: st.applied || null, label: st.applied ? roleTermLabel_(st.applied) : '', at: st.at || '' };
}
function roleTermArg_(term) {
  var t = parseInt(term, 10);
  return (t > 0 && t < 1000) ? t : roleTermOf_(new Date());
}

// 役職の欄が、13の役職の名前だけでできているか（「バイスプレジデント」「BCP委員・スプレディング委員」など）
function roleRosterOnlyLabels_(s) {
  var labels = {}, parts = String(s == null ? '' : s).split(/[・、,，\/／\s　]+/), n = 0, i;
  for (i = 0; i < ROLE_DEFS_.length; i++) labels[roleNorm_(ROLE_DEFS_[i].label)] = true;
  for (i = 0; i < parts.length; i++) {
    if (!parts[i]) continue;
    if (!labels[roleNorm_(parts[i])]) return false;
    n++;
  }
  return n > 0;
}

// 名簿の「役職」を、ある期の役職・チームどおりにするときの変更（名簿はまだ書き換えない）
function roleRosterPlan_(term) {
  var holders = roleHoldersOfTerm_(roleHolderTerms_(), term).holders, byName = {}, i, j;
  // その期のチームを登録してあり（誰か1人はメンバーかリーダーがいる）ときは、役職の欄を全部その期の内容にする
  var tm = roleTeamsOfTerm_(roleTeamTerms_(), term);
  var full = tm.registered && tm.teams.some(function (t) { return t.members.length > 0 || !!t.leader; });
  var entry = function (nm) { var k = normName_(nm); return (byName[k] = byName[k] || { name: nm, labels: [], found: false, role: false }); };
  for (i = 0; i < ROLE_DEFS_.length; i++) {
    var nm = holders[ROLE_DEFS_[i].key];
    if (nm) { var e = entry(nm); e.labels.push(ROLE_DEFS_[i].label); e.role = true; }
  }
  if (full) {
    for (i = 0; i < tm.teams.length; i++) {
      var t = tm.teams[i];
      // 役職のチームでないチーム（メンバーシップ委員会・足したチーム）のリーダーは「（リーダー）」。役職のある方は役職だけ
      if (t.leader && !entry(t.leader).role) entry(t.leader).labels.push(t.name + '（リーダー）');
      for (j = 0; j < t.members.length; j++) entry(t.members[j].name).labels.push(roleTeamMemberLabel_(t, t.members[j]));
    }
  }
  var sh = ensureMemberSheet_(), col = MEMBER_HEADERS_.indexOf('役職') + 1;
  var data = sh.getDataRange().getValues(), changes = [], column = [];
  for (var r = 1; r < data.length; r++) {
    var name = String(data[r][2] == null ? '' : data[r][2]).trim();
    var cur = String(data[r][col - 1] == null ? '' : data[r][col - 1]).trim(), to = cur;
    var h = name ? byName[normName_(name)] : null;
    if (h) { h.found = true; to = h.labels.join('・'); }
    else if (name && cur && (full || roleRosterOnlyLabels_(cur))) to = '';
    column.push([to !== cur ? to : data[r][col - 1]]);
    if (to !== cur) changes.push({ name: name, from: cur, to: to });
  }
  var notFound = [];
  for (var key in byName) if (!byName[key].found) notFound.push(byName[key].name);
  return { term: term, label: roleTermLabel_(term), full: full, sheet: sh, col: col, column: column, changes: changes, notFound: notFound };
}
function roleRosterResult_(plan, applied) {
  var n = plan.changes.length, msg = '';
  if (applied) {
    var what = plan.full ? '役職・チーム' : '担当者';
    msg = n ? 'メンバー名簿の「役職」を、' + plan.term + '期の' + what + 'に合わせて直しました（' + n + '名）。'
            : 'メンバー名簿の「役職」は、' + plan.term + '期の' + what + 'どおりでした。';
  }
  if (plan.notFound.length) msg += (msg ? '\n' : '') + '名簿に見つからない方：' + plan.notFound.join('、');
  return { ok: true, term: plan.term, label: plan.label, full: plan.full, changes: plan.changes, notFound: plan.notFound, message: msg };
}
// 反映する。st（反映の記録）に、反映した期と日を書く（保存は呼ぶ側）
function roleRosterApply_(term, st) {
  var plan = roleRosterPlan_(term);
  if (plan.changes.length) plan.sheet.getRange(2, plan.col, plan.column.length, 1).setValues(plan.column);
  st.applied = term;
  st.at = fmtDate_(new Date());
  return roleRosterResult_(plan, true);
}

// 期が替わったら、その期の担当者を名簿の「役職」に反映する（名簿・役職ごとの入力などを開いたときに呼ぶ）。
// 同じ期のうちは、2回目からは何もしない。その期の担当者がまだ登録されていなければ反映しない
// （登録して保存したときに反映する）。次の期の担当者を前もって反映してあれば、そのまま
function roleRosterAutoSync_() {
  if (ROLE_ROSTER_CHECKED_) return null;
  ROLE_ROSTER_CHECKED_ = true;
  try {
    var now = roleTermOf_(new Date()), st = roleRosterState_(), res = null;
    if (st.seen === now) return null;
    var ready = roleHoldersOfTerm_(roleHolderTerms_(), now).registered || roleTeamsOfTerm_(roleTeamTerms_(), now).registered;
    if (ready && !(st.applied >= now)) {
      res = roleRosterApply_(now, st);
      console.log('[ROLE] 期が替わったので反映: ' + res.message);
    }
    st.seen = now;
    roleRosterSaveState_(st);
    return res;
  } catch (e) {
    console.warn('[ROLE] 担当者をメンバー名簿の役職に反映できませんでした: ' + (e && e.message ? e.message : e));
    return null;
  }
}

// 名簿を取り込んだあと（Spreading・メンバーリスト(OCR)）：役職を、反映してある期の役職・チームに合わせ直す。
// 取り込んだ役職が前の期のままでも、「役職・チーム（半期ごと）」の登録どおりになる。お知らせの文を返す
function roleRosterAfterImport_() {
  try {
    var st = roleRosterState_();
    if (!st.applied) return '';
    var res = roleRosterApply_(st.applied, st);
    roleRosterSaveState_(st);
    return res.changes.length
      ? '\n\n役職は、「役職・チーム（半期ごと）」の' + st.applied + '期の内容に合わせ直しました（' + res.changes.length + '名）。'
      : '';
  } catch (e) {
    console.warn('[ROLE] 取り込みのあと、役職を合わせ直せませんでした: ' + (e && e.message ? e.message : e));
    return '';
  }
}

// 画面から：名簿の「役職」を、ある期の担当者どおりにしたときの変更を確かめる（書き換えない）
function previewRoleHoldersRoster(term) {
  try {
    return roleRosterResult_(roleRosterPlan_(roleTermArg_(term)), false);
  } catch (e) {
    return { ok: false, message: 'メンバー名簿を読めませんでした: ' + (e && e.message ? e.message : e) };
  }
}
// 画面から：名簿の「役職」を、ある期の担当者どおりにする
function applyRoleHoldersToRoster(term) {
  try {
    var st = roleRosterState_(), res = roleRosterApply_(roleTermArg_(term), st);
    st.seen = roleTermOf_(new Date());
    roleRosterSaveState_(st);
    res.state = roleRosterStateView_(st);
    return res;
  } catch (e) {
    return { ok: false, message: 'メンバー名簿の役職を直せませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// --- シートの読み取り ---
// 見比べ用（id・担当の書き方）：全角半角をそろえ、空白を除き、英字は小文字
function roleNorm_(s) {
  return String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]/g, '').toLowerCase();
}
// 項目名を ROLE_INIT_RULES_ に当てるとき：全角半角をそろえ、空白を除く
function roleMatchText_(s) {
  return String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]/g, '');
}

// 「担当」の欄 → 役職のキー（複数書いてあれば全部）。分からない書き方は unknown に入れる
function roleKeysOf_(raw) {
  var parts = String(raw == null ? '' : raw).split(/[・、,，\/／\n]/), keys = [], unknown = [];
  for (var i = 0; i < parts.length; i++) {
    var p = roleNorm_(parts[i]);
    if (!p) continue;
    var hit = '';
    for (var r = 0; r < ROLE_DEFS_.length && !hit; r++) {
      var al = ROLE_DEFS_[r].alias.concat([ROLE_DEFS_[r].label, ROLE_DEFS_[r].sheet]);
      for (var a = 0; a < al.length; a++) if (roleNorm_(al[a]) === p) { hit = ROLE_DEFS_[r].key; break; }
    }
    if (hit) { if (keys.indexOf(hit) < 0) keys.push(hit); }
    else unknown.push(String(parts[i]).trim());
  }
  return { keys: keys, unknown: unknown };
}

// セルの値 → 文字（日付は yyyy/MM/dd）
function roleCellText_(v) {
  if (v == null) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return fmtDate_(v);
  return String(v).replace(/\r\n?/g, '\n').replace(/\s+$/, '');
}

// 「5日前」「当日」→ 何日前か。読めなければ null
function roleDueDays_(due) {
  var s = String(due == null ? '' : due).normalize('NFKC');
  var m = s.match(/(\d+)\s*日前/);
  if (m) return parseInt(m[1], 10);
  if (/当日/.test(s)) return 0;
  if (/前日/.test(s)) return 1;
  return null;
}

// シート1枚ぶんの項目の一覧。見出しの行（「担当」「期日」…）を探し、その下を読む。
//   B列 No.／C〜E列 項目名（C＞D＞E の順に細かくなる）／担当／期日／曜日目安／備考
// 戻り値 { header, cols: {role, due, weekday, note}, items: [...], unknown: [...] }
function roleSheetLayout_(grid) {
  var hr = -1, cols = {};
  for (var r = 0; r < Math.min(12, grid.length) && hr < 0; r++) {
    for (var c = 0; c < grid[r].length; c++) {
      if (roleNorm_(grid[r][c]) === '担当') { hr = r; break; }
    }
  }
  if (hr < 0) return null;
  for (var c2 = 0; c2 < grid[hr].length; c2++) {
    var h = roleNorm_(grid[hr][c2]);
    if (h === '担当') cols.role = c2;
    else if (h === '期日') cols.due = c2;
    else if (h === '曜日目安') cols.weekday = c2;
    else if (h === '備考') cols.note = c2;
  }
  // 項目名は「内容」の列から「担当」の手前まで（通常 C〜E 列）
  cols.label0 = 2;
  for (var c3 = 0; c3 < cols.role; c3++) if (roleNorm_(grid[hr][c3]) === '内容') { cols.label0 = c3; break; }
  return { header: hr, cols: cols };
}

function roleSheetItems_(grid) {
  var lay = roleSheetLayout_(grid);
  if (!lay) return null;
  var cols = lay.cols, items = [], unknown = [], seen = {}, g1 = '', g2 = '', lastUsed = lay.header;
  var width = Math.max(cols.role, cols.due || 0, cols.weekday || 0, cols.note || 0) + 1;
  var cell = function (r, c) { return c == null ? '' : String(grid[r][c] == null ? '' : grid[r][c]).replace(/[\r\n]+/g, ' ').replace(/[\s　]+/g, ' ').trim(); };
  for (var r = lay.header + 1; r < grid.length; r++) {
    for (var u = 0; u < width; u++) if (cell(r, u)) { lastUsed = r; break; }   // 行を足すときの目安（見出し〜備考の列）
    // C＞D＞E の順に細かくなる。いちばん細かいものが項目名、1つ上がその親、C列がまとまりの名前
    var labels = [];
    for (var c = cols.label0; c < cols.role; c++) labels.push(cell(r, c));
    var level = -1;
    for (var k = labels.length - 1; k >= 0; k--) if (labels[k]) { level = k; break; }
    if (labels[0]) { g1 = labels[0]; g2 = ''; }
    if (labels[1]) g2 = labels[1];
    if (level < 0) continue;
    var roleRaw = cell(r, cols.role);
    if (!roleRaw) continue;                                   // 担当の無い行（見出しなど）は項目にしない
    var title = labels[level];
    var group = level === 0 ? '' : g1;
    var parent = level >= 2 ? g2 : '';
    var rk = roleKeysOf_(roleRaw);
    var id = roleNorm_(group) + '>' + roleNorm_(parent) + '>' + roleNorm_(title);
    seen[id] = (seen[id] || 0) + 1;
    if (seen[id] > 1) id += '#' + seen[id];
    var note = cell(r, cols.note);
    items.push({ id: id, row: r, group: group, parent: parent, title: title,
                 roleRaw: roleRaw, roles: rk.keys, due: cell(r, cols.due), weekday: cell(r, cols.weekday),
                 note: String(grid[r][cols.note] == null ? '' : grid[r][cols.note]).trim(),
                 // 備考に「ここに自動生成されます」とある欄（トークスクリプト）は、入力しない。
                 // 「※自動生成の為、全文を記載」（この欄から自動で作る）は入力する欄
                 auto: /ここに自動|自動生成されます|自動で入ります/.test(note) });
    if (rk.unknown.length) unknown.push({ title: title, parent: parent, roleRaw: roleRaw });
  }
  return { header: lay.header, cols: cols, items: items, unknown: unknown, lastUsed: lastUsed };
}

// 足す行（ROLE_EXTRA_ITEMS_）の id。シートに行ができたあとと同じ形にしておく
function roleExtraId_(title) { return roleNorm_(ROLE_EXTRA_GROUP_) + '>>' + roleNorm_(title); }

// --- 読み込み（一覧・役職ごとの入力で使う）---
// シートは1回の実行の中で使い回す
function roleSheetCache_() {
  var cache = {};
  return function (entry) {
    var name = entry.name;
    if (cache[name]) return cache[name];
    var sh = entry.sheet, rows = sh.getLastRow(), last = sh.getLastColumn();
    var grid = (rows && last) ? sh.getRange(1, 1, rows, last).getValues() : [];
    var parsed = roleSheetItems_(grid);
    var byId = {};
    if (parsed) for (var i = 0; i < parsed.items.length; i++) byId[parsed.items[i].id] = parsed.items[i];
    cache[name] = { sheet: sh, grid: grid, parsed: parsed, byId: byId };
    return cache[name];
  };
}

// 開催日の選択肢：これからの開催日（休会日を除く）
function roleMeetingChoices_() {
  var out = [], c = getMeetingCandidates(), w = ['日', '月', '火', '水', '木', '金', '土'];
  for (var i = 0; i < c.length; i++) {
    var d = parseDate_(c[i].dateValue);
    var no = (c[i].display.match(/第(\d+)回/) || [])[1] || '';
    out.push({ dateValue: c[i].dateValue, no: no,
               display: (d.getMonth() + 1) + '/' + d.getDate() + '(' + w[d.getDay()] + ')' + (no ? ' 第' + no + '回' : '') });
  }
  return out;
}

function roleMd_(d) {
  var w = ['日', '月', '火', '水', '木', '金', '土'];
  return (d.getMonth() + 1) + '/' + d.getDate() + '(' + w[d.getDay()] + ')';
}

// 画面の初期表示：開催日 → 役職ごとの項目・いまの値・前回の値・入力状況。
// roleKey を渡すと、その役職の項目だけ初期値を推定する（名簿や参加者シートを読むので少し時間がかかる。
// 入力状況の一覧だけなら推定は要らないので、一覧はすぐ出る）。'*' は全役職（検査用）
function getRoleInputContext(dateStr, roleKey) {
  try {
    routineResetCache_();                                    // シートを足した直後などに備え、毎回読み直す
    var sync = roleRosterAutoSync_();                        // 期が替わっていたら、担当者を名簿の「役職」に反映
    var meetings = roleMeetingChoices_();
    var date = dateStr || (meetings[0] ? meetings[0].dateValue : '');
    var target = parseDate_(date);
    if (!target) return { ok: false, message: '開催日が分かりません。' };
    var ctx = roleBuildContext_(target, roleKey || '');
    ctx.meetings = meetings;
    ctx.rosterState = roleRosterStateView_(roleRosterState_());
    if (sync && sync.changes.length) ctx.rosterSync = '期が替わったので、' + sync.message;
    return ctx;
  } catch (e) {
    console.error('[ROLE] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

function roleBuildContext_(target, roleKey) {
  var key = fmtDate_(target), idx = routineIndexFor_(key), hit = idx[key] || null;
  var today = new Date(); today.setHours(0, 0, 0, 0);
  // 担当者は、この開催日の期（半期）のもの
  var holderAll = roleHolderTerms_(), teamAll = roleTeamTerms_(), holderTerm = roleTermEntry_(holderAll, teamAll, roleTermOf_(target));
  var holders = holderTerm.holders, sheetOf = roleSheetCache_();
  var no = '';
  var ctx = { ok: true, date: key, display: roleMd_(target), today: fmtDate_(today), estimatedFor: roleKey || '',
              found: !!hit, sheetName: hit ? hit.name : '', roles: [], items: {}, order: [], unknown: [],
              holderTerm: holderTerm, holderTerms: roleHolderTermList_(holderAll, holderTerm.term, teamAll),
              members: [], version: (typeof SYSTEM_VERSION_ === 'string') ? SYSTEM_VERSION_ : '' };
  try { ctx.members = (getMemberMaster({ membersOnly: true }).members || []).map(function (m) { return m.name; }); } catch (e) {}
  if (!hit) {
    ctx.message = 'ルーティンチェックシートに ' + key + ' の列が見つかりません。'
      + '「【NN期】ルーティンチェックシート」の1行目に開催日があるか確かめてください。';
    ctx.roles = ROLE_DEFS_.map(function (r) {
      return { key: r.key, label: r.label, holder: holders[r.key] || '', items: [],
               status: { state: 'none', required: 0, filled: 0, missing: [] } };
    });
    return ctx;
  }
  var cur = sheetOf(hit);
  if (!cur.parsed) {
    ctx.found = false;
    ctx.message = '「' + hit.name + '」に「担当」「期日」の見出しの行が見つかりません。';
    return ctx;
  }
  try {
    var rno = routineFindRow_(cur.grid, ['定例会回数']);
    if (rno >= 0) no = String(cur.grid[rno][hit.col - 1] == null ? '' : cur.grid[rno][hit.col - 1]).replace(/\.0$/, '').trim();
  } catch (e) {}
  if (!/^\d+$/.test(no)) { var cnt = meetingCountOf_(target); no = cnt ? String(cnt) : ''; }
  ctx.meetingNo = no;

  // 数式の入っている欄（自動で入るので入力しない）
  var formulas = [];
  try { formulas = cur.sheet.getRange(1, hit.col, cur.grid.length, 1).getFormulas(); } catch (e) {}

  // 前回までの開催日（この日より前、新しい順に6回まで。休会日の空の列も含む）
  var prevDates = Object.keys(idx).filter(function (d) { return d < key; }).sort().reverse().slice(0, 6);

  var items = cur.parsed.items.slice(), present = {};
  for (var i = 0; i < items.length; i++) present[items[i].id] = true;
  // シートにまだ無い「事前MTG 共有事項」の項目も並べる（保存したときに行を足す）
  for (var x = 0; x < ROLE_EXTRA_ITEMS_.length; x++) {
    var ex = ROLE_EXTRA_ITEMS_[x], eid = roleExtraId_(ex.title);
    if (present[eid]) continue;
    items.push({ id: eid, row: -1, group: ROLE_EXTRA_GROUP_, parent: '', title: ex.title,
                 roleRaw: roleDefOf_(ex.role).sheet, roles: [ex.role], due: ex.due || '', weekday: ex.weekday || '',
                 note: '', auto: false, virtual: true });
  }

  for (var j = 0; j < items.length; j++) {
    var it = items[j];
    var raw = it.row >= 0 ? cur.grid[it.row][hit.col - 1] : '';
    var value = roleCellText_(raw);
    var hasFormula = it.row >= 0 && formulas[it.row] && formulas[it.row][0];
    var hist = [];
    for (var p = 0; p < prevDates.length && hist.length < 4; p++) {
      var e2 = idx[prevDates[p]], s2 = sheetOf(e2), pit = s2.byId[it.id];
      if (!pit) continue;
      var pv = roleCellText_(s2.grid[pit.row][e2.col - 1]);
      if (pv) hist.push({ date: prevDates[p], value: pv });
    }
    var dueDays = roleDueDays_(it.due);
    var dueDate = null;
    if (dueDays !== null) { dueDate = new Date(target.getTime()); dueDate.setDate(dueDate.getDate() - dueDays); }
    var extra = null;
    for (var y = 0; y < ROLE_EXTRA_ITEMS_.length; y++) {
      if (roleExtraId_(ROLE_EXTRA_ITEMS_[y].title) === it.id) { extra = ROLE_EXTRA_ITEMS_[y]; break; }
    }
    // 「済」の印の欄（スライドを入れたか等）。短い値に「済」が続けて書かれているもの
    var checks = hist.filter(function (h) { return /済/.test(h.value) && h.value.length <= 40; }).length;
    var isCheck = /スライド画像/.test(it.title) || (hist.length > 0 && checks >= Math.min(2, hist.length));
    var readOnly = !!(hasFormula || it.auto);
    ctx.items[it.id] = {
      id: it.id, roles: it.roles, roleRaw: it.roleRaw, group: it.group, parent: it.parent, title: it.title,
      due: it.due, weekday: it.weekday, note: it.note,
      dueDate: dueDate ? fmtDate_(dueDate) : '', dueLabel: dueDate ? roleMd_(dueDate) : '',
      required: !!it.due && !readOnly, readOnly: readOnly,
      readOnlyWhy: hasFormula ? '数式で自動的に入ります' : (it.auto ? '自動で作られます（' + it.note + '）' : ''),
      kind: (extra && extra.kind) || (isCheck ? 'check' : ''),
      multiline: ROLE_MULTILINE_RE_.test(it.title) || /\n/.test(value) || hist.some(function (h) { return /\n/.test(h.value) || h.value.length > 40; }),
      value: value, orig: value, prev: hist[0] || null, history: hist, virtual: !!it.virtual,
      init: extra ? (extra.init || '') : ''
    };
    ctx.order.push(it.id);
  }
  ctx.unknown = cur.parsed.unknown;

  // 初期値の推定（ほかの項目の値を使うものがあるので、全部そろってから。開いている役職の項目だけ）
  if (roleKey) {
    var env = roleEstimateEnv_(target, prevDates, ctx);
    for (var q = 0; q < ctx.order.length; q++) {
      var item = ctx.items[ctx.order[q]];
      if (item.readOnly || (roleKey !== '*' && item.roles.indexOf(roleKey) < 0)) continue;
      try { item.estimate = roleEstimate_(item, env); } catch (e) { item.estimate = null; }
    }
  }

  // 役職ごとの入力状況
  for (var r = 0; r < ROLE_DEFS_.length; r++) {
    var def = ROLE_DEFS_[r], ids = [];
    for (var z = 0; z < ctx.order.length; z++) if (ctx.items[ctx.order[z]].roles.indexOf(def.key) >= 0) ids.push(ctx.order[z]);
    ctx.roles.push({ key: def.key, label: def.label, holder: holders[def.key] || '', items: ids,
                     status: roleStatus_(ids, ctx.items, today) });
  }
  return ctx;
}

// 役職の入力状況。必須（期日あり）の項目がシートに入っているか
//   state … done（そろっている）／todo（足りない）／late（期日を過ぎて足りない）／none（必須の項目が無い）
function roleStatus_(ids, items, today) {
  var st = { required: 0, filled: 0, missing: [], state: 'none', nextDue: '' };
  var late = false;
  for (var i = 0; i < ids.length; i++) {
    var it = items[ids[i]];
    if (!it.required) continue;
    st.required++;
    if (String(it.value || '').trim()) { st.filled++; continue; }
    var over = !!it.dueDate && parseDate_(it.dueDate).getTime() < today.getTime();
    if (over) late = true;
    st.missing.push({ id: it.id, title: it.title, parent: it.parent, dueLabel: it.dueLabel, late: over });
    if (it.dueDate && (!st.nextDue || it.dueDate < st.nextDue)) st.nextDue = it.dueDate;
  }
  if (st.nextDue) st.nextDueLabel = roleMd_(parseDate_(st.nextDue));
  st.state = !st.required ? 'none' : (!st.missing.length ? 'done' : (late ? 'late' : 'todo'));
  return st;
}

// --- 初期値の推定 ---
// 名簿・参加者シートなどは、必要になったときに1回だけ読む
function roleEstimateEnv_(target, prevDates, ctx) {
  var memo = {};
  var once = function (k, fn) {
    if (!(k in memo)) { try { memo[k] = fn(); } catch (e) { console.warn('[ROLE] ' + k + ': ' + (e && e.message ? e.message : e)); memo[k] = null; } }
    return memo[k];
  };
  var itemByTitle = function (title) {
    for (var i = 0; i < ctx.order.length; i++) {
      var it = ctx.items[ctx.order[i]];
      if (roleMatchText_(it.title) === roleMatchText_(title)) return it;
    }
    return null;
  };
  return {
    target: target, date: fmtDate_(target), prevDates: prevDates, itemByTitle: itemByTitle,
    members: function () { return once('members', function () { return getMemberMaster({ membersOnly: true }).members || []; }); },
    holidays: function () { return once('holidays', function () { return getHolidays(); }); },
    renewal: function () { return once('renewal', function () { var L = computeRenewalLists(fmtDate_(target)); return L && L.ok ? L : null; }); },
    participants: function () {
      return once('participants', function () {
        var name = sheetKeyOf_(target) + '参加者';
        if (!getSS_().getSheetByName(name)) return null;
        var data = loadSheetData(name), people = [];
        for (var i = 0; i < data.rows.length; i++) {
          var p = visitorPostPerson_(data.rows[i]);
          if (p.name && !p.cancelled) people.push(p);
        }
        return { sheet: name, people: people };
      });
    },
    routineInfo: function () { return once('routineInfo', function () { return getRoutineInfo(fmtDate_(target)); }); },
    rotation: function () { return once('rotation', function () { return rotPairFor_(target); }); }
  };
}

// 氏名 → 「名字さん」（同じ名字の方がいれば「田中秀一さん」）
function roleShortName_(name, members) {
  var full = String(name || '').replace(/[\s　]+/g, ' ').trim();
  if (!full) return '';
  var sur = full.split(' ')[0], dup = 0;
  for (var i = 0; i < (members || []).length; i++) {
    if (String(members[i].name || '').replace(/[\s　]+/g, ' ').trim().split(' ')[0] === sur) dup++;
  }
  return (dup > 1 ? full.replace(/ /g, '') : sur) + 'さん';
}
function roleNamesText_(names, members, sep) {
  if (!names.length) return 'なし';
  return names.map(function (n) { return roleShortName_(n, members); }).join(sep || '、');
}

// 1つの項目の初期値。戻り値 { value, source, mode: 'prefill'|'suggest' } または null
//   prefill … 入力欄に入れておく（保存すればシートに入る）
//   suggest … 入力欄の下に「推定」として出すだけ（押せば入る）
function roleEstimate_(it, env) {
  var t = roleMatchText_(it.title), p = roleMatchText_(it.parent), rule = null, i;
  for (i = 0; i < ROLE_INIT_RULES_.length; i++) {
    var ru = ROLE_INIT_RULES_[i];
    if (ru.re.test(t) && (!ru.parent || ru.parent.test(p))) { rule = ru; break; }
  }
  var prev = it.prev, prevDate = prev ? roleMd_(parseDate_(prev.date)) : '';
  var carry = function (why) {
    return prev ? { value: prev.value, source: why || ('前回（' + prevDate + '）と同じ'), mode: 'prefill', basis: 'carry' } : null;
  };
  // スピーカーローテーションの表は前半スライドで作るようになったので、画像の用意は要らない
  if (/^スピーカーローテーション用スライド画像/.test(t)) {
    return { value: '済', source: '前半スライドの表はスピーカーローテーションから作るので、画像は要りません', mode: 'prefill' };
  }
  if (it.kind === 'check') return null;                     // 「済」の印は引き継がない
  if (it.init === 'carry') return carry();
  var est = rule ? rule.est : '';
  var members, L, P, m, names, k;

  switch (est) {
    case 'carry':
      return carry();
    case 'policy': {                                        // 一般規定は開催ごとに1つずつ進む（1〜12番）
      if (!prev) return null;
      var n = routinePolicyNo_(prev.value);
      if (!n) return null;
      var steps = mpMeetingsBetween_(parseDate_(prev.date), env.target, env.holidays() || []);
      var nx = ((n - 1 + Math.max(steps, 1)) % 12) + 1;
      return { value: nx + '番', source: '前回（' + prevDate + '）の ' + n + '番 の次', mode: 'prefill' };
    }
    case 'weekly': {                                        // ウィークリープレゼンの始まり（業種区分の巡回）
      members = (env.members() || []).map(function (x) {
        return { name: x.name, no: x.no, cat: x.cat, blockKey: '' };
      });
      if (!members.length) return null;
      var cycle = mpBlocks_(members), rows = {};
      try { rows = routineRowValues_(ROUTINE_WEEKLY_LABELS_); } catch (e) {}
      var st = mpStartFromRoutine_(cycle, members, rows, env.date, env.holidays() || []);
      if (st && st.from === 'routine') return null;           // その日の記載がもうある
      var gk = (st && st.key) || mpStartFor_(cycle, env.date);
      var blk = null;
      for (i = 0; i < cycle.length; i++) if (cycle[i].gkey === gk) blk = cycle[i];
      if (!blk) return null;
      var first = null;
      for (i = 0; i < members.length; i++) {
        if (members[i].blockKey !== gk) continue;
        if (!first || (parseFloat(members[i].no) || 9999) < (parseFloat(first.no) || 9999)) first = members[i];
      }
      var head = blk.block + (first ? '　' + String(parseFloat(first.no) || '') + '番　' + roleShortName_(first.name, members) : '');
      // 前回の記載から数えられなかったとき（記載の区分が業種区分マスタに無いなど）は、開催日からの計算
      var why = (st && st.key && st.from === 'previous') ? ('前回（' + roleMd_(parseDate_(st.date)) + '）の記載から、業種区分を' + st.steps + 'つ進めて')
              : '業種区分の巡回（開催日からの計算）から';
      return { value: head.replace(/　$/, ''), source: why, mode: 'prefill' };
    }
    case 'renew30': case 'renew60': case 'renew90': {       // 名簿の更新期限日から
      L = env.renewal(); members = env.members() || [];
      if (!L || !members.length || L.noDate.length >= members.length) return null;
      var list = L[est === 'renew30' ? 'd30' : (est === 'renew60' ? 'd60' : 'd90')] || [];
      return { value: roleNamesText_(list.map(function (x) { return x.name; }), members),
               source: '名簿の更新期限日から（残り' + est.slice(-2) + '日以内）', mode: 'prefill' };
    }
    case 'joined': case 'renewed': {                        // 名簿の入会日・更新日が、前回の開催日より後
      members = env.members() || [];
      var fld = est === 'joined' ? 'joinDate' : 'renewDate';
      if (!members.some(function (x) { return x[fld]; })) return null;
      var from = env.prevDates[0] ? parseDate_(env.prevDates[0]) : new Date(env.target.getTime() - 7 * 86400000);
      names = [];
      for (i = 0; i < members.length; i++) {
        var d = parseDate_(members[i][fld]);
        if (d && d.getTime() > from.getTime() && d.getTime() <= env.target.getTime()) names.push(members[i].name);
      }
      return { value: roleNamesText_(names, members),
               source: '名簿の' + (est === 'joined' ? '入会日' : '更新日') + 'から（前回の開催日より後）', mode: 'prefill' };
    }
    case 'leaving': {                                       // 更新状況で「更新しない」にした方
      var marks = getRenewalMarks_(); members = env.members() || []; names = [];
      for (i = 0; i < members.length; i++) {
        m = marks[normName_(members[i].name)];
        if (m && m.status === 'leaving') names.push(members[i].name);
      }
      if (!names.length) return null;
      return { value: roleNamesText_(names, members), source: '更新状況で「更新しない」にした方', mode: 'suggest' };
    }
    case 'subs': case 'visitors': case 'guests': {          // 参加者シート（その開催日のもの）から
      P = env.participants();
      if (!P) return null;
      members = env.members() || [];
      var type = est === 'subs' ? 'sub' : (est === 'visitors' ? 'visitor' : 'guest');
      var hits = P.people.filter(function (x) { return x.type === type; });
      var src = '参加者シート（' + P.sheet + '）から';
      if (est !== 'subs') return { value: String(hits.length), source: src, mode: 'prefill' };
      names = [];
      for (i = 0; i < hits.length; i++) {
        var who = routineMemberName_(hits[i].inviter);
        var nm = who.name || hits[i].inviter;
        if (nm && names.indexOf(nm) < 0) names.push(nm);
      }
      return { value: roleNamesText_(names, members), source: src + '（代理を立てる方）', mode: 'prefill' };
    }
    case 'zero':
      return carry() || { value: '0', source: 'これまでの記載が無いので0', mode: 'prefill', basis: 'carry' };
    case 'regionCount': {                                   // 「リージョン参加者」の欄の人数
      var reg = env.itemByTitle('リージョン参加者');
      var rv = reg ? (reg.value || (reg.estimate && reg.estimate.value) || '') : '';
      if (!rv) return null;
      var cntR = routineIsBlank_(rv.replace(/[\s　]/g, '')) ? 0
        : rv.split(/[、,，・\/／\s　]+/).filter(function (s) { return s.trim(); }).length;
      return { value: String(cntR), source: '「リージョン参加者」の欄（' + rv + '）から', mode: 'prefill' };
    }
    case 'greeting': {                                      // その日のメインプレゼンのお2人
      var ri = env.routineInfo();
      members = env.members() || [];
      var ns = [], gsrc = 'メインプレゼンのお2人';
      if (ri && ri.found && (ri.mainPresenters || []).length) {
        ns = ri.mainPresenters.map(function (x) {
          return x.name ? roleShortName_(x.name, members) : String(x.raw || '').replace(/(さん)?$/, 'さん');
        });
      } else {
        ns = (env.rotation() || []).map(function (n) { return roleShortName_(n, members); });
        gsrc = 'スピーカーローテーションのお2人';
      }
      ns = ns.filter(function (s) { return s && s !== 'さん'; });
      if (!ns.length) return null;
      return { value: ns.join('　'), source: gsrc, mode: 'prefill' };
    }
    case 'mainRotation': {                                  // スピーカーローテーション（書記兼会計が管理）の2名
      var pair = env.rotation() || [];
      if (pair.length < 2) return null;
      members = env.members() || [];
      return { value: '①' + roleShortName_(pair[0], members) + '　②' + roleShortName_(pair[1], members),
               source: 'スピーカーローテーションから', mode: 'prefill' };
    }
    case 'ecNext': case 'coreNext': {                       // ECの前回の共有事項「次週のエデュケーション：○○さん（コアバリュー）」
      var ec = env.itemByTitle('今週の共有事項（エデュケーションコーディネーター）');
      var txt = ec && ec.prev ? ec.prev.value : '';
      var line = '';
      String(txt).split(/\n/).forEach(function (l) { if (!line && /次週/.test(l) && /エデュケーション/.test(l)) line = l; });
      if (!line) return est === 'coreNext' ? null : null;
      var body = line.replace(/^.*?エデュケーション[^：:]*[：:]\s*/, '');
      if (est === 'ecNext') {
        var nm2 = body.replace(/[（(].*$/, '').trim();
        return nm2 ? { value: nm2, source: 'ECの前回の共有事項（' + line.trim() + '）から', mode: 'prefill' } : null;
      }
      var cv = coreValueOf_((body.match(/[（(]([^）)]*)[）)]/) || [])[1] || body);
      return cv ? { value: cv.label, source: 'ECの前回の共有事項（' + line.trim() + '）から', mode: 'prefill' } : null;
    }
  }
  // 決まりの無い項目：前回までの書き方から判断する
  var h = it.history || [];
  if (!h.length) return null;
  if (h.length >= 2 && roleNorm_(h[0].value) === roleNorm_(h[1].value)) return carry('前回・前々回と同じ');
  if (routineIsBlank_(roleNorm_(h[0].value)) && /^(なし|無し)/.test(roleNorm_(h[0].value))) {
    return { value: 'なし', source: '前回（' + prevDate + '）も「なし」', mode: 'prefill', basis: 'carry' };
  }
  return null;
}

// --- 保存 ---
// entries … [{ id, value, orig }]。orig は画面を開いたときのシートの値。
// その後にほかの人がシートを書き換えていたら、上書きせずにお知らせする。
function saveRoleInput(dateStr, roleKey, entries) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) {
    return { ok: false, message: 'ほかの方が保存中です。少し待ってからもう一度お試しください。' };
  }
  try {
    routineResetCache_();
    var target = parseDate_(dateStr);
    if (!target) return { ok: false, message: '開催日が分かりません。' };
    var hit = findRoutineColumn_(target);
    if (!hit) return { ok: false, message: 'ルーティンチェックシートに ' + fmtDate_(target) + ' の列が見つかりません。' };
    var sh = hit.sheet, list = entries || [];
    var readGrid = function () {
      var rows = sh.getLastRow(), last = sh.getLastColumn();
      return sh.getRange(1, 1, rows, last).getValues();
    };
    var grid = readGrid(), parsed = roleSheetItems_(grid);
    if (!parsed) return { ok: false, message: '「' + hit.name + '」に「担当」「期日」の見出しの行が見つかりません。' };
    var byId = {}, i;
    for (i = 0; i < parsed.items.length; i++) byId[parsed.items[i].id] = parsed.items[i];

    // シートにまだ無い項目（事前MTG 共有事項）に書くときは、先に行を足す
    var needExtra = false;
    for (i = 0; i < list.length; i++) {
      if (!byId[list[i].id] && String(list[i].value || '').trim() && roleIsExtraId_(list[i].id)) needExtra = true;
    }
    var added = 0;
    if (needExtra) {
      added = roleAddExtraRows_(sh, parsed, byId);
      grid = readGrid(); parsed = roleSheetItems_(grid); byId = {};
      for (i = 0; i < parsed.items.length; i++) byId[parsed.items[i].id] = parsed.items[i];
    }

    var formulas = sh.getRange(1, hit.col, grid.length, 1).getFormulas();
    var saved = 0, skipped = [];
    for (i = 0; i < list.length; i++) {
      var en = list[i], it = byId[en.id];
      var value = String(en.value == null ? '' : en.value).replace(/\r\n?/g, '\n').replace(/\s+$/, '');
      if (!it) {
        if (value) skipped.push({ id: en.id, title: roleTitleOfId_(en.id), why: 'シートに項目の行が見つかりません' });
        continue;
      }
      if (roleDefOf_(roleKey) && it.roles.indexOf(roleKey) < 0) {
        skipped.push({ id: en.id, title: it.title, why: 'この役職の項目ではありません（担当: ' + it.roleRaw + '）' });
        continue;
      }
      if (formulas[it.row] && formulas[it.row][0]) { skipped.push({ id: en.id, title: it.title, why: '数式が入っている欄です' }); continue; }
      var now = roleCellText_(grid[it.row][hit.col - 1]);
      if (value === now) continue;
      var orig = String(en.orig == null ? '' : en.orig).replace(/\r\n?/g, '\n').replace(/\s+$/, '');
      if (now !== orig) {
        skipped.push({ id: en.id, title: it.title, why: '画面を開いたあとに、シートが「' + now.slice(0, 30) + '」に書き換えられていました' });
        continue;
      }
      var cellR = sh.getRange(it.row + 1, hit.col);
      if (/^\s*\d{1,4}[\/\-.]\d{1,2}([\/\-.]\d{1,4})?\s*$/.test(value)) cellR.setNumberFormat('@');   // 日付に化けないように
      cellR.setValue(value);
      saved++;
    }
    routineResetCache_();
    try { if (SpreadsheetApp.flush) SpreadsheetApp.flush(); } catch (e) {}
    try { lock.releaseLock(); } catch (e) {}
    var msg = saved ? ('保存しました（' + saved + '項目）。') : '変更はありませんでした。';
    if (added) msg += '\nルーティンチェックシートに「' + ROLE_EXTRA_GROUP_ + '」の行（' + added + '行）を足しました。';
    if (skipped.length) {
      msg += '\n保存しなかった項目: ' + skipped.map(function (s) { return s.title + '（' + s.why + '）'; }).join('、');
    }
    var ctx = roleBuildContext_(target, roleDefOf_(roleKey) ? roleKey : '');
    ctx.meetings = roleMeetingChoices_();
    return { ok: true, saved: saved, skipped: skipped, added: added, message: msg, context: ctx };
  } catch (e) {
    console.error('[ROLE] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function roleIsExtraId_(id) {
  for (var i = 0; i < ROLE_EXTRA_ITEMS_.length; i++) if (roleExtraId_(ROLE_EXTRA_ITEMS_[i].title) === id) return true;
  return false;
}
function roleTitleOfId_(id) { return String(id || '').split('>').pop().replace(/#\d+$/, ''); }

// 「事前MTG 共有事項」の行を、シートの項目のいちばん下（1行あけて）に足す。
// 書くのは項目名・担当・期日・曜日目安・備考の列だけ（開催日の列には触らない）。
// すでにある項目は足さない。戻り値は足した行数
function roleAddExtraRows_(sh, parsed, byId) {
  var cols = parsed.cols, todo = [], i;
  for (i = 0; i < ROLE_EXTRA_ITEMS_.length; i++) if (!byId[roleExtraId_(ROLE_EXTRA_ITEMS_[i].title)]) todo.push(ROLE_EXTRA_ITEMS_[i]);
  if (!todo.length) return 0;
  var hasGroup = false;
  for (i = 0; i < parsed.items.length; i++) if (parsed.items[i].group === ROLE_EXTRA_GROUP_) hasGroup = true;
  var width = Math.max(cols.role, cols.due || 0, cols.weekday || 0, cols.note || 0) + 1;
  var rows = [];
  var blank = function () { var r = []; for (var c = 0; c < width; c++) r.push(''); return r; };
  if (!hasGroup) {
    var hr = blank();
    hr[cols.label0] = ROLE_EXTRA_GROUP_;
    if (cols.note != null) hr[cols.note] = ROLE_EXTRA_NOTE_;
    rows.push(hr);
  }
  for (i = 0; i < todo.length; i++) {
    var r = blank();
    r[cols.label0 + 1] = todo[i].title;
    r[cols.role] = roleDefOf_(todo[i].role).sheet;
    if (cols.due != null) r[cols.due] = todo[i].due || '';
    if (cols.weekday != null) r[cols.weekday] = todo[i].weekday || '';
    rows.push(r);
  }
  var start = parsed.lastUsed + 2 + 1;                       // 1行あけた次の行（1から数える）
  var need = start + rows.length - 1 - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
  sh.getRange(start, 1, rows.length, width).setValues(rows);
  return rows.length;
}
