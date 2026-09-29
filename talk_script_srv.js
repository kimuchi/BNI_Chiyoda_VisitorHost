// === トークスクリプト（定例会の台本）===
// 「【NN期】トークスクリプト生成用」シートでしていたこと（台本の文に、その回の担当者やお名前を入れる）を、
// システムの機能にしたもの。
//
//   ひな形 … 「トークスクリプト（ひな形）」シート（時間・発言者・スライド・トーク・メモ の5列）。
//            画面の「ひな形の編集」で直せる（シートで直接直してもよい）。シートが無ければ既定のひな形
//            （talk_script_default.js）を使う。
//   台本   … 開催日を選んで作ると「20261007トークスクリプト」のシートができる（同じ日を作り直すと上書き）。
//
// ひな形の {…} に、その回の内容が入る（使えるものの一覧は talkCatalog_()。画面の差し込みのボタンもこれから作る）。
//   お名前は名字だけ（同じ名字の方がいればフルネーム）で、「さん」はひな形の方に書く。
//   チームや参加者の一覧は「カテゴリーの名字さん」のように、敬称まで入れて「、」でつなぐ。
//   チェックシートの値（{ルーティン:体験談} など）は書いてあるとおり。「〇〇さん」と書いてあって、
//   ひな形の方にも「さん」が続くときは、1つにする。
// 入らなかった差し込みは「〇〇」にして、画面とシートで知らせる。チェックシートが「なし」の項目は、
// その行のメモに「なし」と書く（読み飛ばせるように）。
var TALK_TEMPLATE_SHEET_ = 'トークスクリプト（ひな形）';
var TALK_HEADERS_ = ['時間', '発言者', 'スライド', 'トーク', 'メモ'];
var TALK_FIELDS_ = ['time', 'speaker', 'slide', 'talk', 'memo'];
var TALK_OUT_SUFFIX_ = 'トークスクリプト';     // 「20261007トークスクリプト」
var TALK_BLANK_ = '〇〇';
var TALK_KEY_RE_ = /\{([^{}\n]{1,60})\}/g;

function openTalkScriptDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('talk_script').setWidth(1100).setHeight(780), 'トークスクリプト（台本）');
}

// --- ひな形 ---
// → { rows: [{ time, speaker, slide, talk, memo }], fromSheet }
function talkTemplate_() {
  var sh = getSS_().getSheetByName(TALK_TEMPLATE_SHEET_);
  if (sh && sh.getLastRow() > 1) {
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, TALK_FIELDS_.length).getValues(), rows = [];
    for (var i = 0; i < vals.length; i++) {
      var r = talkRowOf_(vals[i]);
      if (TALK_FIELDS_.some(function (f) { return r[f] !== ''; })) rows.push(r);
    }
    return { rows: rows, fromSheet: true };
  }
  return { rows: TALK_DEFAULT_ROWS_.map(talkRowOf_), fromSheet: false };
}
// シートの1行（配列）→ 行。時間の欄が時刻（Date）になっていたら「7:15」の形にする
function talkRowOf_(a) {
  var r = {};
  for (var j = 0; j < TALK_FIELDS_.length; j++) {
    var v = a && a[j] != null ? a[j] : '';
    if (Object.prototype.toString.call(v) === '[object Date]') v = v.getHours() + ':' + ('0' + v.getMinutes()).slice(-2);
    r[TALK_FIELDS_[j]] = String(v).replace(/\r\n?/g, '\n');
  }
  return r;
}
// 画面から受け取った行をそろえる（長すぎる・多すぎるものは切る）
function talkCleanRows_(rows) {
  var out = [];
  for (var i = 0; i < (rows || []).length && out.length < 500; i++) {
    var r = rows[i] || {}, o = {};
    TALK_FIELDS_.forEach(function (f, j) {
      var v = Array.isArray(r) ? r[j] : r[f];
      o[f] = String(v == null ? '' : v).replace(/\r\n?/g, '\n').slice(0, f === 'talk' || f === 'memo' ? 5000 : 200);
    });
    if (TALK_FIELDS_.some(function (f) { return o[f].trim() !== ''; })) out.push(o);
  }
  return out;
}

// --- 画面から呼ぶ ---
// 開催日の候補・ひな形・差し込みの一覧
function getTalkScriptContext() {
  try {
    var t = talkTemplate_(), sh = getSS_().getSheetByName(TALK_TEMPLATE_SHEET_);
    return { ok: true, meetings: talkMeetings_(), template: t.rows, fromSheet: t.fromSheet,
             templateSheet: TALK_TEMPLATE_SHEET_, templateUrl: sh ? talkSheetUrl_(sh) : '',
             catalog: talkCatalog_(), chapter: chapterLabel_() };
  } catch (e) {
    console.error('[TALK] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み込めませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// ひな形を保存する（シートが無ければ作る）
function saveTalkTemplate(rows) {
  try {
    var clean = talkCleanRows_(rows);
    if (!clean.length) return { ok: false, message: 'ひな形が空です。1行以上入れてください。' };
    var sh = talkWriteSheet_(TALK_TEMPLATE_SHEET_, [TALK_HEADERS_], clean, {});
    return { ok: true, message: 'ひな形を保存しました（' + clean.length + '行）。', count: clean.length, templateUrl: talkSheetUrl_(sh) };
  } catch (e) {
    console.error('[TALK] save ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '保存できませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// ひな形を既定の内容に戻す（シートに書き込む）
function resetTalkTemplate() {
  var res = saveTalkTemplate(TALK_DEFAULT_ROWS_.map(talkRowOf_));
  if (res.ok) {
    res.message = 'ひな形を既定の内容に戻しました（' + res.count + '行）。';
    res.template = TALK_DEFAULT_ROWS_.map(talkRowOf_);
  }
  return res;
}

// その開催日の差し込みの中身と、できあがりの見本。overrides … { キー: 手で直した値 }
function previewTalkScript(dateStr, overrides) {
  try {
    var d = parseDate_(dateStr);
    if (!d) return { ok: false, message: '開催日が分かりません。' };
    var built = talkBuild_(d, overrides);
    return { ok: true, date: fmtDate_(d), label: built.label, values: built.values, rows: built.rows, marks: built.marks,
             missing: built.missing, none: built.none, unknown: built.unknown, sheetName: talkOutName_(d) };
  } catch (e) {
    console.error('[TALK] preview ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '作れませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// その開催日の台本をシートに作る（同じ日のシートがあれば上書き）
function createTalkScript(dateStr, overrides) {
  try {
    var d = parseDate_(dateStr);
    if (!d) return { ok: false, message: '開催日が分かりません。' };
    var built = talkBuild_(d, overrides), name = talkOutName_(d);
    var title = built.label + '　トークスクリプト（' + chapterLabel_() + '）';
    var sh = talkWriteSheet_(name, [[title, '', '', '', ''], TALK_HEADERS_], built.rows, { marks: built.marks });
    try { if (sh.isSheetHidden()) sh.showSheet(); } catch (e) {}   // アーカイブ（非表示）してあった日も、作ったら見えるように
    var msg = '「' + name + '」を作りました（' + built.rows.length + '行）。';
    if (built.missing.length) msg += '\n入らなかった差し込み（「' + TALK_BLANK_ + '」にしました）: ' + built.missing.join('、');
    if (built.unknown.length) msg += '\nひな形の書き方が違う差し込み: ' + built.unknown.map(function (k) { return '{' + k + '}'; }).join('、');
    return { ok: true, message: msg, sheetName: name, url: talkSheetUrl_(sh),
             missing: built.missing, none: built.none, unknown: built.unknown };
  } catch (e) {
    console.error('[TALK] create ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '作れませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// --- 台本を組み立てる ---
function talkOutName_(d) { return sheetKeyOf_(d) + TALK_OUT_SUFFIX_; }

function talkBuild_(d, overrides) {
  var env = talkEnv_(d), tpl = talkTemplate_().rows, cache = {}, ov = overrides || {};
  var resolve = function (key) {
    if (!(key in cache)) {
      var r = null;
      if (Object.prototype.hasOwnProperty.call(ov, key) && String(ov[key]).trim() !== '') r = { value: String(ov[key]), state: 'ok', manual: true };
      else { try { r = talkValue_(key, env); } catch (e) { r = { value: '', state: 'missing' }; } }
      cache[key] = r;
    }
    return cache[key];
  };
  var rows = [], marks = [];
  for (var i = 0; i < tpl.length; i++) {
    var src = tpl[i], row = {}, notes = [], missing = false, none = false;
    for (var j = 0; j < TALK_FIELDS_.length; j++) {
      var f = TALK_FIELDS_[j];
      row[f] = talkFill_(src[f], resolve, function (key, r) {
        if (r.state === 'missing') missing = true;
        if (r.state === 'none') { none = true; if (notes.indexOf(key) < 0) notes.push(key); }
      });
    }
    if (notes.length) row.memo = (row.memo ? row.memo + '\n' : '') + '※チェックシートでは「なし」: ' + notes.join('、');
    rows.push(row);
    marks.push(missing ? 'missing' : (none ? 'none' : ''));
  }
  var values = [], miss = [], nones = [], unknown = [];
  Object.keys(cache).forEach(function (k) {
    var r = cache[k];
    values.push({ key: k, value: r.value, state: r.state, manual: !!r.manual });
    if (r.state === 'missing') miss.push(k);
    if (r.state === 'none') nones.push(k);
    if (r.state === 'unknown') unknown.push(k);
  });
  var no = talkMeetingNo_(env);
  return { rows: rows, marks: marks, values: values, missing: miss, none: nones, unknown: unknown,
           label: (no ? '第' + no + '回 ' : '') + talkDateLabel_(d) };
}

// 1つの欄の {…} を入れる。onUse(key, 結果) … 使った差し込みを知らせる
function talkFill_(text, resolve, onUse) {
  var s = String(text == null ? '' : text);
  return s.replace(TALK_KEY_RE_, function (m, rawKey, at) {
    var key = rawKey.trim(), r = resolve(key);
    onUse(key, r);
    if (r.state === 'unknown') return m;                        // 書き方の違う差し込みは、そのまま残す（気づけるように）
    if (r.state === 'missing') return TALK_BLANK_;
    var v = r.value, after = s.substr(at + m.length, 2);
    // 「〇〇さん」＋ひな形の「さん」、「なし」＋「さん」は、敬称を1つにする（「なし」なら付けない）
    var hon = after.match(/^(さん|様)/);
    if (hon && r.state === 'none') return v + '\u0000';          // うしろの敬称を消す目印
    if (hon && new RegExp(hon[1] + '$').test(v)) v = v.slice(0, -hon[1].length);
    return v;
  }).replace(/\u0000(さん|様)/g, '').replace(/\u0000/g, '');
}

// その開催日の材料（使うものだけ、1回ずつ読む）
function talkEnv_(d) {
  var memo = {}, key = fmtDate_(d);
  var once = function (k, f) {
    if (!(k in memo)) { try { memo[k] = f(); } catch (e) { console.warn('[TALK] ' + k + ': ' + (e && e.message ? e.message : e)); memo[k] = null; } }
    return memo[k];
  };
  var env = {
    date: d, key: key,
    members: function () { return once('members', function () { return getMemberMaster({ membersOnly: true }).members || []; }) || []; },
    holidays: function () { return once('holidays', function () { return getHolidays(); }) || []; },
    routine: function () { return once('routine', function () { var r = getRoutineInfo(key); return r && r.found ? r : null; }); },
    grid: function () {
      return once('grid', function () {
        var hit = findRoutineColumn_(d);
        return hit ? { grid: routineSheetGrid_(hit.name, hit.sheet) || [], col: hit.col } : null;
      });
    },
    holders: function () { return once('holders', function () { return roleHolders_(d); }) || {}; },
    teams: function () { return once('teams', function () { return roleTeams_(d); }) || []; },
    people: function () {
      return once('people', function () {
        var name = sheetKeyOf_(d) + '参加者';
        if (!getSS_().getSheetByName(name)) {
          // 移行前の4桁の名前（0923参加者）しか無い回。ほかの機能と同じく、そちらを読む（その開催日のものに限る）
          var old = Utilities.formatDate(d, 'Asia/Tokyo', 'MMdd'), od = meetingDateFromKey_(old);
          if (!getSS_().getSheetByName(old + '参加者') || !od || sheetKeyOf_(od) !== sheetKeyOf_(d)) return null;
          name = old + '参加者';
        }
        var data = loadSheetData(name), out = [];
        for (var i = 0; i < data.rows.length; i++) {
          var p = visitorPostPerson_(data.rows[i]);
          if (p.name && !p.cancelled) out.push(p);
        }
        return out;
      });
    },
    rotation: function () { return once('rotation', function () { var r = getSpeakerRotationWeeks(key); return r && r.ok ? r.weeks : null; }); },
    renewal: function () { return once('renewal', function () { var r = computeRenewalLists(key); return r && r.ok ? r : null; }); },
    weekly: function () { return once('weekly', function () { return talkWeeklyStart_(d, env.members()); }); }
  };
  return env;
}

// 差し込み1つの値 → { value, state: 'ok'|'none'|'missing'|'unknown' }
function talkValue_(key, env) {
  var ok = function (v) { v = String(v == null ? '' : v).trim(); return v ? { value: v, state: 'ok' } : { value: '', state: 'missing' }; };
  var c = chapterInfo_(), m, i;
  switch (key) {
    case 'チャプター': return ok(chapterLabel_());
    case 'チャプター名': return ok(c.name);
    case 'リージョン': return ok(c.region);
    case '回': return ok(talkMeetingNo_(env));
    case '開催日': return ok(talkDateLabel_(env.date));
    case '期': return ok(roleTermOf_(env.date));
    case '名簿の先頭': return ok(env.members().length ? talkShort_(env.members()[0].name, env.members()) : '');
    case 'ビジターの人数': case 'ゲストの人数': case '代理の人数':
      return env.people() ? { value: String(talkPeople_(env, { 'ビジターの人数': 'visitor', 'ゲストの人数': 'guest', '代理の人数': 'sub' }[key]).length), state: 'ok' }
                          : { value: '', state: 'missing' };
    case 'ビジターの一覧': return talkPeopleText_(env, 'visitor');
    case 'ゲストの一覧': return talkPeopleText_(env, 'guest');
    case '代理の一覧': return talkPeopleText_(env, 'sub');
    case 'メインプレゼン': case 'メインプレゼン1': case 'メインプレゼン2': {
      var ri = env.routine(), ps = ri ? (ri.mainPresenters || []) : [];
      if (ri && !ps.length && ri.mainPresentersRaw && talkIsNone_(ri.mainPresentersRaw)) return { value: 'なし', state: 'none' };
      var names = ps.map(function (p) { return p.name ? talkShort_(p.name, env.members()) : talkBare_(p.raw); });
      if (key === 'メインプレゼン') return ok(names.map(function (n) { return n + 'さん'; }).join('・'));
      return ok(names[key === 'メインプレゼン1' ? 0 : 1] || '');
    }
    case 'スタートアッププレゼン': case '2分30秒プレゼン': {       // 「2分30秒プレゼン」は前の呼び名
      var r2 = env.routine();
      if (!r2) return ok('');
      if (!r2.longPresenter && r2.longPresenterRaw === '' ) {
        var raw2 = talkRoutineItem_(env, ROUTINE_STARTUP_LABELS_);
        if (raw2.state === 'none') return raw2;
      }
      return ok(r2.longPresenter ? talkShort_(r2.longPresenter, env.members()) : talkBare_(r2.longPresenterRaw));
    }
    case 'ウィークリープレゼンの秒数': return ok(chapterSecondsLabel_(c.seconds.weekly));
    case 'スタートアッププレゼンの秒数': return ok(chapterSecondsLabel_(c.seconds.startup));
    case 'ビジタープレゼンの秒数': return ok(chapterSecondsLabel_(c.seconds.visitor));
    case 'リファーラル発表の秒数': return ok(chapterSecondsLabel_(c.seconds.referral));
    case 'ウィークリープレゼンの起点': {
      var w = env.weekly();
      var no = w && w.member ? String(w.member.no == null ? '' : w.member.no).replace(/\.0+$/, '').trim() : '';
      return ok(w && w.member ? (no ? no + '番 ' : '') + talkShort_(w.member.name, env.members()) : '');
    }
    case 'ウィークリープレゼンの業種区分': { var w2 = env.weekly(); return ok(w2 ? w2.block : ''); }
    case 'コアバリュー': { var r3 = env.routine(); return ok(r3 ? (r3.coreValue || r3.coreValueRaw) : ''); }
    case '推薦のことば': case '推薦のことばの件数': case '推薦のことばを書いた方': case '推薦のことばを受けた方':
      return talkRecommendations_(env, key);
    case '真正度確認の紹介者': case '真正度確認の受け手': case '真正度確認の紹介先':
      return talkAuthenticity_(env, key);
    case '前々回の開催日': return ok(talkPrevMeeting_(env, 2));
    case 'スピーカーローテーション': return talkRotation_(env);
    case '更新対象者（30日以内）': case '更新対象者（60日以内）': case '更新対象者（90日以内）': {
      var rl = env.renewal();
      if (!rl) return ok('');
      var list = rl[{ '更新対象者（30日以内）': 'd30', '更新対象者（60日以内）': 'd60', '更新対象者（90日以内）': 'd90' }[key]] || [];
      return list.length ? ok(list.map(function (x) { return talkShort_(x.name, env.members()) + 'さん'; }).join('、'))
                         : { value: 'なし', state: 'none' };
    }
  }
  if ((m = key.match(/^ルーティン[:：]\s*(.+)$/))) return talkRoutineItem_(env, m[1]);
  if ((m = key.match(/^チーム(のメンバー)?[:：]\s*(.+)$/))) return talkTeam_(env, m[2], !!m[1]);
  var cat = key.match(/^(.+)のカテゴリー$/), role = talkRoleOf_(cat ? cat[1] : key);
  if (role) {
    var holder = env.holders()[role.key] || '';
    if (!holder) return { value: '', state: 'missing' };
    return ok(cat ? talkCat_(holder, env.members()) : talkShort_(holder, env.members()));
  }
  return { value: '', state: 'unknown' };
}

// 役職の名前（「プレジデント」「書記兼会計」「VHC」なども）→ ROLE_DEFS_ の役職
function talkRoleOf_(name) {
  var t = roleNorm_(name);
  for (var i = 0; i < ROLE_DEFS_.length; i++) {
    var d = ROLE_DEFS_[i];
    if (roleNorm_(d.label) === t) return d;
    for (var k = 0; k < (d.alias || []).length; k++) if (roleNorm_(d.alias[k]) === t) return d;
  }
  return null;
}

// --- お名前の書き方 ---
// 氏名 → 名字（同じ名字の方が名簿にいればフルネーム。空白は詰める）
function talkShort_(name, members) {
  return roleShortName_(name, members).replace(/さん$/, '');
}
// 名簿のカテゴリー（業務内容）
function talkCat_(name, members) {
  var k = normName_(name);
  for (var i = 0; i < (members || []).length; i++) if (normName_(members[i].name) === k) return String(members[i].title || '').trim();
  return '';
}
// 「カテゴリーの名字さん」
function talkPerson_(name, members) {
  var cat = talkCat_(name, members);
  return (cat ? cat + 'の' : '') + talkShort_(name, members) + 'さん';
}
// 名簿に当たらなかった書き方（「〇〇さん」）から、敬称を外したもの
function talkBare_(raw) { return String(raw == null ? '' : raw).trim().replace(/(さん|様|さま)$/, ''); }

// --- それぞれの差し込み ---
function talkMeetingNo_(env) {
  var ri = env.routine();
  if (ri && ri.meetingNo) return String(ri.meetingNo);
  var n = 0;
  try { n = meetingCountOf_(env.date); } catch (e) {}
  return n ? String(n) : '';
}
function talkDateLabel_(d) {
  return (d.getMonth() + 1) + '月' + d.getDate() + '日(' + CHAPTER_WEEK_[d.getDay()] + ')';
}

// チェックシートの、その開催日の列の値（項目名は空白を除いて、完全一致 → 前方一致の順で探す）
//   label … 項目名。呼び名が揺れる項目は配列で（前にあるものほど優先）
function talkRoutineItem_(env, label) {
  var g = env.grid();
  if (!g) return { value: '', state: 'missing' };
  var want = [].concat(label).map(function (l) { return String(l).replace(/[\s　]/g, '').replace(/[※＊].*$/, ''); });
  var r = routineFindRow_(g.grid, want);
  if (r < 0) return { value: '', state: 'missing' };
  var v = String(g.grid[r][g.col - 1] == null ? '' : g.grid[r][g.col - 1]).trim();
  if (!v) return { value: '', state: 'missing' };
  if (talkIsNone_(v)) return { value: /^[-ー―－—‐・\s　]+$/.test(v) ? 'なし' : v, state: 'none' };   // 「ー」だけなら「なし」と読む
  return { value: v, state: 'ok' };
}

// チェックシートで「なし」の書き方（「なし」「―」のほか、長音の「ー」だけのもの）
function talkIsNone_(v) {
  var t = String(v == null ? '' : v).replace(/[\s　]/g, '');
  return routineIsBlank_(t) || /^[ー―－\-—‐・]+$/.test(t);
}

// チーム（「役職・チーム（半期ごと）」）。name はチームの名前。役職のチームのリーダーは、その役職の担当者
function talkTeam_(env, name, membersOnly) {
  var want = roleNorm_(name), team = null, i;
  var teams = env.teams();
  for (i = 0; i < teams.length; i++) if (roleNorm_(teams[i].name) === want) { team = teams[i]; break; }
  if (!team) {
    var role = talkRoleOf_(name);                               // 「エデュケーションコーディネーター」など役職の名前でも
    for (i = 0; role && i < teams.length; i++) if (teams[i].key === 'role:' + role.key) { team = teams[i]; break; }
  }
  if (!team) return { value: '', state: 'unknown' };
  var def = roleTeamDefOf_(team.key), holders = env.holders(), names = [];
  var leader = (def && def.role) ? (holders[def.role] || '') : (team.leader || (team.key === 'membership' ? holders.vice || '' : ''));
  if (leader && !membersOnly) names.push(leader);
  (team.members || []).forEach(function (m) { if (m.name && names.indexOf(m.name) < 0 && normName_(m.name) !== normName_(leader)) names.push(m.name); });
  if (!names.length) return { value: '', state: 'missing' };
  var ms = env.members();
  return { value: names.map(function (n) { return talkPerson_(n, ms); }).join('、'), state: 'ok' };
}

// 参加者シートの、その種別の方
function talkPeople_(env, type) {
  return (env.people() || []).filter(function (p) { return p.type === type; });
}
function talkPeopleText_(env, type) {
  if (!env.people()) return { value: '', state: 'missing' };
  var ps = talkPeople_(env, type), ms = env.members();
  if (!ps.length) return { value: 'なし', state: 'none' };
  var inv = function (p) { var who = routineMemberName_(p.inviter); return who.name ? talkShort_(who.name, ms) : talkBare_(p.inviter); };
  var nm = function (p) { return p.name + (p.kana && p.kana !== p.name ? '（' + p.kana + '）' : ''); };
  if (type === 'visitor') {
    return { value: ps.map(function (p) {
      return (p.inviter ? inv(p) + 'さんのご招待　' : '') + (p.category ? p.category + '　' : '') + nm(p) + '様';
    }).join('\n'), state: 'ok' };
  }
  if (type === 'guest') {
    return { value: ps.map(function (p) { return (p.inviter ? inv(p) + 'さんご招待の' : '') + nm(p) + '様'; }).join('、'), state: 'ok' };
  }
  return { value: ps.map(function (p) { return (p.inviter ? inv(p) + 'さんの代理として' : '') + nm(p) + '様'; }).join('、'), state: 'ok' };
}

// 推薦のことば（チェックシートの「推薦の言葉」のうち、定例会中に発表する組）
function talkRecommendations_(env, key) {
  var ri = env.routine();
  if (!ri) return { value: '', state: 'missing' };
  var all = ri.recommendations || [], ms = env.members(), raw = String(ri.recommendationsRaw || '').trim();
  var list = all.filter(function (x) { return x.when === 'during'; });
  if (!list.length) {
    if (!raw) return { value: '', state: 'missing' };                                   // まだ書いていない
    if (talkIsNone_(raw) || all.length) {                                                // 「なし」・アフターだけ
      return { value: key === '推薦のことばの件数' ? '0' : 'なし', state: 'none' };
    }
    return key === '推薦のことば' ? { value: raw, state: 'ok' } : { value: '', state: 'missing' };   // 組に分けられない書き方
  }
  var who = function (x) { return x.name ? talkShort_(x.name, ms) : talkBare_(x.raw); };
  var uniq = function (arr) { return arr.filter(function (v, i) { return v && arr.indexOf(v) === i; }); };
  if (key === '推薦のことばの件数') return { value: String(list.length), state: 'ok' };
  if (key === '推薦のことばを書いた方') return { value: uniq(list.map(function (x) { return who(x.giver) + 'さん'; })).join('・'), state: 'ok' };
  if (key === '推薦のことばを受けた方') return { value: uniq(list.map(function (x) { return who(x.receiver) + 'さん'; })).join('・'), state: 'ok' };
  return { value: list.map(function (x) { return who(x.giver) + 'さんから　' + who(x.receiver) + 'さんに　推薦の言葉をいただいております。'; }).join('\n'), state: 'ok' };
}

// 真正度確認（チェックシートの「真正度確認」。「Aさん⇒Bさん　C様」→ 紹介者 Aさん・受け手 Bさん・紹介先 C様）
function talkAuthenticity_(env, key) {
  var r = talkRoutineItem_(env, '真正度確認');
  if (r.state !== 'ok') return r;
  var parts = r.value.replace(/[\r\n]+/g, ' ').split(/[→⇒➡]/);
  if (parts.length < 2) return { value: '', state: 'missing' };
  var from = parts[0].trim(), rest = parts.slice(1).join('').trim();
  var mm = rest.match(/^(.+?(?:さん|様|さま))[\s　、,]*(.*)$/) || rest.match(/^(\S+)[\s　、,]+(.*)$/);
  var to = mm ? mm[1].trim() : rest, target = mm ? mm[2].trim() : '';
  var v = { '真正度確認の紹介者': from, '真正度確認の受け手': to, '真正度確認の紹介先': target }[key];
  return v ? { value: v, state: 'ok' } : { value: '', state: 'missing' };
}

// n 回前の開催日（休会日の週は数えない）→「9月9日」
function talkPrevMeeting_(env, n) {
  var d = new Date(env.date.getTime()), hol = env.holidays(), k = 0;
  for (var i = 0; i < 60 && k < n; i++) {
    d.setDate(d.getDate() - 7);
    if (hol.indexOf(fmtDate_(d)) < 0) k++;
  }
  return k === n ? (d.getMonth() + 1) + '月' + d.getDate() + '日' : '';
}

// 今後4回のメインプレゼン（この回の次から）→ 1回1行
function talkRotation_(env) {
  var weeks = env.rotation();
  if (!weeks) return { value: '', state: 'missing' };
  var ms = env.members(), lines = weeks.slice(1, 5).map(function (w) {
    return w.md + (w.no ? ' 第' + w.no + '回' : '') + '　' + w.people.map(function (p) { return talkShort_(p.name, ms) + 'さん'; }).join('・');
  });
  return lines.length ? { value: lines.join('\n'), state: 'ok' } : { value: '', state: 'missing' };
}

// ウィークリープレゼンの起点（その日の始まりの業種区分の、名簿でいちばん上の方）。ブレイクアウトルームの起点も同じ
function talkWeeklyStart_(d, master) {
  var members = (master || []).map(function (m) { return { name: m.name, cat: m.cat, no: m.no, blockKey: '' }; });
  var cycle = mpBlocks_(members), rows = null, holidays = [];
  try { rows = routineRowValues_(ROUTINE_WEEKLY_LABELS_); holidays = getHolidays(); } catch (e) {}
  var st = null;
  try { st = mpStartFromRoutine_(cycle, members, rows, fmtDate_(d), holidays); } catch (e) {}
  var key = (st && st.key) || mpStartFor_(cycle, fmtDate_(d)), block = '', first = null, i;
  for (i = 0; i < cycle.length; i++) if (cycle[i].gkey === key) block = cycle[i].block;
  for (i = 0; i < members.length && !first; i++) if (members[i].blockKey === key) first = members[i];
  return { block: block, member: first };
}

// --- 開催日の候補（次の4回と、前の2回）---
function talkMeetings_() {
  var out = [], c = [];
  try { c = getMeetingCandidates(); } catch (e) { c = []; }
  var hol = [];
  try { hol = getHolidays(); } catch (e) {}
  if (c.length) {
    var d = parseDate_(c[0].dateValue), past = [];
    for (var i = 0; i < 30 && past.length < 2; i++) {
      d.setDate(d.getDate() - 7);
      if (hol.indexOf(fmtDate_(d)) >= 0) continue;
      var no = 0;
      try { no = meetingCountOf_(d); } catch (e) {}
      past.unshift({ dateValue: fmtDate_(d), display: (d.getFullYear()) + '/' + (d.getMonth() + 1) + '/' + d.getDate()
                     + '(' + CHAPTER_WEEK_[d.getDay()] + ')' + (no ? ' 第' + no + '回' : '') + '（済）' });
    }
    out = past.concat(c.map(function (x) { return { dateValue: x.dateValue, display: x.display }; }));
  }
  return { list: out, next: c.length ? c[0].dateValue : '' };
}

// --- 差し込みの一覧（画面のボタンと、マニュアルの表のもと）---
function talkCatalog_() {
  var g = [];
  g.push({ group: 'チャプター・開催回', items: [
    { key: 'チャプター', desc: '「' + chapterLabel_() + '」' }, { key: 'チャプター名', desc: '「' + chapterInfo_().name + '」' },
    { key: 'リージョン', desc: 'チャプターの設定のリージョン' }, { key: '回', desc: '定例会の回数（数字）' },
    { key: '開催日', desc: '「10月7日(水)」' }, { key: '期', desc: 'その回の期（数字）' }] });
  var sec = chapterInfo_().seconds;
  g.push({ group: 'プレゼンの秒数（チャプターの設定）', items: [
    { key: 'ウィークリープレゼンの秒数', desc: '「' + chapterSecondsLabel_(sec.weekly) + '」' },
    { key: 'スタートアッププレゼンの秒数', desc: '「' + chapterSecondsLabel_(sec.startup) + '」' },
    { key: 'ビジタープレゼンの秒数', desc: '「' + chapterSecondsLabel_(sec.visitor) + '」' },
    { key: 'リファーラル発表の秒数', desc: '「' + chapterSecondsLabel_(sec.referral) + '」' }] });
  var roles = [];
  ROLE_DEFS_.forEach(function (r) {
    roles.push({ key: r.label, desc: 'その期の' + r.label + 'の名字' });
    roles.push({ key: r.label + 'のカテゴリー', desc: 'メンバー名簿のカテゴリー' });
  });
  g.push({ group: '役職（その期の担当者）', items: roles });
  var teams = [];
  try {
    roleTeams_(new Date()).forEach(function (t) {
      teams.push({ key: 'チーム:' + t.name, desc: 'リーダーとメンバー（「カテゴリーの名字さん」を「、」で）' });
      teams.push({ key: 'チームのメンバー:' + t.name, desc: 'リーダーを除いたメンバー' });
    });
  } catch (e) {}
  g.push({ group: 'チーム（役職・チーム（半期ごと））', items: teams });
  g.push({ group: '参加者（参加者シート）', items: [
    { key: 'ビジターの人数', desc: '数字' }, { key: 'ビジターの一覧', desc: '1人1行（招待者・カテゴリー・お名前）' },
    { key: 'ゲストの人数', desc: '数字' }, { key: 'ゲストの一覧', desc: '「〇〇さんご招待の〇〇様」' },
    { key: '代理の人数', desc: '数字' }, { key: '代理の一覧', desc: '「〇〇さんの代理として〇〇様」' }] });
  g.push({ group: 'チェックシート（その回の列）', items: [
    { key: 'メインプレゼン', desc: '「〇〇さん・〇〇さん」' }, { key: 'メインプレゼン1', desc: '1人目の名字' }, { key: 'メインプレゼン2', desc: '2人目の名字' },
    { key: 'スタートアッププレゼン', desc: '名字（前の呼び名 {2分30秒プレゼン} でも同じ）' }, { key: 'ウィークリープレゼンの起点', desc: '「22番 〇〇」（ブレイクアウトルームの起点も同じ）' },
    { key: 'ウィークリープレゼンの業種区分', desc: '始まりの業種区分' }, { key: 'コアバリュー', desc: '「BNI目的と概要」のコアバリュー' },
    { key: '推薦のことば', desc: '「〇〇さんから　〇〇さんに　推薦の言葉を…」を1組1行' }, { key: '推薦のことばの件数', desc: '数字' },
    { key: '推薦のことばを書いた方', desc: '「〇〇さん・〇〇さん」' }, { key: '推薦のことばを受けた方', desc: '「〇〇さん・〇〇さん」' },
    { key: '真正度確認の紹介者', desc: '「真正度確認」の欄の最初の方' }, { key: '真正度確認の受け手', desc: '矢印の先の方' },
    { key: '真正度確認の紹介先', desc: 'そのあとの「〇〇様」' }, { key: '前々回の開催日', desc: '「9月9日」' }]
    .concat(talkRoutineLabels_().map(function (l) { return { key: 'ルーティン:' + l, desc: 'チェックシートの「' + l + '」に書いてあるとおり' }; })) });
  g.push({ group: 'その他', items: [
    { key: 'スピーカーローテーション', desc: '今後4回のメインプレゼン（1回1行）' },
    { key: '更新対象者（30日以内）', desc: 'メンバー名簿の更新期限日から' }, { key: '更新対象者（60日以内）', desc: 'メンバー名簿の更新期限日から' },
    { key: '更新対象者（90日以内）', desc: 'メンバー名簿の更新期限日から' }, { key: '名簿の先頭', desc: 'メンバー名簿のいちばん上の方の名字' }] });
  return g;
}

// チェックシートの項目名（差し込みのボタンに並べる。次回の開催日の列があるシートから）
function talkRoutineLabels_() {
  var out = [], seen = {};
  try {
    var next = getMeetingCandidates()[0], hit = next ? findRoutineColumn_(parseDate_(next.dateValue)) : null;
    if (!hit) {
      var idx = routineIndex_(), keys = Object.keys(idx).sort();
      if (keys.length) hit = idx[keys[keys.length - 1]];
    }
    if (!hit) return out;
    var grid = routineSheetGrid_(hit.name, hit.sheet) || [];
    for (var r = 0; r < Math.min(grid.length, ROUTINE_SCAN_ROWS_); r++) {
      for (var c = 1; c < Math.min(ROUTINE_LABEL_COLS_, (grid[r] || []).length); c++) {
        var s = String(grid[r][c] == null ? '' : grid[r][c]).replace(/[\s　]+/g, '').replace(/[※＊].*$/, '');
        if (!s || /^[\d.]+$/.test(s) || /^(No\.?|内容|定例会開催日)$/.test(s) || s.length > 30 || seen[s]) continue;
        seen[s] = true;
        out.push(s);
      }
    }
  } catch (e) {}
  return out;
}

// --- シートに書く ---
// head … 見出しの行（配列の配列）。opts.marks … 行ごとの印（'missing' は黄色・'none' は灰色の字）
function talkWriteSheet_(name, head, rows, opts) {
  var ss = getSS_(), sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name); else sh.clear();
  var body = rows.map(function (r) {
    return TALK_FIELDS_.map(function (f) { var v = String(r[f] == null ? '' : r[f]); return /^[=+\-@]/.test(v) ? "'" + v : v; });
  });
  var all = head.concat(body);
  var rng = sh.getRange(1, 1, all.length, TALK_FIELDS_.length);
  rng.setNumberFormat('@');                                    // 「7:15」を時刻にしない
  rng.setValues(all);
  try { talkFormat_(sh, head.length, body.length, (opts && opts.marks) || []); } catch (e) { console.warn('[TALK] 書式: ' + e.message); }
  return sh;
}
function talkFormat_(sh, headRows, n, marks) {
  var widths = [60, 150, 180, 560, 300];
  for (var i = 0; i < widths.length; i++) sh.setColumnWidth(i + 1, widths[i]);
  sh.getRange(headRows, 1, 1, TALK_FIELDS_.length).setFontWeight('bold').setBackground('#f2f6ff');
  if (headRows > 1) sh.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sh.setFrozenRows(headRows);
  if (n) sh.getRange(headRows + 1, 1, n, TALK_FIELDS_.length).setWrap(true).setVerticalAlignment('top');
  for (var r = 0; r < marks.length; r++) {
    if (marks[r] === 'missing') sh.getRange(headRows + 1 + r, 1, 1, TALK_FIELDS_.length).setBackground('#fff4cc');
    else if (marks[r] === 'none') sh.getRange(headRows + 1 + r, 1, 1, TALK_FIELDS_.length).setFontColor('#8a8f98');
  }
}
function talkSheetUrl_(sh) {
  try { return getSS_().getUrl() + '#gid=' + sh.getSheetId(); } catch (e) { return ''; }
}
