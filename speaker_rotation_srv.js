// === スピーカーローテーション（メインプレゼンの順番。書記兼会計の業務）===
//
// 「MP選出・ローテーション管理ツール」（1つのHTMLファイル）でしていたことを、このシステムに移したもの。
//
//   並び順（order）… ローテーションの順番に並べたメンバー
//   対象外（excluded）… 並びにはいるが、いまは発表しない方（三役など）
//   起点（anchor）… { date, pointer }。date の開催日は、並びの pointer 番目（0から）から2名
//
// 開催ごとに、並びの次の2名を割り当てる。対象外の方と、メンバー名簿にいない方は飛ばす。
// 休会日（⚙️ 設定 > 休会日）の週は飛ばす。並びの最後まで行ったら最初に戻る。
// ルーティンチェックシートに「メインプレゼン」が書いてある回は、そちらを優先して表示する（確定した回）。
//
// 並び順を直すと、起点の日から先の割り当てが変わる。画面を開くときに、起点を
// 「ルーティンチェックシートのメインプレゼンがまだ空の、いちばん近い開催日」まで進めておくので、
// 確定した回（書記兼会計がもう準備した回）は変わらない。
//
// 使うところ
//   ・役職ごとの入力（書記兼会計）… ローテーションの管理画面、メインプレゼンの初期値
//   ・定例会スライド（前半）… 「スピーカーローテーション」のページの表（5回分。以前は書記兼会計が作った画像）

var ROT_KEY_ = 'BNI_SPEAKER_ROTATION';
var ROT_SLIDE_WEEKS_ = 5;                       // スライドの表に載せる回数
var ROT_LIST_WEEKS_ = 12;                       // 管理画面に並べる回数

// 初期値（まだ保存していないとき）。並びはメンバー名簿の順・対象外なし・起点は次の開催日の先頭から（rotLoad_）。
// ツール（MP選出・ローテーション管理ツール）の状態は、画面の「ツールから取り込む」で入れる。
// メンバーの氏名はコードに書かない（名簿はスプレッドシートにだけ置く）
var ROT_DEFAULT_ = {
  header: 'メインプレゼンテーション（各４分45秒）',
  notes: ['２週間前までに、略歴書と発表用のデータ・資料を書記兼会計までご提出ください。',
          '当日、スピーカーからのご提供で包装した商品（目安1,000～2,000円程度）をお持ちください。',
          '発表者は、「ビジネスビルダー　MSアドオンプログラム」の受講が必須条件となります。',
          '発表者の方のパワーチームに繋がりそうな方をビジターとしてご招待ください。']
};

function rotNorm_(s) { return String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]+/g, ''); }

function rotLoad_() {
  var raw = null, st = null;
  try { raw = PropertiesService.getScriptProperties().getProperty(ROT_KEY_); } catch (e) {}
  try { st = raw ? JSON.parse(raw) : null; } catch (e) { st = null; }
  var d = JSON.parse(JSON.stringify(ROT_DEFAULT_));
  st = st || {};
  if (!Array.isArray(st.order)) st.order = rotRosterOrder_();
  if (!Array.isArray(st.excluded)) st.excluded = [];
  if (!st.anchor || !parseDate_(st.anchor.date)) {
    var holidays = [];
    try { holidays = getHolidays(); } catch (e) {}
    st.anchor = { date: fmtDate_(rotNextMeeting_(holidays)), pointer: 0 };
    // まだ何も保存していないとき：起点（最初に見たときの次の開催日の先頭から）を1回だけ残しておく。
    // 以前は見るたびに「次の開催日から」になり、同じ回の発表者が、見た週によって変わっていた
    // （前半スライドの表・書記兼会計の初期値・告知の文）。並び順はこれまでどおり名簿の順のまま
    if (!raw) {
      try { PropertiesService.getScriptProperties().setProperty(ROT_KEY_, JSON.stringify({ anchor: st.anchor, provisional: true })); }
      catch (e) { console.warn('[ROT] 起点を残せませんでした: ' + (e && e.message ? e.message : e)); }
    }
  }
  st.anchor.pointer = parseInt(st.anchor.pointer, 10) || 0;
  if (typeof st.header !== 'string') st.header = d.header;
  if (!Array.isArray(st.notes)) st.notes = d.notes;
  st.updated = st.updated || '';
  return st;
}

// まだ並びを保存していないときの並び：メンバー名簿の順
function rotRosterOrder_() {
  try {
    return (getMemberMaster({ membersOnly: true }).members || []).map(function (m) { return String(m.name || '').trim(); })
      .filter(function (n) { return !!n; });
  } catch (e) { return []; }
}

function rotSave_(st) {
  st.updated = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(ROT_KEY_, JSON.stringify({
    order: st.order, excluded: st.excluded, anchor: st.anchor, header: st.header, notes: st.notes, updated: st.updated
  }));
  return st;
}

// 飛ばす方：対象外の方。
// メンバー名簿にいない方（退会・名簿の登録漏れ）は自動では飛ばさない（登録漏れで順番がずれないように）。
// 管理画面で「名簿にいない」と知らせるので、退会された方は並びから外してもらう
function rotSkip_(st) {
  var skip = {};
  for (var i = 0; i < (st.excluded || []).length; i++) skip[rotNorm_(st.excluded[i])] = 'excluded';
  return skip;
}

// 並びの start 番目から、count 組ぶん2名ずつ取る → { pairs, next }（next は続きの位置）
function rotPairsFrom_(st, skip, start, count) {
  var order = st.order, n = order.length, pairs = [];
  if (!n) return { pairs: [], next: 0 };
  var idx = ((start % n) + n) % n, guard = 0, limit = n * 4 + 10;
  while (pairs.length < count && guard < limit) {
    var picked = [];
    while (picked.length < 2 && guard < limit) {
      if (!skip[rotNorm_(order[idx])]) picked.push(order[idx]);
      idx = (idx + 1) % n; guard++;
    }
    if (picked.length === 2) pairs.push(picked); else break;
  }
  return { pairs: pairs, next: idx };
}

// 1組さかのぼった位置（ツールの「前の週へ」と同じ）
function rotPrevPointer_(st, skip, pointer) {
  var order = st.order, n = order.length;
  if (!n) return pointer;
  var idx = ((pointer % n) + n) % n, found = [], guard = 0;
  while (found.length < 2 && guard < n * 4 + 10) {
    idx = (idx - 1 + n) % n;
    if (!skip[rotNorm_(order[idx])]) found.push(idx);
    guard++;
  }
  return found.length === 2 ? found[1] : pointer;
}

// 起点の日そのものが休会日なら、その次の開催日を起点とみなす
function rotAnchorDate_(st, holidays) {
  var a = parseDate_(st.anchor.date);
  for (var i = 0; i < 60 && holidays.indexOf(fmtDate_(a)) >= 0; i++) a.setDate(a.getDate() + 7);
  return a;
}

// 開催日 d の2名が並びの何番目から始まるか
function rotPointerAt_(st, skip, d, holidays) {
  var a = rotAnchorDate_(st, holidays), p = st.anchor.pointer || 0, k, i;
  if (d.getTime() >= a.getTime()) {
    k = mpMeetingsBetween_(a, d, holidays);
    return k ? rotPairsFrom_(st, skip, p, k).next : p;
  }
  k = mpMeetingsBetween_(d, a, holidays);
  for (i = 0; i < k; i++) p = rotPrevPointer_(st, skip, p);
  return p;
}

// from（その日を含む）から、休会日を除いた開催日を count 回ぶん
function rotMeetingDates_(from, count, holidays) {
  var out = [], d = new Date(from.getTime());
  for (var i = 0; i < 400 && out.length < count; i++) {
    if (holidays.indexOf(fmtDate_(d)) < 0) out.push(new Date(d.getTime()));
    d.setDate(d.getDate() + 7);
  }
  return out;
}

// 次の開催日（きょう以降で休会日でない回。getMeetingCandidates と同じ数え方）。
// 休会日は呼ぶ側で読んだものを使う（休会日のシートを何度も読まないように）
function rotNextMeeting_(holidays) {
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var d = chapterMeetingBase_().date;                         // チャプターの設定の基準の開催日（その曜日に毎週）
  for (var i = 0; i < 3000; i++) {
    if (d.getTime() >= today.getTime() && holidays.indexOf(fmtDate_(d)) < 0) return d;
    d.setDate(d.getDate() + 7);
  }
  d = new Date(today.getTime());
  while (d.getDay() !== chapterWeekday_()) d.setDate(d.getDate() + 1);
  return rotMeetingDates_(d, 1, holidays)[0];
}

// ルーティンチェックシートの「メインプレゼン」→ { 'yyyy/MM/dd': [氏名, 氏名] }（書いてある回＝確定）
//   fromKey … この日より前の回しか載っていないシート（昔の期）は読まない
function rotRoutineMains_(fromKey) {
  var out = {}, rows = {};
  try { rows = routineRowValues_(['メインプレゼン'], fromKey); } catch (e) { rows = {}; }
  for (var k in rows) {
    var ps = routineMainPresenters_(rows[k]);
    if (ps.length) out[k] = ps.map(function (p) { return p.name || p.raw; });
  }
  return out;
}
function rotRoutineNumbers_(fromKey) {
  try { return routineRowValues_(['定例会回数'], fromKey); } catch (e) { return {}; }
}
// ルーティンチェックシートに列がある開催日（{ 'yyyy/MM/dd': true }）
function rotRoutineDates_() {
  var out = {};
  try { var idx = routineIndex_(); for (var k in idx) out[k] = true; } catch (e) {}
  return out;
}

// 起点を「メインプレゼンがまだ空の、いちばん近い開催日」まで進める（割り当ては変わらない）。
// 並び順を直したとき、確定した回が動かないように
function rotRebase_(st, skip, holidays, mains, next) {
  var dates = rotMeetingDates_(next, 60, holidays), open = next;
  for (var i = 0; i < dates.length; i++) { if (!mains[fmtDate_(dates[i])]) { open = dates[i]; break; } }
  var a = rotAnchorDate_(st, holidays);
  if (open.getTime() <= a.getTime()) return { rebased: false, open: open };
  st.anchor = { date: fmtDate_(open), pointer: rotFirstActive_(st, skip, rotPointerAt_(st, skip, open, holidays)) };
  return { rebased: true, open: open };
}

// p から先で、対象外でない最初の位置（起点の「ここから」を分かりやすくするため）
function rotFirstActive_(st, skip, p) {
  var n = st.order.length;
  for (var i = 0; i < n; i++) {
    var k = (p + i) % n;
    if (!skip[rotNorm_(st.order[k])]) return k;
  }
  return p;
}

// 開催日ごとの2名（count 回ぶん）。ルーティンに書いてある回はそちらを優先する
function rotWeeks_(st, from, count, env) {
  var dates = rotMeetingDates_(from, count, env.holidays), out = [];
  var byName = {};
  for (var i = 0; i < env.members.length; i++) byName[rotNorm_(env.members[i].name)] = env.members[i];
  var w = ['日', '月', '火', '水', '木', '金', '土'];
  // 開催回：最初の回はルーティンチェックシートの「定例会回数」（無ければ数える）。あとは休会日を除いて1つずつ
  // （ルーティンチェックシートの先の回の数字は、書き間違いが残っていることがあるため使わない）
  var base = dates.length ? parseInt(String(env.numbers[fmtDate_(dates[0])] || '').replace(/\.0$/, ''), 10) : 0;
  if (!(base > 0) && dates.length) base = meetingCountOf_(dates[0]) || 0;
  for (var j = 0; j < dates.length; j++) {
    var d = dates[j], key = fmtDate_(d), names, source;
    if (env.mains[key]) { names = env.mains[key].slice(0, 2); source = 'routine'; }
    else {
      var p = rotPairsFrom_(st, env.skip, rotPointerAt_(st, env.skip, d, env.holidays), 1).pairs[0];
      names = p ? p.slice() : []; source = 'rotation';
    }
    var no = base > 0 ? String(base + j) : '';
    out.push({
      date: key, no: no, md: (d.getMonth() + 1) + '/' + d.getDate() + '(' + w[d.getDay()] + ')',
      label: (d.getMonth() + 1) + '月' + d.getDate() + '日', source: source,
      people: names.map(function (nm) {
        var m = byName[rotNorm_(nm)] || null;
        return { name: m ? m.name : nm, title: m ? (m.title || '') : '', collab: m ? (m.collab || '') : '',
                 company: m ? (m.company || '') : '', inMaster: !!m };
      })
    });
  }
  return out;
}

// 割り当てに要るもの。ルーティンチェックシートは from（その日）以降の回が載っているシートだけ読む。
//   from … Date。無ければ次の開催日
function rotEnv_(from) {
  var members = [];
  try { members = getMemberMaster({ membersOnly: true }).members || []; } catch (e) {}
  var holidays = [];
  try { holidays = getHolidays(); } catch (e) {}
  var start = from || rotNextMeeting_(holidays), key = fmtDate_(start);
  return { members: members, holidays: holidays, start: start,
           mains: rotRoutineMains_(key), numbers: rotRoutineNumbers_(key) };
}

// 次回のメインプレゼンターのご案内（Facebookに投稿する文。ツールと同じ文面）
function rotFbText_(week, secretary) {
  if (!week || week.people.length < 2) return '（次回のメインプレゼンターが決まっていないため作れません）';
  var p1 = week.people[0], p2 = week.people[1];
  return '＜【' + week.label + '定例会】メインプレゼンターのご案内＞\n'
    + '本日も、定例会お疲れ様でした。\n'
    + '次回の定例会は' + week.label + 'になります。\n\n'
    + '➊次回のメインプレゼンターをご紹介致します。\n'
    + '（１）' + p1.name + 'さん/' + (p1.title || '') + '\n'
    + '　　　→つながりたい人：' + (p1.collab || '') + '\n'
    + '（２）' + p2.name + 'さん/' + (p2.title || '') + '\n'
    + '　　　→つながりたい人：' + (p2.collab || '') + '\n'
    + 'スピーカーの繋がりたい方を、是非ビジターとしてお呼びしましょう！\n'
    + '➋今後4週間のメイン・プレゼンターは添付画像の通りです。\n'
    + '➌推薦のことばをお待ちしています。（特にメインプレゼンター宛！）\n'
    + '毎週月曜日の正午までに、PDFにて書記兼会計の' + (secretary || '（書記兼会計）') + 'までご提出下さい';
}

// --- 画面から呼ぶ ---
// 管理画面の中身：並び順・対象外・起点・これからの予定・Facebookの文
function getSpeakerRotation() {
  try {
    routineResetCache_();
    var st = rotLoad_(), env = rotEnv_();
    env.skip = rotSkip_(st);
    var next = env.start;
    var rb = rotRebase_(st, env.skip, env.holidays, env.mains, next);
    var weeks = rotWeeks_(st, next, ROT_LIST_WEEKS_, env);
    var inOrder = {}, inMaster = {}, missing = [];
    for (var i = 0; i < st.order.length; i++) inOrder[rotNorm_(st.order[i])] = true;
    for (var j = 0; j < env.members.length; j++) inMaster[rotNorm_(env.members[j].name)] = true;
    if (env.members.length) st.order.forEach(function (n) { if (!inMaster[rotNorm_(n)]) missing.push(n); });
    // 案内文の「書記兼会計の○○まで」は、その回の期（半期）の書記兼会計
    try {
      var hAll = roleHolderTerms_();
      weeks.forEach(function (w) { w.secretary = roleHoldersOfTerm_(hAll, roleTermOf_(parseDate_(w.date))).holders.secretary || ''; });
    } catch (e) {}
    // ルーティンチェックシートで開催回が空の日（休会の予定）が、休会日に入っていなければ知らせる
    var hints = [], limit = new Date(next.getTime()); limit.setDate(limit.getDate() + 7 * 30);
    Object.keys(rotRoutineDates_()).sort().forEach(function (k) {
      var dd = parseDate_(k);
      if (!dd || dd.getTime() < next.getTime() || dd.getTime() > limit.getTime()) return;
      if (!String(env.numbers[k] || '').trim() && env.holidays.indexOf(k) < 0) hints.push(rotMd_(dd));
    });
    return {
      ok: true, order: st.order, excluded: st.excluded, anchor: st.anchor, header: st.header, notes: st.notes,
      updated: st.updated, provisional: !st.updated, rebased: rb.rebased, openDate: fmtDate_(rb.open),
      missing: missing, holidayHints: hints, weeks: weeks, fbText: rotFbText_(weeks[0], (weeks[0] || {}).secretary || ''),
      secretary: (weeks[0] || {}).secretary || '',
      chapter: (function () { try { return chapterLabel_(); } catch (e) { return ''; } })(),   // メインプレゼンターの画像の下の帯
      members: env.members.map(function (m) {
        return { name: m.name, title: m.title || '', collab: m.collab || '', inOrder: !!inOrder[rotNorm_(m.name)] };
      }),
      holidays: env.holidays
    };
  } catch (e) {
    console.error('[ROT] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'スピーカーローテーションの読み込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 画面で直した並び順などを保存する。base は画面を開いたときの updated（ほかの方が先に保存していたら止める）
function saveSpeakerRotation(data) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok: false, message: 'ほかの方が保存中です。少し待ってからもう一度お試しください。' }; }
  try {
    var cur = rotLoad_(), d = data || {};
    if ((d.base || '') !== (cur.updated || '')) {
      return { ok: false, conflict: true, message: '画面を開いたあとに、ほかの方がローテーションを保存しています。開き直してからもう一度直してください。' };
    }
    var seen = {}, order = [];
    (d.order || []).forEach(function (n) {
      var s = String(n == null ? '' : n).replace(/[\r\n]+/g, ' ').trim(), k = rotNorm_(s);
      if (s && !seen[k]) { seen[k] = true; order.push(s); }
    });
    if (!order.length) return { ok: false, message: '並び順が空です。' };
    var excluded = (d.excluded || []).filter(function (n) { return seen[rotNorm_(n)]; });
    var date = parseDate_(d.anchor && d.anchor.date);
    if (!date) return { ok: false, message: '起点の開催日が分かりません。' };
    var pointer = parseInt(d.anchor.pointer, 10);
    if (!(pointer >= 0)) pointer = 0;
    pointer = pointer % order.length;
    var st = { order: order, excluded: excluded, anchor: { date: fmtDate_(date), pointer: pointer },
               header: typeof d.header === 'string' ? d.header.trim() : cur.header,
               notes: Array.isArray(d.notes) ? d.notes.map(function (s) { return String(s).trim(); }).filter(function (s) { return s; }) : cur.notes };
    rotSave_(st);
    var res = getSpeakerRotation();
    res.message = 'スピーカーローテーションを保存しました。';
    return res;
  } catch (e) {
    console.error('[ROT] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// 「MP選出・ローテーション管理ツール」のHTMLにある ROTATION_STORE（JSON）を取り込む。
// 画面でファイルを読み、JSONの部分だけを渡す（写真入りのHTMLは大きいので送らない）
function importSpeakerRotation(json) {
  try {
    var t = typeof json === 'string' ? JSON.parse(json) : json;
    if (!t || !Array.isArray(t.order) || !t.order.length) return { ok: false, message: 'ローテーションの並び順が見つかりません。' };
    var date = parseDate_(String(t.nextMeetingDate || '').replace(/-/g, '/'));
    if (!date) return { ok: false, message: 'ツールの「次回の開催日」が読めません。' };
    var cur = rotLoad_();
    var st = { order: t.order.map(String), excluded: (t.excluded || []).map(String),
               anchor: { date: fmtDate_(date), pointer: parseInt(t.pointer, 10) || 0 },
               header: cur.header, notes: cur.notes };
    rotSave_(st);
    // ツールで休みにしていた日のうち、休会日に無いもの（お知らせだけ。休会日は ⚙️ 設定 で登録）
    var hol = [];
    try { hol = getHolidays(); } catch (e) {}
    var miss = (t.holidays || []).map(function (h) { return String(h).replace(/-/g, '/'); })
      .filter(function (h) { return parseDate_(h) && hol.indexOf(fmtDate_(parseDate_(h))) < 0; });
    var res = getSpeakerRotation();
    res.message = 'ツールのローテーションを取り込みました（' + st.order.length + '名・対象外 ' + st.excluded.length + '名・'
      + rotMd_(date) + ' は並びの' + (st.anchor.pointer + 1) + '番目から）。'
      + (miss.length ? '\nツールでお休みにしていた ' + miss.join('、') + ' が「休会日」にありません。'
         + 'お休みなら ⚙️ 設定 > 休会日 に登録してください（登録しないと、その週にも割り当てます）。' : '');
    return res;
  } catch (e) {
    return { ok: false, message: '取り込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}
function rotMd_(d) { return (d.getMonth() + 1) + '/' + d.getDate(); }

// 定例会スライド（前半）の画面から：その開催日から5回ぶんの予定と、表の見出し・注意書き
function getSpeakerRotationWeeks(dateStr) {
  try {
    var d = parseDate_(dateStr);
    if (!d) return { ok: false, message: '開催日が分かりません。' };
    routineResetCache_();
    var st = rotLoad_(), env = rotEnv_(d);
    env.skip = rotSkip_(st);
    return { ok: true, weeks: rotWeeks_(st, d, ROT_SLIDE_WEEKS_, env), header: st.header, notes: st.notes, provisional: !st.updated };
  } catch (e) {
    return { ok: false, message: 'スピーカーローテーションを読めませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// 役職ごとの入力（書記兼会計のメインプレゼンの初期値）：その開催日のローテーションの2名
// （ルーティンチェックシートは読まない。役職の画面を開く時間を延ばさないように）
function rotPairFor_(d) {
  var st = rotLoad_(), skip = rotSkip_(st), holidays = [];
  try { holidays = getHolidays(); } catch (e) {}
  if (holidays.indexOf(fmtDate_(d)) >= 0) return [];
  var p = rotPairsFrom_(st, skip, rotPointerAt_(st, skip, d, holidays), 1).pairs[0];
  return p || [];
}

// Facebookに添付する表の画像（画面で描いたPNG）を 03_生成物 に保存する。
// 画面からそのままダウンロードできない環境（ダイアログなど）向け。
function saveSpeakerRotationImage(base64, fileName) {
  try {
    var name = String(fileName || 'スピーカーローテーション.png').replace(/[\\\/:*?"<>|]/g, '_');
    if (!/\.png$/i.test(name)) name += '.png';
    var data = String(base64 || '').replace(/^data:image\/png;base64,/, '');
    var bytes = data ? Utilities.base64Decode(data) : [];
    // PNGの頭の目印（\x89PNG）があるものだけ保存する
    var sig = [0x89, 0x50, 0x4E, 0x47], isPng = bytes.length > 8;
    for (var i = 0; isPng && i < sig.length; i++) if ((bytes[i] & 0xff) !== sig[i]) isPng = false;
    if (!isPng) return { ok: false, message: '画像が空です。画面を開き直してください。' };
    var saved = saveOutputFile_(Utilities.newBlob(bytes, 'image/png', name), name);
    return { ok: true, url: saved.url, downloadUrl: saved.downloadUrl, fileName: name,
             message: '「03_生成物」に保存しました。' };
  } catch (e) {
    console.error('[ROT] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '画像を保存できませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// --- 定例会スライド（前半）の「スピーカーローテーション」のページ ---
// 書記兼会計が作っていた表の画像を、スライドの表に置き換える。
// ページは「スピーカーローテーション」の文字で探し、いちばん大きい画像の場所に表と注意書きを置く。
// 画像の無いページ（公式ファイルから作った雛形は、見本の表）なら、いちばん大きい表の場所に置く。
//   rot … { weeks: [...], header, notes }（getSpeakerRotationWeeks の結果）
function applySpeakerRotation_(parts, rot) {
  if (!rot || !rot.weeks || !rot.weeks.length) return null;
  var path = null, tag = 'p:pic', p, x;
  var isRot = function (xml) { return xml && slideText_(xml).replace(/[\s　]/g, '').indexOf('スピーカーローテーション') >= 0; };
  for (p in parts) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(p)) continue;
    x = xmlOf_(parts, p);
    if (isRot(x) && x.indexOf('<p:pic>') >= 0) { path = p; break; }
  }
  if (!path) {
    for (p in parts) {
      if (!/^ppt\/slides\/slide\d+\.xml$/.test(p)) continue;
      x = xmlOf_(parts, p);
      if (isRot(x) && /<a:tbl>/.test(x)) { path = p; tag = 'p:graphicFrame'; break; }
    }
  }
  if (!path) return { message: '「スピーカーローテーション」のページ（表の画像）が見つからないので、表は作りませんでした。' };
  var xml = xmlOf_(parts, path), pics = findTagRanges_(xml, tag), best = null, i;
  for (i = 0; i < pics.length; i++) {
    var seg = xml.substring(pics[i].start, pics[i].end);
    if (tag === 'p:graphicFrame' && !/<a:tbl>/.test(seg)) continue;
    var id = (seg.match(/<p:cNvPr[^>]*\sid="(\d+)"/) || [])[1];
    var g = seg.match(/<a:off x="(-?\d+)" y="(-?\d+)"\s*\/>\s*<a:ext cx="(\d+)" cy="(\d+)"/);
    if (!id || !g) continue;
    var box = { id: id, x: +g[1], y: +g[2], cx: +g[3], cy: +g[4] };
    if (!best || box.cx * box.cy > best.cx * best.cy) best = box;
  }
  if (!best) return { message: '「スピーカーローテーション」のページに表の画像が見つかりません。' };
  xml = removeShape_(xml, best.id);
  var maxId = 0, m, re = /<p:cNvPr[^>]*\sid="(\d+)"/g;
  while ((m = re.exec(xml)) !== null) maxId = Math.max(maxId, parseInt(m[1], 10));
  // 表は画像の場所の上から、注意書きはその下に置く
  var notes = (rot.notes || []).filter(function (s) { return String(s).trim(); });
  var notesH = notes.length ? Math.round(notes.length * 12 * 1.35 * 12700 + 91440) : 0;
  var tableH = best.cy - notesH - (notes.length ? 60000 : 0);
  var built = rotTableXml_(maxId + 1, best.x, best.y, best.cx, tableH, rot);
  var add = built.xml;
  if (notes.length) add += rotNotesXml_(maxId + 2, best.x, best.y + built.height + 60000, best.cx, notesH, notes);
  xml = xml.replace('</p:spTree>', add + '</p:spTree>');
  putXml_(parts, path, xml);
  var names = [];
  rot.weeks.forEach(function (w) { w.people.forEach(function (pp) { if (pp.name) names.push(pp.name); }); });
  return { message: 'スピーカーローテーションの表を作りました（' + rot.weeks.length + '回ぶん: '
             + rot.weeks[0].label + '〜' + rot.weeks[rot.weeks.length - 1].label + '）。', path: path, names: names };
}

// 表：見出しの行＋1回につき3行（第N回／カテゴリー、日付／氏名、つながりたい人）
var ROT_FONT_ = 'Meiryo UI';
function rotTableXml_(id, x, y, cx, maxH, rot) {
  var weeks = rot.weeks, n = weeks.length;
  var colW = [Math.round(cx * 0.17)];
  colW.push(Math.round((cx - colW[0]) / 2));
  colW.push(cx - colW[0] - colW[1]);
  // 行の高さ：見出し・カテゴリー・氏名・つながりたい人（全体の高さに収まるよう比で割る）
  var unit = Math.floor(maxH / (1.1 + n * (1.0 + 1.25 + 0.8)));
  var hHead = Math.round(unit * 1.1), hCat = unit, hName = Math.round(unit * 1.25), hCol = Math.round(unit * 0.8);
  var pt = function (emu) { return emu / 12700; };
  // 行の高さに合う文字の大きさ（上限つき）
  var szHead = Math.min(16, Math.floor(pt(hHead) * 0.55)), szCat = Math.min(15, Math.floor(pt(hCat) * 0.6));
  var szName = Math.min(22, Math.floor(pt(hName) * 0.62)), szCol = Math.min(11, Math.floor(pt(hCol) * 0.62));
  var szNo = Math.min(12, szCat), szDate = Math.min(18, szName);
  var rows = [];
  rows.push(rotTr_(hHead, [
    rotTc_('日　程', szHead, { b: true, fill: 'FF0000', w: colW[0] }),
    rotTc_(rot.header || '', szHead, { b: true, fill: 'FF0000', w: colW[1] + colW[2], span: 2 }),
    '<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ja-JP"/></a:p></a:txBody><a:tcPr/></a:tc>'
  ]));
  for (var i = 0; i < n; i++) {
    var w = weeks[i], a = w.people[0] || {}, b = w.people[1] || {};
    rows.push(rotTr_(hCat, [
      rotTc_(w.no ? '第' + w.no + '回' : '', szNo, { b: true, w: colW[0] }),
      rotTc_(a.title || '', szCat, { b: true, w: colW[1] }),
      rotTc_(b.title || '', szCat, { b: true, w: colW[2] })
    ]));
    rows.push(rotTr_(hName, [
      rotTc_(w.label || '', szDate, { b: true, w: colW[0] }),
      rotTc_(a.name || '', szName, { b: true, color: 'C00000', w: colW[1] }),
      rotTc_(b.name || '', szName, { b: true, color: 'C00000', w: colW[2] })
    ]));
    rows.push(rotTr_(hCol, [
      rotTc_('', szCol, { fill: 'FCD5B4', w: colW[0] }),
      rotTc_(a.collab || '', szCol, { b: true, fill: 'FCD5B4', w: colW[1] }),
      rotTc_(b.collab || '', szCol, { b: true, fill: 'FCD5B4', w: colW[2] })
    ]));
  }
  var height = hHead + n * (hCat + hName + hCol);
  var xml = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="' + id + '" name="スピーカーローテーションの表"/>'
    + '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
    + '<p:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + height + '"/></p:xfrm>'
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>'
    + '<a:tblPr firstRow="1"/><a:tblGrid>' + colW.map(function (v) { return '<a:gridCol w="' + v + '"/>'; }).join('') + '</a:tblGrid>'
    + rows.join('') + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
  return { xml: xml, height: height };
}

function rotTr_(h, cells) { return '<a:tr h="' + h + '">' + cells.join('') + '</a:tr>'; }

// セル。長い文字は1行に収まるまで小さくする（全角1em・半角0.55em・空白0.3emで見積もり、5%のゆとり）
function rotTc_(text, size, o) {
  var s = String(text == null ? '' : text);
  var usable = ((o.w || 1000000) - 91440) / 12700 * 0.95, em = 0;
  for (var i = 0; i < s.length; i++) { var c = s.charCodeAt(i); em += c === 32 ? 0.3 : (c < 128 ? 0.55 : 1); }
  var sz = size;
  if (em && em * sz > usable) sz = Math.max(7, Math.floor(usable / em * 2) / 2);
  var line = function (clr) {
    return '<a:solidFill><a:srgbClr val="' + clr + '"/></a:solidFill>';
  };
  var ln = '<a:lnL w="9525">' + line('7F7F7F') + '</a:lnL><a:lnR w="9525">' + line('7F7F7F') + '</a:lnR>'
         + '<a:lnT w="9525">' + line('7F7F7F') + '</a:lnT><a:lnB w="9525">' + line('7F7F7F') + '</a:lnB>';
  var rPr = '<a:rPr lang="ja-JP" altLang="en-US" sz="' + Math.round(sz * 100) + '"' + (o.b ? ' b="1"' : '') + ' dirty="0">'
          + line(o.color || '000000') + '<a:latin typeface="' + ROT_FONT_ + '"/><a:ea typeface="' + ROT_FONT_ + '"/></a:rPr>';
  var body = '<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/>'
    + (s ? '<a:r>' + rPr + '<a:t>' + escapeXml_(s) + '</a:t></a:r>' : '')
    + '<a:endParaRPr lang="ja-JP" altLang="en-US" sz="' + Math.round(sz * 100) + '" dirty="0"/></a:p></a:txBody>';
  return '<a:tc' + (o.span ? ' gridSpan="' + o.span + '"' : '') + '>' + body
    + '<a:tcPr marL="45720" marR="45720" marT="0" marB="0" anchor="ctr">' + ln
    + (o.fill ? '<a:solidFill><a:srgbClr val="' + o.fill + '"/></a:solidFill>' : '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>')
    + '</a:tcPr></a:tc>';
}

// 表の下の注意書き（※ …）
function rotNotesXml_(id, x, y, cx, cy, notes) {
  var ps = notes.map(function (s) {
    return '<a:p><a:r><a:rPr lang="ja-JP" altLang="en-US" sz="1200" dirty="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
      + '<a:latin typeface="' + ROT_FONT_ + '"/><a:ea typeface="' + ROT_FONT_ + '"/></a:rPr><a:t>' + escapeXml_('※　' + s) + '</a:t></a:r></a:p>';
  }).join('');
  return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="スピーカーローテーションの注意書き"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
    + '<p:spPr><a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>'
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
    + '<p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr>'
    + '<a:lstStyle/>' + ps + '</p:txBody></p:sp>';
}
