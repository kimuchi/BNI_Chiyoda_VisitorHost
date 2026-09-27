// === ルーティンチェックシートから、その日の決めごとを読む ===
//
// チャプターでは開催日ごとの準備を「【NN期】ルーティンチェックシート」で管理している。
// 1行目に開催日、2行目に定例会回数が並び、その下に項目が縦に並ぶ表。
// ここに書いてある「コアバリュー」「スタートアッププレゼンの方」を、
// 各画面の初期値として読み込む（同じことを二度入力しなくて済むように）。
//
//   1行目  … 定例会開催日（1列おきに入っている）
//   2行目  … 定例会回数
//   …行    … B〜E列のどこかに項目名。その開催日の列に値が入る
//
// 期（17期・18期…）が変わるとシートも変わるが、開催日そのもので探すので
// 期の計算は要らない。項目の行番号も期によってずれるため、名前で探す。

var ROUTINE_SHEET_RE_ = /^【(\d+)期】ルーティンチェックシート$/;
var ROUTINE_SCAN_ROWS_ = 150;     // 項目名を探す範囲（これより下は見ない。「役職ごとの入力」で足す行も含む）
var ROUTINE_LABEL_COLS_ = 5;      // 項目名が入りうる列（A〜E）

// BNIのコアバリュー。シートの書き方が揺れる（Givers Gain® / ギバーズゲイン /
// Lifelong Learnign など）ので、英語のキーワード → 日本語のキーワードの順に当てる。
var CORE_VALUES_ = [
  { label: 'Givers Gain',             en: 'givers',     ja: 'ギバーズ' },
  { label: 'Building Relationships',  en: 'building',   ja: '関係構築' },
  { label: 'Lifelong Learning',       en: 'lifelong',   ja: '生涯学習' },
  { label: 'Traditions + Innovation', en: 'tradition',  ja: '伝統' },
  { label: 'Positive Attitude',       en: 'positive',   ja: '前向き' },
  { label: 'Accountability',          en: 'accountab',  ja: '責任' },
  { label: 'Recognition',             en: 'recogni',    ja: '承認' }
];

function coreValueOf_(text) {
  var s = String(text == null ? '' : text), i;
  var en = s.toLowerCase().replace(/[^a-z]/g, '');
  for (i = 0; i < CORE_VALUES_.length; i++) if (en.indexOf(CORE_VALUES_[i].en) >= 0) return CORE_VALUES_[i];
  var ja = s.replace(/[\s　]/g, '');
  for (i = 0; i < CORE_VALUES_.length; i++) if (ja.indexOf(CORE_VALUES_[i].ja) >= 0) return CORE_VALUES_[i];
  return null;
}

// 1行目の開催日。ラベルや数字を日付と誤解しないよう、Dateか「2026/…」形式だけを見る
function routineDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return parseDate_(v);
  var s = String(v == null ? '' : v).trim();
  if (!/^\d{4}[\/\-年]/.test(s)) return null;
  return parseDate_(s);
}

// 開催日 → { シート, 列 } の索引。
// 1回の実行中に何度も呼ばれる（開催日の候補ぶん）ので、1度だけ作って使い回す。
//
// 昔の期のシートは、ルールを知るための控えで、ふだんの準備には使わないので読まない
// （期ごとに1枚ずつ増えていき、全部読むと時間がかかる）。
// 期の新しいシートから読んでいき、ROUTINE_RECENT_DAYS_ 日より前の回しか載っていないシートまで来たら、
// そのシート（前回の開催日を探すのに使うことがある）で止める。
// それより前の日を探すときだけ（fromKey に昔の日を渡したとき）、昔の期のシートも読む。
var ROUTINE_RECENT_DAYS_ = 120;
var ROUTINE_INDEX_ = null;
var ROUTINE_INDEX_FROM_ = '';     // 索引に入っている範囲：この日からあとの回（'' は全部のシート）
function routineRecentKey_() {
  var d = new Date(); d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() - ROUTINE_RECENT_DAYS_);
  return fmtDate_(d);
}
function routineIndex_(fromKey) {
  var from = routineRecentKey_();
  if (fromKey && fromKey < from) from = fromKey;
  if (ROUTINE_INDEX_ && (!ROUTINE_INDEX_FROM_ || ROUTINE_INDEX_FROM_ <= from)) return ROUTINE_INDEX_;
  var sheets = getSS_().getSheets(), list = [], idx = {}, all = true, i;
  for (i = 0; i < sheets.length; i++) {
    var m = sheets[i].getName().match(ROUTINE_SHEET_RE_);
    if (m) list.push({ sheet: sheets[i], name: sheets[i].getName(), term: parseInt(m[1], 10) });
  }
  list.sort(function (a, b) { return b.term - a.term; });          // 期の新しいシートから
  for (i = 0; i < list.length; i++) {
    var last = list[i].sheet.getLastColumn();
    if (last < 2) continue;
    var row1 = list[i].sheet.getRange(1, 1, 1, last).getValues()[0], newest = '';
    for (var c = 0; c < row1.length; c++) {
      var d = routineDate_(row1[c]);
      if (!d) continue;
      var key = fmtDate_(d);
      if (key > newest) newest = key;
      // 同じ日が複数の期にあれば、期の新しい方（先に読んだ方）を採る
      if (!idx[key]) idx[key] = { sheet: list[i].sheet, name: list[i].name, term: list[i].term, col: c + 1 };
    }
    if (newest && newest < from && i < list.length - 1) { all = false; break; }   // ここより昔の期は読まない
  }
  ROUTINE_INDEX_ = idx;
  ROUTINE_INDEX_FROM_ = all ? '' : from;
  return idx;
}

// 開催日から、その日が載っているシートと列を探す（昔の日なら、昔の期のシートも読む）
function findRoutineColumn_(target) {
  var key = fmtDate_(target);
  return routineIndexFor_(key)[key] || null;
}
function routineIndexFor_(key) {
  var idx = routineIndex_();
  if (!idx[key] && ROUTINE_INDEX_FROM_ && key < ROUTINE_INDEX_FROM_) idx = routineIndex_(key);
  return idx;
}

// 項目名（B〜E列のどこか）で行を探す。0から数えた行番号（無ければ -1）。
// まず完全一致で探し、無ければ前方一致で探す。
// 「メインプレゼン」で前方一致だけにすると「メインプレゼン用スライド画像」の行を
// 先に拾ってしまうため、この順番が要る。
function routineFindRow_(grid, labels) {
  var scan = function (exact) {
    for (var r = 0; r < grid.length; r++) {
      for (var c = 1; c < Math.min(ROUTINE_LABEL_COLS_, grid[r].length); c++) {
        var s = String(grid[r][c] == null ? '' : grid[r][c]).replace(/[\s　]/g, '');
        if (!s) continue;
        for (var k = 0; k < labels.length; k++) {
          if (exact ? (s === labels[k]) : (s.indexOf(labels[k]) === 0)) return r;
        }
      }
    }
    return -1;
  };
  var r = scan(true);
  return r >= 0 ? r : scan(false);
}

// ある項目を、開催日ごとに読む → { 'yyyy/MM/dd': 値 }（空欄の日は入れない）。
// 「ウィークリープレゼン」のように、その日に書かれていなければ
// 前回までの記載から数えたい項目に使う。
//   fromKey … 'yyyy/MM/dd'。この日より前の回しか載っていないシート（昔の期）は読まない。
//             省略すると、最近（ROUTINE_RECENT_DAYS_ 日）の回が載っているシートだけ
// シートの中身は1回の実行の中で使い回す（項目を変えて何度読んでも、シートを読むのは1回）。
var ROUTINE_WEEKLY_LABELS_ = ['ウィークリープレゼン'];
// スタートアッププレゼン（新しく入った方の長いプレゼン）の行。
// チャプターによって長さが違い、「2分30秒プレゼン」と書いてあるシートもあるので、どちらでも読む
var ROUTINE_STARTUP_LABELS_ = ['スタートアッププレゼン', 'スタートアップ', '2分30秒プレゼン', '2分30秒'];
var ROUTINE_ROWS_CACHE_ = {};
var ROUTINE_GRID_CACHE_ = {};
function routineRowValues_(labels, fromKey) {
  var from = fromKey || routineRecentKey_();
  var ck = labels.join('|') + '@' + from;
  if (ROUTINE_ROWS_CACHE_[ck]) return ROUTINE_ROWS_CACHE_[ck];
  var idx = routineIndex_(from), bySheet = {}, out = {}, key, name;
  for (key in idx) {
    var e = idx[key];
    if (!bySheet[e.name]) bySheet[e.name] = { sheet: e.sheet, cols: [], last: '' };
    bySheet[e.name].cols.push({ key: key, col: e.col });
    if (key > bySheet[e.name].last) bySheet[e.name].last = key;
  }
  for (name in bySheet) {
    if (bySheet[name].last < from) continue;
    var grid = routineSheetGrid_(name, bySheet[name].sheet);
    if (!grid) continue;
    var r = routineFindRow_(grid, labels);
    if (r < 0) continue;
    for (var i = 0; i < bySheet[name].cols.length; i++) {
      var c = bySheet[name].cols[i], v = grid[r][c.col - 1];
      v = String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').trim();
      if (v) out[c.key] = v;
    }
  }
  ROUTINE_ROWS_CACHE_[ck] = out;
  return out;
}

// ルーティンチェックシート1枚の中身（項目を探す範囲まで）。1回の実行の中で使い回す
function routineSheetGrid_(name, sh) {
  if (ROUTINE_GRID_CACHE_[name]) return ROUTINE_GRID_CACHE_[name];
  var rows = Math.min(ROUTINE_SCAN_ROWS_, sh.getLastRow()), last = sh.getLastColumn();
  if (rows < 1 || last < 1) return null;
  ROUTINE_GRID_CACHE_[name] = sh.getRange(1, 1, rows, last).getValues();
  return ROUTINE_GRID_CACHE_[name];
}

// 読み込んだルーティンチェックシートの控えを捨てる（書き込んだあとや、画面から呼ばれた最初に）
function routineResetCache_() {
  ROUTINE_INDEX_ = null;
  ROUTINE_INDEX_FROM_ = '';
  ROUTINE_ROWS_CACHE_ = {};
  ROUTINE_GRID_CACHE_ = {};
  ROUTINE_ROSTER_ = null;
}

// 「なし」「無し」「休会」などは、指定されていないものとして扱う。
// 「なし（4/2火曜10時時点）」のように但し書きが付くことがあるので、先頭だけ見る。
function routineIsBlank_(s) {
  return !s || /^(なし|無し|ナシ|ない|―|－|-|—|休会|済|未定|\?|？)/.test(s);
}

// 「谷口さん」のような書き方を、メンバー名簿のフルネームに合わせる。
// 「３５番船越さん」「内装業　岩渕さん」のように、番号や業種が前に付くことがあるため、
// そのままで当たらなければ、前置きを落としながら何通りか試す。
function routineMemberName_(raw) {
  var s0 = String(raw == null ? '' : raw).trim();
  if (routineIsBlank_(s0.replace(/[\s　]/g, ''))) return { raw: '', name: '', matched: false };

  var tries = [], push = function (x) {
    x = String(x || '').replace(/[\s　]/g, '');
    if (x && tries.indexOf(x) < 0) tries.push(x);
  };
  push(s0);
  push(s0.replace(/(さん|様|さま|くん|君|ちゃん|氏)?へ$/, '$1'));   // 「山内さんへ」
  push(s0.replace(/(さん|様|さま|くん|君|ちゃん|氏)?へ?$/, ''));    // 「長尾くん」
  push(s0.replace(/^[0-9０-９]+\s*番?/, ''));            // 「３５番船越さん」
  var parts = s0.split(/[\s　]+/);
  if (parts.length > 1) push(parts[parts.length - 1]);   // 「内装業　岩渕さん」
  push(s0.replace(/[（(].*$/, ''));                       // 「谷口さん（代理）」

  var members = getMembersList(), i, k;
  for (k = 0; k < tries.length; k++) {
    var full = matchInviterToMember(tries[k], members);
    for (i = 0; i < members.length; i++) {
      if (normName_(members[i].name) === normName_(full)) {
        return { raw: s0, name: members[i].name, matched: true };
      }
    }
  }
  // 字の違い（「川辺さん」と名簿の「川邉 真由子」など）でも、名字が1人に決まるなら合わせる
  for (k = 0; k < tries.length; k++) {
    var f = routineFoldName_(tries[k]), hit = [];
    if (f.length < 2) continue;
    for (i = 0; i < members.length; i++) {
      if (routineFoldName_(members[i].name).indexOf(f) === 0) hit.push(members[i].name);
    }
    if (hit.length === 1) return { raw: s0, name: hit[0], matched: true };
  }
  return { raw: s0, name: '', matched: false };
}

// 名字でよく使い分けられる字をそろえる（照らし合わせるときだけ使う）
var ROUTINE_NAME_VARIANTS_ = {
  '邉': '辺', '邊': '辺', '齋': '斉', '齊': '斉', '斎': '斉', '髙': '高', '﨑': '崎', '嵜': '崎',
  '澤': '沢', '濱': '浜', '濵': '浜', '廣': '広', '嶋': '島', '嶌': '島', '國': '国', '眞': '真',
  '冨': '富', '瀨': '瀬', '德': '徳', '藏': '蔵', '櫻': '桜', '龍': '竜', '萬': '万', '槇': '槙',
  '桒': '桑', '淵': '渕', '條': '条', '峯': '峰', '藪': '薮', '籔': '薮', '栁': '柳', '惠': '恵',
  '實': '実', '壽': '寿', '豐': '豊', '禮': '礼'
};
function routineFoldName_(s) {
  var t = String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]/g, '').replace(/さん$/, '');
  var out = '';
  for (var i = 0; i < t.length; i++) out += ROUTINE_NAME_VARIANTS_[t.charAt(i)] || t.charAt(i);
  return out;
}

// --- 欄の書き方を、スライドに載せられる形に直す ----------------------------

// 「なし」などをそのまま残す（スライドにも「なし」と出るため）。空欄だけ空にする。
function routineText_(v) {
  return String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').trim();
}

// 「43.0 一般規定」の欄。「2番」「６番」→ 2 / 6
function routinePolicyNo_(v) {
  var s = String(v == null ? '' : v).normalize('NFKC');
  var m = s.match(/(\d+)\s*番/) || s.match(/^\s*(\d+)\s*$/);
  return m ? parseInt(m[1], 10) : 0;
}

// 「22 メインプレゼン」の欄。「①山内さん　②金井さん」「①北川‗②若松」など。
// 番号の丸数字で切り、無ければ読点で切る。それぞれをメンバー名簿の氏名に合わせる。
function routineMainPresenters_(v) {
  var s = String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').trim();
  if (routineIsBlank_(s.replace(/[\s　]/g, ''))) return [];
  var parts;
  if (/[①②③④⑤]/.test(s)) {
    parts = s.split(/[①②③④⑤]/).slice(1);
  } else {
    parts = s.split(/[、,／\/･・]/);
  }
  var out = [];
  for (var i = 0; i < parts.length && out.length < 2; i++) {
    var raw = String(parts[i]).replace(/[‗_＿\-－―]/g, ' ').trim();
    if (!raw) continue;
    var hit = routineMemberName_(raw);
    out.push({ raw: raw, name: hit.name, matched: hit.matched });
  }
  return out;
}

// 「募集カテゴリー」の欄。「赤文字／黒文字」の見出しや但し書きを落として、項目だけ並べる。
//   赤文字
//   ・結婚相談所　・電気工事業　・SNS運用
//   黒文字
//   ・ハウスメーカー　・工務店
// → ['結婚相談所','電気工事業','SNS運用','ハウスメーカー','工務店']
function routineList_(v) {
  var s = String(v == null ? '' : v);
  if (routineIsBlank_(s.replace(/[\s　]/g, ''))) return [];
  var lines = s.split(/[\r\n]+/), out = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    if (/^[（(]/.test(line)) continue;                       // 「（…から更新します）」などの但し書き
    if (/^(赤文字|黒文字|赤|黒)[：:　 ]*$/.test(line)) continue;  // 色の見出し
    var items = line.split(/[・･、,／\/]+/);
    for (var k = 0; k < items.length; k++) {
      var t = items[k].replace(/[\s　]+/g, ' ').trim();
      if (!t) continue;
      if (/^(赤文字|黒文字)$/.test(t)) continue;
      out.push(t);
    }
  }
  return out;
}

// 「25 推薦の言葉」の欄。組ごとに「推薦する人 → 推薦される人」のページを作る。
//   定例会中
//    ⓵藤本さん⇒金井さん、
//    ②山岸さん⇒山内さん
//   アフター
//    ①船越さん→金井さん
// のように、見出しで「いつ発表するか」を分けて書かれる。見出しの下の組にその時間を付ける。
// 「加納さん→近江さん（アフター）」のように、組のうしろの括弧に書かれていることもある（その組だけに付ける）。
//   during … 定例会中（見出しが無いときもこれ）。推薦のことばのページの場所に並べる
//   after  … アフター・定例会後。抽選コーナーのあとに並べる
//   later  … 翌週以降（予定のメモ）。スライドには入れない
// 1行に「①A→B ②C→D」と続けて書かれていても、丸数字で区切って組に分ける。
var ROUTINE_RECO_MARK_RE_ = /[①②③④⑤⑥⑦⑧⑨⑩⓵⓶⓷⓸⓹⓺⓻⓼⓽]/;
var ROUTINE_RECO_SAMPLE_RE_ = /^[\s　]*[○〇◯×✕＊*…]+(さん|様)?[\s　、,]*$/;
function routineRecoWhen_(text, when) {
  if (/翌週|来週|次週/.test(text)) return 'later';
  if (/アフター|定例会後|終了後/.test(text)) return 'after';
  if (/定例会/.test(text)) return 'during';
  return when;
}
function routineRecommendations_(v) {
  var s = String(v == null ? '' : v);
  if (routineIsBlank_(s.replace(/[\s　]/g, ''))) return [];
  var ARROW = /[→⇒➡]/, lines = s.split(/[\r\n]+/), out = [], when = 'during';
  for (var i = 0; i < lines.length; i++) {
    // 組のうしろの括弧に時間が書いてあれば（「加納さん→近江さん（アフター）」）、その行の組だけに使う
    var notes = (lines[i].match(/[（(][^）)]*[）)]/g) || []).join(' ');
    var lineWhen = notes ? routineRecoWhen_(notes, '') : '';
    // 括弧の但し書き（「（先週繰り越し分）」など）と、※以降のメモは落とす
    var line = lines[i].replace(/[（(][^）)]*[）)]/g, '').replace(/[※＊].*$/, '').trim();
    var arrowAt = line.search(ARROW), markAt = line.search(ROUTINE_RECO_MARK_RE_);
    // 見出し（「定例会中」「アフター」「定例会後」「翌週以降」）で、いつ発表する組かを切り替える。
    // 見出しと組が同じ行に書かれていたら、組より前の部分だけで判断する。
    var headEnd = arrowAt < 0 ? line.length : (markAt >= 0 && markAt < arrowAt ? markAt : arrowAt);
    var head = line.substring(0, headEnd);
    when = routineRecoWhen_(head, when);
    if (arrowAt < 0) {
      if (lineWhen && !head.replace(/[\s　【】\[\]]/g, '')) when = lineWhen;       // 「（アフター）」だけの見出し
      continue;
    }
    var body;
    if (markAt >= 0 && markAt < arrowAt) body = line.substring(markAt);          // 丸数字から先が組
    else {
      // 見出しの語（「定例会後：」「【アフター】」「アフター分」など）だけ落とす。うしろに続くお名前は残す
      var kw = head.match(/^.*(定例会中|定例会後|終了後|アフター|定例会)(?:の?分|の部|にて|で)?[\s　:：、。.．\-－―)）】」』］\]]*/);
      body = kw ? line.substring(kw[0].length) : line;
    }
    var segs = body.split(ROUTINE_RECO_MARK_RE_);
    for (var k = 0; k < segs.length; k++) {
      // 行頭の記号（「・船越さん→…」の・など）と、前後の区切りを落とす
      var seg = segs[k].replace(/^[\s　.．、,・･•●◆◇■□▪*＊\-－]+/, '').replace(/[\s　、,]+$/, '');
      if (!ARROW.test(seg)) continue;
      // 1つの区切りに2組以上続けて書かれていたら（「金井さん→西澤さん 藤本さん→長尾さん」）、組ごとに分ける
      var found = [], mm;
      if ((seg.match(/[→⇒➡]/g) || []).length > 1) {
        var re = /([^\s　、,→⇒➡]+)[\s　]*[→⇒➡][\s　]*([^\s　、,→⇒➡]+)/g;
        while ((mm = re.exec(seg)) !== null) found.push({ a: mm[1], b: mm[2], raw: mm[0] });
      } else {
        var parts = seg.split(ARROW);
        found.push({ a: parts[0], b: parts[1], raw: seg });
      }
      for (var f = 0; f < found.length; f++) {
        // 書き方の見本（「○○さん ⇒○○さん」）は組にしない
        if (ROUTINE_RECO_SAMPLE_RE_.test(found[f].a) && ROUTINE_RECO_SAMPLE_RE_.test(found[f].b)) continue;
        var a = routineMemberName_(String(found[f].a).replace(/^[・･•●◆◇■□▪*＊\-－]+/, '').replace(/[、,]\s*$/, ''));
        var b = routineMemberName_(String(found[f].b).replace(/^[・･•●◆◇■□▪*＊\-－]+/, '').replace(/[、,]\s*$/, ''));
        if (!a.raw && !b.raw) continue;
        out.push({ giver: a, receiver: b, raw: found[f].raw, when: lineWhen || when });
      }
    }
  }
  return out;
}

// --- 前半スライドの「新メンバー」「更新メンバー」「バイスプレジデントによる報告」「ネットワーキングリーダー」 ---
// どれもバイスプレジデントが2日前までに書く欄。書き方は人によって揺れるので、形を見て拾う。
var ROUTINE_NEW_LABELS_ = ['新入会', '新メンバー'];
var ROUTINE_RENEW_LABELS_ = ['更新式', '更新メンバー'];
var ROUTINE_VP_LABELS_ = ['バイスプレジデントによる報告'];
var ROUTINE_NL_LABELS_ = ['ネットワーキングリーダー'];

// 括弧の外と中に分ける（入れ子の括弧・全角半角の混ざった括弧も1つの組にする）
//   「徳永京平さん（飲食店（ホルモン焼肉））メンバーリスト53番」
//   → [{外:'徳永京平さん'}, {中:'飲食店（ホルモン焼肉）'}, {外:'メンバーリスト53番'}]
function routineParenSegs_(s) {
  var out = [], depth = 0, buf = '';
  for (var i = 0; i < s.length; i++) {
    var ch = s.charAt(i);
    if (ch === '（' || ch === '(') {
      if (depth === 0) { if (buf) out.push({ text: buf, paren: false }); buf = ''; } else buf += ch;
      depth++;
    } else if ((ch === '）' || ch === ')') && depth > 0) {
      depth--;
      if (depth === 0) { out.push({ text: buf, paren: true }); buf = ''; } else buf += ch;
    } else buf += ch;
  }
  if (buf) out.push({ text: buf, paren: depth > 0 });
  return out;
}

// 名簿の氏名に合わせる。routineMemberName_ より慎重に、名字・氏名がそのまま一致する方を先に採り、
// 前にカテゴリーなどが付いた書き方（「店舗オフィス岩渕」「占いを使ったインサイトカウンセラー神楽」）は
// うしろが名字・氏名に一致する方を採る（2文字以上）。
// 同じ名字の方が2人以上いるときは、書き添えてあるカテゴリー（hint）が名簿のカテゴリーに合う方に決める。決まらなければ空
function routineRosterName_(text, roster, hint) {
  var t = routineFoldName_(String(text).replace(/【[^】]*】/g, ''));
  if (!t) return '';
  var keys = roster.map(function (m) {
    return { name: m.name, title: routineFoldName_(m.title || ''),
             full: routineFoldName_(m.name), sur: routineFoldName_(String(m.name).trim().split(/[\s　]+/)[0]) };
  });
  var h = routineFoldName_(String(hint || '').replace(/[()（）【】「」]/g, ''));
  var one = function (list) {
    if (list.length === 1) return list[0].name;
    if (!h || list.length < 2) return '';
    var byCat = list.filter(function (k) {
      var ti = k.title.replace(/[()（）【】「」]/g, '');
      return ti && (ti.indexOf(h) >= 0 || h.indexOf(ti) >= 0);
    });
    return byCat.length === 1 ? byCat[0].name : '';
  };
  var exact = keys.filter(function (k) { return k.full === t; });
  if (exact.length) return one(exact);
  var sur = keys.filter(function (k) { return k.sur === t; });
  if (sur.length) return one(sur);
  var hit = routineMemberName_(text);
  if (hit.matched) return hit.name;
  var best = [], bestLen = 0;
  keys.forEach(function (k) {
    [k.full, k.sur].forEach(function (w) {
      if (w.length < 2 || t.length <= w.length || t.slice(-w.length) !== w) return;
      if (w.length > bestLen) { best = [k]; bestLen = w.length; }
      else if (w.length === bestLen && best.indexOf(k) < 0) best.push(k);
    });
  });
  if (best.length > 1 && !h) h = routineFoldName_(t.slice(0, -bestLen));      // 前に付いているのがカテゴリー
  return one(best);
}
// 名簿（氏名とカテゴリー）。1回の実行の中で使い回す
var ROUTINE_ROSTER_ = null;
function routineRoster_() {
  if (ROUTINE_ROSTER_) return ROUTINE_ROSTER_;
  var list = [];
  try { list = (getMemberMaster({ membersOnly: true }).members || []).map(function (m) { return { name: m.name, title: m.title || '' }; }); } catch (e) {}
  if (!list.length) {
    try { list = getMembersList().map(function (m) { return { name: m.name, title: '' }; }); } catch (e) {}
  }
  ROUTINE_ROSTER_ = list;
  return list;
}

// 括弧の中が「カテゴリー」らしいか（番号・年数・日付・よみがな・但し書きは違う）
function routineCategoryIn_(t) {
  var s = String(t || '').trim(), m = s.match(/カテゴリー\s*[：:]?\s*([^\s　、,，]+)/);
  if (m) return m[1];
  if (!s || /[0-9０-９]|番|年|月|日|時点|代理|予定|メンバーリスト|さん/.test(s)) return '';
  if (/^[ぁ-ゖー・\s　]+$/.test(s)) return '';                 // よみがな（ひらがなだけ）
  return s;
}
// 名前らしい文字か（「対面BOD」「該当者なし」のような書き込みを外す）
function routineNameLike_(t) {
  var s = String(t || '').replace(/[\s　]/g, '');
  return s.length >= 1 && s.length <= 10 && !/[A-Za-zＡ-Ｚａ-ｚ0-9０-９]/.test(s)
      && !/該当|なし|無し|受賞|以上|次は|最後|お二人|おふたり|ふたり|招待|部門|ビジター|連続|初の|抑え|予定|休会|カテゴリー/.test(s);
}

// 「15 新入会」「16 更新式(更新メンバー)」の欄 → [{ raw, name, matched, years, category }]
//   本間さん、平松さん、谷口さん
//   徳永京平さん（飲食店（ホルモン焼肉））メンバーリスト53番
//   （業務用冷凍冷蔵設備) 福元 良平さん ⏎（LPに特化してWebデザイン）村井 絢香さん
//   鈴村さん（1年更新）、加納さん（1年更新）          … 更新は「1年」「2年」も読む（書いていなければ 0）
//   内装業　店舗オフィス岩渕さん（1年）               … 前に業種・カテゴリーが付く
//   岩渕裕太（いしぶちゆうすけ）さん                   … よみがなの括弧のあとに「さん」
//   小西さん（伊豆澤さんは10/7）                      … 括弧の中のお名前は数えない
//   青木周一さん（…） ⏎ 遠藤さんのあと、31番に入ります … 「○○さんのあと」のような文の中のお名前も数えない
//   深井 宗二郎（ふかた そういちろう）（法人コスト削減） … どこにも「さん」が無ければ、括弧の外をお名前とみなす
// category は、名簿に無い方（まだ名簿に入っていない新メンバー）のときに使う
function routineMemberList_(v) {
  var s = String(v == null ? '' : v);
  if (routineIsBlank_(s.replace(/[\s　]/g, ''))) return [];
  var roster = routineRoster_(), out = [], seen = {};
  var add = function (rawName, shown, years, category) {
    var nm = String(rawName).replace(/【[^】]*】/g, '').replace(/^[\s　、,，・･と]+|[\s　、,，・･]+$/g, '');
    if (!nm || routineIsBlank_(nm.replace(/[\s　]/g, ''))) return;
    var full = routineRosterName_(nm, roster, category), key = full || routineFoldName_(nm);
    if (!full && !routineNameLike_(nm.replace(/^.*[\s　]/, ''))) return;
    if (seen[key]) return;
    seen[key] = true;
    out.push({ raw: shown, name: full, matched: !!full, years: years || 0, category: full ? '' : (category || '') });
  };
  var lines = s.split(/[\r\n]+/).map(function (l) { return l.replace(/[※＊].*$/, '').trim(); });
  var honor = /(さん|様|さま|氏)/;
  if (lines.some(function (l) { return honor.test(routineParenSegs_(l).filter(function (x) { return !x.paren; }).map(function (x) { return x.text; }).join('')); })) {
    lines.forEach(function (line) {
      if (!line || routineIsBlank_(line.replace(/[\s　]/g, ''))) return;
      var segs = routineParenSegs_(line);
      for (var i = 0; i < segs.length; i++) {
        if (segs[i].paren) continue;
        var text = segs[i].text, re = /(さん|様|さま|氏)/g, m, cut = 0;
        while ((m = re.exec(text)) !== null) {
          var after = text.substring(m.index + m[0].length);
          var chunk = text.substring(cut, m.index), lead = cut === 0;
          cut = m.index + m[0].length;
          if (/^[のはがをにへもで]/.test(after)) continue;                 // 「遠藤さんのあと」
          chunk = chunk.replace(/^.*[、,，。；;：:／\/]/, '');                 // 区切りより前（「法人コスト削減：」など）
          var reading = '';
          if (!chunk.trim() && lead && i >= 2 && segs[i - 1].paren && !segs[i - 2].paren) {
            chunk = segs[i - 2].text.replace(/^.*[、,，。；;：:／\/]/, '');   // 「岩渕裕太（いしぶちゆうすけ）さん」
            reading = segs[i - 1].text;
          }
          var next = !after.trim() ? segs[i + 1] : null;
          var ym = (next && next.paren ? next.text : after).match(/^[\s　]*([0-9０-９])\s*年/);
          var years = ym ? parseInt(String(ym[1]).normalize('NFKC'), 10) : 0;
          var prev = lead && !reading && i > 0 && segs[i - 1].paren ? segs[i - 1].text : '';
          var cat = routineCategoryIn_(next && next.paren ? next.text : '') || routineCategoryIn_(prev);
          add(chunk, chunk.trim() + (reading ? '（' + reading + '）' : '') + m[0], years, cat);
        }
      }
    });
    return out;
  }
  // どこにも「さん」が無い書き方：括弧の外の文字を「、」で区切って、それぞれお名前とみなす
  lines.forEach(function (line) {
    if (!line || routineIsBlank_(line.replace(/[\s　]/g, ''))) return;
    var segs = routineParenSegs_(line), cat = '';
    var catM = line.match(/カテゴリー\s*[：:]?\s*([^\s　、,，（(]+(?:[（(][^）)]*[）)])?)/);
    if (catM) cat = catM[1];
    segs.forEach(function (x) { if (x.paren && !cat) cat = routineCategoryIn_(x.text); });
    var plain = segs.filter(function (x) { return !x.paren; }).map(function (x) { return x.text; }).join('').replace(/カテゴリー.*$/, '');
    plain.split(/[、,，・･／\/]+/).forEach(function (p) {
      var t = p.replace(/^.*[：:]/, '').trim();
      if (t) add(t, t, 0, cat);
    });
  });
  return out;
}

// 3,080 のような数（「,」をそろえる）。数でなければそのまま
function routineNum_(s) {
  var t = String(s == null ? '' : s).normalize('NFKC').replace(/[,\s]/g, '');
  if (!/^\d+$/.test(t)) return String(s == null ? '' : s).trim();
  return t.replace(/^0+(?=\d)/, '').replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
// 「54億0,272万円」→「54億272万円」、「4,000万円」→ そのまま。億・万・円の形にそろえる
function routineYen_(s) {
  var t = String(s == null ? '' : s).normalize('NFKC').replace(/\s/g, '');
  var m = t.match(/^(?:([\d,]+)億)?(?:([\d,]+)万)?([\d,]+)?円?$/);
  if (!m || !(m[1] || m[2] || m[3])) return t;
  return (m[1] ? routineNum_(m[1]) + '億' : '') + (m[2] ? routineNum_(m[2]) + '万' : '') + (m[3] ? routineNum_(m[3]) : '') + '円';
}

// 「42 バイスプレジデントによる報告」の欄（読み上げる文がそのまま書いてある）から、スライドの数字を拾う。
//   チャプター設立以来、月間リファーラル数の平均は308件、2026年8月の月間リファーラル数は281件(70件/週）
//   2026年3月から2026年8月の半年間のリファーラル数の合計は1,927件となります。
//   （クリック）チャプターが発足されてから交わされたビジネスのサンキュー額つまり売上は、 54億8,074万円となります。
// → { avg:'308', month:'2026-08', count:'281', perWeek:'70', from:'2026-03', to:'2026-08', total:'1,927', thanks:'54億8,074万円' }
function routineVpReport_(v) {
  var s = String(v == null ? '' : v).normalize('NFKC').replace(/[\r\n]+/g, ' ');
  if (routineIsBlank_(s.replace(/\s/g, ''))) return null;
  var out = {}, m, ym = function (y, mo) { return y + '-' + ('0' + mo).slice(-2); };
  if ((m = s.match(/月間リファーラル数の平均\s*(?:は|:)?\s*([\d,]+)\s*件/))) out.avg = routineNum_(m[1]);
  if ((m = s.match(/(\d{4})\s*年\s*(\d{1,2})\s*月の月間リファーラル数\s*(?:は|:)?\s*([\d,]+)\s*件(?:\s*\(\s*([\d,]+)\s*件\s*\/\s*週\s*\))?/))) {
    out.month = ym(m[1], m[2]); out.count = routineNum_(m[3]); out.perWeek = m[4] ? routineNum_(m[4]) : '';
  }
  if ((m = s.match(/(\d{4})\s*年\s*(\d{1,2})\s*月から\s*(\d{4})\s*年\s*(\d{1,2})\s*月(?:まで)?の[^\d]*?(?:は|:)\s*(?:ちょうど|約)?\s*([\d,]+)\s*件/))) {
    out.from = ym(m[1], m[2]); out.to = ym(m[3], m[4]); out.total = routineNum_(m[5]);
  }
  if ((m = s.match(/サンキュー額[^\d]*?((?:[\d,]+\s*億\s*)?(?:[\d,]+\s*万\s*)?(?:[\d,]+\s*)?円)/))) out.thanks = routineYen_(m[1]);
  return Object.keys(out).length ? out : null;
}

// ネットワーキングリーダーの部門（スライドのページも、この言葉で見分ける）
var NL_KINDS_ = [
  { key: 'ceu',     label: 'CEU',              re: /CEU/i },
  { key: 'thanks',  label: 'サンキュー',        re: /サンキュー/ },
  { key: 'ext',     label: '外部リファーラル',  re: /外部\s*リファーラル/ },
  { key: 'oto',     label: '1to1',             re: /1\s*to\s*1|ワントゥーワン/i },
  { key: 'visitor', label: 'ビジター招待数',    re: /ビジター\s*招待/ }
];
// 「24 ネットワーキングリーダー」の欄（月初の回に、発表の原稿がそのまま書いてある。ほかの週は「ー」）。
//   2026年8月のネットワーキングリーダーの発表をさせて頂きます。
//   先月のCEU部門は23ポイントで、エンタ―テイメントショ―：泉さんです。
//   なんと！4,000万円の売上に貢献頂きました、生命保険(法人)：丘野さんです！
//   外部リファーラル部門。17件で、わたくし、ベリーダンス：山岸です。          … 「わたくし」はバイスプレジデント本人
//   1to1の回数ですが、今回はおふたりいらっしゃいます。21回で、鍼灸師：梅中さんと、…：若松さんです！
// → { month:'2026-08', items: [{ key, label, value:'23', unit:'ポイント', winners:[{ raw, name, matched, category }] }] }
//   数は、部門の見出しのあとの最初の「数＋単位」。受賞者は、そのあとの「…です」までに書かれた方
function routineNetworkingLeaders_(v) {
  var s = String(v == null ? '' : v).normalize('NFKC').replace(/[\r\n]+/g, ' ');
  if (routineIsBlank_(s.replace(/\s/g, '')) || !/[\d]/.test(s)) return null;
  var out = { month: '', items: [] }, m = s.match(/(\d{4})\s*年\s*(\d{1,2})\s*月の\s*ネットワーキングリーダー/);
  if (m) out.month = m[1] + '-' + ('0' + m[2]).slice(-2);
  // 締めの言葉（「それでは、今回の受賞者を代表し…」）から先は読まない
  var endAt = s.search(/それでは[、,]?\s*今回の受賞者|受賞者を代表|今月は代表して|(?:素晴らしい|大きな)貢献を頂きました/);
  var body = endAt >= 0 ? s.substring(0, endAt) : s, roster = routineRoster_();
  var starts = [];
  NL_KINDS_.forEach(function (k) {
    var at = body.search(k.re);
    if (at >= 0) starts.push({ kind: k, at: at });
  });
  starts.sort(function (a, b) { return a.at - b.at; });
  starts.forEach(function (st, i) {
    var sec = body.substring(st.at, i + 1 < starts.length ? starts[i + 1].at : body.length);
    var vm = sec.match(/((?:[\d,]+\s*億\s*)?[\d,]+\s*(?:万\s*)?)(円|件|回|名|人|ポイント|PT|pt|P)(?![a-z])/);
    if (!vm) return;
    var item = { key: st.kind.key, label: st.kind.label, winners: [],
                 value: st.kind.key === 'thanks' ? routineYen_(vm[1] + vm[2]) : routineNum_(vm[1]),
                 unit: st.kind.key === 'thanks' ? '' : vm[2] };
    var rest = sec.substring(sec.indexOf(vm[0]) + vm[0].length), stop = rest.search(/です|でした/);
    var seg = stop >= 0 ? rest.substring(0, stop) : rest;
    if (/該当者?なし|いらっしゃらな/.test(seg)) { out.items.push(item); return; }
    seg.split(/(?:さん|様)\s*(?:と\s*)?[、,]?\s*|\s*と\s*[、,]\s*/).forEach(function (piece) {
      var p = piece.trim(), self = /わたくし|私/.test(p);
      p = p.replace(/わたくし|私/g, '').replace(/^[\s、,!]*(?:で|が)?[\s、,!]*/, '')
           .replace(/^(?:なんと|またもや|こちらも|この部門も|今回も|今月も|そして)[\s、,!]*/, '').trim();
      if (!p) return;
      var colon = p.lastIndexOf(':'), cat = '', raw;
      if (colon >= 0) { cat = p.substring(0, colon); raw = p.substring(colon + 1); }
      else {
        var no = p.lastIndexOf('の');                                    // 古い書き方「贈答用生鮮食品の船越さん」
        if (no >= 0 && p.length - no - 1 >= 1 && p.length - no - 1 <= 8) { cat = p.substring(0, no); raw = p.substring(no + 1); }
        else raw = p;
      }
      raw = raw.replace(/^.*[、,]/, '').trim();
      cat = cat.replace(/^.*[、,]/, '').trim();
      if (!raw || !routineNameLike_(raw)) return;
      var full = routineRosterName_(raw, roster, cat);
      item.winners.push({ raw: raw + (self ? '' : 'さん'), name: full, matched: !!full, category: cat, self: self });
    });
    out.items.push(item);
  });
  return out.items.length ? out : null;
}

// その月の最初の定例会か（休会日の週は飛ばす。前の週に同じ月の開催があれば、最初ではない）
function routineFirstOfMonth_(d) {
  var hol = [];
  try { hol = getHolidays(); } catch (e) {}
  for (var k = 1; k <= 5; k++) {
    var p = new Date(d.getTime()); p.setDate(p.getDate() - 7 * k);
    if (p.getMonth() !== d.getMonth()) return true;
    if (hol.indexOf(fmtDate_(p)) < 0) return false;
  }
  return true;
}

// 開催日の欄を読む。画面から直接呼べる。
// opts.firstHalf … 前半スライドの画面から。新メンバー・更新メンバー・バイスプレジデントによる報告・
//                  ネットワーキングリーダーの欄も読む（バイスプレジデントによる報告が空欄なら、前の回の記載を使う）
function getRoutineInfo(dateStr, opts) {
  try {
    var t = parseDate_(dateStr);
    if (!t) return { ok: false, found: false, message: '開催日を解釈できませんでした。' };
    var hit = findRoutineColumn_(t);
    if (!hit) {
      return { ok: true, found: false,
        message: 'ルーティンチェックシートに ' + fmtDate_(t) + ' の列が見つかりませんでした。' };
    }
    var grid = routineSheetGrid_(hit.name, hit.sheet) || [];     // 候補の日が同じシートなら、読むのは1回

    // 項目の行を探し、その開催日の列の値を返す
    var pick = function (labels) {
      var r = routineFindRow_(grid, labels);
      return { value: r < 0 ? '' : String(grid[r][hit.col - 1] == null ? '' : grid[r][hit.col - 1]).trim() };
    };

    var no = pick(['定例会回数']);
    var core = pick(['BNI目的と概要', 'BNIの目的と概要']);
    var long = pick(ROUTINE_STARTUP_LABELS_);
    var main = pick(['メインプレゼン']);
    var wanted = pick(['募集カテゴリー', 'チャプターが求める', '求める専門分野']);
    var open = pick(['開放カテゴリー']);
    var review = pick(['審査中カテゴリー', '審査中の申込み']);
    var policy = pick(['一般規定']);
    var reco = pick(['推薦の言葉', '推薦のことば']);
    // アンバサダー・ディレクターなど、その日に来られるリージョンの方（「大庭ED・坂上アンバサダー」など）
    var region = pick(['リージョン参加者']);
    // ウィークリープレゼンの始まり（「建築　住まい　22番　熊田さん」など）
    var weekly = pick(ROUTINE_WEEKLY_LABELS_);
    var cv = coreValueOf_(core.value);
    var pres = routineMemberName_(long.value);
    var mains = routineMainPresenters_(main.value);

    return { ok: true, found: true, sheetName: hit.name, term: hit.term,
             date: fmtDate_(t),
             meetingNo: String(no.value || '').replace(/[^\d]/g, ''),
             coreValue: cv ? cv.label : '', coreValueRaw: core.value,
             longPresenter: pres.name, longPresenterRaw: pres.raw,
             longPresenterUnmatched: (!!pres.raw && !pres.matched),
             mainPresenters: mains, mainPresentersRaw: main.value,
             wantedCategories: routineList_(wanted.value), wantedCategoriesRaw: wanted.value,
             openCategory: routineText_(open.value),
             reviewCategory: routineText_(review.value),
             generalPolicy: routinePolicyNo_(policy.value), generalPolicyRaw: policy.value,
             recommendations: routineRecommendations_(reco.value), recommendationsRaw: reco.value,
             regionGuestsRaw: routineText_(region.value),
             weeklyStartRaw: routineText_(weekly.value),
             firstHalf: (opts && opts.firstHalf) ? routineFirstHalf_(t, pick) : null };
  } catch (e) {
    console.error('[ROUTINE] ' + (e && e.stack ? e.stack : e));
    return { ok: false, found: false, message: 'ルーティンチェックシートの読み取りに失敗しました: '
             + (e && e.message ? e.message : e) };
  }
}

// 前半スライドだけで使う欄（getRoutineInfo の opts.firstHalf）
function routineFirstHalf_(t, pick) {
  var nw = pick(ROUTINE_NEW_LABELS_), rn = pick(ROUTINE_RENEW_LABELS_);
  var vp = pick(ROUTINE_VP_LABELS_), nl = pick(ROUTINE_NL_LABELS_);
  var vpRep = routineVpReport_(vp.value), vpFrom = vpRep ? fmtDate_(t) : '';
  if (!vpRep) {
    // その日の欄がまだ空なら、いちばん近い前の回の記載（「できれば前週をコピペ」の運用のため）
    var rows = {}, key = fmtDate_(t), best = '';
    try { rows = routineRowValues_(ROUTINE_VP_LABELS_); } catch (e) {}
    Object.keys(rows).forEach(function (k) { if (k < key && k > best && routineVpReport_(rows[k])) best = k; });
    if (best) { vpRep = routineVpReport_(rows[best]); vpFrom = best; }
  }
  var nlRep = routineNetworkingLeaders_(nl.value);
  if (nlRep) {
    var vice = '';
    try { vice = roleHolders_(t).vice || ''; } catch (e) {}
    nlRep.items.forEach(function (it) {
      it.winners.forEach(function (w) { if (w.self && !w.matched && vice) { w.name = vice; w.matched = true; } });
    });
  }
  return { newMembers: routineMemberList_(nw.value), newMembersRaw: routineText_(nw.value),
           renewMembers: routineMemberList_(rn.value), renewMembersRaw: routineText_(rn.value),
           vpReport: vpRep, vpReportFrom: vpFrom,
           networkingLeaders: nlRep, networkingLeadersRaw: routineText_(nl.value),
           firstOfMonth: routineFirstOfMonth_(t) };
}

// 開催日をまとめて引く（画面の初期表示で、候補ぶんを一度に取るため）
function getRoutineInfoFor_(dates) {
  var out = {};
  for (var i = 0; i < (dates || []).length; i++) {
    try { out[dates[i]] = getRoutineInfo(dates[i]); } catch (e) { out[dates[i]] = { ok: false, found: false }; }
  }
  return out;
}
