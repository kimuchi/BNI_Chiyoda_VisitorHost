// === ルーティンチェックシートから、その日の決めごとを読む ===
//
// チャプターでは開催日ごとの準備を「【NN期】ルーティンチェックシート」で管理している。
// 1行目に開催日、2行目に定例会回数が並び、その下に項目が縦に並ぶ表。
// ここに書いてある「コアバリュー」「2分30秒プレゼンの方」を、
// 各画面の初期値として読み込む（同じことを二度入力しなくて済むように）。
//
//   1行目  … 定例会開催日（1列おきに入っている）
//   2行目  … 定例会回数
//   …行    … B〜E列のどこかに項目名。その開催日の列に値が入る
//
// 期（17期・18期…）が変わるとシートも変わるが、開催日そのもので探すので
// 期の計算は要らない。項目の行番号も期によってずれるため、名前で探す。

var ROUTINE_SHEET_RE_ = /^【(\d+)期】ルーティンチェックシート$/;
var ROUTINE_SCAN_ROWS_ = 90;      // 項目名を探す範囲（これより下は見ない）
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
var ROUTINE_INDEX_ = null;
function routineIndex_() {
  if (ROUTINE_INDEX_) return ROUTINE_INDEX_;
  var sheets = getSS_().getSheets(), idx = {};
  for (var i = 0; i < sheets.length; i++) {
    var m = sheets[i].getName().match(ROUTINE_SHEET_RE_);
    if (!m) continue;
    var last = sheets[i].getLastColumn();
    if (last < 2) continue;
    var row1 = sheets[i].getRange(1, 1, 1, last).getValues()[0];
    for (var c = 0; c < row1.length; c++) {
      var d = routineDate_(row1[c]);
      if (!d) continue;
      var key = fmtDate_(d);
      // 同じ日が複数の期にあれば、期の新しい方を採る
      if (idx[key] && idx[key].term >= parseInt(m[1], 10)) continue;
      idx[key] = { sheet: sheets[i], name: sheets[i].getName(), term: parseInt(m[1], 10), col: c + 1 };
    }
  }
  ROUTINE_INDEX_ = idx;
  return idx;
}

// 開催日から、その日が載っているシートと列を探す
function findRoutineColumn_(target) {
  return routineIndex_()[fmtDate_(target)] || null;
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

// ある項目を、すべての開催日について読む → { 'yyyy/MM/dd': 値 }（空欄の日は入れない）。
// 「ウィークリープレゼン」のように、その日に書かれていなければ
// 前回までの記載から数えたい項目に使う。シートごとに1回だけ読む。
var ROUTINE_WEEKLY_LABELS_ = ['ウィークリープレゼン'];
var ROUTINE_ROWS_CACHE_ = {};
function routineRowValues_(labels) {
  var ck = labels.join('|');
  if (ROUTINE_ROWS_CACHE_[ck]) return ROUTINE_ROWS_CACHE_[ck];
  var idx = routineIndex_(), bySheet = {}, out = {}, key, name;
  for (key in idx) {
    var e = idx[key];
    if (!bySheet[e.name]) bySheet[e.name] = { sheet: e.sheet, cols: [] };
    bySheet[e.name].cols.push({ key: key, col: e.col });
  }
  for (name in bySheet) {
    var sh = bySheet[name].sheet;
    var rows = Math.min(ROUTINE_SCAN_ROWS_, sh.getLastRow()), last = sh.getLastColumn();
    if (rows < 1 || last < 1) continue;
    var grid = sh.getRange(1, 1, rows, last).getValues(), r = routineFindRow_(grid, labels);
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

// 「なし」「無し」「休会」などは、指定されていないものとして扱う。
// 「なし（4/2火曜10時時点）」のように但し書きが付くことがあるので、先頭だけ見る。
function routineIsBlank_(s) {
  return !s || /^(なし|無し|ナシ|ない|―|－|-|—|休会|済|未定|\?|？)/.test(s);
}

// 「谷村さん」のような書き方を、メンバー名簿のフルネームに合わせる。
// 「３５番船木さん」「内装業　石渕さん」のように、番号や業種が前に付くことがあるため、
// そのままで当たらなければ、前置きを落としながら何通りか試す。
function routineMemberName_(raw) {
  var s0 = String(raw == null ? '' : raw).trim();
  if (routineIsBlank_(s0.replace(/[\s　]/g, ''))) return { raw: '', name: '', matched: false };

  var tries = [], push = function (x) {
    x = String(x || '').replace(/[\s　]/g, '');
    if (x && tries.indexOf(x) < 0) tries.push(x);
  };
  push(s0);
  push(s0.replace(/(さん|様|さま|くん|君|ちゃん|氏)?へ$/, '$1'));   // 「山本さんへ」
  push(s0.replace(/(さん|様|さま|くん|君|ちゃん|氏)?へ?$/, ''));    // 「長見くん」
  push(s0.replace(/^[0-9０-９]+\s*番?/, ''));            // 「３５番船木さん」
  var parts = s0.split(/[\s　]+/);
  if (parts.length > 1) push(parts[parts.length - 1]);   // 「内装業　石渕さん」
  push(s0.replace(/[（(].*$/, ''));                       // 「谷村さん（代理）」

  var members = getMembersList(), i, k;
  for (k = 0; k < tries.length; k++) {
    var full = matchInviterToMember(tries[k], members);
    for (i = 0; i < members.length; i++) {
      if (normName_(members[i].name) === normName_(full)) {
        return { raw: s0, name: members[i].name, matched: true };
      }
    }
  }
  // 字の違い（「渡辺さん」と名簿の「渡邉 真理子」など）でも、名字が1人に決まるなら合わせる
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

// 「22 メインプレゼン」の欄。「①山本さん　②金子さん」「①木村‗②若林」など。
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
//    ⓵藤田さん⇒金子さん、
//    ②山口さん⇒山本さん
//   アフター
//    ①船木さん→金子さん
// のように、見出しで「いつ発表するか」を分けて書かれる。見出しの下の組にその時間を付ける。
// 「佐藤さん→近藤さん（アフター）」のように、組のうしろの括弧に書かれていることもある（その組だけに付ける）。
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
    // 組のうしろの括弧に時間が書いてあれば（「佐藤さん→近藤さん（アフター）」）、その行の組だけに使う
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
      // 行頭の記号（「・船木さん→…」の・など）と、前後の区切りを落とす
      var seg = segs[k].replace(/^[\s　.．、,・･•●◆◇■□▪*＊\-－]+/, '').replace(/[\s　、,]+$/, '');
      if (!ARROW.test(seg)) continue;
      // 1つの区切りに2組以上続けて書かれていたら（「金子さん→三澤さん 藤田さん→長見さん」）、組ごとに分ける
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

// 開催日の欄を読む。画面から直接呼べる。
function getRoutineInfo(dateStr) {
  try {
    var t = parseDate_(dateStr);
    if (!t) return { ok: false, found: false, message: '開催日を解釈できませんでした。' };
    var hit = findRoutineColumn_(t);
    if (!hit) {
      return { ok: true, found: false,
        message: 'ルーティンチェックシートに ' + fmtDate_(t) + ' の列が見つかりませんでした。' };
    }
    var rows = Math.min(ROUTINE_SCAN_ROWS_, hit.sheet.getLastRow());
    var grid = hit.sheet.getRange(1, 1, rows, hit.col).getValues();

    // 項目の行を探し、その開催日の列の値を返す
    var pick = function (labels) {
      var r = routineFindRow_(grid, labels);
      return { value: r < 0 ? '' : String(grid[r][hit.col - 1] == null ? '' : grid[r][hit.col - 1]).trim() };
    };

    var no = pick(['定例会回数']);
    var core = pick(['BNI目的と概要', 'BNIの目的と概要']);
    var long = pick(['2分30秒プレゼン', '2分30秒']);
    var main = pick(['メインプレゼン']);
    var wanted = pick(['募集カテゴリー', 'チャプターが求める', '求める専門分野']);
    var open = pick(['開放カテゴリー']);
    var review = pick(['審査中カテゴリー', '審査中の申込み']);
    var policy = pick(['一般規定']);
    var reco = pick(['推薦の言葉', '推薦のことば']);
    // アンバサダー・ディレクターなど、その日に来られるリージョンの方（「吉田ED・坂爪アンバサダー」など）
    var region = pick(['リージョン参加者']);
    // ウィークリープレゼンの始まり（「建築　住まい　22番　熊谷さん」など）
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
             weeklyStartRaw: routineText_(weekly.value) };
  } catch (e) {
    console.error('[ROUTINE] ' + (e && e.stack ? e.stack : e));
    return { ok: false, found: false, message: 'ルーティンチェックシートの読み取りに失敗しました: '
             + (e && e.message ? e.message : e) };
  }
}

// 開催日をまとめて引く（画面の初期表示で、候補ぶんを一度に取るため）
function getRoutineInfoFor_(dates) {
  var out = {};
  for (var i = 0; i < (dates || []).length; i++) {
    try { out[dates[i]] = getRoutineInfo(dates[i]); } catch (e) { out[dates[i]] = { ok: false, found: false }; }
  }
  return out;
}
