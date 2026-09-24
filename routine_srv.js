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
  return { raw: s0, name: '', matched: false };
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

    // 項目名（B〜E列のどこか）で行を探し、その開催日の列の値を返す
    var pick = function (labels) {
      for (var r = 0; r < grid.length; r++) {
        for (var c = 1; c < Math.min(ROUTINE_LABEL_COLS_, grid[r].length); c++) {
          var s = String(grid[r][c] == null ? '' : grid[r][c]).replace(/[\s　]/g, '');
          if (!s) continue;
          for (var k = 0; k < labels.length; k++) {
            if (s.indexOf(labels[k]) === 0) {
              return { row: r + 1, value: String(grid[r][hit.col - 1] == null ? '' : grid[r][hit.col - 1]).trim() };
            }
          }
        }
      }
      return { row: 0, value: '' };
    };

    var no = pick(['定例会回数']);
    var core = pick(['BNI目的と概要', 'BNIの目的と概要']);
    var long = pick(['2分30秒']);
    var cv = coreValueOf_(core.value);
    var pres = routineMemberName_(long.value);

    return { ok: true, found: true, sheetName: hit.name, term: hit.term,
             date: fmtDate_(t),
             meetingNo: String(no.value || '').replace(/[^\d]/g, ''),
             coreValue: cv ? cv.label : '', coreValueRaw: core.value,
             longPresenter: pres.name, longPresenterRaw: pres.raw,
             longPresenterUnmatched: (!!pres.raw && !pres.matched) };
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
