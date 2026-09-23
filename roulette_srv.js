// === 抽選ルーレット作成用シートのビジター招待数 ===
//
// 当日の「〇〇参加者」シートから、No. が V で始まる行（＝ビジター）を招待者ごとに数え、
// 「抽選ルーレット作成用」シートのH列に入れる。
// ゲスト(G)・代理は数えない。招待者が空の行も数えない。
//
// シート上の「ボタン」について:
//   Googleスプレッドシートの図形ボタンはスクリプトからは作れないため、
//   チェックボックスで代用している。チェックを入れると onEdit が拾って実行し、
//   終わったらチェックを自動で外す。結果は右隣のセルに出る。
//   図形にスクリプトを割り当てたい場合は、fillRouletteVisitorCounts を指定すればよい。

var ROULETTE_PREFIX_ = '抽選ルーレット作成用';
var ROULETTE_COUNT_COL_ = 8;        // H列（ビジター招待数）
var ROULETTE_NAME_COL_ = 1;         // A列（姓）。B列が名
var ROULETTE_BTN_CELL_ = 'P1';      // チェックボックス
var ROULETTE_MSG_CELL_ = 'Q1';      // 実行結果

// 氏名の照合キー。スペース・全角半角・「さん」の有無を無視する
function rouletteKey_(name) {
  return String(name == null ? '' : name).normalize('NFKC').replace(/[\s　]/g, '').replace(/さん$/, '').trim();
}

// 当日の参加者シートを探す。
// 新しい開催日を作ると他の開催日は自動で非表示になるため、表示されているものを優先する。
// 名前は MMDD なので年をまたぐと単純な並べ替えでは誤るが、表示中のものを先に見ることで避けている。
function findCurrentVisitorSheet_() {
  var sheets = getSS_().getSheets(), visible = [], all = [];
  for (var i = 0; i < sheets.length; i++) {
    if (!/^\d{4}参加者$/.test(sheets[i].getName())) continue;
    all.push(sheets[i]);
    if (!sheets[i].isSheetHidden()) visible.push(sheets[i]);
  }
  var pick = visible.length ? visible : all;
  if (!pick.length) return null;
  pick.sort(function (a, b) { return a.getName() < b.getName() ? 1 : (a.getName() > b.getName() ? -1 : 0); });
  return pick[0];
}

// 参加者シートから、招待者ごとのビジター人数を数える
function countVisitorsByInviter_(sheet) {
  var data = sheet.getDataRange().getValues(), h = -1;
  for (var i = 0; i < Math.min(10, data.length); i++) {
    if (data[i].indexOf('No.') !== -1) { h = i; break; }
  }
  if (h < 0) throw new Error('「' + sheet.getName() + '」に見出し行（No.）が見つかりません。');
  var hd = data[h], noIdx = hd.indexOf('No.'), invIdx = hd.indexOf('招待者'), nmIdx = hd.indexOf('参加者氏名');
  if (noIdx < 0 || invIdx < 0) throw new Error('「' + sheet.getName() + '」に「No.」または「招待者」の列がありません。');

  var members = getMembersList(), counts = {}, visitors = 0, noInviter = [];
  for (var r = h + 1; r < data.length; r++) {
    var no = String(data[r][noIdx] == null ? '' : data[r][noIdx]).trim();
    if (!/^V/i.test(no)) continue;          // ビジターのみ。ゲスト(G)・代理は対象外
    visitors++;
    // 招待者名はメンバー名簿の表記にそろえる（渡邉/渡辺などの異体字対応）
    var key = rouletteKey_(matchInviterToMember(data[r][invIdx], members));
    if (!key) { noInviter.push(String(data[r][nmIdx >= 0 ? nmIdx : noIdx] || no)); continue; }
    counts[key] = (counts[key] || 0) + 1;
  }
  return { counts: counts, visitors: visitors, noInviter: noInviter, sheetName: sheet.getName() };
}

// 「抽選ルーレット作成用」で始まるシートを全て返す（①②と分かれていても両方に入れる）
function findRouletteSheets_() {
  var sheets = getSS_().getSheets(), out = [];
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf(ROULETTE_PREFIX_) === 0) out.push(sheets[i]);
  }
  return out;
}

// 見出し行（A列が「姓」の行）を探す
function findRouletteHeaderRow_(data) {
  for (var i = 0; i < Math.min(10, data.length); i++) {
    if (String(data[i][0] || '').trim() === '姓') return i;
  }
  return -1;
}

// 本体。メニュー・チェックボックス・図形ボタンのいずれからも呼べる。
function fillRouletteVisitorCounts() {
  var vs = findCurrentVisitorSheet_();
  if (!vs) return { ok: false, message: '参加者シート（例: 0923参加者）が見つかりません。先に「1. CSVから名簿・PDF作成」で作成してください。' };

  var tally;
  try { tally = countVisitorsByInviter_(vs); }
  catch (e) { return { ok: false, message: (e && e.message ? e.message : String(e)) }; }

  var targets = findRouletteSheets_();
  if (!targets.length) return { ok: false, message: '「' + ROULETTE_PREFIX_ + '」で始まるシートが見つかりません。' };

  var members = getMembersList(), filled = 0, written = 0, unknown = [], sheetNames = [];
  for (var t = 0; t < targets.length; t++) {
    var sh = targets[t], data = sh.getDataRange().getValues();
    var hr = findRouletteHeaderRow_(data);
    if (hr < 0) { unknown.push('「' + sh.getName() + '」はA列に「姓」の見出しが無いため飛ばしました'); continue; }
    var out = [];
    for (var r = hr + 1; r < data.length; r++) {
      var sei = String(data[r][ROULETTE_NAME_COL_ - 1] || '').trim();
      var mei = String(data[r][ROULETTE_NAME_COL_] || '').trim();
      if (!sei && !mei) { out.push(['']); continue; }   // 空行はそのまま空に
      // ルーレット側の氏名もメンバー名簿の表記にそろえてから照合する
      var key = rouletteKey_(matchInviterToMember(sei + mei, members));
      var n = tally.counts[key];
      if (n === undefined) n = tally.counts[rouletteKey_(sei + mei)];
      out.push([n === undefined ? 0 : n]);
      if (n) filled++;
      written++;
    }
    if (out.length) sh.getRange(hr + 2, ROULETTE_COUNT_COL_, out.length, 1).setValues(out);
    sheetNames.push(sh.getName());
  }

  var msg = '「' + tally.sheetName + '」のビジター ' + tally.visitors + '名を数えて、'
          + sheetNames.join('／') + ' のH列に入れました（' + written + '行 / 招待者として名前が出たのは ' + filled + '名）。';
  if (tally.noInviter.length) msg += '\n招待者が空のビジター: ' + tally.noInviter.join('、');
  if (unknown.length) msg += '\n' + unknown.join('\n');
  console.log('[ROULETTE] ' + msg.replace(/\n/g, ' / '));
  return { ok: true, message: msg, visitors: tally.visitors, written: written, filled: filled };
}

// メニューから実行
function menuFillRouletteVisitorCounts() {
  var ui = SpreadsheetApp.getUi(), res = fillRouletteVisitorCounts();
  ui.alert('ビジター招待数の記入', (res.ok ? '✅ ' : '❌ ') + res.message, ui.ButtonSet.OK);
}

// シート上のチェックボックス「ボタン」を用意する
function setupRouletteButton() {
  var ui = SpreadsheetApp.getUi(), targets = findRouletteSheets_();
  if (!targets.length) {
    ui.alert('ボタンの設置', '「' + ROULETTE_PREFIX_ + '」で始まるシートが見つかりません。', ui.ButtonSet.OK);
    return;
  }
  var done = [];
  for (var i = 0; i < targets.length; i++) {
    var sh = targets[i];
    sh.getRange(ROULETTE_BTN_CELL_).insertCheckboxes().setValue(false)
      .setNote('チェックを入れると、当日の参加者シートからビジター招待数をH列に入れます。');
    sh.getRange(ROULETTE_MSG_CELL_).setValue('← チェックでビジター招待数を記入')
      .setFontSize(9).setFontColor('#666');
    done.push(sh.getName());
  }
  ui.alert('ボタンの設置',
    done.join('／') + ' の ' + ROULETTE_BTN_CELL_ + ' にチェックボックスを置きました。\n\n'
    + 'チェックを入れると実行され、終わると自動でチェックが外れます。結果は隣の ' + ROULETTE_MSG_CELL_ + ' に出ます。\n\n'
    + '図形のボタンにしたい場合は、挿入＞図形 で図形を作り、右上のメニューから\n'
    + '「スクリプトを割り当て」に fillRouletteVisitorCounts と入力してください。',
    ui.ButtonSet.OK);
}

// チェックボックスを「ボタン」として動かす。
// 単純トリガーなのでスプレッドシートの読み書きのみ。ここでは十分。
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName().indexOf(ROULETTE_PREFIX_) !== 0) return;
    if (e.range.getA1Notation() !== ROULETTE_BTN_CELL_) return;
    if (e.range.getValue() !== true) return;
    var msg = sh.getRange(ROULETTE_MSG_CELL_);
    msg.setValue('実行中...');
    var res = fillRouletteVisitorCounts();
    e.range.setValue(false);                       // 押しっぱなしにしない
    msg.setValue((res.ok ? '✅ ' : '❌ ') + String(res.message).split('\n')[0]);
  } catch (err) {
    try {
      e.range.setValue(false);
      e.range.getSheet().getRange(ROULETTE_MSG_CELL_).setValue('❌ ' + (err && err.message ? err.message : err));
    } catch (e2) {}
  }
}
