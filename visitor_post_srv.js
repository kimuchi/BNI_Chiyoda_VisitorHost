// === ビジター情報の投稿文 ===
//
// 「第535回(9/30)定例会ビジター情報　確定」のような、グループに投稿する文章を作る。
// 材料は「CSVから名簿・PDF作成」で作った参加者シート（YYYYMMDD参加者）。
// SpreadingのCSVの列がそのまま残っているので、種別（ビジター・ゲスト・代理）と
// 入金の状態（支払いステータス）をそこから読む。
//
// 文章の組み立ては画面（visitor_post.html）で行う。入金済み／未入金を画面で
// 切り替えるたびに、その場で文章を作り直せるようにするため。

function openVisitorPostDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('visitor_post').evaluate().setWidth(820).setHeight(760),
    'ビジター情報の投稿文');
}

// SpreadingのCSVは英語の列名で出ることも日本語で出ることもある。
// 英語は取り込み時に日本語へ置き換わる（Payment Status → 支払いステータス など）。
var VP_PAY_STATUS_COLS_ = ['支払いステータス', '支払ステータス', '支払状況', '入金状況', 'Payment Status'];
var VP_FEE_COLS_ = ['費用', 'Fee'];
var VP_PAID_FEE_COLS_ = ['支払い済み', 'Paid Fee'];
var VP_STATUS_COLS_ = ['ステータス', '出席ステータス', 'Status'];

var VP_SETTINGS_KEY_ = 'BNI_VISITOR_POST';
var VP_DEFAULT_FOOTER_ = '※招待者の方は[spreading]にご入力ください。\n\n'
  + '■メンバーの皆様はSpreadingにてビジター様の詳細情報を事前に確認してください。';

function vpCol_(row, names) {
  for (var i = 0; i < names.length; i++) {
    var v = row[names[i]];
    if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

function vpNum_(v) {
  var s = String(v == null ? '' : v).normalize('NFKC').replace(/[^\d.]/g, '');
  return s ? parseFloat(s) : null;
}

// 入金済みか。Spreadingの支払いステータスは「未払い」「支払済み」の2つ
// （英語のCSVなら Unpaid / Paid）。「未払い」も「Unpaid」も「払」「paid」を含むので、
// 未払いの方を先に見る。状態の列が無いときは、費用と支払い済みの金額で判断する。
// どちらでも決まらなければ null（画面で「読めなかった」とお知らせする）。
function visitorPaidOf_(row) {
  var raw = vpCol_(row, VP_PAY_STATUS_COLS_), s = raw.normalize('NFKC').toLowerCase();
  if (s) {
    if (/未|unpaid|not\s*paid|pending/.test(s)) return { paid: false, raw: raw };
    if (/済|paid|完了|入金/.test(s)) return { paid: true, raw: raw };
  }
  var fee = vpNum_(vpCol_(row, VP_FEE_COLS_)), got = vpNum_(vpCol_(row, VP_PAID_FEE_COLS_));
  if (fee !== null && fee > 0 && got !== null) {
    return { paid: got >= fee, raw: raw || ('費用 ' + fee + '／支払い済み ' + got) };
  }
  return { paid: null, raw: raw };
}

// 参加者シートの1行 → 投稿文に使う形
function visitorPostPerson_(r) {
  var t = String(r['種別'] || '').trim(), no = String(r._No || '').trim();
  var type = (/^guest$|ゲスト/i.test(t) || /^G\d/.test(no)) ? 'guest'
           : (/^substitute$|代理/i.test(t) || /^代理/.test(no)) ? 'sub'
           : 'visitor';
  var pay = visitorPaidOf_(r), st = vpCol_(r, VP_STATUS_COLS_);
  var clean = function (v) { return String(v == null ? '' : v).replace(/[\s　]+/g, ' ').trim(); };
  return { no: no, type: type, name: clean(r['参加者氏名']), kana: clean(r['ふりがな']),
           category: clean(r['カテゴリー']), inviter: clean(r['招待者']),
           paid: pay.paid, payRaw: pay.raw,
           cancelled: /キャンセル|cancel/i.test(st), statusRaw: st };
}

// 「20260930参加者」→ 2026/9/30。移行前の「0930参加者」は今年として読む
function visitorPostDateOf_(sheetName) {
  var m = String(sheetName || '').match(/^(\d{8}|\d{4})参加者$/);
  if (!m) return null;
  var k = m[1], y = k.length === 8 ? parseInt(k.slice(0, 4), 10) : new Date().getFullYear();
  var md = k.slice(-4);
  var d = new Date(y, parseInt(md.slice(0, 2), 10) - 1, parseInt(md.slice(2), 10));
  return isNaN(d.getTime()) ? null : d;
}

function getVisitorPostSettings_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(VP_SETTINGS_KEY_);
    var s = raw ? JSON.parse(raw) : {};
    return { footer: (typeof s.footer === 'string') ? s.footer : VP_DEFAULT_FOOTER_ };
  } catch (e) {
    return { footer: VP_DEFAULT_FOOTER_ };
  }
}

// 末尾の決まり文句を保存する（空で保存すると初期の文面に戻る）
function saveVisitorPostFooter(footer) {
  try {
    var text = String(footer == null ? '' : footer).replace(/\r\n?/g, '\n').replace(/\s+$/, '');
    var props = PropertiesService.getScriptProperties();
    if (!text) props.deleteProperty(VP_SETTINGS_KEY_);
    else props.setProperty(VP_SETTINGS_KEY_, JSON.stringify({ footer: text }));
    return { ok: true, footer: text || VP_DEFAULT_FOOTER_,
             message: text ? '末尾の文面を保存しました。次回からこの文面が入ります。' : '末尾の文面を初期のものに戻しました。' };
  } catch (e) {
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 画面の初期表示：参加者シートの一覧（新しい順）と、既定で選ぶシート
function getVisitorPostContext() {
  try {
    var names = getExistingVisitorSheets();          // コード.js（新しい順）
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var sheets = [], pick = '', pickTime = Infinity;
    for (var i = 0; i < names.length; i++) {
      var d = visitorPostDateOf_(names[i]);
      sheets.push({ name: names[i],
                    label: d ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd') : names[i] });
      // 今日以降でいちばん近い開催日を選ぶ（無ければいちばん新しいもの）
      if (d && d.getTime() >= today.getTime() && d.getTime() < pickTime) { pick = names[i]; pickTime = d.getTime(); }
    }
    if (!pick && sheets.length) pick = sheets[0].name;
    return { ok: true, sheets: sheets, defaultSheet: pick, settings: getVisitorPostSettings_() };
  } catch (e) {
    console.error('[VPOST] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 参加者シートを読み、投稿文の材料を返す
function getVisitorPostData(sheetName) {
  try {
    var data = loadSheetData(sheetName);             // コード.js（種別から V01・G01・代理 の番号も振る）
    var people = [];
    for (var i = 0; i < data.rows.length; i++) {
      var p = visitorPostPerson_(data.rows[i]);
      if (p.name) people.push(p);
    }
    var d = visitorPostDateOf_(sheetName), no = '';
    if (d) {
      // 開催回はルーティンチェックシートの「定例会回数」を優先し、無ければ休会日を除いて数える
      try {
        var ri = getRoutineInfo(Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd'));
        if (ri && ri.found && ri.meetingNo) no = String(ri.meetingNo);
      } catch (e) {}
      if (!no) { var c = meetingCountOf_(d); if (c) no = String(c); }
    }
    return { ok: true, sheetName: sheetName, meetingNo: no,
             month: d ? d.getMonth() + 1 : '', day: d ? d.getDate() : '',
             hasPayColumn: data.header.some(function (h) {
               return VP_PAY_STATUS_COLS_.indexOf(h) >= 0 || VP_PAID_FEE_COLS_.indexOf(h) >= 0;
             }),
             people: people };
  } catch (e) {
    console.error('[VPOST] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '「' + sheetName + '」を読めませんでした: ' + (e && e.message ? e.message : e) };
  }
}
