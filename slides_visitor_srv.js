// === ビジター・代理スライド作成（BNI SLIDE GENERATOR の移植）===
// 入力は「yyyyMMdd参加者」シートそのもの。Excelのアップロードは不要。
// テンプレートpptxは ⚙️PowerPointテンプレートの登録 で Drive に常設したものを使う。

// 紹介／代理紹介スライド（3人1枚）のシェイプID。※2人目の氏名は 27 ではなく 28
// labels は左側の見出し（「氏名」と「専門分野／招待者」）。
// 人数が3の倍数でないとき、空き枠にこの見出しだけが残ってしまうため、
// 値と一緒に空にする。運用で手作業で消していたのと同じ扱い。
var SLIDE_BLOCKS_ = [
  { category: 19, inviter: 20, name: 21, labels: [17, 18] },
  { category: 25, inviter: 26, name: 28, labels: [23, 24] },
  { category: 31, inviter: 32, name: 33, labels: [29, 30] }
];

function openVisitorSlideDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('slides_visitor').setWidth(720).setHeight(660),
    'ビジター・代理スライド作成');
}

// 参加者シートを読み、ビジター／代理に振り分ける
function parseParticipantSheet_(sheetName) {
  var sh = getSS_().getSheetByName(sheetName);
  if (!sh) return null;
  var data = sh.getDataRange().getValues(), visitors = [], guests = [], dairi = [], cancelled = [], statusIdx = -1;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var no   = String(row[0] == null ? '' : row[0]).trim();
    var name = String(row[1] == null ? '' : row[1]).trim();
    if (no === 'No.') {                                  // 見出しの行：Spreadingのステータスの列を探す
      statusIdx = row.map(function (h) { return String(h).trim(); }).indexOf('ステータス');
      if (statusIdx < 0) statusIdx = row.map(function (h) { return String(h).trim(); }).indexOf('出席ステータス');
    }
    // ヘッダー行・空行は値で判定して飛ばす（元ツールと同じ方式）
    if (!no || no === 'No.' || !name || name === '参加者氏名') continue;
    // Spreadingでキャンセルになった方は、スライドに入れない（ようこそのページ・プレゼンのページに出さない）
    if (statusIdx >= 0 && /キャンセル|cancel/i.test(String(row[statusIdx] == null ? '' : row[statusIdx]))) { cancelled.push(name); continue; }
    var item = {
      no: no,
      name: name,
      kana:     String(row[2] == null ? '' : row[2]).trim(),
      category: String(row[3] == null ? '' : row[3]).trim(),
      company:  String(row[4] == null ? '' : row[4]).trim(),
      inviter:  String(row[5] == null ? '' : row[5]).trim()
    };
    // No. は V01 / G01 / 代理28 の形。ゲストはビジターと別に数える
    if (/^代理/.test(no)) dairi.push(item);
    else if (/^G/i.test(no)) guests.push(item);
    else visitors.push(item);
  }
  return { visitors: visitors, guests: guests, dairi: dairi, cancelled: cancelled };   // 並び順はシートのまま（ソートしない）
}

function getVisitorSlideContext() {
  try {
    var sheets = getExistingVisitorSheets();          // 既存関数。新しい順
    var tpl = getTemplateStatus().templates, ready = {};
    for (var i = 0; i < tpl.length; i++) ready[tpl[i].kind] = tpl[i].registered;
    // 既定はメールと同じ「次回の定例会」（無ければ直近）。いちばん新しいシートだと、翌週の名簿を先に作ったときに翌週になってしまう
    var ctx = getEmailContext(), def = (ctx && ctx.ok && ctx.defaultSheet) ? ctx.defaultSheet : (sheets.length ? sheets[0] : '');
    return { ok: true, sheets: sheets, defaultSheet: def, templates: ready };
  } catch (e) {
    console.error('[VSLIDE] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '初期表示の取得に失敗しました: ' + (e && e.message ? e.message : e), sheets: [], templates: {} };
  }
}

function previewVisitorSlideData(sheetName) {
  try {
    if (!sheetName) return { ok: false, message: 'シートが選択されていません。' };
    var r = parseParticipantSheet_(sheetName);
    if (!r) return { ok: false, message: 'シート「' + sheetName + '」が見つかりません。' };
    return { ok: true, visitors: r.visitors, guests: r.guests, dairi: r.dairi, cancelled: r.cancelled,
             message: 'ビジター' + r.visitors.length + '名・ゲスト' + r.guests.length
                    + '名・代理' + r.dairi.length + '名を読み込みました。'
                    + (r.cancelled.length ? '\nSpreadingでキャンセルの方（スライドに入れません）: ' + r.cancelled.join('、') : '') };
  } catch (e) {
    console.error('[VSLIDE] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 長い文字を枠に収めるときの下限（pt）。これ以上は小さくせず、折り返す。
var FIT_MIN_ = { presenName: 40, presenCompany: 16, presenCategory: 14,
                 groupName: 16, groupValue: 10 };
// ビジタープレゼンの会社名・【カテゴリー】の大きさ。メンバーのページ（slides_layout.html）と同じ、会社名44pt・カテゴリー32pt
var PRESEN_PT_ = { company: 44, category: 32 };

// プレゼンスライド1枚分（ID25=氏名+様 / ID27=会社名 / ID29=【カテゴリー】）
// 会社名・【カテゴリー】は、1人ずつ小さくすると大きさがバラバラになるので、決まった大きさ（PRESEN_PT_）で入れる。
// 1行に入らなければ同じ大きさで2行にし、増えた高さのぶん【カテゴリー】を下げる。
// 下の図形（カウントダウンなど）まで余地が足りないときだけ、足りる大きさまで小さくする。枠の幅はテンプレートから読む。
// sec … カウントダウンの秒数（チャプターの設定のビジタープレゼン）。テンプレートと違うときだけ作り直す
//       （始め方はテンプレートのまま。最後に鳴るベルの音も残る）
// カウントダウンが終わっても、次の方のページへは自動で進めない（クリックで次へ）。
// テンプレートに「○秒後に次へ」が入っていても外す
function buildPresenXml_(xml, v, sec) {
  var nm = v.name + ' 様', cat = '【' + v.category + '】';
  xml = setTextInShape_(xml, 25, nm);
  xml = setTextInShape_(xml, 27, v.company);
  xml = setCategoryInShape_(xml, 29, v.category);
  xml = fitFontToShape_(xml, 25, nm, FIT_MIN_.presenName);
  xml = fitStackFixed_(xml, { id: 27, text: v.company, pt: PRESEN_PT_.company, minPt: FIT_MIN_.presenCompany, split: true },
                            { id: 29, text: cat, pt: PRESEN_PT_.category, minPt: FIT_MIN_.presenCategory });
  var now = sec ? mpCountdownSeconds_(xml) : 0;
  if (now && sec !== now) xml = mpSetCountdown_(xml, sec);
  return mpNoAutoAdvance_(xml);
}

// 紹介／代理紹介スライド1枚分（3人）。空き枠は空文字で上書きしてダミー文字を消す
// 公式ファイルから作った雛形のように「専門分野：」「招待者：」の見出しが値と同じ枠に入っているときは、
// 見出しを残して後ろに値を入れる（空き枠は見出しごと消す）
function buildGroupXml_(xml, trio) {
  for (var i = 0; i < 3; i++) {
    var v = trio[i], b = SLIDE_BLOCKS_[i];
    var cp = groupLabelOf_(xml, b.category), ip = groupLabelOf_(xml, b.inviter);
    var cat = v ? cp + v.category : '', inv = v ? ip + v.inviter : '';
    xml = setTextInShape_(xml, b.category, cat);
    xml = setTextInShape_(xml, b.inviter,  inv);
    xml = setTextInShape_(xml, b.name,     v ? v.name + ' 様' : '');
    // 長い専門分野・会社名は枠からあふれるので、文字数に応じて小さくする
    if (v) {
      xml = fitFontToShape_(xml, b.category, cat, FIT_MIN_.groupValue);
      xml = fitFontToShape_(xml, b.inviter,  inv, FIT_MIN_.groupValue);
      xml = fitFontToShape_(xml, b.name,     v.name + ' 様', FIT_MIN_.groupName);
    }
    // 空き枠は左の見出しも消す。値だけ空にすると「氏名」「専門分野」「招待者」が残る
    if (!v && b.labels) {
      for (var L = 0; L < b.labels.length; L++) xml = setTextInShape_(xml, b.labels[L], '');
    }
  }
  return xml;
}

// 枠の文字の頭にある見出し（「専門分野：」「招待者：」）。無ければ空
function groupLabelOf_(xml, shapeId) {
  var r = findShapeRange_(xml, shapeId);
  if (!r) return '';
  var t = slideText_(xml.substring(r.start, r.end)).replace(/^[\s　]+/, '');
  var m = t.match(/^([^：:\s　]{1,8}[：:])/);
  return m ? m[1] : '';
}

// 「定例会20260930_（ビジタープレゼン）09301412.pptx」。開催日は参加者シートの名前（20260930参加者）から
function visitorSlideFileName_(sheetName, part) {
  var key = (String(sheetName || '').match(/^\d{8}|^\d{4}/) || [''])[0];
  return slideFileName_(key ? meetingDateFromKey_(key) : null, part);
}

function makeGroups_(list) {
  var groups = [];
  for (var i = 0; i < list.length; i += 3) groups.push([list[i], list[i + 1] || null, list[i + 2] || null]);
  return groups;
}

// 紹介スライドの見出し（シェイプID15）。
// ビジター・ゲスト・代理で同じレイアウトを使い、見出しの文字だけ変える。
var INTRO_TITLE_SHAPE_ID_ = 15;
var INTRO_SECTIONS_ = [
  { key: 'visitors', title: '歓迎 本日のビジター', label: 'ビジター' },
  { key: 'guests',   title: '歓迎 本日のゲスト',   label: 'ゲスト' },
  { key: 'dairi',    title: '代理出席の方々',       label: '代理' }
];

// ビジター紹介・ゲスト紹介・代理紹介を1つのpptxにまとめて作る。
// 3つとも同じレイアウトなので、テンプレートは「ビジター紹介」1つだけを使い、
// 見出し（シェイプID15）を節ごとに差し替える。
// 節ごとに別のテンプレートを使いたい場合は、単体の作成ボタンをお使いください。
function generateAllIntroSlides(sheetName) {
  try {
    if (!sheetName) return { ok: false, message: 'シートが選択されていません。' };
    var baseBlob = getTemplateBlob_('intro');
    if (!baseBlob) return { ok: false, message: '「ビジター紹介（3人1枚）」のテンプレートが未登録です。メニュー「⚙️ 設定」＞「PowerPointテンプレート（ビジター用）」から登録してください。' };

    var parsed = parseParticipantSheet_(sheetName);
    if (!parsed) return { ok: false, message: 'シート「' + sheetName + '」が見つかりません。' };

    var slides = [], counts = [];
    for (var i = 0; i < INTRO_SECTIONS_.length; i++) {
      var sec = INTRO_SECTIONS_[i], list = parsed[sec.key] || [];
      if (!list.length) continue;
      var groups = makeGroups_(list);
      for (var g = 0; g < groups.length; g++) slides.push({ trio: groups[g], title: sec.title });
      counts.push(sec.label + ' ' + list.length + '名（' + groups.length + '枚）');
    }
    if (!slides.length) return { ok: false, message: 'このシートにビジター・ゲスト・代理のいずれもいません。' };

    var outName = visitorSlideFileName_(sheetName, '紹介スライド・まとめて');
    var noTitle = false;
    var blob = buildPptxFromTemplate_(baseBlob, slides, function (xml, item) {
      if (!findShapeRange_(xml, INTRO_TITLE_SHAPE_ID_)) noTitle = true;
      xml = setTextInShape_(xml, INTRO_TITLE_SHAPE_ID_, item.title);
      return buildGroupXml_(xml, item.trio);
    }, outName);
    // 見出しの枠が無いテンプレートでは、ゲスト・代理のページも「本日のビジター」の見出しのままになる。黙っていない
    var titleNote = (noTitle && counts.length > 1) || (noTitle && !parsed.visitors.length)
      ? '\nテンプレートに見出しの枠（図形の番号15）が無いため、ゲスト・代理のページの見出しはテンプレートのままです。' : '';

    var saved = saveOutputFile_(blob, outName);
    console.log('[VSLIDE] all-intro ' + outName + ' slides=' + slides.length);
    return { ok: true, url: saved.url, downloadUrl: saved.downloadUrl, fileName: outName, slideCount: slides.length,
             message: '紹介スライドをまとめて作成しました（' + counts.join(' / ') + '　合計' + slides.length + '枚）。' + titleNote
               + (parsed.cancelled.length ? '\nSpreadingでキャンセルの方は入れていません: ' + parsed.cancelled.join('、') : '') };
  } catch (e) {
    console.error('[VSLIDE] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'スライドの作成に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// type: 'presen' | 'intro' | 'guest' | 'dairi'
function generateVisitorSlides(sheetName, type) {
  try {
    console.log('[VSLIDE] generate type=' + type + ' sheet=' + sheetName);
    if (!sheetName) return { ok: false, message: 'シートが選択されていません。' };
    var def = TEMPLATE_KINDS_[type];
    if (!def) return { ok: false, message: 'スライドの種類が不正です。' };

    var tplBlob = getTemplateBlob_(type);
    if (!tplBlob) return { ok: false, message: '「' + def.label + '」のテンプレートが未登録です。メニュー「⚙️ PowerPointテンプレートの登録」から登録してください。' };

    var parsed = parseParticipantSheet_(sheetName);
    if (!parsed) return { ok: false, message: 'シート「' + sheetName + '」が見つかりません。' };

    var dataList, builder, outName, countLabel;
    if (type === 'presen') {
      if (!parsed.visitors.length) return { ok: false, message: 'このシートにビジターがいません。' };
      dataList = parsed.visitors;
      var vsec = chapterPresenSeconds_().visitor;
      builder = function (xml, v) { return buildPresenXml_(xml, v, vsec); };
      outName = visitorSlideFileName_(sheetName, 'ビジタープレゼン');
      countLabel = parsed.visitors.length + '名・' + dataList.length + '枚・カウントダウン' + chapterSecondsLabel_(vsec);
    } else if (type === 'intro') {
      if (!parsed.visitors.length) return { ok: false, message: 'このシートにビジターがいません。' };
      dataList = makeGroups_(parsed.visitors);
      builder = buildGroupXml_;
      outName = visitorSlideFileName_(sheetName, 'ビジター紹介');
      countLabel = parsed.visitors.length + '名・' + dataList.length + '枚';
    } else if (type === 'guest') {
      if (!parsed.guests.length) return { ok: false, message: 'このシートにゲストがいません。' };
      dataList = makeGroups_(parsed.guests);
      builder = buildGroupXml_;
      outName = visitorSlideFileName_(sheetName, 'ゲスト紹介');
      countLabel = parsed.guests.length + '名・' + dataList.length + '枚';
    } else {
      if (!parsed.dairi.length) return { ok: false, message: 'このシートに代理の方がいません。' };
      dataList = makeGroups_(parsed.dairi);
      builder = buildGroupXml_;
      outName = visitorSlideFileName_(sheetName, '代理紹介');
      countLabel = parsed.dairi.length + '名・' + dataList.length + '枚';
    }

    var blob = buildPptxFromTemplate_(tplBlob, dataList, builder, outName);
    var saved = saveOutputFile_(blob, outName);
    console.log('[VSLIDE] saved ' + outName + ' -> ' + saved.url);
    return { ok: true, message: '「' + def.label + '」を作成しました（' + countLabel + '）。'
               + (parsed.cancelled.length ? '\nSpreadingでキャンセルの方は入れていません: ' + parsed.cancelled.join('、') : ''),
             url: saved.url, downloadUrl: saved.downloadUrl, fileName: outName, slideCount: dataList.length };
  } catch (e) {
    console.error('[VSLIDE] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'スライドの作成に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}
