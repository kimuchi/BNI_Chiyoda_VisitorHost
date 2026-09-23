// === ビジター・代理スライド作成（BNI SLIDE GENERATOR の移植）===
// 入力は「MMdd参加者」シートそのもの。Excelのアップロードは不要。
// テンプレートpptxは ⚙️PowerPointテンプレートの登録 で Drive に常設したものを使う。

// 紹介／代理紹介スライド（3人1枚）のシェイプID。※2人目の氏名は 27 ではなく 28
var SLIDE_BLOCKS_ = [
  { category: 19, inviter: 20, name: 21 },
  { category: 25, inviter: 26, name: 28 },
  { category: 31, inviter: 32, name: 33 }
];

function openVisitorSlideDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('slides_visitor').setWidth(720).setHeight(660),
    'ビジター・代理スライド作成');
}

// 参加者シートを読み、ビジター／代理に振り分ける
function parseParticipantSheet_(sheetName) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return null;
  var data = sh.getDataRange().getValues(), visitors = [], guests = [], dairi = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var no   = String(row[0] == null ? '' : row[0]).trim();
    var name = String(row[1] == null ? '' : row[1]).trim();
    // ヘッダー行・空行は値で判定して飛ばす（元ツールと同じ方式）
    if (!no || no === 'No.' || !name || name === '参加者氏名') continue;
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
  return { visitors: visitors, guests: guests, dairi: dairi };   // 並び順はシートのまま（ソートしない）
}

function getVisitorSlideContext() {
  try {
    var sheets = getExistingVisitorSheets();          // 既存関数。新しい順
    var tpl = getTemplateStatus().templates, ready = {};
    for (var i = 0; i < tpl.length; i++) ready[tpl[i].kind] = tpl[i].registered;
    return { ok: true, sheets: sheets, defaultSheet: sheets.length ? sheets[0] : '', templates: ready };
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
    return { ok: true, visitors: r.visitors, guests: r.guests, dairi: r.dairi,
             message: 'ビジター' + r.visitors.length + '名・ゲスト' + r.guests.length
                    + '名・代理' + r.dairi.length + '名を読み込みました。' };
  } catch (e) {
    console.error('[VSLIDE] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// プレゼンスライド1枚分（ID25=氏名+様 / ID27=会社名 / ID29=【カテゴリー】）
function buildPresenXml_(xml, v) {
  xml = setTextInShape_(xml, 25, v.name + ' 様');
  xml = setTextInShape_(xml, 27, v.company);
  xml = setCategoryInShape_(xml, 29, v.category);
  return xml;
}

// 紹介／代理紹介スライド1枚分（3人）。空き枠は空文字で上書きしてダミー文字を消す
function buildGroupXml_(xml, trio) {
  for (var i = 0; i < 3; i++) {
    var v = trio[i], b = SLIDE_BLOCKS_[i];
    xml = setTextInShape_(xml, b.category, v ? v.category : '');
    xml = setTextInShape_(xml, b.inviter,  v ? v.inviter  : '');
    xml = setTextInShape_(xml, b.name,     v ? v.name + ' 様' : '');
  }
  return xml;
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

    var mmdd = (sheetName.match(/^\d{4}/) || [''])[0];
    var outName = mmdd + '_BNI_紹介スライド_一括.pptx';
    var blob = buildPptxFromTemplate_(baseBlob, slides, function (xml, item) {
      xml = setTextInShape_(xml, INTRO_TITLE_SHAPE_ID_, item.title);
      return buildGroupXml_(xml, item.trio);
    }, outName);

    var saved = saveOutputFile_(blob, outName);
    console.log('[VSLIDE] all-intro ' + outName + ' slides=' + slides.length);
    return { ok: true, url: saved.url, fileName: outName, slideCount: slides.length,
             message: '紹介スライドをまとめて作成しました（' + counts.join(' / ') + '　合計' + slides.length + '枚）。' };
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
    var mmdd = (sheetName.match(/^\d{4}/) || [''])[0];
    if (type === 'presen') {
      if (!parsed.visitors.length) return { ok: false, message: 'このシートにビジターがいません。' };
      dataList = parsed.visitors;
      builder = buildPresenXml_;
      outName = mmdd + '_BNI_プレゼンスライド.pptx';
      countLabel = parsed.visitors.length + '名・' + dataList.length + '枚';
    } else if (type === 'intro') {
      if (!parsed.visitors.length) return { ok: false, message: 'このシートにビジターがいません。' };
      dataList = makeGroups_(parsed.visitors);
      builder = buildGroupXml_;
      outName = mmdd + '_BNI_紹介スライド.pptx';
      countLabel = parsed.visitors.length + '名・' + dataList.length + '枚';
    } else if (type === 'guest') {
      if (!parsed.guests.length) return { ok: false, message: 'このシートにゲストがいません。' };
      dataList = makeGroups_(parsed.guests);
      builder = buildGroupXml_;
      outName = mmdd + '_BNI_ゲスト紹介スライド.pptx';
      countLabel = parsed.guests.length + '名・' + dataList.length + '枚';
    } else {
      if (!parsed.dairi.length) return { ok: false, message: 'このシートに代理の方がいません。' };
      dataList = makeGroups_(parsed.dairi);
      builder = buildGroupXml_;
      outName = mmdd + '_BNI_代理紹介スライド.pptx';
      countLabel = parsed.dairi.length + '名・' + dataList.length + '枚';
    }

    var blob = buildPptxFromTemplate_(tplBlob, dataList, builder, outName);
    var saved = saveOutputFile_(blob, outName);
    console.log('[VSLIDE] saved ' + outName + ' -> ' + saved.url);
    return { ok: true, message: '「' + def.label + '」を作成しました（' + countLabel + '）。', url: saved.url, fileName: outName, slideCount: dataList.length };
  } catch (e) {
    console.error('[VSLIDE] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'スライドの作成に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}
