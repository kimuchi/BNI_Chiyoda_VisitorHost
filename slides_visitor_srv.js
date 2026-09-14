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
  var data = sh.getDataRange().getValues(), visitors = [], dairi = [];
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
    if (/^代理/.test(no)) dairi.push(item); else visitors.push(item);
  }
  return { visitors: visitors, dairi: dairi };   // 並び順はシートのまま（ソートしない）
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
    return { ok: true, visitors: r.visitors, dairi: r.dairi,
             message: 'ビジター' + r.visitors.length + '名・代理' + r.dairi.length + '名を読み込みました。' };
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

// type: 'presen' | 'intro' | 'dairi'
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
