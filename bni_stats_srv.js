// BNIの統計情報。適用月ごとに保存し、過去の開催日には未来の数値を使わない。
var BNI_STATS_PREFIX_ = 'BNI_STATS_MONTH_';
var BNI_STATS_FIELDS_ = [
  { key: 'worldMembers', label: '世界メンバー数', unit: '人', type: 'members', scope: 'world' },
  { key: 'worldChapters', label: '世界チャプター数', unit: 'チャプター', type: 'chapters', scope: 'world' },
  { key: 'worldReferrals', label: '世界リファーラル数', unit: '件', type: 'referrals', scope: 'world' },
  { key: 'worldRevenue', label: '世界売上金額', unit: '円', type: 'revenue', scope: 'world' },
  { key: 'japanMembers', label: '日本メンバー数', unit: '人', type: 'members', scope: 'japan' },
  { key: 'japanChapters', label: '日本チャプター数', unit: 'チャプター', type: 'chapters', scope: 'japan' },
  { key: 'worldCountries', label: '活動する国・地域数', unit: 'か国・地域', type: 'countries', scope: 'world' },
  { key: 'japanReferrals', label: '日本リファーラル数', unit: '件', type: 'referrals', scope: 'japan' },
  { key: 'japanRevenue', label: '日本売上金額', unit: '円', type: 'revenue', scope: 'japan' }
];

function openBniStatsDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('bni_stats').setWidth(660).setHeight(760), 'BNI統計情報（月次更新）');
}
function bniStatsMonth_(v) {
  var s = String(v || '');
  if (!/^20\d{2}-(0[1-9]|1[0-2])$/.test(s)) throw new Error('適用月を年月で選んでください。');
  return s;
}
function bniStatsMonths_() {
  return PropertiesService.getScriptProperties().getKeys().filter(function(k) {
    return /^BNI_STATS_MONTH_20\d{2}-(0[1-9]|1[0-2])$/.test(k);
  }).map(function(k) { return k.slice(BNI_STATS_PREFIX_.length); }).sort().reverse();
}
function bniStatsRecord_(month) {
  var raw = PropertiesService.getScriptProperties().getProperty(BNI_STATS_PREFIX_ + month);
  if (!raw) return null;
  var r = JSON.parse(raw);
  if (!r || r.month !== month || !r.values) throw new Error(month + 'の統計情報を読み込めません。');
  return r;
}
function bniStatsForDate_(date) {
  var month = Utilities.formatDate(date || new Date(), 'Asia/Tokyo', 'yyyy-MM');
  var chosen = bniStatsMonths_().filter(function(m) { return m <= month; })[0];
  return chosen ? bniStatsRecord_(chosen) : null;
}
function getBniStatsSettings(month) {
  requireSheetAccess_();
  month = bniStatsMonth_(month || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM'));
  var months = bniStatsMonths_(), props = PropertiesService.getScriptProperties();
  var own = bniStatsRecord_(month), prev = months.filter(function(m) { return m < month; })[0];
  var base = own || (prev ? bniStatsRecord_(prev) : null);
  return { ok: true, month: month, months: months, fields: BNI_STATS_FIELDS_, saved: !!own,
    sourceMonth: base ? base.month : '', values: base ? base.values : {}, source: base ? base.source || '' : '',
    revision: props.getProperty(BNI_STATS_PREFIX_ + month) || '' };
}
function saveBniStatsSettings(s) {
  var lock = LockService.getScriptLock(), locked = false;
  try {
    requireSheetAccess_();
    s = s || {};
    var month = bniStatsMonth_(s.month), values = {}, count = 0;
    BNI_STATS_FIELDS_.forEach(function(f) {
      var v = String((s.values || {})[f.key] == null ? '' : s.values[f.key]).trim()
        .replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }).replace(/，/g, ',');
      if (v.length > 60 || /[<>{}\x00-\x1f]/.test(v) || (v && !/[0-9]/.test(v))) {
        throw new Error(f.label + 'は数値と単位を60文字以内で入れてください。');
      }
      if (/^[0-9][0-9,.]*$/.test(v)) v += f.unit;
      var figure = v ? bniStatsFigure_(v, f.type, false) : null;
      if (v && (!figure || figure.text !== v)) throw new Error(f.label + 'は数値と単位で入れてください（例: 356,000人超、約4兆5000億円）。');
      values[f.key] = v;
      if (v) count++;
    });
    if (!count) throw new Error('更新する数値を1つ以上入れてください。');
    var source = String(s.source || '').trim();
    if (source.length > 300 || /[\x00-\x1f]/.test(source)) throw new Error('発表元・メモは300文字以内の1行で入れてください。');
    if (typeof s.revision !== 'string') throw new Error('先に設定を読み込んでください。');
    locked = lock.tryLock(10000);
    if (!locked) throw new Error('ほかの方が保存中です。少し待ってから保存してください。');
    var props = PropertiesService.getScriptProperties(), key = BNI_STATS_PREFIX_ + month;
    if ((props.getProperty(key) || '') !== s.revision) throw new Error('ほかの方がこの月の数値を変更しました。読み込み直して確認してください。');
    var record = { month: month, values: values, source: source, updatedAt: new Date().toISOString() };
    props.setProperty(key, JSON.stringify(record));
    return { ok: true, message: month + 'から使う統計情報を保存しました。スライド・台本を作り直すと反映されます。',
      settings: getBniStatsSettings(month) };
  } catch (e) {
    return { ok: false, message: String(e && e.message ? e.message : e) };
  } finally { if (locked) lock.releaseLock(); }
}

function bniStatsText_(text) { return String(text || '').replace(/[\s　]/g, '').toLowerCase(); }
function bniStatsScope_(text) {
  var t = bniStatsText_(text), w = /世界|グローバル|world|global|international/.test(t), j = /日本|国内|japan/.test(t);
  return w === j ? '' : (j ? 'japan' : 'world');
}
function bniStatsField_(text, scope) {
  var t = bniStatsText_(text), type = '';
  if (/メンバー数|会員数|^(?:(?:全世界|世界|日本|国内)の?)?メンバー$|members/.test(t)) type = 'members';
  else if (/チャプター数|チャプター$|chapters/.test(t)) type = 'chapters';
  else if (/リファーラル数|リファーラル件数|リファーラル$|referrals/.test(t)) type = 'referrals';
  else if (/売上|ビジネス(?:総額|金額|額)|サンキュー|取引額|revenue|businessgenerated/.test(t)) type = 'revenue';
  else if (/国・地域数|国と地域|活動国|国数|countries/.test(t)) type = 'countries';
  scope = bniStatsScope_(text) || scope;
  if (type === 'countries' && !scope) scope = 'world';
  return BNI_STATS_FIELDS_.filter(function(f) { return f.scope === scope && f.type === type; })[0] || null;
}
function bniStatsUnitRe_(type) {
  return { members: '(?:人|名)', chapters: 'チャプター', referrals: '件',
    revenue: '(?:円|米?ドル|USD)', countries: '(?:の)?(?:か国|ヶ国|カ国|ヵ国|国)(?:・地域|と地域)?' }[type];
}
// 数字だけを別のランにしてある雛形でも、数字の色と単位の色を保つ。
function bniStatsFigure_(text, type, standalone) {
  var core = '(?:約|およそ)?[0-9０-９][0-9０-９,，.．兆億万千]*';
  var re = new RegExp(core + '\\s*' + bniStatsUnitRe_(type) + '(?:超|以上)?', 'i'), m = re.exec(text);
  if (!m && standalone) m = new RegExp('^\\s*' + core + '(?:超|以上)?\\s*$').exec(text);
  if (!m) return null;
  return { start: m.index, end: m.index + m[0].length, text: m[0] };
}
function bniStatsReplaceFigure_(p, hit, value) {
  return replaceInParagraph_(p, function() {
    var split = function(v) { return v.match(/^([^0-9０-９]*)([0-9０-９][0-9０-９,，.．兆億万千]*)([\s\S]*)$/); };
    var a = split(hit.text), b = split(value);
    if (!a || !b) return [];
    var out = [], at = hit.start, number = b[2];
    if (a[1]) { out.push({ start: at, end: at + a[1].length, value: b[1] }); at += a[1].length; }
    else number = b[1] + number;
    if (!a[3]) number += b[3];
    out.push({ start: at, end: at + a[2].length, value: number });
    at += a[2].length;
    if (a[3]) out.push({ start: at, end: hit.end, value: b[3] });
    return out;
  });
}

// ページ番号ではなく、項目名と近くの数値を対応させる。世界と日本が別ページでも同じ処理。
function applyBniStats_(parts, record) {
  if (!record) return null;
  var values = record.values || {}, found = {}, updated = {}, paths = [], size = fpSize_(parts);
  slideOrder_(parts).forEach(function(path) {
    var xml = xmlOf_(parts, path), before = xml;
    if (!/世界|日本|BNI|活動する国・地域数|統計情報の適用月|global|japan/i.test(slideText_(xml))) return;
    var shapes = riShapes_(xml).filter(function(s) { return s.tag === 'p:sp'; });
    var pageText = shapes.map(function(s) { return s.text; }).join(' ');
    var trusted = shapes.some(function(s) { return s.ay < size.H * 0.3 &&
      /BNIの?(?:世界|日本|グローバル)?(?:目的|概要|実績|統計)|BNI\s*(?:world|global|international|japan)|(?:世界|日本)のBNI/i.test(s.text); });
    var scope = trusted ? (bniStatsScope_(pageText) || (!/日本|国内|japan/i.test(pageText) ? 'world' : '')) : '';
    var labels = [], paragraphs = [];
    shapes.forEach(function(s) {
      s.paras.forEach(function(text, i) {
        var p = { id: s.id, index: i, text: text, x: s.ax, y: s.ay + s.acy * i / s.paras.length,
          cx: s.acx, cy: s.acy / s.paras.length, width: s.cx };
        paragraphs.push(p);
        // グローバル/日本の統計と分からないページでは、一般の「メンバー数」等を変更しない。
        var f = bniStatsField_(text, scope);
        if (f && (bniStatsScope_(text) || trusted)) labels.push({ p: p, field: f });
      });
    });
    // {{世界メンバー数}} などの差し込み口があるテンプレートは最優先。
    BNI_STATS_FIELDS_.forEach(function(f) {
      var token = '{{' + f.label + '}}', alias = f.key === 'worldCountries' ? '{{世界の国・地域数}}' : token;
      if (slideText_(xml).indexOf(token) < 0 && slideText_(xml).indexOf(alias) < 0) return;
      found[f.key] = true;
      if (!values[f.key]) return;
      var map = {}; map[f.label] = values[f.key]; map[alias.slice(2, -2)] = values[f.key];
      xml = replaceTokensInXml_(xml, map); updated[f.key] = true;
    });
    xml = replaceTokensInXml_(xml, { '統計情報の適用月': record.month });
    var edges = [];
    labels.forEach(function(l, li) {
      if (l.p.text.indexOf('{{') >= 0) return;
      paragraphs.forEach(function(p, pi) {
        var same = p === l.p, hit = bniStatsFigure_(p.text, l.field.type, !same);
        if (!hit || (!same && bniStatsText_(hit.text) !== bniStatsText_(p.text))) return;
        var dx = Math.abs(p.x + p.cx / 2 - l.p.x - l.p.cx / 2), dy = p.y - l.p.y - l.p.cy;
        var below = dx < Math.max(p.cx, l.p.cx) * 0.65 && dy >= -l.p.cy && dy < size.H * 0.18;
        var right = Math.abs(p.y + p.cy / 2 - l.p.y - l.p.cy / 2) < Math.max(p.cy, l.p.cy)
          && p.x >= l.p.x && p.x - l.p.x - l.p.cx < size.W * 0.22;
        if (!same && !below && !right) return;
        var score = same ? -1 : dx / size.W + Math.abs(dy) / size.H;
        edges.push({ label: li, para: pi, hit: hit, score: score });
      });
    });
    edges.sort(function(a, b) { return a.score - b.score; });
    var usedLabels = {}, usedParas = {}, changes = [];
    edges.forEach(function(e) {
      if (usedLabels[e.label] || usedParas[e.para]) return;
      if (edges.some(function(other) { return other.label === e.label && other.para !== e.para
          && Math.abs(other.score - e.score) < 0.005; })) return; // 同程度の候補が複数あれば推測で書き換えない
      var l = labels[e.label], p = paragraphs[e.para];
      usedLabels[e.label] = true; usedParas[e.para] = true; found[l.field.key] = true;
      if (values[l.field.key]) changes.push({ p: p, hit: e.hit, key: l.field.key });
    });
    changes.forEach(function(c) {
      xml = fpMapParas_(xml, c.p.id, function(p, text, i) {
        return i === c.p.index ? bniStatsReplaceFigure_(p, c.hit, values[c.key]) : p;
      });
      xml = fpFitPara_(xml, c.p.id, c.p.index, c.p.width, 1);
      updated[c.key] = true;
    });
    if (xml !== before) { putXml_(parts, path, xml); paths.push(path); }
  });
  var missing = BNI_STATS_FIELDS_.filter(function(f) { return values[f.key] && !found[f.key]; }).map(function(f) { return f.label; });
  var msg = 'BNI統計情報: ' + record.month + 'の数値を使用（' + paths.length + 'ページ更新）。';
  if (missing.length) msg += '\n※ 数値欄が見つからない項目: ' + missing.join('、')
    + '。画像だけのページは書き換えられません。文字の見出しと数値、または {{項目名}} の差し込み口を使ってください。';
  return { month: record.month, pages: paths.length, paths: paths, updated: Object.keys(updated), missing: missing, message: msg };
}

// 保存済みの古い台本も、統計の行だけ数値を更新する（ひな形そのものは保存し直さない）。
function bniStatsLegacyField_(slide) {
  var t = bniStatsText_(slide);
  if (/^(?:全)?世界(?:で)?[0-9０-９,，.．]+万?人(?:以上|超)?$/.test(t)) return BNI_STATS_FIELDS_[0];
  if (/^(?:全)?世界(?:で)?[0-9０-９,，.．]+$/.test(t)) return BNI_STATS_FIELDS_[1];
  if (/^[0-9０-９,，]+(?:か国|ヶ国)$/.test(t)) return BNI_STATS_FIELDS_[6];
  return null;
}
function bniStatsTalkSlide_(slide, record) {
  var f = bniStatsLegacyField_(slide);
  return record && f && record.values[f.key] ? f.label : slide;
}
function bniStatsTalkText_(text, slide, record) {
  if (!record) return text;
  var f = bniStatsField_(slide, '') || bniStatsLegacyField_(slide);
  var keys = f ? [f.key] : (/BNIの数値|BNI統計/.test(slide) ? ['worldReferrals', 'worldRevenue'] : []);
  BNI_STATS_FIELDS_.forEach(function(field) {
    if (keys.indexOf(field.key) < 0 || !(record.values || {})[field.key]) return;
    var hit = bniStatsFigure_(text, field.type, false);
    // 以前の台本の「11,000を超えるチャプター」という書き方にも対応する。
    if (!hit && field.type === 'chapters') {
      var m = /[0-9０-９][0-9０-９,，.．]*を超えるチャプター/.exec(text);
      if (m) hit = { start: m.index, end: m.index + m[0].length, text: m[0] };
    }
    if (hit) {
      var value = record.values[field.key], tail = text.substring(hit.end);
      if (/^を超え/.test(tail) && /超$/.test(value)) value = value.slice(0, -1);
      text = text.substring(0, hit.start) + value + tail;
    }
  });
  return text;
}
