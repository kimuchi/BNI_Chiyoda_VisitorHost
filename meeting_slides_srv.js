// === 定例会スライドの自動更新 ===
// メンバー名簿の入会日・更新日・更新期限日から、スライドに載せる内容を自動算出する。
// 手入力していた「新メンバー」「更新メンバー」「更新状況一覧(90/60/30日)」が不要になる。

// 前半と後半は入口を分けてある。どちらも開いたらすぐ、そのテンプレートと
// ルーティンチェックシートを読み込む（読み込み中は画面に「読み込み中…」を出す）。
function openMeetingFirstDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('slides_meeting_first').evaluate().setWidth(900).setHeight(760),
    '定例会スライド（前半）');
}
function openMeetingSecondDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('slides_meeting_second').evaluate().setWidth(900).setHeight(760),
    '定例会スライド（後半）');
}

// テンプレートを開いて調べた結果を、テンプレートが変わるまで覚えておく。
// 数十MBのファイルを開くので、画面を開くたびに読むと毎回待たされるため。
// ファイルを差し替える（別のファイルを登録する・Drive上で更新する）と読み直す。
function templateMemo_(kind, name, compute) {
  var file = getBigTemplateFile_(kind);
  var stamp = file.getId() + ':' + file.getLastUpdated().getTime();
  var props = PropertiesService.getScriptProperties(), key = 'BNI_TPL_MEMO_' + name;
  try {
    var memo = JSON.parse(props.getProperty(key) || 'null');
    if (memo && memo.stamp === stamp) return memo.value;
  } catch (e) {}
  var value = compute(unzipToMap_(file.getBlob()));
  try { props.setProperty(key, JSON.stringify({ stamp: stamp, value: value })); } catch (e) {}   // 大きすぎるときは覚えない
  return value;
}

// 文字列/日付値 → Date（時刻を0時に丸める）。解釈できなければ null
function parseDate_(v) {
  if (!v) return null;
  var d;
  if (Object.prototype.toString.call(v) === '[object Date]') d = new Date(v.getTime());
  else {
    var s = String(v).trim().replace(/[年月]/g, '/').replace(/日/g, '');
    if (!s) return null;
    d = new Date(s);
  }
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

function fmtDate_(d) { return d ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd') : ''; }
function daysBetween_(a, b) { return Math.round((b.getTime() - a.getTime()) / 86400000); }

// 氏名の配列 → 「山田さん、佐藤さん」。該当なしは「該当者なし」
function joinNames_(names) {
  if (!names || !names.length) return '該当者なし';
  var out = [];
  for (var i = 0; i < names.length; i++) out.push(names[i] + 'さん');
  return out.join('、');
}

// 開催日を基準に、新メンバー・更新メンバー・更新期限の一覧を算出する
// === 更新状況のチェック ===
// {氏名: {status:'done'|'leaving', exp:'2026/10/01'}}
// 更新期限日とセットで覚えるのがポイント。「更新済み」の印は、その期限に対するもの。
// 会費レポートを取り込んで期限が伸びれば印は自動で外れ、次の期限でまた案内に出る。
// 「更新しない」は期限に関係なく出さない（退会される方なので案内の対象外）。
var RENEWAL_MARK_KEY_ = 'BNI_RENEWAL_MARKS';

function getRenewalMarks_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(RENEWAL_MARK_KEY_);
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    console.warn('[RENEW] 印の読み込みに失敗: ' + (e && e.message ? e.message : e));
    return {};
  }
}

// status: 'done' | 'leaving' | '' （空で印を外す）
function setRenewalMark(name, status, exp) {
  try {
    if (!name) return { ok: false, message: '氏名がありません。' };
    var marks = getRenewalMarks_(), key = normName_(name);
    if (status === 'done' || status === 'leaving') marks[key] = { status: status, exp: String(exp || '') };
    else delete marks[key];
    PropertiesService.getScriptProperties().setProperty(RENEWAL_MARK_KEY_, JSON.stringify(marks));
    return { ok: true };
  } catch (e) {
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

function computeRenewalLists(meetingDateVal) {
  try {
    var base = parseDate_(meetingDateVal);
    if (!base) return { ok: false, message: '開催日を解釈できませんでした。' };
    var members = (getMemberMaster().members || []);
    var newMembers = [], renewMembers = [], d90 = [], d60 = [], d30 = [], overdue = [], noDate = [];
    var marks = getRenewalMarks_(), done = [], leaving = [];

    for (var i = 0; i < members.length; i++) {
      var m = members[i];
      var join = parseDate_(m.joinDate), renew = parseDate_(m.renewDate), exp = parseDate_(m.expireDate);

      // 新メンバー: 入会日が開催日と同じ年月（＝今月入会）
      if (join && join.getFullYear() === base.getFullYear() && join.getMonth() === base.getMonth()) {
        newMembers.push({ name: m.name, company: m.company, title: m.title, cat: m.cat, date: fmtDate_(join) });
      }
      // 更新メンバー: 更新日が開催日と同じ年月。継続年数は入会日から算出
      if (renew && renew.getFullYear() === base.getFullYear() && renew.getMonth() === base.getMonth()) {
        var years = 0;
        if (join) {
          years = renew.getFullYear() - join.getFullYear();
          if (renew.getMonth() < join.getMonth() ||
             (renew.getMonth() === join.getMonth() && renew.getDate() < join.getDate())) years--;
          if (years < 1) years = 1;
        }
        renewMembers.push({ name: m.name, company: m.company, years: years, date: fmtDate_(renew) });
      }
      // 更新状況: 更新期限日までの残り日数で区分
      if (!exp) { if (m.name) noDate.push(m.name); continue; }
      var left = daysBetween_(base, exp), expStr = fmtDate_(exp);
      // チェック済みの人は案内から外す
      var mk = marks[normName_(m.name)];
      if (mk && mk.status === 'leaving') { leaving.push({ name: m.name, date: expStr }); continue; }
      if (mk && mk.status === 'done' && mk.exp === expStr) { done.push({ name: m.name, date: expStr }); continue; }
      var rec = { name: m.name, date: expStr, left: left };
      if (left < 0) overdue.push(rec);
      else if (left <= 30) d30.push(rec);
      else if (left <= 60) d60.push(rec);
      else if (left <= 90) d90.push(rec);
    }

    var byLeft = function (a, b) { return a.left - b.left; };
    d30.sort(byLeft); d60.sort(byLeft); d90.sort(byLeft); overdue.sort(byLeft);
    var nameOf = function (a) { return a.name; };

    return { ok: true, meetingDate: fmtDate_(base),
      newMembers: newMembers, renewMembers: renewMembers,
      d90: d90, d60: d60, d30: d30, overdue: overdue, noDate: noDate,
      done: done, leaving: leaving,
      text: {
        d90: joinNames_(d90.map(nameOf)),
        d60: joinNames_(d60.map(nameOf)),
        d30: joinNames_(d30.map(nameOf)),
        overdue: joinNames_(overdue.map(nameOf)),
        newMembers: joinNames_(newMembers.map(nameOf)),
        renewMembers: renewMembers.length
          ? renewMembers.map(function (r) { return r.name + 'さん' + (r.years ? '（' + r.years + '年）' : ''); }).join('、')
          : '該当者なし'
      }
    };
  } catch (e) {
    console.error('[MEETING] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '算出に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// ダイアログの初期表示。各項目の初期値をできる限りシートから埋める
function getMeetingSlideContext() {
  try {
    var candidates = getMeetingCandidates();          // 既存関数（開催日・第N回）
    var first = candidates.length ? candidates[0] : null;
    var lists = first ? computeRenewalLists(first.dateValue) : { ok: false };
    var stats = getMeetingStats_();
    var tpl = getBigTemplateStatus().templates, ready = {};
    for (var i = 0; i < tpl.length; i++) ready[tpl[i].kind] = tpl[i].registered;
    var members = getMemberMaster().members || [];
    // その日のコアバリュー・開催回は、ルーティンチェックシートに書いてある
    var routine = first ? getRoutineInfo(first.dateValue) : null;
    return { ok: true, meetings: candidates, defaultMeeting: first,
             lists: lists.ok ? lists : null, stats: stats, templates: ready,
             routine: routine,
             coreValues: CORE_VALUES_.map(function (c) { return c.label; }),
             memberCount: members.length,
             memberNames: members.map(function (m) { return m.name; }),
             members: members.map(function (m) {
               return { no: m.no, name: m.name, company: m.company, title: m.title,
                        hasPhoto: !!findPhotoIdForName_(m.name) };
             }) };
  } catch (e) {
    console.error('[MEETING] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '初期表示の取得に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 実績数値は前回値を初期値にする（毎回ゼロから入力しなくて済む）
var MEETING_STATS_KEY_ = 'BNI_MEETING_STATS';
function getMeetingStats_() {
  var json = PropertiesService.getScriptProperties().getProperty(MEETING_STATS_KEY_);
  if (!json) return {};
  try { return JSON.parse(json); } catch (e) { return {}; }
}
function saveMeetingStats(data) {
  try {
    PropertiesService.getScriptProperties().setProperty(MEETING_STATS_KEY_, JSON.stringify(data || {}));
    return { ok: true, message: '実績数値を保存しました。次回はこの値が初期表示されます。' };
  } catch (e) {
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// --- 差し込み口が無いページの「第○回」「○年○月○日」---
// テンプレートによっては、表紙などに {{開催回}} を置かず、そのまま
// 「第526回 2026年07月22日」と書いてあるページがある。
// そういうページも更新できるよう、形で見つけて書き換える。
function meetingPatternRules_(no, d) {
  var rules = [];
  if (no) {
    rules.push({ re: /第\s*\d{1,5}\s*回/, value: function () { return '第' + no + '回'; } });
  }
  if (d) {
    var y = d.getFullYear(), mo = d.getMonth() + 1, da = d.getDate();
    var pad = function (n, wide) { return wide ? ('0' + n).slice(-2) : String(n); };
    // 「2026年07月22日」… 元の桁数（0埋めの有無）はそのまま引き継ぐ
    rules.push({ re: /(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日/, value: function (m) {
      return y + '年' + pad(mo, m[2].length === 2) + '月' + pad(da, m[3].length === 2) + '日';
    } });
    // 「2026/07/22」「2026-07-22」
    rules.push({ re: /(\d{4})([\/\-])(\d{1,2})\2(\d{1,2})/, value: function (m) {
      return y + m[2] + pad(mo, m[3].length === 2) + m[2] + pad(da, m[4].length === 2);
    } });
  }
  return rules;
}

// --- コアバリューのページ ---
// 前半スライドには7つのコアバリューのページが入っていて、その週の1枚だけを表示する。
// どのページがどのコアバリューかはテンプレート側の文字から判断する。
// ただし「責任」「伝統」のような語はふつうの文にも出るため、
// 「1つだけ当てはまるページ」が3種類以上そろったときにだけ、まとめて切り替える。
function coreValuesInSlide_(xml, useJa) {
  var t = '', m, re = /<a:t(?=[\s>])[^>]*>([\s\S]*?)<\/a:t>/g;
  while ((m = re.exec(xml)) !== null) t += unescapeXml_(m[1]) + ' ';
  var en = t.toLowerCase().replace(/[^a-z]/g, ''), ja = t.replace(/[\s　]/g, ''), found = [];
  for (var i = 0; i < CORE_VALUES_.length; i++) {
    if (en.indexOf(CORE_VALUES_[i].en) >= 0) { found.push(CORE_VALUES_[i]); continue; }
    if (useJa && ja.indexOf(CORE_VALUES_[i].ja) >= 0) found.push(CORE_VALUES_[i]);
  }
  return found;
}

function setSlideShow_(xml, show) {
  return xml.replace(/<p:sld(?=[\s>])([^>]*)>/, function (all, attrs) {
    return '<p:sld' + attrs.replace(/\sshow="[^"]*"/g, '') + (show ? '' : ' show="0"') + '>';
  });
}

function applyCoreValue_(parts, wanted) {
  var w = coreValueOf_(wanted);
  if (!w) return null;
  // まず英語だけで探す。「責任」「伝統」のような語はふつうの文にも出てくるため、
  // 英語で足りているならそちらを使う方が誤検出が少ない。
  var scan = function (useJa) {
    var cand = [], kinds = {}, path;
    for (path in parts) {
      if (!/^ppt\/slides\/slide\d+\.xml$/.test(path)) continue;
      var xml = xmlOf_(parts, path);
      if (!xml) continue;
      var found = coreValuesInSlide_(xml, useJa);
      if (found.length !== 1) continue;
      cand.push({ path: path, xml: xml, cv: found[0] });
      kinds[found[0].label] = true;
    }
    var n = 0;
    for (var k in kinds) n++;
    return { cand: cand, kinds: n };
  };
  var res = scan(false);
  if (res.kinds < 3) res = scan(true);
  var cand = res.cand, kindCount = res.kinds;
  if (kindCount < 3) {
    return { value: w.label, shown: 0, hidden: 0, kinds: kindCount,
             message: 'コアバリューのページを見分けられませんでした（' + kindCount + '種類）。表示の切り替えは行っていません。' };
  }
  var shown = [], hidden = [];
  for (var i = 0; i < cand.length; i++) {
    var on = (cand[i].cv.label === w.label);
    var out = setSlideShow_(cand[i].xml, on);
    if (out !== cand[i].xml) putXml_(parts, cand[i].path, out);
    (on ? shown : hidden).push(cand[i].path.replace(/^.*slide(\d+)\.xml$/, '$1'));
  }
  return { value: w.label, shown: shown.length, hidden: hidden.length, kinds: kindCount,
           message: 'コアバリュー「' + w.label + '」のページ' + shown.length + '枚を表示、他'
                    + hidden.length + '枚を非表示にしました。' };
}

// --- 一般規定のページ ---
// 12枚あるうちの1枚だけを表示にする。何番かはルーティンチェックシートに書いてある。
// ページの見分けには、BNI公式スライドがノートに持っている通し番号（J-04164 …）を使う。
// スライドの位置や文言が変わっても、この番号は付いて回る。
var GENERAL_POLICY_NOTE_BASE_ = 4163;    // J-04164 が「一般規定1番」
var GENERAL_POLICY_MAX_ = 12;

function slideNoteText_(parts, slidePath) {
  var m = slidePath.match(/^ppt\/slides\/slide(\d+)\.xml$/);
  if (!m) return '';
  var rels = xmlOf_(parts, 'ppt/slides/_rels/slide' + m[1] + '.xml.rels');
  if (!rels) return '';
  var nm = rels.match(/Target="\.\.\/notesSlides\/(notesSlide\d+\.xml)"/);
  if (!nm) return '';
  var note = xmlOf_(parts, 'ppt/notesSlides/' + nm[1]);
  if (!note) return '';
  var t = '', mm, re = /<a:t(?=[\s>])[^>]*>([\s\S]*?)<\/a:t>/g;
  while ((mm = re.exec(note)) !== null) t += unescapeXml_(mm[1]);
  return t;
}

function applyGeneralPolicy_(parts, no) {
  if (!no) return null;
  var byNo = {}, path, found = 0;
  for (path in parts) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(path)) continue;
    var id = slideNoteText_(parts, path).match(/J-0(\d{4})/);
    if (!id) continue;
    var n = parseInt(id[1], 10) - GENERAL_POLICY_NOTE_BASE_;
    if (n >= 1 && n <= GENERAL_POLICY_MAX_ && !byNo[n]) { byNo[n] = path; found++; }
  }
  if (found < 6) {
    return { no: no, found: found,
             message: '一般規定のページを見分けられませんでした（' + found + '枚）。表示の切り替えは行っていません。' };
  }
  var shown = 0, hidden = 0;
  for (var k in byNo) {
    var on = (parseInt(k, 10) === no);
    var xml = xmlOf_(parts, byNo[k]), out = setSlideShow_(xml, on);
    if (out !== xml) putXml_(parts, byNo[k], out);
    if (on) shown++; else hidden++;
  }
  return { no: no, found: found, shown: shown, hidden: hidden,
           message: shown ? ('一般規定' + no + '番のページを表示、他' + hidden + '枚を非表示にしました。')
                          : ('一般規定' + no + '番のページが見つかりませんでした（' + found + '枚中）。') };
}

// --- 2人が並ぶページの写真（メインプレゼン・推薦のことば・抽選コーナー）---
// 前半のメインプレゼン、後半の推薦のことば・抽選コーナーは、どれも
// 「左右に写真とお名前が並ぶ」同じ作りなので、1つの処理でまかなう。
// {{○○1氏名}} が載っているページを探し、それぞれの方の写真の枠を決めて差し替える。
//
// 写真の枠は「その方の氏名の差し込み口の真上にある枠」で決める。大きさだけで選ぶと、
// 抽選コーナーの中央下にある飾りの画像（写真と同じくらいの大きさ）を枠と取り違えるため。
// 名前が空・写真が無い方の枠は、テンプレートの元の写真が残らないよう枠ごと消す。
function applyTwoPersonPhotos_(parts, prefix, names, cache) {
  if (!names || !names.length) return null;
  var path = null, p;
  for (p in parts) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(p)) continue;
    var x = xmlOf_(parts, p);
    if (x && x.indexOf(prefix + '1氏名') >= 0) { path = p; break; }
  }
  if (!path) return { message: '' };
  var r = setTwoPersonPhotos_(parts, path, prefix, names, cache);
  var msg = r.done.length ? (prefix + 'の写真を差し替えました: ' + r.done.join('、')) : '';
  if (r.missing.length) msg += (msg ? '／' : '') + '写真が見つからない方（写真なし）: ' + r.missing.join('、');
  if (r.noFrame) msg = prefix + 'のページで写真の枠が見つかりませんでした。';
  return { message: msg, replaced: r.done.length };
}

// 1枚のページの写真を差し替える（差し込み口 {{○○1氏名}} がまだ残っている状態で呼ぶこと）。
// 戻り値 { done: [差し替えた方], missing: [写真が無かった方], noFrame: 枠が無かったか }
function setTwoPersonPhotos_(parts, path, prefix, names, cache) {
  var xml = xmlOf_(parts, path), res = { done: [], missing: [], noFrame: false };
  var prs = xmlOf_(parts, 'ppt/presentation.xml') || '';
  var sz = prs.match(/<p:sldSz\s+cx="(\d+)"/);
  var slideW = sz ? parseInt(sz[1], 10) : 12192000;

  // 写真の枠の候補（背景いっぱいの画像と、小さな飾りは除く）
  var pics = [], ranges = findTagRanges_(xml, 'p:pic');
  for (var i = 0; i < ranges.length; i++) {
    var seg = xml.substring(ranges[i].start, ranges[i].end);
    var id = (seg.match(/<p:cNvPr[^>]*\sid="(\d+)"/) || [])[1];
    var off = seg.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>/);
    var ext = seg.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
    if (!id || !off || !ext || !/<a:blip[^>]*r:embed=/.test(seg)) continue;
    var cx = parseInt(ext[1], 10), cy = parseInt(ext[2], 10);
    if (cy < 1500000) continue;                       // 小さな飾りは対象外
    if (cx > slideW * 0.4) continue;                  // 背景いっぱいの画像は対象外
    pics.push({ id: id, x: parseInt(off[1], 10), y: parseInt(off[2], 10), cx: cx, cy: cy,
                center: parseInt(off[1], 10) + cx / 2, area: cx * cy });
  }
  if (!pics.length) { res.noFrame = true; return res; }
  var slots = pickPhotoFrames_(xml, prefix, Math.min(names.length, 2), pics);

  var rp = relsPathOf_(path), rels = xmlOf_(parts, rp);
  cache = cache || { by: {}, seq: 0 };
  for (var k = 0; k < slots.length; k++) {
    var pic = slots[k];
    if (!pic) continue;
    var photo = names[k] ? mpAddPhoto_(parts, cache, names[k]) : null;
    if (!photo) {
      xml = removeShape_(xml, pic.id);
      if (names[k]) res.missing.push(names[k]);
      continue;
    }
    var box = readShapeGeomEmu_(xml, pic.id);
    if (box && photo.width && photo.height) {
      xml = setSrcRectInPic_(xml, pic.id, coverCrop_(photo.width, photo.height, box.cx, box.cy));
    }
    var set = setPicImage_(xml, rels, pic.id, '../media/' + photo.path.replace('ppt/media/', ''));
    xml = set.xml; rels = set.rels;
    res.done.push(names[k]);
  }
  putXml_(parts, path, xml);
  if (rels) putXml_(parts, rp, rels);
  return res;
}

// お2人が並ぶページ（メインプレゼン・推薦のことば・抽選）の氏名・会社名・カテゴリーの文字箱は、
// 幅が決まっていて、すぐ下に次の段がある。長いお名前や会社名が折り返すと下の段に重なるので、
// 収まらないときだけ文字を小さくする（氏名・会社名は1行、カテゴリーは2行まで）。
// 文字の幅は全角1文字＝1em・半角スペース＝0.3em・ほかの半角＝0.55em で見積もり、
// 太字や代わりのフォント（Meiryo UI が無い環境）でも収まるよう、幅は5%ほどゆとりをみる。
// 差し込み口 {{○○1会社名}} などがまだ残っている状態で呼ぶこと。
var TWO_PERSON_FIT_ = [
  { field: '氏名', lines: 1, min: 20 },
  { field: '会社名', lines: 1, min: 8 },           // 支社・営業所まで入る長い会社名もあるので、小さめまで許す
  { field: 'カテゴリー', lines: 2, min: 9 }
];
function fitTwoPersonText_(xml, prefix, values) {
  for (var k = 1; k <= 2; k++) {
    for (var f = 0; f < TWO_PERSON_FIT_.length; f++) {
      var t = TWO_PERSON_FIT_[f], key = prefix + k + t.field;
      if (values[key]) xml = fitTokenBox_(xml, '{{' + key + '}}', String(values[key]), t.lines, t.min);
    }
  }
  return xml;
}
function fitTwoPersonPage_(parts, prefix, values) {
  for (var path in parts) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(path)) continue;
    var xml = xmlOf_(parts, path);
    if (!xml || xml.indexOf(prefix + '1氏名') < 0) continue;
    var out = fitTwoPersonText_(xml, prefix, values);
    if (out !== xml) putXml_(parts, path, out);
  }
}
function textBoxEm_(s) {
  var em = 0;
  for (var i = 0; i < s.length; i++) {
    var c = s.charCodeAt(i);
    em += c === 32 ? 0.3 : (c < 128 ? 0.55 : 1);
  }
  return em;
}
function fitTokenBox_(xml, token, text, maxLines, minPt) {
  var sps = findTagRanges_(xml, 'p:sp');
  for (var i = 0; i < sps.length; i++) {
    var seg = xml.substring(sps[i].start, sps[i].end);
    if (slideText_(seg).indexOf(token) < 0) continue;
    var ext = seg.match(/<a:ext cx="(\d+)" cy="\d+"\s*\/>/);
    if (!ext) return xml;
    var bp = (seg.match(/<a:bodyPr\b[^>]*>/) || [''])[0];
    var ins = function (k) {                                   // 文字箱の左右の余白（既定は0.1インチ）
      var m = bp.match(new RegExp('\\s' + k + '="(\\d+)"'));
      return m ? parseInt(m[1], 10) : 91440;
    };
    var widthPt = (parseInt(ext[1], 10) - ins('lIns') - ins('rIns')) / 12700 * 0.95;
    var pt = parseInt((seg.match(/<a:rPr\b[^>]*\ssz="(\d+)"/) || [0, 1800])[1], 10) / 100;
    var em = textBoxEm_(text);
    if (em * pt <= widthPt * maxLines) return xml;
    var size = Math.max(minPt || 10, Math.floor(widthPt * maxLines / em * 2) / 2);
    seg = seg.replace(/(<a:(?:rPr|endParaRPr)\b[^>]*?\ssz=")\d+(")/g, '$1' + Math.round(size * 100) + '$2');
    return xml.substring(0, sps[i].start) + seg + xml.substring(sps[i].end);
  }
  return xml;
}

// --- 書記兼会計による報告（更新状況一覧）---
// 後半の「書記兼会計による報告」のページは、「90日以内に更新を迎えるメンバー」などの見出しと
// お名前の2段の表が4つ並ぶ作り。差し込み口を置かなくても、見出しの文字で表を見分けて、
// 下の段を画面の「更新状況一覧」の値に入れ替える（テンプレートを作り直さずに済むように）。
// お名前は「、」の区切りで改行し、多いときは文字を小さくして元の2行ぶんの高さに収める（下の表に重ならないように）。
var RENEWAL_TABLES_ = [
  { key: '更新90', label: '90日以内', re: /90日以内/ },
  { key: '更新60', label: '60日以内', re: /60日以内/ },
  { key: '更新30', label: '30日以内', re: /30日以内/ },
  { key: '更新超過', label: '期限切れ', re: /期限切れ|期限超過/ }
];
function applyRenewalStatus_(parts, map) {
  var want = false, i;
  for (i = 0; i < RENEWAL_TABLES_.length; i++) if (map && map[RENEWAL_TABLES_[i].key] !== undefined) want = true;
  if (!want) return null;
  var done = [], tight = [], path;
  for (path in parts) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(path)) continue;
    var xml = xmlOf_(parts, path);
    if (!xml || slideText_(xml).indexOf('更新を迎えるメンバー') < 0) continue;
    var frames = findTagRanges_(xml, 'p:graphicFrame'), todo = [];
    for (i = 0; i < frames.length; i++) {
      var seg = xml.substring(frames[i].start, frames[i].end);
      var id = (seg.match(/<p:cNvPr[^>]*\sid="(\d+)"/) || [])[1];
      var trs = findTagRanges_(seg, 'a:tr');
      if (!id || trs.length < 2) continue;
      var head = slideText_(seg.substring(trs[0].start, trs[0].end));
      for (var k = 0; k < RENEWAL_TABLES_.length; k++) {
        var t = RENEWAL_TABLES_[k];
        if (t.re.test(head) && map[t.key] !== undefined) { todo.push({ id: id, t: t }); break; }
      }
    }
    for (i = 0; i < todo.length; i++) {
      var text = String(map[todo[i].t.key] == null ? '' : map[todo[i].t.key]).trim() || '該当者なし';
      var fit = setNamesInTableCell_(xml, todo[i].id, 1, 0, text, 2);
      xml = fit.xml;
      if (!fit.fits) tight.push(todo[i].t.label);
      done.push(todo[i].t.label);
    }
    if (todo.length) putXml_(parts, path, xml);
  }
  if (!done.length) return { message: '' };
  var msg = '書記兼会計による報告（更新状況）を書き換えました: ' + done.join('・');
  if (tight.length) msg += '\n更新状況の「' + tight.join('」「') + '」はお名前が多く、文字を小さくしても枠に収まりきりません。'
    + '下の表に重なるときは、画面の欄を短くしてください（名字だけにするなど）。';
  return { message: msg, done: done.length, tight: tight };
}

// 表のセルに、お名前の一覧（「○○ ○○さん、○○ ○○さん、…」）を入れる。
// 元の文字の大きさで maxLines 行ぶんの高さに収まるよう、「、」の区切りで改行し（お名前の途中で
// 折り返さないように）、それでも足りなければ文字を小さくする（小さくすると1行の高さも縮むので、
// 3行になっても同じ高さに収まることがある）。元の大きさより大きくはせず、10pt より小さくはしない
// （それでも収まらなければ fits: false）。文字の幅は全角1文字＝1em・半角＝0.5em で見積もる。
// 戻り値 { xml, pt, lines, fits }
function setNamesInTableCell_(xml, frameId, rowIdx, colIdx, text, maxLines) {
  var res = { xml: xml, pt: 0, lines: [text], fits: true };
  var r = findShapeRange_(xml, frameId);
  if (!r) return res;
  var cols = xml.substring(r.start, r.end).match(/<a:gridCol w="\d+"/g) || [];
  if (!cols[colIdx]) return res;
  var colW = parseInt(cols[colIdx].replace(/\D+/g, ''), 10);
  res.xml = mapTableCell_(xml, frameId, rowIdx, colIdx, function (tc) {
    var pt = parseInt((tc.match(/<a:rPr\b[^>]*\ssz="(\d+)"/) || [0, 1800])[1], 10) / 100;
    var tcPr = (tc.match(/<a:tcPr\b[^>]*>/) || [''])[0];
    var mar = function (k) {                                   // セルの左右の余白（既定は0.1インチ）
      var m = tcPr.match(new RegExp('\\s' + k + '="(\\d+)"'));
      return m ? parseInt(m[1], 10) : 91440;
    };
    // 行末の禁則などで少し余ることがあるので、幅は1割ほどゆとりをみる
    var widthPt = (colW - mar('marL') - mar('marR')) / 12700 * 0.9;
    var layout = function (size) {
      var lines = wrapAtCommas_(text, widthPt / size), h = 0;
      for (var i = 0; i < lines.length; i++) h += Math.max(1, Math.ceil(textEm_(lines[i]) * size / widthPt)) * size;
      return { lines: lines, height: h };
    };
    var size = pt, lay = layout(size);
    while (size > 10 && lay.height > pt * maxLines) { size -= 0.5; lay = layout(size); }
    res.pt = size; res.lines = lay.lines; res.fits = lay.height <= pt * maxLines;
    var tb = findTagRanges_(tc, 'a:txBody');
    if (tb.length) {
      tc = tc.substring(0, tb[0].start) + setTxBodyLines_(tc.substring(tb[0].start, tb[0].end), lay.lines)
         + tc.substring(tb[0].end);
    }
    if (size < pt) tc = tc.replace(/(<a:(?:rPr|endParaRPr)\b[^>]*?\ssz=")\d+(")/g, '$1' + Math.round(size * 100) + '$2');
    return tc;
  });
  return res;
}
function textEm_(s) {
  var em = 0;
  for (var i = 0; i < s.length; i++) em += s.charCodeAt(i) < 128 ? 0.5 : 1;
  return em;
}
// 「、」の区切りで、1行（widthEm 文字ぶん）に入るだけ詰めて行に分ける
function wrapAtCommas_(text, widthEm) {
  var segs = String(text).split('、'), lines = [], cur = '';
  for (var i = 0; i < segs.length; i++) {
    var seg = segs[i] + (i < segs.length - 1 ? '、' : '');
    if (cur && textEm_(cur + seg) > widthEm) { lines.push(cur); cur = seg; }
    else cur += seg;
  }
  if (cur) lines.push(cur);
  return lines.length ? lines : [''];
}

// 表のセル（<a:tc>）を関数で書き換える（行・列は0始まり）
function mapTableCell_(xml, frameId, rowIdx, colIdx, fn) {
  var r = findShapeRange_(xml, frameId);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), tbls = findTagRanges_(seg, 'a:tbl');
  if (!tbls.length) return xml;
  var tbl = seg.substring(tbls[0].start, tbls[0].end), trs = findTagRanges_(tbl, 'a:tr');
  if (rowIdx >= trs.length) return xml;
  var tr = tbl.substring(trs[rowIdx].start, trs[rowIdx].end), tcs = findTagRanges_(tr, 'a:tc');
  if (colIdx >= tcs.length) return xml;
  var tc = fn(tr.substring(tcs[colIdx].start, tcs[colIdx].end));
  tr  = tr.substring(0, tcs[colIdx].start) + tc + tr.substring(tcs[colIdx].end);
  tbl = tbl.substring(0, trs[rowIdx].start) + tr + tbl.substring(trs[rowIdx].end);
  seg = seg.substring(0, tbls[0].start) + tbl + seg.substring(tbls[0].end);
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}

// --- 推薦のことば（何組でも）---
// テンプレートの推薦のことばのページ（{{推薦のことば1氏名}} のあるページ）をひな形に、組の数だけページを作る。
//   定例会中の組          … ひな形のページの場所に続けて並べる（1組目はひな形のページそのもの）
//   アフター・定例会後の組 … 抽選コーナーのページのうしろに並べる
// どのページも、左が推薦する人・右が推薦される人（氏名・会社名・カテゴリー・写真）。
// 定例会中の組が無い日は、ひな形のページを非表示にする（元のお名前と写真は消しておく）。
//   pairs … [{ giver: {name, company, category}, receiver: {…}, after: true/false }]
var RECO_PREFIX_ = '推薦のことば';
function expandRecommendations_(parts, pairs, cache) {
  var model = findSlideWithText_(parts, RECO_PREFIX_ + '1氏名');
  if (!model) {
    return pairs.length ? { message: '推薦のことばのページ（{{推薦のことば1氏名}} のあるページ）が見つかりませんでした。' } : null;
  }
  var during = [], after = [], i;
  for (i = 0; i < pairs.length; i++) (pairs[i].after ? after : during).push(pairs[i]);
  var modelXml = setSlideShow_(xmlOf_(parts, model), true);
  var modelRels = (xmlOf_(parts, relsPathOf_(model)) || '').replace(/<Relationship\b[^>]*notesSlides\/[^>]*\/>/g, '');
  var lottery = findSlideWithText_(parts, '抽選1氏名');
  var entries = slideEntries_(parts), missing = [];

  var fill = function (path, pair) {
    var g = pair.giver || {}, r = pair.receiver || {};
    var ph = setTwoPersonPhotos_(parts, path, RECO_PREFIX_, [g.name || '', r.name || ''], cache);
    missing = missing.concat(ph.missing);
    var vals = {};
    vals[RECO_PREFIX_ + '1氏名'] = g.name || ''; vals[RECO_PREFIX_ + '1会社名'] = g.company || '';
    vals[RECO_PREFIX_ + '1カテゴリー'] = g.category || '';
    vals[RECO_PREFIX_ + '2氏名'] = r.name || ''; vals[RECO_PREFIX_ + '2会社名'] = r.company || '';
    vals[RECO_PREFIX_ + '2カテゴリー'] = r.category || '';
    putXml_(parts, path, replaceTokensInXml_(fitTwoPersonText_(xmlOf_(parts, path), RECO_PREFIX_, vals), vals));
  };
  // 先に、手を付けていないひな形から2組目以降とアフターの組のページを作る
  var clone = function (pair) {
    var add = addSlidePart_(parts, modelXml, modelRels);
    fill(add.path, pair);
    return add;
  };
  var moreDuring = [], afterPages = [];
  for (i = 1; i < during.length; i++) moreDuring.push(clone(during[i]));
  for (i = 0; i < after.length; i++) afterPages.push(clone(after[i]));
  if (during.length) {
    putXml_(parts, model, modelXml);
    fill(model, during[0]);
  } else {
    fill(model, { giver: {}, receiver: {} });
    putXml_(parts, model, setSlideShow_(xmlOf_(parts, model), false));
  }
  // 並び：ひな形のうしろに定例会中の2組目以降、抽選コーナーのうしろにアフター・定例会後の組
  var out = [], placedAfter = false;
  for (i = 0; i < entries.length; i++) {
    out.push(entries[i]);
    if (entries[i].path === model) out = out.concat(moreDuring);
    if (lottery && entries[i].path === lottery) { out = out.concat(afterPages); placedAfter = true; }
  }
  if (!placedAfter) out = out.concat(afterPages);        // 抽選コーナーが無ければ最後に
  setSlideEntries_(parts, out);

  var msg = during.length
    ? ('推薦のことばのページを ' + during.length + '枚作りました（定例会中）')
    : '定例会中の推薦のことばは無いので、そのページは非表示にしました';
  if (after.length) msg += '。アフター・定例会後の ' + after.length + '枚は、'
    + (lottery ? '抽選コーナーのあと' : '最後') + 'に入れました';
  msg += '。';
  if (missing.length) msg += '\n推薦のことばで写真が見つからない方（写真なし）: ' + missing.join('、');
  return { message: msg, during: during.length, after: after.length,
           paths: [model].concat(moreDuring.map(function (x) { return x.path; }), afterPages.map(function (x) { return x.path; })) };
}

// 氏名の差し込み口（{{○○1氏名}}）の文字箱の位置
function tokenBox_(xml, token) {
  var ranges = findTagRanges_(xml, 'p:sp');
  for (var i = 0; i < ranges.length; i++) {
    var seg = xml.substring(ranges[i].start, ranges[i].end);
    if (slideText_(seg).indexOf(token) < 0) continue;
    var off = seg.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>/);
    var ext = seg.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
    if (off && ext) return { y: parseInt(off[2], 10), center: parseInt(off[1], 10) + parseInt(ext[1], 10) / 2 };
  }
  return null;
}

// count 人ぶんの写真の枠を決める（slots[k] が k 人目の枠）。
// 氏名の差し込み口が見つかれば、その真上（左右の位置がいちばん近く、氏名より上）の枠。
// 見つからなければ、大きい順に count 個採って左から割り当てる。
function pickPhotoFrames_(xml, prefix, count, pics) {
  var anchors = [], k, j;
  for (k = 0; k < count; k++) anchors.push(tokenBox_(xml, '{{' + prefix + (k + 1) + '氏名}}'));
  var anchored = true;
  for (k = 0; k < count; k++) if (!anchors[k]) anchored = false;
  if (anchored) {
    var pairs = [];
    for (k = 0; k < count; k++) {
      for (j = 0; j < pics.length; j++) {
        var below = pics[j].y >= anchors[k].y;       // 氏名より下にあるものは写真の枠ではない
        pairs.push({ k: k, j: j, d: Math.abs(pics[j].center - anchors[k].center) + (below ? 1e12 : 0) });
      }
    }
    pairs.sort(function (a, b) { return a.d - b.d; });
    var slots = [], usedK = {}, usedJ = {};
    for (var q = 0; q < pairs.length; q++) {
      if (usedK[pairs[q].k] || usedJ[pairs[q].j]) continue;
      slots[pairs[q].k] = pics[pairs[q].j];
      usedK[pairs[q].k] = usedJ[pairs[q].j] = true;
    }
    return slots;
  }
  var big = pics.slice().sort(function (a, b) { return b.area - a.area; }).slice(0, count);
  big.sort(function (a, b) { return a.center - b.center; });
  return big;
}

// 画像の図形の中身を差し替える。同じ関係IDを別の図形も使っているときは、
// その図形だけ新しい関係IDに付け替える（片方を替えるともう片方も替わってしまうため）。
function setPicImage_(xml, rels, picId, target) {
  var r = findShapeRange_(xml, picId);
  if (!r || !rels) return { xml: xml, rels: rels };
  var seg = xml.substring(r.start, r.end);
  var rid = (seg.match(/<a:blip[^>]*r:embed="([^"]+)"/) || [])[1];
  if (!rid) return { xml: xml, rels: rels };
  if (xml.split('r:embed="' + rid + '"').length - 1 <= 1) {
    return { xml: xml, rels: retargetRel_(rels, rid, target) };
  }
  var max = 0, m, re = /Id="rId(\d+)"/g;
  while ((m = re.exec(rels)) !== null) max = Math.max(max, parseInt(m[1], 10));
  var nid = 'rId' + (max + 1);
  rels = rels.replace('</Relationships>', '<Relationship Id="' + nid + '" Type="'
    + 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + target + '"/></Relationships>');
  seg = seg.replace(new RegExp('(<a:blip[^>]*r:embed=")' + rid + '"'), '$1' + nid + '"');
  return { xml: xml.substring(0, r.start) + seg + xml.substring(r.end), rels: rels };
}

// --- 音楽の差し替えと音量 ---
//
// 音声は <p:pic> に <a:audioFile r:link="…"> と <p14:media r:embed="…"> の2つの関係で
// ぶら下がっている（どちらも同じファイルを指す）。差し替えるときは両方を付け替える。
// 音量は、そのページの <p:timing> の <p:cMediaNode vol="…"> に入っている（0〜100000）。
//
// どのページのどの音かは、図形の名前（曲名がそのまま入っている）で分かるようにしてある。
var AUDIO_EXT_RE_ = /\.(mp3|m4a|wav|wma|aac)$/i;

// スライドに入っている音声を一覧する
function listSlideAudio_(parts, path) {
  var xml = xmlOf_(parts, path);
  if (!xml) return [];
  var n = path.replace(/^.*slide(\d+)\.xml$/, '$1');
  var rels = xmlOf_(parts, 'ppt/slides/_rels/slide' + n + '.xml.rels') || '';
  var out = [], ranges = findTagRanges_(xml, 'p:pic');
  for (var i = 0; i < ranges.length; i++) {
    var seg = xml.substring(ranges[i].start, ranges[i].end);
    var af = seg.match(/<a:audioFile[^>]*r:link="([^"]+)"/)
          || seg.match(/<a:videoFile[^>]*r:link="([^"]+)"/);
    if (!af) continue;
    var isVideo = seg.indexOf('<a:videoFile') >= 0;
    var md = seg.match(/<p14:media[^>]*r:embed="([^"]+)"/);
    var id = (seg.match(/<p:cNvPr id="(\d+)"[^>]*name="([^"]*)"/) || []);
    var tgt = rels.match(new RegExp('Id="' + af[1] + '"[^>]*Target="\\.\\./media/([^"]+)"'));
    // 外部参照（Target="NULL" のまま残っているものなど）は差し替えようがないので飛ばす
    if (!tgt) continue;
    // いまの音量（<p:cMediaNode vol="…">、0〜100000）
    var vol = '';
    var vr = findTagRanges_(xml, 'p:audio').concat(findTagRanges_(xml, 'p:video'));
    for (var v = 0; v < vr.length; v++) {
      var vs = xml.substring(vr[v].start, vr[v].end);
      if (id[1] && vs.indexOf('spid="' + id[1] + '"') < 0) continue;
      var vm = vs.match(/<p:cMediaNode\b[^>]*\svol="(\d+)"/);
      if (vm) { vol = Math.round(parseInt(vm[1], 10) / 1000); break; }
    }
    out.push({ slide: path, spid: id[1] || '', name: id[2] || '', linkRid: af[1],
               embedRid: md ? md[1] : '', file: tgt ? tgt[1] : '', volume: vol, video: isVideo,
               key: path.replace(/^.*\//, '') + '#' + (id[1] || '') });
  }
  return out;
}

function listMeetingAudio_(parts) {
  var out = [], p;
  for (p in parts) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(p)) continue;
    out = out.concat(listSlideAudio_(parts, p));
  }
  out.sort(function (a, b) {
    return parseInt(a.slide.replace(/\D+/g, ''), 10) - parseInt(b.slide.replace(/\D+/g, ''), 10);
  });
  return out;
}

// 音量を変える（0〜100 の％で指定）
function setSlideAudioVolume_(parts, path, spid, percent) {
  var xml = xmlOf_(parts, path);
  if (!xml) return false;
  var vol = Math.max(0, Math.min(100, Number(percent) || 0)) * 1000;
  var changed = false;
  // その音声を指している <p:audio>/<p:video> の <p:cMediaNode> だけを変える
  var tags = ['p:audio', 'p:video'];
  for (var t = 0; t < tags.length; t++) {
    var ranges = findTagRanges_(xml, tags[t]);
    for (var i = ranges.length - 1; i >= 0; i--) {
      var seg = xml.substring(ranges[i].start, ranges[i].end);
      if (spid && seg.indexOf('spid="' + spid + '"') < 0) continue;
      var out = seg.replace(/(<p:cMediaNode\b[^>]*?)\svol="\d+"/, '$1 vol="' + vol + '"');
      if (out === seg) out = seg.replace(/<p:cMediaNode\b/, '<p:cMediaNode vol="' + vol + '"');
      if (out !== seg) { xml = xml.substring(0, ranges[i].start) + out + xml.substring(ranges[i].end); changed = true; }
    }
  }
  if (changed) putXml_(parts, path, xml);
  return changed;
}

// 音楽を差し替える。Driveのファイルをそのまま入れ、関係の指し先を付け替える。
function replaceSlideAudio_(parts, audio, fileId) {
  var file = DriveApp.getFileById(fileId), blob = file.getBlob();
  var name = String(file.getName() || 'music');
  var ext = (name.match(AUDIO_EXT_RE_) || ['.mp3'])[0].toLowerCase();
  var newPath = 'ppt/media/bniMusic' + audio.spid + ext;
  parts[newPath] = blob.setName(newPath);

  var n = audio.slide.replace(/^.*slide(\d+)\.xml$/, '$1');
  var rp = 'ppt/slides/_rels/slide' + n + '.xml.rels', rels = xmlOf_(parts, rp);
  if (!rels) return false;
  var target = '../media/' + newPath.replace('ppt/media/', '');
  rels = retargetRel_(rels, audio.linkRid, target);
  if (audio.embedRid) rels = retargetRel_(rels, audio.embedRid, target);
  putXml_(parts, rp, rels);

  // 図形の名前（曲名）も入れ替えておく。画面の一覧でどの曲か分かるように。
  var xml = xmlOf_(parts, audio.slide);
  xml = xml.replace(new RegExp('(<p:cNvPr id="' + audio.spid + '" name=")[^"]*(")'),
                    '$1' + escapeXml_(name.replace(AUDIO_EXT_RE_, '')) + '$2');
  putXml_(parts, audio.slide, xml);

  // 種類が決まっていないと再生できないので、拡張子の既定を足しておく
  var ct = xmlOf_(parts, '[Content_Types].xml'), e = ext.replace('.', '');
  if (ct && ct.indexOf('Extension="' + e + '"') < 0) {
    var mime = { mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav',
                 wma: 'audio/x-ms-wma', aac: 'audio/aac' }[e] || 'audio/mpeg';
    putXml_(parts, '[Content_Types].xml',
            ct.replace('<Default', '<Default Extension="' + e + '" ContentType="' + mime + '"/><Default'));
  }
  return true;
}

// 画面から指示された差し替え・音量をまとめて適用する。
// music: { 'slide1.xml#3': { fileId: '…', volume: 30 }, … }
function applyMeetingAudio_(parts, music) {
  if (!music) return null;
  var list = listMeetingAudio_(parts), done = [], failed = [];
  for (var i = 0; i < list.length; i++) {
    var a = list[i], key = a.slide.replace(/^.*\//, '') + '#' + a.spid, cfg = music[key];
    if (!cfg) continue;
    try {
      var what = [];
      if (cfg.fileId) { if (replaceSlideAudio_(parts, a, cfg.fileId)) what.push('差し替え'); }
      if (cfg.volume !== undefined && cfg.volume !== '' && cfg.volume !== null) {
        if (setSlideAudioVolume_(parts, a.slide, a.spid, cfg.volume)) what.push('音量' + cfg.volume + '%');
      }
      if (what.length) done.push(key + '（' + what.join('・') + '）');
    } catch (e) {
      failed.push(key + '：' + (e && e.message ? e.message : e));
    }
  }
  var msg = done.length ? ('音楽を設定しました: ' + done.join('、')) : '';
  if (failed.length) msg += (msg ? '\n' : '') + '音楽の設定に失敗: ' + failed.join('、');
  return { message: msg, done: done.length };
}

// 画面用：登録してあるテンプレートの中を調べる。
//   ・入っている音楽・動画（差し替えと音量の一覧に使う）
//   ・リファーラル発表のひな形ページの枠の位置（会社名の収め方を画面で決めるのに使う）
// 後半の画面を開いたときに読む。テンプレートが同じ間は結果を覚えておくので、2回目からはすぐ返る。
function getMeetingTemplateInfo(kind) {
  try {
    if (!BIG_TEMPLATE_KINDS_[kind]) return { ok: false, message: 'スライドの種類が不正です。' };
    var info = templateMemo_(kind, 'info_' + kind, function (parts) {
      var l = listMeetingAudio_(parts);
      for (var i = 0; i < l.length; i++) l[i].slideNo = parseInt(String(l[i].slide).replace(/\D+/g, ''), 10);
      var rb = null, wb = null;
      try { rb = rfLayoutBoxes_(parts, RF_TITLE_); } catch (e) {}
      try { wb = rfLayoutBoxes_(parts, WEEKLY_TITLE_); } catch (e) {}
      return { list: l, referralBoxes: rb, weeklyBoxes: wb };
    });
    var list = info.list || [], boxes = info.referralBoxes, weeklyBoxes = info.weeklyBoxes;
    var msg = list.length ? (list.length + '件の音楽・動画が入っています。')
                          : 'このテンプレートに音楽は入っていません。';
    if (boxes) msg += '／リファーラル発表のひな形が見つかりました。';
    if (weeklyBoxes) msg += '／ウィークリープレゼンのひな形が見つかりました'
      + (weeklyBoxes.hasNext ? '。' : '（NEXT➡の枠はありません）。');
    if (!boxes && !weeklyBoxes) msg += '／発表ページのひな形は見つかりませんでした。';
    return { ok: true, list: list, referralBoxes: boxes, weeklyBoxes: weeklyBoxes, message: msg };
  } catch (e) {
    console.error('[MEETING] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み取りに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 差し替え用の音楽ファイルを、Driveのリンクで覚えておく（画面で選べるように）
var MEETING_MUSIC_KEY_ = 'BNI_MEETING_MUSIC';
function getMeetingMusicFiles() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(MEETING_MUSIC_KEY_);
    var ids = raw ? JSON.parse(raw) : [];
    var out = [];
    for (var i = 0; i < ids.length; i++) {
      try {
        var f = DriveApp.getFileById(ids[i]);
        out.push({ id: ids[i], name: f.getName(), url: f.getUrl(),
                   sizeMB: Math.round((f.getSize() || 0) / 104857.6) / 10 });
      } catch (e) { out.push({ id: ids[i], name: '（開けません）', url: '', sizeMB: 0 }); }
    }
    return { ok: true, files: out };
  } catch (e) {
    return { ok: false, message: '読み込みに失敗しました: ' + (e && e.message ? e.message : e), files: [] };
  }
}

function addMeetingMusicFile(linkOrId) {
  try {
    var id = extractDriveId_(linkOrId);
    if (!id) return { ok: false, message: 'ファイルIDを認識できませんでした。Driveの共有リンクを貼り付けてください。' };
    var f;
    try { f = DriveApp.getFileById(id); }
    catch (e) { return { ok: false, message: 'このIDのファイルを開けませんでした。アクセス権をご確認ください。' }; }
    if (!AUDIO_EXT_RE_.test(f.getName())) {
      return { ok: false, message: '音楽ファイル（mp3 / m4a / wav など）を指定してください（現在: ' + f.getName() + '）。' };
    }
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(MEETING_MUSIC_KEY_), ids = raw ? JSON.parse(raw) : [];
    if (ids.indexOf(id) < 0) ids.push(id);
    props.setProperty(MEETING_MUSIC_KEY_, JSON.stringify(ids));
    return { ok: true, message: '「' + f.getName() + '」を登録しました。', files: getMeetingMusicFiles().files };
  } catch (e) {
    return { ok: false, message: '登録に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

function removeMeetingMusicFile(id) {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(MEETING_MUSIC_KEY_), ids = raw ? JSON.parse(raw) : [];
    var out = [];
    for (var i = 0; i < ids.length; i++) if (ids[i] !== id) out.push(ids[i]);
    props.setProperty(MEETING_MUSIC_KEY_, JSON.stringify(out));
    return { ok: true, message: '登録を外しました。', files: getMeetingMusicFiles().files };
  } catch (e) {
    return { ok: false, message: '削除に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// --- アンバサダー・ディレクターのページ ---
// 前半テンプレートには、メンバー以外の方（アンバサダー・エグゼクティブディレクター）の
// ウィークリープレゼンのページが非表示で入っている。来られる日だけ表示にする。
// どのページかは番号ではなく中身で見分ける：「ウィークリープレゼンテーション」の見出しから
// 「全員終わりましたか？」までの間にある、氏名入りの WEEKLY PRESENTATION のページ。
// メンバープレゼンのページも同じ作りなので、差し込む前に探すこと。
function weeklyGuestPages_(parts) {
  var order = slideOrder_(parts), at = order.indexOf(weeklyAnchor_(parts));
  if (at < 0) return [];
  var sz = (xmlOf_(parts, 'ppt/presentation.xml') || '').match(/<p:sldSz\s+cx="(\d+)"/);
  var slideW = sz ? parseInt(sz[1], 10) : 12192000, out = [];
  for (var i = at + 1; i < order.length; i++) {
    var xml = xmlOf_(parts, order[i]) || '', flat = slideText_(xml).replace(/[\s　]/g, '');
    if (flat.indexOf('終わりましたか') >= 0) break;
    if (flat.toUpperCase().indexOf('WEEKLYPRESENTATION') < 0) continue;
    var sh = presenterShapes_(xml, slideW), name = shapeTextOf_(xml, sh.nameBox);
    // 「苗字　名前」「〇〇　〇〇」のままの下書き（2分30秒のページなど）は人のページではない
    if (!name || /苗字|名前|氏名/.test(name) || /^[〇○◯\s　]+$/.test(name)) continue;
    out.push({ key: order[i].replace(/^.*\//, ''), path: order[i], name: name,
               role: shapeTextOf_(xml, sh.category), hidden: /<p:sld\b[^>]*\sshow="0"/.test(xml) });
  }
  return out;
}

function shapeTextOf_(xml, id) {
  var r = id ? findShapeRange_(xml, id) : null;
  return r ? slideText_(xml.substring(r.start, r.end)).replace(/^[\s　]+|[\s　]+$/g, '') : '';
}

// names … 表示にする方の氏名（画面のチェック）。入っていない方のページは非表示にする。
// auto  … メンバーのページと同じく、カウントダウンが終わったら自動で次へ進めるか
function applyWeeklyGuests_(parts, names, auto) {
  var pages = weeklyGuestPages_(parts), want = {}, shown = [], hidden = [], paths = [], i;
  for (i = 0; i < names.length; i++) want[normName_(names[i])] = true;
  for (i = 0; i < pages.length; i++) {
    var g = pages[i], xml = xmlOf_(parts, g.path), on = !!want[normName_(g.name)];
    xml = setSlideShow_(xml, on);
    if (on) {
      try {
        xml = mpSetCountdown_(xml, WEEKLY_SECONDS_, !auto);
        xml = auto ? mpAutoAdvance_(xml, (WEEKLY_SECONDS_ + 1) * 1000) : mpNoAutoAdvance_(xml);
      } catch (e) {
        console.warn('[MEETING] ' + g.name + 'さんのページのカウントダウンを作り直せませんでした: ' + e.message);
      }
      shown.push(g.name); paths.push(g.path);
    } else {
      hidden.push(g.name);
    }
    putXml_(parts, g.path, xml);
  }
  var msg = '';
  if (!pages.length) {
    if (names.length) msg = '前半スライドにアンバサダー・ディレクターのページが見つかりませんでした。';
  } else if (shown.length) {
    msg = shown.join('さん・') + 'さんのウィークリープレゼンを入れました（メンバーの後／'
        + (auto ? '自動で次へ' : 'クリックで次へ') + '）。'
        + (hidden.length ? hidden.join('さん・') + 'さんのページは非表示です。' : '');
  } else {
    msg = hidden.join('さん・') + 'さんのウィークリープレゼンのページは非表示です。';
  }
  return { message: msg, paths: paths, auto: auto && paths.length > 0, shown: shown, hidden: hidden };
}

// 画面用：前半テンプレートに入っているアンバサダー・ディレクターのページ。
// 27MBほどのファイルを開くので、テンプレートが同じ間は結果を覚えておく。
function getWeeklyGuests() {
  try {
    var guests = templateMemo_('meetingFirst', 'guests', function (parts) {
      return weeklyGuestPages_(parts).map(function (g) {
        return { name: g.name, role: g.role, hidden: g.hidden };
      });
    });
    return { ok: true, guests: guests };
  } catch (e) {
    console.error('[MEETING] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'アンバサダー・ディレクターのページを調べられませんでした: '
             + (e && e.message ? e.message : e), guests: [] };
  }
}

// --- メンバープレゼンのページを前半に差し込む ---
// メンバープレゼンのテンプレートで人数ぶんのページを作り、それを前半スライドの
// 「ウィークリープレゼンテーション」の見出しページの直後へ差し込む。
// アンバサダー・ディレクター・2分30秒の下書きのページは、差し込んだページの後ろに残る
// （アンバサダー・ディレクターは applyWeeklyGuests_ で表示を切り替える）。
function insertMemberPresen_(parts, items) {
  var anchor = weeklyAnchor_(parts);
  if (!anchor) return { message: '前半スライドに「ウィークリープレゼンテーション」の見出しページが見つかりませんでした。' };
  var file = getBigTemplateFile_(MP_TEMPLATE_KIND_);
  var src = unzipToMap_(file.getBlob());
  var built = buildMemberPresenSlides_(src, items);
  var sp = spliceSlides_(parts, src, slideOrder_(src), anchor);
  var auto = false;
  for (var i = 0; i < items.length; i++) if (items[i].autoAdvanceMs) auto = true;
  var msg = 'メンバープレゼンのページを ' + sp.paths.length + '枚 差し込みました（'
          + (auto ? '自動で次へ' : 'クリックで次へ') + '）。';
  if (built.noPhoto && built.noPhoto.length) msg += '\n写真が見つからない方: ' + built.noPhoto.join('、');
  if (sp.missingLayout.length) msg += '\n※ 同じ名前のレイアウトが無いページがありました（見た目が変わる可能性があります）。';
  return { message: msg, paths: sp.paths, auto: auto };
}

// pptxの中身を書き換える本体。Driveの読み書きから切り離してあるので、
// 手元で同じ処理を走らせて確かめられる（tools/check_meeting_output.js）。
function editMeetingSlides_(parts, map, rules, o) {
  var touched = 0, byPattern = 0, path;
  // 発表のページを人数ぶんに増やす（先にページを増やしてから文字を差し替える）
  //   前半 … ウィークリープレゼン（メンバープレゼンと同じ内容）
  //   後半 … リファーラル発表
  // 写真は1つの控えで取り込む。同じ方の写真は1枚だけ入れて使い回し、
  // 別の処理が同じ名前の画像を作って上書きし合うこと（写真の取り違い）も起きない。
  var photoCache = { by: {}, seq: 0 };
  var referral = (o.referral && o.referral.length)
    ? expandPresenterSlides_(parts, o.referral,
        { title: RF_TITLE_, label: 'リファーラル発表', seconds: RF_SECONDS_, photoCache: photoCache }) : null;
  // 前半：アンバサダー・ディレクターのページの表示を切り替える（差し込みより先に。同じ作りのため）
  var guests = o.weeklyGuests ? applyWeeklyGuests_(parts, o.weeklyGuests, o.weeklyAuto !== false) : null;
  // 前半：メンバープレゼンのページを差し込む（メンバープレゼンのテンプレートから作る）
  var weekly = (o.memberPresen && o.memberPresen.length) ? insertMemberPresen_(parts, o.memberPresen) : null;
  // 自動で進むページを作ったときだけ「保存済みのタイミングを使用」を入れる。
  // そのとき、自分で作っていないページに残っている自動送り（0.4秒など）は外す。
  // 元は全部無視されていた値なので、外しても元の動きのまま。
  var ours = [], anyAuto = false;
  [referral, weekly, guests].forEach(function (r) {
    if (!r) return;
    if (r.auto) anyAuto = true;
    ours = ours.concat(r.paths || []);
  });
  var stripped = anyAuto ? mpUseTimingsOnly_(parts, ours) : 0;
  // 写真の差し替えは {{ }} を消す前に行う（差し込み口の名前でページを探すため）
  // 推薦のことば：組の数だけページを作る（定例会中はその場所、アフター・定例会後は抽選コーナーのあと）
  var reco = (o.recommendPairs && o.recommendPairs.length !== undefined)
    ? expandRecommendations_(parts, o.recommendPairs, photoCache) : null;
  // 書記兼会計による報告（更新状況一覧）
  var renewal = applyRenewalStatus_(parts, map);
  var photoMsgs = [], TWO = [
    { prefix: 'メインプレゼン', names: o.mainPresenters },
    { prefix: '推薦のことば',   names: reco ? null : o.recommenders },
    { prefix: '抽選',           names: o.lottery }
  ];
  for (var w = 0; w < TWO.length; w++) {
    var ph = applyTwoPersonPhotos_(parts, TWO[w].prefix, TWO[w].names || [], photoCache);
    if (ph && ph.message) photoMsgs.push(ph.message);
    fitTwoPersonPage_(parts, TWO[w].prefix, map);      // 長いお名前・会社名が下の段に重ならないように
  }
  var photos = { message: photoMsgs.join('\n') };
  for (path in parts) {
    if (!/^ppt\/(slides|notesSlides)\/[^\/]+\.xml$/.test(path)) continue;   // 本文だけ
    var xml = xmlOf_(parts, path), before = xml;
    if (!xml) continue;
    if (xml.indexOf('{{') !== -1) xml = replaceTokensInXml_(xml, map);
    if (rules && rules.length) {
      var pr = replacePatternsInXml_(xml, rules);
      if (pr.changed) { xml = pr.xml; byPattern += pr.changed; }
    }
    if (xml !== before) { putXml_(parts, path, xml); touched++; }
  }
  var core = o.coreValue ? applyCoreValue_(parts, o.coreValue) : null;
  var policy = o.generalPolicy ? applyGeneralPolicy_(parts, o.generalPolicy) : null;
  var audio = o.music ? applyMeetingAudio_(parts, o.music) : null;
  return { touched: touched, byPattern: byPattern, core: core, policy: policy,
           photos: photos, audio: audio, referral: referral, weekly: weekly, guests: guests,
           reco: reco, renewal: renewal };
}

// 定例会スライドを生成する。テンプレート内の {{キー}} を置換する方式。
// pptxのままサーバー側で加工し、Driveへ保存してURLを返す。
// opts: { patterns: true/false（差し込み口が無いページの第○回・日付も直す）,
//         coreValue: 'Givers Gain' など,
//         memberPresen: [...]（前半に差し込むメンバープレゼンのページ）,
//         weeklyGuests: ['坂爪　達也', …]（表示にするアンバサダー・ディレクター。null なら触らない）,
//         weeklyAuto: true/false（ウィークリープレゼンを自動で次へ進めるか） }
function generateMeetingSlides(kind, values, meetingDateVal, opts) {
  try {
    if (!BIG_TEMPLATE_KINDS_[kind]) return { ok: false, message: 'スライドの種類が不正です。' };
    var map = values || {}, o = opts || {};
    var mmdd = '';
    var d = parseDate_(meetingDateVal);
    if (d) mmdd = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyyMMdd');
    var label = BIG_TEMPLATE_KINDS_[kind].label;
    var outName = (mmdd ? mmdd + '_' : '') + label + '.pptx';

    var rules = (o.patterns === false) ? []
              : meetingPatternRules_(String(map['開催回'] || '').replace(/[^\d]/g, ''), d);

    var r = editPptxOnServer_(kind, outName, function (parts) {
      return editMeetingSlides_(parts, map, rules, o);
    });

    var info = r.info || {};
    var msg = '「' + label + '」を作成しました（' + r.partCount + 'パーツ／' +
              (r.timing.合計 / 1000).toFixed(1) + '秒）。';
    if (!info.touched) {
      msg += '\n※ 書き換える箇所が見つかりませんでした（{{ }} も「第○回」も見つかりません）。';
    } else {
      msg += '\n' + info.touched + '枚のスライドを書き換えました';
      msg += info.byPattern ? '（うち「第○回・日付」を直した段落 ' + info.byPattern + 'か所）。' : '。';
    }
    if (info.core) msg += '\n' + info.core.message;
    if (info.policy) msg += '\n' + info.policy.message;
    if (info.photos && info.photos.message) msg += '\n' + info.photos.message;
    if (info.audio && info.audio.message) msg += '\n' + info.audio.message;
    if (info.referral && info.referral.message) msg += '\n' + info.referral.message;
    if (info.weekly && info.weekly.message) msg += '\n' + info.weekly.message;
    if (info.guests && info.guests.message) msg += '\n' + info.guests.message;
    if (info.reco && info.reco.message) msg += '\n' + info.reco.message;
    if (info.renewal && info.renewal.message) msg += '\n' + info.renewal.message;
    return { ok: true, message: msg, url: r.saved.url, downloadUrl: r.saved.downloadUrl,
             fileName: outName, touched: info.touched, core: info.core, policy: info.policy,
             timing: r.timing };
  } catch (e) {
    console.error('[MEETING] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'スライドの作成に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}
