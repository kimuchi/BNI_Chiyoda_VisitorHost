// === 定例会スライドの自動更新 ===
// メンバー名簿の入会日・更新日・更新期限日から、スライドに載せる内容を自動算出する。
// 手入力していた「新メンバー」「更新メンバー」「更新状況一覧(90/60/30日)」が不要になる。

function openMeetingSlideDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('slides_meeting').setWidth(900).setHeight(760),
    '定例会スライドの自動更新');
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
               return { name: m.name, company: m.company, title: m.title };
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

// --- メインプレゼンのお2人の写真 ---
// {{メインプレゼン1氏名}} が載っているページを探し、大きな写真2枚を左右で割り当てる。
// 背景いっぱいの画像や小さな飾りを拾わないよう、大きさで絞る。
function applyMainPresenterPhotos_(parts, names) {
  if (!names || !names.length) return null;
  var path = null, xml = null, p;
  for (p in parts) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(p)) continue;
    var x = xmlOf_(parts, p);
    if (x && x.indexOf('メインプレゼン1氏名') >= 0) { path = p; xml = x; break; }
  }
  if (!path) return { message: '' };

  var prs = xmlOf_(parts, 'ppt/presentation.xml') || '';
  var sz = prs.match(/<p:sldSz\s+cx="(\d+)"/);
  var slideW = sz ? parseInt(sz[1], 10) : 12192000;

  var pics = [], ranges = findTagRanges_(xml, 'p:pic');
  for (var i = 0; i < ranges.length; i++) {
    var seg = xml.substring(ranges[i].start, ranges[i].end);
    var id = (seg.match(/<p:cNvPr[^>]*\sid="(\d+)"/) || [])[1];
    var off = seg.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>/);
    var ext = seg.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
    var rid = (seg.match(/<a:blip[^>]*r:embed="([^"]+)"/) || [])[1];
    if (!id || !off || !ext || !rid) continue;
    var cx = parseInt(ext[1], 10), cy = parseInt(ext[2], 10);
    if (cy < 1500000) continue;                       // 小さな飾りは対象外
    if (cx > slideW * 0.4) continue;                  // 背景いっぱいの画像は対象外
    pics.push({ id: id, rid: rid, cx: cx, cy: cy, center: parseInt(off[1], 10) + cx / 2 });
  }
  pics.sort(function (a, b) { return a.center - b.center; });
  if (pics.length < 1) return { message: 'メインプレゼンのページで写真の枠が見つかりませんでした。' };

  var rp = 'ppt/slides/_rels/' + path.replace(/^.*\//, '') + '.rels';
  var rels = xmlOf_(parts, rp), cache = { by: {}, seq: 0 }, done = [], miss = [];
  for (var k = 0; k < pics.length && k < names.length; k++) {
    if (!names[k]) continue;
    var photo = mpAddPhoto_(parts, cache, names[k]);
    if (!photo) { miss.push(names[k]); continue; }
    var box = readShapeGeomEmu_(xml, pics[k].id);
    if (box && photo.width && photo.height) {
      xml = setSrcRectInPic_(xml, pics[k].id, coverCrop_(photo.width, photo.height, box.cx, box.cy));
    }
    rels = retargetRel_(rels, pics[k].rid, '../media/' + photo.path.replace('ppt/media/', ''));
    done.push(names[k]);
  }
  putXml_(parts, path, xml);
  if (rels) putXml_(parts, rp, rels);
  var msg = done.length ? ('メインプレゼンの写真を差し替えました: ' + done.join('、')) : '';
  if (miss.length) msg += (msg ? '／' : '') + '写真が見つからない方: ' + miss.join('、');
  return { message: msg, replaced: done.length };
}

// pptxの中身を書き換える本体。Driveの読み書きから切り離してあるので、
// 手元で同じ処理を走らせて確かめられる（tools/check_meeting_output.js）。
function editMeetingSlides_(parts, map, rules, o) {
  var touched = 0, byPattern = 0, path;
  // 写真の差し替えは {{ }} を消す前に行う（差し込み口の名前でページを探すため）
  var photos = applyMainPresenterPhotos_(parts, o.mainPresenters || []);
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
  return { touched: touched, byPattern: byPattern, core: core, policy: policy, photos: photos };
}

// 定例会スライドを生成する。テンプレート内の {{キー}} を置換する方式。
// pptxのままサーバー側で加工し、Driveへ保存してURLを返す。
// opts: { patterns: true/false（差し込み口が無いページの第○回・日付も直す）,
//         coreValue: 'Givers Gain' など }
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
    return { ok: true, message: msg, url: r.saved.url, downloadUrl: r.saved.downloadUrl,
             fileName: outName, touched: info.touched, core: info.core, policy: info.policy,
             timing: r.timing };
  } catch (e) {
    console.error('[MEETING] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'スライドの作成に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}
