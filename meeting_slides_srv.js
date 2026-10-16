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
  var stamp = file.getId() + ':' + file.getLastUpdated().getTime() + ':native1';
  var props = PropertiesService.getScriptProperties(), key = 'BNI_TPL_MEMO_' + name;
  try {
    var memo = JSON.parse(props.getProperty(key) || 'null');
    if (memo && memo.stamp === stamp) return memo.value;
  } catch (e) {}
  var parts = unzipToMap_(file.getBlob());
  if (typeof prepareNativeMeetingTemplate_ === 'function') prepareNativeMeetingTemplate_(parts);
  var value = compute(parts);
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

// 氏名の配列 → 「山田さん、加納さん」。該当なしは「該当者なし」
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
    var tpl = getBigTemplateStatus().templates, ready = {};
    for (var i = 0; i < tpl.length; i++) ready[tpl[i].kind] = tpl[i].registered;
    var members = getMemberMaster().members || [];
    // その日のコアバリュー・開催回は、ルーティンチェックシートに書いてある。
    // 前半の新メンバー・更新メンバー・バイスプレジデントによる報告・ネットワーキングリーダーも（firstHalf）
    var routine = first ? getRoutineInfo(first.dateValue, { firstHalf: true }) : null;
    return { ok: true, meetings: candidates, defaultMeeting: first,
             lists: lists.ok ? lists : null, templates: ready,
             routine: routine, seconds: chapterPresenSeconds_(),
             coreValues: CORE_VALUES_.map(function (c) { return c.label; }),
             memberCount: members.length,
             memberNames: members.map(function (m) { return m.name; }),
             members: members.map(function (m) {
               return { no: m.no, name: m.name, company: m.company, title: m.title, collab: m.collab || '',
                        hasPhoto: !!findPhotoIdForName_(m.name) };
             }) };
  } catch (e) {
    console.error('[MEETING] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '初期表示の取得に失敗しました: ' + (e && e.message ? e.message : e) };
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

// コアバリューのページを見分ける（テンプレートの字で）。
// まず英語だけで探す。「責任」「伝統」のような語はふつうの文にも出てくるため、
// 英語で足りているならそちらを使う方が誤検出が少ない。
function coreValuePages_(parts) {
  var scan = function (useJa) {
    var cand = [], kinds = {}, path;
    for (path in parts) {
      if (!/^ppt\/slides\/slide\d+\.xml$/.test(path)) continue;
      var xml = xmlOf_(parts, path);
      if (!xml) continue;
      var found = coreValuesInSlide_(xml, useJa);
      if (found.length !== 1) continue;
      cand.push({ path: path, cv: found[0] });
      kinds[found[0].label] = true;
    }
    var n = 0;
    for (var k in kinds) n++;
    return { cand: cand, kinds: n, useJa: useJa };
  };
  var res = scan(false);
  return res.kinds < 3 ? scan(true) : res;
}

// pages … 差し込み・書き換えの前に coreValuePages_ で見分けたもの（渡さなければ、ここで見分ける）。
// 以前は書き換えのあとで見分けていたため、メンバーの会社名・カテゴリーに「Building」「Positive」などの字があると、
// そのメンバーのページやスピーカーローテーションのページまで、コアバリューのページとして非表示にしていた
function applyCoreValue_(parts, wanted, pages) {
  var w = coreValueOf_(wanted);
  if (!w) return null;
  var res = pages || coreValuePages_(parts);
  var cand = res.cand, kindCount = res.kinds;
  if (kindCount < 3) {
    return { value: w.label, shown: 0, hidden: 0, kinds: kindCount,
             message: 'コアバリューのページを見分けられませんでした（' + kindCount + '種類）。表示の切り替えは行っていません。' };
  }
  var shown = [], hidden = [];
  for (var i = 0; i < cand.length; i++) {
    // いまの中身で見直す（書き換えでコアバリューのページでなくなったものは触らない）
    var xml = xmlOf_(parts, cand[i].path), now = xml ? coreValuesInSlide_(xml, res.useJa) : [];
    if (now.length !== 1 || now[0].label !== cand[i].cv.label) continue;
    var on = (cand[i].cv.label === w.label);
    var out = setSlideShow_(xml, on);
    if (out !== xml) putXml_(parts, cand[i].path, out);
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
//
// メインプレゼン・抽選は、1人の週がある（opt.single）。そのときは、もう1人の文字箱も消し、残る方の写真の枠と
// 文字箱を左右のまん中に動かす（twoPersonSinglePlan_）。①②のどちらの欄に入れても同じ。推薦のことばは、いつも2人。
// ディレクター（名簿に無く、前半テンプレートにアンバサダー・ディレクターのページがある方）がメインプレゼンを担当する週は、
// 写真をそのページの写真にする（opt.photoOf。twoPersonDirectors_・twoPersonGuestPhoto_）。
function applyTwoPersonPhotos_(parts, prefix, names, cache, opt) {
  if (!names || !names.length) return null;
  var path = null, p;
  for (p in parts) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(p)) continue;
    var x = xmlOf_(parts, p);
    if (x && x.indexOf(prefix + '1氏名') >= 0) { path = p; break; }
  }
  if (!path) return { message: '' };
  var r = setTwoPersonPhotos_(parts, path, prefix, names, cache, opt);
  var msg = r.done.length ? (prefix + 'の写真を差し替えました: ' + r.done.join('、')) : '';
  // どのページの方かが分かるように、ページの名前を付ける（後半は、メインプレゼンと抽選の両方に同じ方が出ることがある）
  if (r.missing.length) msg += (msg ? '／' : '') + prefix + 'で写真が見つからない方（写真なし）: ' + r.missing.join('、');
  if (r.noFrame) msg = prefix + 'のページで写真の枠が見つかりませんでした。';
  r.guests.forEach(function (n) {
    msg += (msg ? '\n' : '') + (prefix === 'メインプレゼン' ? '' : prefix + 'の') + 'ディレクター（' + n
      + 'さん）の写真は、前半テンプレートのその方のページから入れました';
  });
  if (r.single && (r.single.moved || r.single.removed)) {
    msg += (msg ? '\n' : '') + prefix + 'は1人なので、' + (r.single.moved ? 'まん中に置きました'
      : 'もう1人の欄を消しました' + (r.single.blocked ? '（まん中に動かすと飾りの画像に重なるので、位置はテンプレートのままです）' : ''));
  }
  return { message: msg, replaced: r.done.length, single: r.single };
}

// 1枚のページの写真を差し替える（差し込み口 {{○○1氏名}} がまだ残っている状態で呼ぶこと）。
//   opt.single  … 1人だけのとき、もう1人の欄を消し、残る方をまん中に置く（メインプレゼン・抽選）
//   opt.photoOf … 氏名 → 名簿の写真より先に使う写真 { path, width, height }（ディレクターのページの写真）。無ければ null
// 戻り値 { done: [差し替えた方], missing: [写真が無かった方], noFrame: 枠が無かったか,
//          guests: [ディレクターのページの写真を入れた方], single: { name, moved, blocked, removed }（1人のとき） }
function setTwoPersonPhotos_(parts, path, prefix, names, cache, opt) {
  opt = opt || {};
  var xml = xmlOf_(parts, path), res = { done: [], missing: [], noFrame: false, guests: [], single: null };
  var prs = xmlOf_(parts, 'ppt/presentation.xml') || '';
  var sz = prs.match(/<p:sldSz\s+cx="(\d+)"/);
  var slideW = sz ? parseInt(sz[1], 10) : 12192000;
  // いつも2人ぶん（1人ぶんだけ渡されても、2人目の枠にテンプレートの元の写真を残さない）
  var two = [String(names[0] || ''), String(names[1] || '')];

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
  res.noFrame = !pics.length;
  var slots = pics.length ? pickPhotoFrames_(xml, prefix, 2, pics) : [];
  // 1人だけのときの置き方は、写真の枠を消す前に決める（写真の無い方も、枠のある形でまん中を決める）
  var one = two[0] ? (two[1] ? -1 : 0) : (two[1] ? 1 : -1);
  var plan = (opt.single && one >= 0) ? twoPersonSinglePlan_(xml, prefix, one, slots, slideW) : null;

  var rp = relsPathOf_(path), rels = xmlOf_(parts, rp);
  cache = cache || { by: {}, seq: 0 };
  for (var k = 0; k < slots.length; k++) {
    var pic = slots[k];
    if (!pic) continue;
    var own = (two[k] && opt.photoOf) ? opt.photoOf(two[k]) : null;
    var photo = own || (two[k] ? mpAddPhoto_(parts, cache, two[k]) : null);
    if (!photo) {
      xml = removeShape_(xml, pic.id);
      if (two[k]) res.missing.push(two[k]);
      continue;
    }
    var box = readShapeGeomEmu_(xml, pic.id);
    if (box && photo.width && photo.height) {
      xml = setSrcRectInPic_(xml, pic.id, coverCrop_(photo.width, photo.height, box.cx, box.cy));
    }
    var set = setPicImage_(xml, rels, pic.id, relPathFrom_(posixDir_(path), photo.path));
    xml = set.xml; rels = set.rels;
    res.done.push(two[k]);
    if (own) res.guests.push(two[k]);
  }
  if (plan) {
    xml = twoPersonSingleApply_(xml, plan);
    res.single = { name: two[one], moved: plan.moved, blocked: plan.blocked, removed: plan.remove.length };
  }
  putXml_(parts, path, xml);
  if (rels) putXml_(parts, rp, rels);
  return res;
}

// --- 1人だけの週（メインプレゼン・抽選）---
// k … 残る方（0＝1人目の欄・1＝2人目の欄）。残る方の文字箱（{{○○N氏名}}・会社名・カテゴリー）と写真の枠をひとまとまりにして、
// 左右のまん中に来る動かし量（dx）を決め、もう1人の文字箱を消す。上下の位置はそのまま。ほかの飾り（題名・動画など）は動かさない。
// 動かした先が、ほかの画像（抽選コーナーの中央下のプレゼントの絵など）に新しく重なるときは、動かさない（blocked）。
// 位置の書いていない文字箱（レイアウトから受け継ぐもの）があるときも、動かさない。
var TWO_PERSON_FIELDS_ = ['氏名', '会社名', 'カテゴリー'];
function twoPersonSinglePlan_(xml, prefix, k, slots, slideW) {
  var geo = shapeBoxes_(xml);
  var move = twoPersonSlotShapes_(xml, prefix, k + 1, 2 - k);
  if (slots[k]) move.push(slots[k].id);
  var plan = { remove: twoPersonSlotShapes_(xml, prefix, 2 - k, k + 1), move: move, dx: 0, sx: {}, moved: false, blocked: false };
  if (!move.length) return plan;
  var lo = Infinity, hi = -Infinity, i, j;
  for (i = 0; i < move.length; i++) {
    var b = geo[move[i]];
    if (!b) return plan;
    lo = Math.min(lo, b.x); hi = Math.max(hi, b.x + b.cx);
    plan.sx[move[i]] = b.sx;
  }
  plan.dx = Math.round(slideW / 2 - (lo + hi) / 2);
  if (!plan.dx) return plan;
  // ほかの画像（2人の写真の枠でないもの）に、動かしたら新しく重なるか（もともと重なっている背景などは数えない）
  var skip = {};
  move.forEach(function (id) { skip[id] = true; });
  slots.forEach(function (s) { if (s) skip[s.id] = true; });
  var hit = function (a, dx, c) { return a.x + dx < c.x + c.cx && c.x < a.x + dx + a.cx && a.y < c.y + c.cy && c.y < a.y + a.cy; };
  for (var id in geo) {
    if (skip[id] || !geo[id].pic) continue;
    for (j = 0; j < move.length; j++) {
      if (hit(geo[move[j]], plan.dx, geo[id]) && !hit(geo[move[j]], 0, geo[id])) { plan.blocked = true; return plan; }
    }
  }
  plan.moved = true;
  return plan;
}
// n 人目の欄の文字箱（その方の差し込み口があり、もう1人＝m 人目の差し込み口の無い文字箱）の図形ID
function twoPersonSlotShapes_(xml, prefix, n, m) {
  var out = [], ranges = findTagRanges_(xml, 'p:sp');
  var has = function (t, no) {
    return TWO_PERSON_FIELDS_.some(function (f) { return t.indexOf('{{' + prefix + no + f + '}}') >= 0; });
  };
  for (var i = 0; i < ranges.length; i++) {
    var seg = xml.substring(ranges[i].start, ranges[i].end), t = slideText_(seg);
    var id = (seg.match(/<p:cNvPr[^>]*\sid="(\d+)"/) || [])[1];
    if (id && has(t, n) && !has(t, m) && out.indexOf(id) < 0) out.push(id);
  }
  return out;
}
function twoPersonSingleApply_(xml, plan) {
  plan.remove.forEach(function (id) { xml = removeShape_(xml, id); });
  if (!plan.moved) return xml;
  plan.move.forEach(function (id) {
    var r = findShapeRange_(xml, id);
    if (!r) return;                                   // 写真が見つからず、枠を消した方
    // グループの中の図形は、グループの子の座標で動かす（グループが縮めていれば、そのぶん大きく）
    var d = Math.round(plan.dx / (plan.sx[id] || 1));
    var seg = xml.substring(r.start, r.end).replace(/(<a:off\s+x=")(-?\d+)(")/, function (all, a, x, b) {
      return a + (parseInt(x, 10) + d) + b;
    });
    xml = xml.substring(0, r.start) + seg + xml.substring(r.end);
  });
  return xml;
}
// ページの上での図形（文字箱・画像）の位置（EMU）→ { 図形ID: { x, y, cx, cy, sx, pic } }。
// グループの中の図形は、グループの子の座標（chOff・chExt）をページの座標に直す（sx … 横の倍率）。位置の書いていない図形は入れない
function shapeBoxes_(xml) {
  var groups = [], out = {};
  var nums = function (s, re) { var m = s.match(re); return m ? [parseInt(m[1], 10), parseInt(m[2], 10)] : null; };
  var walk = function (from, to) {
    findTagRanges_(xml.substring(from, to), 'p:grpSp').forEach(function (r) {
      var s = from + r.start, e = from + r.end;
      var pr = (xml.substring(s, e).match(/<p:grpSpPr\b[^>]*>([\s\S]*?)<\/p:grpSpPr>/) || [])[1] || '';
      groups.push({ start: s, end: e, off: nums(pr, /<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"/), ext: nums(pr, /<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/),
                    choff: nums(pr, /<a:chOff\s+x="(-?\d+)"\s+y="(-?\d+)"/), chext: nums(pr, /<a:chExt\s+cx="(\d+)"\s+cy="(\d+)"/) });
      walk(xml.indexOf('>', s) + 1, e);                // 入れ子のグループ
    });
  };
  walk(0, xml.length);
  ['p:sp', 'p:pic'].forEach(function (tag) {
    findTagRanges_(xml, tag).forEach(function (r) {
      var seg = xml.substring(r.start, r.end);
      var id = (seg.match(/<p:cNvPr[^>]*\sid="(\d+)"/) || [])[1];
      var off = nums(seg, /<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"/), ext = nums(seg, /<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
      if (!id || !off || !ext) return;
      var b = { x: off[0], y: off[1], cx: ext[0], cy: ext[1], sx: 1, pic: tag === 'p:pic' };
      groups.filter(function (g) { return g.start < r.start && r.end <= g.end; })
        .sort(function (a, c) { return c.start - a.start; })          // 内側のグループから
        .forEach(function (g) {
          if (!g.off || !g.ext || !g.choff || !g.chext) return;
          var fx = g.chext[0] ? g.ext[0] / g.chext[0] : 1, fy = g.chext[1] ? g.ext[1] / g.chext[1] : 1;
          b = { x: g.off[0] + (b.x - g.choff[0]) * fx, y: g.off[1] + (b.y - g.choff[1]) * fy,
                cx: b.cx * fx, cy: b.cy * fy, sx: b.sx * fx, pic: b.pic };
        });
      out[id] = b;
    });
  });
  return out;
}

// --- ディレクターがメインプレゼンを担当する週 ---
// 画面は、ディレクターを前半テンプレートのアンバサダー・ディレクターのページの氏名で送ってくる（名簿の方と同じ欄）。
// 名簿に無く、そのページがある方をディレクターとみなし、写真（と、画面から来なければ会社名・カテゴリー）をそのページから入れる。
// 前半はこのファイルのそのページ、後半は前半のテンプレート（o.mainPageSource。抽選だけのときは o.guestSource）のページ。
// メンバーのページを差し込む前に呼ぶこと（差し込んだメンバーのページも、同じ WEEKLY PRESENTATION の作りのため）。
//   戻り値 { normName_(送られた氏名): { name, role, company, photo: 'ppt/media/…', src } }
function twoPersonDirectors_(parts, o) {
  var names = [].concat(o.mainPresenters || [], o.lottery || []).filter(function (n) { return n; }), out = {};
  if (!names.length) return out;
  var keys = o.rosterKeys || (o.rosterKeys = rosterNameKeys_());
  var others = names.filter(function (n) { return !keys[normName_(n)]; });
  if (!others.length) return out;
  var src = o.mainPageSource || o.guestSource || parts, pages = [];
  try { pages = weeklyGuestPages_(src); }
  catch (e) { console.warn('[MEETING] アンバサダー・ディレクターのページを調べられませんでした: ' + (e && e.message ? e.message : e)); }
  others.forEach(function (n) {
    var g = guestByName_(pages, n);
    if (g) out[normName_(n)] = { name: g.name, role: g.role || '', company: g.company || '', photo: g.photo || '', src: src };
  });
  return out;
}
// 名簿の方の氏名（normName_）→ true。読めなければ空（そのときは、どの方もアンバサダー・ディレクターのページで探す）
function rosterNameKeys_() {
  var keys = {};
  try { (getMemberMaster({ membersOnly: true }).members || []).forEach(function (m) { keys[normName_(m.name)] = true; }); }
  catch (e) { console.warn('[MEETING] メンバー名簿を読めませんでした: ' + (e && e.message ? e.message : e)); }
  return keys;
}
// 名簿に無い方（後半で、抽選だけにディレクターを選んだとき、前半のテンプレートを開くかどうか）
function twoPersonNonMembers_(names, o) {
  var list = (names || []).filter(function (n) { return n; });
  if (!list.length) return [];
  var keys = o.rosterKeys || (o.rosterKeys = rosterNameKeys_());
  return list.filter(function (n) { return !keys[normName_(n)]; });
}
// ディレクターの写真 → { path, width, height }（ページに写真の無い方は null）。
// 前半は同じファイルの画像をそのまま使う。後半は前半のテンプレートから画像を写す（種類も [Content_Types].xml に足す）。1人1回だけ
function twoPersonGuestPhoto_(parts, g) {
  if (!g || !g.photo) return null;
  if (g.got !== undefined) return g.got;
  g.got = null;
  try {
    var path = g.photo;
    if (g.src !== parts) {
      var ext = (path.match(/\.([^.\/]+)$/) || [])[1] || 'png', n = 1, dst, b = g.src[path];
      if (!b) return null;
      do { dst = 'ppt/media/guestphoto' + (n++) + '.' + ext; } while (parts[dst]);
      parts[dst] = b.copyBlob ? b.copyBlob().setName(dst) : b.setName(dst);
      ctPartType_(parts, g.src, path, dst);
      path = dst;
    }
    var size = null;
    try { size = imageSizeOf_(parts[path].getBytes()); } catch (e) {}
    g.got = { path: path, width: size ? size.width : 0, height: size ? size.height : 0 };
  } catch (e) {
    console.warn('[MEETING] ' + g.name + 'さんのページの写真を使えませんでした: ' + (e && e.message ? e.message : e));
  }
  return g.got;
}
// ディレクターの氏名・会社名・カテゴリーが画面から来なければ（サーバーから直接呼んだときなど）、その方のページの文字で埋める。
// カテゴリーは、そのページのカテゴリーの枠の文字（「BNI東京千代田リージョン」など）を【】で囲む（もう囲んであればそのまま）
function twoPersonDirectorValues_(values, two, directors) {
  var out = {}, k;
  for (k in values) out[k] = values[k];
  two.forEach(function (t) {
    if (t.prefix === RECO_PREFIX_) return;
    (t.names || []).slice(0, 2).forEach(function (n, i) {
      var g = n ? directors[normName_(n)] : null, p = t.prefix + (i + 1);
      if (!g) return;
      if (!out[p + '氏名']) out[p + '氏名'] = n;
      if (!out[p + '会社名'] && g.company) out[p + '会社名'] = CB_.normalize(g.company);
      if (!out[p + 'カテゴリー'] && g.role) out[p + 'カテゴリー'] = roleInBrackets_(g.role);
    });
  });
  return out;
}
function roleInBrackets_(role) {
  var r = CB_.normalize(role);
  return !r ? '' : (/^【[\s\S]*】$/.test(r) ? r : '【' + r + '】');
}

// お2人が並ぶページ（メインプレゼン・推薦のことば・抽選）の氏名・会社名・カテゴリーの文字箱は、
// 幅が決まっていて、すぐ下に次の段がある。長いお名前や会社名が折り返すと下の段に重なるので、
// 収まらないときだけ文字を小さくする（氏名・会社名は1行、カテゴリーは2行まで）。
// 文字の幅は全角1文字＝1em・半角スペース＝0.3em・ほかの半角＝0.55em で見積もり、
// 太字や代わりのフォント（Meiryo UI が無い環境）でも収まるよう、幅は5%ほどゆとりをみる。
// 会社名（company）は、ほかの出力と同じ会社名の見積もり（company_break.js の CB_.lineWidthPt。英字は字ごとの幅・余裕こみ）で、
// 枠の内側の幅に1行で入る大きさにする（英字の多い会社名は 0.55em では縮め足りず、PowerPoint が2行に折り返して下の段に重なった）。
// 測る会社名は名簿の書き方にそろえたもの（CB_.normalize）。入れる値も同じにそろえておくこと（twoPersonCompanies_）。
// 差し込み口 {{○○1会社名}} などがまだ残っている状態で呼ぶこと。
var TWO_PERSON_FIT_ = [
  { field: '氏名', lines: 1, min: 20 },
  { field: '会社名', lines: 1, min: 8, company: true },   // 支社・営業所まで入る長い会社名もあるので、小さめまで許す
  { field: 'カテゴリー', lines: 2, min: 9 }
];
function fitTwoPersonText_(xml, prefix, values) {
  for (var k = 1; k <= 2; k++) {
    for (var f = 0; f < TWO_PERSON_FIT_.length; f++) {
      var t = TWO_PERSON_FIT_[f], key = prefix + k + t.field;
      if (!values[key]) continue;
      var text = t.company ? CB_.normalize(values[key]) : String(values[key]);
      xml = fitTokenBox_(xml, '{{' + key + '}}', text, t.lines, t.min, t.company);
    }
  }
  return xml;
}
// お2人のページ（prefixes）の会社名を、名簿の書き方にそろえた値の写し（CB_.normalize：セルの中の改行を外す・㈱を「株式会社」に）。
// 画面から改行の残った会社名が来ても、測った文字と入れる文字を同じにする（改行が残ると、1行の枠で2行になり下の段に重なる）
function twoPersonCompanies_(values, prefixes) {
  var out = {}, k;
  for (k in values) out[k] = values[k];
  for (var p = 0; p < prefixes.length; p++) {
    for (var n = 1; n <= 2; n++) {
      k = prefixes[p] + n + '会社名';
      if (out[k]) out[k] = CB_.normalize(out[k]);
    }
  }
  return out;
}
// 1人の週は、1人目の欄を消して2人目の欄だけが残ることもある（twoPersonSinglePlan_）ので、どちらかの差し込み口のあるページ
function fitTwoPersonPage_(parts, prefix, values) {
  for (var path in parts) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(path)) continue;
    var xml = xmlOf_(parts, path);
    if (!xml || (xml.indexOf(prefix + '1氏名') < 0 && xml.indexOf(prefix + '2氏名') < 0)) continue;
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
// company … 会社名。CB_.lineWidthPt（余裕こみ）で、枠の内側の幅（左右の余白を除く）に入る大きさを 0.5pt 刻みで
function fitTokenBox_(xml, token, text, maxLines, minPt, company) {
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
    var innerPt = (parseInt(ext[1], 10) - ins('lIns') - ins('rIns')) / 12700, widthPt = innerPt * 0.95;
    var pt = parseInt((seg.match(/<a:rPr\b[^>]*\ssz="(\d+)"/) || [0, 1800])[1], 10) / 100;
    var size;
    if (company) {
      var per = CB_.lineWidthPt(text, 1);                      // 1pt あたりの幅（余裕こみ）
      if (per <= 0 || per * pt <= innerPt * maxLines + 1e-6) return xml;
      size = Math.floor(innerPt * maxLines / per * 2) / 2;
      if (per * size > innerPt * maxLines + 1e-6) size -= 0.5; // 割り算の端数で、わずかに超えないように
      size = Math.max(minPt || 10, size);
    } else {
      var em = textBoxEm_(text);
      if (em * pt <= widthPt * maxLines) return xml;
      size = Math.max(minPt || 10, Math.floor(widthPt * maxLines / em * 2) / 2);
    }
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
// 組に受け取ったスライド（slideImage。reco_slide_srv.js の recoLoadSlides_ で開いたもの）があれば、
// その組のページのすぐ後ろに1枚のページとして入れる（recoImageSlide_）。
//   pairs … [{ giver: {name, company, category}, receiver: {…}, after: true/false, slideImage }]
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
    // 会社名は名簿の書き方にそろえる（画面から改行の残った会社名が来ても、測った文字と入れる文字を同じに。twoPersonCompanies_ と同じ）
    var vals = {};
    vals[RECO_PREFIX_ + '1氏名'] = g.name || ''; vals[RECO_PREFIX_ + '1会社名'] = CB_.normalize(g.company);
    vals[RECO_PREFIX_ + '1カテゴリー'] = g.category || '';
    vals[RECO_PREFIX_ + '2氏名'] = r.name || ''; vals[RECO_PREFIX_ + '2会社名'] = CB_.normalize(r.company);
    vals[RECO_PREFIX_ + '2カテゴリー'] = r.category || '';
    putXml_(parts, path, replaceTokensInXml_(fitTwoPersonText_(xmlOf_(parts, path), RECO_PREFIX_, vals), vals));
  };
  // 先に、手を付けていないひな形から2組目以降とアフターの組のページを作る
  var clone = function (pair) {
    var add = addSlidePart_(parts, modelXml, modelRels);
    fill(add.path, pair);
    return add;
  };
  // 受け取ったスライドのページ（その組のページのすぐ後ろに並べる）。placed … どの組のページのあとか（お知らせ用）
  var slides = [], seq = 0, placed = [];
  var slideOf = function (pair, pagePath) {
    if (!pair.slideImage) return [];
    var s = recoImageSlide_(parts, modelRels, pair.slideImage, ++seq);
    slides.push(s.path);
    placed.push({ pair: pair, page: pagePath, path: s.path });
    return [s];
  };
  var moreDuring = [], afterPages = [], firstSlide = during.length ? slideOf(during[0], model) : [], c;
  for (i = 1; i < during.length; i++) { c = clone(during[i]); moreDuring = moreDuring.concat([c], slideOf(during[i], c.path)); }
  for (i = 0; i < after.length; i++) { c = clone(after[i]); afterPages = afterPages.concat([c], slideOf(after[i], c.path)); }
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
    if (entries[i].path === model) out = out.concat(firstSlide, moreDuring);
    if (lottery && entries[i].path === lottery) { out = out.concat(afterPages); placedAfter = true; }
  }
  if (!placedAfter) out = out.concat(afterPages);        // 抽選コーナーが無ければ最後に
  setSlideEntries_(parts, out);

  // 何枚目に入れたか（PowerPoint の左の一覧の番号。非表示のページも数える）。このあとページの並びは変えない
  var pos = {};
  slideOrder_(parts).forEach(function (p, k) { pos[p] = k + 1; });
  var span = function (paths) {
    var n = paths.map(function (p) { return pos[p]; }).filter(Boolean);
    if (!n.length) return '';
    var lo = Math.min.apply(null, n), hi = Math.max.apply(null, n);
    return lo === hi ? lo + '枚目' : lo + '〜' + hi + '枚目';
  };
  var pathsOf = function (list) { return list.map(function (x) { return x.path; }); };
  var who = function (pair) { return ((pair.giver || {}).name || '（推薦する方）') + ' → ' + ((pair.receiver || {}).name || '（推薦される方）'); };
  var msg = during.length
    ? ('推薦のことばのページを ' + during.length + '枚作りました（定例会中。' + span([model].concat(pathsOf(firstSlide), pathsOf(moreDuring))) + '）')
    : '定例会中の推薦のことばは無いので、そのページは非表示にしました';
  if (after.length) msg += '。アフター・定例会後の ' + after.length + '枚は、'
    + (lottery ? '抽選コーナー（' + pos[lottery] + '枚目）のあと' : '最後') + '（' + span(pathsOf(afterPages)) + '）に入れました';
  msg += '。';
  if (slides.length) {
    msg += '\n受け取ったスライド ' + slides.length + '枚を、その組の推薦のことばのページのすぐあとに入れました: '
      + placed.map(function (x) { return who(x.pair) + ' … ' + pos[x.path] + '枚目'; }).join('、') + '。';
  }
  if (missing.length) msg += '\n推薦のことばで写真が見つからない方（写真なし）: ' + missing.join('、');
  var pagePaths = function (list) { return list.filter(function (x) { return slides.indexOf(x.path) < 0; }).map(function (x) { return x.path; }); };
  return { message: msg, during: during.length, after: after.length, slides: slides,
           paths: [model].concat(pagePaths(moreDuring), pagePaths(afterPages)) };
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
  var label = escapeXml_(name.replace(AUDIO_EXT_RE_, ''));   // 関数で置き換える（曲のファイル名の「$1」などを記号として読まない）
  xml = xml.replace(new RegExp('(<p:cNvPr id="' + audio.spid + '" name=")[^"]*(")'),
                    function (all, a, b) { return a + label + b; });
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

// --- 音楽を止める（後半：リファーラル発表のあと、推薦のことばのページに入ったところ）---
// テンプレートの音楽は「○枚のスライドにわたって再生」で鳴っている（数はテンプレートの並びで決めてある）。
// リファーラル発表のページを人数ぶんに増やす・欠席の方のページを作らないと、その数がずれて、
// 推薦のことばのページになっても音楽が止まらなかった。そこで
//   ・止めるページ（stopPath。非表示なら、そのあとの最初の表示のページ）の画面切り替えに「前のサウンドを停止」を付ける
//   ・それより前のページから鳴っていて、そのページまで届く音楽は、そのページの手前で止まる枚数に直す
//     （表示のページだけ数える）。「前のサウンドを停止」で止まる合図（onStopAudio）も付けておく
//   戻り値 { slide: 何枚目で止めるか, fixed: [枚数を直した音楽], message }
function stopMusicAt_(parts, stopPath) {
  var order = slideOrder_(parts), stop = order.indexOf(stopPath);
  if (stop < 0) return null;
  var shown = function (p) { return !/<p:sld\b[^>]*\sshow="0"/.test(xmlOf_(parts, p) || ''); };
  while (stop < order.length && !shown(order[stop])) stop++;
  if (stop >= order.length) return null;
  var fixed = [];
  listMeetingAudio_(parts).forEach(function (a) {
    var at = order.indexOf(a.slide);
    if (a.video || at < 0 || at >= stop || !shown(a.slide)) return;
    var n = 0;
    for (var k = at; k < stop; k++) if (shown(order[k])) n++;
    var xml = xmlOf_(parts, a.slide), next = mediaPlayAcross_(xml, a.spid, n);
    if (next !== xml) { putXml_(parts, a.slide, next); fixed.push((a.name || '音楽') + '（' + (at + 1) + '枚目から' + n + '枚）'); }
  });
  putXml_(parts, order[stop], transitionStopSound_(xmlOf_(parts, order[stop])));
  return { slide: stop + 1, fixed: fixed,
           message: '音楽：推薦のことばのページ（' + (stop + 1) + '枚目）に入ったところで止まるようにしました'
             + (fixed.length ? '（鳴らす枚数を直した音楽: ' + fixed.join('、') + '）' : '') + '。' };
}
// spid の音楽が n 枚より多く鳴る設定なら、n 枚にする（<p:cMediaNode numSld>。無いときは1枚）。
// 「前のサウンドを停止」で止まる合図（終わりの条件 onStopAudio）が無ければ足す
function mediaPlayAcross_(xml, spid, n) {
  var rs = findTagRanges_(xml, 'p:audio');
  for (var i = rs.length - 1; i >= 0; i--) {
    var seg = xml.substring(rs[i].start, rs[i].end);
    if (seg.indexOf('spid="' + spid + '"') < 0) continue;
    var m = seg.match(/<p:cMediaNode\b[^>]*>/);
    if (!m) continue;
    var now = +((m[0].match(/\snumSld="(\d+)"/) || [])[1] || 1);
    if (now <= n) continue;
    var tag = /\snumSld="/.test(m[0]) ? m[0].replace(/\snumSld="\d+"/, ' numSld="' + n + '"') : m[0].replace(/^<p:cMediaNode\b/, '<p:cMediaNode numSld="' + n + '"');
    seg = seg.replace(m[0], tag);
    if (!/evt="onStopAudio"/.test(seg)) {
      var END = '<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>';
      var ct = seg.match(/<p:cTn\b[^>]*?(\/?)>/);
      if (ct && ct[1]) seg = seg.replace(ct[0], ct[0].replace(/\/>$/, '>') + END + '</p:cTn>');
      else if (ct) {
        var at = seg.indexOf(ct[0]) + ct[0].length, sc = seg.indexOf('</p:stCondLst>', at), close = seg.indexOf('</p:cTn>', at);
        var put = (sc >= 0 && (close < 0 || sc < close)) ? sc + '</p:stCondLst>'.length : at;
        seg = seg.substring(0, put) + END + seg.substring(put);
      }
    }
    xml = xml.substring(0, rs[i].start) + seg + xml.substring(rs[i].end);
  }
  return xml;
}
// 画面切り替えに「前のサウンドを停止」（<p:sndAc><p:endSnd/></p:sndAc>）を付ける。切り替えが無いページには足す
function transitionStopSound_(xml) {
  var SND = '<p:sndAc><p:endSnd/></p:sndAc>';
  if (!/<p:transition\b/.test(xml)) {
    var tr = '<p:transition>' + SND + '</p:transition>';
    return xml.indexOf('</p:clrMapOvr>') >= 0 ? xml.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + tr) : xml.replace('</p:cSld>', '</p:cSld>' + tr);
  }
  return xml.replace(/<p:transition\b([^>]*?)\/>/g, '<p:transition$1>' + SND + '</p:transition>')
    .replace(/(<p:transition\b[^>]*>)([\s\S]*?)(<\/p:transition>)/g, function (all, open, body, close) {
      if (body.indexOf(SND) >= 0) return all;
      body = body.replace(/<p:sndAc>[\s\S]*?<\/p:sndAc>/g, '');
      var ext = body.indexOf('<p:extLst>');
      return open + (ext >= 0 ? body.substring(0, ext) + SND + body.substring(ext) : body + SND) + close;
    });
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
    // 「苗字　名前」「〇〇　〇〇」のままの下書き（スタートアッププレゼンのページなど）は人のページではない
    if (!name || /苗字|名前|氏名/.test(name) || /^[〇○◯\s　]+$/.test(name)) continue;
    // role … カテゴリーの枠の文字（「エグゼクティブディレクター」「BNI東京千代田リージョン」など、テンプレートに書いてあるとおり）。
    // company … 会社名の枠の文字。photo … 写真の画像（ppt/media/…）。メインプレゼンをディレクターが担当する週に使う
    out.push({ key: order[i].replace(/^.*\//, ''), path: order[i], name: name,
               role: shapeTextOf_(xml, sh.category), company: shapeTextOf_(xml, sh.company),
               photo: sh.photo ? picMediaPath_(parts, order[i], xml, sh.photo) : '',
               hidden: /<p:sld\b[^>]*\sshow="0"/.test(xml) });
  }
  return out;
}

// 画像の図形が使っている画像（ppt/media/…）。見つからなければ空
function picMediaPath_(parts, slidePath, xml, picId) {
  var r = findShapeRange_(xml, picId);
  if (!r) return '';
  var rid = (xml.substring(r.start, r.end).match(/<a:blip[^>]*r:embed="([^"]+)"/) || [])[1];
  var rels = rid ? xmlOf_(parts, relsPathOf_(slidePath)) || '' : '';
  var m = rels.match(new RegExp('<Relationship\\b[^>]*\\bId="' + rid + '"[^>]*>'));
  var t = m ? (m[0].match(/Target="([^"]+)"/) || [])[1] : '';
  if (!t || /TargetMode="External"/.test(m[0])) return '';
  var media = t.replace(/^\.\.\//, 'ppt/').replace(/^\/+/, '');
  return parts[media] ? media : '';
}

// アンバサダー・ディレクターのうち、その書き方（「見本 太郎」「見本」＝名字だけ、「見本ED」は guest に名字だけ）に当たる方。
// 氏名がそのまま同じ方 → 名字（氏名の頭）が同じ方が1人だけ のとき。当たらなければ null
function guestByName_(guests, text) {
  var t = normName_(String(text || '').replace(/(さん|様)$/, ''));
  if (!t) return null;
  var same = (guests || []).filter(function (g) { return normName_(g.name) === t; });
  if (same.length) return same[0];
  var head = (guests || []).filter(function (g) { return normName_(g.name).indexOf(t) === 0; });
  return head.length === 1 ? head[0] : null;
}

// 同じ方のウィークリープレゼンのページが、氏名の枠では見分けられない作りになっていても拾う：
// 見出しから「全員終わりましたか？」までの WEEKLY PRESENTATION のページで、見分けた方の氏名の文字があるもの
function weeklyGuestNamePages_(parts, guests) {
  var order = slideOrder_(parts), at = order.indexOf(weeklyAnchor_(parts)), got = {}, out = [], i, k;
  if (at < 0) return out;
  for (k = 0; k < guests.length; k++) got[guests[k].path] = true;
  for (i = at + 1; i < order.length; i++) {
    var xml = xmlOf_(parts, order[i]) || '', text = slideText_(xml), flat = normName_(text);
    if (flat.indexOf('終わりましたか') >= 0) break;
    if (got[order[i]] || flat.toUpperCase().indexOf('WEEKLYPRESENTATION') < 0) continue;
    for (k = 0; k < guests.length; k++) {
      if (normName_(guests[k].name) && flat.indexOf(normName_(guests[k].name)) >= 0) {
        out.push({ path: order[i], name: guests[k].name });
        break;
      }
    }
  }
  return out;
}

// アンバサダー・ディレクターの紹介のページ（「ウィークリープレゼンテーション」の見出しより前）。
// その方の氏名と、「アンバサダー」「ディレクター」「リージョン」の文字があるページ
var GUEST_INTRO_RE_ = /アンバサダー|ディレクター|リージョン|ambassador|director/i;
function weeklyGuestIntroPages_(parts, guests) {
  var order = slideOrder_(parts), at = order.indexOf(weeklyAnchor_(parts)), out = [];
  if (at < 0) at = order.length;
  for (var i = 0; i < at; i++) {
    var xml = xmlOf_(parts, order[i]) || '', text = slideText_(xml), flat = normName_(text);
    if (!GUEST_INTRO_RE_.test(text)) continue;
    for (var k = 0; k < guests.length; k++) {
      if (normName_(guests[k].name) && flat.indexOf(normName_(guests[k].name)) >= 0) {
        out.push({ path: order[i], name: guests[k].name, hidden: /<p:sld\b[^>]*\sshow="0"/.test(xml) });
        break;
      }
    }
  }
  return out;
}

function shapeTextOf_(xml, id) {
  var r = id ? findShapeRange_(xml, id) : null;
  return r ? slideText_(xml.substring(r.start, r.end)).replace(/^[\s　]+|[\s　]+$/g, '') : '';
}

// names … 表示にする方の氏名（画面のチェック）。入っていない方のページは非表示にする。
// auto  … メンバーのページと同じく、カウントダウンが終わったら自動で次へ進めるか
// カウントダウンは、メンバーと同じウィークリープレゼンテーションの秒数（チャプターの設定）
// その方の紹介のページ（見出しより前）も、同じように表示・非表示にする
function applyWeeklyGuests_(parts, names, auto) {
  var pages = weeklyGuestPages_(parts), want = {}, shown = [], hidden = [], paths = [], i;
  var sec = chapterPresenSeconds_().weekly;
  for (i = 0; i < names.length; i++) want[normName_(names[i])] = true;
  pages = pages.concat(weeklyGuestNamePages_(parts, pages));
  var intros = weeklyGuestIntroPages_(parts, pages), introShown = [], introHidden = [], introPages = [];
  for (i = 0; i < intros.length; i++) {
    var on0 = !!want[normName_(intros[i].name)];
    putXml_(parts, intros[i].path, setSlideShow_(xmlOf_(parts, intros[i].path), on0));
    (on0 ? introShown : introHidden).push(intros[i].name);
    introPages.push({ path: intros[i].path, name: intros[i].name, shown: on0 });
  }
  for (i = 0; i < pages.length; i++) {
    var g = pages[i], xml = xmlOf_(parts, g.path), on = !!want[normName_(g.name)];
    xml = setSlideShow_(xml, on);
    if (on) {
      try {
        xml = mpSetCountdown_(xml, sec, !auto);
        xml = auto ? mpAdvanceAfterCountdown_(xml, sec) : mpNoAutoAdvance_(xml);
      } catch (e) {
        console.warn('[MEETING] ' + g.name + 'さんのページのカウントダウンを作り直せませんでした: ' + e.message);
      }
      if (shown.indexOf(g.name) < 0) shown.push(g.name);
      paths.push(g.path);
    } else if (hidden.indexOf(g.name) < 0) {
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
  if (introShown.length) msg += (msg ? '\n' : '') + introShown.join('さん・') + 'さんの紹介のページを表示しました。';
  if (introHidden.length) msg += (msg ? '\n' : '') + introHidden.join('さん・') + 'さんの紹介のページは非表示です。';
  return { message: msg, paths: paths, auto: auto && paths.length > 0, shown: shown, hidden: hidden,
           introShown: introShown, introHidden: introHidden, introPages: introPages };
}

// 画面用：前半テンプレートに入っているアンバサダー・ディレクターのページ。
// 27MBほどのファイルを開くので、テンプレートが同じ間は結果を覚えておく。
function getWeeklyGuests() {
  try {
    var guests = templateMemo_('meetingFirst', 'guests3', function (parts) {
      var pages = weeklyGuestPages_(parts), intro = {};
      weeklyGuestIntroPages_(parts, pages).forEach(function (p) { intro[normName_(p.name)] = true; });
      return pages.map(function (g) {
        return { name: g.name, role: g.role, company: g.company || '', hasPhoto: !!g.photo,
                 hidden: g.hidden, intro: !!intro[normName_(g.name)] };
      });
    });
    return { ok: true, guests: guests };
  } catch (e) {
    console.error('[MEETING] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'アンバサダー・ディレクターのページを調べられませんでした: '
             + (e && e.message ? e.message : e), guests: [] };
  }
}

// アンバサダー・ディレクターの写真（前半テンプレートのその方のページの写真）→ { ok, photos: { 氏名: 'data:image/…;base64,…' } }。
// メインプレゼンターの画像（書記兼会計の画面）を、ディレクターが担当する週に作るため
function getWeeklyGuestPhotos() {
  try {
    var parts = unzipToMap_(getBigTemplateFile_('meetingFirst').getBlob()), photos = {};
    weeklyGuestPages_(parts).forEach(function (g) {
      if (!g.photo || photos[g.name]) return;
      var b = parts[g.photo], ext = g.photo.replace(/^.*\./, '').toLowerCase();
      var type = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : 'image/jpeg';
      photos[g.name] = 'data:' + type + ';base64,' + Utilities.base64Encode(b.getBytes());
    });
    return { ok: true, photos: photos };
  } catch (e) {
    console.error('[MEETING] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'アンバサダー・ディレクターの写真を読めませんでした: ' + (e && e.message ? e.message : e), photos: {} };
  }
}

// --- メンバープレゼンのページを前半に差し込む ---
// メンバープレゼンのテンプレートで人数ぶんのページを作り、それを前半スライドの
// 「ウィークリープレゼンテーション」の見出しページの直後へ差し込む。
// アンバサダー・ディレクター・スタートアッププレゼンの下書きのページは、差し込んだページの後ろに残る
// （アンバサダー・ディレクターは applyWeeklyGuests_ で表示を切り替える）。
function insertMemberPresen_(parts, items) {
  var anchor = weeklyAnchor_(parts);
  if (!anchor) return { message: '前半スライドに「ウィークリープレゼンテーション」の見出しページが見つかりませんでした。' };
  var file, embeddedSource = null, embeddedModels = [];
  try { file = getBigTemplateFile_(MP_TEMPLATE_KIND_); }
  catch (e) {
    if (String(e && e.message ? e.message : e).indexOf('テンプレートが未登録') < 0) throw e;
    var ov = mpFindModelSlide_(parts, [MP_OVERVIEW_.title, MP_OVERVIEW_.nextName, MP_OVERVIEW_.photo, MP_OVERVIEW_.table]);
    var iv = mpFindModelSlide_(parts, [MP_INDIVIDUAL_.photo, MP_INDIVIDUAL_.nameBox, MP_INDIVIDUAL_.company,
      MP_INDIVIDUAL_.category, MP_INDIVIDUAL_.nextName, MP_INDIVIDUAL_.nextLabel]);
    if (ov && iv && rfFindModels_(parts, WEEKLY_TITLE_).some(function (m) { return m.path === iv.path; })) {
      embeddedSource = {};
      for (var p in parts) embeddedSource[p] = parts[p].copyBlob();
      embeddedModels = [ov.path, iv.path];
    } else {
      var embeddedCache = { by: {}, seq: 0 };
      var embedded = expandPresenterSlides_(parts, items.filter(function (it) { return it.kind === 'individual'; }),
        { title: WEEKLY_TITLE_, label: 'メンバープレゼン', seconds: chapterPresenSeconds_().weekly,
          photoCache: embeddedCache,
          excludePaths: weeklyGuestPages_(parts).map(function (g) { return g.path; }) });
      if (embedded) { embedded.gone = embeddedCache.gone || []; embedded.opened = embeddedCache.opened || 0; }
      if (embedded && embedded.paths && embedded.paths.length) {
        embedded.message = '定例会テンプレート内のひな形を使いました。' + embedded.message;
      }
      return embedded || { message: 'メンバープレゼンのひな形がないため、個人ページを省略しました。' };
    }
  }
  var src = embeddedSource || unzipToMap_(file.getBlob());
  if (typeof prepareNativeMeetingTemplate_ === 'function') prepareNativeMeetingTemplate_(src);
  var built = buildMemberPresenSlides_(src, items);
  if (embeddedModels.length) setSlideEntries_(parts, slideEntries_(parts).filter(function (e) {
    return embeddedModels.indexOf(e.path) < 0;
  }));
  var sp = spliceSlides_(parts, src, slideOrder_(src), anchor);
  // auto … 個人ページがカウントダウンのあと自動で次へ進むか。業種区分の扉ページは、いつも1秒で次へ進む
  // （どちらかがあれば、ファイルの「保存済みのタイミングを使用」を入れる）
  var auto = false, flip = false;
  for (var i = 0; i < items.length; i++) { if (items[i].auto) auto = true; if (items[i].kind === 'overview') flip = true; }
  var msg = 'メンバープレゼンのページを ' + sp.paths.length + '枚 差し込みました（'
          + (auto ? '自動で次へ' : 'クリックで次へ') + (flip ? '・業種区分のページは1秒で次へ' : '') + '）。';
  if (embeddedSource) msg = '定例会テンプレート内のひな形を使いました。' + msg;
  if (built.noPhoto && built.noPhoto.length) msg += '\n写真が見つからない方: ' + built.noPhoto.join('、');
  if (sp.missingLayout.length) msg += '\n※ 同じ名前のレイアウトが無いページがありました（見た目が変わる可能性があります）。';
  return { message: msg, paths: sp.paths, auto: auto || flip, gone: built.gone || [], opened: built.opened || 0 };
}

// pptxの中身を書き換える本体。Driveの読み書きから切り離してあるので、
// 手元で同じ処理を走らせて確かめられる（tools/check_meeting_output.js）。
function editMeetingSlides_(parts, map, rules, o) {
  map = Object.assign({}, map || {});
  if (typeof prepareNativeMeetingTemplate_ === 'function') prepareNativeMeetingTemplate_(parts);
  if (map['チャプター'] === undefined && typeof chapterLabel_ === 'function') map['チャプター'] = chapterLabel_();
  if (map['開催年月日'] === undefined && o.meetingDate) {
    map['開催年月日'] = o.meetingDate.getFullYear() + '年' + (o.meetingDate.getMonth() + 1) + '月' + o.meetingDate.getDate() + '日';
  }
  var touched = 0, byPattern = 0, path;
  // 発表のページを人数ぶんに増やす（先にページを増やしてから文字を差し替える）
  //   前半 … ウィークリープレゼン（メンバープレゼンと同じ内容）
  //   後半 … リファーラル発表
  // 写真は1つの控えで取り込む。同じ方の写真は1枚だけ入れて使い回し、
  // 別の処理が同じ名前の画像を作って上書きし合うこと（写真の取り違い）も起きない。
  var photoCache = { by: {}, seq: 0 };
  var visitors = typeof expandNativeVisitorPages_ === 'function' ? expandNativeVisitorPages_(parts, o.meetingDate) : null;
  // メインプレゼン・抽選のディレクター（名簿に無く、前半テンプレートにアンバサダー・ディレクターのページがある方）。
  // ページを差し込む・写す前に探す（差し込んだメンバーのページも同じ作りのため）
  var directors = twoPersonDirectors_(parts, o);
  // 後半：1枚目に、前半の最後のページ（メインプレゼンテーション）を入れる（前半のテンプレートから写す）。
  // お名前・会社名・カテゴリーは差し込み口 {{メインプレゼン1氏名}} など、写真はあとのお2人の写真の処理で入る
  var mainPage = o.mainPageSource ? insertMainPresenPage_(parts, o.mainPageSource)
               : (o.mainPageNote ? { message: o.mainPageNote } : null);
  // コアバリューのページは、差し込み・書き換えの前に、テンプレートの字で見分けておく（メンバーの会社名などに英語の語が
  // あっても、そのページをコアバリューのページとして非表示にしないように）
  var corePages = o.coreValue ? coreValuePages_(parts) : null;
  // 前半：役職のメンバー紹介（リーダーシップチーム・コーディネーター・チームのページ）を、その期の方にする。
  // 画面で外したとき・その期の役職が読めないときは、差し込み口を空にするだけ（role_intro_srv.js）
  var roles = applyRoleIntro_(parts, o.roleIntro === false ? null : (o.roleIntroData || null), photoCache);
  // 前半：新メンバー・更新メンバー（1人1枚）／バイスプレジデントによる報告の数／ネットワーキングリーダー
  // （どれもルーティンチェックシートの内容を画面で確かめたもの。渡されなければテンプレートのまま。meeting_pages_srv.js）
  var members = o.memberPages ? applyMemberPages_(parts, o.memberPages, photoCache) : null;
  var vp = o.vpReport ? applyVpReport_(parts, o.vpReport, o.meetingDate || null) : null;
  var leaders = o.networkingLeaders ? applyNetworkingLeaders_(parts, o.networkingLeaders, photoCache) : null;
  var referral = (o.referral && o.referral.length)
    ? expandPresenterSlides_(parts, o.referral,
        { title: RF_TITLE_, label: 'リファーラル発表', seconds: chapterPresenSeconds_().referral, photoCache: photoCache }) : null;
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
  // 後半：リファーラル発表のあと、推薦のことばのページに入ったところで音楽を止める（ページの並びが決まってから）
  var recoPath = (reco && reco.paths && reco.paths[0]) || findSlideWithText_(parts, RECO_PREFIX_ + '1氏名');
  var musicStop = recoPath ? stopMusicAt_(parts, recoPath) : null;
  // 書記兼会計による報告（更新状況一覧）
  var renewal = applyRenewalStatus_(parts, map);
  var photoMsgs = [], TWO = [
    { prefix: 'メインプレゼン', names: o.mainPresenters },
    { prefix: '推薦のことば',   names: reco ? null : o.recommenders },
    { prefix: '抽選',           names: o.lottery }
  ];
  // ディレクター：会社名・カテゴリーが画面から来なければ、その方のページの文字で埋める
  map = twoPersonDirectorValues_(map || {}, TWO, directors);
  // お2人のページの会社名は、名簿の書き方にそろえてから測り、同じ文字を入れる（後半の1枚目のメインプレゼンのページも）
  map = twoPersonCompanies_(map, TWO.map(function (t) { return t.prefix; }));
  // ディレクターの写真は、名簿の写真より先に、その方のページの写真を使う
  var guestPhotoOf = function (name) { return twoPersonGuestPhoto_(parts, directors[normName_(name)]); };
  for (var w = 0; w < TWO.length; w++) {
    // メインプレゼン・抽選は1人の週がある（もう1人の欄を消し、まん中に置く）。推薦のことばはいつも2人
    var ph = applyTwoPersonPhotos_(parts, TWO[w].prefix, TWO[w].names || [], photoCache,
                                   { single: TWO[w].prefix !== RECO_PREFIX_, photoOf: guestPhotoOf });
    if (ph && ph.message) photoMsgs.push(ph.message);
    fitTwoPersonPage_(parts, TWO[w].prefix, map);      // 長いお名前・会社名が下の段に重ならないように
  }
  var photos = { message: photoMsgs.join('\n') };
  var stats = (typeof applyBniStats_ === 'function') ? applyBniStats_(parts, o.bniStats) : null;
  var statsTouched = {};
  if (stats) stats.paths.forEach(function(p) { statsTouched[p] = true; touched++; });
  for (path in parts) {
    if (!/^ppt\/(slides|notesSlides)\/[^\/]+\.xml$/.test(path)) continue;   // 本文だけ
    var xml = xmlOf_(parts, path), before = xml;
    if (!xml) continue;
    if (xml.indexOf('{{') !== -1) xml = replaceTokensInXml_(xml, map);
    if (rules && rules.length) {
      var pr = replacePatternsInXml_(xml, rules);
      if (pr.changed) { xml = pr.xml; byPattern += pr.changed; }
    }
    if (xml !== before) { putXml_(parts, path, xml); if (!statsTouched[path]) touched++; }
  }
  // 後半の1枚目に写したメインプレゼンのページ：後半で入れる値の無い差し込み口（前半だけのもの）は空にする
  if (mainPage && mainPage.path) {
    var mx = xmlOf_(parts, mainPage.path), left = mx ? Object.keys(listTokensInXml_(mx)) : [];
    if (left.length) {
      var blank = {};
      left.forEach(function (t) { blank[t] = ''; });
      putXml_(parts, mainPage.path, replaceTokensInXml_(mx, blank));
    }
  }
  // 前半：スピーカーローテーションの表（以前は書記兼会計が作った画像）。
  // 「第○回」の書き換えのあとに入れる（表の中の先の回の番号まで、今回の番号にしないように）
  var rotation = o.speakerRotation ? applySpeakerRotation_(parts, o.speakerRotation) : null;
  var core = o.coreValue ? applyCoreValue_(parts, o.coreValue, corePages) : null;
  var policy = o.generalPolicy ? applyGeneralPolicy_(parts, o.generalPolicy) : null;
  var audio = o.music ? applyMeetingAudio_(parts, o.music) : null;
  // 写真の索引にあるのに開けなかった方（写真なしで作った）。1枚も開けないときは、写真の無いスライドで
  // 同じ開催日のスライドを置き換えないよう止める（写真のフォルダを見る権限が無い・索引が古いなど）
  var gone = (photoCache.gone || []).slice();
  ((weekly && weekly.gone) || []).forEach(function (n) { if (gone.indexOf(n) < 0) gone.push(n); });
  if (gone.length && !((photoCache.opened || 0) + ((weekly && weekly.opened) || 0))) {
    throw new Error('メンバーの写真を1枚も開けませんでした（' + gone.length + '名）。写真のフォルダ（02_メンバー写真）を見る権限があるかを確かめ、'
      + '「⚙️ 設定」＞「メンバー写真」で「索引を作り直す」を押してから、もう一度お試しください。');
  }
  return { touched: touched, byPattern: byPattern, core: core, policy: policy, photoGone: gone,
           photos: photos, audio: audio, referral: referral, weekly: weekly, guests: guests,
           reco: reco, musicStop: musicStop, renewal: renewal, rotation: rotation, roles: roles,
           members: members, vp: vp, leaders: leaders, mainPage: mainPage, visitors: visitors, stats: stats };
}

// 後半の1枚目に、メインプレゼンのページ（前半の最後のページ。差し込み口 {{メインプレゼン1氏名}} のあるページ）を入れる。
// 前半のテンプレート（src）から、見た目（レイアウト・マスター）ごと写す。後半のテンプレートにもうあれば、そのまま使う
//   戻り値 { path: 入れたページ, message }
function insertMainPresenPage_(parts, src) {
  var mark = '{{メインプレゼン1氏名}}';
  if (findSlideWithText_(parts, mark)) return { message: '' };
  var model = findSlideWithText_(src, mark);
  if (!model) {
    return { message: '前半のテンプレートに、メインプレゼンのページ（差し込み口 {{メインプレゼン1氏名}} のあるページ）が見つからないため、'
      + '後半の1枚目には入れませんでした。' };
  }
  fitSourceToTargetSize_(parts, src);                         // ページの大きさが違えば、後半の大きさに合わせる
  var made = importSlideCopies_(parts, src, model, 1).slides[0];
  setSlideEntries_(parts, [made].concat(slideEntries_(parts).filter(function (e) { return e.path !== made.path; })));
  syncSections_(parts);                                       // PowerPoint のセクションも、並びに合わせる
  return { path: made.path, message: '後半の1枚目に、メインプレゼンのページ（前半のテンプレートの最後のページ）を入れました。' };
}

// ファイル名の（ ）の中
var MEETING_FILE_PARTS_ = { meetingFirst: '前半', meetingSecond: '後半' };

// 前半のファイルに後半も同梱されている場合は、同じ開催日の材料を補う。
// 画面や呼び出し元から指定された値は優先する。
function combinedMeetingOptions_(parts, map, o, date) {
  var hasSecond = rfFindModels_(parts, RF_TITLE_).length > 0;
  slideOrder_(parts).forEach(function (p) {
    var xml = xmlOf_(parts, p) || '';
    if (/\{\{(?:推薦のことば|抽選|更新90|更新60|更新30|更新超過)/.test(xml)
        || slideText_(xml).indexOf('更新を迎えるメンバー') >= 0) hasSecond = true;
  });
  if (!hasSecond) return;
  var routine = getRoutineInfo(date);
  if (!routine || !routine.ok) throw new Error('同梱の後半部分に使うルーティンチェックシートを読み取れませんでした。');
  var members = getMemberMaster().members || [];
  var person = function (rec) {
    var name = rec && (rec.name || rec.raw) || '';
    var m = members.filter(function (m) { return normName_(m.name) === normName_(name); })[0] || {};
    return { name: name, company: m.company || '', category: m.title || '' };
  };
  if (o.recommendPairs === undefined) o.recommendPairs = (routine.recommendations || [])
    .filter(function (p) { return p.when !== 'later'; })
    .map(function (p) { return { giver: person(p.giver), receiver: person(p.receiver), after: p.when === 'after' }; });
  if (o.lottery === undefined) o.lottery = (routine.mainPresenters || [])
    .filter(function (p) { return p.kind !== 'director' && p.matched; }).map(function (p) { return p.name; });
  o.lottery.forEach(function (name, i) {
    var p = person({ name: name }), prefix = '抽選' + (i + 1);
    if (map[prefix + '氏名'] === undefined) map[prefix + '氏名'] = p.name;
    if (map[prefix + '会社名'] === undefined) map[prefix + '会社名'] = p.company;
    if (map[prefix + 'カテゴリー'] === undefined) map[prefix + 'カテゴリー'] = roleInBrackets_(p.category);
  });
  var lists = computeRenewalLists(date);
  if (!lists.ok) throw new Error(lists.message || '更新状況を読み取れませんでした。');
  ['90', '60', '30', '超過'].forEach(function (k, i) {
    var key = '更新' + k;
    if (map[key] === undefined) map[key] = lists.text[['d90', 'd60', 'd30', 'overdue'][i]];
  });
  if (o.referral === undefined && rfFindModels_(parts, RF_TITLE_).length) {
    var skip = {};
    // 画面で選び直した欠席の方を優先する。
    (o.absentNames || ((routine.absentees || {}).names) || []).forEach(function (n) { skip[normName_(n)] = true; });
    var list = members.filter(function (m) { return m.name && !skip[normName_(m.name)]; })
      .sort(function (a, b) { return (parseFloat(a.no) || 9999) - (parseFloat(b.no) || 9999); });
    o.referral = list.map(function (m, i) {
      return { name: m.name, photoName: m.name, auto: false, seconds: chapterPresenSeconds_().referral,
        companyLines: [CB_.normalize(m.company || '')], categoryLines: [m.title || ''],
        nextName: list[i + 1] ? list[i + 1].name : '' };
    });
  }
}

// 定例会スライドを生成する。テンプレート内の {{キー}} を置換する方式。
// pptxのままサーバー側で加工し、Driveへ保存してURLを返す。
// opts: { patterns: true/false（差し込み口が無いページの第○回・日付も直す）,
//         coreValue: 'Givers Gain' など,
//         memberPresen: [...]（前半に差し込むメンバープレゼンのページ）,
//         weeklyGuests: ['坂上　達彦', …]（表示にするアンバサダー・ディレクター。null なら触らない）,
//         weeklyAuto: true/false（ウィークリープレゼンを自動で次へ進めるか）,
//         roleIntro: true/false（前半の役職のメンバー紹介を、その開催日の期の役職・チームにするか。既定は true）,
//         memberPages: { newMembers: [{ name, raw, category }], renewMembers: [{ name, raw, category, years }] },
//         vpReport: { avg, month, count, perWeek, from, to, total, thanks, weekCount, weekExt },
//         networkingLeaders: { show, month, items: [{ key, value, unit, winners: [{ name, raw, category }] }] } }
function generateMeetingSlides(kind, values, meetingDateVal, opts) {
  try {
    if (!BIG_TEMPLATE_KINDS_[kind]) return { ok: false, message: 'スライドの種類が不正です。' };
    var map = Object.assign({}, values || {}), o = Object.assign({}, opts || {});
    var d = parseDate_(meetingDateVal);
    var label = BIG_TEMPLATE_KINDS_[kind].label;
    // 「定例会20260930_（前半）09301412.pptx」（開催日と、作った日時）
    var outName = slideFileName_(d, MEETING_FILE_PARTS_[kind] || label);

    var rules = (o.patterns === false) ? []
              : meetingPatternRules_(String(map['開催回'] || '').replace(/[^\d]/g, ''), d);
    o.meetingDate = d;            // 表紙・「○/○速報」の日付
    if (typeof bniStatsForDate_ === 'function') o.bniStats = bniStatsForDate_(d);
    // 前半の役職のメンバー紹介：その開催日の期の「役職・チーム（半期ごと）」
    if (kind === 'meetingFirst' && o.roleIntro !== false && !o.roleIntroData) {
      try { o.roleIntroData = riDataOfDate_(d || new Date()); }
      catch (e) { console.warn('[MEETING] 役職・チームを読めませんでした: ' + (e && e.message ? e.message : e)); }
    }
    // 後半の推薦のことば：組に付いた受け取ったスライド（画像）を開いておく（reco_slide_srv.js）
    var recoGone = (kind === 'meetingSecond' && o.recommendPairs && o.recommendPairs.length) ? recoLoadSlides_(o.recommendPairs) : '';
    // 後半の1枚目のメインプレゼンのページは、前半のテンプレートから写す（前半のテンプレートを開いておく）
    if (kind === 'meetingSecond' && o.mainPage) {
      try { o.mainPageSource = bigTemplateParts_('meetingFirst'); }
      catch (e) {
        o.mainPageNote = '後半の1枚目にメインプレゼンのページを入れられませんでした（前半のテンプレートを開けません: '
          + (e && e.message ? e.message : e) + '）。';
      }
    }
    // 後半で、1枚目のページは入れずに抽選だけにディレクター（名簿に無い方）を選んだときも、
    // その方の写真・会社名・カテゴリーは前半のテンプレートのその方のページから入れる（前半のテンプレートを開いておく）
    if (kind === 'meetingSecond' && !o.mainPageSource && twoPersonNonMembers_(o.lottery, o).length) {
      try { o.guestSource = bigTemplateParts_('meetingFirst'); }
      catch (e) { console.warn('[MEETING] 前半のテンプレートを開けません（抽選のディレクターの写真）: ' + (e && e.message ? e.message : e)); }
    }

    var r = editPptxOnServer_(kind, outName, function (parts) {
      if (typeof prepareNativeMeetingTemplate_ === 'function') prepareNativeMeetingTemplate_(parts);
      if (kind === 'meetingFirst') combinedMeetingOptions_(parts, map, o, d);
      return editMeetingSlides_(parts, map, rules, o);
    });

    var info = r.info || {};
    var msg = '「' + label + '」を作成しました（' + r.partCount + 'パーツ／' +
              (r.timing.合計 / 1000).toFixed(1) + '秒）。';
    if (!info.touched && !(info.stats && info.stats.updated.length)) {
      msg += '\n※ 書き換える箇所が見つかりませんでした（{{ }} も「第○回」も見つかりません）。';
    } else if (info.touched) {
      msg += '\n' + info.touched + '枚のスライドを書き換えました';
      msg += info.byPattern ? '（うち「第○回・日付」を直した段落 ' + info.byPattern + 'か所）。' : '。';
    }
    if (info.core) msg += '\n' + info.core.message;
    if (info.policy) msg += '\n' + info.policy.message;
    if (info.photos && info.photos.message) msg += '\n' + info.photos.message;
    if (info.photoGone && info.photoGone.length) {
      msg += '\n⚠ 写真の索引にあるのに、写真のファイルを開けなかった方（写真なしで作りました）: ' + info.photoGone.join('、')
           + '。「⚙️ 設定」＞「メンバー写真」で「索引を作り直す」を押してください。';
    }
    if (info.audio && info.audio.message) msg += '\n' + info.audio.message;
    if (info.referral && info.referral.message) msg += '\n' + info.referral.message;
    if (info.weekly && info.weekly.message) msg += '\n' + info.weekly.message;
    if (info.guests && info.guests.message) msg += '\n' + info.guests.message;
    if (info.reco && info.reco.message) msg += '\n' + info.reco.message;
    if (recoGone) msg += '\n⚠ ' + recoGone;
    if (info.musicStop && info.musicStop.message) msg += '\n' + info.musicStop.message;
    if (info.renewal && info.renewal.message) msg += '\n' + info.renewal.message;
    if (info.rotation && info.rotation.message) msg += '\n' + info.rotation.message;
    if (info.roles && info.roles.message) msg += '\n' + info.roles.message;
    if (info.members && info.members.message) msg += '\n' + info.members.message;
    if (info.vp && info.vp.message) msg += '\n' + info.vp.message;
    if (info.leaders && info.leaders.message) msg += '\n' + info.leaders.message;
    if (info.stats && info.stats.message) msg += '\n' + info.stats.message;
    if (info.mainPage && info.mainPage.message) msg += '\n' + info.mainPage.message;
    if (info.visitors && info.visitors.message) msg += '\n' + info.visitors.message;
    return { ok: true, message: msg, url: r.saved.url, downloadUrl: r.saved.downloadUrl,
             fileName: outName, touched: info.touched, core: info.core, policy: info.policy,
             timing: r.timing };
  } catch (e) {
    console.error('[MEETING] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'スライドの作成に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}
