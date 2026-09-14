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
function computeRenewalLists(meetingDateVal) {
  try {
    var base = parseDate_(meetingDateVal);
    if (!base) return { ok: false, message: '開催日を解釈できませんでした。' };
    var members = (getMemberMaster().members || []);
    var newMembers = [], renewMembers = [], d90 = [], d60 = [], d30 = [], overdue = [], noDate = [];

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
      var left = daysBetween_(base, exp);
      var rec = { name: m.name, date: fmtDate_(exp), left: left };
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
    return { ok: true, meetings: candidates, defaultMeeting: first,
             lists: lists.ok ? lists : null, stats: stats, templates: ready,
             memberCount: members.length,
             memberNames: members.map(function (m) { return m.name; }) };
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

// 定例会スライドを生成する。テンプレート内の {{キー}} を置換する方式。
// pptxのままサーバー側で加工し、Driveへ保存してURLを返す。
function generateMeetingSlides(kind, values, meetingDateVal) {
  try {
    if (!BIG_TEMPLATE_KINDS_[kind]) return { ok: false, message: 'スライドの種類が不正です。' };
    var map = values || {};
    var mmdd = '';
    var d = parseDate_(meetingDateVal);
    if (d) mmdd = Utilities.formatDate(d, 'Asia/Tokyo', 'MMdd');
    var label = BIG_TEMPLATE_KINDS_[kind].label;
    var outName = (mmdd ? mmdd + '_' : '') + label + '.pptx';

    var replaced = 0, touched = 0;
    var r = editPptxOnServer_(kind, outName, function (parts) {
      for (var path in parts) {
        if (!/^ppt\/(slides|notesSlides)\/[^\/]+\.xml$/.test(path)) continue;   // 本文だけ
        var xml = xmlOf_(parts, path);
        if (!xml || xml.indexOf('{{') === -1) continue;
        var before = xml;
        xml = replaceTokensInXml_(xml, map);
        if (xml !== before) { putXml_(parts, path, xml); touched++; replaced++; }
      }
      return { touched: touched };
    });

    var msg = '「' + label + '」を作成しました（' + r.partCount + 'パーツ／' +
              (r.timing.合計 / 1000).toFixed(1) + '秒）。';
    if (!touched) {
      msg += '\n※ テンプレート内に {{ }} の差し込み口が見つかりませんでした。文字は差し替わっていません。';
    } else {
      msg += '\n' + touched + '枚のスライドを書き換えました。';
    }
    return { ok: true, message: msg, url: r.saved.url, fileName: outName, touched: touched, timing: r.timing };
  } catch (e) {
    console.error('[MEETING] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'スライドの作成に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}
