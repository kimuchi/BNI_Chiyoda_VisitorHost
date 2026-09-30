// === 役職のメンバー紹介（定例会スライド・前半）===
// 前半スライドの「リーダーシップチーム」「各コーディネーター」「メンバーシップ委員会」「ビジターホストチーム」
// 「サポートチーム」のページに、その開催日の期の「役職・チーム（半期ごと）」の方を入れる
// （お名前・写真・会社名・カテゴリー・ローマ字）。
//
// 差し込み口（雛形に書いておく。公式ファイルから作った雛形には入っている）
//   役職の担当者     … {{プレジデント氏名}} {{プレジデント会社名}} {{プレジデントカテゴリー}} {{プレジデントローマ字}}
//                      （役職の名前は「役職・チーム（半期ごと）」の13の役職の名前）
//                      {{プレジデントカテゴリー（）}} は「（カテゴリー）」（中の（）は「」にする。カテゴリーが空なら何も出さない）
//   チームのメンバー … {{メンバーシップ委員会1氏名}}（1人目。会社名・カテゴリー・ローマ字も同じ）、
//                      {{ビジターホストの一覧}}（お名前を「、」でつなぐ）。チームの名前は既定の名前
//   写真の枠         … 画像の図形の名前を {{プレジデント写真}} {{メンバーシップ委員会1写真}} にする
//                      （PowerPoint の「オブジェクトの選択と表示」で付けられる）
// 差し込み口の無いページ（Activeチャプターの雛形など）は、ページの作りから枠を見分ける（riDetectSlots_）。
// そのときは、お名前が替わった枠だけ書き換える（同じ方のままの枠は、写真の切り抜きやローマ字もそのまま残す）。
// 担当者の居ない役職の1人のページ（プレジデントのページなど）は非表示にする。
// 担当者・チームが登録されていない期は、差し込み口の無いページには触らない（テンプレートのまま）。

// ページの文字から役職を見分ける書き方（全角半角・空白・大文字小文字は無視。長いものから当てる）
var RI_ROLE_WORDS_ = [
  ['president', ['チャプタープレジデント', 'プレジデント', 'President', 'Chapter President']],
  ['vice', ['バイスプレジデント', 'Vice President']],
  ['secretary', ['書記兼会計', 'Secretary & Treasurer', 'Secretary&Treasurer', 'Secretary/Treasurer']],
  ['vhc', ['ビジターホストコーディネーター', 'Visitor Host Coordinator']],
  ['ec', ['エデュケーションコーディネーター', 'Education Coordinator']],
  ['mentor', ['メンターコーディネーター', 'Mentor Coordinator']],
  ['web', ['webマスター', 'ウェブマスター']],
  ['support', ['メンバーサポート委員長', 'メンバーサポート委員']],
  ['training', ['トレーニング促進委員長', 'トレーニング委員長', 'トレーニング促進委員', 'トレーニング委員']],
  ['event', ['イベント委員＆1to1促進委員', 'イベント委員長', 'イベント委員', '1to1委員長']],
  ['bcp', ['BCP委員長', 'BCP委員']],
  ['spreading', ['スプレディング委員長', 'Spreading委員長', 'スプレディング委員', 'Spreading委員']],
  ['gbc', ['グローバルビジネスコーディネーター']]
];
// 表の「コーディネーター」の行の短い書き方（「エデュケーション｜氏名」のように、見出しとお名前が同じ行）
var RI_ROW_WORDS_ = [['ec', ['エデュケーション']], ['mentor', ['メンター']], ['vhc', ['ビジターホスト']]];
// チームの見出し
var RI_TEAM_WORDS_ = [['membership', ['メンバーシップ委員会']], ['role:vhc', ['ビジターホストチーム', 'ビジターホスト']]];
var RI_TEAM_MAX_ = 40;                                 // チームの差し込み口の番号（1〜40）
var RI_FIELDS_ = ['氏名', '会社名', 'カテゴリー', 'ローマ字'];

function riNorm_(s) { return String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]/g, '').toLowerCase(); }

// 文字 → 役職（またはチーム）のキー。contains … 含んでいればよい（サポートチームのページの見出し）
function riWordOf_(text, table, contains) {
  var t = riNorm_(text), best = '', bestLen = 0;
  if (!t) return '';
  for (var i = 0; i < table.length; i++) {
    for (var j = 0; j < table[i][1].length; j++) {
      var w = riNorm_(table[i][1][j]);
      if ((contains ? t.indexOf(w) >= 0 : t === w) && w.length > bestLen) { best = table[i][0]; bestLen = w.length; }
    }
  }
  return best;
}

// お名前らしい文字か（「氏名」「First & last name」の見本も含める）
function riNameLike_(text) {
  var s = String(text == null ? '' : text).normalize('NFKC').replace(/\s+/g, ' ').trim();
  if (!s) return false;
  if (/^(氏名|お名前|名前|first ?& ?last name)$/i.test(s)) return true;
  if (s.length > 20 || /[0-9()（）【】「」:：,、。@＆&]/.test(s)) return false;
  if (/(委員|コーディネーター|マスター|チーム|プレジデント|会計|リーダー|担当)/.test(s)) return false;
  if (/[぀-ヿ㐀-鿿豈-﫿々]/.test(s)) return /^[A-Za-z぀-ヿ㐀-鿿豈-﫿々〆ヶ・. '-]+$/.test(s);
  return /^[A-Za-z][A-Za-z.'-]*( [A-Za-z][A-Za-z.'-]*)+$/.test(s);
}
// 「氏名, 氏名, 氏名」のように、お名前を並べた1つのマス目か
function riNameList_(text) {
  var parts = String(text == null ? '' : text).split(/[,，、]/).map(function (x) { return x.trim(); }).filter(Boolean);
  return parts.length >= 2 && parts.every(riNameLike_);
}

// --- ふりがな → ローマ字（パスポートの書き方。名・姓の順、大文字）---
var RI_KANA_ = {
  'きゃ': 'kya', 'きゅ': 'kyu', 'きょ': 'kyo', 'しゃ': 'sha', 'しゅ': 'shu', 'しょ': 'sho', 'ちゃ': 'cha', 'ちゅ': 'chu',
  'ちょ': 'cho', 'にゃ': 'nya', 'にゅ': 'nyu', 'にょ': 'nyo', 'ひゃ': 'hya', 'ひゅ': 'hyu', 'ひょ': 'hyo', 'みゃ': 'mya',
  'みゅ': 'myu', 'みょ': 'myo', 'りゃ': 'rya', 'りゅ': 'ryu', 'りょ': 'ryo', 'ぎゃ': 'gya', 'ぎゅ': 'gyu', 'ぎょ': 'gyo',
  'じゃ': 'ja', 'じゅ': 'ju', 'じょ': 'jo', 'ぢゃ': 'ja', 'ぢゅ': 'ju', 'ぢょ': 'jo', 'びゃ': 'bya', 'びゅ': 'byu',
  'びょ': 'byo', 'ぴゃ': 'pya', 'ぴゅ': 'pyu', 'ぴょ': 'pyo', 'しぇ': 'she', 'じぇ': 'je', 'ちぇ': 'che', 'てぃ': 'ti',
  'でぃ': 'di', 'ふぁ': 'fa', 'ふぃ': 'fi', 'ふぇ': 'fe', 'ふぉ': 'fo', 'うぃ': 'wi', 'うぇ': 'we', 'うぉ': 'wo',
  'あ': 'a', 'い': 'i', 'う': 'u', 'え': 'e', 'お': 'o', 'か': 'ka', 'き': 'ki', 'く': 'ku', 'け': 'ke', 'こ': 'ko',
  'さ': 'sa', 'し': 'shi', 'す': 'su', 'せ': 'se', 'そ': 'so', 'た': 'ta', 'ち': 'chi', 'つ': 'tsu', 'て': 'te', 'と': 'to',
  'な': 'na', 'に': 'ni', 'ぬ': 'nu', 'ね': 'ne', 'の': 'no', 'は': 'ha', 'ひ': 'hi', 'ふ': 'fu', 'へ': 'he', 'ほ': 'ho',
  'ま': 'ma', 'み': 'mi', 'む': 'mu', 'め': 'me', 'も': 'mo', 'や': 'ya', 'ゆ': 'yu', 'よ': 'yo',
  'ら': 'ra', 'り': 'ri', 'る': 'ru', 'れ': 're', 'ろ': 'ro', 'わ': 'wa', 'ゐ': 'i', 'ゑ': 'e', 'を': 'o', 'ん': 'n',
  'が': 'ga', 'ぎ': 'gi', 'ぐ': 'gu', 'げ': 'ge', 'ご': 'go', 'ざ': 'za', 'じ': 'ji', 'ず': 'zu', 'ぜ': 'ze', 'ぞ': 'zo',
  'だ': 'da', 'ぢ': 'ji', 'づ': 'zu', 'で': 'de', 'ど': 'do', 'ば': 'ba', 'び': 'bi', 'ぶ': 'bu', 'べ': 'be', 'ぼ': 'bo',
  'ぱ': 'pa', 'ぴ': 'pi', 'ぷ': 'pu', 'ぺ': 'pe', 'ぽ': 'po', 'ゔ': 'vu',
  'ぁ': 'a', 'ぃ': 'i', 'ぅ': 'u', 'ぇ': 'e', 'ぉ': 'o', 'ゃ': 'ya', 'ゅ': 'yu', 'ょ': 'yo', 'ゎ': 'wa'
};
function riRomaji_(kana) {
  var s = String(kana == null ? '' : kana).normalize('NFKC').trim();
  if (!s) return '';
  s = s.replace(/[ァ-ヶ]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0x60); });   // カタカナ → ひらがな
  var words = s.split(/\s+/).map(riKanaWord_);
  for (var i = 0; i < words.length; i++) if (!words[i]) return '';                                          // 読めない文字がある
  if (words.length >= 2) words = words.slice(1).concat([words[0]]);                                          // 名・姓の順
  return words.join(' ').toUpperCase();
}
function riKanaWord_(w) {
  var out = '', i = 0, sokuon = false;
  while (i < w.length) {
    var two = w.substr(i, 2), one = w.charAt(i), r;
    if (RI_KANA_[two] && two.length === 2) { r = RI_KANA_[two]; i += 2; }
    else if (one === 'っ') { sokuon = true; i++; continue; }
    else if (one === 'ー' || one === '-') { i++; continue; }                                                  // 長音は書かない
    else if (RI_KANA_[one]) { r = RI_KANA_[one]; i++; }
    else if (/[a-z]/i.test(one)) { r = one.toLowerCase(); i++; }
    else return '';
    if (sokuon) { out += r.indexOf('ch') === 0 ? 't' : r.charAt(0); sokuon = false; }
    out += r;
  }
  // 長い音は書かない（おう・おお → o、うう → u）。B・M・P の前の「ん」は m
  return out.replace(/ou/g, 'o').replace(/oo/g, 'o').replace(/uu/g, 'u').replace(/n(?=[bmp])/g, 'm');
}

// --- その期の方 ---
// 名簿の方 → { name, company, category, romaji }（名簿に居ない方は、お名前だけ）
function riPerson_(name, byName) {
  var nm = String(name == null ? '' : name).trim();
  if (!nm) return null;
  var m = byName[normName_(nm)] || {};
  return { name: m.name || nm, company: m.company || '', category: m.title || m.category || '', romaji: riRomaji_(m.kana || '') };
}
// holders … { president: '氏名', … }、teams … [{ key, name, members: [{ name }] }]（その期に登録してあれば）、
// members … メンバー名簿。holdersOk / teamsOk … 使えるか（どちらも無ければ、差し込み口の無いページには触らない）
function riDataFrom_(term, label, holders, holdersFrom, teams, teamsRegistered, members) {
  var byName = {};
  (members || []).forEach(function (m) { byName[normName_(m.name)] = m; });
  var roles = {}, anyHolder = false;
  ROLE_DEFS_.forEach(function (r) {
    roles[r.key] = holders && holders[r.key] ? riPerson_(holders[r.key], byName) : null;
    if (roles[r.key]) anyHolder = true;
  });
  var tms = {}, names = {}, anyMember = false;
  (teams || []).forEach(function (t) {
    tms[t.key] = (t.members || []).map(function (m) { return riPerson_(m.name, byName); }).filter(Boolean);
    names[t.key] = t.name;
    if (tms[t.key].length) anyMember = true;
  });
  return { term: term, label: label, roles: roles, teams: tms, teamNames: names, holdersFrom: holdersFrom,
           holdersOk: holdersFrom != null && anyHolder, teamsOk: !!teamsRegistered && anyMember };
}
// その開催日の期
function riDataOfDate_(date) {
  var term = roleTermOf_(date), h = roleHoldersOfTerm_(roleHolderTerms_(), term), tm = roleTeamsOfTerm_(roleTeamTerms_(), term);
  return riDataFrom_(term, roleTermLabel_(term), h.holders, h.from, tm.teams, tm.registered, getMemberMaster().members || []);
}

// 画面（定例会スライド（前半））に出す、入れる方の一覧
function getRoleIntroPreview(dateStr) {
  try {
    var d = parseDate_(dateStr || '') || new Date(), data = riDataOfDate_(d);
    return { ok: true, term: data.term, label: data.label, holdersOk: data.holdersOk, holdersFrom: data.holdersFrom,
             teamsOk: data.teamsOk,
             holders: ROLE_DEFS_.map(function (r) { return { label: r.label, name: data.roles[r.key] ? data.roles[r.key].name : '' }; }),
             teams: Object.keys(data.teams).filter(function (k) { return data.teams[k].length; })
               .map(function (k) { return { name: data.teamNames[k] || k, count: data.teams[k].length }; }) };
  } catch (e) {
    return { ok: false, message: '役職・チームを読めませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// --- 差し込み口の名前 ---
function riRoleBase_(key) { var r = roleDefOf_(key); return r ? r.label : ''; }
function riTeamBase_(key, data) {
  var d = roleTeamDefOf_(key);
  return d ? d.name : ((data && data.teamNames && data.teamNames[key]) || '');
}
// 差し込み口 → 値、写真の枠の名前 → 方。
// 写真の枠は、担当者（チーム）が登録してある期のときだけ入れ替える（未登録なら、雛形の写真の枠をそのまま残す）
function riValues_(data) {
  var map = {}, persons = {};
  var put = function (base, p, usable) {
    if (usable) persons[base] = p;
    var c = p ? p.category : '', inner = c.replace(/[（(]/g, '「').replace(/[）)]/g, '」');
    map[base + '氏名'] = p ? p.name : ''; map[base + '会社名'] = p ? p.company : '';
    map[base + 'カテゴリー'] = c; map[base + 'ローマ字'] = p ? p.romaji : '';
    map[base + 'カテゴリー（）'] = c ? '（' + inner + '）' : '';
    map[base + 'カテゴリー()'] = c ? '(' + inner + ')' : '';
  };
  var hOk = !!(data && data.holdersOk), tOk = !!(data && data.teamsOk);
  ROLE_DEFS_.forEach(function (r) { put(r.label, hOk ? data.roles[r.key] : null, hOk); });
  var keys = ROLE_TEAM_DEFS_.map(function (d) { return d.key; });
  if (data) Object.keys(data.teams).forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k); });
  keys.forEach(function (k) {
    var base = riTeamBase_(k, data), list = tOk ? (data.teams[k] || []) : [];
    if (!base) return;
    for (var n = 1; n <= Math.max(RI_TEAM_MAX_, list.length); n++) put(base + n, list[n - 1] || null, tOk);
    map[base + 'の一覧'] = list.map(function (p) { return p.name; }).join('、');
  });
  return { map: map, persons: persons };
}

// --- ページの図形（グループの中も。スライド上の位置に直す）---
// 入れ子のグループも数える（findTagRanges_ は外側のグループの中を飛ばすため）
function riAllRanges_(xml, tag) {
  var out = [], stack = [], re = new RegExp('<(\\/?)' + tag.replace(':', '\\:') + '(?=[\\s>/])', 'g'), m;
  while ((m = re.exec(xml)) !== null) {
    var gt = xml.indexOf('>', m.index);
    if (gt < 0) break;
    if (m[1] === '/') { var st = stack.pop(); if (st != null) out.push({ start: st, end: gt + 1 }); }
    else if (xml.charAt(gt - 1) === '/') out.push({ start: m.index, end: gt + 1 });
    else stack.push(m.index);
  }
  return out.sort(function (a, b) { return a.start - b.start; });
}
function riShapes_(xml) {
  var list = [];
  ['p:grpSp', 'p:sp', 'p:pic', 'p:graphicFrame'].forEach(function (tag) {
    riAllRanges_(xml, tag).forEach(function (r) {
      var seg = xml.substring(r.start, r.end), nv = seg.match(/<p:cNvPr\b[^>]*>/);
      if (!nv) return;
      var off = seg.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>/), ext = seg.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
      var s = { tag: tag, id: (nv[0].match(/\sid="(\d+)"/) || [])[1], name: unescapeXml_((nv[0].match(/\sname="([^"]*)"/) || [])[1] || ''),
                start: r.start, end: r.end, x: off ? +off[1] : 0, y: off ? +off[2] : 0, cx: ext ? +ext[1] : 0, cy: ext ? +ext[2] : 0 };
      if (tag === 'p:grpSp') {
        var ch = seg.match(/<a:chOff\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>\s*<a:chExt\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
        s.ch = ch ? { x: +ch[1], y: +ch[2], cx: +ch[3], cy: +ch[4] } : { x: s.x, y: s.y, cx: s.cx, cy: s.cy };
      }
      if (tag === 'p:sp') {
        s.paras = findTagRanges_(seg, 'a:p').map(function (p) { return slideText_(seg.substring(p.start, p.end)); });
        s.lines = [];
        findTagRanges_(seg, 'a:p').forEach(function (p) {
          seg.substring(p.start, p.end).split(/<a:br\b[^>]*\/>/).forEach(function (q) { s.lines.push(slideText_(q)); });
        });
        s.text = s.paras.join('');
        s.title = /<p:ph\b[^>]*\stype="(?:title|ctrTitle)"/.test(seg);
        var sz = 0, m, re = /<a:(?:rPr|endParaRPr|defRPr)\b[^>]*\ssz="(\d+)"/g;
        while ((m = re.exec(seg)) !== null) sz = Math.max(sz, +m[1]);
        s.sz = sz;
      }
      if (tag === 'p:pic') s.photo = /<a:blip\b[^>]*r:embed=/.test(seg) && !/<a:videoFile\b|<a:audioFile\b/.test(seg);
      if (tag === 'p:graphicFrame') s.table = seg.indexOf('<a:tbl>') >= 0;
      list.push(s);
    });
  });
  list.sort(function (a, b) { return a.start - b.start; });
  var groups = list.filter(function (s) { return s.tag === 'p:grpSp'; });
  list.forEach(function (s) {
    var chain = groups.filter(function (g) { return g !== s && g.start < s.start && g.end >= s.end; });
    s.parent = chain.length ? chain[chain.length - 1].id : null;
    var b = { x: s.x, y: s.y, cx: s.cx, cy: s.cy };
    for (var i = chain.length - 1; i >= 0; i--) {
      var g = chain[i], sx = g.ch.cx ? g.cx / g.ch.cx : 1, sy = g.ch.cy ? g.cy / g.ch.cy : 1;
      b = { x: g.x + (b.x - g.ch.x) * sx, y: g.y + (b.y - g.ch.y) * sy, cx: b.cx * sx, cy: b.cy * sy };
    }
    s.ax = b.x; s.ay = b.y; s.acx = b.cx; s.acy = b.cy;
  });
  return list;
}
// 表のマス目 [[{ text, merged }]]（列の番号は <a:tc> の並び。つないだマス目も1つと数える）
function riTableCells_(seg) {
  return findTagRanges_(seg, 'a:tr').map(function (tr) {
    var trx = seg.substring(tr.start, tr.end);
    return findTagRanges_(trx, 'a:tc').map(function (tc) {
      var tcx = trx.substring(tc.start, tc.end), head = tcx.match(/^<a:tc\b[^>]*>/)[0];
      return { text: slideText_(tcx).trim(), merged: /\s(?:hMerge|vMerge)="1"/.test(head) };
    });
  });
}
function riGridCols_(seg) {
  var out = [], m, re = /<a:gridCol\b[^>]*\sw="(\d+)"/g;
  while ((m = re.exec(seg)) !== null) out.push(+m[1]);
  return out;
}
function riOverlap_(a, b) {
  var w = Math.min(a.ax + a.acx, b.ax + b.acx) - Math.max(a.ax, b.ax);
  return w > 0 ? w / Math.max(1, Math.min(a.acx, b.acx)) : 0;
}
// ページの題（題の枠があればそれ。無ければ上の方のいちばん大きな文字）
function riTitleOf_(shapes, H) {
  var t = shapes.filter(function (s) { return s.tag === 'p:sp' && s.title && s.text.trim(); });
  if (t.length) return t[0];
  var c = shapes.filter(function (s) { return s.tag === 'p:sp' && s.text.trim() && s.ay + s.acy / 2 < H * 0.3; });
  c.sort(function (a, b) { return (b.sz - a.sz) || (a.ay - b.ay); });
  return c[0] || null;
}
// 写真の枠の候補（小さな飾り・背景いっぱいの画像・動画は除く）
function riPhotos_(shapes, W, H) {
  return shapes.filter(function (s) { return s.tag === 'p:pic' && s.photo && s.acy >= H * 0.12 && s.acx <= W * 0.8; });
}
// ある枠の下（または上）のいちばん近い写真
function riPicNear_(pics, box, below, used) {
  var best = null, bestD = Infinity;
  pics.forEach(function (p) {
    if (used && used[p.id] || riOverlap_(p, box) < 0.5) return;
    var d = below ? p.ay - (box.ay + box.acy * 0.5) : (box.ay + box.acy * 0.5) - (p.ay + p.acy);
    if (d < -box.acy * 0.5 || d >= bestD) return;
    best = p; bestD = d;
  });
  return best;
}
// 1人ぶんのまとまり（カテゴリーの文字の下にお名前の文字、のグループ）
function riPersonUnits_(shapes) {
  var units = [];
  shapes.forEach(function (g) {
    if (g.tag !== 'p:grpSp') return;
    var kids = shapes.filter(function (c) { return c.parent === g.id && c.tag === 'p:sp' && c.text.trim(); });
    if (kids.length < 2 || kids.length > 3) return;
    kids.sort(function (a, b) { return a.ay - b.ay; });
    var name = kids[kids.length - 1];
    if (!riNameLike_(name.text) || kids.some(function (k) { return riWordOf_(k.text, RI_ROLE_WORDS_, true); })) return;
    units.push({ id: g.id, ax: g.ax, ay: g.ay, acx: g.acx, acy: g.acy, name: name, cats: kids.slice(0, -1) });
  });
  return units;
}

// --- 枠を見分ける ---
// 戻り値 { slots: [{ kind: 'holder'|'team'|'teamList', key, n, current, fields: [{ field, loc }], photo }], single: 役職のキー }
//   loc … { type: 'cell', id, row, col } / { type: 'shape', id } / { type: 'para', id, idx, wrap: ['（', '）'] }
function riDetectSlots_(xml, W, H) {
  var shapes = riShapes_(xml), pics = riPhotos_(shapes, W, H), out = { slots: [], single: '' };
  riDetectTables_(shapes, xml, pics, out, H);
  if (out.slots.length) return out;
  if (riDetectSingle_(shapes, pics, out, W, H)) return out;
  var title = riTitleOf_(shapes, H), team = title ? riWordOf_(title.text, RI_TEAM_WORDS_, false) : '';
  if (team && riDetectGrid_(shapes, pics, team, out)) return out;
  riDetectLabels_(shapes, pics, title, out);
  return out;
}

// 表：見出しのマス目が役職 → その下がお名前（リーダーシップチーム）。見出しと同じ行にお名前（コーディネーターの表）。
// 見出しがチーム → その下にメンバー（1人1マス、または「氏名, 氏名」の1マス）。
// 題がチーム（ビジターホストチーム）で、見出しの無いお名前の表 → 表のお名前を上から順に
function riDetectTables_(shapes, xml, pics, out, H) {
  var title = riTitleOf_(shapes, H), titleTeam = title ? riWordOf_(title.text, RI_TEAM_WORDS_, false) : '';
  shapes.forEach(function (f) {
    if (f.tag !== 'p:graphicFrame' || !f.table) return;
    var seg = xml.substring(f.start, f.end), rows = riTableCells_(seg), cols = riGridCols_(seg), used = {}, found = 0;
    var cell = function (r, c) { return rows[r] && rows[r][c] && !rows[r][c].merged ? rows[r][c] : null; };
    var loc = function (r, c) { return { type: 'cell', id: f.id, row: r, col: c }; };
    var r, c;
    for (r = 0; r < rows.length; r++) {
      var real = [];
      for (c = 0; c < rows[r].length; c++) if (!rows[r][c].merged) real.push(c);
      // 見出しとお名前が同じ行
      if (real.length === 2) {
        var k2 = riWordOf_(rows[r][real[0]].text, RI_ROLE_WORDS_, false) || riWordOf_(rows[r][real[0]].text, RI_ROW_WORDS_, false);
        var nm = rows[r][real[1]].text;
        if (k2 && (!nm || riNameLike_(nm))) {
          out.slots.push({ kind: 'holder', key: k2, current: nm, fields: [{ field: '氏名', loc: loc(r, real[1]) }] });
          used[r + ':' + real[0]] = used[r + ':' + real[1]] = true; found++;
          continue;
        }
      }
      for (c = 0; c < rows[r].length; c++) {
        var h = cell(r, c);
        if (!h || used[r + ':' + c]) continue;
        var role = riWordOf_(h.text, RI_ROLE_WORDS_, false), team = role ? '' : riWordOf_(h.text, RI_TEAM_WORDS_, false);
        if (role) {
          var below = cell(r + 1, c);
          if (!below || (below.text && !riNameLike_(below.text))) continue;
          // 写真：表の下で、その列の真下にあるもの
          var x0 = f.ax, i;
          for (i = 0; i < c && i < cols.length; i++) x0 += cols[i];
          var colBox = { ax: x0, ay: f.ay, acx: cols[c] || f.acx, acy: f.acy * 0.2 };
          var pic = riPicNear_(pics, colBox, true, null);
          out.slots.push({ kind: 'holder', key: role, current: below.text, fields: [{ field: '氏名', loc: loc(r + 1, c) }],
                           photo: pic && pic.ay >= f.ay + f.acy * 0.3 ? pic.id : null });
          used[(r + 1) + ':' + c] = true; found++;
        } else if (team) {
          var first = cell(r + 1, c);
          if (first && riNameList_(first.text)) {
            out.slots.push({ kind: 'teamList', key: team, current: first.text, fields: [{ field: 'の一覧', loc: loc(r + 1, c) }] });
            used[(r + 1) + ':' + c] = true; found++;
            continue;
          }
          for (var rr = r + 1, n = 1; rr < rows.length; rr++) {
            var b = cell(rr, c);
            if (!b || (b.text && !riNameLike_(b.text))) break;
            out.slots.push({ kind: 'team', key: team, n: n++, current: b.text, fields: [{ field: '氏名', loc: loc(rr, c) }] });
            used[rr + ':' + c] = true; found++;
          }
        }
      }
    }
    // 題がチームで、見出しの無いお名前の表：お名前の入った最初の行から、左上から順に
    if (!found && titleTeam) {
      var start = -1, cnt = 0;
      for (r = 0; r < rows.length; r++) {
        for (c = 0; c < rows[r].length; c++) if (cell(r, c) && riNameLike_(rows[r][c].text)) { cnt++; if (start < 0) start = r; }
      }
      if (cnt < 3) return;
      for (r = start, n = 1; r < rows.length; r++) {
        for (c = 0; c < rows[r].length; c++) {
          var e = cell(r, c);
          if (!e || (e.text && !riNameLike_(e.text))) continue;
          out.slots.push({ kind: 'team', key: titleTeam, n: n++, current: e.text, fields: [{ field: '氏名', loc: loc(r, c) }] });
        }
      }
    }
  });
}

// 1人のページ（題が役職の名前。大きな写真と、お名前・ローマ字・会社名・（カテゴリー）の文字）
function riDetectSingle_(shapes, pics, out, W, H) {
  var title = riTitleOf_(shapes, H);
  if (!title) return false;
  var key = '';
  title.lines.forEach(function (p) { key = key || riWordOf_(p, RI_ROLE_WORDS_, false); });
  if (!key) return false;
  var big = pics.filter(function (p) { return p.acx * p.acy >= W * H * 0.08; })
    .sort(function (a, b) { return b.acx * b.acy - a.acx * a.acy; })[0];
  var boxes = shapes.filter(function (s) {
    if (s.tag !== 'p:sp' || s === title || s.title || s.sz < 3600) return false;
    var first = s.paras.filter(function (p) { return p.trim(); })[0];
    return first && riNameLike_(first);
  }).sort(function (a, b) { return b.sz - a.sz; });
  if (!big || !boxes.length) return false;
  var box = boxes[0], lines = [];
  box.paras.forEach(function (t, i) { if (t.trim()) lines.push({ t: t.trim(), i: i }); });
  var at = function (i) { return { type: 'para', id: box.id, idx: i }; };
  var fields = [{ field: '氏名', loc: at(lines[0].i) }], k = 1;
  var paren = function (t) { var m = String(t).match(/^([（(])[\s\S]*([）)])$/); return m ? [m[1], m[2]] : null; };
  if (lines[k] && /^[A-Za-z][A-Za-z .'-]*$/.test(lines[k].t.normalize('NFKC').replace(/\s+/g, ' '))) fields.push({ field: 'ローマ字', loc: at(lines[k++].i) });
  if (lines[k] && !paren(lines[k].t)) fields.push({ field: '会社名', loc: at(lines[k++].i) });
  if (lines[k]) { var loc = at(lines[k].i); loc.wrap = paren(lines[k].t); fields.push({ field: 'カテゴリー', loc: loc }); }
  out.single = key;
  out.slots.push({ kind: 'holder', key: key, current: lines[0].t, fields: fields, photo: big.id });
  return true;
}

// 題がチーム（メンバーシップ委員会）で、写真の下にカテゴリーとお名前が並ぶページ。左上から順に
function riDetectGrid_(shapes, pics, team, out) {
  var units = riPersonUnits_(shapes).map(function (u) { u.pic = riPicNear_(pics, u, false, null); return u; })
    .filter(function (u) { return u.pic; });
  if (units.length < 2) return false;
  units.sort(function (a, b) { return Math.abs(a.ay - b.ay) > a.acy ? a.ay - b.ay : a.ax - b.ax; });
  units.forEach(function (u, i) {
    var f = [{ field: '氏名', loc: { type: 'shape', id: u.name.id, w: u.name.acx, h: u.name.acy } }];
    if (u.cats.length) f.push({ field: 'カテゴリー', loc: { type: 'shape', id: u.cats[0].id, w: u.cats[0].acx, h: u.cats[0].acy } });
    out.slots.push({ kind: 'team', key: team, n: i + 1, current: u.name.text.trim(), fields: f, photo: u.pic.id });
  });
  return true;
}

// 役職の見出しの下に写真、その下にカテゴリーとお名前（各サポートチームのページ）
function riDetectLabels_(shapes, pics, title, out) {
  var units = riPersonUnits_(shapes), inUnit = {}, got = [], usedPic = {}, usedUnit = {};
  units.forEach(function (u) { u.cats.concat([u.name]).forEach(function (s) { inUnit[s.id] = true; }); });
  shapes.forEach(function (s) {
    if (s.tag !== 'p:sp' || s === title || s.title || inUnit[s.id] || riNorm_(s.text).length > 30) return;
    var key = riWordOf_(s.text, RI_ROLE_WORDS_, true);
    if (!key) return;
    var pic = riPicNear_(pics, s, true, usedPic);
    if (!pic) return;
    var best = null, bestD = Infinity;
    units.forEach(function (u) {
      if (usedUnit[u.id] || riOverlap_(u, pic) < 0.5) return;
      var d = u.ay - (pic.ay + pic.acy * 0.5);
      if (d >= 0 && d < bestD) { best = u; bestD = d; }
    });
    if (!best) return;
    usedPic[pic.id] = usedUnit[best.id] = true;
    var f = [{ field: '氏名', loc: { type: 'shape', id: best.name.id, w: best.name.acx, h: best.name.acy } }];
    if (best.cats.length) f.push({ field: 'カテゴリー', loc: { type: 'shape', id: best.cats[0].id, w: best.cats[0].acx, h: best.cats[0].acy } });
    got.push({ kind: 'holder', key: key, current: best.name.text.trim(), fields: f, photo: pic.id });
  });
  if (got.length >= 2) out.slots = out.slots.concat(got);
}

// --- 枠に差し込み口を書く ---
function riSlotBase_(slot, data) {
  if (slot.kind === 'holder') return riRoleBase_(slot.key);
  var t = riTeamBase_(slot.key, data);
  return t ? (slot.kind === 'teamList' ? t : t + slot.n) : '';
}
function riSetPara_(xml, id, idx, text) {
  var r = findShapeRange_(xml, id);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), ps = findTagRanges_(seg, 'a:p');
  if (!ps[idx]) return xml;
  seg = seg.substring(0, ps[idx].start) + oneRunParagraph_(seg.substring(ps[idx].start, ps[idx].end), text) + seg.substring(ps[idx].end);
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}
function riNamePic_(xml, id, name) {
  var r = findShapeRange_(xml, id);
  if (!r) return xml;
  var nm = escapeXml_(name);                         // 関数で置き換える（お名前の「$&」「$1」を記号として読まない）
  var seg = xml.substring(r.start, r.end).replace(/(<p:cNvPr\b[^>]*\sname=")[^"]*(")/, function (all, a, b) { return a + nm + b; });
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}
// fits … 書いた枠のうち、入れた文字に合わせて小さくする図形 [{ id, lines, w }]（渡したときだけ足す）
function riWriteSlots_(xml, slots, data, fits) {
  slots.forEach(function (s) {
    var base = riSlotBase_(s, data);
    if (!base) return;
    s.fields.forEach(function (f) {
      var l = f.loc, tok = '{{' + base + f.field + (l.wrap ? (l.wrap[0] === '(' ? '()' : '（）') : '') + '}}';
      if (l.type === 'cell') xml = setTableCellText_(xml, l.id, l.row, l.col, tok);
      else if (l.type === 'shape') {
        xml = setParagraphsInShape_(xml, l.id, [tok]);
        if (fits && l.w) fits.push({ id: l.id, lines: f.field === '氏名' ? 1 : 0, w: l.w, h: l.h });
      }
      else xml = riSetPara_(xml, l.id, l.idx, tok);
    });
    if (s.photo) xml = riNamePic_(xml, s.photo, '{{' + base + '写真}}');
  });
  return xml;
}
// 差し込み口だけ入れる（公式ファイルから雛形を作るとき）
function riTokenizePage_(xml, W, H) {
  var det = riDetectSlots_(xml, W, H);
  return det.slots.length ? riWriteSlots_(xml, det.slots, null) : xml;
}

// --- 差し込み口を、その期の方で埋める ---
// （PowerPoint が差し込み口を複数のランに分けていても見つかるよう、文字だけをつないで探す）
function riHasTokens_(xml, map) {
  var t = slideText_(xml), re = /\{\{([^{}]{1,60})\}\}/g, m;
  while ((m = re.exec(t)) !== null) if (Object.prototype.hasOwnProperty.call(map, m[1])) return true;
  return false;
}
// 文字の幅の見積もり（em。全角1・半角0.55・空白0.3）
function riTextEm_(t) {
  var w = 0, s = String(t == null ? '' : t);
  for (var i = 0; i < s.length; i++) {
    var ch = s.charAt(i);
    w += (ch === ' ' || ch === '　') ? (ch === ' ' ? 0.3 : 1) : (s.charCodeAt(i) < 0x2000 ? 0.55 : 1);
  }
  return w;
}
// 図形の文字が lines 行に収まるよう、文字を小さくする（収まっていれば何もしない。元の大きさの半分まで）。
// lines が 0 なら、枠の高さに入る行数（カテゴリーは、1行ぶんの高さの枠なら1行に）
function riFitShape_(xml, id, lines, wEmu, hEmu) {
  var r = findShapeRange_(xml, id);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), sz = 0, m, re = /<a:(?:rPr|endParaRPr)\b[^>]*\ssz="(\d+)"/g;
  while ((m = re.exec(seg)) !== null) sz = Math.max(sz, +m[1]);
  var text = slideText_(seg).trim();
  if (!sz || !text) return xml;
  if (!lines) lines = Math.max(1, Math.floor((hEmu || 0) / (sz / 100 * 12700 * 1.2)));
  var avail = Math.max(1, wEmu - 2 * 91440) / 12700 * 0.95 * lines;        // 左右の余白を除いた幅（pt）× 行数
  var need = riTextEm_(text) * sz / 100;
  if (need <= avail) return xml;
  var to = Math.max(Math.floor(sz / 2 / 50) * 50, Math.floor(sz * avail / need / 50) * 50);
  seg = seg.replace(/(<a:(?:rPr|endParaRPr)\b[^>]*\ssz=")\d+(")/g, '$1' + to + '$2');
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}

// 写真の枠 {{○○写真}}：その方の写真に替える。居ない方・写真の無い方の枠は消す（前の方の写真を残さない）
function riFillPhotos_(parts, path, xml, persons, cache, res) {
  var want = [];
  findTagRanges_(xml, 'p:pic').forEach(function (r) {
    var nv = xml.substring(r.start, r.end).match(/<p:cNvPr\b[^>]*>/);
    var id = nv && (nv[0].match(/\sid="(\d+)"/) || [])[1], nm = nv && (nv[0].match(/\sname="\{\{([^"{}]+)写真\}\}"/) || [])[1];
    if (!id || !nm) return;
    nm = unescapeXml_(nm);
    if (Object.prototype.hasOwnProperty.call(persons, nm)) want.push({ id: id, base: nm });
  });
  if (!want.length) return xml;
  var rp = relsPathOf_(path), rels = xmlOf_(parts, rp);
  want.forEach(function (w) {
    var p = persons[w.base], photo = p ? mpAddPhoto_(parts, cache, p.name) : null;
    if (!photo) {
      xml = removeShape_(xml, w.id);
      if (p && res.noPhoto.indexOf(p.name) < 0) res.noPhoto.push(p.name);
      return;
    }
    var box = readShapeGeomEmu_(xml, w.id);
    if (box && photo.width && photo.height) xml = setSrcRectInPic_(xml, w.id, coverCrop_(photo.width, photo.height, box.cx, box.cy));
    var set = setPicImage_(xml, rels, w.id, '../media/' + photo.path.replace('ppt/media/', ''));
    xml = set.xml; rels = set.rels;
    res.photos++;
  });
  if (rels) putXml_(parts, rp, rels);
  return xml;
}

// === ネットワーキング学習コーナー ===
// 担当は、その期のエデュケーションコーディネーター（役職・チーム（半期ごと））。
//   ・「担当：」の文字のある枠（Activeチャプターの雛形）… 1行目にカテゴリー、2行目に「担当：お名前」。
//     雛形の斜体・自動縮小（文字がとても小さくなる）をやめ、カテゴリー20pt、「担当：」20pt・お名前40ptの太字にする
//     （お名前は枠の幅に収まらないときだけ小さくする）。枠の高さが足りなければ下へ伸ばす
//     （以前はお名前も「担当：」と同じ32ptで、自動縮小も残っていたため、Googleスライドなどでは小さく出た）
//   ・公式ファイルの作り（「氏名／学習トピック」の枠）… 「氏名」をお名前にする（学習トピックはそのまま）
//   ・写真 … その枠の上の写真（無ければいちばん大きな写真）を、エデュケーションコーディネーターの写真にする。
//     写真が無い方のときは枠を外し、「スピーカーの写真を挿入」のような見本の文字も消す
var RI_LEARN_CAT_PT_ = 20, RI_LEARN_LABEL_PT_ = 20, RI_LEARN_NAME_PT_ = 40;
function riLearnPt_(text, want, wEmu) {
  var avail = Math.max(1, (wEmu || 0) - 2 * 91440) / 12700 * 0.95, em = riTextEm_(text);
  return em * want <= avail ? want : Math.max(Math.floor(want / 2), Math.floor(avail / em));
}
// 「担当：」のうしろのお名前の大きさ。「担当：」と合わせて枠の幅に収まる大きさ（お名前だけ小さくする）
function riLearnNamePt_(label, name, wEmu) {
  var avail = Math.max(1, (wEmu || 0) - 2 * 91440) / 12700 * 0.95 - riTextEm_(label) * RI_LEARN_LABEL_PT_, em = riTextEm_(name);
  if (!em || em * RI_LEARN_NAME_PT_ <= avail) return RI_LEARN_NAME_PT_;
  return Math.max(Math.floor(RI_LEARN_NAME_PT_ / 2), Math.floor(avail / em));
}
function riRewriteLearnBox_(xml, box, p, H) {
  var r = findShapeRange_(xml, box.id);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), tb = findTagRanges_(seg, 'p:txBody')[0];
  if (!tb) return xml;
  var body = seg.substring(tb.start, tb.end);
  // 文字の色・書体は、雛形の最初の文字のまま。斜体・大きさ・太字は決め直す
  var rPr = (body.match(/<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/) || ['<a:rPr lang="ja-JP"/>'])[0]
    .replace(/\s(?:i|sz|b|dirty|err)="[^"]*"/g, '');
  var run = function (text, pt, bold) {
    return '<a:r>' + rPr.replace(/^<a:rPr\b/, '<a:rPr sz="' + pt * 100 + '" b="' + (bold ? 1 : 0) + '" i="0"')
         + '<a:t>' + escapeXml_(text) + '</a:t></a:r>';
  };
  var para = function (runs, endPt) {
    return '<a:p><a:pPr algn="ctr"/>' + runs + '<a:endParaRPr lang="ja-JP" sz="' + endPt * 100 + '"/></a:p>';
  };
  // 自動縮小はやめる（normAutofit が残っていると、Googleスライドなどが枠に合わせて字を小さくする）
  var bodyPr = (body.match(/<a:bodyPr\b[^>]*\/>|<a:bodyPr\b[^>]*>[\s\S]*?<\/a:bodyPr>/) || ['<a:bodyPr/>'])[0]
    .replace(/<a:(?:normAutofit|spAutoFit|noAutofit)\b[^>]*\/>/g, '');
  bodyPr = /\/>$/.test(bodyPr) ? bodyPr.replace(/\/>$/, '><a:noAutofit/></a:bodyPr>')
         : /<\/a:prstTxWarp>/.test(bodyPr) ? bodyPr.replace('</a:prstTxWarp>', '</a:prstTxWarp><a:noAutofit/>')
         : bodyPr.replace(/^(<a:bodyPr\b[^>]*>)/, '$1<a:noAutofit/>');
  var lst = (body.match(/<a:lstStyle\b[^>]*\/>|<a:lstStyle\b[^>]*>[\s\S]*?<\/a:lstStyle>/) || ['<a:lstStyle/>'])[0];
  var label = '担当：', catPt = p.category ? riLearnPt_(p.category, RI_LEARN_CAT_PT_, box.acx) : 0;
  var namePt = riLearnNamePt_(label, p.name, box.acx);
  var paras = (p.category ? para(run(p.category, catPt, false), catPt) : '')
            + para(run(label, RI_LEARN_LABEL_PT_, true) + run(p.name, namePt, true), namePt);
  seg = seg.substring(0, tb.start) + '<p:txBody>' + bodyPr + lst + paras + '</p:txBody>' + seg.substring(tb.end);
  xml = xml.substring(0, r.start) + seg + xml.substring(r.end);
  // 枠の高さが2行ぶんに足りなければ、下へ伸ばす（スライドの下にはみ出すなら、そのぶん上へ）。グループの中の枠はそのまま
  var g = box.parent ? null : readShapeGeomEmu_(xml, box.id);
  var need = Math.round((catPt + namePt) * 1.2 * 12700 + 2 * 45720);
  if (g && g.cy < need) {
    var y = g.y;
    if (H && y + need > H) y = Math.max(0, H - need);
    xml = setShapeGeomEmu_(xml, box.id, { y: y, cy: need });
  }
  return xml;
}
function riLearningCorner_(xml, data, W, H) {
  var ec = data && data.holdersOk ? data.roles.ec : null;
  if (!ec) return { xml: xml, done: false };
  var shapes = riShapes_(xml), pics = riPhotos_(shapes, W, H), pic = null;
  var box = shapes.filter(function (s) { return s.tag === 'p:sp' && /担当\s*[：:]/.test(s.text); })[0];
  if (box) {
    xml = riRewriteLearnBox_(xml, box, ec, H);
    pic = riPicNear_(pics, box, false, null);
  } else {
    box = shapes.filter(function (s) {
      return s.tag === 'p:sp' && s.lines.some(function (l) { return l.trim() === '氏名'; });
    })[0];
    if (!box) return { xml: xml, done: false };
    var r = findShapeRange_(xml, box.id), seg = xml.substring(r.start, r.end);
    seg = replaceTokensInXml_(seg.replace(/(<a:t(?:\s[^>]*)?>)氏名(<\/a:t>)/, '$1{{学習コーナー担当}}$2'), { '学習コーナー担当': ec.name });
    xml = xml.substring(0, r.start) + seg + xml.substring(r.end);
  }
  if (!pic) pic = pics.slice().sort(function (a, b) { return b.acx * b.acy - a.acx * a.acy; })[0] || null;
  if (pic) {
    xml = riNamePic_(xml, pic.id, '{{' + riRoleBase_('ec') + '写真}}');
    // 写真の無い方：写真の枠は外すので、後ろの「スピーカーの写真を挿入」のような見本の文字も消す
    if (!findPhotoIdForName_(ec.name)) {
      shapes.forEach(function (s) {
        if (s.tag === 'p:sp' && /写真/.test(s.text) && riOverlap_(s, pic) > 0.5) xml = setParagraphsInShape_(xml, s.id, ['']);
      });
    }
  }
  return { xml: xml, done: true, name: ec.name };
}

// 前半スライドの役職紹介のページを、その期の方にする。
// data … riDataFrom_ / riDataOfDate_ の結果。null なら、差し込み口を空にするだけ（写真の枠はそのまま）
function applyRoleIntro_(parts, data, cache) {
  // filled … 入れたお名前 [{ path, base, name }]（検査で使う）
  var res = { pages: [], hidden: [], noPhoto: [], overflow: [], photos: 0, filled: [] };
  var prs = xmlOf_(parts, 'ppt/presentation.xml') || '';
  var sz = prs.match(/<p:sldSz\s+cx="(\d+)"\s+cy="(\d+)"/), W = sz ? +sz[1] : 12192000, H = sz ? +sz[2] : 6858000;
  var order = slideOrder_(parts), anchor = weeklyAnchor_(parts), stop = anchor ? order.indexOf(anchor) : order.length;
  var vals = riValues_(data);
  cache = cache || { by: {}, seq: 0 };
  for (var i = 0; i < order.length; i++) {
    var path = order[i], xml = xmlOf_(parts, path), before = xml, fits = [];
    if (!xml) continue;
    var hasTok = riHasTokens_(xml, vals.map) || /\sname="\{\{[^"{}]+写真\}\}"/.test(xml);
    // 差し込み口の無いページ：ページの作りから枠を見分け、お名前が替わった枠だけ差し込み口を入れる
    if (!hasTok && data && (data.holdersOk || data.teamsOk) && i < stop) {
      var det = riDetectSlots_(xml, W, H);
      if (det.single && data.holdersOk && !data.roles[det.single]) {
        xml = setSlideShow_(xml, false);
        res.hidden.push(riRoleBase_(det.single));
      } else if (det.slots.length) {
        var counts = {};
        var write = det.slots.filter(function (s) {
          if (s.kind === 'holder') {
            if (!data.holdersOk) return false;
            var p = data.roles[s.key];
            return normName_(s.current) !== normName_(p ? p.name : '');
          }
          if (!data.teamsOk) return false;
          var list = data.teams[s.key] || [];
          if (s.kind === 'teamList') return normName_(s.current) !== normName_(list.map(function (x) { return x.name; }).join(''));
          counts[s.key] = (counts[s.key] || 0) + 1;
          var q = list[s.n - 1];
          return normName_(s.current) !== normName_(q ? q.name : '');
        });
        Object.keys(counts).forEach(function (k) {
          var n = (data.teams[k] || []).length;
          if (n > counts[k]) res.overflow.push(riTeamBase_(k, data) + '（' + n + '名のうち' + counts[k] + '名ぶんの枠）');
        });
        if (write.length) xml = riWriteSlots_(xml, write, data, fits);
      }
    }
    // ネットワーキング学習コーナー（担当はその期のエデュケーションコーディネーター）
    if (data && i < stop && /学習コーナー/.test(slideText_(xml))) {
      var lc = riLearningCorner_(xml, data, W, H);
      if (lc.done) { xml = lc.xml; res.learning = lc.name; }
    }
    if (xml !== before || hasTok) {
      var tm, txt = slideText_(xml), tre = /\{\{([^{}]{1,60})氏名\}\}/g;
      while ((tm = tre.exec(txt)) !== null) {
        if (Object.prototype.hasOwnProperty.call(vals.map, tm[1] + '氏名')) res.filled.push({ path: path, base: tm[1], name: vals.map[tm[1] + '氏名'] });
      }
      tre = /\{\{([^{}]{1,60})の一覧\}\}/g;
      while ((tm = tre.exec(txt)) !== null) {
        if (Object.prototype.hasOwnProperty.call(vals.map, tm[1] + 'の一覧')) res.filled.push({ path: path, base: tm[1] + 'の一覧', name: vals.map[tm[1] + 'の一覧'] });
      }
      xml = replaceTokensInXml_(xml, vals.map);
      fits.forEach(function (f) { xml = riFitShape_(xml, f.id, f.lines, f.w, f.h); });
      if (data) xml = riFillPhotos_(parts, path, xml, vals.persons, cache, res);
    }
    if (xml !== before) {
      putXml_(parts, path, xml);
      res.pages.push(i + 1);
    }
  }
  res.message = riMessage_(data, res);
  return res;
}
function riMessage_(data, res) {
  if (!data) return '';
  var head = '役職のメンバー紹介（' + data.term + '期）: ';
  if (!data.holdersOk && !data.teamsOk) return head + '役職・チームが未登録のため、テンプレートのままにしました。';
  var msg = head + (res.pages.length ? res.pages.length + '枚のページを、この期の方にしました（写真 ' + res.photos + '枚）。'
                                      : '入れ替えるところはありませんでした（テンプレートのとおりです）。');
  if (!data.teamsOk) msg += 'チームのメンバーは未登録のため、チームのページはテンプレートのままです。';
  if (res.hidden.length) msg += '\n担当者が未登録の役職のページは非表示にしました: ' + res.hidden.join('、');
  if (res.overflow.length) msg += '\nページの枠より人数が多いチーム: ' + res.overflow.join('、');
  if (res.learning) msg += '\nネットワーキング学習コーナー：担当 ' + res.learning + 'さん（その期のエデュケーションコーディネーター）';
  if (res.noPhoto.length) msg += '\n写真が見つからない方（写真の枠を外しました）: ' + res.noPhoto.join('、');
  return msg;
}
