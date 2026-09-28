// === メンバー名簿マスタ（メンバーブック／メンバープレゼン／Zoom案内の共通データ基盤）===
// 既存の「メンバーリスト」(A=No/B=氏名、OCRで更新) は割り振り表などのキーとして使い続け、
// 詳細情報はこの「メンバー名簿」シートに持つ。氏名で紐付ける。
// 行順＝メンバーブックの掲載順＝メンバープレゼンの発表順。

var MEMBER_SHEET_ = 'メンバー名簿';
// メンバーリスト(OCR)で読むPDFの列（No/氏名/ふりがな/カテゴリー/会社名/役職/メモ）と、
// メンバーブックで使う項目を1枚に統合した。これが全機能の正本になる。
// 「役職」はBNIの役職（役職・チーム（半期ごと）から期ごとに入る。Spreadingの position もこれ）。
// 「会社での役職」は、代表取締役・支店長などの会社での肩書き（メンバーブックに載せる）。あとから足した列なので最後にある
var MEMBER_HEADERS_ = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ',
                       '写真ファイル名', '一言コメント', '紹介してほしい人', '協業したい人',
                       '入会日', '更新日', '更新期限日', '会社での役職'];
var COVER_SHEET_ = 'メンバーブック表紙';
var CAT_SHEET_ = '業種区分マスタ';

// 業種区分（キー・表示ラベル・色1・色2・ブロック表示名・巡回順）。
// キーはSpreadingのカテゴリ設定のグループ名と同じにしてある。名簿の「業種区分」列と
// この キー が一致したときに、メンバーブックのカードに色が付く。
// 色1は冊子のカード帯の地色で、実際の配布版PDFから採取した値。
// 巡回順は、Spreadingのグループの並び順に合わせてある。
var DEFAULT_CATEGORIES_ = [
  ['企業サポート',   '企業サポート',   '#FFDE58', '#F5C400', '企業サポート',   1],
  ['研修・教育',     '研修・教育',     '#C1FF72', '#8FD43A', '研修・教育',     2],
  ['不動産関連',     '不動産関連',     '#37B5FF', '#0B8FE0', '不動産関連',     3],
  ['建築・住まい',   '建築・住まい',   '#AAB5D9', '#7A88B8', '建築＆住まい',   4],
  ['プロモーション', 'プロモーション', '#FF66C3', '#E0329B', 'プロモーション', 5],
  ['暮らし・生活',   '暮らし・生活',   '#5CE1E6', '#1FBCC2', '暮らし・生活',   6],
  ['美容と健康',     '美容と健康',     '#7DD957', '#4FAF2A', '美容・健康',     7],
  ['飲食・エンタメ', '飲食・エンタメ', '#CB6BE6', '#A33BC2', '飲食・エンタメ', 8]
];

function openMemberMasterDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('member_master').setWidth(1000).setHeight(720),
    'メンバー名簿の管理');
}

function ensureMemberSheet_() {
  var ss = getSS_(), sh = ss.getSheetByName(MEMBER_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(MEMBER_SHEET_);
    sh.appendRow(MEMBER_HEADERS_);
    sh.getRange(1, 1, 1, MEMBER_HEADERS_.length).setFontWeight('bold').setBackground('#f2f6ff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 以前の既定の業種区分（2026年9月23日の版より前）。
// そのころに作られた「業種区分マスタ」には、この行がそのまま残っていて、今の既定の行はうしろに足されている。
// すると同じ区分が2行になり（「建築住まい」と「建築・住まい」など）、巡回順も古いまま（プロモーションが4番）なので、
// ウィークリープレゼンの順番がずれる（建築＆住まいの次が不動産関連、プロモーションの次が美容・健康になる）。
// 自動で入ったままの行（6つの値がこのとおりの行）だけを今の既定に直す。手で直した行は触らない。
var LEGACY_CATEGORIES_ = [
  ['企業サポート', '企業サポート', '#1f4e79', '#2e75b6', '企業サポート', 1],
  ['研修教育',     '研修・教育',   '#375623', '#548235', '研修・教育',   2],
  ['建築住まい',   '建築・住まい', '#7f6000', '#bf9000', '建築＆住まい', 3],
  ['プロモーション', 'プロモーション', '#833c00', '#c55a11', 'プロモーション', 4],
  ['美容健康',     '美容・健康',   '#7b2d52', '#c0507f', '美容・健康',   5],
  ['金融保険',     '金融・保険',   '#1f3864', '#2f5597', '金融・保険',   6],
  ['不動産',       '不動産',       '#4d3b63', '#7030a0', '不動産',       7],
  ['暮らしサービス', '暮らしサービス', '#255e5e', '#38859c', '暮らしサービス', 8]
];

// 区分の書き方の違い（「建築・住まい」「建築住まい」、「美容と健康」「美容健康」）をそろえたもの
function catNormKey_(s) {
  return String(s == null ? '' : s).normalize('NFKC').replace(/[\s・･＆&と]/g, '');
}

// 昔の既定のまま残っている行を、今の既定に直した行の並びを返す。
//   ・今の既定に同じキーがある（企業サポート・プロモーション）… 今の既定の値（色・巡回順）にする
//   ・書き方違いの同じ区分が今の既定にある（研修教育・建築住まい・美容健康）… 消す（今の既定の行を使う）
//   ・今の既定に無い区分（金融保険・不動産・暮らしサービス）… そのまま（使われていなければ「使われていない業種区分を消す」で消せる）
function repairLegacyCategoryRows_(rows) {
  var cur = {}, curNorm = {}, legacy = {}, out = [], changed = [], i;
  for (i = 0; i < DEFAULT_CATEGORIES_.length; i++) {
    cur[DEFAULT_CATEGORIES_[i][0]] = DEFAULT_CATEGORIES_[i];
    curNorm[catNormKey_(DEFAULT_CATEGORIES_[i][0])] = DEFAULT_CATEGORIES_[i];
  }
  for (i = 0; i < LEGACY_CATEGORIES_.length; i++) legacy[LEGACY_CATEGORIES_[i][0]] = LEGACY_CATEGORIES_[i];
  var same = function (row, lg) {
    for (var c = 0; c < 6; c++) {
      var a = String(row[c] == null ? '' : row[c]).trim(), b = String(lg[c]);
      if (c === 5 ? Number(a) !== lg[5] : a.toLowerCase() !== b.toLowerCase()) return false;
    }
    return true;
  };
  for (i = 0; i < rows.length; i++) {
    var key = String(rows[i][0] == null ? '' : rows[i][0]).trim(), lg = legacy[key];
    if (lg && same(rows[i], lg)) {
      if (cur[key]) { out.push(cur[key].slice()); changed.push(key); continue; }
      if (curNorm[catNormKey_(key)]) { changed.push(key); continue; }
    }
    out.push(rows[i].slice(0, 6));
  }
  return { rows: out, changed: changed };
}

function ensureCategorySheet_() {
  var ss = getSS_(), sh = ss.getSheetByName(CAT_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(CAT_SHEET_);
    sh.appendRow(['キー', '表示ラベル', '色1', '色2', 'ブロック表示名', '巡回順']);
    sh.getRange(2, 1, DEFAULT_CATEGORIES_.length, 6).setValues(DEFAULT_CATEGORIES_);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f2f6ff');
    sh.setFrozenRows(1);
    sh.hideSheet();
    return sh;
  }
  // 昔の既定のまま残っている行を、今の既定に直す（上の LEGACY_CATEGORIES_）。
  // 区分が増えたとき（Spreadingのグループが変わったときなど）に足りない行を補う。
  // それ以外の行の色や並びは、手で直されている可能性があるのでそのまま残す。
  var data = sh.getDataRange().getValues(), body = data.slice(1);
  var fixed = repairLegacyCategoryRows_(body), have = {};
  for (var i = 0; i < fixed.rows.length; i++) have[String(fixed.rows[i][0]).trim()] = true;
  var add = [];
  for (var j = 0; j < DEFAULT_CATEGORIES_.length; j++) {
    if (!have[DEFAULT_CATEGORIES_[j][0]]) add.push(DEFAULT_CATEGORIES_[j].slice());
  }
  if (fixed.changed.length) {
    var all = fixed.rows.concat(add);
    if (all.length) sh.getRange(2, 1, all.length, 6).setValues(all);
    if (body.length > all.length) sh.getRange(2 + all.length, 1, body.length - all.length, 6).clearContent();
    console.log('[CAT] 昔の既定のまま残っていた業種区分を今の既定に直しました: ' + fixed.changed.join('・')
      + (add.length ? '（' + add.length + '件追加）' : ''));
  } else if (add.length) {
    sh.getRange(sh.getLastRow() + 1, 1, add.length, 6).setValues(add);
    console.log('[CAT] 業種区分を' + add.length + '件追加しました');
  }
  return sh;
}

// 業種区分の色や並びを初期値に戻す（冊子の配色に合わせ直したいとき）
function resetCategoryMaster() {
  try {
    var sh = ensureCategorySheet_();
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
    sh.getRange(2, 1, DEFAULT_CATEGORIES_.length, 6).setValues(DEFAULT_CATEGORIES_);
    return { ok: true, message: '業種区分を初期値に戻しました（' + DEFAULT_CATEGORIES_.length + '区分）。',
             categories: getCategoryMaster() };
  } catch (e) {
    return { ok: false, message: '戻せませんでした: ' + (e && e.message ? e.message : e) };
  }
}

function toDateStr_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  }
  return String(v).trim();
}

// opts.membersOnly … 名簿の行だけ返す（表紙・業種区分マスタを読まないぶん速い。サーバーの中から使う）
function getMemberMaster(opts) {
  try {
    // 期が替わっていたら、役職・チーム（半期ごと）を「役職」に反映してから読む（role_input_srv.js）
    if (!(opts && opts.membersOnly) && typeof roleRosterAutoSync_ === 'function') roleRosterAutoSync_();
    var sh = ensureMemberSheet_(), data = sh.getDataRange().getValues(), members = [];
    var col = function (r, i) { return String(r[i] == null ? '' : r[i]).trim(); };
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!col(r, 2)) continue;                       // 氏名が無い行は飛ばす
      members.push({
        no:        col(r, 0),
        cat:       col(r, 1),
        name:      col(r, 2),
        kana:      col(r, 3),
        title:     col(r, 4),     // カテゴリー（業務内容）
        company:   col(r, 5),
        role:      col(r, 6),     // 役職
        memo:      col(r, 7),
        photoFile: col(r, 8),
        comment:   col(r, 9),
        refer:     col(r, 10),
        collab:    col(r, 11),
        joinDate:   toDateStr_(r[12]),
        renewDate:  toDateStr_(r[13]),
        expireDate: toDateStr_(r[14]),
        position:  col(r, 15)     // 会社での役職（肩書き）
      });
    }
    if (opts && opts.membersOnly) return { ok: true, members: members };
    return { ok: true, members: members, cover: getCoverInfo_(), categories: getCategoryMaster(), backup: getMemberBackupInfo_() };
  } catch (e) {
    console.error('[MEMBER] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: 'メンバー名簿の読み込みに失敗しました: ' + (e && e.message ? e.message : e), members: [] };
  }
}

function saveMemberMaster(members, cover) {
  try {
    if (!members) return { ok: false, message: '保存するデータがありません。' };
    var sh = ensureMemberSheet_();
    sh.clear();
    sh.appendRow(MEMBER_HEADERS_);
    sh.getRange(1, 1, 1, MEMBER_HEADERS_.length).setFontWeight('bold').setBackground('#f2f6ff');
    sh.setFrozenRows(1);
    var rows = [];
    for (var i = 0; i < members.length; i++) {
      var m = members[i] || {};
      rows.push([m.no || '', m.cat || '', m.name || '', m.kana || '', m.title || '',
                 m.company || '', m.role || '', m.memo || '', m.photoFile || '',
                 m.comment || '', m.refer || '', m.collab || '',
                 m.joinDate || '', m.renewDate || '', m.expireDate || '', m.position || '']);
    }
    if (rows.length) sh.getRange(2, 1, rows.length, MEMBER_HEADERS_.length).setValues(rows);
    if (cover) saveCoverInfo_(cover);
    console.log('[MEMBER] saved=' + rows.length);
    return { ok: true, message: 'メンバー名簿を保存しました（' + rows.length + '名）。' };
  } catch (e) {
    console.error('[MEMBER] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}


// 氏名から写真ファイル名を自動解決して写真列を埋める
function autoFillMemberPhotos() {
  try {
    var m = getMemberMaster();
    if (!m.ok) return m;
    var members = m.members, filled = 0, miss = [];
    for (var i = 0; i < members.length; i++) {
      var id = findPhotoIdForName_(members[i].name);
      if (!id) { miss.push(members[i].name); continue; }
      try {
        var nm = DriveApp.getFileById(id).getName();
        if (members[i].photoFile !== nm) { members[i].photoFile = nm; filled++; }
      } catch (e) { miss.push(members[i].name); }
    }
    var res = saveMemberMaster(members, null);
    if (!res.ok) return res;
    return { ok: true, message: filled + '名の写真を紐付けました。未照合 ' + miss.length + '名。', filled: filled, missing: miss };
  } catch (e) {
    return { ok: false, message: '写真の自動照合に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// --- 業種区分マスタ ---
function getCategoryMaster() {
  var sh = ensureCategorySheet_(), data = sh.getDataRange().getValues(), list = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0] || '').trim()) continue;
    list.push({ key: String(data[i][0]).trim(), label: String(data[i][1] || '').trim(),
                bg: String(data[i][2] || '').trim(), bg2: String(data[i][3] || '').trim(),
                block: String(data[i][4] || '').trim(), order: Number(data[i][5]) || 0 });
  }
  list.sort(function (a, b) { return a.order - b.order; });
  return list;
}

function saveCategoryMaster(rows) {
  try {
    var sh = ensureCategorySheet_();
    sh.clear();
    sh.appendRow(['キー', '表示ラベル', '色1', '色2', 'ブロック表示名', '巡回順']);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f2f6ff');
    var vals = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i] || {};
      vals.push([r.key || '', r.label || '', r.bg || '', r.bg2 || '', r.block || '', Number(r.order) || (i + 1)]);
    }
    if (vals.length) sh.getRange(2, 1, vals.length, 6).setValues(vals);
    return { ok: true, message: '業種区分マスタを保存しました。' };
  } catch (e) {
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// --- メンバーブック表紙情報 ---
// === メンバーブックの表紙ページ ===
// 期ごとに書き換わる文章（プレジデント挨拶など）と、めったに変わらない定型文
// （理念・用語の説明など）をまとめて1件のJSONで持つ。
// 項目が増えてもプロパティを増やさずに済む。
var COVER_KEY_ = 'BNI_MB_COVER';

// チャプター名・期が入る項目（title・term・aboutTitle・prole）は defaultCover_() で、チャプターの設定から作る
var DEFAULT_COVER_ = {
  philosophyTitle: 'BNIの理念　Givers Gain®（ギバーズゲイン）',
  philosophy: '他の人におしみなくビジネスを提供することで、自分も他の人からビジネスを提供してもらえる、「与えるものは与えられる」という考え方に基づいて運営されています。',
  about: '個性豊かなメンバーが、様々な分野のプロフェッショナルとして、お互いのビジネス発展、売上UP、人脈の拡大のためにサポートし合うビジネスチームです。',
  benefitsTitle: 'BNIに参加するメリット（BNIを活用する５つのベネフィット＋1）',
  benefits: '１.大きなマーケティングチーム　２.競合のいないビジネス環境　３.継続的な新規顧客の紹介\n４.国内外に広がる人脈　５.長く続く有意義な信頼関係　＋１.生涯学習',
  scheduleFrom: '7:15',
  scheduleTo: '9:15',
  schedule: [
    'オープンネットワーキング',
    'ビジターの歓迎と、リーダーシップチーム及びサポートチームの紹介',
    'BNIのコアバリュー、目的と概要',
    'ネットワーキングに関する学習コーナー',
    'BNIネットワーキングリーダーの発表',
    '新メンバー及び更新メンバーの歓迎',
    '全メンバーによるウィークリープレゼンテーション',
    'ビジターを再度歓迎/ビジターによるプレゼンテーション',
    'ビジネスブレイクアウトルーム',
    'バイスプレジデントによる報告',
    'メンバーシップ委員会による報告',
    '書記兼会計によるスピーカーローテーションの発表と、今週のスピーカーの紹介',
    'スピーカーによるメインプレゼンテーション',
    'リファーラルと推薦のことばの発表とビジターから良かった点のシェア',
    'リファーラル真正度の確認',
    '書記兼会計による報告',
    'プレジデントからビジターへの感謝',
    'BNIによる発表、お知らせ、特別レポート',
    '商品抽選（ビジターを同伴、またはリファーラルを提供したメンバーが対象）',
    '閉会'
  ].join('\n'),
  termsTitle: 'BNIの用語の説明',
  terms: 'チャプター\n1つの専門分野に1名で構成されたビジネスチームのことです。\n'
       + 'リファーラル\nメンバー間で交わされるビジネスや、信頼に基づく人脈の紹介です。\n'
       + 'サンキュー\nそのリファーラルによって発生した売上金額のことで、感謝の意を込めて、サンキューと呼んでいます。\n'
       + 'カテゴリー\n専門業種のことで、BNIでは1つの専門分野に対して加入できるのは1名のみであるという規定があります。',
  pname: '',
  ptext: '',
  photoFile: ''
};

// 表紙の初期値（チャプター名・期を入れる）。期を渡さなければ、いまの期
function defaultCover_(term) {
  var c = {}, k, t = term || roleTermOf_(new Date()), name = chapterInfo_().name;
  for (k in DEFAULT_COVER_) c[k] = DEFAULT_COVER_[k];
  c.title = 'BNI ' + name + ' chapter Member Book';
  c.term = t + '期';
  c.aboutTitle = 'BNI ' + chapterLabel_() + 'とは？';
  c.prole = chapterLabel_() + '\n第' + t + '期プレジデント';
  return c;
}

// --- プレジデント設定（期ごと）---
// プレジデントの氏名・肩書き・挨拶文は期ごとに変わるので、期ごとに持つ（スクリプトのプロパティ BNI_MB_PRESIDENT_24
// = { pname, prole, ptext } のように、期ごとに1つ。1つにまとめると、何年分もたまったときにプロパティの大きさの上限
// （1つ 9KB）を超えて保存できなくなるため）。表紙のほかの文章（理念・用語の説明など）は期によらず BNI_MB_COVER に1件。
// 期ごとの中身に項目が無いときは既定：氏名＝その期の「役職・チーム（半期ごと）」のプレジデント、
// 肩書き＝「○○チャプター\n第N期プレジデント」、挨拶文＝空。既定と同じ値は保存しない
// （期の番号・チャプター名・担当者を直しても、既定のままの項目は付いていく）
var COVER_PRES_PREFIX_ = 'BNI_MB_PRESIDENT_';
var COVER_PRES_MOVED_KEY_ = 'BNI_MB_PRESIDENTS_MOVED';    // 期ごとにする前の1件を引き継いだ印
var COVER_PRES_FIELDS_ = ['pname', 'prole', 'ptext'];

// 「24期」「第24期」「24」→ 24（読めなければ 0）
function coverTermNo_(v) {
  var m = String(v == null ? '' : v).normalize('NFKC').match(/\d+/), n = m ? parseInt(m[0], 10) : 0;
  return n > 0 ? n : 0;
}

// 保存してある表紙（期ごとにする前は、プレジデントの項目もこの1件に入っていた）
function coverSaved_() {
  var props = PropertiesService.getScriptProperties(), saved = {};
  try {
    var raw = props.getProperty(COVER_KEY_);
    if (raw) saved = JSON.parse(raw) || {};
  } catch (e) {
    console.warn('[MBOOK] 表紙設定の読み込みに失敗: ' + (e && e.message ? e.message : e));
  }
  // 旧版で個別のプロパティに保存していた分を引き継ぐ
  var old = { term: 'BNI_MB_TERM', pname: 'BNI_MB_PNAME', ptext: 'BNI_MB_PTEXT', photoFile: 'BNI_MB_COVER_PHOTO' };
  for (var o in old) {
    var v = props.getProperty(old[o]);
    if (v && !saved[o]) saved[o] = v;
  }
  return saved;
}

// ある期のプレジデント設定の既定
function coverPresDefault_(term) {
  var pname = '';
  try {
    var h = roleHoldersOfTerm_(roleHolderTerms_(), term);
    if (h.registered) pname = h.holders.president || '';      // その期として登録してある担当者だけ（前の期の方を出さない）
  } catch (e) {}
  return { pname: pname, prole: chapterLabel_() + '\n第' + term + '期プレジデント', ptext: '' };
}

// 期ごとのプレジデント設定 { 期: { pname, prole, ptext } }。
// まだ期ごとに保存していないときは、1件だけ保存していた設定を、その「期」（「23期」など。読めなければいまの期）のものとして読む
function coverPresidentTerms_() {
  var props = PropertiesService.getScriptProperties(), raw = props.getProperties() || {}, all = {}, found = false, k;
  for (k in raw) {
    if (k.indexOf(COVER_PRES_PREFIX_) !== 0) continue;
    var tail = k.slice(COVER_PRES_PREFIX_.length), n = parseInt(tail, 10);
    if (!(n > 0) || String(n) !== tail) continue;
    found = true;
    try { var v = JSON.parse(raw[k]); if (v && typeof v === 'object') all[n] = v; } catch (e) {}
  }
  if (found || raw[COVER_PRES_MOVED_KEY_]) return all;
  var old = coverSaved_();
  if (old.pname || old.ptext) {
    var t = coverTermNo_(old.term) || roleTermOf_(new Date()), def = coverPresDefault_(t), e = {};
    COVER_PRES_FIELDS_.forEach(function (k) {
      if (old[k] == null) return;                             // 無かった項目は既定
      var v = String(old[k]);
      if (v !== def[k]) e[k] = v;
    });
    all[t] = e;
  }
  return all;
}

// ある期のプレジデント設定（保存してある項目＋既定）
function coverPresidentOf_(term, all) {
  var e = (all || coverPresidentTerms_())[term] || null, def = coverPresDefault_(term);
  var out = { term: term, label: term + '期', range: roleTermLabel_(term), saved: !!e, defaults: def };
  COVER_PRES_FIELDS_.forEach(function (k) {
    out[k] = (e && Object.prototype.hasOwnProperty.call(e, k)) ? String(e[k] == null ? '' : e[k]) : def[k];
  });
  return out;
}

// 画面で選べる期：いまの期とその前後、保存してある期、選んでいる期
function coverPresidentList_(selTerm) {
  var all = coverPresidentTerms_(), cur = roleTermOf_(new Date()), set = {};
  [cur - 1, cur, cur + 1, coverTermNo_(selTerm)].forEach(function (t) { if (t > 0) set[t] = true; });
  Object.keys(all).forEach(function (t) { var n = parseInt(t, 10); if (n > 0) set[n] = true; });
  return Object.keys(set).map(function (t) { return parseInt(t, 10); }).sort(function (a, b) { return a - b; })
    .map(function (t) { var p = coverPresidentOf_(t, all); p.current = (t === cur); return p; });
}

// 期ごとのプレジデント設定を書く（all にある期を書き、all に無い期のプロパティは消す）。引き継いだ印も付ける
function coverPresidentWrite_(all) {
  var props = PropertiesService.getScriptProperties(), raw = props.getProperties() || {}, k;
  for (k in raw) {
    if (k.indexOf(COVER_PRES_PREFIX_) === 0 && !all[k.slice(COVER_PRES_PREFIX_.length)]) props.deleteProperty(k);
  }
  for (k in all) props.setProperty(COVER_PRES_PREFIX_ + k, JSON.stringify(all[k]));
  props.setProperty(COVER_PRES_MOVED_KEY_, '1');
}

// 期の番号を付け直したとき（チャプターの設定）、期ごとのプレジデント設定も同じだけずらす。
// 付け直す前の番号のうちに、担当者をずらす前に呼ぶ（期ごとにする前の1件を、前の番号の期・前の番号の担当者で読むため）
function coverShiftTerms_(delta) {
  delta = parseInt(delta, 10);
  if (!delta) return;
  var all = coverPresidentTerms_(), out = {};
  Object.keys(all).forEach(function (t) {
    var n = parseInt(t, 10);
    if (n > 0 && n + delta > 0) out[n + delta] = all[t];
  });
  coverPresidentWrite_(out);
}

// 表紙の中身（期を渡さなければ、いまの期のプレジデント設定）。termNo … どの期のプレジデント設定か
function getCoverInfo_(term) {
  var cover = {}, def = defaultCover_(), saved = coverSaved_(), k;
  for (k in def) cover[k] = def[k];
  for (k in saved) if (COVER_PRES_FIELDS_.indexOf(k) < 0 && k !== 'term' && k !== 'termNo') cover[k] = saved[k];
  var t = coverTermNo_(term) || roleTermOf_(new Date()), p = coverPresidentOf_(t);
  cover.termNo = t;
  cover.term = p.label;
  COVER_PRES_FIELDS_.forEach(function (f) { cover[f] = p[f]; });
  return cover;
}

// 表紙を保存する。プレジデントの項目（氏名・肩書き・挨拶文）は、term の期（無ければ cover.term の期・いまの期）に保存する。
// 渡さなかった項目（undefined）はそのまま
function saveCoverInfo_(cover, term) {
  if (!cover) return;
  var props = PropertiesService.getScriptProperties();
  // 先に期ごとの設定を作ってから（期ごとにする前の1件を引き継ぐため）、表紙の1件からプレジデントの項目を外す
  var all = coverPresidentTerms_();
  if (COVER_PRES_FIELDS_.some(function (k) { return cover[k] !== undefined; })) {
    var t = coverTermNo_(term) || coverTermNo_(cover.termNo) || coverTermNo_(cover.term) || roleTermOf_(new Date());
    var def = coverPresDefault_(t), e = all[t] || {};
    COVER_PRES_FIELDS_.forEach(function (k) {
      if (cover[k] === undefined) return;
      var v = String(cover[k] == null ? '' : cover[k]);
      if (v === def[k]) delete e[k]; else e[k] = v;
    });
    all[t] = e;
  }
  coverPresidentWrite_(all);
  var cur = coverSaved_(), shared = {}, k;
  var skip = function (x) { return COVER_PRES_FIELDS_.indexOf(x) >= 0 || x === 'term' || x === 'termNo'; };
  for (k in cur) if (!skip(k)) shared[k] = cur[k];
  for (k in cover) if (!skip(k) && cover[k] !== undefined) shared[k] = cover[k];
  props.setProperty(COVER_KEY_, JSON.stringify(shared));
}

// 画面から呼ぶ用。term … プレジデント設定を保存する期
function saveMemberBookCover(cover, term) {
  try {
    var t = coverTermNo_(term) || coverTermNo_(cover && cover.termNo) || roleTermOf_(new Date());
    saveCoverInfo_(cover, t);
    return { ok: true, message: '表紙の設定を保存しました（プレジデントの設定は' + t + '期のものとして保存）。',
             cover: getCoverInfo_(t), presidents: coverPresidentList_(t) };
  } catch (e) {
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 表紙の文章を初期値に戻す（期ごとのプレジデント設定は消さない）。term … 画面で選んでいる期
function resetMemberBookCoverText(term) {
  try {
    var props = PropertiesService.getScriptProperties();
    coverPresidentWrite_(coverPresidentTerms_());           // 期ごとにする前の1件を先に引き継ぐ
    var def = defaultCover_(), next = {}, keep = coverSaved_();
    for (var k in def) if (COVER_PRES_FIELDS_.indexOf(k) < 0 && k !== 'term') next[k] = def[k];
    if (keep.photoFile) next.photoFile = keep.photoFile;
    props.setProperty(COVER_KEY_, JSON.stringify(next));
    var t = coverTermNo_(term) || roleTermOf_(new Date());
    return { ok: true, message: '定型文を初期値に戻しました（プレジデントの設定はそのままです）。',
             cover: getCoverInfo_(t), presidents: coverPresidentList_(t) };
  } catch (e) {
    return { ok: false, message: '戻せませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// === メンバーブックHTML / TSV からの一括取り込み ===

// 現行のメンバーブックHTML（DATA_STORE に JSON が埋まっている）から取り込む
// 取り込んだ内容を名簿へ反映する共通処理。
// 氏名で照合し、**取り込み側に値がある項目だけ**を上書きする。
// 空欄は「消したい」ではなく「その形式に無い／未入力」なので触らない。
// 名簿にしか居ない人も消さない。
// （以前は取り込みのたびに名簿を丸ごと書き換えており、
//   ふりがな・No・写真・入会日など、取り込み元に無い項目が消えていた）
var MEMBER_FIELDS_ = ['no', 'cat', 'kana', 'title', 'company', 'role', 'memo', 'photoFile',
                      'comment', 'refer', 'collab', 'joinDate', 'renewDate', 'expireDate', 'position'];

function mergeMembersInto_(incoming) {
  var cur = getMemberMaster().members || [], byName = {}, i, f;
  for (i = 0; i < cur.length; i++) byName[normName_(cur[i].name)] = i;
  var updated = 0, added = 0, addedNames = [];

  for (var k = 0; k < (incoming || []).length; k++) {
    var e = incoming[k] || {};
    if (!e.name) continue;
    var key = normName_(e.name), idx = byName[key];
    if (idx === undefined) {
      var row = { no: '', cat: '', name: e.name, kana: '', title: '', company: '', role: '', memo: '',
                  photoFile: '', comment: '', refer: '', collab: '',
                  joinDate: '', renewDate: '', expireDate: '', position: '' };
      for (f = 0; f < MEMBER_FIELDS_.length; f++) {
        if (e[MEMBER_FIELDS_[f]]) row[MEMBER_FIELDS_[f]] = e[MEMBER_FIELDS_[f]];
      }
      cur.push(row); byName[key] = cur.length - 1; added++; addedNames.push(e.name);
      continue;
    }
    var m = cur[idx], hit = 0;
    for (f = 0; f < MEMBER_FIELDS_.length; f++) {
      var fld = MEMBER_FIELDS_[f];
      if (e[fld] && String(e[fld]) !== String(m[fld] || '')) { m[fld] = e[fld]; hit++; }
    }
    if (hit) updated++;
  }
  return { members: cur, updated: updated, added: added, addedNames: addedNames };
}

function importMemberBookHtml(base64) {
  try {
    if (!base64) return { ok: false, message: 'ファイルデータが空です。' };
    var html = Utilities.newBlob(Utilities.base64Decode(base64), 'text/html', 'mb.html').getDataAsString('UTF-8');
    var m = html.match(/<script[^>]*id=["']DATA_STORE["'][^>]*>([\s\S]*?)<\/script>/i);
    if (!m) return { ok: false, message: 'このHTMLに DATA_STORE が見つかりません。メンバーブック管理ツールで保存したHTMLをお選びください。' };
    var json = m[1].trim();
    if (!json) return { ok: false, message: 'DATA_STORE が空でした。メンバーを登録した状態で保存したHTMLをお使いください。' };
    var data = JSON.parse(json);
    var src = data.members || [], members = [];
    for (var i = 0; i < src.length; i++) {
      var s = src[i] || {};
      members.push({
        no: '', cat: s.cat || '', name: s.name || '', kana: '', title: s.title || '',
        company: s.company || '', role: '', memo: '', photoFile: '',
        comment: s.comment || '', refer: s.refer || '', collab: s.collab || '',
        joinDate: '', renewDate: '', expireDate: ''
      });
    }
    if (!members.length) return { ok: false, message: 'メンバーが0件でした。' };

    // 表紙は、HTML側に値が入っている項目だけ反映する（空で上書きしない）
    var cover = null;
    if (data.cover) {
      cover = {};
      if (data.cover.term)  cover.term  = data.cover.term;
      if (data.cover.pname) cover.pname = data.cover.pname;
      if (data.cover.ptext) cover.ptext = data.cover.ptext;
      if (!Object.keys(cover).length) cover = null;
    }

    var merged = mergeMembersInto_(members);
    memberBackup_('メンバーブックHTMLの取り込み');
    var res = saveMemberMaster(merged.members, cover);
    if (!res.ok) return res;
    console.log('[MEMBER] html merged updated=' + merged.updated + ' added=' + merged.added);
    var msg = 'メンバーブックHTMLから ' + members.length + '名を読み取りました。'
            + '（更新 ' + merged.updated + '名 / 新規 ' + merged.added + '名 / 名簿は計 ' + merged.members.length + '名）\n'
            + 'HTMLに無い項目（No・ふりがな・役職・メモ・写真ファイル名・入会日・更新日・更新期限日・会社での役職）は'
            + 'そのまま残しています。\n取り込む前の名簿は控えてあります（「取り込む前に戻す」で戻せます）。';
    if (merged.addedNames.length) msg += '\n新しく追加: ' + merged.addedNames.join('、');
    return { ok: true, message: msg, updated: merged.updated, added: merged.added };
  } catch (e) {
    console.error('[MEMBER] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '取り込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// タブ区切りテキストから取り込む（見出し行は任意）
// 列: No / 業種区分 / 氏名 / ふりがな / カテゴリー / 会社名 / 役職 / メモ / 写真 / 一言 / 紹介 / 協業 / 入会日 / 更新日 / 更新期限日 / 会社での役職
function importMemberTsv(text, replaceAll) {
  try {
    if (!text || !String(text).trim()) return { ok: false, message: '貼り付けられたデータが空です。' };
    var lines = String(text).replace(/\r/g, '').split('\n'), members = [];
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (!line.trim()) continue;
      var c = line.split('\t');
      var name = (c[2] || '').trim();
      if (!name || name === '氏名') continue;       // 見出し行・空行を飛ばす
      members.push({
        no: (c[0] || '').trim(), cat: (c[1] || '').trim(), name: name, kana: (c[3] || '').trim(),
        title: (c[4] || '').trim(), company: (c[5] || '').trim(), role: (c[6] || '').trim(),
        memo: (c[7] || '').trim(), photoFile: (c[8] || '').trim(), comment: (c[9] || '').trim(),
        refer: (c[10] || '').trim(), collab: (c[11] || '').trim(),
        joinDate: (c[12] || '').trim(), renewDate: (c[13] || '').trim(), expireDate: (c[14] || '').trim(),
        position: (c[15] || '').trim()
      });
    }
    if (!members.length) return { ok: false, message: '取り込める行がありませんでした。タブ区切りで、3列目が氏名になっているかご確認ください。' };
    var out, note;
    if (replaceAll) {
      out = members;
      note = '名簿を貼り付けた内容だけに置き換えました';
    } else {
      // 空欄で既存の内容を消さないよう、値のある列だけを反映する
      var merged = mergeMembersInto_(members);
      out = merged.members;
      note = '更新 ' + merged.updated + '名 / 新規 ' + merged.added + '名。'
           + '貼り付けた表で空欄だった列は、元の内容を残しています';
    }
    memberBackup_('表計算からの貼り付け');
    var res = saveMemberMaster(out, null);
    if (!res.ok) return res;
    return { ok: true, message: members.length + '行を取り込みました（' + note + ' / 名簿は計 ' + out.length + '名）。'
             + '\n取り込む前の名簿は控えてあります（「取り込む前に戻す」で戻せます）。',
             imported: members.length, total: out.length };
  } catch (e) {
    console.error('[MEMBER] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '取り込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// === メンバーリスト(OCR)の読み取り結果をメンバー名簿へ統合する ===
// PDFに載っている項目（No・ふりがな・業種区分・カテゴリー・会社名・役職・メモ）だけを上書きし、
// 写真・一言コメント・紹介/協業・日付・会社での役職など、PDFに無い項目は既存の値を残す。
//
// 名簿の行は【氏名】で探す（番号では探さない）。番号はPDFの番号に直すだけで、行の氏名は書き換えない。
// 以前は番号で行を探して氏名を書き換えていたため、メンバーが増えて番号がずれると、
// ある方の行（写真・一言・日付など）が次々と別の方の名前に書き換わってしまった。
//   1. 同じ氏名（空白・全角半角の違いは無視）
//   2. 1文字だけ違う氏名（読み取りの誤り・異体字。3文字以上で、名簿に当てはまる方が1人だけのとき）… 氏名は名簿のまま
//   3. どちらも無ければ、新しい方として足す
// PDFに載っていない方は消さない（知らせるだけ）。

// 2つの氏名（normName_ 済み）が、1文字だけ違う（置き換え・足りない・多い）か
function ocrNameNear_(a, b) {
  if (a === b || Math.abs(a.length - b.length) > 1) return false;
  var i, d = 0;
  if (a.length === b.length) {
    for (i = 0; i < a.length; i++) if (a.charAt(i) !== b.charAt(i) && ++d > 1) return false;
    return d === 1;
  }
  var s = a.length < b.length ? a : b, l = a.length < b.length ? b : a, j = 0;
  for (i = 0; i < l.length; i++) {
    if (j < s.length && s.charAt(j) === l.charAt(i)) j++;
    else if (++d > 1) return false;
  }
  return true;
}

var OCR_FIELDS_ = [['no', 'No'], ['kana', 'ふりがな'], ['cat', '業種区分'], ['title', 'カテゴリー'],
                   ['company', '会社名'], ['role', '役職'], ['memo', 'メモ']];

// 読み取った方（extracted）を名簿（cur）に当てはめる。名簿はまだ変えない
function ocrMergePlan_(extracted, cur) {
  var used = {}, keys = cur.map(function (m) { return normName_(m.name); }), pending = [];
  var plan = { matches: [], adds: [], noName: [], missing: [] };
  (extracted || []).forEach(function (e) {
    var key = normName_(e.name);
    if (!key) { plan.noName.push(e); return; }
    for (var i = 0; i < cur.length; i++) {
      if (!used[i] && keys[i] === key) { used[i] = true; plan.matches.push({ idx: i, e: e, near: false }); return; }
    }
    pending.push(e);
  });
  pending.forEach(function (e) {
    var key = normName_(e.name), cand = [];
    if (key.length >= 3) {
      for (var i = 0; i < cur.length; i++) if (!used[i] && keys[i].length >= 3 && ocrNameNear_(key, keys[i])) cand.push(i);
    }
    if (cand.length === 1) { used[cand[0]] = true; plan.matches.push({ idx: cand[0], e: e, near: true }); }
    else plan.adds.push(e);
  });
  plan.matches.forEach(function (x) {
    var m = cur[x.idx];
    x.changes = [];
    OCR_FIELDS_.forEach(function (f) {
      var v = x.e[f[0]];
      if (v && String(v) !== String(m[f[0]] || '')) x.changes.push({ key: f[0], field: f[1], from: String(m[f[0]] || ''), to: String(v) });
    });
  });
  for (var n = 0; n < cur.length; n++) if (!used[n]) plan.missing.push(n);
  return plan;
}

// 当てはめた結果を名簿（cur）に書き込み、No順に並べる。番号が重なった方を返す
function ocrApplyPlan_(plan, cur) {
  plan.matches.forEach(function (x) {
    x.changes.forEach(function (c) { cur[x.idx][c.key] = c.to; });
  });
  plan.adds.forEach(function (e) {
    cur.push({ no: e.no, cat: e.cat, name: e.name, kana: e.kana, title: e.title, company: e.company, role: e.role, memo: e.memo,
               photoFile: '', comment: '', refer: '', collab: '', joinDate: '', renewDate: '', expireDate: '', position: '' });
  });
  // PDFに載っていた順（No順）に並べ替える。掲載順＝発表順の意味を持つため
  cur.sort(function (a, b) {
    var na = parseInt(a.no, 10), nb = parseInt(b.no, 10);
    if (isNaN(na) && isNaN(nb)) return 0;
    if (isNaN(na)) return 1;
    if (isNaN(nb)) return -1;
    return na - nb;
  });
  var byNo = {};
  cur.forEach(function (m) { if (m.no) (byNo[m.no] = byNo[m.no] || []).push(m.name); });
  return Object.keys(byNo).filter(function (k) { return byNo[k].length > 1; })
    .map(function (k) { return { no: k, names: byNo[k] }; });
}

// 画面・確認用に、当てはめた結果をまとめる（名簿は変えない）
function ocrPlanSummary_(extracted, cur) {
  var plan = ocrMergePlan_(extracted, cur);
  var copy = cur.map(function (m) { var o = {}; for (var k in m) o[k] = m[k]; return o; });
  var dup = ocrApplyPlan_(ocrMergePlan_(extracted, copy), copy);
  return {
    total: (extracted || []).length,
    updates: plan.matches.filter(function (x) { return x.changes.length; }).map(function (x) {
      return { name: cur[x.idx].name, pdfName: x.e.name, near: x.near,
               changes: x.changes.map(function (c) { return { field: c.field, from: c.from, to: c.to }; }) };
    }),
    unchanged: plan.matches.filter(function (x) { return !x.changes.length; }).length,
    near: plan.matches.filter(function (x) { return x.near; }).map(function (x) { return { no: x.e.no, pdf: x.e.name, roster: cur[x.idx].name }; }),
    adds: plan.adds.map(function (e) { return { no: e.no, name: e.name, cat: e.cat }; }),
    missing: plan.missing.map(function (i) { return { no: cur[i].no, name: cur[i].name }; }),
    noName: plan.noName.map(function (e) { return { no: e.no }; }),
    dupNos: dup
  };
}

// 確認用の文（ダイアログを使わない取り込みの確認・反映したあとの知らせ）
function ocrSummaryText_(sm, done) {
  var list = function (a, f) { return a.slice(0, 12).map(f).join('、') + (a.length > 12 ? ' ほか' + (a.length - 12) + '名' : ''); };
  var t = '読み取り ' + sm.total + '名：変わる方 ' + sm.updates.length + '名 / 変わらない方 ' + sm.unchanged + '名 / 新しく足す方 ' + sm.adds.length + '名';
  if (sm.adds.length) t += '\n・新しく足す方（名簿に同じ氏名が無い）: ' + list(sm.adds, function (x) { return 'No' + (x.no || '?') + ' ' + x.name; });
  if (sm.near.length) t += '\n・氏名が1文字違う方（名簿の氏名のまま' + (done ? '更新しました' : '更新します') + '）: '
    + list(sm.near, function (x) { return 'PDF「' + x.pdf + '」→ 名簿「' + x.roster + '」'; });
  if (sm.missing.length) t += '\n・PDFに無い方（消しません）: ' + list(sm.missing, function (x) { return (x.no ? 'No' + x.no + ' ' : '') + x.name; });
  if (sm.dupNos.length) t += '\n・番号が重なる方: ' + list(sm.dupNos, function (x) { return 'No' + x.no + '（' + x.names.join('・') + '）'; });
  if (sm.noName.length) t += '\n・氏名を読み取れなかった行（飛ばします）: ' + list(sm.noName, function (x) { return 'No' + (x.no || '?'); });
  return t;
}

// 読み取った方を名簿に反映する（確認のあと、画面から呼ぶ）。いまの名簿は、書き換える前に控えておく
function mergeMembersFromOcr_(extracted, label) {
  try {
    var cur = getMemberMaster().members || [];
    var sm = ocrPlanSummary_(extracted, cur);
    var plan = ocrMergePlan_(extracted, cur);
    memberBackup_(label || 'メンバーリスト(OCR)の取り込み');
    ocrApplyPlan_(plan, cur);
    var res = saveMemberMaster(cur, null);
    if (!res.ok) return res;
    var msg = '「メンバー名簿」に反映しました（名簿は計 ' + cur.length + '名）。\n' + ocrSummaryText_(sm, true)
            + '\n写真・一言コメント・日付など、PDFに無い項目は残しています。'
            + '\n取り込む前の名簿は控えてあります（メンバー名簿の画面の「取り込む前に戻す」で戻せます）。';
    if (typeof roleRosterAfterImport_ === 'function') msg += roleRosterAfterImport_();
    console.log('[OCR] merged updated=' + sm.updates.length + ' added=' + sm.adds.length);
    return { ok: true, message: msg, updated: sm.updates.length, added: sm.adds.length, total: cur.length, summary: sm };
  } catch (e) {
    console.error('[OCR] merge ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '名簿への反映に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 画面の「名簿に反映」。読み取った内容（確認した一覧のもと）を受け取って反映する
function applyMemberListOcr(extracted) {
  var lock = LockService.getScriptLock();
  try {
    if (!extracted || !extracted.length) return { ok: false, message: '反映する内容がありません。もう一度読み取ってください。' };
    if (!lock.tryLock(20000)) return { ok: false, message: 'ほかの方が名簿を保存中です。少し待ってから、もう一度押してください。' };
    return mergeMembersFromOcr_(extracted);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// === 取り込む前の名簿の控え ===
// 名簿をまとめて書き換える取り込み（メンバーリスト(OCR)・Spreading・メンバーブックHTML・表計算からの貼り付け・
// BNI公式レポート）の前に、いまの名簿を隠しシートに写しておく。
// 「取り込む前に戻す」で、その名簿と入れ替える（もう一度押すと、戻す前の名簿に戻る）
var MEMBER_BACKUP_SHEET_ = 'メンバー名簿_取り込み前';
var MEMBER_BACKUP_KEY_ = 'BNI_MEMBER_BACKUP';           // { at: 控えた日時, label: 何の前か, rows: 人数 }

function memberBackupSheet_() {
  var ss = getSS_(), sh = ss.getSheetByName(MEMBER_BACKUP_SHEET_);
  if (sh) return sh;
  var active = null;
  try { active = ss.getActiveSheet(); } catch (e) {}
  sh = ss.insertSheet(MEMBER_BACKUP_SHEET_);
  try { sh.hideSheet(); } catch (e) {}
  try { if (active) ss.setActiveSheet(active); } catch (e) {}   // 作った控えのシートに画面が移らないように
  return sh;
}
function memberCopyGrid_(from, to) {
  var data = from.getDataRange().getValues();
  to.clear();
  if (data.length && data[0].length) to.getRange(1, 1, data.length, data[0].length).setValues(data);
  return Math.max(0, data.length - 1);
}
function memberBackup_(label) {
  var rows = memberCopyGrid_(ensureMemberSheet_(), memberBackupSheet_());
  var info = { at: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'), label: label || '取り込み', rows: rows };
  PropertiesService.getScriptProperties().setProperty(MEMBER_BACKUP_KEY_, JSON.stringify(info));
  return info;
}
function getMemberBackupInfo_() {
  try {
    var info = JSON.parse(PropertiesService.getScriptProperties().getProperty(MEMBER_BACKUP_KEY_) || 'null');
    if (!info || !getSS_().getSheetByName(MEMBER_BACKUP_SHEET_)) return null;
    return info;
  } catch (e) { return null; }
}
// 画面の「取り込む前に戻す」。控えの名簿といまの名簿を入れ替える
function restoreMemberBackup() {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(20000)) return { ok: false, message: 'ほかの方が名簿を保存中です。少し待ってから、もう一度押してください。' };
    var info = getMemberBackupInfo_();
    if (!info) return { ok: false, message: '戻せる名簿の控えがありません。' };
    var ss = getSS_(), bk = ss.getSheetByName(MEMBER_BACKUP_SHEET_), sh = ensureMemberSheet_();
    if (bk.getLastRow() < 1) return { ok: false, message: '控えの名簿が空です。' };
    var tmp = sh.getDataRange().getValues();
    var rows = memberCopyGrid_(bk, sh);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#f2f6ff');
    sh.setFrozenRows(1);
    bk.clear();
    if (tmp.length && tmp[0].length) bk.getRange(1, 1, tmp.length, tmp[0].length).setValues(tmp);
    var now = { at: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'), label: '「取り込む前に戻す」を押す前', rows: Math.max(0, tmp.length - 1) };
    PropertiesService.getScriptProperties().setProperty(MEMBER_BACKUP_KEY_, JSON.stringify(now));
    return { ok: true, message: info.at + '（' + info.label + '）の名簿に戻しました（' + rows + '名）。'
             + '\nもう一度押すと、戻す前の名簿に戻ります。', backup: now };
  } catch (e) {
    console.error('[MEMBER] restore ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '戻せませんでした: ' + (e && e.message ? e.message : e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// メモ欄に「ビジターホスト」等が入っているメンバーを、ビジターホスト設定に反映する
function applyVisitorHostsFromMemo() {
  try {
    var members = getMemberMaster().members || [], hosts = [], names = [];
    for (var i = 0; i < members.length; i++) {
      // メモ欄と役職欄の両方を見る。OCR取込はメモ欄に、Spreading取込は役職欄に入るため。
      var memo = (members[i].memo || '') + ' ' + (members[i].role || '');
      // 「ビジターホスト」「ビジホス」などの表記ゆれを拾う
      if (!/ビジ(ター)?ホス(ト)?/.test(memo)) continue;
      if (!members[i].no) continue;
      hosts.push(String(members[i].no)); names.push(members[i].name);
    }
    if (!hosts.length) return { ok: false, message: 'メモ欄に「ビジターホスト」の記載があるメンバーが見つかりませんでした。' };
    saveVisitorHosts(hosts);
    return { ok: true, message: hosts.length + '名をビジターホストに設定しました。\n' + names.join('、'),
             count: hosts.length, names: names };
  } catch (e) {
    return { ok: false, message: '反映に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 旧「メンバーリスト」シートの内容をメンバー名簿へ移し、旧シートを隠す
function migrateMemberListSheet() {
  try {
    var ss = getSS_(), old = ss.getSheetByName('メンバーリスト');
    if (!old) return { ok: false, message: '「メンバーリスト」シートはありません。移行は不要です。' };
    var data = old.getDataRange().getValues(), src = [];
    for (var i = 1; i < data.length; i++) {
      var no = String(data[i][0] == null ? '' : data[i][0]).trim();
      var nm = String(data[i][1] == null ? '' : data[i][1]).trim();
      if (no && nm) src.push({ no: no, name: normalizeSpace(nm), kana: '', cat: '', title: '', company: '', role: '', memo: '' });
    }
    if (!src.length) {
      old.hideSheet();
      return { ok: true, message: '「メンバーリスト」は空だったため、非表示にしました。' };
    }
    var res = mergeMembersFromOcr_(src, '旧メンバーリストからの移行');
    if (!res.ok) return res;
    old.hideSheet();
    return { ok: true, message: res.message + '\n旧「メンバーリスト」シートは非表示にしました（データは残っています）。' };
  } catch (e) {
    console.error('[MEMBER] migrate ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '移行に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// === BNI公式レポート（Excel）の取り込み ===
// 「メンバーシップ期間レポート」から入会日、「会費レポート」から更新期限日を読む。
// どちらも拡張子は .xls だが中身は SpreadsheetML（XML）なので、そのまま解析できる。

// SpreadsheetML → 行の二次元配列
function parseSpreadsheetMlRows_(xml) {
  var rows = [], rowRe = /<Row[^>]*>([\s\S]*?)<\/Row>/g, m;
  while ((m = rowRe.exec(xml)) !== null) {
    var cells = [], cellRe = /<Cell[^>]*>([\s\S]*?)<\/Cell>|<Cell[^>]*\/>/g, c;
    while ((c = cellRe.exec(m[1])) !== null) {
      var inner = c[1] || '';
      var d = inner.match(/<Data[^>]*>([\s\S]*?)<\/Data>/);
      cells.push(d ? unescapeXml_(d[1].replace(/<[^>]+>/g, '')).trim() : '');
    }
    rows.push(cells);
  }
  return rows;
}

// 'YYYY-MM-DDT00:00:00.000' → 'YYYY/MM/DD'
function isoToDateStr_(v) {
  var m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? (m[1] + '/' + m[2] + '/' + m[3]) : '';
}

function rowHas_(row, text) {
  for (var i = 0; i < row.length; i++) if (String(row[i]).indexOf(text) !== -1) return true;
  return false;
}

// files = [{name, base64}] を1つ以上。レポートの種類は中身の見出しから自動判別する
function importMembershipReports(files) {
  try {
    if (!files || !files.length) return { ok: false, message: 'ファイルが選択されていません。' };
    var join = {}, expire = {}, states = {}, kinds = [];

    for (var f = 0; f < files.length; f++) {
      var xml = Utilities.newBlob(Utilities.base64Decode(files[f].base64), 'text/xml', files[f].name || 'r.xls')
                         .getDataAsString('UTF-8');
      var rows = parseSpreadsheetMlRows_(xml);
      if (!rows.length) continue;

      // --- メンバーシップ期間レポート（姓・名・Recent Start date）---
      var hi = -1;
      for (var i = 0; i < rows.length; i++) if (rowHas_(rows[i], 'Recent Start date')) { hi = i; break; }
      if (hi >= 0) {
        var n1 = 0;
        for (var r = hi + 1; r < rows.length; r++) {
          var row = rows[r];
          if (row.length < 7 || !row[1]) continue;
          var nm = (row[1] + ' ' + row[2]).trim();
          var d = isoToDateStr_(row[6]);          // Recent Start date = 最新の入会日
          if (!nm || !d) continue;
          join[normName_(nm)] = d; n1++;
        }
        kinds.push('メンバーシップ期間レポート（' + n1 + '名）');
        continue;
      }

      // --- 会費レポート（メンバー名・更新日）---
      var hj = -1;
      for (var j = 0; j < rows.length; j++) if (rowHas_(rows[j], 'AutoRenewal')) { hj = j; break; }
      if (hj >= 0) {
        var n2 = 0;
        for (var k = hj + 1; k < rows.length; k++) {
          var rw = rows[k];
          if (rw.length < 6) continue;
          var name2 = String(rw[1] || '').trim();
          var d2 = isoToDateStr_(rw[5]);          // 更新日 = 会費がいつまでか
          if (!name2 || name2 === 'メンバー名' || !d2) continue;
          expire[normName_(name2)] = d2;
          states[normName_(name2)] = String(rw[4] || '').trim();
          n2++;
        }
        kinds.push('会費レポート（' + n2 + '名）');
        continue;
      }
      kinds.push('「' + (files[f].name || '不明') + '」は種類を判別できませんでした');
    }

    if (!Object.keys(join).length && !Object.keys(expire).length) {
      return { ok: false, message: 'レポートの内容を読み取れませんでした。BNI公式サイトから出力した .xls をそのままお選びください。\n' + kinds.join('\n') };
    }

    // 名簿へ反映（氏名で照合。日付以外は触らない）
    var members = getMemberMaster().members || [];
    var setJoin = 0, setExp = 0, pending = [], unmatched = {};
    for (var q in join) unmatched[q] = true;
    for (var q2 in expire) unmatched[q2] = true;

    for (var mi = 0; mi < members.length; mi++) {
      var key = normName_(members[mi].name);
      if (join[key]) { members[mi].joinDate = join[key]; setJoin++; delete unmatched[key]; }
      if (expire[key]) {
        members[mi].expireDate = expire[key]; setExp++; delete unmatched[key];
        if (states[key] && states[key].indexOf('Active') === -1) pending.push(members[mi].name + '（' + states[key] + '）');
      }
    }
    memberBackup_('BNI公式レポートの取り込み');
    var res = saveMemberMaster(members, null);
    if (!res.ok) return res;

    var left = [];
    for (var u in unmatched) left.push(u);
    var msg = kinds.join(' / ') + ' を読み込みました。\n'
            + '入会日 ' + setJoin + '名 / 更新期限日 ' + setExp + '名 を更新しました。';
    if (left.length) msg += '\n※ 名簿に見つからなかった氏名 ' + left.length + '件: ' + left.slice(0, 10).join('、');
    if (pending.length) msg += '\n※ 更新手続き中の方 ' + pending.length + '名: ' + pending.slice(0, 10).join('、');
    console.log('[REPORT] join=' + setJoin + ' expire=' + setExp + ' unmatched=' + left.length);
    return { ok: true, message: msg, joinSet: setJoin, expireSet: setExp, unmatched: left, pending: pending };
  } catch (e) {
    console.error('[REPORT] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '取り込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}
