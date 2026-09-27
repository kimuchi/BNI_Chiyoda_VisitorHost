// マニュアル（MANUAL.md / manual.html）に載せる画面のスクリーンショットを、架空のデータで撮る。
//
//   node tools/make_manual_shots.js <routine.json> [出力先のフォルダ（既定: docs/images）] [撮る画面の名前…]
//
// 撮った画面は、幅 960px までに縮めて WebP（画質82）にする（Python の Pillow を使う）。
// MANUAL.md から ![説明](docs/images/名前.webp) で載せ、python3 tools/build_manual.py で manual.html に埋め込む。
//
// ・ルーティンチェックシート（routine.json）は、項目の並び（内容・担当・期日）だけを使う。
//   開催日の列の値はすべて消して、架空の値を入れる。備考などの「○○さん」も「見本さん」にする
// ・名簿・参加者・担当者・チーム・スピーカーローテーションは、すべて架空の方にする
//   （実在の方の名前・写真は写さない。実在のメンバーと同じ姓も使わない）
// ・画面は本物の HTML を、サーバーの代わりの関数表（google.script.run）で動かし、Chromium で撮る。
//   サーバーの返事は、本物のサーバーのコード（lib_role_fixture.js で動かすもの）で作る
// ・「今日」は 2026/09/26（次の定例会は 9/30）

const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const { makeRoleServer } = require('./lib_role_fixture');

const ROOT = path.join(__dirname, '..');
const ROUTINE_PATH = process.argv[2];
const OUT = path.resolve(process.argv[3] || path.join(ROOT, 'docs', 'images'));
const ONLY = process.argv.slice(4);
if (!ROUTINE_PATH) { console.error('使い方: node tools/make_manual_shots.js <routine.json> [出力先] [画面の名前…]'); process.exit(1); }

// ===================== 架空のデータ =====================
// 名簿（No・氏名・業種区分・カテゴリー・会社名）
const FAKE_MEMBERS = [
  ['青木 一郎', '企業サポート', '税理士', '青木会計事務所'], ['上田 花子', '企業サポート', '社会保険労務士', 'うえだ社労士事務所'],
  ['高橋 健太', '企業サポート', 'ITコンサルタント', '株式会社タカハシIT'], ['伊藤 美咲', '企業サポート', '司法書士', '伊藤司法書士事務所'],
  ['野村 大輔', '研修・教育', '研修講師', 'ノムラ研修'], ['中村 由美', '研修・教育', '英会話講師', 'なかむら英語教室'],
  ['小林 翔太', '不動産関連', '不動産売買仲介', '小林不動産'], ['加藤 真理', '不動産関連', '賃貸管理', 'カトウ住宅管理'],
  ['大庭 誠', '不動産関連', '土地家屋調査士', '大庭測量'], ['山田 恵', '建築・住まい', 'リフォーム', '山田リフォーム'],
  ['佐々木 拓也', '建築・住まい', '外壁塗装', 'ササキ塗装'], ['松本 陽子', '建築・住まい', 'インテリアコーディネーター', '松本デザイン室'],
  ['井上 直樹', '建築・住まい', '設計事務所', '井上設計'], ['林 彩', 'プロモーション', 'Web制作', 'はやしウェブ工房'],
  ['斎藤 浩二', 'プロモーション', '動画制作', 'サイトウ映像'], ['清水 麻衣', 'プロモーション', 'カメラマン', 'しみず写真室'],
  ['森 修', '暮らし・生活', '保険代理店', '森保険サービス'], ['池田 香織', '暮らし・生活', '家事代行', 'いけだ家事サポート'],
  ['橋本 亮', '暮らし・生活', '自動車販売', 'ハシモト自動車'], ['阿部 智子', '暮らし・生活', '行政書士', '阿部行政書士事務所'],
  ['石川 剛', '美容と健康', '整体院', 'いしかわ整体'], ['山下 舞', '美容と健康', 'ヘアサロン', 'サロン・ド・ヤマシタ'],
  ['中島 聡', '美容と健康', '歯科医院', 'なかじま歯科'], ['石井 裕子', '美容と健康', 'エステサロン', 'エステ石井'],
  ['小川 学', '飲食・エンタメ', '和食店', '割烹おがわ'], ['前田 愛', '飲食・エンタメ', 'カフェ', 'カフェ前田'],
  ['岡田 隆', '飲食・エンタメ', 'イベント企画', 'オカダ企画'], ['長谷川 真由美', '企業サポート', '経営コンサルタント', '長谷川経営研究所'],
  ['後藤 淳', '不動産関連', '不動産投資コンサル', 'ゴトウ資産'], ['近江 京子', '暮らし・生活', '終活カウンセラー', 'こんどう相談室'],
  ['西村 悟', '建築・住まい', '電気工事', 'ニシムラ電設'], ['福田 明美', 'プロモーション', '印刷', '福田印刷'],
  ['太田 勇気', '企業サポート', '弁護士', '太田法律事務所'], ['三浦 千尋', '美容と健康', 'ネイルサロン', 'ネイル三浦'],
  ['藤井 優', '飲食・エンタメ', 'ワインバー', 'バー藤井'], ['松尾 奈々', '研修・教育', '学習塾', 'まつだ塾'],
].map((m, i) => ({ no: String(i + 1), name: m[0], cat: m[1], title: m[2], company: m[3],
                     expireDate: '2027/' + String(1 + (i % 12)).padStart(2, '0') + '/15' }));
const N = (i) => FAKE_MEMBERS[i].name;

// 担当者（23期・24期）とチーム
const HOLDERS_23 = { president: N(0), vice: N(1), secretary: N(2), vhc: N(3), mentor: N(4), ec: N(5), web: N(13),
  support: N(16), training: N(20), event: N(25), bcp: N(18), spreading: N(15), gbc: N(32) };
const HOLDERS_24 = { president: N(6), vice: N(7), secretary: N(8), vhc: N(9), mentor: N(10), ec: N(11), web: N(14),
  support: N(17), training: N(21), event: N(24), bcp: N(26), spreading: N(31), gbc: N(27) };
const TEAMS_24 = [
  { key: 'membership', name: 'メンバーシップ委員会', members: [{ name: N(12) }, { name: N(19) }, { name: N(29) }] },
  { key: 'role:ec', name: 'エデュケーションコーディネーター', members: [{ name: N(35) }] },
  { key: 'role:vhc', name: 'ビジターホスト', members: [{ name: N(28), note: 'サブリーダー' }, { name: N(3) }, { name: N(22) },
    { name: N(23) }, { name: N(30) }, { name: N(33) }, { name: N(34) }] },
  { key: 'role:web', name: 'Webチーム', members: [{ name: N(13) }, { name: N(15) }] },
  { key: 'role:mentor', name: 'メンターコーディネーター', members: [{ name: N(4), note: 'メンター' }, { name: N(0), note: 'メンター' }] },
  { key: 'role:event', name: 'イベント委員＆1to1促進委員', members: [{ name: N(25) }] },
  { key: 'role:training', name: 'トレーニング促進委員', members: [{ name: N(20) }] },
];

// ===================== ルーティンチェックシート（項目の並びだけ使う）=====================
function scrubRoutine(src) {
  const out = {};
  for (const name of Object.keys(src)) {
    const g = src[name].map((r) => r.slice());
    const head = g[0] || [];
    const first = head.findIndex((v) => /^\d{4}\/\d{2}\/\d{2}$/.test(String(v)));
    const from = first < 0 ? head.length : first;
    for (let r = 0; r < g.length; r++) {
      for (let c = 0; c < g[r].length; c++) {
        if (c >= from) { if (r >= 2) g[r][c] = ''; continue; }
        g[r][c] = String(g[r][c] == null ? '' : g[r][c]).replace(/[一-龥々ぁ-んァ-ヶー]{1,6}さん/g, '見本さん');
      }
    }
    out[name] = g;
  }
  return out;
}
// その日・その項目（B〜E列の名前）の欄に値を入れる
function setRoutine(R, date, title, value, parent) {
  for (const name of Object.keys(R)) {
    const g = R[name], c = (g[0] || []).indexOf(date);
    if (c < 0) continue;
    let under = '';
    for (let r = 2; r < g.length; r++) {
      const labels = [g[r][2], g[r][3], g[r][4]].map((v) => String(v || '').replace(/\s/g, ''));
      if (labels[1] && !labels[2]) under = labels[1];
      if (labels.some((v) => v && v === title.replace(/\s/g, '')) && (!parent || under === parent || labels[1] === parent)) {
        g[r][c] = value;
        return true;
      }
    }
  }
  return false;
}

const ROUTINE = scrubRoutine(JSON.parse(fs.readFileSync(ROUTINE_PATH, 'utf8')));
const PREV = '2026/09/23', NEXT = '2026/09/30';
const surname = (n) => n.split(' ')[0] + 'さん';
[
  ['一般規定', '2番'], ['遅刻・欠席担当(7:00開始)', surname(N(17))], ['名札・バッチの注意', 'なし'], ['リファーラルの注意', 'なし'],
  ['代理', surname(N(16))], ['欠席', surname(N(26))], ['リージョン参加者', 'なし'], ['体験談', 'なし'],
  ['エデュケーション', surname(N(5))], ['審査中カテゴリー', 'なし'], ['新入会', 'なし'], ['更新式(更新メンバー)', 'なし'],
  ['ウィークリープレゼン', '建築・住まい　10番　' + surname(N(9))], ['募集カテゴリー', '税理士、司法書士、Webデザイナー'],
  ['退会者', 'なし'], ['開放カテゴリー', 'なし'], ['メインプレゼン', '①' + N(22) + 'さん　②' + N(23) + 'さん'],
  ['推薦の言葉', N(0) + 'さん→' + N(1) + 'さん'], ['2分30秒プレゼン', N(10) + 'さん'], ['BNI目的と概要', 'Givers Gain®（与える者は与えられる）'],
  ['一般規定', '2番'],
].forEach(([t, v]) => setRoutine(ROUTINE, PREV, t, v));
// 9/30：バイスプレジデントの項目の一部と、事前MTGの共有事項（役職ごと）は入力済み
[
  ['一般規定', '3番'], ['遅刻・欠席担当(7:00開始)', surname(N(17))], ['名札・バッチの注意', 'なし'],
  ['メインプレゼン', '①' + N(10) + 'さん　②' + N(11) + 'さん'], ['2分30秒プレゼン', N(12) + 'さん'],
  ['推薦の言葉', '①' + N(0) + 'さん→' + N(10) + 'さん　②' + N(1) + 'さん→' + N(11) + 'さん\nアフター：' + N(2) + 'さん→' + N(12) + 'さん'],
].forEach(([t, v]) => setRoutine(ROUTINE, NEXT, t, v));

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'manual-shots-'));
const routineFile = path.join(tmp, 'routine.json'), membersFile = path.join(tmp, 'members.json');
fs.writeFileSync(routineFile, JSON.stringify(ROUTINE));
fs.writeFileSync(membersFile, JSON.stringify(FAKE_MEMBERS));
const S = makeRoleServer(routineFile, membersFile);
const F = S.F;
// ウェブアプリのトップページの一覧（webapp_srv.js）は、別の場所で読む（検査用のシートの読み方を置き換えないように）
const WEB = { console };
vm.createContext(WEB);
vm.runInContext(fs.readFileSync(path.join(ROOT, 'webapp_srv.js'), 'utf8'), WEB, { filename: 'webapp_srv.js' });
const WEBAPP_PAGES = vm.runInContext('WEBAPP_PAGES_', WEB);

// 担当者・チーム・ローテーション（架空）
F.saveRoleHolders(HOLDERS_23, 23, NEXT);
F.saveRoleHolders(HOLDERS_24, 24, '2026/10/07', TEAMS_24);
S.props.BNI_SPEAKER_ROTATION = JSON.stringify({
  order: FAKE_MEMBERS.map((m) => m.name), excluded: [N(0), N(1), N(2)], anchor: { date: '2026/10/07', pointer: 5 },
  header: 'メインプレゼンテーション（各４分45秒）',
  notes: ['２週間前までに、略歴書と発表用のデータ・資料を書記兼会計までご提出ください。'], updated: '',
});
// 事前MTGの共有事項（9/30）
const shared = (role, text) => F.saveRoleInput(NEXT, role, [{ id: F.roleExtraId_('今週の共有事項（' + F.roleDefOf_(role).label + '）'), value: text, orig: '' }]);
shared('president', '・24期への引継ぎ資料を9/28までに共有します\n・10/7 は24期として初めての定例会です');
shared('vice', '・ビジター3名（うち1名はオンライン）\n・代理1名');
shared('vhc', '・ビジターフォローは当日中に連絡をお願いします');
shared('ec', '・エデュケーション：「リファーラルの質を高める3つの質問」');
// そのほかの項目（お願い事項・直近のイベント・人数・代理・欠席）も、画面に出る項目の名前で入れる
function fill(role, pairs) {
  const ctx = F.getRoleInputContext(NEXT, role), r = ctx.roles.find((x) => x.key === role);
  const entries = pairs.map(([title, value]) => {
    const id = r.items.find((i) => ctx.items[i].title.replace(/\s/g, '') === title.replace(/\s/g, ''));
    return id ? { id, value, orig: ctx.items[id].value || '' } : null;
  }).filter(Boolean);
  F.saveRoleInput(NEXT, role, entries);
}
fill('president', [['お願い事項（プレジデントから）', '・10/7 の24期キックオフに、ビジターをお誘いください']]);
fill('vice', [['代理', surname(N(16))], ['欠席', surname(N(26)) + '、' + surname(N(30))], ['人数：ビジター', '3'], ['人数：ゲスト', '1'],
  ['人数：見学', '0'], ['人数：リージョン参加者', '1']]);
fill('secretary', [['直近のイベント', '・10/10（土）24期キックオフ交流会（18:30〜）\n・10/21（火）1to1促進ランチ会']]);

// ===================== 画面を動かすための「サーバー」 =====================
// google.script.run の代わり。返事は関数名と引数で引く（引数が合わなければ、その関数の既定の返事）
function stubScript(answers) {
  return '<script>var __R=' + JSON.stringify(answers).replace(/</g, '\\u003c') + ';'
    + 'function __pick(n,a){var e=__R[n];if(!e)return null;var k=JSON.stringify(a);'
    + 'return (e.byArgs&&Object.prototype.hasOwnProperty.call(e.byArgs,k))?e.byArgs[k]:(e.any===undefined?null:e.any);}'
    + 'var google={script:{host:{close:function(){}},get run(){var ok=null,p=new Proxy({},{get:function(_,n){'
    + 'if(n==="withSuccessHandler")return function(f){ok=f;return p;};'
    + 'if(n==="withFailureHandler"||n==="withUserObject")return function(){return p;};'
    + 'return function(){var v=__pick(n,[].slice.call(arguments));setTimeout(function(){ok&&ok(JSON.parse(JSON.stringify(v)));},15);};}});return p;}}};'
    + '</script>';
}
// Apps Script のテンプレート（<? ?>・<?= ?>・<?!= ?>）を、渡した値で展開する
function evalTemplate(file, vars) {
  const src = fs.readFileSync(path.join(ROOT, file), 'utf8');
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  let code = 'var __o=[];with(__v){', i = 0, m;
  const re = /<\?(!=|=)?([\s\S]*?)\?>/g;
  while ((m = re.exec(src))) {
    code += '__o.push(' + JSON.stringify(src.slice(i, m.index)) + ');';
    if (m[1] === '=') code += '__o.push(__e(' + m[2] + '));';
    else if (m[1] === '!=') code += '__o.push(String(' + m[2] + '));';
    else code += m[2] + '\n';
    i = re.lastIndex;
  }
  code += '__o.push(' + JSON.stringify(src.slice(i)) + ');}return __o.join("");';
  const include = (n) => fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8');
  return new Function('__v', '__e', code)(Object.assign({ include }, vars), esc);
}
function writePage(name, html, answers) {
  const out = html.replace(/<head>/i, '<head><meta charset="utf-8">' + stubScript(answers || {}));
  const f = path.join(tmp, name + '.html');
  fs.writeFileSync(f, out);
  return f;
}

// ===================== 撮る画面 =====================
const ctxOverview = F.getRoleInputContext('', '');
const ctxVice = F.getRoleInputContext(NEXT, 'vice');
const answersRole = {
  getSystemVersion: { any: '' },                       // 版は写さない（版を上げるたびに古くなるため）
  getRoleInputContext: { any: ctxOverview, byArgs: { [JSON.stringify(['', ''])]: ctxOverview, [JSON.stringify([NEXT, 'vice'])]: ctxVice } },
  getSpeakerRotation: { any: F.getSpeakerRotation() },
  getPreMeetingPreview: { any: F.getPreMeetingPreview(NEXT) },
};
const rolePage = (view, role) => evalTemplate('role_input.html', { params: { role: role || '', view: view || '' } });
const vhAnswers = {
  getMembersList: { any: F.getMembersList() }, getVisitorHosts: { any: [] }, getMemberPriorities: { any: {} },
  getVisitorHostsFromRoles: { any: F.getVisitorHostsFromRoles() },
};
// ---- 毎週の作業（参加者シートは、番号つきの架空のビジター・ゲスト・代理に差し替える）----
const VISITORS = [
  ['赤坂 太郎', 'あかさか たろう', '経営コンサルタント', '赤坂経営研究所', '代表', N(0), 'Visitor', '有効', '支払済み', 'オンライン参加'],
  ['白石 花子', 'しらいし はなこ', 'フラワーショップ', 'フラワーしらいし', '店長', N(6), 'Visitor', '有効', '未払い', ''],
  ['黒田 健', 'くろだ けん', '税理士', '黒田税理士事務所', '所長', N(13), 'Visitor', '有効', '支払済み', ''],
  ['緑川 由紀', 'みどりかわ ゆき', 'ヨガ講師', 'スタジオみどり', '代表', N(21), 'Visitor', '有効', '支払済み', ''],
  ['青井 誠一', 'あおい せいいち', '保険代理店', '青井保険事務所', '代表', N(16), 'Visitor', 'キャンセル', '未払い', ''],
  ['桃井 さくら', 'ももい さくら', '司会業', '桃井企画', '代表', N(25), 'Guest', '有効', '支払済み', ''],
  ['灰田 次郎', 'はいだ じろう', '家事代行', 'いけだ家事サポート', '', N(17), 'Substitute', '有効', '', N(17) + 'さんの代理'],
];
const romaji = ['akasaka', 'shiraishi', 'kuroda', 'midorikawa', 'aoi', 'momoi', 'haida'];
const PART_HEAD = ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', '招待者', 'メール', '種別', 'ステータス', '支払いステータス', 'メモ（ビジターリストに表示）'];
// 番号は本物の参加者シートと同じ形（ビジター V01…・ゲスト G01…・代理は「代理」＋代理を立てた方の番号）
let vNo = 0, gNo = 0;
const noOf = (v) => v[6] === 'Visitor' ? 'V' + String(++vNo).padStart(2, '0') : v[6] === 'Guest' ? 'G' + String(++gNo).padStart(2, '0')
  : '代理' + (FAKE_MEMBERS.find((m) => m.name === v[5]) || {}).no;
const partRows = VISITORS.map((v, i) => [noOf(v), v[0], v[1], v[2], v[3], v[4], v[5], romaji[i] + '@example.com', v[6], v[7], v[8], v[9]]);
S.sandbox.Utilities.parseCsv = (text) => text.replace(/\r/g, '').split('\n').filter((l) => l !== '').map((l) => {
  const out = []; let cur = '', q = false;
  for (let i = 0; i < l.length; i++) {
    const c = l[i];
    if (q) { if (c === '"' && l[i + 1] === '"') { cur += '"'; i++; } else if (c === '"') q = false; else cur += c; }
    else if (c === '"') q = true; else if (c === ',') { out.push(cur); cur = ''; } else cur += c;
  }
  out.push(cur);
  return out;
});
const CSV_HEAD = ['Name', 'Furigana', 'Business Category', 'Company Name', 'Job Title', 'Inviter', 'Email', 'Type', 'Status', 'Payment Status', 'Memo For Printing'];
const csvText = [CSV_HEAD].concat(partRows.map((r) => r.slice(1))).map((r) => r.map((v) => '"' + String(v).replace(/"/g, '""') + '"').join(',')).join('\n');
const answersCsv = { getMeetingCandidates: { any: F.getMeetingCandidates() }, getExistingVisitorSheets: { any: ['20260930参加者'] } };
const csvResult = F.analyzeCsvData(csvText);
{
  const i = S.sheets.findIndex((s) => s.getName() === '20260930参加者');
  S.sheets[i] = S.makeSheet('20260930参加者', [PART_HEAD].concat(partRows));
}
const answersWeekly = {
  getMeetingCandidates: { any: F.getMeetingCandidates() },
  getAllocationData: { any: F.getAllocationData(NEXT) },
  getEmailContext: { any: F.getEmailContext() },
  generateEmailDrafts: { any: F.generateEmailDrafts('20260930参加者') },
  getVisitorPostContext: { any: F.getVisitorPostContext() },
  getVisitorPostData: { any: F.getVisitorPostData('20260930参加者') },
};
// ---- 定例会スライド（前半・後半）。テンプレートは登録済みのつもり。アンバサダーなどは架空の方 ----
const slideCtx = F.getMeetingSlideContext();
slideCtx.templates = { meetingFirst: true, meetingSecond: true, memberPresen: true };
(slideCtx.members || []).forEach((m) => { m.hasPhoto = true; });
const mpCtx = F.getMemberPresenContext();
(mpCtx.members || []).forEach((m) => { m.hasPhoto = true; });
mpCtx.template = { kind: 'memberPresen', registered: true };
const answersSlides = {
  getSystemVersion: answersRole.getSystemVersion,
  getMeetingSlideContext: { any: slideCtx },
  getMemberPresenContext: { any: mpCtx },
  getWeeklyGuests: { any: { ok: true, guests: [{ name: '見本　アンバサダー', role: 'Activeチャプター担当アンバサダー', hidden: true },
                                                { name: '見本　ディレクター', role: 'エグゼクティブディレクター', hidden: true }] } },
  getRoutineInfo: { any: F.getRoutineInfo(NEXT) },
  computeRenewalLists: { any: F.computeRenewalLists(NEXT) },
  getSpeakerRotationWeeks: { any: F.getSpeakerRotationWeeks(NEXT) },
  getMeetingTemplateInfo: { any: { ok: true, message: '',
    list: [{ slide: 'ppt/slides/slide1.xml', slideNo: 1, spid: '3', name: '入場曲', volume: 25, video: false, key: 'slide1.xml#3' },
           { slide: 'ppt/slides/slide21.xml', slideNo: 21, spid: '3', name: '抽選BGM', volume: 60, video: false, key: 'slide21.xml#3' }],
    referralBoxes: { companyTall: { x: 5087424, y: 2849608, cx: 6893161, cy: 1446550 }, categoryLow: 4071101, categoryWidth: 6893161, hasNext: true, slides: 1 } } },
  getMeetingMusicFiles: { any: { ok: true, files: [{ id: 'f1', name: '入場曲.mp3', sizeMB: 3.2 }, { id: 'f2', name: '抽選BGM.mp3', sizeMB: 2.1 }] } },
};
// ---- 設定：メンバー名簿（役職は24期の内容）・休会日・大きなスライド ----
const DEF_CATS = vm.runInContext('DEFAULT_CATEGORIES_', S.sandbox);
const roleOfFake = {};
Object.keys(HOLDERS_24).forEach((k) => { roleOfFake[HOLDERS_24[k]] = F.roleDefOf_(k).label; });
const masterAnswer = { ok: true, members: FAKE_MEMBERS.map((m) => Object.assign({ kana: '', role: roleOfFake[m.name] || '', memo: '', photoFile: '',
    comment: '', refer: '', collab: '', joinDate: '', renewDate: '', expireDate: '' }, m)),
  categories: DEF_CATS.map((r) => ({ key: r[0], label: r[1], bg: r[2], bg2: r[3], block: r[4], order: r[5] })), cover: {} };
const bigStatus = F.getBigTemplateStatus();
(bigStatus.templates || []).forEach((t, i) => {
  if (t.builtin) return;
  t.registered = true; t.fileName = ['BNI_定例会前半.pptx', 'BNI_定例会後半.pptx', 'BNI_メンバープレゼン.pptx'][i] || (t.label + '.pptx');
  t.sizeMB = [34.2, 31.8, 2.4][i] || 1; t.url = 'https://drive.google.com/';
});
const answersSettings = { getMemberMaster: { any: masterAnswer }, getHolidays: { any: F.getHolidays() }, getBigTemplateStatus: { any: bigStatus },
                          getChapterSettings: { any: F.getChapterSettings() } };

// ---- トークスクリプト（既定のひな形。9/30 は上の架空の参加者・チェックシートの値で作る）----
for (const f of ['talk_script_default.js', 'talk_script_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), S.sandbox, { filename: f });
}
// 台本は前日〜当日に作るので、9/30 のチェックシートも埋まっているものとして、架空の値を入れる
// （ここより前に作った画面の返事は、入れる前の値のまま）
{
  const live = {};
  S.sheets.forEach((sh) => { if (/ルーティンチェックシート/.test(sh.getName())) live[sh.getName()] = sh._grid; });
  const sn = (i) => surname(N(i));
  [
    ['トークスクリプト追加事項', '9/30 は23期の最後の定例会です。閉会の前に、プレジデントから半年間のお礼を述べます'],
    ['その他注意事項', 'なし'], ['担当割り振り', 'ホスト　林・清水／スライド　斎藤／スポットライト　西村／ルーレット　高橋'],
    ['プレジデントより', '10/7 から24期です。役職の引継ぎをお願いします'], ['バイスプレジデントより', 'ビジターは3名の予定です'],
    ['書記兼会計より', '10月分のチャプター運営費のお振込みをお願いします'], ['その他のお知らせ', 'なし'],
    ['リージョン参加者', 'なし'], ['BNI目的と概要', 'Building Relationships（人間関係の構築）'], ['体験談', sn(19)],
    ['エデュケーション', sn(5)], ['ネットワーキングリーダー', 'なし'], ['新入会', 'なし'], ['更新式(更新メンバー)', sn(8)],
    ['バイスプレジデントによる報告', '9月の月間リファーラル数は○件、サンキュー額は○円でした（スライドを読み上げる）'],
    ['メンバーシップから報告', 'なし'], ['審査中カテゴリー', 'なし'], ['募集カテゴリー', '税理士、司法書士、Webデザイナー'],
    ['開放カテゴリー', 'なし'], ['真正度確認', sn(4) + '⇒' + sn(9) + '　見本 太郎様'],
    ['リマインダー・特別報告BNIからのお知らせ', 'サンキューの入力をお願いします'], ['アフターMTG', '10/7 のアフターMTGは24期の顔合わせです'],
    ['本日の招待者', [0, 6, 13, 21, 25].map(sn).join('、')], ['更新対象者60日前', 'なし'],
  ].forEach(([t, v]) => {
    // 項目は、トークスクリプトと同じ探し方（項目名が同じか、項目名で始まる行）で探す
    const hit = Object.values(live).some((g) => {
      const c = (g[0] || []).indexOf(NEXT), r = c < 0 ? -1 : F.routineFindRow_(g, [t.replace(/\s/g, '')]);
      if (r < 0) return false;
      g[r][c] = v;
      return true;
    });
    if (!hit) console.warn('（トークスクリプトの見本）チェックシートに無い項目: ' + t);
  });
  F.routineResetCache_();
}
const talkCtx = F.getTalkScriptContext();
const talkPv = F.previewTalkScript(NEXT, {});
const answersTalk = { getTalkScriptContext: { any: talkCtx }, previewTalkScript: { any: talkPv } };

const homeStatus = { ok: true, latest: { date: NEXT }, title: F.chapterSystemTitle_(), checks: [
  { label: 'チャプター', ready: true, detail: F.chapterLabel_() + '・いまの期 ' + F.roleTermOf_(new F.Date()) + '期・次回 ' + F.getMeetingCandidates()[0].display },
  { label: '素材フォルダ', ready: true, detail: 'BNI素材（見本）' }, { label: 'メンバーリスト（割り振り用）', ready: true, detail: FAKE_MEMBERS.length + '名' },
  { label: 'メンバー名簿（冊子・スライド用）', ready: true, detail: FAKE_MEMBERS.length + '名' },
  { label: 'メンバー写真', ready: false, detail: '30枚（未照合 6名）', fixLabel: '管理する', fixFn: 'openMemberPhotoDialog' },
  { label: 'PowerPointテンプレート', ready: true, detail: '3種類とも登録済み' }, { label: '大きなスライド', ready: true, detail: '登録済み' },
  { label: 'Gemini API', ready: true, detail: '設定済み' } ] };

const SHOTS = [
  { name: 'webapp_home', file: () => writePage('webapp_home', evalTemplate('webapp_home.html', {
      groups: WEBAPP_PAGES, status: { ok: true, name: 'BNI名簿システム（見本）', url: '', user: '' },
      appUrl: 'https://script.google.com/macros/s/xxxx/exec', version: '',
      title: 'Activeチャプター 名簿システム', footer: 'BNI東京千代田リージョン ｜ Activeチャプター' })), width: 900, height: 760 },
  { name: 'menu_home', file: () => writePage('menu_home', fs.readFileSync(path.join(ROOT, 'menu_home.html'), 'utf8'), { getHomeStatus: { any: homeStatus } }),
    width: 980, height: 760 },
  { name: 'role_overview', file: () => writePage('role_overview', rolePage(), answersRole), width: 960,
    shoot: async (p) => p.screenshot({ clip: await clipOf(p, 'h3', '#cards', 460) }) },
  { name: 'role_teams', file: () => writePage('role_teams', rolePage(), answersRole), width: 960, element: '#holderBox',
    before: async (p) => { await p.selectOption('#holderTerm', '24'); await p.evaluate(() => changeHolderTerm()); await p.waitForTimeout(150); } },
  { name: 'role_input_vice', file: () => writePage('role_vice', rolePage('', 'vice'), answersRole), width: 960,
    shoot: async (p) => p.screenshot({ clip: await clipOf(p, 'h3', '#items', 620) }) },
  { name: 'premtg_preview', file: () => writePage('premtg', rolePage('premtg'), answersRole), width: 960, element: '#premtgBox', wait: 900 },
  { name: 'rotation', file: () => writePage('rotation', rolePage('rotation', 'secretary'), answersRole), width: 960,
    shoot: async (p) => p.screenshot({ clip: await clipOf(p, '#rotView', null, 820) }) },
  { name: 'visitor_host', file: () => writePage('visitor_host', fs.readFileSync(path.join(ROOT, 'visitor_host.html'), 'utf8'), vhAnswers), width: 460, height: 700,
    before: async (p) => { await p.click('button:has-text("読み取る")'); await p.waitForTimeout(150); } },
  // 毎週の作業
  { name: 'csv_start', file: () => writePage('csv', readHtml('dialog.html'), answersCsv), width: 820, height: 560 },
  { name: 'csv_editor', file: () => writePage('csv2', readHtml('dialog.html'), answersCsv), width: 1000, height: 620,
    before: async (p) => { await p.evaluate((r) => { document.getElementById('step1').classList.add('hidden'); showEditor(r); }, csvResult);
      await p.waitForTimeout(200); } },
  { name: 'email', file: () => writePage('email', readHtml('email.html'), answersWeekly), width: 900, height: 760, wait: 900 },
  { name: 'allocation', file: () => writePage('allocation', readHtml('allocation.html'), answersWeekly), width: 1200, height: 820,
    before: async (p) => {
      await p.click('button:has-text("データを読み込む")'); await p.waitForTimeout(500);
      // 何人か割り振った状態にする（ファシリテーター・ルームメンバー・オリエンテーション）
      await p.evaluate(() => {
        state.facilAlloc = { V01: '10', V02: '29', V03: '4', V04: '34' };
        state.roomAlloc = { V01: ['23', '24'], V02: ['31'], V03: ['35', '36'], V04: ['22'] };
        state.orienAlloc = { V01: ['17'], V03: ['12'] };
        render();
      });
      await p.waitForTimeout(200);
    } },
  { name: 'visitor_post', file: () => writePage('visitor_post', readHtml('visitor_post.html'), answersWeekly), width: 960, height: 820, wait: 900 },
  // 定例会スライド
  { name: 'slides_first', file: () => writePage('first', evalTemplate('slides_meeting_first.html', {}), answersSlides), width: 1000, height: 900, wait: 1200 },
  { name: 'slides_second', file: () => writePage('second', evalTemplate('slides_meeting_second.html', {}), answersSlides), width: 1000, height: 900, wait: 1200 },
  // トークスクリプト（中身を確かめたところ・ひな形の編集）
  { name: 'talk_script', file: () => writePage('talk_script', readHtml('talk_script.html'), answersTalk), width: 1100, height: 900,
    before: async (p) => { await p.click('#btnPreview'); await p.waitForTimeout(300); } },
  { name: 'talk_template', file: () => writePage('talk_template', readHtml('talk_script.html'), answersTalk), width: 1100, height: 760,
    before: async (p) => {
      await p.click('#tabEdit'); await p.waitForTimeout(150);
      await p.evaluate(() => { const r = document.getElementById('r26_talk'); if (r) window.scrollTo(0, r.getBoundingClientRect().top + window.scrollY - 140); });
      await p.waitForTimeout(100);
    } },
  // 設定
  { name: 'member_master', file: () => writePage('member_master', readHtml('member_master.html'), answersSettings), width: 1280, height: 640, wait: 700 },
  { name: 'holiday', file: () => writePage('holiday', readHtml('holiday.html'), answersSettings), width: 520, height: 560 },
  { name: 'chapter_settings', file: () => writePage('chapter_settings', readHtml('chapter_settings.html'), answersSettings), width: 580, height: 740 },
  { name: 'big_templates', file: () => writePage('big_templates', readHtml('big_templates.html'), answersSettings), width: 900, height: 700, wait: 700 },
];
function readHtml(f) { return fs.readFileSync(path.join(ROOT, f), 'utf8'); }

// 画面の中で、start の上端から end の下端まで（end が無ければ高さ h まで）を切り出す
async function clipOf(p, start, end, maxH) {
  const a = await p.$eval(start, (e) => { const r = e.getBoundingClientRect(); return { x: 0, y: r.top + window.scrollY }; });
  let bottom = a.y + (maxH || 600);
  if (end) {
    const b = await p.$eval(end, (e) => { const r = e.getBoundingClientRect(); return r.bottom + window.scrollY; });
    bottom = Math.min(b + 8, a.y + (maxH || 99999));
  }
  const w = await p.evaluate(() => document.documentElement.clientWidth);
  return { x: 0, y: Math.max(0, a.y - 8), width: w, height: Math.max(40, bottom - a.y + 8) };
}

function loadPlaywright() {
  try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); }
}

// PNG を WebP にして出力先へ（幅 960px まで・画質82）
const TO_WEBP = `
import os, sys
from PIL import Image
src, dst = sys.argv[1], sys.argv[2]
for f in sorted(os.listdir(src)):
    if not f.endswith('.png'):
        continue
    im = Image.open(os.path.join(src, f)).convert('RGB')
    w, h = im.size
    if w > 960:
        im = im.resize((960, round(h * 960 / w)), Image.LANCZOS)
    im.save(os.path.join(dst, f[:-4] + '.webp'), 'WEBP', quality=82, method=6)
`;

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const PNG = path.join(tmp, 'png');
  fs.mkdirSync(PNG, { recursive: true });
  const { chromium } = loadPlaywright();
  const browser = await chromium.launch();
  const done = [];
  for (const s of SHOTS) {
    if (ONLY.length && ONLY.indexOf(s.name) < 0) continue;
    const page = await browser.newPage({ viewport: { width: s.width || 960, height: s.height || 900 }, deviceScaleFactor: 1 });
    await page.goto('file://' + s.file());
    await page.waitForTimeout(s.wait || 500);
    if (s.before) await s.before(page);
    const out = path.join(PNG, s.name + '.png');
    if (s.shoot) fs.writeFileSync(out, await s.shoot(page));
    else if (s.element) await (await page.$(s.element)).screenshot({ path: out });
    else await page.screenshot({ path: out, fullPage: !!s.full });
    await page.close();
    done.push(s.name);
  }
  await browser.close();
  require('child_process').execFileSync('python3', ['-c', TO_WEBP, PNG, OUT], { stdio: 'inherit' });
  console.log('撮った画面 ' + done.length + '枚 → ' + path.relative(ROOT, OUT) + '/（WebP）: ' + done.join('、'));
})().catch((e) => { console.error(e); process.exit(1); });
