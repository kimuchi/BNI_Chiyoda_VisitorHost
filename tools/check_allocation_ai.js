// 割り振り表の「AIに提案させる」（callGeminiAutoAllocation）を確かめる。名簿・ビジターは作り物。Gemini には送らない。
//   1. 待機メンバーの情報は、メンバー名簿（メンバーブックと同じ中身）から文字のまま渡す（メンバーブックのPDFは渡さない）
//      … 業種区分・カテゴリー・会社名・会社での役職・BNIの役職・一言・紹介・協業。ふりがな・メモ・日付は渡さない。
//        待機メンバーでない方（招待者など）の情報は渡さない。AI参考資料はこれまでどおり添付する
//   2. 名簿に一言・紹介・協業がまったく入っていないときだけ、これまでどおりメンバーブックのPDFも渡す
//   3. 返ってきた割り振りの反映（同じ方を2つのルームに入れない・ファシリテーターはルームに入れない）
//
//   node tools/check_allocation_ai.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

// 作り物の名簿（No・氏名・業種区分・カテゴリー・会社名・…）
const ROSTER = [
  { no: '01', name: '見本 一郎', cat: '企業サポート', title: '税理士', company: '見本会計', position: '所長', role: 'ビジターホスト',
    comment: '会社の数字を経営の味方に', refer: '創業3年以内の社長', collab: '司法書士・社会保険労務士', kana: 'みほん いちろう', memo: '内部のメモ', joinDate: '2020/01/01' },
  { no: '02', name: '見本 二郎', cat: '不動産関連', title: '不動産売買仲介', company: '見本不動産', position: '代表', role: '',
    comment: '', refer: '相続を考えている方', collab: '税理士', kana: '', memo: '', joinDate: '' },
  { no: '03', name: '見本 三郎', cat: '建築・住まい', title: 'リフォーム', company: '見本工務店', position: '', role: 'プレジデント',
    comment: '住まいの困りごとを解決します', refer: '賃貸オーナー', collab: '不動産管理会社', kana: '', memo: '', joinDate: '' },
  { no: '04', name: '見本 四郎', cat: 'プロモーション', title: 'Web制作', company: '見本ウェブ', position: '', role: '',
    comment: '', refer: '', collab: '', kana: '', memo: '', joinDate: '' },
  { no: '05', name: '見本 五郎', cat: '暮らし・生活', title: '保険代理店', company: '見本保険', position: '', role: '',
    comment: '保険の見直しならお任せを', refer: '子育て世帯', collab: 'ファイナンシャルプランナー', kana: '', memo: '', joinDate: '' },
  { no: '06', name: '見本 六郎', cat: '美容と健康', title: '整体院', company: '見本整体', position: '院長', role: '',
    comment: '招待者の方の一言（送らない）', refer: '招待者の方の紹介（送らない）', collab: '', kana: '', memo: '', joinDate: '' },
];
const CATS = ['企業サポート', '不動産関連', '建築・住まい', 'プロモーション', '暮らし・生活', '美容と健康'].map((k) => ({ key: k, label: k === '美容と健康' ? '美容・健康' : k }));

function makeServer(roster, geminiText) {
  const props = { GEMINI_API_KEY: 'test-key', MEMBER_BOOK_ID: 'mb-pdf', AI_REF_DOCS: J([{ id: 'doc-1', name: '参考資料.txt' }]) };
  const files = { 'mb-pdf': { type: 'application/pdf', data: Buffer.from('%PDF-1.4 メンバーブック') },
                  'doc-1': { type: 'text/plain', data: Buffer.from('コンタクトサークルのアンケート') } };
  const fetches = [];
  const box = {
    console,
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); } }) },
    DriveApp: { getFileById: (id) => {
      if (!files[id]) throw new Error('no file ' + id);
      return { getBlob: () => ({ getBytes: () => Array.from(files[id].data).map((b) => (b > 127 ? b - 256 : b)), getContentType: () => files[id].type }) };
    } },
    Utilities: { base64Encode: (a) => Buffer.from(a.map((b) => b & 255)).toString('base64') },
    UrlFetchApp: { fetch: (url, opt) => {
      fetches.push({ url, payload: JSON.parse(opt.payload) });
      return { getContentText: () => J({ candidates: [{ content: { parts: [{ text: geminiText }] } }] }) };
    } },
  };
  vm.createContext(box);
  for (const f of ['コード.js', 'assets.js', 'member_master_srv.js']) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), box, { filename: f });
  }
  box.getMemberMaster = () => ({ ok: true, members: roster.map((m) => Object.assign({}, m)) });
  box.getCategoryMaster = () => CATS;
  return { box, fetches };
}

// 画面から渡る状態（getAllocationData と同じ形）。招待者の 06 は待機メンバーに入らない
const state = () => ({
  visitors: [
    { no: 'V01', name: '見本 花子', cat: '司法書士', inviter: '見本 六郎',
      details: { 'メンバーになる確度': '3', '種別': 'Visitor', 'メモ（非表示）': '', 'メモ（ビジターリストに表示）': '相続に強い方とつながりたい' } },
    { no: 'V02', name: '見本 桜', cat: 'ファイナンシャルプランナー', inviter: '',
      details: { 'メンバーになる確度': '1', '種別': 'Visitor', 'メモ（非表示）': '', 'メモ（ビジターリストに表示）': '' } },
    { no: 'G01', name: '見本 梅', cat: '司会', inviter: '',
      details: { 'メンバーになる確度': '', '種別': 'Guest', 'メモ（非表示）': '', 'メモ（ビジターリストに表示）': '' } },
  ],
  pool: ROSTER.filter((m) => m.no !== '06').map((m) => ({ no: m.no, name: m.name })),
  hosts: ['01', '02'], facilAlloc: {}, orienAlloc: {}, roomAlloc: {}, connectReq: { V01: '見本 二郎' }, mergedWith: {}, priorities: {},
});
// Gemini の返事：V01 に 03・04、V02 に 04（重複）・05、G01 に 02（ファシリテーター）
const ANSWER = J({ allocations: [{ visitorNo: 'V01', room: ['03', '04'] }, { visitorNo: 'V02', room: ['04', '05'] }, { visitorNo: 'G01', room: ['02'] }] });

// ---------- 1. 名簿の情報を文字で渡す（PDFは渡さない）----------
{
  const { box, fetches } = makeServer(ROSTER, ANSWER);
  const res = box.callGeminiAutoAllocation(state(), 5);
  ck(fetches.length === 1, 'Gemini を呼んだ回数: ' + fetches.length);
  const parts = fetches[0].payload.contents[0].parts, prompt = parts[0].text;
  const inline = parts.slice(1).map((p) => p.inlineData && p.inlineData.mimeType);
  ck(J(inline) === J(['text/plain']), '添付：AI参考資料だけのはず（メンバーブックのPDFは渡さない）: ' + J(inline));
  const m = prompt.match(/待機メンバー（プロフィール）: (\[.*\])\n/);
  ck(!!m, '待機メンバーのプロフィールが無い');
  const prof = m ? JSON.parse(m[1]) : [];
  ck(J(prof.map((x) => x.no)) === J(['01', '02', '03', '04', '05']), '待機メンバーの並び: ' + J(prof.map((x) => x.no)));
  const one = prof.find((x) => x.no === '01') || {};
  ck(one['業種区分'] === '企業サポート' && one['カテゴリー'] === '税理士' && one['会社名'] === '見本会計' && one['会社での役職'] === '所長'
     && one['BNIの役職'] === 'ビジターホスト' && one['一言'] === '会社の数字を経営の味方に' && one['紹介してほしい人'] === '創業3年以内の社長'
     && one['協業したい人'] === '司法書士・社会保険労務士', '名簿の中身: ' + J(one));
  ck(!('一言' in (prof.find((x) => x.no === '02') || {})), '空の項目は入れない');
  ck(J(prof.find((x) => x.no === '04')) === J({ no: '04', name: '見本 四郎', '業種区分': 'プロモーション', 'カテゴリー': 'Web制作', '会社名': '見本ウェブ' }),
     '一言・紹介・協業の無い方: ' + J(prof.find((x) => x.no === '04')));
  ['みほん いちろう', '内部のメモ', '2020/01/01', '招待者の方の一言', '招待者の方の紹介'].forEach((w) => ck(!prompt.includes(w), `送らない情報が入っている: ${w}`));
  ck(prompt.includes('待機メンバーのプロフィール（メンバー名簿の内容）') && prompt.includes('添付された参考資料') && !prompt.includes('メンバーブックPDF'),
     '指示文の、使う情報の書き方');
  ck(prompt.includes('紹介してほしい人') && prompt.includes('協業したい人') && /つながりの考え方/.test(prompt), 'つながりの考え方の指示が無い');
  // 3. 返事の反映：ファシリテーター（01・02）はルームに入らない。同じ方は1つのルームだけ
  ck(res.facilAlloc.V01 === '01' && res.facilAlloc.V02 === '02', 'ファシリテーター: ' + J(res.facilAlloc));
  ck(J(res.roomAlloc.V01) === J(['03', '04']) && J(res.roomAlloc.V02) === J(['05']) && J(res.roomAlloc.G01) === J([]), 'ルームメンバー: ' + J(res.roomAlloc));
}

// ---------- 2. 名簿に一言・紹介・協業がまったく無いときは、これまでどおりPDFも渡す ----------
{
  const bare = ROSTER.map((m) => Object.assign({}, m, { comment: '', refer: '', collab: '' }));
  const { box, fetches } = makeServer(bare, ANSWER);
  box.callGeminiAutoAllocation(state(), 5);
  const parts = fetches[0].payload.contents[0].parts, prompt = parts[0].text;
  const inline = parts.slice(1).map((p) => p.inlineData && p.inlineData.mimeType);
  ck(J(inline) === J(['application/pdf', 'text/plain']), '名簿にメンバーブックの内容が無いときは、PDFと参考資料: ' + J(inline));
  ck(prompt.includes('添付されたメンバーブックPDF'), 'PDFを渡すときの指示文');
}

// ---------- 名簿に居ない方（古い「メンバーリスト」シートだけの方）は番号と氏名だけ ----------
{
  const { box } = makeServer(ROSTER, ANSWER);
  const prof = box.allocationMemberProfiles_([{ no: '99', name: '見本 名簿外' }, { no: '03', name: '見本　三郎' }]);
  ck(J(prof[0]) === J({ no: '99', name: '見本 名簿外' }) && prof[1]['カテゴリー'] === 'リフォーム', '名簿に居ない方・全角の空白: ' + J(prof));
}

console.log(`割り振り表のAI: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 待機メンバーの情報はメンバー名簿から文字で渡す（PDFは名簿に一言・紹介・協業が無いときだけ）・送らない情報・返事の反映');
