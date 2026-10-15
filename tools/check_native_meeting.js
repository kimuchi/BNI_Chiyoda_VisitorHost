// 手元の日本語定例会テンプレートを、専用メンバーテンプレートなしで生成する。
// node tools/check_native_meeting.js <template.pptx> <output-directory>
// 原本・実際の名簿やシートは変更しない。検査の氏名と数値は架空。
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');
const { makeEnv } = require('./lib_sheet_fake');
const { readZip } = require('./lib_zip');
const input = process.argv[2], outDir = process.argv[3];
if (!input || !outDir) throw new Error('template.pptx と出力先を指定してください');
fs.mkdirSync(outDir, { recursive: true });
const env = makeEnv({ now: new Date(2026, 9, 20, 9, 0) });
env.reset([['20261020参加者', false, [
  ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール', 'ステータス'],
  ['V01', '来訪 一郎', '', 'ITコンサルタント', '来訪株式会社', '見本 一郎', '', 'Visitor', '', ''],
  ['G01', 'ゲスト 花子', '', '保険', 'ゲスト商事', '試験 花子', '', 'Guest', '', ''],
  ['V02', '取消 二郎', '', '印刷', '取消会社', '見本 一郎', '', 'Visitor', '', 'キャンセル']
]]], {});
const F = Object.assign({}, env.globals);
vm.createContext(F);
const root = path.join(__dirname, '..');
for (const file of fs.readdirSync(root).filter(f => f.endsWith('.js')).sort()) {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), F, { filename: file });
}
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const blob = F.Utilities.newBlob([...fs.readFileSync(input)], 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'native.pptx');
const fresh = () => F.unzipToMap_(blob.copyBlob());
const parts = fresh(), order = F.slideOrder_(parts);
const media = Object.keys(parts).filter(p => p.startsWith('ppt/media/'));
ck(F.prepareNativeMeetingTemplate_(parts), '日本語形式を見分けられない');
ck(JSON.stringify(F.slideOrder_(parts)) === JSON.stringify(order), '準備でページ順が変わった');
const prepared = F.zipFromMap_(parts, 'prepared.pptx');
fs.writeFileSync(path.join(outDir, 'prepared.pptx'), prepared._buf);
ck(!F.prepareNativeMeetingTemplate_(parts), '準備処理を繰り返して枠を増やした');
ck(F.mpFindModelSlide_(parts, [F.MP_INDIVIDUAL_.photo, F.MP_INDIVIDUAL_.nameBox, F.MP_INDIVIDUAL_.company,
  F.MP_INDIVIDUAL_.category, F.MP_INDIVIDUAL_.nextName, F.MP_INDIVIDUAL_.nextLabel]), 'メンバー個人ページが見つからない');
ck(F.rfLayoutBoxes_(parts, F.RF_TITLE_), 'リファーラルの組版枠が見つからない');
const members = [
  { no: '1', name: '見本 一郎', company: '見本会計株式会社', title: '税理士', blockKey: 'sample', block: '企業サポート' },
  { no: '2', name: '試験 花子', company: '試験商事株式会社', title: '印刷', blockKey: 'sample', block: '企業サポート' },
  { no: '3', name: '架空 三郎', company: '架空保険株式会社', title: '保険', blockKey: 'sample', block: '企業サポート' }
];
F.getMemberMaster = () => ({ ok: true, members });
F.getMembersList = () => members;
const zlib = require('zlib');
const chunk = (type, data) => {
  const td = Buffer.concat([Buffer.from(type), data]), len = Buffer.alloc(4), crc = Buffer.alloc(4);
  len.writeUInt32BE(data.length); crc.writeUInt32BE(zlib.crc32(td) >>> 0);
  return Buffer.concat([len, td, crc]);
};
const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(40, 0); ihdr.writeUInt32BE(60, 4); ihdr[8] = 8; ihdr[9] = 2;
const row = Buffer.alloc(1 + 40 * 3, 200); row[0] = 0;
const png = Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr),
  chunk('IDAT', zlib.deflateSync(Buffer.concat(Array.from({ length: 60 }, () => row)))), chunk('IEND', Buffer.alloc(0))]);
env.addFile('PHOTO', F.Utilities.newBlob([...png], 'image/png', 'sample.png'), 'sample.png');
const readFile = F.DriveApp.getFileById;
F.DriveApp.getFileById = id => {
  const file = readFile(id), readBlob = file.getBlob;
  file.getBlob = () => readBlob().copyBlob();
  return file;
};
F.findPhotoIdForName_ = () => 'PHOTO';
F.getBigTemplateFile_ = kind => {
  if (kind === 'memberPresen') throw new Error('「メンバープレゼン」のテンプレートが未登録です。');
  return { getBlob: () => blob.copyBlob(), getName: () => 'native.pptx', getId: () => 'NATIVE',
    getLastUpdated: () => new Date(2026, 9, 15) };
};
const layout = (() => {
  const { loadPage } = require('./lib_minidom');
  // 組版は本番と同じ company_break.js と slides_layout.html。
  const page = loadPage('slides_layout.html');
  vm.runInContext(fs.readFileSync(path.join(root, 'company_break.js'), 'utf8'), page.sandbox);
  return page.sandbox;
})();
const mpCtx = { members, blocks: [{ gkey: 'sample', block: '企業サポート', count: 3 }],
  rowsPerPage: 7, seconds: { weekly: 30, startup: 150 },
  layoutBoxes: F.templateMemo_('meetingFirst', 'native_mp_test', F.mpLayoutBoxes_) };
const items = layout.memberPresenItems(mpCtx, 'sample', '試験 花子', true, { '架空 三郎': true });
const roleData = F.riDataFrom_(24, '2026年10月〜2027年3月',
  { president: '見本 一郎', vice: '試験 花子', secretary: '架空 三郎', ec: '見本 一郎', mentor: '試験 花子', vhc: '架空 三郎' }, 24,
  [{ key: 'membership', name: 'メンバーシップ委員会', members: [{ name: '試験 花子' }] },
    { key: 'role:vhc', name: 'ビジターホスト', members: [{ name: '架空 三郎' }] }], true, members);
F.getRoutineInfo = () => ({ ok: true, found: true, absentees: { names: ['架空 三郎'] },
  mainPresenters: [{ name: '見本 一郎', matched: true, kind: 'member' }],
  recommendations: [{ giver: { name: '見本 一郎' }, receiver: { name: '試験 花子' }, when: 'during' }] });
F.computeRenewalLists = () => ({ ok: true, text: { d90: '見本 一郎', d60: '試験 花子', d30: '該当者なし', overdue: '架空 三郎' } });
F.chapterLabel_ = () => 'サンプルチャプター';
let output;
F.saveOutputFile_ = b => { output = b; return { url: 'https://drive.example/TEST', downloadUrl: 'https://drive.example/TEST' }; };
const values = { 開催回: '42', 審査中カテゴリー: 'ITコンサルタント',
  メインプレゼン1氏名: '見本 一郎', メインプレゼン1会社名: '見本会計株式会社', メインプレゼン1カテゴリー: '【税理士】',
  メインプレゼン2氏名: '', メインプレゼン2会社名: '', メインプレゼン2カテゴリー: '' };
for (let i = 1; i <= 12; i++) values['求める専門分野' + i] = i === 1 ? '弁護士' : '';
const weeks = [0, 1, 2, 3, 4].map(i => ({ date: '2026/10/' + (20 + 7 * i), no: String(42 + i), md: '10/' + (20 + 7 * i),
  people: [{ name: '見本 一郎', title: '税理士', company: '見本会計株式会社', collab: '' }], label: (i + 1) + '回目' }));
const result = F.generateMeetingSlides('meetingFirst', values, '2026/10/20', {
  memberPresen: items, mainPresenters: ['見本 一郎'], roleIntroData: roleData, weeklyGuests: [], absentNames: ['架空 三郎'],
  coreValue: 'Givers Gain', speakerRotation: { weeks, header: 'メインプレゼンテーション', notes: [] },
  memberPages: { newMembers: [{ name: '試験 花子' }], renewMembers: [{ name: '見本 一郎', years: 2 }] },
  vpReport: { avg: '308', visitorAvg: '12.5', total: '1,927' },
  networkingLeaders: { show: true, items: [
    { key: 'ext', value: '17', unit: '件', winners: [{ name: '見本 一郎' }] },
    { key: 'visitor', value: '3', unit: '名', winners: [{ name: '試験 花子' }] }] }
});
ck(result.ok, '全体生成が失敗: ' + result.message);
if (output) {
  const outFile = path.join(outDir, 'meeting-native-sample.pptx');
  fs.writeFileSync(outFile, output._buf);
  const files = readZip(output._buf), outParts = F.unzipToMap_(output);
  const text = p => F.slideText_(F.xmlOf_(outParts, p) || '');
  const texts = F.slideOrder_(outParts).map(text);
  ck(texts[0].includes('サンプルチャプター') && texts[0].includes('第42回') && texts[0].includes('2026年10月20日'), '表紙のチャプター・開催回・日付が更新されない');
  ck(texts.some(t => t.includes('学習コーナー') && t.includes('見本 一郎') && t.includes('学習トピック')),
    '改行を含む学習コーナーの担当者が入らない');
  ck(texts.some(t => t.includes('弁護士') && t.includes('ITコンサルタント')), '募集・審査中カテゴリーが入らない');
  ck(texts.some(t => t.includes('来訪 一郎') && t.includes('来訪株式会社'))
    && texts.some(t => t.includes('ゲスト 花子') && t.includes('ゲスト商事'))
    && !texts.some(t => t.includes('取消 二郎')), '同梱ビジター紹介・キャンセル除外が違う');
  ck(texts.some(t => t.includes('月間ビジター数の平均') && t.includes('12.5名') && t.includes('1,927件')), 'バイスプレジデントの数値が入らない');
  F.applyVpReport_(outParts, { visitorAvg: '14.2' }, new Date(2026, 9, 20));
  const updatedVp = F.slideOrder_(outParts).map(text).find(t => t.includes('月間ビジター数の平均'));
  ck(updatedVp.includes('14.2名') && !updatedVp.includes('12.14.2'), '小数のビジター平均を再更新できない');
  ck(texts.some(t => t.includes('リファーラル数 第1位') && t.includes('見本 一郎') && t.includes('試験 花子')), 'ネットワーキングリーダーが入らない');
  const weekly = F.rfFindModels_(outParts, F.WEEKLY_TITLE_);
  ck(weekly.length === 2 && weekly.some(m => F.slideText_(m.xml).includes('見本 一郎'))
    && weekly.some(m => F.slideText_(m.xml).includes('試験 花子')), '同梱のウィークリー個人ページが人数ぶん生成されない');
  ck(weekly.every(m => !F.slideText_(m.xml).includes('架空 三郎')), '欠席の方の個人ページがある');
  ck(weekly.some(m => F.mpCountdownSeconds_(m.xml) === 150) && weekly.some(m => F.mpCountdownSeconds_(m.xml) === 30),
    'スタートアップと通常のカウントダウン秒数が違う');
  ck(F.rfFindModels_(outParts, F.RF_TITLE_).length === 2, '同梱の後半のリファーラルが生成されない');
  ck(texts.some(t => t.includes('推薦のことば') && t.includes('見本 一郎') && t.includes('試験 花子')), '推薦のことばに氏名が入らない');
  ck(texts.some(t => t.includes('90日以内') && t.includes('見本 一郎') && t.includes('試験 花子')), '更新状況が入らない');
  ck(texts.some(t => t.includes('チャプタープレジデント') && t.includes('見本 一郎') && t.includes('試験 花子')), '役職の氏名が入らない');
  ck(!texts.some(t => /\{\{/.test(t)), '差し込み口が出力に残った');
  const original = readZip(fs.readFileSync(input));
  ck(media.every(p => files[p] && files[p].equals(original[p])), '元の画像・動画のデータを壊した');
  try { execFileSync('python3', [path.join(__dirname, 'pptx_integrity.py'), path.join(outDir, 'prepared.pptx'), outFile]); ck(true, ''); }
  catch (e) { ck(false, 'PowerPointの内部参照が壊れた: ' + String(e.stdout || e.message)); }
}
if (fails.length) { console.log(fails.join('\n')); console.log('NG: ' + fails.length + '/' + checks); process.exit(1); }
console.log('OK: 日本語定例会テンプレート ' + checks + '項目、出力先 ' + outDir);
