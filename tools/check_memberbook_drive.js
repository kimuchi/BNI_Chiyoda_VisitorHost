// ドライブのメンバーブック（メールで送るPDF）の差し替え（memberbook_srv.js の saveMemberBookPdfToDrive・コード.js の
// uploadMemberBookBlob_・memberBookFileInfo_）を、見せかけのドライブで確かめる。名前はすべて架空。
//
//   node tools/check_memberbook_drive.js
//
// 確かめること
//   ・登録してあるファイル（MEMBER_BOOK_ID）の中身を差し替え、差し替えたファイルの名前・持ち主・ドライブの更新日時を知らせる
//   ・差し替えるのは、リンク集・案内メールのリンク（MEMBER_BOOK_URL）のファイル。覚えているファイルID（MEMBER_BOOK_ID）が
//     それと違っていたら、リンクの方にそろえる（前は ID の方を差し替え、リンクで見るファイルが替わらなかった）
//   ・差し替えたのにドライブの更新日時が変わらなければ ⚠ で知らせる
//   ・画面に出す、登録してあるファイルの様子（getMemberBookFileInfo）

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

const env = makeEnv({ now: new Date(2026, 10, 10, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
env.reset([], { MEMBER_BOOK_ID: 'OLDFILE1234567890', MEMBER_BOOK_URL: 'https://drive.google.com/file/d/LINKFILE1234567890/view?usp=drivesdk' });

let clock = Date.UTC(2026, 9, 14, 1, 2, 0), stuck = false, updates = 0;
const files = { LINKFILE1234567890: { name: 'MemberBook.pdf', owner: '見本 一郎', modified: clock, data: 'old' },
                OLDFILE1234567890: { name: 'MemberBook.pdf', owner: '試験 二郎', modified: clock, data: 'other' } };
F.Drive = { Files: {
  get: (id) => {
    const f = files[id];
    if (!f) throw new Error('File not found: ' + id);
    return { id, name: f.name, webViewLink: 'https://drive.test/file/d/' + id + '/view', modifiedTime: new Date(f.modified).toISOString(),
             trashed: false, size: '2097152', owners: [{ displayName: f.owner }] };
  },
  update: (meta, id, blob) => { updates++; if (blob) { files[id].data = blob.getDataAsString(); if (!stuck) files[id].modified = (clock += 3600000); } return { id }; },
} };
F.DriveApp.getFileById = (id) => ({ getUrl: () => 'https://drive.test/file/d/' + id + '/view' });
const b64 = Buffer.from('%PDF-1.4 見本').toString('base64');

// ===== 差し替え =====
let r = F.saveMemberBookPdfToDrive(b64, 'MemberBook.pdf', false);
ck(r.ok && files.LINKFILE1234567890.data === '%PDF-1.4 見本' && files.OLDFILE1234567890.data === 'other' && updates === 1,
   'リンク集のリンクのファイルを差し替えていない: ' + J([r, files, updates]));
ck(r.url === 'https://drive.google.com/file/d/LINKFILE1234567890/view?usp=drivesdk' && env.props.MEMBER_BOOK_ID === 'LINKFILE1234567890'
   && env.props.MEMBER_BOOK_URL === 'https://drive.google.com/file/d/LINKFILE1234567890/view?usp=drivesdk',
   'リンクが変わった・覚えているファイルIDをリンクにそろえない: ' + J([r.url, env.props]));
ck(/差し替えたファイル：「MemberBook\.pdf」（持ち主 見本 一郎）　ドライブの更新日時 2026\/10\/14 11:02/.test(r.message) && !/⚠/.test(r.message),
   '差し替えたファイルの知らせ: ' + J(r.message));
ck(r.file && r.file.name === 'MemberBook.pdf' && r.file.modified === '2026/10/14 11:02', '画面に渡すファイルの様子: ' + J(r.file));

// ===== 更新日時が変わらない =====
stuck = true;
r = F.saveMemberBookPdfToDrive(b64, 'MemberBook.pdf', false);
ck(r.ok && /⚠ 差し替えたはずですが、ドライブのファイルの更新日時が変わっていません/.test(r.message), '更新日時が変わらないときの知らせ: ' + J(r.message));
stuck = false;

// ===== アップロードから（メンバーブック(PDF)の更新）=====
r = F.uploadMemberBookBlob_(F.Utilities.newBlob('%PDF-1.4 アップロード', 'application/pdf', 'x.pdf'));
ck(!r.error && /更新しました/.test(r.msg) && /差し替えたファイル：「MemberBook\.pdf」/.test(r.msg) && r.url === 'https://drive.google.com/file/d/LINKFILE1234567890/view?usp=drivesdk' && files.LINKFILE1234567890.data === '%PDF-1.4 アップロード', 'アップロードからの差し替え: ' + J(r));

// ===== 画面に出す様子 =====
const info = F.getMemberBookFileInfo();
ck(info.ok && info.file.name === 'MemberBook.pdf' && info.file.owner === '見本 一郎' && /^2026\/10\/14 /.test(info.file.modified) && info.file.size === 2,
   '登録してあるファイルの様子: ' + J(info));
ck(F.memberBookDriveInfo_().file.name === 'MemberBook.pdf', '編集画面に渡す様子にファイルが無い');
env.props.MEMBER_BOOK_URL = '';
env.props.MEMBER_BOOK_ID = 'GONEFILE1234567890';
ck(/not found/.test(F.getMemberBookFileInfo().file.error), '見つからないファイルの様子: ' + J(F.getMemberBookFileInfo()));

// ===== 03_生成物 への保存（saveOutputFile_）=====
// ゴミ箱の同じ名前のファイルは使わない。ほかの方の同じ名前のファイルを更新して共有を変えられなくても、もうリンクで見られれば成功
{
  const mk = (id, o) => Object.assign({ id, trashed: false, access: 'PRIVATE', canShare: true, data: '' }, o);
  const store = [mk('T1', { trashed: true }), mk('O1', { access: 'ANYONE_WITH_LINK', canShare: false })];
  const fobj = (x) => ({ getId: () => x.id, isTrashed: () => x.trashed, getUrl: () => 'https://drive.test/' + x.id,
    setTrashed: (t) => { x.trashed = t; }, getSharingAccess: () => x.access,
    setSharing: () => { if (!x.canShare) throw new Error('見本：共有を変える権限がありません'); x.access = 'ANYONE_WITH_LINK'; } });
  let byName = [];
  const folder = { getName: () => '03_生成物', getFilesByName: () => { let i = 0; const a = byName.map(fobj); return { hasNext: () => i < a.length, next: () => a[i++] }; },
                   createFile: () => { const x = mk('N' + store.length, { canShare: false }); store.push(x); return fobj(x); } };
  F.getAssetFolder_ = () => folder;
  F.DriveApp.Access = { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK', ANYONE: 'ANYONE' };
  F.DriveApp.Permission = { VIEW: 'VIEW' };
  F.Drive.Files.update = (meta, id) => ({ id });
  byName = [store[0], store[1]];
  let s1 = F.saveOutputFile_(F.Utilities.newBlob('x', 'text/plain', 'a.txt'), 'a.txt');
  ck(s1.id === 'O1' && s1.shared === true && s1.shareError === '', 'ゴミ箱のファイルを使った・もうリンクで見られるのに失敗にした: ' + J(s1));
  // フォルダが「リンクを知っている全員」で共有されている：新しく作ったファイルもそのまま見られるので、共有は変えない
  folder.createFile = () => { const x = mk('L' + store.length, { canShare: false, access: 'ANYONE_WITH_LINK' }); store.push(x); return fobj(x); };
  byName = [];
  s1 = F.saveOutputFile_(F.Utilities.newBlob('x', 'text/plain', 'c.txt'), 'c.txt');
  ck(/^L/.test(s1.id) && s1.shared === true, 'フォルダの共有を受け継いだファイルなのに失敗にした: ' + J(s1));
  folder.createFile = () => { const x = mk('N' + store.length, { canShare: false }); store.push(x); return fobj(x); };
  byName = [store[0]];
  s1 = F.saveOutputFile_(F.Utilities.newBlob('x', 'text/plain', 'b.txt'), 'b.txt');
  ck(/^N/.test(s1.id) && s1.shared === false && /見本：共有を変える権限がありません/.test(s1.shareError), '共有できないときの理由: ' + J(s1));
}

if (fails.length) { console.log('NG ' + fails.length + '件 / ' + checks + '件の検査'); fails.forEach((f) => console.log('  - ' + f)); process.exit(1); }
console.log('ドライブのメンバーブックの差し替え: 検査 ' + checks + ' 件 OK');
