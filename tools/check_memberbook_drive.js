// ドライブのメンバーブック（メールで送るPDF）の差し替え（memberbook_srv.js の saveMemberBookPdfToDrive・コード.js の
// uploadMemberBookBlob_・memberBookFileInfo_）を、見せかけのドライブで確かめる。名前はすべて架空。
//
//   node tools/check_memberbook_drive.js
//
// 確かめること
//   ・登録してあるファイル（MEMBER_BOOK_ID）の中身を差し替え、差し替えたファイルの名前・持ち主・ドライブの更新日時を知らせる
//   ・覚えている URL（MEMBER_BOOK_URL）が別のファイルのものでも、差し替えたファイルの URL にそろえる
//     （同じ名前の「MemberBook.pdf」がいくつもあり、見ているファイルと差し替えたファイルが違っていた）
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
env.reset([], { MEMBER_BOOK_ID: 'MB1', MEMBER_BOOK_URL: 'https://drive.test/file/d/OTHER/view' });

let clock = Date.UTC(2026, 9, 14, 1, 2, 0), stuck = false, updates = 0;
const files = { MB1: { name: 'MemberBook.pdf', owner: '見本 一郎', modified: clock, data: 'old' } };
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
ck(r.ok && files.MB1.data === '%PDF-1.4 見本' && updates === 1, '差し替えていない: ' + J([r, files.MB1, updates]));
ck(r.url === 'https://drive.test/file/d/MB1/view' && env.props.MEMBER_BOOK_URL === 'https://drive.test/file/d/MB1/view',
   '覚えている URL が別のファイルのまま: ' + J([r.url, env.props.MEMBER_BOOK_URL]));
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
ck(!r.error && /更新しました/.test(r.msg) && /差し替えたファイル：「MemberBook\.pdf」/.test(r.msg) && r.url === 'https://drive.test/file/d/MB1/view', 'アップロードからの差し替え: ' + J(r));

// ===== 画面に出す様子 =====
const info = F.getMemberBookFileInfo();
ck(info.ok && info.file.name === 'MemberBook.pdf' && info.file.owner === '見本 一郎' && /^2026\/10\/14 /.test(info.file.modified) && info.file.size === 2,
   '登録してあるファイルの様子: ' + J(info));
ck(F.memberBookDriveInfo_().file.name === 'MemberBook.pdf', '編集画面に渡す様子にファイルが無い');
env.props.MEMBER_BOOK_ID = 'GONE';
ck(/not found/.test(F.getMemberBookFileInfo().file.error), '見つからないファイルの様子: ' + J(F.getMemberBookFileInfo()));

if (fails.length) { console.log('NG ' + fails.length + '件 / ' + checks + '件の検査'); fails.forEach((f) => console.log('  - ' + f)); process.exit(1); }
console.log('ドライブのメンバーブックの差し替え: 検査 ' + checks + ' 件 OK');
