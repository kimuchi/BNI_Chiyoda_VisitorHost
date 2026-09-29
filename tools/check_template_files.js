// 「PowerPointテンプレート（ビジター用）」の登録（assets.js・template_files.html）を、実データなしで確かめる。
//
//   node tools/check_template_files.js
//
//   ・登録してあるのに、その方のアカウントで開けないテンプレートを「未登録」と出さない（理由と直し方を出す）
//   ・登録済みのテンプレートを置き換える前に確かめる（チャプター全体のテンプレートが替わるため）
//   ・設定した素材フォルダを開けない方は、登録・写真の追加・写真の索引の作り直しを止める
//     （別のフォルダに置いたり、みんなの写真の索引を空にしたりしないように）

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const { loadPage } = require('./lib_minidom');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

const env = makeEnv({ now: new Date(2026, 8, 29, 10, 0, 0) });
const srv = Object.assign({}, env.globals);
vm.createContext(srv);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), srv, { filename: f });
const KINDS = vm.runInContext('TEMPLATE_KINDS_', srv);
const kinds = Object.keys(KINDS);

env.reset([['メンバー名簿', false, [['No', '業種区分', '氏名'], ['1', '', '見本 一郎']]]], {});
env.addFile('TPL1', new env.FakeBlob('pptx', 'application/zip', kinds[0] + '_見本.pptx'), kinds[0] + '_見本.pptx');
env.props[KINDS[kinds[0]].prop] = 'TPL1';            // 1つ目：開ける
env.props[KINDS[kinds[1]].prop] = 'TPL2';            // 2つ目：登録してあるが、このアカウントでは開けない
const st = srv.getTemplateStatus();
const row = (k) => st.templates.find((t) => t.kind === k);
ck(row(kinds[0]).registered === true && !row(kinds[0]).unreadable, '開ける登録が「登録済み」にならない');
ck(row(kinds[1]).registered === false && /開けない|権限|許可/.test(row(kinds[1]).unreadable || ''), '開けない登録の理由が無い: ' + JSON.stringify(row(kinds[1])));
ck(row(kinds[2]).registered === false && !row(kinds[2]).unreadable, '未登録のものに理由が付いた');

// 画面：開けない登録は「未登録」と出さない。登録済みを置き換える前に確かめる
const page = loadPage('template_files.html', { fails, server: { getTemplateStatus: () => JSON.parse(JSON.stringify(st)), saveTemplateBase64: () => ({ ok: true, message: '登録しました。' }) } });
page.step('開く', () => page.window.onload());
const listHtml = page.els.list.innerHTML, part = (k) => (listHtml.split('id="s_' + k + '"')[1] || '').split('</div></div>')[0];
ck(/このアカウントでは開けません/.test(part(kinds[1])) && !/未登録/.test(part(kinds[1])), '開けない登録を「未登録」と出した: ' + part(kinds[1]).slice(0, 120));
ck(/未登録/.test(part(kinds[2])), '未登録のものが「未登録」と出ない: ' + part(kinds[2]).slice(0, 120));
page.sandbox.FileReader = class { readAsDataURL() { this.result = 'data:application/zip;base64,AAAA'; this.onload(); } };
page.els['f_' + kinds[1]].files = [{ name: 'new.pptx' }];
page.step('開けない登録を置き換える', () => page.run(`up('${kinds[1]}')`));
ck(page.log.confirms.length === 1 && /登録済み/.test(page.log.confirms[0]), '登録済みを置き換える前に確かめない');
page.els['f_' + kinds[2]].files = [{ name: 'new.pptx' }];
page.step('未登録に登録する', () => page.run(`up('${kinds[2]}')`));
ck(page.log.confirms.length === 1, '未登録のものを登録するのに確かめた');

// 素材フォルダを開けない方：登録・写真の追加・索引の作り直しを止める（別のフォルダに置かない・索引を空にしない）
env.props[vm.runInContext('ASSET_ROOT_KEY_', srv)] = 'NOT_SHARED';
env.denied.add('NOT_SHARED');                            // この方には共有されていない素材フォルダ
const reg = srv.saveTemplateBase64(kinds[2], Buffer.from('x').toString('base64'), 'new.pptx');
ck(reg.ok === false, '素材フォルダを開けない方の登録を止めない: ' + JSON.stringify(reg));
const up = srv.uploadMemberPhotosBase64([{ name: '見本 一郎.png', base64: 'AAAA' }]);
ck(up.ok === false && /素材フォルダ/.test(up.message), '素材フォルダを開けない方の写真の追加を止めない: ' + JSON.stringify(up));
const idx = srv.rebuildPhotoIndex();
ck(idx.ok === false && /素材フォルダ/.test(idx.message), '素材フォルダを開けない方の索引の作り直しを止めない: ' + JSON.stringify(idx));

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('PowerPointテンプレートの登録: 検査 ' + checks + ' 件 OK: 開けない登録を未登録と出さない・置き換える前に確かめる・素材フォルダを開けない方は止める');
