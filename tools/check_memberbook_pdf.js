// メンバーブックのPDFを編集画面の中で作る（memberbook_pdf.html）ことと、ドライブのメンバーブック（メールで送るPDF）を
// 差し替える（memberbook_srv.js の saveMemberBookPdfToDrive）ことを確かめる。名簿・写真は作り物。
//
//   node tools/check_memberbook_pdf.js [Googleフォントを写したディレクトリ]
//   （css.txt と gstatic/ のディレクトリを渡すと、書体（Noto Serif JP）を埋め込むところも確かめる。渡さなければパソコンの書体で描く）
//
// 確かめること
//   1. サーバー：登録してあるファイルの中身だけを差し替える（URL・名前はそのまま）。ゴミ箱にあれば戻してから。
//      開けない・差し替えられないときは新しく作らない（URLを変えない）。「作り直す」ときだけ新しく作る。
//      まだ無いときは作って登録する（リンクを知っている全員が見られるように）。ほかの方が更新中・空のデータ。
//      「メンバーブック(PDF)の更新」（アップロード）も、差し替えた日時を残す
//   2. 画面の中で作ったPDF：PDFとして読める（相互参照表・ページ・A4）。ページの画像（JPEG）の大きさ。
//      見えない文字（検索・コピー用）：お名前・会社名・表紙の題が入る。枠で切って見えていない文字は入れない。
//      見た目がふつうに開いた冊子（印刷と同じ）と同じ。写真は縮めてから使う。大きくなりすぎたら小さくして作り直す
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();

const ROOT = path.join(__dirname, '..');
const FONT_DIR = process.argv[2] || '';
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

// ===================== 1. サーバー（ドライブの差し替え）=====================
// Drive の作り物：files … id → { trashed, content, name, parents, shared, owner }。
//   readOnly … 見られるが編集の権限が無いファイル（差し替えると 403）。transient … 差し替えで一時的なエラー
function server(opts) {
  const o = opts || {};
  const props = Object.assign({}, o.props || {});
  const files = Object.assign({}, o.files || {});
  const log = [];
  let seq = 0;
  const box = {
    console: { log() {}, warn() {}, error() {} },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => !o.busy, releaseLock() {} }) },
    Utilities: {
      sleep: (ms) => { log.push(['sleep', ms]); },
      base64Decode: (s) => Array.from(Buffer.from(s, 'base64')).map((b) => (b > 127 ? b - 256 : b)),
      base64Encode: (a) => Buffer.from(a.map((b) => b & 255)).toString('base64'),
      newBlob: (bytes, type, name) => {
        let nm = name;
        return { _bytes: Buffer.from(bytes.map((b) => b & 255)), getContentType: () => type, getName: () => nm, setName(n) { nm = n; return this; } };
      },
    },
    Drive: { Permissions: {
      create: (res, id) => {
        if (o.noShare) throw new Error('API call to drive.permissions.create failed with error: The user does not have sufficient permissions for this file.');
        log.push(['share-api', id, J(res)]); files[id].shared = true; return { id: 'anyoneWithLink' };
      },
    }, Files: {
      get: (id, args) => {
        log.push(['get', id, args && args.fields]);
        if (o.scope) throw new Error('API call to drive.files.get failed with error: Insufficient Permission: Request had insufficient authentication scopes.');
        if (!files[id]) throw new Error('API call to drive.files.get failed with error: File not found: ' + id + '.');
        if (args && /owners/.test(args.fields)) return { owners: [{ displayName: files[id].owner || '見本 管理者', emailAddress: 'owner@example.com' }] };
        return { id, trashed: !!files[id].trashed };
      },
      update: (res, id, blob, args) => {
        log.push(['update', id, J(res), blob ? blob._bytes.toString('latin1').slice(0, 8) : null, J(args || null)]);
        if (!files[id]) throw new Error('API call to drive.files.update failed with error: File not found: ' + id + '.');
        if (o.readOnly && o.readOnly.includes(id) && blob) throw new Error('API call to drive.files.update failed with error: The user does not have sufficient permissions for file ' + id + '.');
        if (o.transient && blob) throw new Error('API call to drive.files.update failed with error: Internal Error');
        if (o.untrashDenied && res && res.trashed === false) throw new Error('API call to drive.files.update failed with error: The user does not have sufficient permissions for file ' + id + '.');
        if (res && res.trashed === false) files[id].trashed = false;
        if (blob) files[id].content = blob._bytes;
        return { id };
      },
      create: (res, blob, args) => {
        if (o.parentReadOnly && res.parents) { log.push(['create-denied', J(res.parents)]); throw new Error('API call to drive.files.create failed with error: Insufficient permissions for the specified parent.'); }
        const id = 'new' + (++seq);
        log.push(['create', id, res.name, J(res.parents || null), J(args || null)]);
        files[id] = { name: res.name, parents: res.parents || null, content: blob._bytes, trashed: false };
        return { id };
      },
    } },
    DriveApp: {
      Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK', ANYONE: 'ANYONE', PRIVATE: 'PRIVATE' }, Permission: { VIEW: 'VIEW' },
      getFileById: (id) => {
        if (o.photos && id in o.photos) {
          if (o.photos[id] === 'broken') throw new Error('写真を読めない');
          if (o.photos[id] === 'gone') throw new Error('No item with the given ID could be found. Or perhaps you do not have permission to access it.');
          return { getBlob: () => ({ getContentType: () => 'image/jpeg', getBytes: () => Array.from(Buffer.from(o.photos[id])) }) };
        }
        return {
          getUrl: () => 'https://drive.test/file/d/' + id + '/view',
          setSharing: (a, p) => {
            if (o.noShare || (o.shareFlaky && o.shareFlaky-- > 0)) throw new Error('共有は組織の設定で禁止');
            log.push(['share', id, a, p]); files[id].shared = true;
          },
          getSharingAccess: () => (files[id] && files[id].shared ? 'ANYONE_WITH_LINK' : 'PRIVATE'),
        };
      },
    },
  };
  vm.createContext(box);
  for (const f of ['コード.js', 'assets.js', 'memberbook_srv.js']) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), box, { filename: f });
  // 素材フォルダの 03_生成物（無いときは作れない → マイドライブ）
  box.getAssetFolder_ = (kind) => { if (o.noFolder) throw new Error('素材フォルダが無い'); return { getId: () => 'OUT-' + kind }; };
  box.findPhotoIdForName_ = (n) => (o.photoIds || {})[n] || '';
  return { box, props, files, log };
}
const PDF64 = Buffer.from('%PDF-1.4 見本のPDF').toString('base64');
const reg = { MEMBER_BOOK_ID: 'mb', MEMBER_BOOK_URL: 'https://drive.test/file/d/mb/view' };
const pdfBlob = (S) => S.box.Utilities.newBlob(Array.from(Buffer.from('%PDF-1.4')), 'application/pdf', 'a.pdf');
{
  // まだ無いとき：素材フォルダの 03_生成物 に新しく作って登録（リンク共有）
  let S = server();
  let r = S.box.saveMemberBookPdfToDrive(PDF64, 'MemberBook.pdf');
  ck(r.ok && r.created && S.props.MEMBER_BOOK_ID === 'new1' && S.props.MEMBER_BOOK_URL === 'https://drive.test/file/d/new1/view' && r.url === S.props.MEMBER_BOOK_URL,
     'まだ無いとき: ' + J(r) + ' ' + J(S.props));
  ck(S.files.new1.shared && J(S.files.new1.parents) === J(['OUT-output']) && S.files.new1.content.toString('latin1').startsWith('%PDF') && /新しく作って登録/.test(r.message),
     'まだ無いとき（置き場所・共有・中身・知らせ）: ' + J(S.log));
  ck(!!S.props.MEMBER_BOOK_UPDATED && /^\d{4}-\d{2}-\d{2}T/.test(S.props.MEMBER_BOOK_UPDATED), '差し替えた日時が残らない: ' + S.props.MEMBER_BOOK_UPDATED);
  ck(r.downloadUrl === 'https://drive.google.com/uc?export=download&id=new1', 'ダウンロードのURL: ' + r.downloadUrl);
  S = server({ noFolder: true });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(r.ok && r.created && S.files.new1.parents === null, '素材フォルダが無いときはマイドライブに作る: ' + J(S.log));
  S = server({ parentReadOnly: true });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(r.ok && r.created && S.files.new1.parents === null && S.log.some((x) => x[0] === 'create-denied'), '03_生成物に書けないときはマイドライブに作る: ' + J(r) + ' ' + J(S.log));

  // 登録してあるとき：中身だけ差し替える（新しく作らない・URLそのまま・共有ドライブでも）
  S = server({ props: reg, files: { mb: { name: '配布用.pdf', content: Buffer.from('old'), trashed: false } } });
  r = S.box.saveMemberBookPdfToDrive(PDF64, 'MemberBook.pdf');
  ck(r.ok && !r.created && r.url === reg.MEMBER_BOOK_URL && S.props.MEMBER_BOOK_ID === 'mb' && S.props.MEMBER_BOOK_URL === reg.MEMBER_BOOK_URL,
     '登録してあるとき: ' + J(r) + ' ' + J(S.props));
  ck(!S.log.some((x) => x[0] === 'create') && S.files.mb.content.toString('latin1').startsWith('%PDF') && S.files.mb.name === '配布用.pdf',
     '中身だけ差し替えていない（新しく作った・名前が変わった）: ' + J(S.log));
  ck(S.log.filter((x) => x[0] === 'update').every((x) => x[2] === '{}' && /supportsAllDrives/.test(x[4])), '差し替えで名前などを書き換えた・共有ドライブに対応していない: ' + J(S.log));
  ck(/URLはそのまま/.test(r.message) && !!S.props.MEMBER_BOOK_UPDATED, '差し替えの知らせ・日時: ' + r.message);

  // ゴミ箱にあるとき：戻してから差し替える
  S = server({ props: reg, files: { mb: { content: Buffer.from('old'), trashed: true } } });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(r.ok && !r.created && S.files.mb.trashed === false && S.files.mb.content.toString('latin1').startsWith('%PDF'), 'ゴミ箱にあるとき: ' + J(S.log));

  // 見つからない（削除した）：新しく作らない。登録（URL）はそのまま。「新しく作り直す」を出す
  S = server({ props: reg, files: {} });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(!r.ok && r.canRecreate && !S.log.some((x) => x[0] === 'create') && S.props.MEMBER_BOOK_URL === reg.MEMBER_BOOK_URL && !S.props.MEMBER_BOOK_UPDATED,
     '削除したとき: ' + J(r) + ' ' + J(S.log));
  ck(/見つかりませんでした/.test(r.message) && /新しく作り直す/.test(r.message), '削除したときの知らせ: ' + r.message);
  // 編集の権限が無い：持ち主に編集者にしてもらう。作り直しは出さない（URLが変わるため）。権限の許可（OAuth）の案内は出さない
  S = server({ props: reg, files: { mb: { owner: '見本 管理者' } }, readOnly: ['mb'] });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(!r.ok && !r.canRecreate && !S.log.some((x) => x[0] === 'create') && S.props.MEMBER_BOOK_URL === reg.MEMBER_BOOK_URL,
     '編集の権限が無いとき: ' + J(r) + ' ' + J(S.log));
  ck(/権限がありません/.test(r.message) && /見本 管理者/.test(r.message) && /編集者/.test(r.message) && !/動作確認/.test(r.message) && !/「新しく作り直す」を押すと/.test(r.message),
     '編集の権限が無いときの知らせ: ' + r.message);
  // このアプリへのGoogleドライブの許可が足りない：許可し直す案内（持ち主に頼む案内ではない）
  S = server({ props: reg, files: { mb: {} }, scope: true });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(!r.ok && !r.canRecreate && /動作確認/.test(r.message) && !/編集者/.test(r.message) && !S.log.some((x) => x[0] === 'create'), '許可が足りないとき: ' + J(r));
  // 持ち主がゴミ箱に入れた（戻せるのは持ち主だけ）：戻してもらう案内
  S = server({ props: reg, files: { mb: { trashed: true } }, untrashDenied: true });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(!r.ok && !r.canRecreate && /ゴミ箱/.test(r.message) && /戻してもらって/.test(r.message) && !S.log.some((x) => x[0] === 'create'), '持ち主のゴミ箱にあるとき: ' + J(r));
  // 一時的なエラー：もう一度。作り直しは出さない
  S = server({ props: reg, files: { mb: {} }, transient: true });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(!r.ok && !r.canRecreate && !S.log.some((x) => x[0] === 'create') && /もう一度/.test(r.message), '一時的なエラー: ' + J(r));
  // 「新しく作り直す」：新しく作って、登録を替える
  S = server({ props: reg, files: {} });
  r = S.box.saveMemberBookPdfToDrive(PDF64, 'MemberBook.pdf', true);
  ck(r.ok && r.created && S.props.MEMBER_BOOK_ID === 'new1' && S.props.MEMBER_BOOK_URL !== reg.MEMBER_BOOK_URL && /作り直して/.test(r.message), '作り直す: ' + J(r));
  // 共有を設定できない（組織の設定）：作ったことは知らせて、共有を確かめるよう添える
  S = server({ noShare: true });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(r.ok && r.created && /共有を設定できませんでした/.test(r.message), '共有できないとき: ' + J(r));
  // 作った直後に共有を断られても、少し待ってやり直して共有する
  S = server({ shareFlaky: 1 });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(r.ok && r.created && S.files.new1.shared && !/共有を設定できませんでした/.test(r.message) && S.log.some((x) => x[0] === 'sleep'),
     '共有を一時的に断られたとき、やり直さない: ' + J(r) + ' ' + J(S.log));
  // ほかの方が更新中・空のデータ
  S = server({ props: reg, files: { mb: {} }, busy: true });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(!r.ok && /更新中/.test(r.message) && !S.log.length, 'ほかの方が更新中: ' + J(r));
  S = server({ props: reg, files: { mb: {} } });
  r = S.box.saveMemberBookPdfToDrive('');
  ck(!r.ok && !S.log.length, '空のデータ: ' + J(r));
  // 編集画面に出す様子
  S = server({ props: Object.assign({ MEMBER_BOOK_UPDATED: '2026-09-28T06:00:00.000Z' }, reg) });
  ck(J(S.box.memberBookDriveInfo_()) === J({ id: 'mb', url: reg.MEMBER_BOOK_URL, updated: '2026-09-28T06:00:00.000Z' }), 'ドライブの様子: ' + J(S.box.memberBookDriveInfo_()));

  // 「メンバーブック(PDF)の更新」（アップロード）：差し替えはURLそのまま。見つからないときだけ新しく作り、URLが変わったと知らせる。
  // 編集の権限が無い・一時的なエラーのときは、新しく作らない（勝手にURLが変わらないように）
  S = server({ props: reg, files: { mb: {} } });
  let u = S.box.uploadMemberBookBlob_(pdfBlob(S));
  ck(/更新しました/.test(u.msg) && u.url === reg.MEMBER_BOOK_URL && !!S.props.MEMBER_BOOK_UPDATED && !S.log.some((x) => x[0] === 'create'), 'アップロードで差し替え: ' + J(u));
  S = server();
  u = S.box.uploadMemberBookBlob_(pdfBlob(S));
  ck(u.msg === '新規登録しました。' && S.props.MEMBER_BOOK_ID === 'new1' && S.files.new1.shared && !!S.props.MEMBER_BOOK_UPDATED, 'アップロードで新しく登録: ' + J(u));
  S = server({ props: reg, files: {} });
  u = S.box.uploadMemberBookBlob_(pdfBlob(S));
  ck(/見つからない/.test(u.msg) && /URLが変わりました/.test(u.msg) && S.props.MEMBER_BOOK_ID === 'new1', 'アップロード（削除してあった）: ' + J(u));
  for (const [label, o] of [['編集の権限が無い', { readOnly: ['mb'] }], ['一時的なエラー', { transient: true }]]) {
    S = server(Object.assign({ props: reg, files: { mb: {} } }, o));
    u = S.box.uploadMemberBookBlob_(pdfBlob(S));
    ck(u.error && !u.msg && !S.log.some((x) => x[0] === 'create') && S.props.MEMBER_BOOK_URL === reg.MEMBER_BOOK_URL && !S.props.MEMBER_BOOK_UPDATED,
       `アップロード（${label}）: 新しく作った・URLが変わった: ` + J(u) + ' ' + J(S.log));
  }
  S = server({ noShare: true });
  u = S.box.uploadMemberBookBlob_(pdfBlob(S));
  ck(/共有を設定できませんでした/.test(u.msg), 'アップロード（共有できない）の知らせ: ' + J(u));

  // 写真の実体：写真の無い方（missing）・一覧にあるのにファイルが無い方（gone）・あるのに読めなかった方（failed）を分ける
  S = server({ photoIds: { '見本 一郎': 'p1', '見本 花子': 'p2', '見本 四郎': 'p4' }, photos: { p1: 'JPEG', p2: 'broken', p4: 'gone' } });
  const ph = S.box.getMemberPhotosBase64(['見本 一郎', '見本 花子', '見本 三郎', '見本 四郎']);
  ck(ph.ok && Object.keys(ph.map).join() === '見本 一郎' && J(ph.missing) === J(['見本 三郎']) && J(ph.failed) === J(['見本 花子']) && J(ph.gone) === J(['見本 四郎']),
     '写真の実体: ' + J(ph));
}

// ===================== 2. 画面の中で作るPDF =====================
// PDF を読む（相互参照表の位置・オブジェクト・ストリーム）
function readPdf(buf) {
  const s = buf.toString('latin1'), out = { errors: [], objs: {} };
  if (!s.startsWith('%PDF-1.')) out.errors.push('先頭が %PDF- でない');
  const m = s.match(/startxref\n(\d+)\n%%EOF\n?$/);
  if (!m) { out.errors.push('startxref が無い'); return out; }
  const x = +m[1], hm = s.slice(x).match(/^xref\n0 (\d+)\n0000000000 65535 f \n/);
  if (!hm) { out.errors.push('相互参照表が無い'); return out; }
  const n = +hm[1];
  for (let i = 1; i < n; i++) {
    const e = s.substr(x + hm[0].length + (i - 1) * 20, 20);
    if (!/^\d{10} 00000 n \n$/.test(e)) { out.errors.push(i + ' 番の相互参照の行: ' + J(e)); continue; }
    const off = parseInt(e, 10), head = i + ' 0 obj\n';
    if (s.substr(off, head.length) !== head) { out.errors.push(i + ' 番の位置が違う'); continue; }
    const at = off + head.length, dm = s.slice(at).match(/^(<<[\s\S]*?>>)\n(stream\n)?/);
    if (!dm) { out.errors.push(i + ' 番の中身が読めない'); continue; }
    const obj = { dict: dm[1] };
    if (dm[2]) {
      const len = +(dm[1].match(/\/Length (\d+)/) || [])[1], from = at + dm[0].length;
      obj.data = buf.subarray(from, from + len);
      if (s.substr(from + len, 18) !== '\nendstream\nendobj\n') out.errors.push(i + ' 番のストリームの長さが違う');
    } else if (s.substr(at + dm[1].length, 8) !== '\nendobj\n') out.errors.push(i + ' 番の終わりが違う');
    out.objs[i] = obj;
  }
  const tr = s.slice(x).match(/trailer\n<<\/Size (\d+)\/Root (\d+) 0 R\/Info (\d+) 0 R>>/);
  if (!tr || +tr[1] !== n) out.errors.push('trailer: ' + (tr && tr[0]));
  out.root = tr && out.objs[+tr[2]]; out.info = tr && out.objs[+tr[3]];
  return out;
}
const ucs2 = (h) => { let t = ''; for (let i = 0; i + 4 <= h.length; i += 4) t += String.fromCharCode(parseInt(h.substr(i, 4), 16)); return t; };
const jpegSize = (b) => {
  if (b[0] !== 0xFF || b[1] !== 0xD8) return null;
  for (let i = 2; i + 9 < b.length;) {
    if (b[i] !== 0xFF) return null;
    const mk = b[i + 1];
    if ((mk >= 0xC0 && mk <= 0xC3) || (mk >= 0xC5 && mk <= 0xC7) || (mk >= 0xC9 && mk <= 0xCB) || (mk >= 0xCD && mk <= 0xCF)) return { h: b.readUInt16BE(i + 5), w: b.readUInt16BE(i + 7) };
    i += 2 + b.readUInt16BE(i + 2);
  }
  return null;
};
// PDF のページ → { box, image: { w, h, jpeg }, text: 行の文字の並び, lines }
function pagesOf(pdf) {
  const pagesObj = pdf.objs[+((pdf.root && pdf.root.dict.match(/\/Pages (\d+) 0 R/)) || [])[1]];
  const kids = pagesObj ? (pagesObj.dict.match(/\/Kids\[([^\]]*)\]/)[1].match(/\d+(?= 0 R)/g) || []).map(Number) : [];
  return kids.map((k) => {
    const d = pdf.objs[k].dict, im = pdf.objs[+d.match(/\/Im0 (\d+) 0 R/)[1]], ct = pdf.objs[+d.match(/\/Contents (\d+) 0 R/)[1]];
    const c = ct.data.toString('latin1'), lines = [];
    const re = /([\d.-]+) 0 0 ([\d.-]+) ([\d.-]+) ([\d.-]+) Tm <([0-9A-F]*)> Tj/g;
    let mm;
    while ((mm = re.exec(c)) !== null) lines.push({ sx: +mm[1], size: +mm[2], x: +mm[3], y: +mm[4], s: ucs2(mm[5]) });
    return { dict: d, box: (d.match(/\/MediaBox\[([^\]]*)\]/) || [])[1], content: c, lines, text: lines.map((l) => l.s).join('\n'),
             image: { dict: im.dict, jpeg: im.data, size: jpegSize(im.data) } };
  });
}

// 作り物の冊子（22名：表紙＋2ページ）
const CATS = [['企業サポート', '#FFDE58'], ['研修・教育', '#C1FF72'], ['不動産関連', '#37B5FF'], ['建築・住まい', '#AAB5D9']].map(([key, bg]) => ({ key, label: key, bg }));
const LONG = '見本の長い文です。';
const MEMBERS = [];
for (let i = 0; i < 22; i++) {
  const k = i % 5;
  MEMBERS.push({ no: String(i + 1), name: ['見本 一郎', '見本 花子', '見本 太郎左衛門', '見本 三郎', 'Mihon Taro'][k] + (i >= 5 ? i : ''), cat: CATS[i % 4].key,
    title: ['税理士', '不動産売買仲介（相続・事業承継）', '内装工事', '生命保険（法人）', '行政書士（建設業許可・相続）'][k],
    company: ['見本株式会社', '一般社団法人見本フォーラム', '(株)見本', '見本生命保険株式会社 東京本店', 'MIHON Holdings Co., Ltd.'][k],
    position: ['代表取締役', '', '取締役 兼 営業本部長', '支店長', ''][k], role: ['ビジターホスト', 'Webチーム', '', 'プレジデント', ''][k],
    comment: ['見本の一言です', LONG.repeat(2), LONG.repeat(3), '切れて見えない始めの文。' + LONG.repeat(40) + '切れて見えない終わりの文', ''][k],
    refer: ['税理士・弁護士', '新しく店舗を出す飲食店のオーナー', LONG.repeat(3), LONG.repeat(9), '司法書士'][k],
    collab: ['司法書士', '保険・不動産・士業', LONG.repeat(2), LONG.repeat(9), ''][k] });
}
// 特別な文字：異体字セレクタ付きの名前・濁点が分かれた名前（NFD）・BMP の外の字・PowerPoint の改行（U+000B）などの制御文字
MEMBERS[19].name = '葛\u{E0100}城 見本';
MEMBERS[20].name = '見本 カ\u3099ク';
MEMBERS[21].name = '\u{20BB7}田 見本';
MEMBERS[19].comment = '一行め\u000B二行め';
MEMBERS[20].company = '見本\u000C商事';
// 組になっていないサロゲート（絵文字の半分）：末尾に上位だけ・下位が2つ続く
MEMBERS[18].company = '見本物産\uD83C';
MEMBERS[18].refer = '税理士\uDE00\uDE01・弁護士';
const COVER = { title: 'BNI 見本 chapter Member Book \u{20BB7}', term: '24期', pname: '見本 会長', prole: '見本チャプター\n第24期プレジデント',
  ptext: 'あいさつの文です。'.repeat(40), philosophyTitle: 'BNIの理念', philosophy: '理念の本文です。'.repeat(6), aboutTitle: 'チャプターとは？',
  about: '紹介の本文です。'.repeat(5), benefitsTitle: 'メリット', benefits: '1.メリット　2.メリット\n3.メリット',
  scheduleFrom: '7:15', scheduleTo: '9:15', schedule: Array.from({ length: 12 }, (_, i) => '項目' + (i + 1)).join('\n'),
  termsTitle: 'BNIの用語の説明', terms: '①チャプター\n説明の文です。\n②リファーラル\n説明の文です。' };

async function routeFonts(ctx) {
  if (!FONT_DIR) { await ctx.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.abort()); return; }
  const css = fs.readFileSync(path.join(FONT_DIR, 'css.txt'), 'utf8'), hdr = { 'access-control-allow-origin': '*' };
  await ctx.route('https://fonts.googleapis.com/**', (r) => r.fulfill({ status: 200, contentType: 'text/css', body: css, headers: hdr }));
  await ctx.route('https://fonts.gstatic.com/**', (r) => {
    const f = path.join(FONT_DIR, 'gstatic', r.request().url().replace('https://fonts.gstatic.com/', '').replace(/\//g, '_'));
    if (!fs.existsSync(f)) return r.fulfill({ status: 404, body: '' });
    r.fulfill({ status: 200, contentType: 'font/woff2', body: fs.readFileSync(f), headers: hdr });
  });
}
const strip = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8').replace(/^<script>|<\/script>\s*$/g, '');

(async () => {
  // 文字の色つきのにじみ（LCD）は、ふつうの画面にだけ付き、SVG の画像には付かない。比べるときに差にならないよう、はじめから外す
  const browser = await pw.chromium.launch({ args: ['--disable-lcd-text'] });
  const ctx = await browser.newContext();
  await routeFonts(ctx);
  const page = await ctx.newPage();
  await page.route('https://mb.test/', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: '<!DOCTYPE html><html><body></body></html>' }));
  await page.goto('https://mb.test/');
  const R = await page.evaluate(async ({ render, pdfjs, members, cats, cover }) => {
    window.members = members; window.cats = cats; window.cover = cover; window.photosB64 = {};
    window.catOf = (k) => cats.find((c) => c.key === k) || null;
    (0, eval)(render); (0, eval)(pdfjs);
    // 写真：大きなもの（元の写真の代わり）と、小さなもの
    const photo = (hue, w, h) => { const c = document.createElement('canvas'); c.width = w; c.height = h; const g = c.getContext('2d');
      g.fillStyle = 'hsl(' + hue + ',60%,55%)'; g.fillRect(0, 0, w, h); g.fillStyle = '#fff'; g.beginPath(); g.arc(w / 2, h / 3, w / 4, 0, 7); g.fill();
      return c.toDataURL('image/jpeg', 0.92); };
    const photos = {};
    members.forEach((m, i) => { if (i % 3 !== 2) photos[m.name] = photo(i * 37, i % 2 ? 2400 : 300, i % 2 ? 3200 : 400); });
    photos[cover.pname] = photo(200, 3000, 2000);
    // 背景を抜いた（透明な）PNG の写真
    const png = (() => { const c = document.createElement('canvas'); c.width = 1000; c.height = 1300; const g = c.getContext('2d');
      g.fillStyle = '#c33'; g.beginPath(); g.arc(500, 500, 300, 0, 7); g.fill(); return c.toDataURL('image/png'); })();
    photos[members[2].name] = png;
    // ブラウザーで開けない写真（HEIC を .jpg にしたもの・空のファイル）
    photos[members[5].name] = 'data:image/jpeg;base64,AAAAGGZ0eXBoZWljAAAAAG1pZjE=';
    const size = (d) => new Promise((ok) => { const im = new Image(); im.onload = () => ok([im.naturalWidth, im.naturalHeight]); im.src = d; });
    const sr = await mbpSmallPhotos(photos), small = sr.map;
    const sizes = {};
    for (const n of Object.keys(small)) sizes[n] = await size(small[n]);
    const alpha = await new Promise((ok) => { const im = new Image(); im.onload = () => { const c = document.createElement('canvas'); c.width = im.naturalWidth; c.height = im.naturalHeight;
      const g = c.getContext('2d'); g.drawImage(im, 0, 0); ok(g.getImageData(1, 1, 1, 1).data[3]); }; im.src = small[members[2].name]; });
    const pngKind = small[members[2].name].slice(0, 15);
    const again = (await mbpSmallPhotos(photos)).map;
    const steps = [];
    const html = bookHtml({ photos: small });
    const t0 = performance.now();
    const pdf = await mbpBuildPdf(html, cover.title, (t) => steps.push(t));
    const ms = performance.now() - t0;
    // 書体の埋め込み（Googleフォントを読めたときだけ）
    const fr = await mbpOpenBook(html), loaded = mbpFontsLoaded(fr.contentDocument);
    const css = loaded ? await mbpFontCss(fr.contentDocument, fr.contentDocument.querySelectorAll('.page')[1].textContent, {}) : '';
    fr.parentNode.removeChild(fr);
    // 枠で切ってある文字は入れない（高さ20pxの枠に3行ぶんの文字）
    const tdoc = document.implementation.createHTMLDocument('t');
    tdoc.body.innerHTML = '<div class="page" style="position:relative;width:793px;height:1122px;">'
      + '<div style="position:absolute;left:10px;top:10px;width:160px;height:20px;overflow:hidden;font-size:16px;line-height:20px;">'
      + 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめも</div>'
      + '<div style="position:absolute;left:10px;top:100px;">見える文字<br>二行め</div></div>';
    const holder = document.createElement('div'); holder.appendChild(document.adoptNode(tdoc.body.firstChild)); document.body.appendChild(holder);
    const clipLines = mbpPageText(holder.querySelector('.page'));
    holder.remove();
    // 大きくなりすぎたとき：小さくして作り直す
    const keep = MBP_MAX_BYTES; MBP_MAX_BYTES = 400 * 1024;
    const steps2 = [];
    const pdf2 = await mbpBuildPdf(html, cover.title, (t) => steps2.push(t));
    MBP_MAX_BYTES = keep;
    return { pdf, ms, steps, sizes, alpha, pngKind, bad: sr.bad, sameAgain: Object.keys(small).every((n) => again[n] === small[n]), small: Object.keys(small).length,
             unchangedSmall: small[members[4].name] === photos[members[4].name], loaded, css: css.length, cssRules: (css.match(/@font-face/g) || []).length,
             clipLines, pdf2: { base64: pdf2.base64, bytes: pdf2.bytes }, steps2, html };
  }, { render: strip('memberbook_render.html'), pdfjs: strip('memberbook_pdf.html'), members: MEMBERS, cats: CATS, cover: COVER });

  // 写真を縮める（短い辺 360px まで。小さな写真はそのまま。2回目は覚えたものを使う）。透明な PNG は透明のまま
  ck(R.small === 17 && J(R.bad) === J([MEMBERS[5].name]), '写真の数・開けない写真: ' + R.small + ' ' + J(R.bad));
  ck(Object.values(R.sizes).every(([w, h]) => Math.min(w, h) <= 360), '写真が縮んでいない: ' + J(R.sizes));
  ck(J(R.sizes[MEMBERS[1].name]) === J([360, 480]) && J(R.sizes[COVER.pname]) === J([540, 360]) && J(R.sizes[MEMBERS[2].name]) === J([360, 468]),
     '縮めた写真の大きさ（縦横の比はそのまま・短い辺で）: ' + J(R.sizes));
  ck(R.pngKind === 'data:image/png;' && R.alpha === 0, '透明な PNG の写真が白く塗られた: ' + J([R.pngKind, R.alpha]));
  ck(R.unchangedSmall && R.sameAgain, '小さな写真はそのまま・2回目は覚えたもの: ' + J([R.unchangedSmall, R.sameAgain]));

  // PDF の作り
  const buf = Buffer.from(R.pdf.base64, 'base64'), pdf = readPdf(buf), pages = pagesOf(pdf);
  ck(!pdf.errors.length, 'PDFの作りが正しくない: ' + J(pdf.errors.slice(0, 5)));
  ck(buf.length === R.pdf.bytes && R.pdf.pages === 3 && pages.length === 3, 'ページ数: ' + R.pdf.pages + ' / ' + pages.length);
  ck(pages.every((p) => p.box === '0 0 595.28 841.89'), 'ページの大きさ（A4）: ' + J(pages.map((p) => p.box)));
  ck(pages.every((p) => p.image.size && Math.abs(p.image.size.w - 1984) <= 2 && Math.abs(p.image.size.h - 2806) <= 3
     && p.image.dict.includes('/Width ' + p.image.size.w + '/') && p.image.dict.includes('/Height ' + p.image.size.h + '/')
     && /\/Filter\/DCTDecode/.test(p.image.dict)), 'ページの画像（JPEG・約240dpi）: ' + J(pages.map((p) => [p.image.size, p.image.dict])));
  ck(pages.every((p) => /^q 595\.28 0 0 841\.89 0 0 cm \/Im0 Do Q\n/.test(p.content) && /BT 3 Tr \/F1 1 Tf\n/.test(p.content)), '画像を全面に・文字は描かない（3 Tr）');
  const font = Object.values(pdf.objs).find((o) => /\/Subtype\/Type0/.test(o.dict));
  ck(font && /\/Encoding\/UniJIS-UCS2-H/.test(font.dict) && /\/ToUnicode \d+ 0 R/.test(font.dict), '文字の書体（UniJIS-UCS2-H・ToUnicode）: ' + (font && font.dict));
  const info = pdf.info ? pdf.info.dict : '', tt = (info.match(/\/Title<FEFF([0-9A-F]*)>/) || [])[1];
  ck(tt && ucs2(tt) === COVER.title, 'PDFの題（BMP の外の字もそのまま）: ' + info);
  ck(R.pdf.bytes < 6 * 1024 * 1024, 'PDFが大きすぎる: ' + R.pdf.bytes);
  ck(R.steps.length === 3 && /3\/3ページ/.test(R.steps[2]), '進み具合の知らせ: ' + J(R.steps));

  // 見えない文字（検索・コピー用）
  const flat = (t) => t.replace(/\s/g, '');
  // 検索用の文字：合成した形（NFC）・異体字セレクタは除く・BMP の外の字は〓・制御文字は除く
  const searchable = (x) => flat(x.replace(/[\uD800-\uDFFF]/gu, '').normalize('NFC').replace(/[\uFE00-\uFE0F]|[\u{E0100}-\u{E01EF}]/gu, '').replace(/[\u0000-\u001F]/g, '')
    .replace(/[\u{10000}-\u{10FFFF}]/gu, '\u3013'));
  ck(flat(pages[0].text).includes(searchable(COVER.title)) && flat(pages[0].text).includes('タイムスケジュール') && flat(pages[0].text).includes('見本会長'),
     '表紙の文字: ' + pages[0].text.slice(0, 120));
  MEMBERS.forEach((m, i) => {
    const p = pages[1 + Math.floor(i / 18)], t = flat(p.text);
    ck(t.includes(searchable(m.name)) && t.includes(searchable(m.company)), `${m.no} ${m.name}: お名前・会社名が文字に無い`);
  });
  ck(pages[2].lines.some((l) => l.s === '葛城 見本') && pages[2].lines.some((l) => l.s === '見本 ガク') && pages[2].lines.some((l) => l.s === '〓田 見本'),
     '特別な文字の名前が1行の文字になっていない: ' + J(pages[2].lines.filter((l) => /見本/.test(l.s) && l.s.length < 8).map((l) => l.s)));
  ck(!pages.some((p) => /[\u0000-\u001F]/.test(p.text.replace(/\n/g, ''))) && flat(pages[2].text).includes('一行め二行め') && flat(pages[2].text).includes('見本商事'),
     '制御文字（U+000B など）があってもPDFが作れて、文字に制御文字が入らない');
  ck(pages[1].lines.some((l) => l.s === '見本 花子'), 'お名前が1行の文字として入っていない（検索できない）');
  // とても長い一言（入りきらず、枠で上下が切れる）：見えている行だけ入る
  ck(!flat(pages[1].text).includes('切れて見えない終わりの文') && !flat(pages[1].text).includes('切れて見えない始めの文')
     && flat(pages[1].text).includes('見本の長い文です。見本の長い文です。'), '枠で切ってある文字も入っている（または一言が無い）');
  if (process.env.DUMP) fs.writeFileSync(path.join(process.env.DUMP, 'text2.txt'), pages[1].lines.map((l) => J(l)).join('\n'));
  ck(J(R.clipLines.map((l) => l.s)) === J(['あいうえおかきくけこ', '見える文字', '二行め']), '枠で切ってある文字を入れない: ' + J(R.clipLines.map((l) => l.s)));
  ck(R.clipLines.every((l) => l.h > 10 && l.h < 20 && l.w > 20), '文字の位置と大きさ（pt）: ' + J(R.clipLines));
  // 文字の位置：お名前の行が、画像の中のお名前のところ（カードの中）にある
  const nm = pages[1].lines.find((l) => l.s === '見本 一郎');
  ck(nm && nm.x > 10 && nm.x < 150 && nm.y > 700 && nm.y < 800 && nm.size > 9 && nm.size < 18, 'お名前の文字の位置: ' + J(nm));

  // 書体の埋め込み
  if (FONT_DIR) ck(R.loaded && R.cssRules > 0 && R.css > 10000, '書体を埋め込んでいない: ' + J([R.loaded, R.cssRules, R.css]));
  else ck(!R.loaded && R.css === 0, '書体を読めないのに埋め込もうとした: ' + J([R.loaded, R.css]));

  // 大きくなりすぎたとき：小さくして作り直す
  const pages2 = pagesOf(readPdf(Buffer.from(R.pdf2.base64, 'base64')));
  ck(R.steps2.some((t) => /小さくして作り直し/.test(t)) && pages2.length === 3 && pages2[1].image.size.w < 1984 && R.pdf2.bytes < R.pdf.bytes,
     '大きすぎるときに作り直していない: ' + J([R.steps2.slice(-1), R.pdf2.bytes, pages2.map((p) => p.image.size)]));

  // 見た目：ふつうに開いた冊子（印刷と同じ）の画面と比べる（1/5 に縮めて、色の差の平均と、大きく違う点の割合）
  const ctx2 = await browser.newContext({ deviceScaleFactor: 2.5 });
  await routeFonts(ctx2);
  const p2 = await ctx2.newPage();
  await p2.route('https://mb.test/book', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: R.html }));
  await p2.goto('https://mb.test/book');
  await p2.waitForFunction(() => document.documentElement.getAttribute('data-fitted') === '1', null, { timeout: 30000 });
  await p2.addStyleTag({ content: '.page{margin:0!important;box-shadow:none!important;} body{padding:0!important;background:#fff!important;}' });
  const els = await p2.$$('.page');
  for (let i = 0; i < els.length; i++) {
    // そのページだけを出して撮る（ページの上端が小数の位置にあると、文字の置き方が1pxほどずれて比べにくい）
    await p2.evaluate((k) => document.querySelectorAll('.page').forEach((p, j) => { p.style.display = j === k ? '' : 'none'; }), i);
    const shotBuf = await els[i].screenshot(), shot = shotBuf.toString('base64'), jpg = pages[i].image.jpeg.toString('base64');
    if (process.env.DUMP) { fs.writeFileSync(path.join(process.env.DUMP, `shot${i + 1}.png`), shotBuf); fs.writeFileSync(path.join(process.env.DUMP, `pdf${i + 1}.jpg`), pages[i].image.jpeg); }
    const d = await p2.evaluate(async ({ a, b }) => {
      const load = (src) => new Promise((ok) => { const im = new Image(); im.onload = () => ok(im); im.src = src; });
      const [ia, ib] = await Promise.all([load('data:image/png;base64,' + a), load('data:image/jpeg;base64,' + b)]);
      // 1/10 に縮めて（JPEG の文字の輪郭のにじみを消す）、明るさの差が 24 を超える点の割合。
      // 書体を埋め込まずにパソコンの書体で描くと 2〜11%、正しく描けていれば 0.4% 以下
      const W = 198, H = 280, px = (im) => {                // 10×10 の点ごとの明るさの平均
        const c = document.createElement('canvas'); c.width = W * 10; c.height = H * 10;
        const g = c.getContext('2d'); g.drawImage(im, 0, 0);
        const d = g.getImageData(0, 0, W * 10, H * 10).data, out = new Float32Array(W * H);
        for (let y = 0; y < H * 10; y++) for (let x = 0; x < W * 10; x++) {
          const k = (y * W * 10 + x) * 4;
          out[Math.floor(y / 10) * W + Math.floor(x / 10)] += (0.299 * d[k] + 0.587 * d[k + 1] + 0.114 * d[k + 2]) / 100;
        }
        return out;
      };
      const x = px(ia), y = px(ib);
      let big = 0;
      for (let k = 0; k < x.length; k++) if (Math.abs(x[k] - y[k]) > 24) big++;
      return { big: big / (W * H) };
    }, { a: shot, b: jpg });
    // 書体を渡さないとき（パソコンの書体）：この検査の環境には冊子の明朝体が無く、総称の serif になる。総称の書体は SVG の画像の中では
    // 別の書体に当たることがあり、文字が1px足らずずれるので、大きく壊れていない（白いページ・写真が無いなど）ことだけを見る
    const limit = FONT_DIR ? 0.01 : 0.1;
    ck(d.big < limit, `${i + 1}ページめの見た目が、ふつうに開いた冊子と違う: 違う点 ${(d.big * 100).toFixed(2)}%`);
    if (process.env.V) console.log(`  ${i + 1}ページめ: 違う点 ${(d.big * 100).toFixed(3)}%`);
  }
  if (process.env.V) console.log('  PDF: ' + R.pdf.bytes + ' bytes, ' + Math.round(R.ms) + 'ms');

  // 書体のファイルの1つが読めない（404）・届かない（止まる）とき（書体を渡したときだけ）：
  // 冊子はその字をパソコンの書体で収めるので、PDFでもその字は埋め込まずに同じ書体で描く（入れると字の幅が変わる）。止まっても待ち続けない
  if (FONT_DIR) {
    const css = fs.readFileSync(path.join(FONT_DIR, 'css.txt'), 'utf8');
    const cp = '見'.codePointAt(0), rule = (css.match(/@font-face\s*\{[^}]*\}/g) || []).find((b) => {
      const ur = (b.match(/unicode-range:\s*([^;]+);/) || [])[1] || '';
      return ur.split(',').some((t) => { const m = t.trim().replace(/^U\+/i, '').split('-').map((x) => parseInt(x, 16)); return cp >= m[0] && cp <= (m[1] || m[0]); });
    });
    const bad = rule.match(/url\(\s*['"]?([^'")]+)/)[1];
    for (const mode of ['404', '止まる']) {
      const c3 = await browser.newContext({ deviceScaleFactor: 1 });
      await routeFonts(c3);
      await c3.route((u) => u.href === bad, (r) => { if (mode === '404') r.fulfill({ status: 404, body: '' }); /* 止まる：返さない */ });
      const p3 = await c3.newPage();
      await p3.route('https://mb.test/', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: '<!DOCTYPE html><html><body></body></html>' }));
      await p3.goto('https://mb.test/');
      const t0 = Date.now();
      const give = new Promise((ok) => setTimeout(() => ok({ ok: false, message: '60秒たっても終わらない' }), 60000));
      const F3 = await Promise.race([give, p3.evaluate(async ({ render, pdfjs, members, cats, cover }) => {
        window.members = members; window.cats = cats; window.cover = cover; window.photosB64 = {};
        window.catOf = (k) => cats.find((c) => c.key === k) || null;
        (0, eval)(render); (0, eval)(pdfjs);
        try { const r = await mbpBuildPdf(bookHtml({ photos: {} }), cover.title, () => {}); return { ok: true, base64: r.base64 }; }
        catch (e) { return { ok: false, message: String(e && e.message || e) }; }
      }, { render: strip('memberbook_render.html'), pdfjs: strip('memberbook_pdf.html'), members: MEMBERS.slice(0, 18), cats: CATS, cover: COVER })]);
      const sec = (Date.now() - t0) / 1000;
      ck(F3.ok && sec < 40, `書体のファイルが${mode}とき: PDFを作れない・時間がかかりすぎる（${sec.toFixed(1)}秒）: ` + (F3.message || ''));
      if (F3.ok && mode === '404') {
        // 冊子（同じく書体の1つが読めない）と見比べる
        const pg = pagesOf(readPdf(Buffer.from(F3.base64, 'base64')))[1];
        const c4 = await browser.newContext({ deviceScaleFactor: 2.5 });
        await routeFonts(c4);
        await c4.route((u) => u.href === bad, (r) => r.fulfill({ status: 404, body: '' }));
        const p4 = await c4.newPage();
        const html = await p3.evaluate(() => bookHtml({ photos: {} }));
        await p4.route('https://mb.test/book', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html }));
        await p4.goto('https://mb.test/book');
        await p4.waitForFunction(() => document.documentElement.getAttribute('data-fitted') === '1', null, { timeout: 30000 });
        await p4.addStyleTag({ content: '.page{margin:0!important;box-shadow:none!important;} body{padding:0!important;background:#fff!important;}' });
        await p4.evaluate(() => document.querySelectorAll('.page').forEach((p, j) => { p.style.display = j === 1 ? '' : 'none'; }));
        const shot = (await (await p4.$$('.page'))[1].screenshot()).toString('base64');
        const d = await p4.evaluate(async ({ a, b }) => {
          const load = (src) => new Promise((ok) => { const im = new Image(); im.onload = () => ok(im); im.src = src; });
          const [ia, ib] = await Promise.all([load('data:image/png;base64,' + a), load('data:image/jpeg;base64,' + b)]);
          const W = 198, H = 280, px = (im) => { const c = document.createElement('canvas'); c.width = W * 10; c.height = H * 10; const g = c.getContext('2d'); g.drawImage(im, 0, 0);
            const d = g.getImageData(0, 0, W * 10, H * 10).data, out = new Float32Array(W * H);
            for (let y = 0; y < H * 10; y++) for (let x = 0; x < W * 10; x++) { const k = (y * W * 10 + x) * 4; out[Math.floor(y / 10) * W + Math.floor(x / 10)] += (0.299 * d[k] + 0.587 * d[k + 1] + 0.114 * d[k + 2]) / 100; }
            return out; };
          const x = px(ia), y = px(ib); let big = 0;
          for (let k = 0; k < x.length; k++) if (Math.abs(x[k] - y[k]) > 24) big++;
          return big / (W * H);
        }, { a: shot, b: pg.image.jpeg.toString('base64') });
        ck(d < 0.02, `書体のファイルが読めないとき: 冊子と見た目が違う（${(d * 100).toFixed(2)}%）`);
        if (process.env.V) console.log(`  書体の1つが読めないとき: 違う点 ${(d * 100).toFixed(3)}%`);
        await c4.close();
      }
      if (process.env.V) console.log(`  書体のファイルが${mode}とき: ${sec.toFixed(1)}秒`);
      await c3.close();
    }
  }
  await browser.close();

  console.log(`メンバーブックのPDF（画面の中で作る・ドライブの差し替え）: 検査 ${checks} 件` + (FONT_DIR ? '（書体を埋め込む）' : '（パソコンの書体）'));
  if (fails.length) {
    console.log(`NG: ${fails.length} 件`);
    fails.forEach((f) => console.log('   ' + f));
    process.exit(1);
  }
  console.log('OK: ドライブのメンバーブックは中身だけ差し替え（URLそのまま・差し替えられないときは新しく作らない）・PDFの作り（A4・画像・検索できる文字・枠で切った文字は入れない）・'
    + '見た目は印刷と同じ・写真を縮める・大きすぎるときは小さくして作り直す');
})().catch((e) => { console.error(e); process.exit(1); });
