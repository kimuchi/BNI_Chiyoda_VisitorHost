// トークスクリプト（talk_script_srv.js・talk_script.html）を、ルーティンチェックシートの写しで確かめる。
//
//   node tools/check_talk_script.js <routine.json> <members.json>
//
// 確かめること
//   ・開催日の候補（次の4回と、前の2回）・既定のひな形・差し込みの一覧（チェックシートの項目名も）
//   ・差し込みの中身：チャプター・回・役職の担当者とカテゴリー・チーム・参加者（ビジター・ゲスト・代理）・
//     チェックシートの項目・メインプレゼン・推薦のことば・真正度確認・ウィークリープレゼンの起点・ローテーション
//   ・「〇〇さん」＋「さん」は1つに、「なし」＋「さん」は敬称を付けない、入らなかったものは「〇〇」、
//     書き方の違う差し込みはそのまま残して知らせる
//   ・画面で直した値（この回だけ）・シートへの書き出し（見出し・作り直し・数式にしない）
//   ・ひな形の保存と読み込み（時刻になった「時間」・空の行）・既定に戻す
//   ・既定のひな形の差し込みは、どれも使える名前で、チャプター名は {チャプター}。docs/samples の見本と同じ
//   ・画面：台本を作る（見本・直した値でシート）／ひな形の編集（差し込みを入れる・行の足し引き・並べ替え・保存・既定に戻す）

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeRoleServer } = require('./lib_role_fixture');
const { loadPage } = require('./lib_minidom');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

const S = makeRoleServer(process.argv[2], process.argv[3]);
for (const f of ['talk_script_default.js', 'talk_script_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), S.sandbox, { filename: f });
}
const F = S.F;
// シートを足せるようにする（台本とひな形のシート）
let gid = 100;
F.getSS_ = () => ({
  getSheets: () => S.sheets,
  getSheetByName: (n) => S.sheets.find((s) => s.getName() === n) || null,
  insertSheet: (n) => {
    const sh = S.makeSheet(n, []);
    sh.clear = () => { sh._grid.length = 0; sh._fmt.length = 0; };
    const id = gid++; sh.getSheetId = () => id;
    // 書式は呼ばれたことだけ控える（行の色で「入らなかった行」を見分けられるか確かめる）
    sh._fmt = [];
    sh.setColumnWidth = (c, w) => { sh._fmt.push(['width', c, w]); };
    sh.setFrozenRows = (r) => { sh._fmt.push(['frozen', r]); };
    const baseRange = sh.getRange;
    sh.getRange = (r, c, nr, nc) => {
      const rg = baseRange(r, c, nr, nc);
      for (const m of ['setFontWeight', 'setBackground', 'setFontSize', 'setWrap', 'setVerticalAlignment', 'setFontColor']) {
        rg[m] = (v) => { sh._fmt.push([m, r, v]); return rg; };
      }
      return rg;
    };
    S.sheets.push(sh);
    return sh;
  },
  getUrl: () => 'https://docs.google.com/spreadsheets/d/TEST/edit',
});
const sheetOf = (n) => S.sheets.find((s) => s.getName() === n);
const DEF = vm.runInContext('TALK_DEFAULT_ROWS_', S.sandbox);
const valOf = (r, k) => (r.values.find((v) => v.key === k) || {});

// ===== 1. 画面の材料 =====
let ctx = F.getTalkScriptContext();
ck(ctx.ok && ctx.meetings.next === '2026/09/30', '開催日の候補: ' + J(ctx.meetings));
const ds = ctx.meetings.list.map((m) => m.dateValue);
ck(J(ds) === J(['2026/09/16', '2026/09/23', '2026/09/30', '2026/10/07', '2026/10/14', '2026/10/21'])
   && /（済）$/.test(ctx.meetings.list[1].display) && /第535回$/.test(ctx.meetings.list[2].display), '前の2回と次の4回: ' + J(ctx.meetings.list));
ck(ctx.template.length === DEF.length && !ctx.fromSheet && ctx.template[0].talk === DEF[0][3], '既定のひな形（シートが無いとき）: ' + ctx.template.length);
const keys = ctx.catalog.flatMap((g) => g.items.map((i) => i.key));
ck(['チャプター', '回', 'プレジデント', 'プレジデントのカテゴリー', 'ビジターの一覧', 'メインプレゼン1', 'ウィークリープレゼンの起点',
    '推薦のことば', '真正度確認の紹介者', 'スピーカーローテーション', 'ルーティン:体験談', 'ルーティン:トークスクリプト追加事項', 'チーム:ビジターホスト']
   .every((k) => keys.includes(k)), '差し込みの一覧に無いもの: ' + ['ルーティン:体験談', 'チーム:ビジターホスト'].filter((k) => !keys.includes(k)));
ck(!keys.some((k) => /^ルーティン:(No\.?|内容|\d+)$/.test(k)), '差し込みの一覧に項目名でないもの: ' + keys.filter((k) => /^ルーティン:(No|内容|\d)/.test(k)));

// ===== 2. 既定のひな形 =====
const defText = DEF.map((r) => r.join('\t')).join('\n');
ck(!/Active|アクティブ|千代田/.test(defText) && /\{チャプター\}/.test(defText) && /\{リージョン\}/.test(defText), '既定のひな形にチャプター固有の語が残っている');
let pv = F.previewTalkScript('2026/09/23');
ck(pv.ok && pv.unknown.length === 0, '既定のひな形に使えない差し込みがある: ' + J(pv.unknown));
// docs/samples の見本と同じ
{
  const tsv = fs.readFileSync(path.join(ROOT, 'docs', 'samples', 'トークスクリプトのひな形.tsv'), 'utf8');
  const rows = [], row = [], s = tsv; let cell = '', q = false, cur = [];
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (q) { if (c === '"' && s[i + 1] === '"') { cell += '"'; i++; } else if (c === '"') q = false; else cell += c; continue; }
    if (c === '"' && cell === '') { q = true; continue; }
    if (c === '\t') { cur.push(cell); cell = ''; continue; }
    if (c === '\n') { cur.push(cell); rows.push(cur); cur = []; cell = ''; continue; }
    cell += c;
  }
  void row;
  ck(J(rows[0]) === J(['時間', '発言者', 'スライド', 'トーク', 'メモ']) && J(rows.slice(1)) === J(DEF),
     'docs/samples/トークスクリプトのひな形.tsv が既定のひな形と違う（' + (rows.length - 1) + '行）');
}

// ===== 3. 差し込みの中身（9/23 … チェックシートが埋まっている回）=====
const holders = F.roleHolders_(F.parseDate_('2026/09/23'));
const members = F.getMemberMaster().members;
const short = (n) => F.roleShortName_(n, members).replace(/さん$/, '');
const catOf = (n) => (members.find((m) => m.name === n) || {}).title || '';
ck(pv.label === '第534回 9月23日(水)', '見出し: ' + pv.label);
ck(valOf(pv, '回').value === '534' && valOf(pv, 'チャプター').value === 'Activeチャプター' && valOf(pv, 'リージョン').value === 'BNI東京千代田リージョン',
   'チャプター・回: ' + J(['回', 'チャプター', 'リージョン'].map((k) => valOf(pv, k).value)));
ck(valOf(pv, 'プレジデント').value === short(holders.president) && valOf(pv, 'プレジデントのカテゴリー').value === catOf(holders.president)
   && valOf(pv, '書記兼会計').value === short(holders.secretary), '役職: ' + J(['プレジデント', 'プレジデントのカテゴリー', '書記兼会計'].map((k) => valOf(pv, k).value)));
ck(valOf(pv, 'チーム:エデュケーションコーディネーター').value === catOf(holders.ec) + 'の' + short(holders.ec) + 'さん',
   'チーム（役職のチームのリーダーは担当者）: ' + valOf(pv, 'チーム:エデュケーションコーディネーター').value);
ck(valOf(pv, 'メインプレゼン').value === '泉さん・舩川さん' && valOf(pv, 'メインプレゼン1').value === '泉' && valOf(pv, 'メインプレゼン2').value === '舩川',
   'メインプレゼン: ' + J(['メインプレゼン', 'メインプレゼン1', 'メインプレゼン2'].map((k) => valOf(pv, k).value)));
ck(valOf(pv, 'ウィークリープレゼンの起点').value === '22番 熊田', 'ウィークリープレゼンの起点（ブレイクアウトルームも）: ' + valOf(pv, 'ウィークリープレゼンの起点').value);
// スタートアッププレゼン（前の呼び名 {2分30秒プレゼン} のままのひな形でも同じ値）と、プレゼンの秒数（チャプターの設定）
{
  const env = F.talkEnv_(F.parseDate_('2026/09/23'));
  const st = F.talkValue_('スタートアッププレゼン', env), old = F.talkValue_('2分30秒プレゼン', env);
  ck(J(st) === J(old) && st.state === 'ok' && valOf(pv, 'スタートアッププレゼン').value === st.value,
     'スタートアッププレゼン（前の呼び名も）: ' + J([st, old, valOf(pv, 'スタートアッププレゼン')]));
  ck(['ウィークリープレゼンの秒数', 'スタートアッププレゼンの秒数', 'ビジタープレゼンの秒数', 'リファーラル発表の秒数']
       .map((k) => F.talkValue_(k, env).value).join('/') === '30秒/2分30秒/20秒/7秒', 'プレゼンの秒数（既定）: '
     + ['ウィークリープレゼンの秒数', 'スタートアッププレゼンの秒数', 'ビジタープレゼンの秒数', 'リファーラル発表の秒数'].map((k) => F.talkValue_(k, env).value).join('/'));
  const vrow = pv.rows.find((r) => /ビジター様のプレゼンテーションタイム/.test(r.talk));
  ck(vrow && /プレゼンの時間は20秒です/.test(vrow.talk), '台本のビジタープレゼンの秒数: ' + (vrow && vrow.talk.slice(0, 60)));
  const wrow = pv.rows.find((r) => /ウィークリープレゼンテーションのコーナー/.test(r.talk));
  ck(wrow && /の2分30秒のスタートアッププレゼンになります/.test(wrow.talk) && /ビジター様の20秒のプレゼンタイム/.test(wrow.talk),
     '台本のスタートアッププレゼン: ' + (wrow && (wrow.talk.match(/\(そして[^)]*\)/) || [''])[0]));
}
ck(valOf(pv, '推薦のことばの件数').value === '2' && valOf(pv, '推薦のことば').value.split('\n').length === 2
   && /^川邉さんから　泉さんに　推薦の言葉/.test(valOf(pv, '推薦のことば').value) && valOf(pv, '推薦のことばを受けた方').value === '泉さん・舩川さん',
   '推薦のことば: ' + J(valOf(pv, '推薦のことば')));
ck(valOf(pv, '真正度確認の紹介者').value === '深井さん' && valOf(pv, '真正度確認の受け手').value === '羽生さん' && valOf(pv, '真正度確認の紹介先').value === '古賀様',
   '真正度確認: ' + J(['真正度確認の紹介者', '真正度確認の受け手', '真正度確認の紹介先'].map((k) => valOf(pv, k).value)));
ck(valOf(pv, '前々回の開催日').value === '9月9日', '前々回の開催日: ' + valOf(pv, '前々回の開催日').value);
ck(valOf(pv, 'ルーティン:エデュケーション').value === '福元さん' && valOf(pv, 'ルーティン:体験談').state === 'none'
   && valOf(pv, 'ルーティン:ネットワーキングリーダー').state === 'none' && valOf(pv, 'ルーティン:ネットワーキングリーダー').value === 'なし'
   && valOf(pv, 'コアバリュー').value === 'Accountability',
   'チェックシートの項目: ' + J(['ルーティン:エデュケーション', 'ルーティン:体験談', 'ルーティン:ネットワーキングリーダー', 'コアバリュー'].map((k) => valOf(pv, k))));
const rot = valOf(pv, 'スピーカーローテーション').value.split('\n');
ck(rot.length === 4 && /^9\/30\(水\) 第535回　山内さん・金井さん$/.test(rot[0]), 'スピーカーローテーション（次の回から4回）: ' + J(rot));
ck(valOf(pv, 'ビジターの一覧').state === 'missing' && pv.missing.includes('ビジターの一覧'), '参加者シートが無い回: ' + J(valOf(pv, 'ビジターの一覧')));
// 台本の行：「なし」は読み飛ばせるようにメモへ、入らなかったものは「〇〇」
const share = pv.rows.find((r) => /シェアストーリ/.test(r.talk));
ck(share && /※チェックシートでは「なし」: ルーティン:体験談/.test(share.memo), '「なし」をメモに書いていない: ' + J(share));
const newMem = pv.rows.find((r) => /新入会のメンバーがいます/.test(r.talk));
ck(newMem && /ご紹介しましょう！ なしです！/.test(newMem.talk), '「なし」＋「さん」: ' + (newMem && newMem.talk.slice(0, 60)));
const vcall = pv.rows.find((r) => /名のビジター様にご参加/.test(r.talk));
ck(vcall && /本日は〇〇名のビジター様/.test(vcall.talk), '入らなかった差し込みが「〇〇」になっていない: ' + (vcall && vcall.talk.slice(0, 200)));
const edu = pv.rows.find((r) => /ネットワーキング学習コーナーです/.test(r.talk));
ck(edu && /それでは福元さんご準備/.test(edu.talk), 'チェックシートの値（敬称つき）: ' + (edu && edu.talk));

// ===== 4. 参加者（9/30 … 参加者シートがある回）=====
pv = F.previewTalkScript('2026/09/30');
const V = valOf(pv, 'ビジターの一覧').value.split('\n');
ck(valOf(pv, 'ビジターの人数').value === '2' && V.length === 2 && /見本 一郎（みほん いちろう）様$/.test(V[0]) && /さんのご招待　税理士　/.test(V[0])
   && !/見本 三郎/.test(V.join('')), 'ビジター（キャンセルの方は入れない）: ' + J(V));
ck(/さんご招待の見本 花子（みほん はなこ）様$/.test(valOf(pv, 'ゲストの一覧').value), 'ゲスト: ' + valOf(pv, 'ゲストの一覧').value);
ck(valOf(pv, '代理の一覧').value === short(S.SUB_FOR.name) + 'さんの代理として見本 代理（みほん だいり）様', '代理: ' + valOf(pv, '代理の一覧').value);

// ===== 5. 差し込みの入れ方（敬称・入らない・書き方の違い）=====
{
  const R = { A: { value: '山田さん', state: 'ok' }, B: { value: 'なし', state: 'none' }, C: { value: '', state: 'missing' },
              D: { value: '', state: 'unknown' }, E: { value: '見本', state: 'ok' } };
  const used = [];
  const fill = (t) => F.talkFill_(t, (k) => R[k] || R.D, (k, r) => used.push(k + ':' + r.state));
  ck(fill('{A}さん、{E}さん') === '山田さん、見本さん', '敬称を1つに: ' + fill('{A}さん、{E}さん'));
  ck(fill('{B}さんです') === 'なしです', '「なし」に敬称を付けない: ' + fill('{B}さんです'));
  ck(fill('{C}さん') === '〇〇さん' && fill('{ D }') === '{ D }', '入らない・書き方の違い: ' + fill('{C}さん') + ' / ' + fill('{ D }'));
  ck(used.includes('A:ok') && used.includes('C:missing') && used.includes('D:unknown'), '使った差し込みの知らせ: ' + J(used));
}

// ===== 6. 画面で直した値・シートへの書き出し =====
pv = F.previewTalkScript('2026/09/23', { 'プレジデント': '見本' });
ck(valOf(pv, 'プレジデント').value === '見本' && valOf(pv, 'プレジデント').manual && pv.rows.some((r) => /プレジデントを務めます.*の見本です/.test(r.talk)),
   '画面で直した値: ' + J(valOf(pv, 'プレジデント')));
let made = F.createTalkScript('2026/09/23');
let sh = sheetOf('20260923トークスクリプト');
ck(made.ok && sh && made.sheetName === '20260923トークスクリプト' && /#gid=\d+$/.test(made.url), '台本のシート: ' + J(made));
ck(sh && /第534回 9月23日\(水\)　トークスクリプト（Activeチャプター）/.test(sh._grid[0][0]) && J(sh._grid[1].slice(0, 5)) === J(['時間', '発言者', 'スライド', 'トーク', 'メモ'])
   && sh.getLastRow() === 2 + DEF.length, '見出しと行数: ' + (sh && [sh._grid[0][0], sh.getLastRow()]));
ck(sh && !sh._grid.slice(2).some((r) => /\{[^}]+\}/.test(r.join(''))), 'シートに {…} が残っている');
ck(/入らなかった差し込み/.test(made.message) && made.missing.includes('ビジターの一覧'), '入らなかったもののお知らせ: ' + made.message.slice(0, 120));
{
  const marks = F.previewTalkScript('2026/09/23').marks;
  const yellow = sh ? sh._fmt.filter((f) => f[0] === 'setBackground' && f[2] === '#fff4cc').map((f) => f[1] - 3) : [];
  const want = marks.map((m, i) => (m === 'missing' ? i : -1)).filter((i) => i >= 0);
  ck(sh && sh._fmt.filter((f) => f[0] === 'width').length === 5 && sh._fmt.some((f) => f[0] === 'frozen' && f[1] === 2)
     && want.length > 0 && J(yellow) === J(want), '書式（列の幅・見出しの固定・入らなかった行の色）: ' + J(yellow) + ' / ' + J(want));
}
made = F.createTalkScript('2026/09/23');
ck(made.ok && S.sheets.filter((s) => s.getName() === '20260923トークスクリプト').length === 1 && sheetOf('20260923トークスクリプト').getLastRow() === 2 + DEF.length,
   '作り直すと上書き');
ck(vm.runInContext('ARCHIVE_SHEET_PATTERN_', S.sandbox).test('20260923トークスクリプト'), 'シートの整理（アーカイブ）の対象に入っていない');

// ===== 7. ひな形の保存と読み込み =====
const mine = DEF.map((r) => ({ time: r[0], speaker: r[1], slide: r[2], talk: r[3], memo: r[4] }));
mine[0].talk = '【見本】{チャプター}の{回}回目です。{存在しない}';
mine.splice(1, 0, { time: '', speaker: '', slide: '', talk: '', memo: '' });          // 空の行は保存しない
mine.push({ time: '9:00', speaker: 'プレジデント', slide: '', talk: '=1+1', memo: '' });
let sv = F.saveTalkTemplate(mine);
const tsh = sheetOf('トークスクリプト（ひな形）');
ck(sv.ok && sv.count === DEF.length + 1 && tsh && J(tsh._grid[0].slice(0, 5)) === J(['時間', '発言者', 'スライド', 'トーク', 'メモ']), 'ひな形の保存: ' + J(sv));
ck(tsh && tsh._grid[tsh.getLastRow() - 1][3] === "'=1+1", '数式にしない（先頭に \'）: ' + (tsh && tsh._grid[tsh.getLastRow() - 1][3]));
tsh._grid[1][0] = new F.Date(2026, 0, 1, 7, 15);                                   // シートで「時間」が時刻になったとき
ctx = F.getTalkScriptContext();
ck(ctx.fromSheet && ctx.template.length === DEF.length + 1 && ctx.template[0].time === '7:15' && ctx.template[0].talk === mine[0].talk,
   'ひな形の読み込み: ' + J(ctx.template[0]));
made = F.createTalkScript('2026/09/30');
sh = sheetOf('20260930トークスクリプト');
ck(made.ok && sh._grid[2][3] === '【見本】Activeチャプターの535回目です。{存在しない}' && made.unknown.includes('存在しない') && /書き方が違う差し込み/.test(made.message),
   '保存したひな形で作る（書き方の違う差し込みは残す）: ' + (sh && sh._grid[2][3]));
ck(!F.saveTalkTemplate([{ time: '', talk: '' }]).ok, '空のひな形を保存した');
sv = F.resetTalkTemplate();
ck(sv.ok && sv.template.length === DEF.length && F.getTalkScriptContext().template[0].talk === DEF[0][3], '既定に戻す: ' + J(sv).slice(0, 100));

// ===== 8. 画面 =====
const calls = [];
const server = {
  getTalkScriptContext: () => { calls.push(['ctx']); return F.getTalkScriptContext(); },
  previewTalkScript: (d, o) => { calls.push(['preview', d, o]); return F.previewTalkScript(d, o); },
  createTalkScript: (d, o) => { calls.push(['create', d, o]); return F.createTalkScript(d, o); },
  saveTalkTemplate: (r) => { calls.push(['save', r]); return F.saveTalkTemplate(r); },
  resetTalkTemplate: () => { calls.push(['reset']); return F.resetTalkTemplate(); },
};
const page = loadPage('talk_script.html', { server, fails });
const { els, run, step } = page;
step('開く', () => page.window.onload());
ck(els.date.value === '2026/09/30' && els.date.options.length === 6, '開催日の選択: ' + els.date.value);
ck(els.r0_talk && els.r0_talk.value === DEF[0][3] && /\{プレジデント\}/.test(els.pal.innerHTML), 'ひな形の行・差し込みのボタン');
ck(/既定のひな形/.test(els.tplNote.innerHTML) === false && /シート「トークスクリプト（ひな形）」/.test(els.tplNote.innerHTML), 'ひな形の出どころ: ' + els.tplNote.innerHTML);
step('中身を確かめる', () => run('preview()'));
ck(/第535回 9月30日\(水\)/.test(els.sum.innerHTML) && els.v_0 && /できあがりの見本（107行）/.test(els.pv.innerHTML), '見本: ' + els.sum.innerHTML.slice(0, 80));
{
  const mk = F.previewTalkScript('2026/09/30').marks;
  const n = (re) => (els.pv.innerHTML.match(re) || []).length;
  ck(n(/class="row missing"/g) === mk.filter((m) => m === 'missing').length && n(/class="row none"/g) === mk.filter((m) => m === 'none').length,
     '見本の行の色（入らなかった・なし）: ' + n(/class="row missing"/g) + '/' + mk.filter((m) => m === 'missing').length);
}
const iPres = run("S.vals.findIndex(function(v){ return v.key==='プレジデント'; })");
els['v_' + iPres].value = '手入力';
step('シートに作る（直した値）', () => run('create()'));
const last = calls.filter((c) => c[0] === 'create').pop();
ck(last && last[1] === '2026/09/30' && J(last[2]) === J({ 'プレジデント': '手入力' }), '直した値を渡していない: ' + J(last && last[2]));
ck(/20260930トークスクリプト/.test(els.made.innerHTML) && sheetOf('20260930トークスクリプト')._grid.some((r) => /手入力です/.test(r[3])), '直した値でシートを作る');
step('ひな形の編集を開く', () => run("showTab('edit')"));
ck(els.edit.hidden === false && els.make.hidden === true, 'タブの切り替え');
step('差し込みを入れる', () => { run("S.last='r0_talk'"); run('insertKey(0,1)'); });
ck(/\{チャプター名\}$/.test(els.r0_talk.value) && run('S.rows[0].talk') === els.r0_talk.value && run('S.dirty') === true
   && /保存していない変更/.test(els.tplNote.innerHTML), '差し込みを入れる: ' + els.r0_talk.value.slice(-20));
step('行を足す・消す・並べ替える', () => { run('addRow(0)'); run("S.rows[1].talk='足した行'"); run('moveRow(1,-1)'); });
ck(run('S.rows.length') === DEF.length + 1 && run('S.rows[0].talk') === '足した行', '行の足し算・並べ替え: ' + run('S.rows[0].talk'));
step('行を消す', () => run('delRow(0)'));
ck(run('S.rows.length') === DEF.length && page.log.confirms.length >= 1, '行を消す（確かめてから）');
step('保存する', () => run('saveTpl()'));
const saved = calls.filter((c) => c[0] === 'save').pop();
ck(saved && saved[1].length === DEF.length && /\{チャプター名\}$/.test(saved[1][0].talk) && run('S.dirty') === false && /保存しました/.test(els.msg.textContent),
   '保存: ' + els.msg.textContent);
step('既定に戻す', () => run('resetTpl()'));
ck(run('S.rows[0].talk') === DEF[0][3] && /既定の内容に戻しました/.test(els.msg.textContent), '既定に戻す: ' + els.msg.textContent);

console.log(`トークスクリプト: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 候補・一覧・差し込みの中身（役職・チーム・参加者・チェックシート・推薦・真正度・起点・ローテーション）・敬称・直した値・シート・ひな形の保存と読み込み・画面');
