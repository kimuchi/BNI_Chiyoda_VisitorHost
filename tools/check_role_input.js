// 「役職ごとの入力」のサーバー側（role_input_srv.js）を、ルーティンチェックシートの実物で確かめる。
// 本番と同じファイルを、スプレッドシートまわりだけ差し替えてNodeで動かす。
// シートは書き込めるようにしてあるので、保存（行の追加・ほかの人との衝突・数式の保護）まで通す。
//
//   node tools/check_role_input.js <routine.json> <members.json>
//
//   routine.json … tools/routine_dump.py で書き出したルーティンチェックシート
//   members.json … メンバー名簿（{ no, name, cat, … } の配列）
//
// 参加者シート（YYYYMMDD参加者）は、架空のビジター・ゲストと、名簿の方の代理で作る。

const vm = require('vm');
const { makeRoleServer } = require('./lib_role_fixture');

const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

const { F, sandbox, sheets, members, SUB_FOR } = makeRoleServer(process.argv[2], process.argv[3]);
const roleOf = (ctx, key) => ctx.roles.find((r) => r.key === key);
const itemsOf = (ctx, key) => roleOf(ctx, key).items.map((id) => ctx.items[id]);
// 項目名は完全一致で探し、無ければ先頭一致（「遅刻・欠席担当(7:00開始)」）
const find = (ctx, key, title, parent) => {
  const ok = (it) => parent === undefined || it.parent === parent;
  const list = itemsOf(ctx, key), n = (x) => String(x).replace(/\s/g, '');
  return list.find((it) => it.title === title && ok(it)) || list.find((it) => n(it.title).indexOf(n(title)) === 0 && ok(it));
};
const short = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').slice(0, 40);

// ===================== 9/30（23期）=====================
let ctx = F.getRoleInputContext('', '*');
ck(ctx.ok, '読み込みに失敗: ' + ctx.message);
ck(ctx.date === '2026/09/30', '既定の開催日が ' + ctx.date + '（9/30 のはず）');
ck(ctx.sheetName === '【23期】ルーティンチェックシート', 'シートが ' + ctx.sheetName);
ck(ctx.meetingNo === '535', '開催回が ' + ctx.meetingNo);
ck(ctx.meetings.length >= 2 && ctx.meetings[0].dateValue === '2026/09/30' && /第535回/.test(ctx.meetings[0].display),
   '開催日の選択肢: ' + JSON.stringify(ctx.meetings.slice(0, 2)));
ck(ctx.roles.length === 13 && ctx.roles[0].key === 'president' && ctx.roles[12].key === 'gbc', '役職の数・並び');
ck(roleOf(ctx, 'president').holder === '熊谷 龍威', 'プレジデントの担当者の初期値');

console.log(`${ctx.display} 第${ctx.meetingNo}回（${ctx.sheetName}）`);
for (const r of ctx.roles) {
  const st = r.status;
  console.log(`  ${r.label}（${r.holder}） ${st.state} 必須${st.required}・入力済み${st.filled}`
    + (st.missing.length ? ` 足りない: ${st.missing.map((m) => m.title).slice(0, 5).join('、')}${st.missing.length > 5 ? ' …' : ''}` : ''));
}

// 役職ごとの項目（シートの「担当」の列から）
const titles = (key) => itemsOf(ctx, key).map((it) => (it.parent ? it.parent + '>' : '') + it.title);
const has = (key, list) => list.every((t) => titles(key).indexOf(t) >= 0);
ck(has('president', ['リージョン参加者', '体験談', 'ウィークリープレゼン', '2分30秒プレゼン', 'プレジデントより',
                     '各担当者とポジティブな挨拶', 'BNI目的と概要', '一般規定'].slice(0, 7)),
   'プレジデントの項目: ' + titles('president').join(' / '));
ck(has('vice', ['メンバーシップから報告', '遅刻・欠席担当(7:00開始)', '代理・欠席>代理', '代理・欠席>欠席', '代理・欠席>医療欠席',
                '審査中カテゴリー', '新入会', '一般規定', '真正度確認 ※前々回の外部リファから選ぶ',
                'お願い事項（バイスプレジデントから）', '人数：ビジター', '今週の共有事項（バイスプレジデント）']),
   'バイスプレジデントの項目: ' + titles('vice').join(' / '));
ck(has('secretary', ['スピーカーローテーション用スライド画像', 'メインプレゼン', 'メインプレゼン>スライド', '推薦の言葉',
                     '推薦の言葉>スライド', '書記兼会計より', '直近のイベント', '卒業コメント']),
   '書記兼会計の項目: ' + titles('secretary').join(' / '));
ck(has('ec', ['エデュケーション', 'エデュケーション>スライド', '今週の共有事項（エデュケーションコーディネーター）']),
   'ECの項目: ' + titles('ec').join(' / '));
ck(has('vhc', ['割振表', 'ビジター情報', 'ビジターの紹介']), 'ビジターホストの項目: ' + titles('vhc').join(' / '));
ck(has('web', ['本日の招待者']), 'webマスターの項目: ' + titles('web').join(' / '));
for (const k of ['mentor', 'support', 'training', 'event', 'bcp', 'spreading', 'gbc']) {
  const ts = titles(k);
  ck(ts.length === 1 && /^今週の共有事項（/.test(ts[0]), k + ' の項目: ' + ts.join(' / '));
}
ck(ctx.unknown.length === 0, '担当が分からない項目: ' + JSON.stringify(ctx.unknown));

// いまの値・前回の値・必須・期日
const mp = find(ctx, 'secretary', 'メインプレゼン');
ck(mp && /山本/.test(mp.value) && /金子/.test(mp.value), 'メインプレゼンのいまの値: ' + (mp && mp.value));
ck(mp && mp.prev && mp.prev.date === '2026/09/23' && /瞳/.test(mp.prev.value), 'メインプレゼンの前回: ' + JSON.stringify(mp && mp.prev));
const reg = find(ctx, 'president', 'リージョン参加者');
ck(reg && reg.required && reg.dueDate === '2026/09/27' && reg.dueLabel === '9/27(日)', 'リージョン参加者の期日: ' + JSON.stringify(reg && [reg.required, reg.dueDate, reg.dueLabel]));
const talk = itemsOf(ctx, 'president').find((it) => it.title === 'トークスクリプト');
ck(talk && talk.readOnly && !talk.required, 'トークスクリプト（自動生成）が入力欄になっている');
const sub = find(ctx, 'vice', '代理', '代理・欠席');
ck(sub && !sub.required, '代理（期日なし）が必須になっている');
ck(find(ctx, 'vice', '人数：ビジター').virtual && find(ctx, 'vice', '人数：ビジター').required, '人数：ビジター（足す行）');

// 初期値の推定
const est = (key, title, parent) => { const it = find(ctx, key, title, parent); return it && it.estimate ? it.estimate : null; };
const show = (key, title, parent) => { const e = est(key, title, parent); return e ? `${short(e.value)}〈${e.source}〉` : '(推定なし)'; };
console.log('\n初期値の推定（9/30）:');
[['vice', '一般規定'], ['president', 'ウィークリープレゼン'], ['vice', '遅刻・欠席担当'], ['vice', 'リファーラルの注意'],
 ['vice', '名札・バッチの注意'], ['vice', '医療欠席', '代理・欠席'], ['vice', '代理', '代理・欠席'], ['vice', '人数：ビジター'],
 ['vice', '人数：ゲスト'], ['vice', '人数：見学'], ['vice', '更新対象者 30日前 ※翌月更新予定者'], ['vice', '新入会'],
 ['president', '各担当者とポジティブな挨拶'], ['president', '体験談'], ['vice', '審査中カテゴリー'],
 ['vice', 'バイスプレジデントによる報告'], ['secretary', 'メインプレゼン>スライド'.split('>')[1], 'メインプレゼン']]
  .forEach(([k, t, p]) => console.log(`  ${(p ? p + '>' : '') + t}: ${show(k, t, p)}`));

ck((est('vice', '一般規定') || {}).value === '3番', '一般規定の推定: ' + show('vice', '一般規定'));
ck(/^プロモーション　\d+番　.+さん$/.test((est('president', 'ウィークリープレゼン') || {}).value || ''),
   'ウィークリープレゼンの推定（建築・住まいの次）: ' + show('president', 'ウィークリープレゼン'));
ck((est('vice', '遅刻・欠席担当') || {}).value === '藤田さん', '遅刻・欠席担当: ' + show('vice', '遅刻・欠席担当'));
ck((est('vice', 'リファーラルの注意') || {}).value === '全員', 'リファーラルの注意: ' + show('vice', 'リファーラルの注意'));
ck((est('vice', '名札・バッチの注意') || {}).value === '福王さん', '名札・バッチの注意: ' + show('vice', '名札・バッチの注意'));
ck((est('vice', '代理', '代理・欠席') || {}).value === SUB_FOR.name.split(' ')[0] + 'さん', '代理（参加者シートから）: ' + show('vice', '代理', '代理・欠席'));
ck((est('vice', '人数：ビジター') || {}).value === '2', 'ビジター人数（キャンセルを除く）: ' + show('vice', '人数：ビジター'));
ck((est('vice', '人数：ゲスト') || {}).value === '1', 'ゲスト人数: ' + show('vice', '人数：ゲスト'));
ck((est('vice', '人数：見学') || {}).value === '0', '見学人数: ' + show('vice', '人数：見学'));
ck((est('vice', '更新対象者 30日前 ※翌月更新予定者') || {}).value === members[10].name.split(' ')[0] + 'さん'
   || /さん/.test((est('vice', '更新対象者 30日前 ※翌月更新予定者') || {}).value || ''),
   '更新対象者30日: ' + show('vice', '更新対象者 30日前 ※翌月更新予定者'));
ck(/さん$/.test((est('vice', '新入会') || {}).value || ''), '新入会（入会日から）: ' + show('vice', '新入会'));
ck(/山本さん/.test((est('president', '各担当者とポジティブな挨拶') || {}).value || ''), 'ポジティブな挨拶: ' + show('president', '各担当者とポジティブな挨拶'));
ck((est('president', '体験談') || {}).value === 'なし', '体験談（前回なし）: ' + show('president', '体験談'));
ck(est('secretary', 'スライド', 'メインプレゼン') === null, '「済」の欄を引き継いでいる: ' + show('secretary', 'スライド', 'メインプレゼン'));
ck(find(ctx, 'secretary', 'スライド', 'メインプレゼン').kind === 'check', '「済」の欄が check になっていない');
ck((est('vice', '審査中カテゴリー') || {}).value && /システムプロデューサー/.test(est('vice', '審査中カテゴリー').value), '審査中カテゴリー（前回と同じ）: ' + show('vice', '審査中カテゴリー'));
ck(/チャプター設立以来/.test((est('vice', 'バイスプレジデントによる報告') || {}).value || ''), 'バイスプレジデントによる報告（前回と同じ）: ' + show('vice', 'バイスプレジデントによる報告'));
ck(find(ctx, 'vice', 'バイスプレジデントによる報告').required, '「※自動生成の為、全文を記載」の欄が入力できない扱い');

// 入力状況
const vst = roleOf(ctx, 'vice').status, sst = roleOf(ctx, 'secretary').status, ost = roleOf(ctx, 'mentor').status;
ck(vst.state === 'todo' && vst.required > 10 && vst.missing.length > 0, 'バイスの状況: ' + JSON.stringify(vst).slice(0, 200));
ck(sst.missing.some((m) => m.title === 'スライド' && m.parent === '推薦の言葉'), '書記兼会計の足りない項目に推薦の言葉のスライドが無い');
ck(!sst.missing.some((m) => m.title === 'メインプレゼン'), 'メインプレゼン（入力済み）が足りない扱い');
ck(ost.required === 1 && ost.missing.length === 1 && ost.state === 'todo', 'メンターコーディネーターの状況: ' + JSON.stringify(ost));

// ===================== 保存 =====================
const s23 = sheets.find((s) => s.getName() === '【23期】ルーティンチェックシート');
const col930 = s23._grid[0].indexOf('2026/09/30') + 1;
const rowOfTitle = (sh, title) => {
  const n = (x) => String(x || '').replace(/\s/g, ''), t = n(title);
  const i = sh._grid.findIndex((r) => [r[2], r[3], r[4]].some((v) => n(v) === t));
  return i >= 0 ? i : sh._grid.findIndex((r) => [r[2], r[3], r[4]].some((v) => n(v) && n(v).indexOf(t) === 0));
};
const entry = (c, key, title, value, parent) => { const it = find(c, key, title, parent); return { id: it.id, value, orig: it.orig }; };

// バイスプレジデント：一般規定・人数（足す行）・代理
let res = F.saveRoleInput('2026/09/30', 'vice', [
  entry(ctx, 'vice', '一般規定', '3番'),
  entry(ctx, 'vice', '人数：ビジター', '2'),
  entry(ctx, 'vice', '人数：ゲスト', '1'),
  entry(ctx, 'vice', '代理', '船木さん', '代理・欠席'),
  entry(ctx, 'vice', '今週の共有事項（バイスプレジデント）', '・チャプターサイズ目標53名\n・引継式の予算承認'),
  entry(ctx, 'secretary', 'メインプレゼン', '書き換え'),            // ほかの役職の項目は保存しない
]);
ck(res.ok, '保存に失敗: ' + res.message);
ck(res.saved === 5, '保存した数: ' + res.saved + ' / ' + res.message);
ck(res.added === 1 + 22, '足した行: ' + res.added);
ck(res.skipped.length === 1 && /この役職の項目ではありません/.test(res.skipped[0].why), 'ほかの役職の項目: ' + JSON.stringify(res.skipped));
console.log('\n保存（バイス）: ' + res.message.replace(/\n/g, ' ／ '));
const rPolicy = rowOfTitle(s23, '一般規定');
ck(String(s23._grid[rPolicy][col930 - 1]) === '3番', '一般規定がシートに入っていない: ' + s23._grid[rPolicy][col930 - 1]);
const rHead = s23._grid.findIndex((r) => String(r[2] || '') === '事前MTG 共有事項');
const rLastOld = 52;                                         // 元の最後の項目（真正度確認）の行
ck(rHead === rLastOld + 2, '「事前MTG 共有事項」の見出しの行: ' + rHead + '（' + (rLastOld + 2) + ' のはず）');
ck(String(s23._grid[rHead + 1][3]) === 'お願い事項（プレジデントから）' && String(s23._grid[rHead + 1][5]) === 'プレジ',
   '足した1行目: ' + JSON.stringify(s23._grid[rHead + 1].slice(0, 9)));
const rVis = rowOfTitle(s23, '人数：ビジター');
ck(rVis > rHead && String(s23._grid[rVis][5]) === 'バイス' && String(s23._grid[rVis][6]) === '2日前'
   && String(s23._grid[rVis][col930 - 1]) === '2', '人数：ビジターの行: ' + JSON.stringify(s23._grid[rVis].slice(0, 9)) + ' 値=' + s23._grid[rVis][col930 - 1]);
const rShare = rowOfTitle(s23, '今週の共有事項（バイスプレジデント）');
ck(/チャプターサイズ/.test(String(s23._grid[rShare][col930 - 1])), '共有事項がシートに入っていない');
ck(!s23._writes.some((w) => w.col === col930 && w.row - 1 === rowOfTitle(s23, 'メインプレゼン')), 'ほかの役職の項目を書き換えた');
// 足した行のあとも、ほかの機能の読み取りは変わらない
F.ROUTINE_INDEX_ = null;
vm.runInContext('ROUTINE_INDEX_ = null; ROUTINE_ROWS_CACHE_ = {};', sandbox);
const ri = F.getRoutineInfo('2026/09/30');
ck(ri.found && (ri.mainPresenters || []).length === 2 && ri.generalPolicy === 3, 'ルーティンの読み取り（行を足したあと）: '
   + JSON.stringify({ mp: (ri.mainPresenters || []).map((m) => m.name), gp: ri.generalPolicy }));

// 保存後の状況：足した行の項目も、入力済みとして数える
ctx = res.context;
ck(!roleOf(ctx, 'vice').status.missing.some((m) => /人数：ビジター|共有事項/.test(m.title)), '保存した項目がまだ足りない扱い');
ck(!find(ctx, 'vice', '人数：ビジター').virtual, '足した行が virtual のまま');

// 2回目の保存では行を足さない
res = F.saveRoleInput('2026/09/30', 'vice', [entry(ctx, 'vice', '人数：見学', '0')]);
ck(res.ok && res.added === 0 && res.saved === 1, '2回目の保存: ' + res.message);
ctx = res.context;

// ほかの人がシートを書き換えていたら、上書きしない
const beforeCtx = ctx;
const rLate = rowOfTitle(s23, '遅刻・欠席担当');
s23._grid[rLate][col930 - 1] = '別の方';
res = F.saveRoleInput('2026/09/30', 'vice', [entry(beforeCtx, 'vice', '遅刻・欠席担当', '藤田さん')]);
ck(res.ok && res.saved === 0 && res.skipped.length === 1 && /書き換えられていました/.test(res.skipped[0].why),
   '衝突: ' + res.message);
ck(s23._grid[rLate][col930 - 1] === '別の方', '衝突したのに上書きした');

// 数式の入った欄は読み取りだけ・保存しない
const rTalk = s23._grid.findIndex((r) => String(r[2] || '') === 'トークスクリプト');
s23._formulas[rTalk + ',' + (col930 - 1)] = '=CONCAT(A1)';
const rPos = rowOfTitle(s23, 'その他のお知らせ');
s23._formulas[rPos + ',' + (col930 - 1)] = '=A1';
ctx = F.getRoleInputContext('2026/09/30', '*');
const pos = find(ctx, 'president', 'その他のお知らせ');
ck(pos.readOnly && !pos.required && /数式/.test(pos.readOnlyWhy), '数式の欄: ' + JSON.stringify([pos.readOnly, pos.required, pos.readOnlyWhy]));
res = F.saveRoleInput('2026/09/30', 'president', [{ id: pos.id, value: '上書き', orig: pos.orig }]);
ck(res.saved === 0 && /数式/.test(res.message), '数式の欄に書いた: ' + res.message);

// 日付に見える値は文字のまま（「10/7」）
res = F.saveRoleInput('2026/09/30', 'president', [entry(ctx, 'president', '体験談', '10/7')]);
ck(res.saved === 1 && s23._grid[rowOfTitle(s23, '体験談')][col930 - 1] === '10/7', '「10/7」の保存');

// ===================== 10/7（24期）：前回は23期のシートの 9/30 =====================
ctx = F.getRoleInputContext('2026/10/07', '*');
ck(ctx.ok && ctx.sheetName === '【24期】ルーティンチェックシート' && ctx.meetingNo === '536', '10/7 のシート: ' + ctx.sheetName + ' 第' + ctx.meetingNo + '回');
ck((est('vice', '一般規定') || {}).value === '4番', '10/7 の一般規定（9/30 の 3番 の次）: ' + show('vice', '一般規定'));
const share = find(ctx, 'vice', '今週の共有事項（バイスプレジデント）');
ck(share.virtual && share.prev && share.prev.date === '2026/09/30' && (share.estimate || {}).value === share.prev.value,
   '10/7 の共有事項（前回の内容を引き継ぐ）: ' + JSON.stringify({ v: share.virtual, prev: share.prev, est: share.estimate }).slice(0, 200));
ck((est('vice', '人数：見学') || {}).value === '0', '10/7 の見学人数: ' + show('vice', '人数：見学'));
ck(/^暮らし・生活/.test((est('president', 'ウィークリープレゼン') || {}).value || ''), '10/7 のウィークリープレゼン: ' + show('president', 'ウィークリープレゼン'));
// スピーカーローテーション用スライド画像：表はシステムで作るので「済」
ck((est('secretary', 'スピーカーローテーション用スライド画像') || {}).value === '済', 'ローテーションの画像: ' + show('secretary', 'スピーカーローテーション用スライド画像'));
// メインプレゼン（書記兼会計）は、スピーカーローテーションの2名（10/7 は渡邉さん・岡本さん）
ck((est('secretary', 'メインプレゼン') || {}).value === '①渡邉さん　②岡本さん', '10/7 のメインプレゼン: ' + show('secretary', 'メインプレゼン'));
ck((est('president', '各担当者とポジティブな挨拶') || {}).value === '渡邉さん　岡本さん', '10/7 のポジティブな挨拶（ローテーションから）: ' + show('president', '各担当者とポジティブな挨拶'));
console.log('\n10/7（24期）: 一般規定 ' + show('vice', '一般規定') + ' ／ ウィークリープレゼン ' + show('president', 'ウィークリープレゼン')
  + ' ／ メインプレゼン ' + show('secretary', 'メインプレゼン'));
res = F.saveRoleInput('2026/10/07', 'mentor', [entry(ctx, 'mentor', '今週の共有事項（メンターコーディネーター）', 'パスポートプログラム進行中')]);
const s24 = sheets.find((s) => s.getName() === '【24期】ルーティンチェックシート');
ck(res.ok && res.added === 23 && s24._grid.some((r) => String(r[2] || '') === '事前MTG 共有事項'), '24期のシートに行を足していない: ' + res.message);
ck(roleOf(res.context, 'mentor').status.state === 'done', 'メンターコーディネーターが入力済みにならない');

// ===================== 担当者・担当が分からない項目 =====================
res = F.saveRoleHolders({ president: '見本 太郎', nobody: 'x' });
ck(res.ok && F.roleHolders_().president === '見本 太郎' && F.roleHolders_().vice === '船木 雄大', '担当者の保存');
const rUnk = rowOfTitle(s24, '割振表');
s24._grid[rUnk][5] = 'VH・広報';
ctx = F.getRoleInputContext('2026/10/07', '*');
ck(ctx.unknown.length === 1 && ctx.unknown[0].roleRaw === 'VH・広報', '担当が分からない項目: ' + JSON.stringify(ctx.unknown));
ck(itemsOf(ctx, 'vhc').some((it) => it.title === '割振表'), '「VH・広報」の VH の方にも入っていない');

// 列の無い開催日
ctx = F.getRoleInputContext('2030/01/02');
ck(ctx.ok && !ctx.found && /列が見つかりません/.test(ctx.message), '列の無い開催日: ' + JSON.stringify(ctx).slice(0, 160));

console.log(`\n役職ごとの入力: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.slice(0, 40).forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 項目（担当の列から）・推定・入力状況・保存（行の追加・衝突・数式・役職の制限）');
