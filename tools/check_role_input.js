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

const { F, sandbox, sheets, members, SUB_FOR, props } = makeRoleServer(process.argv[2], process.argv[3]);
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
ck(roleOf(ctx, 'president').holder === '熊田 龍平', 'プレジデントの担当者の初期値');

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
ck(mp && /山内/.test(mp.value) && /金井/.test(mp.value), 'メインプレゼンのいまの値: ' + (mp && mp.value));
ck(mp && mp.prev && mp.prev.date === '2026/09/23' && /泉/.test(mp.prev.value), 'メインプレゼンの前回: ' + JSON.stringify(mp && mp.prev));
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
// 業種区分マスタに以前の既定の行が残っていても（建築住まい 3・プロモーション 4・美容健康 5 のまま）、1つ進める
{
  const realCats = F.getCategoryMaster;
  const LEGACY = require('vm').runInContext('LEGACY_CATEGORIES_', sandbox);
  const CUR = require('vm').runInContext('DEFAULT_CATEGORIES_', sandbox);
  const lk = new Set(LEGACY.map((r) => r[0]));
  F.getCategoryMaster = () => LEGACY.concat(CUR.filter((r) => !lk.has(r[0])))
    .map((r) => ({ key: r[0], label: r[1], block: r[4], order: r[5] })).sort((a, b) => a.order - b.order);
  const c2 = F.getRoleInputContext('2026/09/30', 'president');
  const w2 = c2.items[c2.order.find((k) => c2.items[k].title === 'ウィークリープレゼン')].estimate || {};
  ck(/^プロモーション　\d+番　.+さん$/.test(w2.value || '') && /前回（9\/23\(水\)）の記載から、業種区分を1つ進めて/.test(w2.source || ''),
     '昔の行が残った業種区分マスタでのウィークリープレゼンの推定: ' + JSON.stringify(w2));
  F.getCategoryMaster = realCats;
}
ck((est('vice', '遅刻・欠席担当') || {}).value === '藤本さん', '遅刻・欠席担当: ' + show('vice', '遅刻・欠席担当'));
ck((est('vice', 'リファーラルの注意') || {}).value === '全員', 'リファーラルの注意: ' + show('vice', 'リファーラルの注意'));
ck((est('vice', '名札・バッチの注意') || {}).value === '福元さん', '名札・バッチの注意: ' + show('vice', '名札・バッチの注意'));
ck((est('vice', '代理', '代理・欠席') || {}).value === SUB_FOR.name.split(' ')[0] + 'さん', '代理（参加者シートから）: ' + show('vice', '代理', '代理・欠席'));
ck((est('vice', '人数：ビジター') || {}).value === '2', 'ビジター人数（キャンセルを除く）: ' + show('vice', '人数：ビジター'));
ck((est('vice', '人数：ゲスト') || {}).value === '1', 'ゲスト人数: ' + show('vice', '人数：ゲスト'));
ck((est('vice', '人数：見学') || {}).value === '0', '見学人数: ' + show('vice', '人数：見学'));
ck((est('vice', '更新対象者 30日前 ※翌月更新予定者') || {}).value === members[10].name.split(' ')[0] + 'さん'
   || /さん/.test((est('vice', '更新対象者 30日前 ※翌月更新予定者') || {}).value || ''),
   '更新対象者30日: ' + show('vice', '更新対象者 30日前 ※翌月更新予定者'));
ck(/さん$/.test((est('vice', '新入会') || {}).value || ''), '新入会（入会日から）: ' + show('vice', '新入会'));
ck(/山内さん/.test((est('president', '各担当者とポジティブな挨拶') || {}).value || ''), 'ポジティブな挨拶: ' + show('president', '各担当者とポジティブな挨拶'));
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
  entry(ctx, 'vice', '代理', '船越さん', '代理・欠席'),
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
res = F.saveRoleInput('2026/09/30', 'vice', [entry(beforeCtx, 'vice', '遅刻・欠席担当', '藤本さん')]);
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
// メインプレゼン（書記兼会計）は、スピーカーローテーションの2名（10/7 は川邉さん・岡林さん）
ck((est('secretary', 'メインプレゼン') || {}).value === '①川邉さん　②岡林さん', '10/7 のメインプレゼン: ' + show('secretary', 'メインプレゼン'));
ck((est('president', '各担当者とポジティブな挨拶') || {}).value === '川邉さん　岡林さん', '10/7 のポジティブな挨拶（ローテーションから）: ' + show('president', '各担当者とポジティブな挨拶'));
console.log('\n10/7（24期）: 一般規定 ' + show('vice', '一般規定') + ' ／ ウィークリープレゼン ' + show('president', 'ウィークリープレゼン')
  + ' ／ メインプレゼン ' + show('secretary', 'メインプレゼン'));
res = F.saveRoleInput('2026/10/07', 'mentor', [entry(ctx, 'mentor', '今週の共有事項（メンターコーディネーター）', 'パスポートプログラム進行中')]);
const s24 = sheets.find((s) => s.getName() === '【24期】ルーティンチェックシート');
ck(res.ok && res.added === 23 && s24._grid.some((r) => String(r[2] || '') === '事前MTG 共有事項'), '24期のシートに行を足していない: ' + res.message);
ck(roleOf(res.context, 'mentor').status.state === 'done', 'メンターコーディネーターが入力済みにならない');

// ===================== 担当者・担当が分からない項目 =====================
res = F.saveRoleHolders({ president: '見本 太郎', nobody: 'x' });
ck(res.ok && F.roleHolders_().president === '見本 太郎' && F.roleHolders_().vice === '船越 雄一', '担当者の保存');

// 担当者は半期ごと（4〜9月・10〜3月。2026年9月までが23期、10月からが24期）
const termOf = (d) => F.roleTermOf_(F.parseDate_(d));
const holderOn = (d, k) => F.roleHolders_(F.parseDate_(d))[k];
ck(termOf('2026/09/30') === 23 && termOf('2026/10/07') === 24 && termOf('2027/03/31') === 24 && termOf('2027/04/07') === 25
   && termOf('2026/04/01') === 23 && termOf('2026/03/25') === 22,
   '期の分け方: ' + ['2026/03/25', '2026/04/01', '2026/09/30', '2026/10/07', '2027/03/31', '2027/04/07'].map(termOf).join(','));
ck(F.roleTermLabel_(23) === '2026年4月〜9月' && F.roleTermLabel_(24) === '2026年10月〜2027年3月', '期の月: ' + F.roleTermLabel_(24));
// いま保存したのは今日（9/26）の期＝23期。24期は、期ごとにする前の担当者（初期値）のまま
ck(holderOn('2026/09/30', 'president') === '見本 太郎' && holderOn('2026/10/07', 'president') === '熊田 龍平',
   '23期と24期の担当者が分かれていない: ' + holderOn('2026/10/07', 'president'));
const c24 = F.getRoleInputContext('2026/10/07', '');
ck(c24.holderTerm.term === 24 && c24.holderTerm.registered && roleOf(c24, 'president').holder === '熊田 龍平'
   && c24.holderTerms.map((t) => t.term).join(',') === '23,24,25' && !c24.holderTerms[2].registered && c24.holderTerms[2].from === 24,
   '10/7 の担当者の期: ' + JSON.stringify(c24.holderTerms.map((t) => [t.term, t.registered, t.from])));
// 次の期（25期）を前もって登録する。24期の回はそのまま
res = F.saveRoleHolders({ president: '次期 花子' }, 25, '2026/10/07');
const t25 = (res.terms || []).find((t) => t.term === 25) || {};
ck(res.ok && res.term === 25 && res.current.term === 24 && res.current.holders.president === '熊田 龍平'
   && t25.registered && t25.holders.president === '次期 花子' && t25.holders.vice === '船越 雄一' && /25期（2027年4月〜9月）/.test(res.message),
   '25期の登録: ' + JSON.stringify(res).slice(0, 240));
ck(holderOn('2027/04/07', 'president') === '次期 花子' && holderOn('2027/03/31', 'president') === '熊田 龍平', '25期の担当者が4月から使われない');
// 登録していない先の期（26期）は前の期（25期）の担当者。画面は「まだ登録されていません」と知らせる
const h26 = F.roleHoldersOfTerm_(F.roleHolderTerms_(), 26);
ck(!h26.registered && h26.from === 25 && h26.holders.president === '次期 花子', '26期（未登録）: ' + JSON.stringify(h26).slice(0, 160));
// 事前MTGの役職のページ・ローテーションの案内文も、その回の期の担当者
ck(F.premtgData_(F.parseDate_('2026/09/30')).roles.find((r) => r.key === 'president').holder === '見本 太郎'
   && F.premtgData_(F.parseDate_('2026/10/07')).roles.find((r) => r.key === 'president').holder === '熊田 龍平', '事前MTGの担当者の期');
res = F.saveRoleHolders({ secretary: '見本 書記' }, 24, '2026/10/07');
let rotNow = F.getSpeakerRotation();
ck(rotNow.weeks[0].secretary === '原口 雅樹' && rotNow.weeks[1].secretary === '見本 書記' && /書記兼会計の原口 雅樹まで/.test(rotNow.fbText),
   'ローテーションの案内文の書記兼会計: ' + rotNow.weeks.slice(0, 2).map((w) => w.date + ' ' + w.secretary).join(' / '));
F.saveRoleHolders({ secretary: '原口 雅樹' }, 24, '2026/10/07');
// 期ごとにする前に保存していた担当者は、24期の担当者として読む
{
  const keep = props.BNI_ROLE_HOLDERS_TERMS;
  delete props.BNI_ROLE_HOLDERS_TERMS;
  props.BNI_ROLE_HOLDERS = JSON.stringify({ president: '前の 保存', vice: '' });
  // 書いていない役職は空欄（担当者の初期値はコードに持たない）
  ck(holderOn('2026/10/07', 'president') === '前の 保存' && holderOn('2026/10/07', 'vice') === '' && holderOn('2026/10/07', 'secretary') === ''
     && holderOn('2026/09/30', 'president') === '前の 保存', '前の形の担当者: ' + JSON.stringify(F.roleHolders_(F.parseDate_('2026/10/07'))).slice(0, 120));
  props.BNI_ROLE_HOLDERS_TERMS = keep;
  delete props.BNI_ROLE_HOLDERS;
}

// ===================== 担当者を、メンバー名簿の「役職」に反映する =====================
{
  const roster = sheets.find((s) => s.getName() === 'メンバー名簿');
  const RC = roster._grid[0].indexOf('役職');
  const roleIn = (n) => { const r = roster._grid.find((row) => row[2] === n); return r ? r[RC] : undefined; };
  const setRole = (n, v) => { roster._grid.find((row) => row[2] === n)[RC] = v; };
  const rosterState = () => JSON.parse(props.BNI_ROLE_ROSTER_STATE || '{}');
  const autoSync = () => { F.ROLE_ROSTER_CHECKED_ = false; return F.roleRosterAutoSync_(); };
  const fromTo = (res, n) => { const c = (res.changes || []).find((x) => x.name === n); return c ? c.from + '→' + c.to : '（変更なし）'; };
  // ここまでの保存で23期を反映しているので、名簿と記録を最初に戻す
  roster._grid.slice(1).forEach((row) => { row[RC] = ''; });
  [['見本 前任', 'バイスプレジデント'], ['見本 兼任', 'BCP委員・スプレディング委員'], ['見本 ウェブ', 'Ｗｅｂマスター'],
   ['見本 経歴', 'ビジホス　過去の経験は、書記兼会計、トレーニング委員'], ['見本 ホスト', 'ビジターホスト'],
   ['熊田 龍平', '見本の旧役職']].forEach((x) => setRole(x[0], x[1]));
  delete props.BNI_ROLE_ROSTER_STATE;

  // 確かめる（名簿はまだ直さない）：24期の担当者どおりにするときの変更
  let p = F.previewRoleHoldersRoster(24);
  ck(p.ok && fromTo(p, '熊田 龍平') === '見本の旧役職→プレジデント' && fromTo(p, '船越 雄一') === '→バイスプレジデント'
     && fromTo(p, '見本 前任') === 'バイスプレジデント→' && fromTo(p, '見本 兼任') === 'BCP委員・スプレディング委員→'
     && fromTo(p, '見本 ウェブ') === 'Ｗｅｂマスター→' && fromTo(p, '見本 経歴') === '（変更なし）' && fromTo(p, '見本 ホスト') === '（変更なし）'
     && p.changes.length === 16 && !p.notFound.length,
     '名簿の役職の変更（24期）: ' + p.changes.map((c) => c.name + ' ' + c.from + '→' + c.to).join(' / '));
  ck(roleIn('熊田 龍平') === '見本の旧役職' && !props.BNI_ROLE_ROSTER_STATE, '確かめただけで名簿を直した');
  // 反映する
  let res = F.applyRoleHoldersToRoster(24);
  ck(res.ok && res.changes.length === 16 && /24期の担当者に合わせて直しました（16名）/.test(res.message)
     && roleIn('熊田 龍平') === 'プレジデント' && roleIn('原口 雅樹') === '書記兼会計' && roleIn('田村 浩美') === 'イベント委員＆1to1促進委員'
     && roleIn('見本 前任') === '' && roleIn('見本 ウェブ') === '' && roleIn('見本 経歴') === 'ビジホス　過去の経験は、書記兼会計、トレーニング委員'
     && roleIn('見本 ホスト') === 'ビジターホスト' && res.state.applied === 24 && res.state.at === '2026/09/26',
     '名簿の役職に反映（24期）: ' + res.message + ' ' + JSON.stringify(res.state));
  res = F.applyRoleHoldersToRoster(24);
  ck(res.ok && !res.changes.length && /24期の担当者どおりでした/.test(res.message), '2回目の反映: ' + res.message);
  // 次の期（24期）を前もって反映してあれば、今の期（23期）に入っても元に戻さない
  ck(autoSync() === null && roleIn('熊田 龍平') === 'プレジデント' && rosterState().seen === 23, '前もって反映した次の期を戻した');
  // 期が替わったら反映する（22期のときに確かめて、22期を反映していたことにする → 23期に入った）
  props.BNI_ROLE_ROSTER_STATE = JSON.stringify({ applied: 22, seen: 22 });
  res = autoSync();
  ck(res && res.term === 23 && fromTo(res, '熊田 龍平') === 'プレジデント→' && roleIn('原口 雅樹') === '書記兼会計'
     && res.notFound.join() === '見本 太郎' && rosterState().applied === 23 && rosterState().seen === 23,
     '期が替わったときの反映（23期）: ' + (res && res.message));
  ck(autoSync() === null, '同じ期のうちに、もう一度反映した');
  // 今の期の担当者がまだ無いときは反映しない（登録して保存したときに反映する）
  const keepTerms = props.BNI_ROLE_HOLDERS_TERMS;
  const only24 = JSON.parse(keepTerms); delete only24['23'];
  props.BNI_ROLE_HOLDERS_TERMS = JSON.stringify(only24);
  props.BNI_ROLE_ROSTER_STATE = JSON.stringify({ applied: 22, seen: 22 });
  ck(autoSync() === null && rosterState().seen === 23 && rosterState().applied === 22, '未登録の期を反映した');
  res = F.saveRoleHolders({ president: '熊田 龍平' }, 23, '2026/09/30');
  ck(res.ok && res.roster && res.roster.term === 23 && roleIn('熊田 龍平') === 'プレジデント' && /メンバー名簿の「役職」も直しました/.test(res.message)
     && res.rosterState.applied === 23, '今の期を登録したときの反映: ' + res.message);
  props.BNI_ROLE_HOLDERS_TERMS = keepTerms;
  // 反映してある期の担当者を直すと、名簿も直す（前の担当者の役職は空欄に）
  F.applyRoleHoldersToRoster(24);
  res = F.saveRoleHolders({ bcp: '丘野 秀人' }, 24, '2026/10/07');
  ck(res.ok && roleIn('丘野 秀人') === 'BCP委員' && roleIn('梅中 公平') === '' && /メンバー名簿の「役職」も直しました（2名）/.test(res.message),
     '反映してある期の担当者を直したとき: ' + res.message);
  // 反映していない期（25期）を直しても、名簿はそのまま
  res = F.saveRoleHolders({ bcp: '丘野 秀人' }, 25, '2026/10/07');
  ck(res.ok && !res.roster && roleIn('梅中 公平') === '', '反映していない期を直したのに名簿を直した');
  F.saveRoleHolders({ bcp: '梅中 公平' }, 24, '2026/10/07');
  ck(roleIn('梅中 公平') === 'BCP委員' && roleIn('丘野 秀人') === '', '担当者を戻したときの名簿: ' + roleIn('丘野 秀人'));
  // 取り込み（Spreadingなど）で役職が前の期のままに戻っても、13の役職は担当者に合わせ直す
  setRole('熊田 龍平', 'ビジターホスト'); setRole('見本 前任', 'バイスプレジデント');
  const note = F.roleRosterAfterImport_();
  ck(roleIn('熊田 龍平') === 'プレジデント' && roleIn('見本 前任') === '' && /24期の内容に合わせ直しました（2名）/.test(note), '取り込みのあと: ' + note);

  // ===================== チーム（半期ごと）=====================
  // まだ登録していない：既定のチーム（メンバーシップ委員会と、役職ごとのチーム）だけで、メンバーは空
  const DEF_TEAMS = 'メンバーシップ委員会・エデュケーションコーディネーター・ビジターホスト・Webチーム・メンターコーディネーター・'
    + 'イベント委員＆1to1促進委員・メンバーサポート委員・トレーニング委員・BCP委員・スプレディング委員会・グローバルビジネスコーディネーター';
  let tt = F.roleTeamsOfTerm_(F.roleTeamTerms_(), 24);
  ck(!tt.registered && tt.from === null && tt.teams.map((t) => t.name).join('・') === DEF_TEAMS
     && tt.teams.every((t) => !t.members.length && !t.leader), '未登録のチーム: ' + tt.teams.map((t) => t.name).join('・'));
  // 24期のチームを登録する（キーの無い前の形も読む。名前・メンバーの空欄や重なりはそろえる。担当者は変えない）
  res = F.saveRoleHolders(null, 24, '2026/10/07', [
    { key: 'role:vhc', name: 'ビジターホスト', members: [{ name: '山内 登志夫', note: 'サブリーダー' }, { name: '丘野 秀人' }, '成田 幸司', '丘野　秀人', ''] },
    { name: ' メンバーシップ委員会 ', members: ['西澤 浩平'] },
    { key: 'role:ec', name: 'エデュケーションコーディネーター', members: [{ name: '原口 雅樹' }] },
    { key: 'role:mentor', name: 'メンターコーディネーター', members: [{ name: '田村 秀二', note: 'メンター' }] },
    { key: 'role:web', name: 'Webチーム', members: [{ name: '文野 雅彦' }, { name: '原口 雅樹' }] },
    { key: 'role:vhc', name: 'ビジターホスト', members: ['見本 重なり'] },
    { name: '', members: ['見本 空欄'] },
    { name: '広報チーム', leader: '北川 光男', members: ['加納 祐介'] }]);
  tt = F.roleTeamsOfTerm_(F.roleTeamTerms_(), 24);
  const teamOf = (k) => tt.teams.find((t) => t.key === k) || { members: [] };
  ck(res.ok && /24期（2026年10月〜2027年3月）の役職・チームを保存しました/.test(res.message) && tt.registered && tt.teams.length === 12
     && JSON.stringify(teamOf('role:vhc').members) === JSON.stringify([{ name: '山内 登志夫', note: 'サブリーダー' }, { name: '丘野 秀人', note: '' }, { name: '成田 幸司', note: '' }])
     && teamOf('membership').members.map((m) => m.name).join() === '西澤 浩平' && teamOf('custom:1').name === '広報チーム'
     && teamOf('custom:1').leader === '北川 光男' && holderOn('2026/10/07', 'president') === '熊田 龍平',
     'チームの保存: ' + res.message + ' ' + JSON.stringify(tt.teams.filter((t) => t.members.length || t.leader)));
  // 名簿の役職：チームを登録した期は、役職の欄を全部その期の内容にする（24期は反映してある期なので、保存したときに直す）
  ck(res.roster && res.roster.full && roleIn('熊田 龍平') === 'プレジデント' && roleIn('丘野 秀人') === 'ビジターホスト'
     && roleIn('山内 登志夫') === 'メンバーサポート委員・ビジターホスト（サブリーダー）' && roleIn('西澤 浩平') === 'メンバーシップ委員会'
     && roleIn('原口 雅樹') === '書記兼会計・エデュケーションコーディネーター（サポート）・Webチーム'
     && roleIn('田村 秀二') === 'メンターコーディネーター（メンター）' && roleIn('北川 光男') === '広報チーム（リーダー）'
     && roleIn('加納 祐介') === '広報チーム' && roleIn('見本 経歴') === '' && roleIn('見本 ホスト') === ''
     && /メンバー名簿の「役職」も直しました/.test(res.message),
     'チームを登録した期の名簿: ' + JSON.stringify(['山内 登志夫', '原口 雅樹', '田村 秀二', '北川 光男', '見本 経歴'].map(roleIn)));
  // メンバーシップ委員会のリーダーに、役職の無い方を選ぶと「（リーダー）」、役職のある方は役職だけ
  ck(F.roleTeamMemberLabel_({ key: 'role:vhc', name: 'ビジターホスト' }, { name: 'x', note: '' }) === 'ビジターホスト'
     && F.roleTeamMemberLabel_({ key: 'role:training', name: 'トレーニング委員' }, { name: 'x', note: '' }) === 'トレーニング委員（サポート）'
     && F.roleTeamMemberLabel_({ key: 'role:training', name: 'トレーニング促進委員' }, { name: 'x', note: '' }) === 'トレーニング促進委員',
     'チームでの役割の書き方');
  // 25期（未登録）は、24期の顔ぶれを引き継がない（メンバー・リーダーは空）。チームの名前と足したチームだけ24期のもの。
  // 画面に渡す期の一覧にも入る
  const t25b = F.roleTeamsOfTerm_(F.roleTeamTerms_(), 25);
  ck(!t25b.registered && t25b.from === null && t25b.last === 24 && t25b.teams.length === 12
     && t25b.teams.every((t) => !t.members.length && !t.leader) && t25b.teams.some((t) => t.name === '広報チーム')
     && JSON.stringify(t25b.teams.map((t) => [t.key, t.name])) === JSON.stringify(F.roleTeamsOfTerm_(F.roleTeamTerms_(), 24).teams.map((t) => [t.key, t.name])),
     '25期（未登録）のチーム: ' + JSON.stringify(t25b.teams.filter((t) => t.members.length || t.leader || /広報|トレーニング/.test(t.name))));
  // 23期（24期より前・未登録）も、24期の顔ぶれを出さない
  const t23b = F.roleTeamsOfTerm_(F.roleTeamTerms_(), 23);
  ck(!t23b.registered && t23b.from === null && t23b.teams.every((t) => !t.members.length && !t.leader), '23期（未登録）のチーム: ' + JSON.stringify(t23b.teams[2]));
  // その期のチーム（トークスクリプトなどで使う）も、未登録の期は空
  ck(F.roleTeams_(new Date(2027, 4, 12)).every((t) => !t.members.length && !t.leader), '未登録の期のチーム（roleTeams_）に顔ぶれが入っている');
  F.ROLE_ROSTER_CHECKED_ = true;
  const c24t = F.getRoleInputContext('2026/10/07', '');
  const e24 = c24t.holderTerms.find((t) => t.term === 24), e25 = c24t.holderTerms.find((t) => t.term === 25);
  ck(c24t.holderTerm.teamsRegistered && e24.teams.length === 12 && !e25.teamsRegistered && e25.teamsFrom === null && e25.teamsLast === 24
     && e25.teams.every((t) => !t.members.length),
     '画面に渡す委員会・チーム: ' + JSON.stringify(c24t.holderTerms.map((t) => [t.term, t.teamsRegistered, t.teamsFrom, t.teamsLast])));
  // チームがみな空の期は、13の役職だけを直す（役職の欄を全部書き直さない）
  F.saveRoleHolders(null, 25, '2026/10/07', [{ key: 'role:vhc', name: 'ビジターホスト', members: [] }]);
  ck(F.previewRoleHoldersRoster(25).full === false, 'みな空のチームで、役職の欄を全部書き直そうとした');
  delete props.BNI_ROLE_TEAMS_25;

  // 割り振り：役職から読み取る。9月（期の最後の月）は引継ぎの時期なので、次の期（24期）を選んでおく
  let vh = F.getVisitorHostsFromRoles();
  ck(vh.ok && vh.now === 23 && vh.pick === 24 && vh.handover && vh.terms.map((t) => t.term).join() === '23,24'
     && vh.terms[1].names.join('、') === '藤本 礼子、山内 登志夫、丘野 秀人、成田 幸司' && vh.terms[1].registered
     && !vh.terms[0].registered && vh.terms[0].from === null && vh.terms[0].names.join('、') === '藤本 礼子',
     'ビジターホストの読み取り: ' + JSON.stringify(vh));
  // 次の期の委員会・チームが未登録なら、今の期
  const keep24 = props.BNI_ROLE_TEAMS_24;
  delete props.BNI_ROLE_TEAMS_24;
  vh = F.getVisitorHostsFromRoles();
  ck(vh.ok && vh.pick === 23 && !vh.handover && vh.terms[1].names.join('、') === '藤本 礼子', '次の期が未登録のときの読み取り: ' + JSON.stringify(vh));
  props.BNI_ROLE_TEAMS_24 = keep24;
}
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
console.log('OK: 項目（担当の列から）・推定・入力状況・保存（行の追加・衝突・数式・役職の制限）・役職とチーム（半期ごと・名簿の役職への反映・ビジターホストの読み取り）');
