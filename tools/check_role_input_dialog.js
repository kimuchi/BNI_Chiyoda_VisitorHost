// 「役職ごとの入力」の画面（role_input.html）を、Node上の簡易DOMで動かしてみる。
// サーバーの代わりに、本番と同じ role_input_srv.js（ルーティンチェックシートの写しにつないだもの）を呼ぶので、
// 一覧 → 役職の入力 → 保存 → シートへの書き込み まで通しで確かめられる。
//
//   node tools/check_role_input_dialog.js <routine.json> <members.json>

const { loadPage } = require('./lib_minidom');
const { makeRoleServer } = require('./lib_role_fixture');

const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

const S = makeRoleServer(process.argv[2], process.argv[3]);
const calls = [];
const server = {
  getSystemVersion: () => 'test',
  getRoleInputContext: (d, r) => { calls.push(['ctx', d, r]); return S.F.getRoleInputContext(d, r); },
  saveRoleInput: (d, r, e) => { calls.push(['save', d, r, e]); return S.F.saveRoleInput(d, r, e); },
  saveRoleHolders: (m) => { calls.push(['holders', m]); return S.F.saveRoleHolders(m); },
  getSpeakerRotation: () => { calls.push(['rot']); return S.F.getSpeakerRotation(); },
  saveSpeakerRotation: (d) => { calls.push(['rotSave', d]); return S.F.saveSpeakerRotation(d); },
  importSpeakerRotation: (j) => { calls.push(['rotImport', j]); return S.F.importSpeakerRotation(j); },
};
// 本番はサーバーがURLの役職（?role=…）を画面に埋め込む。ここでは同じ置き換えをしてから読む
function open(role, view) {
  return loadPage('role_input.html', {
    server, fails,
    preprocess: (p) => p.replace(/<\?\s*var roleParam[\s\S]*?\?>/, '')
      .replace('<?= roleParam ?>', role || '').replace('<?= viewParam ?>', view || ''),
  });
}
const shown = (el) => !!el && el.style.display !== 'none';
const idxOf = (page, title, parent) => page.run(`ids.findIndex(function(id){ var it=ctx.items[id];
  return it.title.replace(/\\s/g,'').indexOf(${JSON.stringify(title.replace(/\s/g, ''))})===0
    && (${JSON.stringify(parent === undefined ? null : parent)}===null || it.parent===${JSON.stringify(parent || '')}); })`);

// ===================== 一覧 =====================
let page = open('');
let { els, window, run, step } = page;
page.flush();
window.onload();
ck(shown(els.loading) && /読み込み中/.test(els.loading.innerHTML), '開いた直後に「読み込み中」が出ていない');
ck(els.saveBtn.disabled === true, '読み込み中に保存ボタンが押せる');
step('一覧を読み込む', () => {});
ck(!shown(els.loading), '読み込みが終わっても「読み込み中」が消えない');
ck(shown(els.overview) && !shown(els.roleView), '役職の指定が無いのに一覧が出ていない');
ck(els.meeting.value === '2026/09/30' && /第535回/.test(els.meeting.options[0].text), '開催日の選択: ' + els.meeting.value);
ck(/23期/.test(els.sheetNote.innerText) && /第535回/.test(els.sheetNote.innerText), 'シートの表示: ' + els.sheetNote.innerText);
ck((els.cards.innerHTML.match(/入力する →/g) || []).length === 13, 'カードが13枚ない');
ck(/未入力あり 13/.test(els.summary.innerHTML), '入力状況のまとめ: ' + els.summary.innerHTML.replace(/<[^>]+>/g, ' '));
ck(/メンターコーディネーター[\s\S]*?伊五澤[\s\S]*?今週の共有事項（9\/28/.test(els.cards.innerHTML), 'メンターコーディネーターのカード');
ck(!!els.hd_0 && els.hd_0.value === '熊谷 龍威', '担当者の選択: ' + (els.hd_0 && els.hd_0.value));

// ===================== バイスプレジデントの入力 =====================
step('バイスプレジデントを開く', () => run("openRole('vice')"));
ck(shown(els.roleView) && !shown(els.overview) && els.roleName.innerText === 'バイスプレジデント', '役職の画面が出ていない');
const iPolicy = idxOf(page, '一般規定'), iLate = idxOf(page, '遅刻・欠席担当'), iSub = idxOf(page, '代理', '代理・欠席');
const iVis = idxOf(page, '人数：ビジター'), iShare = idxOf(page, '今週の共有事項（バイスプレジデント）'), iNew = idxOf(page, '新入会');
ck(iPolicy >= 0 && els['f_' + iPolicy].value === '3番', '一般規定の初期値: ' + (els['f_' + iPolicy] || {}).value);
ck(iLate >= 0 && els['f_' + iLate].value === '藤田さん', '遅刻・欠席担当の初期値: ' + (els['f_' + iLate] || {}).value);
ck(iVis >= 0 && els['f_' + iVis].value === '2', 'ビジター人数の初期値: ' + (els['f_' + iVis] || {}).value);
ck(iSub >= 0 && els['f_' + iSub].value === S.SUB_FOR.name.split(' ')[0] + 'さん', '代理の初期値: ' + (els['f_' + iSub] || {}).value);
ck(/<b>1名<\/b>/.test(els['it_' + iSub].innerHTML), '代理の人数が出ていない');
ck(/推定で入れています（参加者シート/.test(els['it_' + iSub].innerHTML), '推定の出どころが出ていない');
ck(/前回（9\/23）/.test(els['it_' + iSub].innerHTML), '前回の値が出ていない');
ck(/必須・9\/28\(月\)まで/.test(els['it_' + iPolicy].innerHTML), '必須と期日の表示: ' + els['it_' + iPolicy].innerHTML.slice(0, 200));
ck(/保存していない変更 \d+件/.test(els.dirty.innerText), '未保存の件数: ' + els.dirty.innerText);
ck(/必須 <b>\d+<\/b> 項目のうち/.test(els.progress.innerHTML), '進み具合: ' + els.progress.innerHTML);
ck(els.saveBtn.disabled === false, '保存ボタンが押せない');

// 前回と同じ → 人数も変わる → 推定に戻す
step('代理を前回と同じにする', () => run(`usePrev(${iSub})`));
ck(/、/.test(els['f_' + iSub].value) && /<b>2名<\/b>/.test(els['it_' + iSub].innerHTML), '前回と同じにしたあと: ' + els['f_' + iSub].value);
step('代理を推定に戻す', () => run(`useEst(${iSub})`));
ck(els['f_' + iSub].value === S.SUB_FOR.name.split(' ')[0] + 'さん', '推定に戻らない: ' + els['f_' + iSub].value);
// 入力欄を直接書き換える
step('共有事項を書く', () => { els['f_' + iShare].value = '・引継式の予算承認\n・チャプターサイズ目標53名'; run(`edit(${iShare})`); });
step('新入会を「なし」にする', () => run(`setVal(${iNew},'なし')`));
ck(els['f_' + iNew].value === 'なし', '「なし」ボタン');

const nMissBefore = S.F.getRoleInputContext('2026/09/30').roles.find((r) => r.key === 'vice').status.missing.length;
step('保存する', () => run('save()'));
const sv = calls.filter((c) => c[0] === 'save').pop();
ck(sv && sv[1] === '2026/09/30' && sv[2] === 'vice', '保存の呼び出し: ' + JSON.stringify(sv && sv.slice(0, 3)));
const sent = sv ? sv[3] : [];
const sentOf = (t) => sent.find((e) => e.id.indexOf(t) >= 0);
ck(sent.length > 10 && sentOf('一般規定') && sentOf('一般規定').value === '3番' && sentOf('一般規定').orig === '',
   '保存した内容: ' + JSON.stringify(sent.slice(0, 3)));
ck(!sent.some((e) => /メインプレゼン/.test(e.id)), 'ほかの役職の項目を送っている');
ck(/保存しました（\d+項目）/.test(els.msg.innerHTML) && /行（23行）を足しました/.test(els.msg.innerHTML), '保存のお知らせ: ' + els.msg.innerHTML);
ck(/入力済み/.test(els['it_' + iPolicy].innerHTML) && !/未保存/.test(els['it_' + iPolicy].innerHTML), '保存後の表示: ' + els['it_' + iPolicy].innerHTML.slice(0, 200));
ck(/変更はありません/.test(els.dirty.innerText), '保存後も未保存の変更が残っている: ' + els.dirty.innerText);
const s23 = S.sheets.find((s) => s.getName() === '【23期】ルーティンチェックシート');
const col = s23._grid[0].indexOf('2026/09/30');
const cellOf = (title) => { const r = s23._grid.findIndex((row) => [row[2], row[3], row[4]].some((v) => String(v || '').replace(/\s/g, '') === title)); return r >= 0 ? s23._grid[r][col] : undefined; };
ck(cellOf('一般規定') === '3番' && String(cellOf('人数：ビジター')) === '2' && /引継式/.test(cellOf('今週の共有事項（バイスプレジデント）') || ''),
   'シートに入っていない: ' + JSON.stringify([cellOf('一般規定'), cellOf('人数：ビジター'), cellOf('今週の共有事項（バイスプレジデント）')]));
const nMissAfter = S.F.getRoleInputContext('2026/09/30').roles.find((r) => r.key === 'vice').status.missing.length;
ck(nMissAfter < nMissBefore, '保存しても足りない項目が減らない: ' + nMissBefore + ' → ' + nMissAfter);

// 一覧に戻ると、バイスの足りない項目が減っている（未保存が無いので確認は出ない）
const confirmsBefore = page.log.confirms.length;
step('一覧に戻る', () => run('showOverview()'));
ck(page.log.confirms.length === confirmsBefore, '保存したのに「保存していない入力があります」が出た');
ck(shown(els.overview), '一覧に戻らない');

// 担当者を変える
step('担当者を変える', () => { els.hd_0.value = els.hd_0.options[2].value; run('saveHolders()'); });
const newHolder = els.hd_0.options[2] ? els.hd_0.options[2].value : '';
ck(calls.some((c) => c[0] === 'holders' && c[1].president === newHolder), '担当者の保存が呼ばれていない');
ck(newHolder && els.cards.innerHTML.indexOf(newHolder) >= 0, 'カードの担当者が変わらない');

// 開催日を変える（10/7 は 24期のシート）
step('10/7 に変える', () => { els.meeting.value = '2026/10/07'; run('changeMeeting()'); });
ck(/24期/.test(els.sheetNote.innerText) && /第536回/.test(els.sheetNote.innerText), '10/7 に変わらない: ' + els.sheetNote.innerText);

// ===================== URLで役職を指定して開く =====================
page = open('ec');
({ els, window, run, step } = page);
page.flush();
window.onload();
step('ECを開く', () => {});
ck(shown(els.roleView) && els.roleName.innerText === 'エデュケーションコーディネーター', '役職を指定して開いても入力画面にならない');
const iEdu = idxOf(page, 'エデュケーション');
ck(iEdu >= 0 && !!els['f_' + iEdu], 'エデュケーションの入力欄が無い');
// 書きかけで一覧に戻ろうとすると確かめる
step('書きかけで戻る', () => { els['f_' + iEdu].value = '見本さん'; run(`edit(${iEdu})`); run('showOverview()'); });
ck(page.log.confirms.length === 1 && /保存していない入力があります/.test(page.log.confirms[0]), '書きかけで戻っても確認が出ない');

// ===================== スピーカーローテーション（書記兼会計）=====================
page = open('secretary', 'rotation');
({ els, window, run, step } = page);
page.flush();
window.onload();
step('ローテーションを開く', () => {});
ck(shown(els.rotView) && !shown(els.roleView) && !shown(els.overview), 'ローテーションの画面が出ていない');
ck(shown(els.rotOpenBtn) === true, '書記兼会計の画面に「スピーカーローテーションの管理」ボタンが無い');
ck(/9\/30\(水\) <b>次回<\/b>/.test(els.rotWeeks.innerHTML) && /確定/.test(els.rotWeeks.innerHTML)
   && (els.rotWeeks.innerHTML.match(/<tr/g) || []).length === 13, '予定の表: ' + els.rotWeeks.innerHTML.slice(0, 200));
ck(/10\/14 の回から/.test(els.rotAnchorNote.innerHTML) && /仲宗根 愛里さん・葉山 成男さん/.test(els.rotAnchorNote.innerHTML), '起点の説明: ' + els.rotAnchorNote.innerHTML);
ck((els.rotOrder.innerHTML.match(/<tr/g) || []).length === 46 && /仲宗根 愛里 <span class="badge req">10\/14の回はここから/.test(els.rotOrder.innerHTML), '並び順の表');
ck(/休会日に入っていません/.test(els.rotMissing.innerHTML) && /11\/11/.test(els.rotMissing.innerHTML), '休会日のお知らせが出ていない');
ck(/＜【9月30日定例会】メインプレゼンターのご案内＞/.test(els.rotFb.value), 'Facebookの文: ' + els.rotFb.value.slice(0, 40));
step('Facebookの文をコピー', () => run('copyFb()'));
ck(page.log.copies === 1, 'コピーされない');

const weekCell = (md) => { const m = els.rotWeeks.innerHTML.match(new RegExp(md.replace('/', '\\/') + '\\(水\\)[^]*?<\\/tr>')); return m ? m[0].replace(/<[^>]+>/g, ' ') : ''; };
step('入れ替え', () => { els.rotSwapA.value = '長見 響児'; els.rotSwapB.value = '徳山 京介'; run('rotSwap()'); });
ck(/星本 充輝\s+徳山 京介/.test(weekCell('10/28')) && /保存していない変更/.test(els.rotDirty.innerText), '入れ替えが予定に出ない: ' + weekCell('10/28'));
step('上下に動かす', () => run('rotMove(0,1)'));
ck(run('rOrder[0]') === '金子 美緒' && run('rOrder[1]') === '山本 登一郎' && run('rAnchor.pointer') === 5, '上下に動かしたあと');
step('外す', () => run('rotDel(2)'));
ck(run('rAnchor.pointer') === 4 && run('rOrder[rAnchor.pointer]') === '仲宗根 愛里' && /外しますか/.test(page.log.confirms.slice(-1)[0] || ''),
   '前を外したのに起点がずれた: ' + run('rOrder[rAnchor.pointer]'));
step('入れる', () => { els.rotAddName.value = '渡邉 真理子'; els.rotAddPos.value = '0'; run('rotAdd()'); });
ck(run('rOrder[0]') === '渡邉 真理子' && run('rOrder[rAnchor.pointer]') === '仲宗根 愛里', '前に入れたのに起点がずれた');
const iNaka = run("rOrder.indexOf('中込 渉')");
step('中込さんを対象外にする', () => { els['rx_' + iNaka].checked = false; run(`rotToggle(${iNaka})`); });
ck(/竹田 明日翔\s+星本 充輝/.test(weekCell('10/21')), '対象外にしたのに予定に出る: ' + weekCell('10/21'));
step('見出しを直して保存', () => { els.rotHeader.value = 'メインプレゼンテーション（各５分）'; run('rotSave()'); });
const rs = calls.filter((c) => c[0] === 'rotSave').pop();
ck(rs && rs[1].anchor.date === '2026/10/14' && rs[1].anchor.pointer === 5 && rs[1].excluded.indexOf('中込 渉') >= 0 && rs[1].base === '',
   '保存した内容: ' + JSON.stringify(rs && { a: rs[1].anchor, b: rs[1].base }));
ck(/保存しました/.test(els.rotMsg.innerHTML) && !/保存していない変更/.test(els.rotDirty.innerText), '保存のお知らせ: ' + els.rotMsg.innerHTML);
const after = S.F.getSpeakerRotation();
ck(after.header === 'メインプレゼンテーション（各５分）' && after.order[0] === '渡邉 真理子'
   && after.weeks.find((w) => w.md.indexOf('10/21') === 0).people.map((p) => p.name).join('・') === '竹田 明日翔・星本 充輝',
   'サーバーの状態: ' + JSON.stringify({ h: after.header, o: after.order.slice(0, 3) }));
step('書記兼会計の入力に戻る', () => run('closeRotation()'));
ck(shown(els.roleView) && !shown(els.rotView) && els.roleName.innerText === '書記兼会計', '書記兼会計の入力に戻らない');

console.log(`役職ごとの入力の画面: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.slice(0, 30).forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 一覧（入力状況・担当者）／役職の入力（推定・前回・人数・保存・シートへの書き込み）／URLでの役職指定／スピーカーローテーション');
