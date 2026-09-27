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
  saveRoleHolders: (m, t, d) => { calls.push(['holders', m, t, d]); return S.F.saveRoleHolders(m, t, d); },
  previewRoleHoldersRoster: (t) => { calls.push(['rosterPreview', t]); return S.F.previewRoleHoldersRoster(t); },
  applyRoleHoldersToRoster: (t) => { calls.push(['rosterApply', t]); return S.F.applyRoleHoldersToRoster(t); },
  getSpeakerRotation: () => { calls.push(['rot']); return S.F.getSpeakerRotation(); },
  saveSpeakerRotation: (d) => { calls.push(['rotSave', d]); return S.F.saveSpeakerRotation(d); },
  importSpeakerRotation: (j) => { calls.push(['rotImport', j]); return S.F.importSpeakerRotation(j); },
  getPreMeetingPreview: (d) => { calls.push(['premtgPreview', d]); return S.F.getPreMeetingPreview(d); },
  generatePreMeetingSlides: (d) => { calls.push(['premtgMake', d]); return S.F.generatePreMeetingSlides(d); },
  saveSpeakerRotationImage: (b, n) => { calls.push(['rotImage', n]); return S.F.saveSpeakerRotationImage(b, n); },
};
// 事前MTGのパワポの保存先（Drive）は使わない
S.F.saveOutputFile_ = (blob, name) => ({ id: 'x', url: 'https://example/' + name, downloadUrl: 'https://example/dl/' + name });
// 本番はサーバーがURLの役職（?role=…）を画面に埋め込む。ここでは同じ置き換えをしてから読む
function open(role, view, now) {
  return loadPage('role_input.html', {
    server, fails, now,
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

// 担当者は半期ごと（4〜9月・10〜3月）。9/30 は23期。まだ登録していないので、24期の担当者を出している
const termTexts = () => els.holderTerm.options.map((o) => o.text).join(' / ');
const presidentCard = () => (els.cards.innerHTML.match(/プレジデント<\/span><div class="who">([^<]*)</) || [])[1];
ck(els.holderTerm.value === '23' && /23期（2026年4月〜9月）・この開催日・未登録/.test(termTexts()) && /24期（2026年10月〜2027年3月）/.test(termTexts()),
   '担当者の期の選択: ' + termTexts());
ck(/23期の担当者はまだ登録されていません。24期の担当者を出しています/.test(els.holderNote.innerText) && !shown(els.holderWarn),
   '未登録の期の説明: ' + els.holderNote.innerText + ' / ' + els.holderWarn.innerHTML);
// 担当者を変える（23期として保存）
step('担当者を変える', () => { els.hd_0.value = els.hd_0.options[2].value; run('saveHolders()'); });
const newHolder = els.hd_0.options[2] ? els.hd_0.options[2].value : '';
const hc23 = calls.filter((c) => c[0] === 'holders').pop();
ck(hc23 && hc23[1].president === newHolder && hc23[2] === 23 && hc23[3] === '2026/09/30', '担当者の保存: ' + JSON.stringify(hc23 && hc23.slice(2)));
ck(newHolder && presidentCard() === newHolder, 'カードの担当者が変わらない: ' + presidentCard());
ck(els.holderNote.innerText === '' && /23期（2026年4月〜9月）の担当者を保存しました/.test(els.holderMsg.innerText), '保存のお知らせ: ' + els.holderMsg.innerText);

// 開催日を変える（10/7 は 24期のシート。担当者も24期）
step('10/7 に変える', () => { els.meeting.value = '2026/10/07'; run('changeMeeting()'); });
ck(/24期/.test(els.sheetNote.innerText) && /第536回/.test(els.sheetNote.innerText), '10/7 に変わらない: ' + els.sheetNote.innerText);
ck(els.holderTerm.value === '24' && els.hd_0.value === '熊谷 龍威' && presidentCard() === '熊谷 龍威',
   '10/7（24期）の担当者: ' + els.holderTerm.value + ' ' + presidentCard());
// 次の期（25期）の担当者を前もって登録する。10/7 のカードは24期のまま
step('25期を選ぶ', () => { els.holderTerm.value = '25'; run('renderHolders()'); });
ck(/25期の担当者はまだ登録されていません。24期の担当者を出しています/.test(els.holderNote.innerText) && els.hd_0.value === '熊谷 龍威',
   '25期を選んだとき: ' + els.holderNote.innerText);
step('25期のプレジデントを登録', () => { els.hd_0.value = els.hd_0.options[3].value; run('saveHolders()'); });
const hc25 = calls.filter((c) => c[0] === 'holders').pop();
ck(hc25 && hc25[2] === 25 && hc25[3] === '2026/10/07' && els.holderTerm.value === '25' && els.holderNote.innerText === ''
   && els.hd_0.value === hc25[1].president, '25期の保存: ' + JSON.stringify(hc25 && hc25.slice(2)) + ' ' + els.holderTerm.value);
ck(presidentCard() === '熊谷 龍威', '25期を登録したら、10/7（24期）のカードが変わった: ' + presidentCard());
// メンバー名簿の「役職」に反映する。いまは23期（今の期）を保存したときに反映した23期の担当者
step('24期を選ぶ', () => { els.holderTerm.value = '24'; run('renderHolders()'); });
ck(/23期（2026年4月〜9月）の担当者を反映しています（2026\/09\/26）/.test(els.rosterNote.innerText) && /自動で反映します/.test(els.rosterNote.innerText),
   '名簿への反映の様子: ' + els.rosterNote.innerText);
// 担当者の変更を保存していないときは、先に保存してもらう
step('保存せずに反映しようとする', () => { els.hd_1.value = els.hd_1.options[4].value; run('rosterRoles()'); });
ck(/先に「この期の担当者を保存」/.test(els.holderMsg.innerText) && !calls.some((c) => c[0] === 'rosterPreview'), '保存していない変更があるのに反映した');
step('選び直す', () => run('renderHolders()'));
const nConf = page.log.confirms.length;
step('名簿の役職に反映', () => run('rosterRoles()'));
const conf = page.log.confirms[nConf] || '';
ck(/24期（2026年10月〜2027年3月）の担当者に合わせて直します/.test(conf) && /・熊谷 龍威：（空欄） → プレジデント/.test(conf)
   && new RegExp('・' + newHolder + '：プレジデント → （空欄）').test(conf), '反映の前の確かめ: ' + conf.slice(0, 300));
ck(calls.some((c) => c[0] === 'rosterApply' && c[1] === 24) && /24期の担当者に合わせて直しました（\d+名）/.test(els.holderMsg.innerText)
   && /24期（2026年10月〜2027年3月）の担当者を反映しています/.test(els.rosterNote.innerText),
   '名簿の役職に反映したあと: ' + els.holderMsg.innerText + ' / ' + els.rosterNote.innerText);
step('もう一度反映', () => run('rosterRoles()'));
ck(/もう24期の担当者どおりです/.test(els.holderMsg.innerText) && page.log.confirms.length === nConf + 1, '2回目の反映: ' + els.holderMsg.innerText);
// 期が替わったのに、新しい期の担当者がまだ無いときは、一覧の上で知らせる
step('新しい期の担当者が未登録', () => run("ctx.holderTerm={term:26,label:'2027年10月〜2028年3月',registered:false,from:25,holders:{}}; showOverview(true)"));
ck(shown(els.holderWarn) && /26期（2027年10月〜2028年3月）の担当者がまだ登録されていません。いまは25期の担当者を出しています/.test(els.holderWarn.innerHTML),
   '新しい期のお知らせ: ' + els.holderWarn.innerHTML);

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
// ローテーションを先に読む。書記兼会計の入力（初期値の推定に時間がかかる）は「← 書記兼会計の入力」を押したときに読む
let nCalls = calls.length;
const callsSince = () => calls.slice(nCalls).map((c) => c.slice(0, 3));
page = open('secretary', 'rotation');
({ els, window, run, step } = page);
page.flush();
window.onload();
ck(callsSince().some((c) => c[0] === 'rot') && !callsSince().some((c) => c[0] === 'ctx' && c[2] === 'secretary'),
   '開いてすぐに読むのがローテーションだけになっていない: ' + JSON.stringify(callsSince()));
step('ローテーションを開く', () => {});
ck(shown(els.rotView) && !shown(els.roleView) && !shown(els.overview), 'ローテーションの画面が出ていない');
ck(els.meeting.value === '2026/09/30' && /23期/.test(els.sheetNote.innerText), '上の帯（開催日・シート）が出ていない: ' + els.sheetNote.innerText);
ck(shown(els.rotOpenBtn) === true, '書記兼会計の画面に「スピーカーローテーションの管理」ボタンが無い');
ck(/9\/30\(水\) <b>次回<\/b>/.test(els.rotWeeks.innerHTML) && /確定/.test(els.rotWeeks.innerHTML)
   && (els.rotWeeks.innerHTML.match(/<tr/g) || []).length === 13, '予定の表: ' + els.rotWeeks.innerHTML.slice(0, 200));
ck(/10\/14 の回から/.test(els.rotAnchorNote.innerHTML) && /仲宗根 愛里さん・葉山 成男さん/.test(els.rotAnchorNote.innerHTML), '起点の説明: ' + els.rotAnchorNote.innerHTML);
ck((els.rotOrder.innerHTML.match(/<tr/g) || []).length === 46 && /仲宗根 愛里 <span class="badge req">10\/14の回はここから/.test(els.rotOrder.innerHTML), '並び順の表');
ck(/休会日に入っていません/.test(els.rotMissing.innerHTML) && /11\/11/.test(els.rotMissing.innerHTML), '休会日のお知らせが出ていない');
ck(/＜【9月30日定例会】メインプレゼンターのご案内＞/.test(els.rotFb.value), 'Facebookの文: ' + els.rotFb.value.slice(0, 40));
step('Facebookの文をコピー', () => run('copyFb()'));
ck(page.log.copies === 1, 'コピーされない');

// Facebookに添付する画像（前半スライドの表と同じ作り。ご案内する回から5回分）
const imgOf = () => JSON.parse(Buffer.from(String(run('rotImgData')).split(',')[1] || '', 'base64').toString() || '{"texts":[]}');
const drawn = () => imgOf().texts.map((t) => t.text);
ck(els.rotFbWeek.options.length === 12 && els.rotFbWeek.value === '0', 'ご案内する回（ふつうの日は次回）: ' + els.rotFbWeek.value);
ck(els.rotFb.value === S.F.getSpeakerRotation().fbText, '画面の投稿文が、サーバーの文面と違う');
let dr = drawn();
ck(['日　程', 'メインプレゼンテーション（各４分45秒）', '第535回', '9月30日', '10月28日'].every((t) => dr.includes(t)) && !dr.includes('11月4日'),
   '画像の中身（9/30〜10/28 の5回分）: ' + dr.slice(0, 12).join(' '));
ck(dr.some((t) => /^※　２週間前までに/.test(t)), '画像に注意書きが無い');
ck(run('rotImgName') === '20260930_スピーカーローテーション.png' && els.rotImg.style.display !== 'none', '画像の名前・表示: ' + run('rotImgName'));
step('10/7 のご案内にする', () => { els.rotFbWeek.value = '1'; run('renderFb()'); });
dr = drawn();
ck(/^＜【10月7日定例会】メインプレゼンターのご案内＞/.test(els.rotFb.value) && /（１）渡邉 真理子さん\//.test(els.rotFb.value), '10/7 の投稿文: ' + els.rotFb.value.slice(0, 60));
ck(dr.includes('10月7日') && dr.includes('11月4日') && !dr.includes('9月30日'), '10/7 からの画像: ' + dr.slice(0, 12).join(' '));
ck(imgOf().texts.filter((t) => t.color === '#C00000').slice(0, 2).map((t) => t.text).join('・') === '渡邉 真理子・岡本 翔太',
   '画像の1回目のお2人（赤）: ' + imgOf().texts.filter((t) => t.color === '#C00000').slice(0, 2).map((t) => t.text).join('・'));
step('画像を保存', () => run('rotImgDownload()'));
ck(page.log.downloads.length === 1 && page.log.downloads[0].name === '20261007_スピーカーローテーション.png'
   && /^data:image\/png;base64,/.test(page.log.downloads[0].href), '画像の保存: ' + JSON.stringify(page.log.downloads.map((d) => d.name)));
step('画像をDriveに保存', () => run('rotImgDrive()'));
// （Nodeの画面では画像の中身が本物のPNGではないので、サーバーは「画像が空です」と断る。呼び出しとお知らせを確かめる）
ck(calls.some((c) => c[0] === 'rotImage' && c[1] === '20261007_スピーカーローテーション.png') && /画像が空です/.test(els.rotImgMsg.innerText),
   'Driveに保存: ' + (els.rotImgMsg.innerText || els.rotImgMsg.innerHTML));
step('見出しを変えると画像も描き直す', () => { els.rotHeader.value = '見本の見出し'; run('renderFb()'); });
ck(drawn().includes('見本の見出し'), '見出しを変えても画像が変わらない');
step('ご案内する回を次回に戻す', () => { els.rotHeader.value = S.F.getSpeakerRotation().header; els.rotFbWeek.value = '0'; run('renderFb()'); });

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
// 並びを直すと、画像も直した並びで描き直す（10/28 は星本さん・徳山さん）
ck(drawn().includes('徳山 京介') && drawn().indexOf('徳山 京介') < drawn().indexOf('長見 響児'), '並びを直したあとの画像: ' + drawn().filter((t) => /さん|[一-龥]{2} /.test(t)).join(' '));
step('書記兼会計の入力に戻る', () => run('closeRotation()'));
ck(shown(els.roleView) && !shown(els.rotView) && els.roleName.innerText === '書記兼会計', '書記兼会計の入力に戻らない');
ck(callsSince().some((c) => c[0] === 'ctx' && c[2] === 'secretary') && Object.keys(run('fields')).length > 0,
   '戻ったときに書記兼会計の入力を読んでいない: ' + JSON.stringify(callsSince().filter((x) => x[0] === 'ctx')));
step('もう一度ローテーションを開いて戻る', () => { run('openRotation()'); });
nCalls = calls.length;
step('2回目は読み直さずに戻る', () => run('closeRotation()'));
ck(shown(els.roleView) && !shown(els.rotView) && !callsSince().length, '2回目に戻るとき: ' + JSON.stringify(callsSince()));

// ローテーションの読み込み中に「← 書記兼会計の入力」を押しても、読み終わってから書記兼会計の入力になる
nCalls = calls.length;
page = open('secretary', 'rotation');
({ els, window, run, step } = page);
page.flush();
window.onload();
step('読み込み中に戻る', () => run('closeRotation()'));
ck(shown(els.roleView) && !shown(els.rotView) && !shown(els.loading) && els.roleName.innerText === '書記兼会計'
   && callsSince().some((c) => c[0] === 'ctx' && c[2] === 'secretary'),
   '読み込み中に戻ったあと: ' + JSON.stringify({ role: shown(els.roleView), rot: shown(els.rotView), loading: shown(els.loading), calls: callsSince() }));

// 定例会の日（9/30）に開くと、終わったあとに投稿するので「ご案内する回」は次の回（10/7）
page = open('secretary', 'rotation', '2026-09-30T20:00:00');
({ els, window, run, step } = page);
page.flush();
window.onload();
step('定例会の日にローテーションを開く', () => {});
ck(els.rotFbWeek.value === '1' && /次の回を選んであります/.test(els.rotFbWeekNote.innerText) && /^＜【10月7日定例会】/.test(els.rotFb.value),
   '定例会の日のご案内する回: ' + els.rotFbWeek.value + ' ' + els.rotFb.value.slice(0, 20));

// ===================== 事前MTG（朝イチMTG）のパワポ =====================
// 前のリンク（?p=premtg）・前のメニューの「事前MTG（朝イチMTG）のパワポ」は、一覧を開いてすぐ中身の確かめを出す
page = open('', 'premtg');
({ els, window, run, step } = page);
page.flush();
window.onload();
step('事前MTGを開く', () => {});
ck(shown(els.overview) && calls.some((c) => c[0] === 'premtgPreview' && c[1] === '2026/09/30'), '開いてすぐ中身の確かめが出ない');
const pm = els.premtgOut.innerHTML;
ck(/📅 直近のイベント/.test(pm) && /✏️ お願い事項/.test(pm) && /💻 定例会関連/.test(pm), 'まとめの3つのまとまりが出ていない');
ck(/・欠席：/.test(pm) && /・ビジター：2名/.test(pm), '定例会関連の中身（バイスが保存したビジター2名）: ' + pm.replace(/<[^>]+>/g, ' ').slice(0, 300));
ck(/役職のページ（1枚）/.test(pm) && /🎯 バイスプレジデント/.test(pm) && /今週の共有事項が無い役職/.test(pm), '役職のページの割り付け: ' + pm.replace(/<[^>]+>/g, ' ').slice(0, 400));
ck(/空欄の項目/.test(pm) && /ひな形：既定のもの/.test(pm), '空欄の項目・ひな形の表示');
step('パワポを作る', () => run('premtgMake()'));
ck(calls.some((c) => c[0] === 'premtgMake' && c[1] === '2026/09/30'), '作る呼び出しが無い');
ck(/✅/.test(els.premtgOut.innerHTML) && /20260930_BNI事前MTG\.pptx/.test(els.premtgOut.innerHTML)
   && /PowerPointをダウンロード/.test(els.premtgOut.innerHTML), '作ったあとの表示: ' + els.premtgOut.innerHTML.replace(/<[^>]+>/g, ' ').slice(0, 300));
ck(els.premtgBtn.disabled === false, '作ったあとも「パワポを作る」が押せない');
step('開催日を変える', () => { els.meeting.value = '2026/10/07'; run('changeMeeting()'); });
ck(els.premtgOut.innerHTML === '', '開催日を変えても前の結果が残っている');
step('中身を確かめる（10/7）', () => run('premtgPreview()'));
ck(calls.filter((c) => c[0] === 'premtgPreview').pop()[1] === '2026/10/07' && /役職のページ（0枚）/.test(els.premtgOut.innerHTML),
   '10/7 の確かめ: ' + els.premtgOut.innerHTML.replace(/<[^>]+>/g, ' ').slice(0, 200));

console.log(`役職ごとの入力の画面: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.slice(0, 30).forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 一覧（入力状況・担当者（半期ごと）・名簿の役職への反映）／役職の入力（推定・前回・人数・保存・シートへの書き込み）／URLでの役職指定／スピーカーローテーション（Facebookの文と画像）／事前MTGのパワポ');
