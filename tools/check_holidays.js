// 休会日を選択してそのまま保存→読み戻し→開催日の候補から除外、を実際のChromiumで確かめる。
process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const { launch, openGasPage } = require('./lib_gas_page');
const { loadPage } = require('./lib_minidom');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const env = makeEnv({ now: new Date(2026, 9, 15, 10, 0, 0) });
const srv = Object.assign({}, env.globals);
vm.createContext(srv);
const root = path.join(__dirname, '..');
for (const file of fs.readdirSync(root).filter(f => f.endsWith('.js')).sort()) {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), srv, { filename: file });
}
srv.chapterMeetingBase_ = () => ({ date: new Date(2026, 9, 21), count: 10 });
const J = x => JSON.stringify(x);
async function main() {
  const browser = await launch();
  try {
    for (const timezoneId of ['Asia/Tokyo', 'America/Los_Angeles']) {
      env.reset([['休会日', true, []]], {});
      const context = await browser.newContext({ timezoneId });
      const { page } = await openGasPage(context, 'holiday.html', srv, { fails });
      await page.waitForFunction(() => loaded);
      await page.locator('#newHoliday').fill('2026-10-28');
      await page.locator('#saveBtn').click(); // 「追加」を押さずに保存
      await page.waitForFunction(() => /保存しました/.test(document.getElementById('saveNote').textContent));
      ck(J(srv.getHolidays()) === J(['2026/10/28']), timezoneId + '：選択中の日付が保存されず0日になる');
      ck(/休会日 1日/.test(await page.locator('#saveNote').textContent()), timezoneId + '：保存した件数が1日にならない');
      const candidates = srv.getMeetingCandidates();
      ck(!candidates.some(c => c.dateValue === '2026/10/28'), '休会日が開催日の候補に残る');
      ck(candidates.some(c => c.dateValue === '2026/11/04' && /第11回/.test(c.display)), '休会日の次の開催回がずれる');
      await page.reload();
      await page.waitForFunction(() => loaded);
      ck((await page.locator('#holidayList').textContent()).includes('2026/10/28'), '開き直すと登録した日が消える');
      await page.locator('#newHoliday').fill('2026-10-28');
      await page.locator('#saveBtn').click();
      await page.waitForFunction(() => /保存しました/.test(document.getElementById('saveNote').textContent));
      ck(srv.getHolidays().length === 1, '同じ日を保存すると重複する');
      await page.locator('#newHoliday').fill('2026-12-30');
      await page.getByRole('button', { name: '追加', exact: true }).click();
      await page.locator('#saveBtn').click();
      await page.waitForFunction(() => /保存しました/.test(document.getElementById('saveNote').textContent));
      ck(J(srv.getHolidays()) === J(['2026/10/28', '2026/12/30']), timezoneId + '：追加した日付が前日になる');
      // 空にする操作は、一覧から削除して保存する。
      while (await page.locator('#holidayList button').count()) await page.locator('#holidayList button').first().click();
      await page.locator('#saveBtn').click();
      await page.waitForFunction(() => /休会日 0日/.test(document.getElementById('saveNote').textContent));
      ck(srv.getHolidays().length === 0, '一覧を空にしたあとも古い休会日が残る');
      await context.close();
    }
    // 読み込み・保存の返事が届くまで、日付の追加や二重保存をさせない。
    env.reset([['休会日', true, []]], {});
    const pending = loadPage('holiday.html', { server: srv });
    pending.run('window.onload()');
    ck(pending.els.addBtn.disabled && pending.els.newHoliday.disabled && pending.els.saveBtn.disabled,
      '読み込み中に編集できて、読み込み結果で追加した日が消える');
    pending.els.newHoliday.value = '2026-10-28';
    pending.run('addHoliday()');
    ck(pending.run('holidays.length') === 0, '読み込み完了前に一覧を書き換えた');
    pending.flush();
    pending.run('save(); save()');
    ck(pending.log.calls.filter(c => c.name === 'saveHolidays').length === 1 && pending.els.newHoliday.disabled,
      '保存中の二重保存・編集を防げない');
    pending.flush();
    ck(!pending.els.saveBtn.disabled && J(srv.getHolidays()) === J(['2026/10/28']), '返事を受け取った後に保存を再開できない');
    const failedContext = await browser.newContext({ timezoneId: 'Asia/Tokyo' });
    const { page: failedPage } = await openGasPage(failedContext, 'holiday.html', Object.assign({}, srv, {
      saveHolidays: () => { throw new Error('書き込みに失敗しました'); }
    }), { fails });
    await failedPage.waitForFunction(() => loaded);
    await failedPage.locator('#newHoliday').fill('2026-11-04');
    await failedPage.locator('#saveBtn').click();
    await failedPage.waitForFunction(() => /保存できませんでした/.test(document.getElementById('saveNote').textContent));
    ck(J(srv.getHolidays()) === J(['2026/10/28']) && (await failedPage.locator('#holidayList').textContent()).includes('2026/11/04'),
      '保存に失敗したときに元の設定や入力した日付を消した');
    ck(await failedPage.locator('#saveBtn').isEnabled(), '保存失敗後にやり直せない');
    await failedContext.close();
    // 不正な日付が届いても、既存の設定を消さない。
    env.reset([['休会日', true, [['2026/12/30']]]], {});
    let rejected = false;
    try { srv.saveHolidays(['2026/02/30']); } catch (e) { rejected = true; }
    ck(rejected && J(srv.getHolidays()) === J(['2026/12/30']), '不正な日付の保存で既存の休会日が消えた');
    srv.saveHolidays(['2026-12-30', '2026/12/30', '2027/1/6']);
    ck(J(srv.getHolidays()) === J(['2026/12/30', '2027/01/06']), '形式をそろえて重複を除けない');
    env.reset([['休会日', true, [['休会日'], [new Date(2026, 11, 30)], ['2026/12/30'], ['２０２７／１／６']]]], {});
    ck(J(srv.getHolidays()) === J(['2026/12/30', '2027/01/06']), '既存の日付セル・全角の文字から休会日を読み戻せない');
  } finally { await browser.close(); }
}
main().then(() => {
  if (fails.length) { console.log('NG ' + fails.length + '件 / ' + checks + '件'); console.log(fails.join('\n')); process.exit(1); }
  console.log('休会日の保存・読み戻し・開催日の候補（ブラウザ）: 検査 ' + checks + ' 件 OK');
}).catch(e => { console.log(fails.join('\n')); console.error(e); process.exit(1); });
