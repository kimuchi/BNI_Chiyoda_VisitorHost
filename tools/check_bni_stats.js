// 架空の統計で、月次保存→開催日で選択→1/2ページのPPTXと台本への反映を検査する。
process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const { launch, openGasPage } = require('./lib_gas_page');
const root = path.join(__dirname, '..'), fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const env = makeEnv({ now: new Date(2026, 9, 16, 10) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const file of fs.readdirSync(root).filter(f => f.endsWith('.js')).sort()) vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), F, { filename: file });
env.reset([], {});
const values = { worldMembers: '360,123人超', worldChapters: '11,789チャプター', worldReferrals: '1,888万件超',
  worldRevenue: '約4兆5678億円', japanMembers: '13,210人', japanChapters: '401チャプター' };
let saved = F.saveBniStatsSettings({ month: '2026-09', revision: '', values });
ck(saved.ok, '月次設定を保存できない: ' + saved.message);
const september = F.bniStatsForDate_(new Date(2026, 8, 30));
ck(september && september.month === '2026-09', '9月の設定が9月末に適用されない');
ck(F.bniStatsForDate_(new Date(2026, 7, 31)) === null, '保存した月より前の開催日に未来の数値を使った');
const next = F.getBniStatsSettings('2026-10');
ck(!next.saved && next.sourceMonth === '2026-09' && next.values.worldMembers === values.worldMembers, '翌月の入力用に前回の数値を表示できない');
ck(F.bniStatsForDate_(new Date(2026, 9, 21)).month === '2026-09', '保存前の翌月設定が適用された');
const octoberValues = Object.assign({}, values, { worldMembers: '３６１，２３４人超', japanChapters: '405' });
saved = F.saveBniStatsSettings({ month: '2026-10', revision: next.revision, values: octoberValues, source: '架空の月次資料' });
ck(saved.ok && saved.settings.values.worldMembers === '361,234人超' && saved.settings.values.japanChapters === '405チャプター', '全角数字や数字だけの入力を表示用にそろえられない');
const october = F.bniStatsForDate_(new Date(2026, 9, 21));
ck(october.month === '2026-10' && F.bniStatsForDate_(new Date(2026, 8, 30)).values.worldMembers === values.worldMembers, '翌月の保存で過去の数値を上書きした');
const stale = F.getBniStatsSettings('2026-10');
saved = F.saveBniStatsSettings({ month: '2026-10', revision: stale.revision, values: Object.assign({}, october.values, { japanMembers: '13,333人' }) });
ck(saved.ok, '同じ月の更新ができない');
const conflict = F.saveBniStatsSettings({ month: '2026-10', revision: stale.revision, values });
ck(!conflict.ok && /ほかの方/.test(conflict.message) && F.bniStatsRecord_('2026-10').values.japanMembers === '13,333人', '古い画面から別の方の保存を消した');
const snapshot = JSON.stringify(env.props);
for (const input of [{ month: '2026-13', values }, { month: '2026-11', values: {} },
  { month: '2026-11', values: { worldMembers: '<img src=x>' } }]) {
  ck(!F.saveBniStatsSettings(Object.assign({ revision: '' }, input)).ok, '不正な設定を保存した');
}
ck(JSON.stringify(env.props) === snapshot, '入力エラーで保存済みの値を変更した');
const originalAccess = F.requireSheetAccess_;
F.requireSheetAccess_ = () => { throw new Error('編集権限がありません'); };
ck(!F.saveBniStatsSettings({ month: '2026-11', revision: '', values }).ok, '権限のない方が保存できた');
let denied = false; try { F.getBniStatsSettings(); } catch (_) { denied = true; }
ck(denied, '権限のない方が設定を読み込めた'); F.requireSheetAccess_ = originalAccess;

// 実際のスクリーンショットと同じ項目と配置。数字は赤、単位は黒、ランは分割。
let seq = 1;
const esc = x => F.escapeXml_(x);
const run = (t, color) => '<a:r><a:rPr lang="ja-JP" sz="3200"><a:solidFill><a:srgbClr val="' + (color || '000000') + '"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>' + esc(t) + '</a:t></a:r>';
function shape(text, x, y, w, h, numeric) {
  const id = String(++seq), m = numeric && text.match(/^([^0-9０-９]*)([0-9０-９][0-9０-９,，.．兆億万千]*)(.*)$/);
  const tx = m ? run(m[1] + m[2].slice(0, 2), 'E31B36') + run(m[2].slice(2), 'E31B36') + run(m[3]) : run(text);
  return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="Text ' + id + '"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + w + '" cy="' + h + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>' + tx + '</a:p></p:txBody></p:sp>';
}
function slide(title, rows) {
  seq = 1;
  let xml = shape(title, 3000000, 100000, 6500000, 450000, false);
  for (const r of rows) {
    xml += shape(r.label, r.x, r.y, 3000000, 350000, false);
    xml += shape(r.value, r.x, r.y + 400000, 3200000, 500000, true);
  }
  return '<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + xml + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
}
const rows = [
  { key: 'worldMembers', label: '世界メンバー数', value: '355,000人超', x: 1500000, y: 900000 },
  { key: 'worldChapters', label: '世界チャプター数', value: '11,600チャプター', x: 7800000, y: 1250000 },
  { key: 'worldReferrals', label: '世界リファーラル数', value: '1,790万件超', x: 700000, y: 3500000 },
  { key: 'worldRevenue', label: '世界売上金額数', value: '約4兆3500億円', x: 700000, y: 4650000 },
  { key: 'japanChapters', label: '日本チャプター数', value: '389チャプター', x: 6200000, y: 5050000 },
  { key: 'japanMembers', label: '日本メンバー数', value: '12,732人', x: 9400000, y: 4850000 }
];
function partsFor(xmls) {
  const p = { 'ppt/presentation.xml': '<p:presentation><p:sldIdLst>' + xmls.map((_, i) => '<p:sldId id="' + (256 + i) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships>' + xmls.map((_, i) => '<Relationship Id="rId' + (i + 1) + '" Target="slides/slide' + (i + 1) + '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>').join('') + '</Relationships>' };
  xmls.forEach((x, i) => { p['ppt/slides/slide' + (i + 1) + '.xml'] = x; });
  Object.keys(p).forEach(k => { p[k] = F.Utilities.newBlob(p[k], 'text/xml', k); });
  return p;
}
const text = (p, n) => F.slideText_(F.xmlOf_(p, 'ppt/slides/slide' + n + '.xml'));
const oneOriginal = slide('BNIの目的と概要', rows), one = partsFor([oneOriginal]);
const oneResult = F.applyBniStats_(one, september);
ck(oneResult.pages === 1 && oneResult.missing.length === 0, '1ページの数値欄を見つけられない: ' + JSON.stringify(oneResult));
for (const r of rows) ck(text(one, 1).includes(values[r.key]), '1ページに反映されない: ' + r.label);
const firstNumber = F.riShapes_(F.xmlOf_(one, 'ppt/slides/slide1.xml')).find(s => s.text === values.worldMembers);
ck(!!firstNumber, '世界メンバーの数字と単位が壊れた');
if (firstNumber) {
  const range = F.findShapeRange_(F.xmlOf_(one, 'ppt/slides/slide1.xml'), firstNumber.id);
  const sp = F.xmlOf_(one, 'ppt/slides/slide1.xml').slice(range.start, range.end);
  ck(/srgbClr val="E31B36"[\s\S]*?<a:t>360,123<\/a:t>/.test(sp) && /srgbClr val="000000"[\s\S]*?<a:t>人超<\/a:t>/.test(sp), '数字と単位の色を保てない');
}
const worldRows = rows.filter(r => r.key.startsWith('world')).map((r, i) => Object.assign({}, r, { label: r.label.replace('世界', ''), x: 1000000 + (i % 2) * 5500000, y: 1100000 + Math.floor(i / 2) * 2400000 }));
const japanRows = rows.filter(r => r.key.startsWith('japan')).map((r, i) => Object.assign({}, r, { label: r.label.replace('日本', ''), x: 1000000 + i * 5500000, y: 2300000 }));
const two = partsFor([slide('BNIの目的と概要（世界）', worldRows), slide('BNI Japan', japanRows)]);
const twoResult = F.applyBniStats_(two, september);
ck(twoResult.pages === 2 && twoResult.missing.length === 0, '世界・日本の2ページを更新できない: ' + JSON.stringify(twoResult));
for (const r of rows) ck(text(two, r.key.startsWith('world') ? 1 : 2).includes(values[r.key]), '2ページで取り違えた: ' + r.label);
ck(!text(two, 2).includes(values.worldMembers), '日本のページに世界の数値を入れた');
const unchanged = partsFor([oneOriginal]);
F.applyBniStats_(unchanged, null);
ck(F.xmlOf_(unchanged, 'ppt/slides/slide1.xml') === oneOriginal, '未設定のときにテンプレートを書き換えた');
const partial = partsFor([oneOriginal]);
F.applyBniStats_(partial, { month: '2026-09', values: { japanMembers: values.japanMembers } });
ck(text(partial, 1).includes('355,000人超') && text(partial, 1).includes(values.japanMembers), '空欄の項目も変更した');
const ordinaryXml = slide('バイスプレジデントによる報告', [{ label: 'メンバー数', value: '52人', x: 1000000, y: 1200000 }]);
const ordinary = partsFor([ordinaryXml]); F.applyBniStats_(ordinary, september);
ck(F.xmlOf_(ordinary, 'ppt/slides/slide1.xml') === ordinaryXml, 'チャプター内の数値を世界の数値に変えた');
const inline = partsFor([slide('BNIの目的と概要', [{ label: '世界メンバー数 355,000人超', value: '2026年10月現在', x: 1000000, y: 1200000 }])]);
F.applyBniStats_(inline, september);
ck(text(inline, 1).includes(values.worldMembers) && text(inline, 1).includes('2026年10月現在'), '同じ段落の数値を変更できない、または年月を変更した');
const tokens = partsFor([slide('BNI統計', [{ label: '{{世界メンバー数}}', value: '{{統計情報の適用月}}', x: 1000000, y: 1200000 }])]);
F.applyBniStats_(tokens, september);
ck(text(tokens, 1).includes(values.worldMembers) && text(tokens, 1).includes('2026-09'), '明示した差し込み口に反映できない');
const absent = F.applyBniStats_(partsFor([ordinaryXml]), september);
ck(absent.pages === 0 && absent.missing.length === 6 && /見つからない/.test(absent.message), '数値欄が無いときに反映したと表示した');
const allValues = Object.assign({}, values, { worldCountries: '80か国・地域', japanReferrals: '123万件', japanRevenue: '1兆2345億円' });
const allRecord = { month: '2026-10', values: allValues };
const allWorldRows = worldRows.concat([{ key: 'worldCountries', label: '国と地域数', value: '77か国・地域' }])
  .map((r, i) => Object.assign({}, r, { x: 1000000 + (i % 2) * 5500000, y: 900000 + Math.floor(i / 2) * 1700000 }));
const allJapanRows = japanRows.concat([{ key: 'japanReferrals', label: 'リファーラル数', value: '100万件' },
  { key: 'japanRevenue', label: '売上金額', value: '9999億円' }])
  .map((r, i) => Object.assign({}, r, { x: 1000000 + (i % 2) * 5500000, y: 1100000 + Math.floor(i / 2) * 2400000 }));
const all = partsFor([slide('BNI Global', allWorldRows), slide('日本のBNI', allJapanRows)]);
const allResult = F.applyBniStats_(all, allRecord);
ck(allResult.missing.length === 0 && allResult.updated.length === 9, '国・地域数や日本の件数・売上金額を更新できない: ' + JSON.stringify(allResult));
ck(text(all, 1).includes('80か国・地域') && text(all, 2).includes('123万件') && text(all, 2).includes('1兆2345億円'), '任意の項目の値を取り違えた');
ck(F.bniStatsTalkText_('現在、世界77の国と地域で活動しています。', '77か国', allRecord).includes('80か国・地域'), '台本の国・地域数を更新できない');
const groupStart = '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="1000" name="Statistics"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="300000" y="100000"/><a:ext cx="10972800" cy="6172200"/><a:chOff x="0" y="0"/><a:chExt cx="12192000" cy="6858000"/></a:xfrm></p:grpSpPr>';
const groupedXml = oneOriginal.replace(/(<p:grpSpPr\/>)([\s\S]*?)(<\/p:spTree>)/, (_, a, b, c) => a + groupStart + b + '</p:grpSp>' + c);
const grouped = partsFor([groupedXml]);
ck(F.applyBniStats_(grouped, september).missing.length === 0 && text(grouped, 1).includes(values.japanMembers), 'グループ化された項目名と数値を対応できない');

// 任意の実物テンプレートを渡した場合は、統計ページを組み込んだPPTXも出力して内部参照を検査する。
if (process.argv[2]) {
  const { execFileSync } = require('child_process');
  const template = F.unzipToMap_(F.Utilities.newBlob([...fs.readFileSync(process.argv[2])], 'application/zip', 'template.pptx'));
  const templateOrder = F.slideOrder_(template), outDir = process.argv[3] || '/tmp/bni-stats-samples';
  fs.mkdirSync(outDir, { recursive: true });
  for (const [name, xmls] of [['one-page', [oneOriginal]], ['two-pages', [slide('BNI Global', allWorldRows), slide('BNI Japan', allJapanRows)]]]) {
    const parts = Object.assign({}, template);
    xmls.forEach((xml, i) => F.putXml_(parts, templateOrder[3 + i], xml));
    const result = F.applyBniStats_(parts, allRecord);
    const file = path.join(outDir, name + '.pptx');
    fs.writeFileSync(file, F.zipFromMap_(parts, name + '.pptx')._buf);
    ck(result.pages === xmls.length, name + 'のPPTXを更新できない');
    try { execFileSync('python3', [path.join(__dirname, 'pptx_integrity.py'), file]); ck(true, ''); }
    catch (e) { ck(false, name + 'の内部参照が壊れた: ' + String(e.stdout || e.message)); }
  }
}

// 実際の生成入口が、ブラウザが送った値より保存済みの適用月を優先する。
let generated;
F.editPptxOnServer_ = (kind, name, edit) => {
  generated = partsFor([oneOriginal]);
  return { info: edit(generated), partCount: 1, timing: { 合計: 1 }, saved: { url: 'https://example.com/file' } };
};
const generatedResult = F.generateMeetingSlides('meetingFirst', {}, '2026/09/30', { roleIntro: false, bniStats: { month: '2099-01', values: { worldMembers: '999人' } } });
ck(generatedResult.ok && text(generated, 1).includes(values.worldMembers), '生成入口で開催月に合う保存済みの数値を使わない: ' + generatedResult.message);
ck(generatedResult.message.includes('2026-09'), '生成時にどの月の統計か分からない');
ck(generatedResult.touched === 1 && !generatedResult.message.includes('書き換える箇所が見つかりません'), '統計を更新したのに更新なしと表示した');

const oldTemplate = F.talkTemplate_();
const beforeTalk = JSON.stringify(oldTemplate);
const built = F.talkBuild_(new Date(2026, 8, 30), {});
const rowBy = label => built.rows.find(r => r.slide === label);
ck(rowBy('世界メンバー数').talk.includes(values.worldMembers.replace(/超$/, '')), '既存の台本のメンバー数・見出しを更新できない');
ck(rowBy('世界チャプター数').talk.includes(values.worldChapters), '既存の台本のチャプター数・見出しを更新できない');
ck(rowBy('BNIの数値各種').talk.includes(values.worldReferrals.replace(/超$/, '')) && rowBy('BNIの数値各種').talk.includes(values.worldRevenue), '既存の台本の件数や金額が同期しない');
ck(JSON.stringify(F.talkTemplate_()) === beforeTalk, '台本のひな形自体を書き換えた');
const resolved = F.talkValue_('世界メンバー数', F.talkEnv_(new Date(2026, 8, 30)));
ck(resolved.state === 'ok' && resolved.value === values.worldMembers, '台本の差し込みに対応していない');
ck(F.talkValue_('日本売上金額', F.talkEnv_(new Date(2026, 8, 30))).state === 'missing', '未入力の項目に架空の数値を出した');
ck(F.talkCatalog_().some(g => g.items.some(i => i.key === '世界メンバー数')), '台本の編集画面に差し込みが無い');

async function browserChecks() {
  const browser = await launch();
  try {
    const { page } = await openGasPage(browser, 'bni_stats.html', F, { fails });
    await page.waitForFunction(() => current && !busy);
    ck(await page.locator('#month').inputValue() === '2026-10', '現在の月を表示しない');
    await page.locator('#worldMembers').fill('362,111人超');
    await page.locator('#saveBtn').click();
    await page.waitForFunction(() => document.getElementById('msg').textContent.includes('保存しました'));
    ck(F.bniStatsRecord_('2026-10').values.worldMembers === '362,111人超', 'ブラウザで保存した値が反映されない');
    await page.reload(); await page.waitForFunction(() => current && !busy);
    ck(await page.locator('#worldMembers').inputValue() === '362,111人超', '開き直すと保存した値が消える');
    await page.locator('#month').fill('2026-11');
    await page.locator('#month').dispatchEvent('change');
    await page.waitForFunction(() => current && current.month === '2026-11' && !busy);
    ck((await page.locator('#state').textContent()).includes('2026-10'), '新しい月の引継元が分からない');
    await page.locator('#japanMembers').fill('13,555人');
    await page.locator('#saveBtn').click();
    await page.waitForFunction(() => document.getElementById('msg').textContent.includes('保存しました'));
    ck(F.bniStatsForDate_(new Date(2026, 9, 21)).values.japanMembers === '13,333人' && F.bniStatsForDate_(new Date(2026, 10, 4)).values.japanMembers === '13,555人', '翌月の更新で当月の数値が変わった');
    const loadFailed = await openGasPage(browser, 'bni_stats.html', { getBniStatsSettings: () => { throw new Error('承認が必要です'); } }, { fails });
    await loadFailed.page.waitForFunction(() => document.getElementById('msg').textContent.includes('承認'));
    ck(await loadFailed.page.locator('#saveBtn').isDisabled(), '読み込めない状態で空の値を保存できる');
    await loadFailed.page.close();
    const failedSave = await openGasPage(browser, 'bni_stats.html', Object.assign({}, F, {
      saveBniStatsSettings: () => { throw new Error('保存の通信に失敗しました'); }
    }), { fails });
    await failedSave.page.waitForFunction(() => current && !busy);
    await failedSave.page.locator('#worldMembers').fill('399,999人');
    await failedSave.page.locator('#saveBtn').click();
    await failedSave.page.waitForFunction(() => document.getElementById('msg').textContent.includes('保存できませんでした'));
    ck(await failedSave.page.locator('#worldMembers').inputValue() === '399,999人' && await failedSave.page.locator('#saveBtn').isEnabled(), '保存の通信失敗で入力内容を消した・保存を再開できない');
    ck(F.bniStatsRecord_('2026-10').values.worldMembers === '362,111人超', '保存の通信失敗で保存済みの数値が変わった');
    await failedSave.page.close();
    await page.close();
  } finally { await browser.close(); }
}
browserChecks().then(() => {
  if (fails.length) { console.log('NG ' + fails.length + '/' + checks); console.log(fails.join('\n')); process.exit(1); }
  console.log('BNI統計情報: 検査 ' + checks + ' 件 OK');
}).catch(e => { console.log(fails.join('\n')); console.error(e); process.exit(1); });
