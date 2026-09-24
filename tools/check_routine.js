// ルーティンチェックシートの読み取りを、実物のデータで確かめる。
// 本番と同じ routine_srv.js を、SpreadsheetApp まわりだけ差し替えてNodeで動かす。
//
//   node tools/check_routine.js <routine.json>

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const DATA = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));

// --- シートを真似る ---
function fakeSheet(name, grid) {
  return {
    getName: () => name,
    getLastRow: () => grid.length,
    getLastColumn: () => (grid[0] ? grid[0].length : 0),
    getRange(r, c, nr, nc) {
      return {
        getValues() {
          const out = [];
          for (let i = 0; i < nr; i++) {
            const row = grid[r - 1 + i] || [];
            out.push(row.slice(c - 1, c - 1 + nc));
          }
          return out;
        },
      };
    },
  };
}
const sheets = Object.keys(DATA).map((n) => fakeSheet(n, DATA[n]));

// --- 名簿（検証用。実際の氏名は使わない）---
const MEMBERS = ['谷村 大輔', '竹田 明日翔', '本田 圭吾', '竹中 直人', '桒原 美穂',
                 '野崎 佳子', '豊田 恵', '高瀬 舟', '平井 成美', '舩山 ちひろ', '杉浦 太郎',
                 '山本 登一郎', '金子 高志', '瞳 ゆり', '田中 秀一', '小池 美咲',
                 '木村 光範', '若林 勇貴', '石渕 裕介', '分銅 雅一', '三澤 浩三', '上野 誠']
  .map((n) => ({ no: '1', name: n }));

const sandbox = {
  console,
  SpreadsheetApp: {},
  getSS_: () => ({ getSheets: () => sheets }),
  getMembersList: () => MEMBERS,
  normName_: (s) => String(s == null ? '' : s).replace(/[\s　]/g, ''),
  parseDate_(v) {
    if (!v) return null;
    let d;
    if (Object.prototype.toString.call(v) === '[object Date]') d = new Date(v.getTime());
    else {
      const s = String(v).trim().replace(/[年月]/g, '/').replace(/日/g, '');
      if (!s) return null;
      d = new Date(s);
    }
    if (isNaN(d.getTime())) return null;
    d.setHours(0, 0, 0, 0);
    return d;
  },
  fmtDate_(d) {
    if (!d) return '';
    const p = (n) => ('0' + n).slice(-2);
    return d.getFullYear() + '/' + p(d.getMonth() + 1) + '/' + p(d.getDate());
  },
  // コード.js の照合をそのまま持ってくる（姓だけでも当たるように）
  matchInviterToMember(inviterName, list) {
    if (!inviterName) return '';
    const s = String(inviterName).replace(/[\s　さん]/g, '');
    if (!s) return '';
    for (const m of list) {
      const t = m.name.replace(/[\s　]/g, '');
      if (t === s || t.startsWith(s) || s.startsWith(t)) return m.name;
    }
    return String(inviterName);
  },
};
sandbox.global = sandbox;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'routine_srv.js'), 'utf8'), sandbox, { filename: 'routine_srv.js' });

// --- 全開催日を通して読んでみる ---
let dates = [];
for (const name of Object.keys(DATA)) {
  const row1 = DATA[name][0] || [];
  row1.forEach((v) => { if (/^\d{4}\/\d{2}\/\d{2}$/.test(String(v))) dates.push(String(v)); });
}
dates = [...new Set(dates)].sort();

let ok = 0, noSheet = 0, withCore = 0, withPres = 0, unmatched = [], badCore = [];
for (const d of dates) {
  const r = sandbox.getRoutineInfo(d);
  if (!r.ok) { console.log('  読み取り失敗', d, r.message); continue; }
  if (!r.found) { noSheet++; continue; }
  ok++;
  if (r.coreValue) withCore++;
  else if (r.coreValueRaw && !/^(済|休会|なし|無し)/.test(r.coreValueRaw)) badCore.push([d, r.coreValueRaw]);
  if (r.longPresenter) withPres++;
  if (r.longPresenterUnmatched) unmatched.push([d, r.longPresenterRaw]);
}
console.log(`開催日 ${dates.length} 件 → 列が見つかった ${ok} 件・見つからない ${noSheet} 件`);
console.log(`  コアバリューを判別できた: ${withCore} 件`);
console.log(`  2分30秒プレゼンの方あり : ${withPres} 件`);
if (badCore.length) {
  console.log(`  判別できなかったコアバリュー ${badCore.length} 件:`);
  badCore.slice(0, 8).forEach(([d, v]) => console.log(`    ${d}  ${JSON.stringify(v).slice(0, 60)}`));
}
if (unmatched.length) {
  console.log(`  名簿と一致しなかった方 ${unmatched.length} 件（検証用の名簿には少ししか入れていないため）:`);
  unmatched.slice(0, 8).forEach(([d, v]) => console.log(`    ${d}  ${v}`));
}

// --- スライドに載せる項目 ---
console.log('\nスライドに載せる項目:');
for (const d of ['2026/09/09', '2026/09/16', '2026/09/23', '2026/09/30']) {
  const r = sandbox.getRoutineInfo(d);
  if (!r.found) { console.log(`  ${d} (該当なし)`); continue; }
  const mp = (r.mainPresenters || []).map((m) => `${m.raw}→${m.name || '(未一致)'}`).join(' / ');
  console.log(`  ${d}`);
  console.log(`     メインプレゼン : ${mp || '(なし)'}`);
  console.log(`     一般規定       : ${r.generalPolicy || '(なし)'}番  ［${r.generalPolicyRaw}］`);
  console.log(`     求める専門分野 : ${JSON.stringify(r.wantedCategories)}`);
  console.log(`     開放カテゴリー : ${r.openCategory || '(空欄)'}`);
  console.log(`     審査中         : ${r.reviewCategory || '(空欄)'}`);
}

// --- 決まった日で中身を確かめる ---
const samples = ['2026/04/01', '2026/04/08', '2026/09/30', '2026/10/07'];
console.log('\n抜き取り:');
for (const d of samples) {
  const r = sandbox.getRoutineInfo(d);
  console.log(`  ${d}  ${r.found ? r.sheetName : '(該当なし)'}`
    + (r.found ? `  第${r.meetingNo}回  コアバリュー=${r.coreValue || '(不明)'} `
               + `${r.coreValueRaw ? '［' + String(r.coreValueRaw).replace(/\s+/g, ' ').slice(0, 30) + '］' : ''}`
               + `  2分30秒=${r.longPresenter || (r.longPresenterRaw ? r.longPresenterRaw + '(未一致)' : 'なし')}` : ''));
}
