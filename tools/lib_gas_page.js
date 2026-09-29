// 画面のHTMLを本物のブラウザ（Chromium）で開き、google.script.run を Node 側のサーバー（vm で動かした本番のコード）につなぐ。
// check_csv_dialog.js・check_email_dialog.js で使う。
//
//   const { launch, openGasPage } = require('./lib_gas_page');
//   const browser = await launch();
//   const { page, dialogs, calls } = await openGasPage(browser, 'dialog.html', srv, { fails });
//
// サーバーの関数の戻り値は JSON を通して渡す（本物と同じく、関数や Date はそのままでは届かない）。
// 例外は withFailureHandler に Error として届く。alert・confirm は dialogs に残し、OK を押す。

const fs = require('fs');
const path = require('path');
const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();

const ROOT = path.join(__dirname, '..');

function launch() { return pw.chromium.launch(); }

async function openGasPage(browser, file, srv, opt) {
  const o = opt || {};
  const fails = o.fails || [];
  const calls = o.calls || [];
  const page = await browser.newPage();
  const dialogs = [];
  page.on('dialog', async (d) => { dialogs.push(d.type() + ':' + d.message()); await d.accept(); });
  page.on('pageerror', (e) => fails.push(file + ' の画面のエラー: ' + e.message));
  await page.exposeFunction('__gas', (name, argsJson) => {
    const args = JSON.parse(argsJson);
    calls.push({ name, args });
    try {
      if (typeof srv[name] !== 'function') return JSON.stringify({ ok: false, message: 'サーバーに無い関数: ' + name });
      const r = srv[name](...args);
      return JSON.stringify({ ok: true, r: r === undefined ? null : r });
    } catch (e) { return JSON.stringify({ ok: false, message: String(e && e.message ? e.message : e) }); }
  });
  await page.addInitScript(() => {
    function runner() {
      let ok = null, ng = null;
      const r = new Proxy({}, { get(_, name) {
        if (name === 'withSuccessHandler') return (f) => { ok = f; return r; };
        if (name === 'withFailureHandler') return (f) => { ng = f; return r; };
        if (name === 'withUserObject') return () => r;
        return (...args) => {
          window.__gas(String(name), JSON.stringify(args)).then((s) => {
            const x = JSON.parse(s);
            if (x.ok) { if (ok) ok(x.r); } else if (ng) ng(new Error(x.message));
          });
        };
      } });
      return r;
    }
    window.google = { script: { get run() { return runner(); } } };
  });
  const url = 'https://' + file.replace(/\W/g, '-') + '.test/';
  await page.route(url, (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: fs.readFileSync(path.join(ROOT, file), 'utf8') }));
  await page.goto(url, { waitUntil: 'load' });
  return { page, dialogs, calls };
}

module.exports = { launch, openGasPage };
