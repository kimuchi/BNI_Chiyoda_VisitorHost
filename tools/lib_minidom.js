// 画面のHTMLを、ブラウザ無しでNode上で動かすための簡易DOM。
// id の付いた部品（input・select・textarea など）だけを持ち、innerHTML で差し込まれた
// 部品も拾う。google.script.run はサーバーの代わりの関数表に置き換える。
//
// 構文は正しくても、ボタンを押したときに初めて出るエラー（関数が無い・値の渡し忘れ・
// 共通部品の読み込み漏れ）を見つけるために使う。
//
//   const { loadPage } = require('./lib_minidom');
//   const page = loadPage('slides_meeting_first.html', { server: { getX: () => ({ ok: true }) } });
//   page.step('開く', () => page.window.onload());

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

function attr(tag, name) {
  const m = tag.match(new RegExp('\\s' + name + '="([^"]*)"'));
  return m ? m[1] : null;
}
function unesc(s) {
  return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'").replace(/&amp;/g, '&');
}
function escHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// 文字の幅は全角1・半角0.5で測る（canvasの代わり）。
// 描いた文字（fillText）は texts に残す（画像に何を描いたかを確かめるため）
function fakeCtx(texts) {
  let size = 44;
  const noop = () => {};
  return {
    set font(v) { const m = /(\d+(?:\.\d+)?)px/.exec(v); size = m ? parseFloat(m[1]) : 44; },
    get font() { return size + 'px'; },
    measureText(t) {
      let w = 0;
      for (const ch of String(t)) w += ch.charCodeAt(0) < 128 ? 0.5 : 1.0;
      return { width: w * size };
    },
    fillText(t, x, y) { if (texts) texts.push({ text: String(t), size, x, y, color: this.fillStyle }); },
    scale: noop, fillRect: noop, strokeRect: noop, save: noop, restore: noop, beginPath: noop, rect: noop, clip: noop,
    fillStyle: '', strokeStyle: '', lineWidth: 1, textAlign: '', textBaseline: '',
  };
}
// canvas の代わり。toDataURL は、描いた文字と大きさをJSONにしたものを base64 で返す
function fakeCanvas() {
  const texts = [];
  const cv = { width: 0, height: 0, _texts: texts };
  cv.getContext = () => fakeCtx(texts);
  cv.toDataURL = () => 'data:image/png;base64,'
    + Buffer.from(JSON.stringify({ width: cv.width, height: cv.height, texts })).toString('base64');
  return cv;
}

function loadPage(file, opts) {
  const o = opts || {};
  const server = o.server || {};
  const fails = o.fails || [];
  const els = {};
  const log = { alerts: [], confirms: [], copies: 0, calls: [], downloads: [] };

  function parseOptions(el, html) {
    el.options = [];
    const re = /<option\b([^>]*)>([\s\S]*?)<\/option>/g;
    let m, sel = 0;
    while ((m = re.exec(html)) !== null) {
      if (/\sselected/.test(m[1])) sel = el.options.length;
      const v = attr('<o ' + m[1] + '>', 'value');
      el.options.push({ value: unesc(v === null ? m[2] : v), text: unesc(m[2]) });
    }
    el.selectedIndex = el.options.length ? sel : -1;
  }
  function makeEl(id, tag) {
    const el = { id, tag, checked: false, disabled: false, style: {}, textContent: '', innerText: '',
                 _html: '', _value: '', options: [], selectedIndex: -1,
                 focus() {}, select() { log.selected = this.id; }, appendChild(c) { return c; }, removeChild(c) { return c; } };
    Object.defineProperty(el, 'innerHTML', {
      get() { return this._html; },
      set(v) {
        this._html = String(v);
        if (this.tag === 'select') parseOptions(this, this._html);
        else scan(this._html);
      },
    });
    Object.defineProperty(el, 'value', {
      get() {
        if (this.tag !== 'select') return this._value;
        return this.selectedIndex >= 0 && this.options[this.selectedIndex] ? this.options[this.selectedIndex].value : '';
      },
      set(v) {
        if (this.tag !== 'select') { this._value = v == null ? '' : String(v); return; }
        this.selectedIndex = this.options.findIndex((x) => x.value === String(v));   // 無い値なら -1（ブラウザと同じ）
      },
    });
    return el;
  }
  function scan(html) {
    const re = /<(input|select|textarea|div|span|table|button|a|label|code|img|b|strong|p|pre|td|tr|ul|li|section|h[1-6])\b([^>]*)>/g;
    let m;
    while ((m = re.exec(html)) !== null) {
      const id = attr(m[0], 'id');
      if (!id) continue;
      const el = makeEl(id, m[1]);
      if (m[1] === 'input') {
        el.type = attr(m[0], 'type') || 'text';
        el.checked = /\schecked(?=[\s>\/])/.test(m[0]);
        el.value = unesc(attr(m[0], 'value') || '');
      }
      if (m[1] === 'select') {
        const end = html.indexOf('</select>', m.index);
        parseOptions(el, html.substring(m.index, end));
      }
      if (m[1] === 'textarea') {
        const end = html.indexOf('</textarea>', m.index);
        el.value = unesc(html.substring(m.index + m[0].length, end));
      }
      els[id] = el;
    }
  }

  // HTMLを読む（<?!= include('…') ?> はここで展開する）。
  // ほかのスクリプトレット（サーバーから渡す値）は、o.preprocess で本番の結果に置き換えてから渡す
  let page = fs.readFileSync(path.join(ROOT, file), 'utf8');
  page = page.replace(/<\?!=\s*include\('([^']+)'\)\s*\?>/g,
    (_, n) => fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8'));
  if (o.preprocess) page = o.preprocess(page);
  if (/<\?/.test(page)) fails.push(file + ': スクリプトレットが残っている');
  scan(page.replace(/<script[^>]*>[\s\S]*?<\/script>/g, ''));
  const js = [...page.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]).join('\n');

  const document = {
    getElementById: (id) => els[id] || null,
    createElement: (tag) => {
      if (tag === 'canvas') return fakeCanvas();
      const x = { textContent: '', style: {}, dataset: {}, className: '', appendChild(c) { return c; },
                  click() { if (this.download) log.downloads.push({ name: this.download, href: this.href }); } };
      // 作った部品に innerHTML で入れた id 付きの部品（カードの中の入力欄など）も拾う
      Object.defineProperty(x, 'innerHTML', { get() { return this._html != null ? this._html : escHtml(this.textContent); },
                                              set(v) { this._html = String(v); scan(this._html); } });
      return x;
    },
    body: { appendChild() {}, removeChild() {} },
    // 「button」のようにタグ名だけのときは、id の付いたその部品を返す（ほかは空）
    querySelectorAll: (sel) => (/^[a-z]+$/.test(sel) ? Object.values(els).filter((e) => e.tag === sel) : []),
    execCommand: (cmd) => { if (cmd === 'copy') log.copies++; return true; },
  };

  // google.script.run の代わり。返事は flush() で順に届ける（本物と同じく後から届く）
  const queue = [];
  function runner() {
    let ok = null, ng = null;
    const r = new Proxy({}, {
      get(_, name) {
        if (name === 'withSuccessHandler') return (f) => { ok = f; return r; };
        if (name === 'withFailureHandler') return (f) => { ng = f; return r; };
        return (...args) => {
          log.calls.push({ name: String(name), args });
          if (!server[name]) { fails.push('サーバーに無い関数を呼んでいる: ' + String(name)); return; }
          let res;
          // サーバーで例外が出たときは、本物と同じく withFailureHandler に届ける（無ければそのまま投げる）
          try { res = server[name](...args); }
          catch (e) { if (ng) { const h = ng; queue.push(() => h(e)); return; } throw e; }
          if (ok) queue.push(() => ok(res));
        };
      },
    });
    return r;
  }
  function flush() { while (queue.length) queue.shift()(); }
  // 返事を1つだけ届ける（読み込みの途中の画面を確かめるため）。届けたら true
  function flushOne() { if (!queue.length) return false; queue.shift()(); return true; }

  const window = {};
  // o.now … 画面の「きょう」（'2026-09-30T10:00:00' など）。無ければ本当のきょう
  const RealDate = Date;
  const PageDate = o.now ? class extends RealDate { constructor(...a) { if (a.length) super(...a); else super(new RealDate(o.now).getTime()); } } : RealDate;
  const sandbox = { window, document, console, Date: PageDate,
                    navigator: {},
                    alert: (m) => { log.alerts.push(m); },
                    confirm: (m) => { log.confirms.push(m); return true; },
                    google: { script: { get run() { return runner(); } } } };
  vm.createContext(sandbox);
  vm.runInContext(js, sandbox, { filename: file });

  const run = (code) => vm.runInContext(code, sandbox);
  function step(label, fn) {
    try { fn(); flush(); } catch (e) {
      fails.push(label + ' でエラー: ' + (e && e.stack ? e.stack.split('\n').slice(0, 3).join(' / ') : e));
    }
  }
  return { els, document, window, sandbox, run, step, flush, flushOne, pending: () => queue.length, log, fails };
}

module.exports = { loadPage };
