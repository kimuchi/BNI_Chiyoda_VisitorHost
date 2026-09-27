// Apps Script の API を、検査に要るぶんだけ Node で真似る。
//
//   const { makeGas } = require('./lib_gas_fake');
//   const gas = makeGas({ files: { OFFICIAL: '/path/official.pptx' } });   // Drive のファイルID → 手元のファイル
//   vm.createContext(Object.assign(sandbox, gas.globals));
//
// Blob は Buffer で持つ。getBytes() / getContent() は Apps Script と同じく符号付きの数の配列を返す。
// Drive API の範囲指定のダウンロード（Range: bytes=a-b）は、手元のファイルのその範囲を返す（206）。
// Blob が 50MB を超えると止める（Apps Script の上限）。範囲の指定が無いと、ファイル全体（200）。

const fs = require('fs');
const zlib = require('zlib');
const { readZip, writeZip } = require('./lib_zip');

const LIMIT = 50 * 1024 * 1024;

class FakeBlob {
  constructor(buf, type, name) {
    this._buf = buf;
    this._type = type || '';
    this._name = name || '';
    if (buf.length > LIMIT) throw new Error('Blob が 50MB を超えました（' + buf.length + '）: ' + name);
  }
  getBytes() { return signed(this._buf); }
  getDataAsString() { return this._buf.toString('utf8'); }
  getContentType() { return this._type; }
  setContentType(t) { this._type = t; return this; }
  getName() { return this._name; }
  setName(n) { this._name = n; return this; }
  copyBlob() { return new FakeBlob(Buffer.from(this._buf), this._type, this._name); }
}
function signed(buf) {
  const out = new Array(buf.length);
  for (let i = 0; i < buf.length; i++) out[i] = buf[i] > 127 ? buf[i] - 256 : buf[i];
  return out;
}
function toBuf(data) {
  if (Buffer.isBuffer(data)) return data;
  if (typeof data === 'string') return Buffer.from(data, 'utf8');
  if (Array.isArray(data)) return Buffer.from(data.map((b) => b & 255));
  return Buffer.from(data);
}

function makeGas(opt) {
  opt = opt || {};
  const files = opt.files || {};          // ID → 手元のパス
  const props = opt.props || {};
  const log = { fetches: 0, fetchedBytes: 0, created: [] };
  const Utilities = {
    newBlob: (data, type, name) => new FakeBlob(toBuf(data == null ? '' : data), type, name),
    ungzip: (blob) => new FakeBlob(zlib.gunzipSync(blob._buf), '', ''),
    gzip: (blob) => new FakeBlob(zlib.gzipSync(blob._buf), 'application/x-gzip', blob.getName() + '.gz'),
    zip: (blobs, name) => {
      const f = {};
      blobs.forEach((b) => { f[b.getName()] = b._buf; });
      return new FakeBlob(writeZip(f), 'application/zip', name);
    },
    unzip: (blob) => {
      const f = readZip(blob._buf);
      return Object.keys(f).map((n) => new FakeBlob(f[n], '', n));
    },
    formatDate: (d, tz, fmt) => {
      const p = (n) => (n < 10 ? '0' : '') + n;
      return String(fmt).replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate()))
        .replace('HH', p(d.getHours())).replace('mm', p(d.getMinutes()));
    },
    base64Decode: (s) => signed(Buffer.from(String(s), 'base64')),
    base64Encode: (b) => Buffer.from(Array.isArray(b) ? b.map((x) => x & 255) : b).toString('base64'),
  };
  const sizeOf = (id) => fs.statSync(files[id]).size;
  const readRange = (id, a, b) => {
    const fd = fs.openSync(files[id], 'r'), len = b - a + 1, buf = Buffer.alloc(len);
    fs.readSync(fd, buf, 0, len, a);
    fs.closeSync(fd);
    return buf;
  };
  const response = (code, buf) => ({
    getResponseCode: () => code,
    getContent: () => signed(buf),
    getBlob: () => new FakeBlob(buf, 'application/octet-stream', ''),
    getContentText: () => buf.length < 2000 ? buf.toString('utf8') : '',
  });
  const UrlFetchApp = {
    fetch(url, o) {
      const m = String(url).match(/\/drive\/v3\/files\/([^?]+)\?alt=media/);
      if (!m || !files[decodeURIComponent(m[1])]) return response(404, Buffer.from('not found'));
      const id = decodeURIComponent(m[1]), size = sizeOf(id);
      const auth = o && o.headers && o.headers.Authorization;
      if (!/^Bearer /.test(auth || '')) return response(401, Buffer.from('no auth'));
      const r = o && o.headers && o.headers.Range && o.headers.Range.match(/^bytes=(\d+)-(\d+)$/);
      log.fetches++;
      if (!r) {
        if (size > LIMIT) throw new Error('範囲を指定せずに 50MB を超えるファイルを取ろうとしました');
        log.fetchedBytes += size;
        return response(200, fs.readFileSync(files[id]));
      }
      const a = +r[1], b = Math.min(+r[2], size - 1);
      if (b - a + 1 > LIMIT) throw new Error('1回に 50MB を超える範囲を取ろうとしました');
      log.fetchedBytes += b - a + 1;
      return response(206, readRange(id, a, b));
    },
    fetchAll(reqs) { return reqs.map((q) => this.fetch(q.url, q)); },
  };
  const fileObj = (id) => ({
    getId: () => id, getName: () => require('path').basename(files[id]), getSize: () => sizeOf(id),
    getMimeType: () => 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    getUrl: () => 'https://drive.google.com/file/d/' + id + '/view',
    getBlob: () => new FakeBlob(fs.readFileSync(files[id]), 'application/zip', require('path').basename(files[id])),
  });
  const DriveApp = {
    getFileById: (id) => { if (!files[id]) throw new Error('ファイルがありません: ' + id); return fileObj(id); },
  };
  const PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null),
      setProperty: (k, v) => { props[k] = String(v); },
      deleteProperty: (k) => { delete props[k]; },
    }),
  };
  const ScriptApp = { getOAuthToken: () => 'test-token' };
  return { globals: { Utilities, UrlFetchApp, DriveApp, PropertiesService, ScriptApp, console }, log, props, FakeBlob };
}

module.exports = { makeGas, FakeBlob, signed };
