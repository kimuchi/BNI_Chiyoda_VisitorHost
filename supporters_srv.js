// === スーパーサポーター（他チャプターのメンバー）===
//
// リージョン参加者として、他チャプターのメンバーがスーパーサポーターとして来られる（来られるかは回によって違う）。
// ⚙️ 設定 ＞ スーパーサポーター で、お名前・チャプター・会社名・カテゴリーを登録しておく。
// 写真は ⚙️ 設定 ＞ メンバー写真 に、メンバーと同じく「氏名.jpg」で入れる（名簿の方と同じように照合する）。
// 前半スライドを作る画面で、その回に来られる方を選ぶと、メンバーのウィークリープレゼンのあとにその方のページを入れる
// （初めは、ルーティンチェックシートの「リージョン参加者」に名字が書いてある方を選んでおく）。

var SUPPORTERS_PROP_ = 'BNI_SUPER_SUPPORTERS';
var SUPPORTERS_MAX_ = 30;

function superSupporters_() {
  var raw = PropertiesService.getScriptProperties().getProperty(SUPPORTERS_PROP_) || '[]', list = [];
  try { list = JSON.parse(raw) || []; } catch (e) { list = []; }
  return list.filter(function (s) { return s && s.name; });
}

// 画面用：登録した方と、写真があるか
function getSuperSupporters() {
  try {
    return { ok: true, supporters: superSupporters_().map(function (s) {
      var has = false;
      try { has = !!findPhotoIdForName_(s.name); } catch (e) {}
      return { name: s.name, chapter: s.chapter || '', company: s.company || '', category: s.category || '', hasPhoto: has };
    }) };
  } catch (e) {
    return { ok: false, message: 'スーパーサポーターを読めませんでした: ' + (e && e.message ? e.message : e), supporters: [] };
  }
}

// list … [{ name, chapter, company, category }]（並びのとおりにページを作る）
function saveSuperSupporters(list) {
  try {
    var clean = function (v, n) { return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, n); };
    var out = [], seen = {};
    (list || []).forEach(function (s) {
      var name = clean(s && s.name, 40).replace(/(さん|様)$/, '').trim();
      if (!name || seen[normName_(name)]) return;
      seen[normName_(name)] = true;
      out.push({ name: name, chapter: clean(s.chapter, 60), company: clean(s.company, 120), category: clean(s.category, 120) });
    });
    if (out.length > SUPPORTERS_MAX_) return { ok: false, message: 'スーパーサポーターは ' + SUPPORTERS_MAX_ + '名までです。' };
    // 名簿のメンバーと同じお名前は、写真の照合で取り違えるので止める
    var members = {};
    try { (getMemberMaster({ membersOnly: true }).members || []).forEach(function (m) { members[photoKey_(m.name)] = m.name; }); } catch (e) {}
    var dup = out.filter(function (s) { return members[photoKey_(s.name)]; }).map(function (s) { return s.name; });
    if (dup.length) return { ok: false, message: 'メンバー名簿に同じお名前の方がいます（' + dup.join('、') + '）。名簿のメンバーはスーパーサポーターに入れられません。' };
    PropertiesService.getScriptProperties().setProperty(SUPPORTERS_PROP_, JSON.stringify(out));
    var res = getSuperSupporters();
    res.message = 'スーパーサポーターを保存しました（' + out.length + '名）。';
    return res;
  } catch (e) {
    return { ok: false, message: '保存できませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// 写真の照合に使うお名前（名簿のメンバーと同じ扱いにする）
function superSupporterNames_() {
  try { return superSupporters_().map(function (s) { return s.name; }); } catch (e) { return []; }
}
