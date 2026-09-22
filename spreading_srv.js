// === Spreading（MCP）のデータをメンバー名簿へ取り込む ===
//
// SpreadingのMCPサーバー: https://active.spreading.tokyo/mcp_server*rpc
//   ※URLの「*」は誤字ではない。サーバーが公開している設定にもこの形で書かれている。
//
// 認証はOAuth 2.1（認可コード＋PKCE、スコープは mcp.read の読み取り専用）。
// 動的クライアント登録(DCR)には対応しておらず、client_id にURLを渡す
// CIMD方式を宣言している。GAS側から直接ログインする仕組みはまだ用意していないため、
// 当面はClaudeなどのMCP対応クライアントで取得した結果(JSON)を貼り付けて取り込む。
//
// 貼り付けるのは member_list の実行結果だけでよい。
// 各メンバーが category_groups（業種区分とその並び順）を持っているので、
// 業種区分も番号の並び順もこれ1つから再現できる。
//
// member_list の主なキー:
//   name         氏名
//   furigana     ふりがな
//   category     カテゴリー（業務内容）      → 名簿の「カテゴリー」列
//   company_name 会社名
//   position     チャプター内の役割           → 名簿の「役職」列
//   date_join    入会日
//   status       0=有効 / 1=無効
//   category_groups[] {category_id, category_name, sort, member_sort}
//
// 取得できないもの（Spreadingに項目が無い）:
//   写真ファイル名・一言コメント・紹介してほしい人・協業したい人・更新日・更新期限日
//   → これらは取り込み時に上書きせず、そのまま残す。

function openSpreadingDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('spreading').setWidth(820).setHeight(660),
    'Spreadingからメンバー名簿を更新');
}

// 貼り付けられた文字列からJSONを取り出す。
// 前後に説明文が付いていても拾えるように、最初の { か [ から最後の } か ] までを切り出す。
// MCPの応答は {"ok":true,"data":{"items":[...]}} だが、items配列だけ貼られても受け付ける。
function parseSpreadingItems_(text) {
  var s = String(text == null ? '' : text).trim();
  if (!s) return null;
  var first = s.search(/[\{\[]/);
  if (first < 0) return null;
  var last = Math.max(s.lastIndexOf('}'), s.lastIndexOf(']'));
  if (last <= first) return null;
  var obj = JSON.parse(s.substring(first, last + 1));
  if (Object.prototype.toString.call(obj) === '[object Array]') return obj;
  if (obj.items) return obj.items;
  if (obj.data && obj.data.items) return obj.data.items;
  return null;
}

// Spreadingのメンバー1件を、メンバー名簿の項目名に変換する
function mapSpreadingMember_(m) {
  var groups = m.category_groups || [];
  var g = groups.length ? groups[0] : null;
  return {
    name:     String(m.name || '').trim(),
    // ふりがなは全角スペース区切りで入っていることが多いので半角にそろえる
    kana:     String(m.furigana || '').replace(/　/g, ' ').trim(),
    title:    String(m.category || '').trim(),
    company:  String(m.company_name || '').trim(),
    role:     String(m.position || '').trim(),
    cat:      g ? String(g.category_name || '').trim() : '',
    joinDate: String(m.date_join || '').trim(),
    _groupSort:  g ? Number(g.sort || 0) : 9999,
    _memberSort: g ? Number(g.member_sort || 0) : 9999,
    _status:  Number(m.status || 0)
  };
}

// 業種区分の並び（グループのsort → グループ内のmember_sort）で 1 から番号を振る。
// 紙のメンバーリストと同じ並びになるはずだが、実際の番号と一致するかは
// 取り込み前のプレビューで必ず確認すること。
function assignSpreadingNumbers_(list) {
  list.sort(function (a, b) {
    if (a._groupSort !== b._groupSort) return a._groupSort - b._groupSort;
    return a._memberSort - b._memberSort;
  });
  for (var i = 0; i < list.length; i++) list[i].no = String(i + 1);
  return list;
}

// 取り込み内容を組み立てる。プレビューと実行の両方から使う。
function buildSpreadingPlan_(memberJson, renumber) {
  var items = parseSpreadingItems_(memberJson);
  if (!items || !items.length) {
    return { ok: false, message: 'メンバー一覧を読み取れませんでした。member_list の実行結果をそのまま貼り付けてください。' };
  }

  var incoming = [], skipped = 0;
  for (var i = 0; i < items.length; i++) {
    var m = mapSpreadingMember_(items[i]);
    if (!m.name) { skipped++; continue; }
    if (m._status !== 0) { skipped++; continue; }   // 無効メンバーは取り込まない
    incoming.push(m);
  }
  if (!incoming.length) return { ok: false, message: '有効なメンバーが1件も見つかりませんでした。' };
  if (renumber) assignSpreadingNumbers_(incoming);

  var cur = getMemberMaster().members || [];
  var byName = {};
  for (var j = 0; j < cur.length; j++) byName[normName_(cur[j].name)] = j;

  var fields = [['kana', 'ふりがな'], ['title', 'カテゴリー'], ['company', '会社名'],
                ['role', '役職'], ['cat', '業種区分'], ['joinDate', '入会日']];
  var plan = { ok: true, renumber: !!renumber, incoming: incoming, skipped: skipped,
               updates: [], adds: [], missing: [], unchanged: 0 };
  var matched = {};

  for (var k = 0; k < incoming.length; k++) {
    var e = incoming[k], idx = byName[normName_(e.name)];
    if (idx === undefined) { plan.adds.push({ no: e.no || '', name: e.name, cat: e.cat }); continue; }
    matched[idx] = true;
    var before = cur[idx], changes = [];
    for (var f = 0; f < fields.length; f++) {
      var key = fields[f][0];
      // Spreading側が空の項目は「消したい」ではなく「未入力」なので上書きしない
      if (e[key] && String(e[key]) !== String(before[key] || '')) {
        changes.push({ field: fields[f][1], from: String(before[key] || ''), to: String(e[key]) });
      }
    }
    if (renumber && e.no && String(e.no) !== String(before.no || '')) {
      changes.push({ field: 'No', from: String(before.no || ''), to: String(e.no) });
    }
    if (changes.length) plan.updates.push({ no: before.no || '', name: before.name, changes: changes });
    else plan.unchanged++;
  }

  // 名簿にはいるがSpreadingの有効メンバーに居ない人。退会者か氏名の表記ゆれ。
  // 消すと写真や一言コメントまで失われるので、報告するだけで自動削除はしない。
  for (var n = 0; n < cur.length; n++) {
    if (!matched[n]) plan.missing.push({ no: cur[n].no || '', name: cur[n].name });
  }
  return plan;
}

// 取り込み前の確認用。名簿は変更しない。
function previewSpreadingMembers(memberJson, renumber) {
  try {
    var plan = buildSpreadingPlan_(memberJson, renumber);
    if (!plan.ok) return plan;
    return {
      ok: true, preview: true,
      total: plan.incoming.length, unchanged: plan.unchanged, skipped: plan.skipped,
      updates: plan.updates, adds: plan.adds, missing: plan.missing
    };
  } catch (e) {
    console.error('[SPREADING] preview ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み取りに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 実際にメンバー名簿へ反映する。
// 写真・一言コメント・紹介してほしい人・協業したい人・メモ・更新日・更新期限日は触らない。
function importSpreadingMembers(memberJson, renumber) {
  try {
    var plan = buildSpreadingPlan_(memberJson, renumber);
    if (!plan.ok) return plan;

    var cur = getMemberMaster().members || [];
    var byName = {};
    for (var i = 0; i < cur.length; i++) byName[normName_(cur[i].name)] = i;

    var updated = 0, added = 0;
    for (var k = 0; k < plan.incoming.length; k++) {
      var e = plan.incoming[k], idx = byName[normName_(e.name)];
      if (idx === undefined) {
        cur.push({ no: e.no || '', cat: e.cat, name: e.name, kana: e.kana, title: e.title,
                   company: e.company, role: e.role, memo: '',
                   photoFile: '', comment: '', refer: '', collab: '',
                   joinDate: e.joinDate, renewDate: '', expireDate: '' });
        added++;
      } else {
        var m = cur[idx];
        if (e.kana) m.kana = e.kana;
        if (e.title) m.title = e.title;
        if (e.company) m.company = e.company;
        if (e.role) m.role = e.role;
        if (e.cat) m.cat = e.cat;
        if (e.joinDate) m.joinDate = e.joinDate;
        if (plan.renumber && e.no) m.no = String(e.no);
        updated++;
      }
    }

    // No順に並べ替える（番号なしは末尾）
    cur.sort(function (a, b) {
      var na = parseInt(a.no, 10), nb = parseInt(b.no, 10);
      if (isNaN(na) && isNaN(nb)) return 0;
      if (isNaN(na)) return 1;
      if (isNaN(nb)) return -1;
      return na - nb;
    });

    var res = saveMemberMaster(cur, null);
    if (!res.ok) return res;

    var msg = 'Spreadingの' + plan.incoming.length + '名を反映しました。'
            + '（更新 ' + updated + '名 / 新規 ' + added + '名 / 名簿は計 ' + cur.length + '名）';
    if (plan.missing.length) {
      msg += '\n\n名簿にあってSpreadingの有効メンバーに居ない方が ' + plan.missing.length + '名います。'
           + '退会された方か、氏名の書き方が違う可能性があります。自動では消していません:\n'
           + plan.missing.map(function (x) { return x.name; }).join('、');
    }
    msg += '\n\n写真・一言コメント・紹介してほしい人・協業したい人・更新期限日はそのまま残しています。';
    console.log('[SPREADING] updated=' + updated + ' added=' + added);
    return { ok: true, message: msg, updated: updated, added: added, total: cur.length };
  } catch (e) {
    console.error('[SPREADING] import ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '取り込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}
