// === 会社名の改行（どの出力でも同じ切れ目で折る）===
// メンバーのページ・リファーラル発表・ビジタープレゼン・新メンバー・ネットワーキングリーダー・役職のメンバー紹介・
// メンバーブック・ビジター様リストのPDFなど、会社名を入れるところは、どれもこの決め方で行を分ける。
// 以前は出力ごとに別の決め方（ブラウザで字を測って詰めて折る・PowerPoint に任せる・字の種類の変わり目など）だったため、
// 「一般社団法人日本見／本協会」「株式会社見本コン／サルティング」のように語の途中で折れ、同じ会社でも出力ごとに折れ方が違った。
//
// 決め方：
//   ・切れ目の良さを、字の並びだけで決める（どこで使っても同じになるように。幅や字の大きさには左右されない）
//       0 法人の種類の前後（「株式会社／見本コンサルティング」「見本テクノロジーズ／株式会社」）・空白
//       1 かっこの前後・「・」「／」のあと・英字と日本語の変わり目
//       2 よく使う語の切れ目（「不動産／コンサルティング」「見本総合／法律事務所」）・漢字とカタカナの変わり目
//       4〜 漢字の続きの途中（2字ずつの切れ目を少しよしとする）・カタカナ語の途中（なるべく使わない）
//       使わない：法人の種類・英単語・数字の途中、行の頭に来てはいけない字（ー・ッ・）・、など）の前、行の終わりに来てはいけない字（（「など）のあと
//   ・枠に入る行の分け方のうち、切れ目の良さ＋行の長さの釣り合いがいちばん良いものを採る。
//     字を少し小さくすれば良い切れ目で入るときは、語の途中で折るより小さくする
//   ・字の幅は、どこでも同じ見積もり（全角1字＝1、英数字は字ごとの幅）で測る。PCの書体に左右されない
//
// サーバー（GAS）では CB_ をそのまま使う。画面（ブラウザ）には include('company_break') で、この関数の中身を
// そのまま渡す（webapp_srv.js の include）。どちらも同じ計算になる。
function companyBreakLib_() {
  // 法人の種類（長いものを先に見る）
  var CORP = ['特定非営利活動法人', '社会保険労務士法人', '土地家屋調査士法人', '有限責任事業組合', '一般社団法人', '一般財団法人',
              '公益社団法人', '公益財団法人', '社会福祉法人', '司法書士法人', '行政書士法人', '医療法人社団', '医療法人財団',
              '弁護士法人', '税理士法人', '弁理士法人', '監査法人', '医療法人', '学校法人', '宗教法人', '農事組合法人', 'NPO法人',
              '株式会社', '有限会社', '合同会社', '合資会社', '合名会社'];
  // 漢字の続きの中の語（この語の前・うしろで折れる。会社名によく出るもの）
  var KANJI_WORDS = ['法律事務所', '会計事務所', '特許事務所', '設計事務所', '税理士事務所', '司法書士事務所', '行政書士事務所',
                     '社労士事務所', '社会保険労務士', '土地家屋調査士', '一級建築士', '二級建築士', '建築士', '税理士', '司法書士',
                     '行政書士', '社労士', '弁護士', '弁理士', '会計士', '事務所', '研究所', '相談所', '製作所', '工務店', '代理店',
                     '営業所', '出張所', '事業所', '事業部', '協議会', '連合会', '動物病院', '整骨院', '接骨院', '鍼灸院', '治療院',
                     '美容室', '美容院', '不動産', '総合', '法律', '会計', '協会', '工業', '建設', '建築', '設計', '商事', '商会', '商店',
                     '産業', '興業', '企画', '技研', '電気', '電機', '設備', '運輸', '物流', '印刷', '製作', '製造', '販売', '保険',
                     '組合', '本店', '本社', '支社', '支店', '医院', '歯科', '病院', '薬局', '教室', '学院', '学園', '食品', '酒造',
                     '日本', '東京', '大阪', '国際', '情報', '通信', '支援', '環境', '開発', '管理', '経営', '労務', '人事', '事務',
                     '貿易', '物産', '工房', '写真', '映像', '広告', '出版', '旅行', '住宅', '住設', '塗装', '内装', '電設', '測量'];
  // カタカナの続きの中の語（この語の前・うしろで折れる）
  var KATA_WORDS = ['インターナショナル', 'コミュニケーションズ', 'コミュニケーション', 'ソリューションズ', 'ソリューション',
                    'エンジニアリング', 'ホールディングス', 'コーポレーション', 'プロダクションズ', 'プロダクション', 'コンサルティング',
                    'コンサルタンツ', 'コンサルタント', 'マーケティング', 'テクノロジーズ', 'テクノロジー', 'トレーディング',
                    'マネジメント', 'クリエイティブ', 'ファイナンシャル', 'ファイナンス', 'イノベーション', 'リノベーション',
                    'インベストメント', 'エンタープライズ', 'エンターテイメント', 'エンタテインメント', 'パートナーズ', 'パートナー',
                    'プランニング', 'ネットワーク', 'エージェンシー', 'エージェント', 'アドバイザーズ', 'アドバイザー', 'ウェルネス',
                    'ビューティー', 'ビューティ', 'システムズ', 'システム', 'サービシーズ', 'サービス', 'デザイン', 'オフィス',
                    'センター', 'スタジオ', 'ワークス', 'ファーム', 'グループ', 'ジャパン', 'カンパニー', 'ライフ', 'ホームズ',
                    'ホーム', 'ハウス', 'ケア', 'サポート', 'クリニック', 'アカデミー', 'スクール', 'メディア', 'ラボ', 'リサーチ',
                    'ビジネス', 'ソフトウェア', 'ソフト', 'テック', 'エステ', 'サロン', 'スペース', 'プロジェクト', 'リンク',
                    'キャピタル', 'ベンチャーズ', 'フーズ', 'フード', 'ダイニング', 'レストラン', 'カフェ', 'ショップ', 'ストア',
                    'トラベル', 'モーターズ', 'ホテル', 'リゾート', 'リフォーム', 'ハウジング', 'エステート', 'アセット', 'プロパティ',
                    'デジタル', 'クリエイト', 'プロ', 'ネット', 'ウェブ', 'アート', 'フォト', 'ミュージック', 'ファクトリー',
                    'インダストリー', 'エナジー', 'エネルギー', 'ロジスティクス', 'トータル', 'グローバル', 'ワールド', 'アソシエイツ',
                    'ドットコム', 'ジム', 'フィットネス', 'ヨガ', 'スポーツ', 'ペット', 'ベーカリー', 'キッチン', 'ガーデン'];
  var NO_START = /[ー‐゠－~〜・･、。，．,.：:；;！？!?）)」』】〕〉》］\]｝}々ゝゞヽヾぁぃぅぇぉっゃゅょゎゕゖァィゥェォッャュョヮヵヶ％%]/;
  var NO_END = /[（(「『【〔〈《［\[｛{]/;
  var PEN = [0, 1, 2, 3, 5, 7, 10];     // 切れ目の良さ（0〜6）→ 減点。-1（法人の種類の前後）は少し加点
  var pen = function (r) { return r < 0 ? -1 : PEN[r]; };
  var MARGIN = 1.05;                     // 字の幅の見積もりの余裕
  var ORPHAN = 8;                        // 1〜2字だけの行の減点
  var SIZE_W = 20;                       // 字を小さくしたときの減点（元の大きさからの割合 × これ）
  var BAL_W = 1.5;                       // 行の長さの釣り合いの減点の重さ

  // 名簿の文字をそろえる：セルの中の改行・タブを外し（英数字どうしのあいだだけ空白に）、空白を1つにまとめ、㈱などを開く
  function normalize(raw) {
    var t = String(raw == null ? '' : raw);
    t = t.replace(/[ \t\u3000]*[\r\n\u000B\u000C\u0085\u2028\u2029]+[ \t\u3000]*/g, function (m, at, all) {
      return /[A-Za-z0-9]/.test(all.charAt(at - 1)) && /[A-Za-z0-9]/.test(all.charAt(at + m.length)) ? ' ' : '';
    });
    t = t.replace(/\t/g, ' ').replace(/[ \u3000]{2,}/g, function (m) { return m.indexOf('\u3000') >= 0 ? '\u3000' : ' '; });
    return t.replace(/㈱/g, '株式会社').replace(/㈲/g, '有限会社').replace(/㈾/g, '合資会社').replace(/^[ \u3000]+|[ \u3000]+$/g, '');
  }

  // 1字の幅（全角1字＝1）。英数字は太字の欧文書体くらいの幅で、少し広めに見積もる
  function charEm(c) {
    var k = c.charCodeAt(0);
    if (k >= 0xFF61 && k <= 0xFF9F) return 0.55;                    // 半角カナ
    if (k >= 128) return 1;
    if (c === ' ') return 0.3;
    if (/[il.,:;|!'`]/.test(c)) return 0.3;
    if (/[fjrtI()\[\]\-\/]/.test(c)) return 0.42;
    if (/[mw]/.test(c)) return 0.92;
    if (/[MW]/.test(c)) return 1;
    if (/[A-Z&@%#]/.test(c)) return 0.74;
    return 0.6;                                                    // 小文字・数字・そのほか
  }
  function widthEm(s) {
    s = String(s == null ? '' : s);
    var w = 0;
    for (var i = 0; i < s.length; i++) w += charEm(s.charAt(i));
    return w;
  }
  function fits(line, pt, widthPt) { return widthEm(line) * pt * MARGIN <= widthPt + 1e-6; }

  function kind(c) {
    if (!c) return '';
    if (/[ \u3000]/.test(c)) return 'sp';
    if (/[A-Za-z0-9Ａ-Ｚａ-ｚ０-９]/.test(c)) return 'lat';
    if (/[一-鿿㐀-䶿々〆ヶ]/.test(c)) return 'kan';
    if (/[ァ-ヺーヽヾｦ-ﾟ]/.test(c)) return 'kat';
    if (/[ぁ-ゖゝゞ]/.test(c)) return 'hir';
    return 'pun';
  }
  // 全角の英数字を半角に（字数は変えない。法人の種類を探すときだけ使う）
  function fold(s) {
    return s.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  }
  // 続き（同じ種類の字の並び）の中の語を探す。見つかった語の [始まり, 終わり) の一覧（長い語を先に・重ならないように）
  function wordsIn(s, from, to, list) {
    var out = [], used = {};
    for (var w = 0; w < list.length; w++) {
      var word = list[w], at = s.indexOf(word, from);
      while (at >= 0 && at + word.length <= to) {
        var free = true;
        for (var q = at; q < at + word.length; q++) if (used[q]) { free = false; break; }
        if (free) {
          for (q = at; q < at + word.length; q++) used[q] = true;
          out.push([at, at + word.length]);
        }
        at = s.indexOf(word, at + 1);
      }
    }
    return out;
  }

  // 切れ目ごとの良さ。r[i] … i字目の前で折るときの良さ（0〜6。null は折らない）。空白の位置で折ると、空白は消える
  function cutRanks(text, opt) {
    var s = String(text), n = s.length, r = [], i, f = fold(s);
    for (i = 0; i <= n; i++) r.push(null);
    var lock = {}, force = {};
    // 法人の種類：中では折らない。前後は良い切れ目
    var used = {};
    for (var c = 0; c < CORP.length; c++) {
      var at = f.indexOf(CORP[c]);
      while (at >= 0) {
        var e = at + CORP[c].length, free = true;
        for (var q = at; q < e; q++) if (used[q]) { free = false; break; }
        if (free) {
          for (q = at; q < e; q++) used[q] = true;
          for (q = at + 1; q < e; q++) lock[q] = true;
          if (e < n) force[e] = -1;                                     // うしろ（「株式会社／見本…」「…株式会社／東京支社」）
          if (at > 0 && force[at] == null) force[at] = e >= n ? -1 : 1; // 前（「見本…／株式会社」）
        }
        at = f.indexOf(CORP[c], at + 1);
      }
    }
    // 続き（カタカナ・漢字）の中の語の切れ目
    var wordCut = {}, wordIn = {};
    var runs = function (k, list) {
      for (var a = 0; a < n;) {
        if (kind(s.charAt(a)) !== k) { a++; continue; }
        var b = a;
        while (b < n && kind(s.charAt(b)) === k) b++;
        wordsIn(s, a, b, list).forEach(function (x) {
          if (x[0] > a) wordCut[x[0]] = true;
          if (x[1] < b) wordCut[x[1]] = true;
          for (var p = x[0] + 1; p < x[1]; p++) wordIn[p] = true;
        });
        a = b;
      }
    };
    runs('kan', KANJI_WORDS);
    runs('kat', KATA_WORDS);
    for (i = 1; i < n; i++) {
      var a = s.charAt(i - 1), b = s.charAt(i), ka = kind(a), kb = kind(b), v;
      if (ka === 'sp') continue;                                         // 空白のうしろ（空白の前で折る）
      if (kb === 'sp') { r[i] = force[i] != null && force[i] < 0 ? -1 : 0; continue; }
      if (lock[i]) continue;
      if (NO_START.test(b) || NO_END.test(a)) continue;
      if (force[i] != null) { r[i] = force[i]; continue; }
      if (ka === 'lat' && kb === 'lat') continue;                        // 英単語・数字の途中
      if (/[・･／\/]/.test(a)) v = 1;
      else if (/[（(「『【〔［]/.test(b) || /[）)」』】〕］]/.test(a)) v = 1;
      // opt.latinJoin（カテゴリー）：「DX推進」「IT導入」のような英字＋日本語の語は、「（」の前・「・」のあとより折りにくく
      else if (ka === 'lat' || kb === 'lat') v = opt && opt.latinJoin && ka === 'lat' && (kb === 'kan' || kb === 'kat') ? 2 : 1;
      else if (/[＆&＋+]/.test(a)) v = 2;
      else if (/[＆&＋+]/.test(b)) v = 3;
      else if (ka !== kb) {
        if ((ka === 'kan' && kb === 'kat') || (ka === 'kat' && kb === 'kan')) v = 2;
        else if (ka === 'hir' && (kb === 'kan' || kb === 'kat')) v = 2;
        else if (ka === 'kat' && kb === 'hir') v = 3;
        else if (ka === 'kan' && kb === 'hir') v = 4;
        else v = 3;
      } else if (ka === 'kan' || ka === 'kat') {
        if (wordCut[i] && !wordIn[i]) v = 2;
        else if (wordIn[i]) v = 6;
        else if (ka === 'kan') {
          // 漢字は2字ずつの語が多い。数えはじめは、漢字の続きの頭・法人の種類のうしろ・語の切れ目から
          var st = i - 1;
          while (st > 0 && kind(s.charAt(st - 1)) === 'kan' && force[st] == null && !wordCut[st]) st--;
          v = (i - st) % 2 === 0 ? 4 : 5;
        } else v = 6;
      } else v = 6;                                                      // ひらがなの続きなど
      r[i] = v;
    }
    return r;
  }

  // 行の中身（前後の空白を外す）
  function piece(s, a, b) { return s.slice(a, b).replace(/^[ \u3000]+|[ \u3000]+$/g, ''); }
  function orphan(t) { return t.replace(/[ \u3000]/g, '').length <= 2 ? ORPHAN : 0; }

  // n 行に分ける切れ目のうち、いちばん良いもの。widths[k] … k 行目の幅（pt）。入らなければ null
  //   { cuts: [i…], lines: [...], cost }
  function bestSplit(s, ranks, nLines, pt, widths) {
    var cand = [];
    for (var i = 1; i < s.length; i++) if (ranks[i] != null) cand.push(i);
    var best = null, W = function (k) { return widths[Math.min(k, widths.length - 1)]; };
    var walk = function (k, from, cuts, cost, worst) {
      if (best && cost >= best.cost) return;
      if (k === nLines - 1) {
        var last = piece(s, from, s.length);
        if (!last || !fits(last, pt, W(k))) return;
        var lines = [], prev = 0, all = cuts.concat([s.length]);
        for (var j = 0; j < all.length; j++) { lines.push(piece(s, prev, all[j])); prev = all[j]; }
        var ws = lines.map(widthEm), mx = Math.max.apply(null, ws), mn = Math.min.apply(null, ws);
        var total = cost + orphan(last) + BAL_W * (mx - mn) / (mx + mn || 1);
        if (!best || total < best.cost) best = { cuts: cuts.slice(), lines: lines, cost: total, worst: worst };
        return;
      }
      for (var c = 0; c < cand.length; c++) {
        var at = cand[c];
        if (at <= from) continue;
        var line = piece(s, from, at);
        if (!line) continue;
        if (!fits(line, pt, W(k))) break;                                // ここから先はもっと長くなる
        walk(k + 1, at, cuts.concat([at]), cost + pen(ranks[at]) + orphan(line), Math.max(worst, ranks[at]));
      }
    };
    walk(0, 0, [], 0, 0);
    return best;
  }

  // 会社名の行の分け方と字の大きさを決める。
  //   o.width    … 2行以上のときの1行の幅（pt。枠の内側）
  //   o.width1   … 1行で入れるときの幅（pt。無ければ width）
  //   o.pt       … いつもの大きさ。o.minPt … 2行で小さくしてよい下限。o.hardMinPt … 3行以上での下限（無ければ minPt）
  //   o.maxLines … 何行まで（既定 2）。o.step … 小さくするときの刻み（既定：30pt以上は2pt・16pt以上は1pt・それより小さいと0.5pt）
  //   o.shrinkFirst … 行を増やす前に、1行のまま小さくしてよいか（既定 false：いつもの大きさのまま2行にする）
  //   o.linePenalty … 3行目からの1行ごとの減点（既定 3。大きくすると、行を増やすより字を小さくする）
  //   o.accept(lines, pt) … ほかの決まり（下の図形まで入るかなど）。false を返した分け方・大きさは使わない
  //   o.latinJoin … 英字のあとに漢字・カタカナが続くところ（「DX推進」）を折りにくくする（カテゴリー。会社名では使わない）
  // → { lines: [...], pt, natural（語の途中で折っていない）}
  function layout(raw, o) {
    o = o || {};
    var s = normalize(raw), base = o.pt || 44, minPt = o.minPt || base, hard = o.hardMinPt || minPt;
    var w = o.width, w1 = o.width1 || w, maxLines = o.maxLines || 2;
    var step = o.step || (base >= 30 ? 2 : base >= 16 ? 1 : 0.5), lineCost = o.linePenalty == null ? 3 : o.linePenalty;
    if (!s) return { lines: [''], pt: base, natural: true };
    if (fits(s, base, w1) && (!o.accept || o.accept([s], base))) return { lines: [s], pt: base, natural: true };
    var ranks = cutRanks(s, o), best = null;
    var consider = function (lines, pt, cost, worst) {
      if (o.accept && !o.accept(lines, pt)) return;
      if (!best || cost < best.cost - 1e-9) best = { lines: lines, pt: pt, cost: cost, worst: worst };
    };
    var floor = Math.min(minPt, hard);
    for (var pt = base; pt >= floor - 1e-9; pt -= step) {
      var sizeCost = SIZE_W * (1 - pt / base);
      if (best && sizeCost - (maxLines - 1) >= best.cost) break;         // これより小さくしても良くならない
      if (o.shrinkFirst && fits(s, pt, w1)) consider([s], pt, sizeCost, 0);
      for (var n = 2; n <= maxLines; n++) {
        if (pt < (n === 2 ? minPt : hard) - 1e-9) continue;
        var sp = bestSplit(s, ranks, n, pt, [w]);
        if (sp) consider(sp.lines, pt, sp.cost + sizeCost + (n - 2) * lineCost, sp.worst);
      }
    }
    if (best) return { lines: best.lines, pt: best.pt, natural: best.worst <= 3 };
    // どの大きさでも入らない（とても長い）：いちばん小さい大きさで、行の幅いっぱいに詰めて折る（英単語の途中も）。
    // fallback … 決まりの中で入らなかったしるし（呼ぶ側で、さらに小さくするなどを決める）
    return { lines: greedy(s, floor, w, maxLines), pt: floor, natural: false, fallback: true };
  }
  // 詰めて折る（入りきらないときだけ使う）。行の頭に来てはいけない字は、前の字と一緒に次の行へ
  function greedy(s, pt, width, maxLines) {
    var out = [], cur = '';
    for (var i = 0; i < s.length; i++) {
      var c = s.charAt(i);
      if (cur && out.length < maxLines - 1 && !fits(cur + c, pt, width)) {
        var j = cur.length, next = c;
        while (j > 1 && (NO_START.test(next) || NO_END.test(cur.charAt(j - 1)))) { next = cur.charAt(j - 1); j--; }
        out.push(cur.slice(0, j).replace(/[ \u3000]+$/, ''));
        cur = (cur.slice(j) + c).replace(/^[ \u3000]+/, '');
      } else cur += c;
    }
    if (cur) out.push(cur);
    return out;
  }

  // 幅に関係なく、2行に分けるときの良い順の分け方（HTMLのメンバーブックのように、あとで枠に合わせて選ぶところ用）
  //   → [[1行目, 2行目], …]（最大 max 通り）
  function twoLineOptions(raw, max) {
    var s = normalize(raw), ranks = cutRanks(s), out = [];
    for (var i = 1; i < s.length; i++) {
      if (ranks[i] == null) continue;
      var a = piece(s, 0, i), b = piece(s, i, s.length);
      if (!a || !b) continue;
      var wa = widthEm(a), wb = widthEm(b);
      out.push({ lines: [a, b], cost: pen(ranks[i]) + orphan(a) + orphan(b) + BAL_W * Math.abs(wa - wb) / (wa + wb || 1) });
    }
    out.sort(function (x, y) { return x.cost - y.cost; });
    return out.slice(0, max || 3).map(function (x) { return x.lines; });
  }

  // 1行の見積もりの幅（pt。余裕こみ）。1行に縮めるところ（2人のページなど）で、PowerPoint が折り返さない大きさを出すのに使う
  function lineWidthPt(line, pt) { return widthEm(line) * pt * MARGIN; }

  return { normalize: normalize, widthEm: widthEm, lineWidthPt: lineWidthPt, fits: fits, cutRanks: cutRanks, layout: layout,
           twoLineOptions: twoLineOptions, MARGIN: MARGIN };
}
var CB_ = companyBreakLib_();
