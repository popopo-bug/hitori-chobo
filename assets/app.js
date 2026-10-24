/*! ひとり帳簿 v1.0.0 (web) */
(function(g){g.HC=g.HC||{};g.HC.edition="web";g.HC.version="1.0.0";})(window);
/* ---- core.js ---- */
/* ひとり帳簿 — core utilities */
(function (g) {
  'use strict';
  const HC = (g.HC = g.HC || {});
  HC.edition = HC.edition || 'web'; // build で 'web' | 'pro' | 'trial' に置換

  const FW = /[０-９．，－ー―‐＋]/g;
  const FW_MAP = { '．': '.', '，': ',', '－': '-', 'ー': '-', '―': '-', '‐': '-', '＋': '+' };

  /** 全角数字・カンマ・円記号まじりの文字列を数値に。空や不正は 0。（ ）や△は負数 */
  function num(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    if (v == null) return 0;
    let s = String(v).replace(FW, (c) => FW_MAP[c] || String.fromCharCode(c.charCodeAt(0) - 0xfee0));
    s = s.replace(/[,\s¥￥円]/g, '');
    let neg = false;
    if (/^\(.*\)$/.test(s) || /^（.*）$/.test(s)) {
      neg = true;
      s = s.slice(1, -1);
    }
    if (/^[△▲]/.test(s)) {
      neg = true;
      s = s.slice(1);
    }
    if (s === '' || s === '-' || s === '.') return 0;
    const n = Number(s);
    if (!isFinite(n)) return 0;
    return neg ? -n : n;
  }

  /** 3桁区切り（負数対応） */
  function yen(n) {
    n = Math.round(num(n));
    const s = Math.abs(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (n < 0 ? '-' : '') + s;
  }

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function uid() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  function pad(n, w) {
    return String(n).padStart(w || 2, '0');
  }

  /** ローカル日付 'YYYY-MM-DD' */
  function ymd(d) {
    d = d || new Date();
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  }

  function parseYmd(s) {
    const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(s || '').trim());
    if (!m) return null;
    const d = new Date(+m[1], +m[2] - 1, +m[3]);
    if (isNaN(d) || d.getMonth() !== +m[2] - 1) return null;
    return d;
  }

  /** 'YYYY-MM-DD' → '2026年10月23日' */
  function dateJa(s) {
    const d = parseYmd(s);
    if (!d) return s ? String(s) : '';
    return d.getFullYear() + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日';
  }

  /** 'YYYY-MM-DD' → '10/23' */
  function md(s) {
    const d = parseYmd(s);
    return d ? d.getMonth() + 1 + '/' + d.getDate() : '';
  }

  function clone(o) {
    return o == null ? o : JSON.parse(JSON.stringify(o));
  }

  function debounce(fn, ms) {
    let t;
    return function () {
      const args = arguments, self = this;
      clearTimeout(t);
      t = setTimeout(() => fn.apply(self, args), ms);
    };
  }

  /** 照合用に正規化（全角英数→半角、半角カナ→全角、小文字、空白除去） */
  function norm(s) {
    return String(s == null ? '' : s)
      .normalize('NFKC')
      .toLowerCase()
      .replace(/[\s　・･\-－ー_.,、。()（）［］\[\]「」『』"'*＊/\\]/g, '');
  }

  /** ファイル名に使えない文字を除去 */
  function safeFileName(s) {
    return String(s || '')
      .replace(/[\\/:*?"<>|\r\n\t]/g, '')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 80);
  }

  function download(filename, content, mime) {
    const blob = content instanceof Blob ? content : new Blob([content], { type: mime || 'application/octet-stream' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 1000);
  }

  /** CSV（Excel向けに UTF-8 BOM 付き） */
  function toCsv(rows) {
    const body = rows
      .map((r) =>
        r
          .map((c) => {
            const s = c == null ? '' : String(c);
            return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
          })
          .join(',')
      )
      .join('\r\n');
    return '﻿' + body + '\r\n';
  }

  function toast(msg, kind) {
    if (typeof document === 'undefined') return;
    let box = document.querySelector('.hc-toasts');
    if (!box) {
      box = document.createElement('div');
      box.className = 'hc-toasts';
      box.setAttribute('role', 'status');
      box.setAttribute('aria-live', 'polite');
      document.body.appendChild(box);
    }
    const el = document.createElement('div');
    el.className = 'hc-toast' + (kind ? ' is-' + kind : '');
    el.textContent = msg;
    box.appendChild(el);
    setTimeout(() => el.classList.add('is-out'), 2600);
    setTimeout(() => el.remove(), 3100);
  }

  /** モーダル。buttons = [[ラベル, 値], ...]（最後が主ボタン）。閉じたときの値を返す */
  function dialog(title, body, buttons) {
    return new Promise((resolve) => {
      const dl = document.createElement('dialog');
      dl.className = 'hc-dialog';
      dl.innerHTML =
        '<form method="dialog"><div class="dlg-head">' + esc(title) + '</div><div class="dlg-body">' + body + '</div><div class="dlg-foot">' +
        buttons.map((b, i) => '<button class="btn ' + (i === buttons.length - 1 ? 'btn-primary' : 'btn-ghost') + '" value="' + esc(b[1]) + '">' + esc(b[0]) + '</button>').join('') +
        '</div></form>';
      document.body.appendChild(dl);
      // ボタンを押したときに確定する（close イベントが遅れる環境でも動くように）。Esc で閉じたときは close で受ける
      let done = false;
      const finish = (v) => {
        if (done) return;
        done = true;
        if (dl.open) dl.close(v);
        dl.remove();
        resolve(v);
      };
      dl.querySelector('form').addEventListener('submit', (ev) => {
        ev.preventDefault();
        finish(ev.submitter ? ev.submitter.value : '');
      });
      dl.addEventListener('close', () => finish(dl.returnValue || ''));
      dl.showModal();
    });
  }

  // ---- localStorage（キーの頭に hc1:） ----
  const PREFIX = 'hc1:';
  const store = {
    get(key, fallback) {
      try {
        const v = g.localStorage.getItem(PREFIX + key);
        return v == null ? fallback : JSON.parse(v);
      } catch (e) {
        return fallback;
      }
    },
    set(key, val) {
      try {
        g.localStorage.setItem(PREFIX + key, JSON.stringify(val));
        return true;
      } catch (e) {
        toast('ブラウザの保存容量がいっぱいです。古い年の取引をCSVに書き出して削除するなど、容量を空けてください。', 'error');
        return false;
      }
    },
    remove(key) {
      try {
        g.localStorage.removeItem(PREFIX + key);
      } catch (e) { /* noop */ }
    },
    keys() {
      const out = [];
      try {
        for (let i = 0; i < g.localStorage.length; i++) {
          const k = g.localStorage.key(i);
          if (k && k.indexOf(PREFIX) === 0) out.push(k.slice(PREFIX.length));
        }
      } catch (e) { /* noop */ }
      return out;
    },
    usage() {
      let bytes = 0;
      try {
        store.keys().forEach((k) => (bytes += (k.length + (g.localStorage.getItem(PREFIX + k) || '').length) * 2));
      } catch (e) { /* noop */ }
      return bytes;
    }
  };

  HC.util = { num, yen, esc, uid, pad, ymd, parseYmd, dateJa, md, clone, debounce, norm, safeFileName, download, toCsv, toast, dialog };
  HC.store = store;
})(typeof window !== 'undefined' ? window : globalThis);

/* ---- kamoku.js ---- */
/* ひとり帳簿 — 勘定科目と「何費？」辞書（Web版の科目さがし・Pro版の自動分類で共通） */
(function (g) {
  'use strict';
  const HC = (g.HC = g.HC || {});

  // kind: income 収入 / cogs 売上原価 / expense 経費
  // std: 決算書・収支内訳書に印刷済みの科目（false は空欄に書く科目）
  // rate: 家事按分（事業で使う割合 %）の初期値 / hidden: 最初は入力の選択肢に出さない
  const ACCOUNTS = [
    { id: 'sales', name: '売上', kind: 'income', std: true, desc: '商品・サービスの売上、報酬、原稿料、広告収入など（手数料を引く前の金額）' },
    { id: 'other_income', name: '雑収入', kind: 'income', std: true, desc: '事業に付随する、売上以外の収入（事業に関する補助金など）' },
    { id: 'kaji', name: '家事消費', kind: 'income', std: true, hidden: true, desc: '売り物を自分や家族のために使った・食べた分（お店向け）' },
    { id: 'purchase', name: '仕入', kind: 'cogs', std: true, desc: '売るための商品・材料、同人誌やグッズの製作費（年末に売れ残った分は在庫になる）' },
    { id: 'tax', name: '租税公課', kind: 'expense', std: true, desc: '個人事業税、事業分の固定資産税・自動車税、収入印紙など（所得税・住民税は×）' },
    { id: 'freight', name: '荷造運賃', kind: 'expense', std: true, desc: '商品の発送料、梱包材' },
    { id: 'utility', name: '水道光熱費', kind: 'expense', std: true, desc: '電気・ガス・水道（自宅兼事務所なら家事按分）' },
    { id: 'travel', name: '旅費交通費', kind: 'expense', std: true, desc: '電車・バス・タクシー・飛行機、出張の宿泊費、高速道路代' },
    { id: 'comm', name: '通信費', kind: 'expense', std: true, desc: 'スマホ・ネット回線、切手、サーバー・ドメイン、仕事用のサブスク' },
    { id: 'ad', name: '広告宣伝費', kind: 'expense', std: true, desc: '広告、チラシ、名刺、ノベルティ、イベント出展料' },
    { id: 'entertain', name: '接待交際費', kind: 'expense', std: true, desc: '取引先との会食、手土産、お祝い・お悔やみ' },
    { id: 'insurance', name: '損害保険料', kind: 'expense', std: true, desc: '事業用の火災保険・自動車保険・賠償責任保険（生命保険は×）' },
    { id: 'repair', name: '修繕費', kind: 'expense', std: true, desc: 'パソコン・機材・事務所の修理' },
    { id: 'supplies', name: '消耗品費', kind: 'expense', std: true, desc: '文房具、画材、10万円未満のパソコン・周辺機器・家具・ソフト' },
    { id: 'depreciation', name: '減価償却費', kind: 'expense', std: true, auto: true, desc: '10万円以上のパソコン・車など（「固定資産」に登録すると自動で計算）' },
    { id: 'welfare', name: '福利厚生費', kind: 'expense', std: true, hidden: true, desc: '従業員の健康診断・慶弔など（自分だけのための支出は×）' },
    { id: 'wages', name: '給料賃金', kind: 'expense', std: true, hidden: true, desc: '従業員・アルバイトの給料' },
    { id: 'outsourcing', name: '外注工賃', kind: 'expense', std: true, desc: '仕事の一部を外部に頼んだ費用（デザイン・イラスト・動画編集・アシスタントなど）' },
    { id: 'interest', name: '利子割引料', kind: 'expense', std: true, hidden: true, desc: '事業用の借入金の利息' },
    { id: 'rent', name: '地代家賃', kind: 'expense', std: true, desc: '事務所・自宅の家賃（事業分）、月極駐車場、コワーキング、トランクルーム' },
    { id: 'baddebt', name: '貸倒金', kind: 'expense', std: true, hidden: true, desc: '回収できなくなった売掛金' },
    { id: 'fees', name: '支払手数料', kind: 'expense', std: false, desc: '振込手数料、BOOTH・FANBOX・Skebなどの販売手数料、決済手数料、税理士報酬' },
    { id: 'books', name: '新聞図書費', kind: 'expense', std: false, desc: '仕事の本・雑誌・新聞・電子書籍・資料' },
    { id: 'training', name: '研修費', kind: 'expense', std: false, desc: 'セミナー・講座・勉強会の参加費' },
    { id: 'meeting', name: '会議費', kind: 'expense', std: false, desc: '打ち合わせの飲食代、会議室の利用料' },
    { id: 'car', name: '車両費', kind: 'expense', std: false, desc: 'ガソリン代、車検、点検・整備（事業分）' },
    { id: 'dues', name: '諸会費', kind: 'expense', std: false, hidden: true, desc: '商工会・業界団体・協会の会費' },
    { id: 'misc', name: '雑費', kind: 'expense', std: true, desc: 'どの科目にも当てはまらない、少額で頻度の低いもの' }
  ];

  // 決算書・収支内訳書の経費の並び（'*' は自分で科目名を書く空欄）
  const FORM_ROWS = {
    // 収支内訳書（一般用）：給料賃金〜利子割引料のあと「その他の経費」に租税公課〜福利厚生費、空欄5行、雑費
    white: ['wages', 'outsourcing', 'depreciation', 'baddebt', 'rent', 'interest', 'tax', 'freight', 'utility', 'travel', 'comm', 'ad', 'entertain', 'insurance', 'repair', 'supplies', 'welfare', '*', 'misc'],
    // 青色申告決算書（一般用）：租税公課〜貸倒金、空欄6行、雑費
    blue: ['tax', 'freight', 'utility', 'travel', 'comm', 'ad', 'entertain', 'insurance', 'repair', 'supplies', 'depreciation', 'welfare', 'wages', 'outsourcing', 'interest', 'rent', 'baddebt', '*', 'misc']
  };
  const BLANK_ROWS = { white: 5, blue: 6 };
  // 収支内訳書で「その他の経費」の小計に入る科目（給料賃金〜利子割引料以外）
  const WHITE_MAIN = ['wages', 'outsourcing', 'depreciation', 'baddebt', 'rent', 'interest'];

  // ---- 「何費？」辞書 ----
  // [見出し語|別名…, 科目id, ひとこと, 印]  印：h=家事按分することが多い / x=経費にならない / f=10万円以上なら固定資産
  // 科目id が null のものは「経費にならない（控除など別の扱い）」
  const DICT = [
    // 売上・収入
    ['売上|売り上げ|代金の入金|報酬|ギャラ', 'sales', '手数料や源泉徴収を引く前の金額で記帳。引かれた手数料は支払手数料、源泉徴収税額は確定申告で精算'],
    ['原稿料|執筆料|ライティング報酬|記事の報酬', 'sales', '源泉徴収されていれば、その金額も記録しておく（確定申告で差し引ける）'],
    ['デザイン料|イラスト料|作画料|制作料|撮影料|講演料', 'sales', '源泉徴収の対象になることが多い報酬。差し引かれた額を記録'],
    ['BOOTHの売上|ブースの売上|BOOTH売上', 'sales', '売れた金額（手数料を引く前）で記帳。BOOTHの手数料は支払手数料'],
    ['FANBOXの支援金|pixivFANBOX|ファンボックス|Fantia|ファンティア|メンバーシップ|ファンクラブ', 'sales', '支援額（手数料を引く前）を売上に。手数料は支払手数料'],
    ['Skebの報酬|スケブ|コミッションの報酬|有償依頼の報酬', 'sales', '依頼の報酬は売上。差し引かれた手数料は支払手数料'],
    ['同人誌の売上|即売会の売上|イベントの売上|コミケの売上|頒布', 'sales', 'イベントごとに「何を何部、いくらで」をメモしておくと安心'],
    ['ハンドメイドの売上|minne|ミンネ|Creema|クリーマ', 'sales', '売れた金額（手数料・送料を分けて）を記帳'],
    ['YouTubeの広告収入|アドセンス|AdSense|広告収入|再生数の収益', 'sales', '振り込まれた月ではなく、確定した収益を記帳するのが基本'],
    ['投げ銭|スーパーチャット|スパチャ|ギフト|チップ', 'sales', '配信サービスの手数料を引く前の金額がわかれば、その額で'],
    ['アフィリエイト収入|アフィリエイト報酬|紹介報酬', 'sales', '事業として行っていれば売上'],
    ['note の売上|有料記事の売上|有料noteの売上', 'sales', '売上は手数料を引く前の金額、noteの手数料は支払手数料'],
    ['補助金|助成金|給付金', 'other_income', '事業に関するものは雑収入（課税されないものもあるので、交付元の案内を確認）'],
    ['売り物を自分で使った|自家消費|家事消費', 'kaji', '売り物を自分や家族のために使った分（お店向け）'],
    ['預金の利息|利息の受取|リソク', null, '預金の利息は源泉分離課税で終わっているので、帳簿の売上には入れない'],

    // 仕入
    ['仕入|商品の仕入れ|仕入れ', 'purchase', '売るための商品。年末に売れ残った分は在庫として差し引く'],
    ['同人誌の印刷代|同人誌の印刷費|本の印刷代|製本代|入稿', 'purchase', '売る本の製作費は仕入（年末の在庫は売上原価から引く）'],
    ['グッズの製作費|アクリルキーホルダー|アクスタ|缶バッジ|ステッカーの製作|グッズ制作', 'purchase', '売るグッズの製作費は仕入。売れ残りは在庫に'],
    ['ハンドメイドの材料|材料費|ビーズ|レジン|布地|生地|パーツ', 'purchase', '作品の材料は仕入（自分の練習用は消耗品費でも）'],
    ['委託販売の商品|卸の仕入', 'purchase', '仕入。委託先の手数料は支払手数料'],

    // 租税公課
    ['個人事業税|事業税', 'tax', '払った年の経費'],
    ['固定資産税', 'tax', '自宅兼事務所なら事業に使う割合だけ', 'h'],
    ['自動車税|軽自動車税|自動車重量税', 'tax', '車を仕事にも使うなら事業分だけ', 'h'],
    ['収入印紙|印紙代|印紙税', 'tax', '契約書・領収書に貼った印紙'],
    ['消費税の納付|消費税を納めた', 'tax', '税込経理なら、納めた年の租税公課（免税事業者は関係なし）'],
    ['登録免許税|印鑑証明書|住民票の写し|証明書の発行手数料', 'tax', '仕事の手続きのためなら租税公課'],
    ['所得税|住民税|復興特別所得税|予定納税', null, '経費にならない（自分にかかる税金）', 'x'],
    ['延滞税|加算税|延滞金|罰金|交通違反の反則金|駐車違反', null, '経費にならない', 'x'],

    // 荷造運賃
    ['送料|発送料|配送料|宅配便|ゆうパック|ゆうパケット|クリックポスト|ネコポス|宅急便|佐川急便|ヤマト運輸', 'freight', '商品の発送は荷造運賃（書類の郵送は通信費）'],
    ['梱包材|ダンボール|段ボール|緩衝材|プチプチ|OPP袋|封筒（発送用）', 'freight', '発送用の資材は荷造運賃（消耗品費にしている人も。毎年同じ科目に）'],
    ['イベント会場への搬入|搬入|直接搬入|宅配搬入', 'freight', '会場に本やグッズを送った送料'],

    // 水道光熱費
    ['電気代|電気料金|電力|でんき', 'utility', '自宅兼事務所なら、使う時間や部屋の広さで按分', 'h'],
    ['ガス代|ガス料金', 'utility', '仕事でほとんど使わないなら経費にしないのが無難', 'h'],
    ['水道代|水道料金', 'utility', '仕事との関係が説明しにくいので、自宅では入れない人が多い', 'h'],
    ['灯油|ストーブの燃料', 'utility', '仕事部屋の暖房分', 'h'],

    // 旅費交通費
    ['電車代|電車賃|運賃|切符|バス代|地下鉄', 'travel', '仕事の移動。日付・区間・目的をメモ'],
    ['新幹線|特急券|飛行機|航空券|フェリー', 'travel', '出張・取材・イベント参加の移動'],
    ['タクシー|タクシー代|配車アプリ', 'travel', '仕事の移動なら旅費交通費'],
    ['宿泊費|ホテル代|出張の宿泊|旅館', 'travel', '仕事の出張。観光の分は分ける'],
    ['高速道路|高速代|ETC|有料道路', 'travel', '仕事の移動分（車両費にまとめる人も）', 'h'],
    ['コインパーキング|駐車場代（一時）|時間貸し駐車場', 'travel', '仕事で出かけた先の駐車場（月極は地代家賃）'],
    ['交通系ICのチャージ|Suicaのチャージ|PASMOのチャージ|ICカードのチャージ', 'travel', 'チャージした時ではなく、仕事で使った分を経費に（履歴を印字・保存）'],
    ['通勤定期|定期券|定期代', 'travel', '事務所への通勤分。私用も多いなら按分', 'h'],
    ['即売会への交通費|イベントへの交通費|取材の交通費|ロケハン', 'travel', '仕事の移動として旅費交通費'],
    ['シェアサイクル|レンタサイクル', 'travel', '仕事の移動分'],

    // 通信費
    ['スマホ代|携帯電話|携帯代|スマートフォンの料金|通話料', 'comm', '仕事と私用で使うなら按分（使った時間などで）', 'h'],
    ['インターネット回線|光回線|プロバイダ|Wi-Fi|ポケットWi-Fi|ネット代', 'comm', '自宅兼事務所なら按分', 'h'],
    ['固定電話|電話代', 'comm', '仕事用の回線なら全額', 'h'],
    ['切手|はがき|レターパック|郵便料金|書類の郵送', 'comm', '書類の郵送は通信費（商品の発送は荷造運賃）'],
    ['レンタルサーバー|サーバー代|ドメイン|独自ドメイン|ホスティング', 'comm', '仕事のサイトなら通信費'],
    ['クラウドストレージ|Dropbox|iCloud|Google One|OneDrive', 'comm', '仕事のデータ保存に使う分', 'h'],
    ['Adobe|Creative Cloud|クリエイティブクラウド|Photoshop|Illustrator|Premiere', 'comm', '月額・年額のソフトは通信費にする人が多い（支払手数料でも。毎年同じ科目に）'],
    ['CLIP STUDIO の月額|クリスタの月額|ソフトの月額|サブスク|サブスクリプション|SaaS', 'comm', '仕事で使うソフトの利用料（買い切りのソフトは消耗品費）'],
    ['ChatGPT|Claude|生成AI|AIツール|AIサービス', 'comm', '仕事で使うAIサービスの利用料'],
    ['Zoom|ビデオ会議|Web会議|Slack|Notion|Canva', 'comm', '仕事で使うツールの利用料'],

    // 広告宣伝費
    ['広告|広告費|Web広告|SNS広告|X広告|Twitter広告|Instagram広告|Google広告', 'ad', '仕事の宣伝のための広告'],
    ['名刺|名刺の印刷', 'ad', '広告宣伝費か消耗品費（毎年同じ科目に）'],
    ['チラシ|フライヤー|ポスター|のぼり|ポップ|看板', 'ad', '宣伝用の印刷物（3年以上使う看板で高額なら固定資産）'],
    ['ノベルティ|無料配布|配布ペーパー|試供品|サンプル品', 'ad', '宣伝のために配るもの'],
    ['サークル参加費|イベント出展料|出展料|ブース代|申込費', 'ad', '即売会・展示会の参加費（支払手数料にする人も。毎年同じ科目に）'],
    ['ポートフォリオサイト|ホームページ制作|LP制作', 'ad', '宣伝用のサイト制作（高額で長く使うものは固定資産になることも）'],
    ['求人広告|アルバイト募集', 'ad', '人を募集する広告'],

    // 接待交際費
    ['接待|会食|取引先との飲食|取引先との食事', 'entertain', '相手・人数・目的をメモ。1人や友人との食事は×'],
    ['手土産|お土産（取引先）|差し入れ', 'entertain', '取引先・仕事仲間へのもの'],
    ['お中元|お歳暮|お祝い|御祝儀|香典|お花代|慶弔', 'entertain', '取引先とのお付き合い'],
    ['取引先とのゴルフ|ゴルフ代', 'entertain', '取引先との付き合いの分だけ'],

    // 損害保険料
    ['火災保険|地震保険', 'insurance', '自宅兼事務所なら事業分だけ', 'h'],
    ['自動車保険|任意保険|自賠責保険', 'insurance', '車を仕事にも使うなら事業分だけ', 'h'],
    ['賠償責任保険|フリーランス保険|業務の保険', 'insurance', '仕事のための保険'],
    ['生命保険|医療保険|がん保険|学資保険|個人年金保険', null, '経費にならない（生命保険料控除で申告）', 'x'],

    // 修繕費
    ['修理|修理代|パソコンの修理|機材の修理|カメラの修理|画面割れの修理', 'repair', '元の状態に戻す修理は修繕費（性能が上がる改造は資産になることも）'],
    ['事務所の修繕|内装の修理|原状回復', 'repair', '事務所部分の修理'],

    // 消耗品費
    ['文房具|事務用品|~ペン|ノート|コピー用紙|プリンターのインク|トナー|印刷用紙', 'supplies', '少額の事務用品'],
    ['画材|インク|スクリーントーン|原稿用紙|コピック|絵の具|筆', 'supplies', '作品づくりの道具（売る作品の材料は仕入）'],
    ['パソコン|ノートパソコン|PC|~Mac|MacBook|iPad|タブレット', 'supplies', '10万円未満なら消耗品費。10万円以上は固定資産（減価償却）', 'f'],
    ['液タブ|ペンタブ|板タブ|液晶タブレット', 'supplies', '10万円未満なら消耗品費。10万円以上は固定資産', 'f'],
    ['モニター|ディスプレイ|キーボード|マウス|ヘッドホン|イヤホン|Webカメラ|マイク|オーディオインターフェース', 'supplies', '10万円未満なら消耗品費。私用もあるなら按分', 'f'],
    ['外付けHDD|SSD|USBメモリ|SDカード|ケーブル|充電器|モバイルバッテリー', 'supplies', '仕事用の周辺機器'],
    ['プリンター|スキャナー|複合機', 'supplies', '10万円未満なら消耗品費', 'f'],
    ['カメラ|レンズ|三脚|照明|撮影機材', 'supplies', '10万円未満なら消耗品費。10万円以上は固定資産', 'f'],
    ['スマホ本体|スマートフォン本体|携帯電話の本体|機種変更', 'supplies', '本体代は10万円未満なら消耗品費（分割払いでも本体の価格で判断）。私用もあるなら按分', 'hf'],
    ['机|デスク|椅子|チェア|本棚|棚|収納', 'supplies', '10万円未満なら消耗品費', 'f'],
    ['ソフトウェア（買い切り）|買い切りのソフト|CLIP STUDIO PAINT|フォント|素材集', 'supplies', '10万円未満の買い切りソフトは消耗品費', 'f'],
    ['イラスト素材|写真素材|3Dモデル素材|音源素材|BGM素材|効果音', 'supplies', '作品に使う素材（売る作品に組み込むなら仕入でも）'],
    ['電池|電球|LED|延長コード|掃除用品', 'supplies', '仕事場で使う分', 'h'],
    ['ゲーム機|ゲームソフト', 'supplies', '実況・レビューなど仕事で使うものだけ。私用は×', 'h'],
    ['衣装|撮影用の衣装|ウィッグ|小道具', 'supplies', '仕事専用のものだけ（普段も着る服は×）'],
    ['作業着|制服|ユニフォーム', 'supplies', '仕事でしか使わないもの'],

    // 減価償却
    ['10万円以上のパソコン|高額な機材|10万円以上の備品|車の購入|自動車の購入|バイクの購入', 'depreciation', '「固定資産」に登録すると、耐用年数に合わせて毎年の減価償却費を計算', 'f'],

    // 福利厚生費・給料
    ['従業員の健康診断|社員旅行|従業員の慶弔', 'welfare', '従業員のための支出（自分だけのためのものは×）'],
    ['アルバイト代|パート代|従業員の給料|給与', 'wages', '源泉徴収・年末調整が必要なことも'],
    ['自分への給料|事業主の給料|生活費', null, '経費にならない（個人事業主は自分に給料を払えない）', 'x'],

    // 外注工賃
    ['外注|外注費|業務委託|下請け', 'outsourcing', '仕事の一部を外部の人・会社に頼んだ費用'],
    ['デザインの外注|イラストの依頼|動画編集の外注|翻訳の依頼|ライターへの依頼|コーディングの外注', 'outsourcing', '頼んだ相手が個人なら、源泉徴収が必要なことも'],
    ['アシスタント代|作画アシスタント|背景の外注', 'outsourcing', '業務委託なら外注工賃（雇っているなら給料賃金）'],
    ['Skebで依頼|有償依頼|コミッション依頼|ココナラで依頼|クラウドワークスで依頼|ランサーズで依頼', 'outsourcing', '仕事に使う作品・作業を依頼した費用'],

    // 利子割引料
    ['借入金の利息|ローンの利息|融資の利息', 'interest', '事業用の借入の利息だけ（元本の返済は経費にならない）', 'h'],
    ['借入金の返済|ローンの返済（元本）|元本', null, '元本の返済は経費にならない', 'x'],

    // 地代家賃
    ['家賃|事務所の家賃|自宅の家賃|賃料', 'rent', '自宅兼事務所なら床面積の割合などで按分', 'h'],
    ['月極駐車場|駐車場代（月極）', 'rent', '仕事にも使う車なら按分', 'h'],
    ['コワーキング|レンタルオフィス|シェアオフィス|バーチャルオフィス', 'rent', '月額で借りる仕事場'],
    ['トランクルーム|倉庫|在庫の保管', 'rent', '在庫や機材の保管場所'],
    ['レンタルスペース|スタジオ代|撮影スタジオ|貸し会議室', 'rent', '仕事で借りた場所（打ち合わせだけなら会議費でも）'],
    ['住宅ローン', null, 'ローンの返済は経費にならない（持ち家は減価償却・固定資産税・利息の事業分で考える）', 'x'],

    // 支払手数料
    ['振込手数料|送金手数料|ATM手数料', 'fees', '仕事のお金の振込・引き出しにかかったもの'],
    ['BOOTHの手数料|BOOTHのサービス利用料|ブースの手数料|BOOTH手数料', 'fees', '売上は手数料を引く前の金額で、手数料はここ'],
    ['FANBOXの手数料|Fantiaの手数料|Skebの手数料|noteの手数料|ココナラの手数料|minneの手数料|Creemaの手数料|販売手数料|システム利用料', 'fees', 'プラットフォームに引かれた手数料'],
    ['決済手数料|カード決済の手数料|Square|Stripe|PayPal|AirPAY', 'fees', 'キャッシュレス決済の手数料'],
    ['税理士報酬|税理士への報酬|記帳代行|弁護士費用|司法書士', 'fees', '支払った相手が個人なら、源泉徴収が必要なことも'],
    ['クレジットカードの年会費|カード年会費', 'fees', '仕事専用のカードなら経費（私用もあるなら按分）', 'h'],
    ['代引き手数料|代金引換の手数料', 'fees', '支払手数料'],
    ['不動産の仲介手数料|事務所の仲介手数料', 'fees', '事務所を借りたときの仲介手数料（20万円以上の礼金などは繰延資産になることも）'],

    // 新聞図書費
    ['本|書籍|技術書|ビジネス書|専門書|電子書籍|Kindle', 'books', '仕事の本・資料'],
    ['雑誌|新聞|業界紙|定期購読', 'books', '仕事の情報収集のためのもの'],
    ['資料|資料集|写真集|画集|図鑑|資料用の本', 'books', '作品づくりの資料（仕事で使う理由を説明できるもの）'],
    ['有料記事|有料note|オンラインサロン（学習）|ニュースサイトの購読', 'books', '仕事の情報収集（研修費にする人も）'],

    // 研修費
    ['セミナー|講座|オンライン講座|Udemy|勉強会|ワークショップ', 'training', '今の仕事に必要な知識・技術のためのもの'],
    ['資格の受験料|検定料|試験料', 'training', '今の仕事に直接必要な資格なら'],

    // 会議費
    ['打ち合わせ|打合せ|カフェで打ち合わせ|ミーティング', 'meeting', '相手・目的をメモ（ひとりの飲食は×）'],
    ['会議室|会議室のレンタル', 'meeting', '打ち合わせのための場所代'],

    // 車両費
    ['ガソリン|ガソリン代|燃料代|給油', 'car', '車を私用にも使うなら走行距離などで按分', 'h'],
    ['車検|点検|整備|オイル交換|タイヤ交換|洗車', 'car', '事業分だけ（修繕費にする人も）', 'h'],

    // 諸会費
    ['会費|年会費|協会の会費|商工会|組合費|業界団体', 'dues', '仕事に関係する団体の会費'],

    // 雑費
    ['ごみ処理|事業ごみ|粗大ごみ（仕事用）', 'misc', '仕事で出たごみの処理費用'],
    ['クリーニング（衣装）|仕事着のクリーニング', 'misc', '仕事専用の衣装・作業着のもの'],
    ['引っ越し代（事務所）|事務所の移転', 'misc', '事務所の移転費用'],

    // 経費にならないもの（控除で申告）
    ['国民健康保険|国保|国民年金|厚生年金|介護保険料', null, '経費にならない（確定申告の「社会保険料控除」で全額差し引ける）', 'x'],
    ['小規模企業共済|iDeCo|イデコ|個人型確定拠出年金', null, '経費にならない（「小規模企業共済等掛金控除」で全額差し引ける）', 'x'],
    ['ふるさと納税|寄付|寄附', null, '経費にならない（「寄附金控除」で申告）', 'x'],
    ['病院代|医療費|薬代|健康診断（自分）|人間ドック', null, '経費にならない（条件を満たせば「医療費控除」）', 'x'],
    ['昼食|~ランチ|自分の食事|飲み会（友人）|おやつ', null, '自分の食事は経費にならない', 'x'],
    ['私服|普段着|美容院|化粧品|~ジム', null, '私生活にも使うものは経費にしにくい（仕事専用と説明できる場合を除く）', 'x'],
    ['家族旅行|帰省|旅行（私用）', null, '私用の旅行は経費にならない', 'x']
  ];

  // CSV取り込み用：明細に出てくる店名・会社名（半角カナ・英字も NFKC で正規化して照合）
  const VENDORS = [
    ['AMAZON|アマゾン|ヨドバシ|ビックカメラ|ヤマダ電機|ヤマダデンキ|ケーズデンキ|コジマ|エディオン|ジョーシン|ドスパラ|ツクモ|ソフマップ|ロフト|東急ハンズ|ハンズ|無印良品|ダイソー|セリア|キャンドゥ|アスクル|ASKUL|モノタロウ|世界堂|画材', 'supplies'],
    ['JR|ジェイアール|SUICA|スイカ|PASMO|パスモ|ICOCA|メトロ|東急|小田急|京王|西武|東武|京急|京成|阪急|阪神|近鉄|南海|名鉄|西鉄|えきねっと|スマートEX|エクスプレス予約|全日空|ANA SKY|JAL|PEACH AVIATION|ジェットスター|スカイマーク|タクシー|TAXI|GO タクシー|DIDI|S.RIDE|東横イン|アパホテル|ルートイン|楽天トラベル|じゃらん|BOOKING.COM|NEXCO|ETC|タイムズ|TIMES|三井のリパーク|リパーク', 'travel'],
    ['NTT|ドコモ|DOCOMO|AHAMO|アハモ|AU PAY|AUひかり|KDDI|UQ|ソフトバンク|SOFTBANK|ワイモバイル|Y!MOBILE|YMOBILE|LINEMO|楽天モバイル|POVO|IIJ|MINEO|マイネオ|BIGLOBE|ビッグローブ|ソネット|SO-NET|OCN|NURO|ニューロ|フレッツ|ドコモ光|ソフトバンク光|AU光|XSERVER|エックスサーバー|さくらインターネット|SAKURA|ONAMAE|お名前|CONOHA|ロリポップ|AWS|AMAZON WEB SERVICES|GITHUB|VERCEL|NETLIFY|CLOUDFLARE|GOOGLE ONE|GOOGLE WORKSPACE|GSUITE|DROPBOX|ICLOUD|APPLE.COM|MICROSOFT|OFFICE 365|ADOBE|アドビ|ZOOM|SLACK|NOTION|CANVA|FIGMA|OPENAI|CHATGPT|ANTHROPIC|CLAUDE|CELSYS|セルシス|日本郵便|郵便局|切手', 'comm'],
    ['東京電力|関西電力|中部電力|東北電力|九州電力|北海道電力|北陸電力|中国電力|四国電力|沖縄電力|電力|でんき|デンキ|東京ガス|大阪ガス|東邦ガス|西部ガス|ガス料金|ガス代|ガスリヨウキン|水道|スイドウ', 'utility'],
    ['ヤマト|クロネコ|佐川|SAGAWA|西濃|ゆうパック|クリックポスト|ネコポス|宅急便', 'freight'],
    ['ENEOS|エネオス|出光|IDEMITSU|コスモ石油|COSMO|SHELL|シェル|キグナス|ガソリン|オートバックス|イエローハット', 'car'],
    ['スターバックス|STARBUCKS|ドトール|DOUTOR|タリーズ|TULLY|コメダ|エクセルシオール|ルノアール|サンマルク|ベローチェ', 'meeting'],
    ['紀伊國屋|紀伊国屋|丸善|ジュンク堂|有隣堂|三省堂|HONTO|AMAZON KINDLE|KINDLE|キンドル|BOOK☆WALKER|BOOKWALKER|コミックシーモア|メロンブックス|とらのあな|書店|ブックス', 'books'],
    ['UDEMY|ユーデミー|SCHOO|ストアカ|COURSERA|ドットインストール', 'training'],
    ['振込手数料|フリコミテスウリヨウ|テスウリヨウ|手数料|PAYPAL|ペイパル|STRIPE|SQUARE', 'fees'],
    ['収入印紙|印紙代|事業税', 'tax'],
    ['家賃|ヤチン|賃料|管理費|駐車場|月極', 'rent'],
    ['損保|損害保険|東京海上|三井住友海上|損保ジャパン|あいおいニッセイ|ソニー損保|イーデザイン|SBI損保|アクサ損保', 'insurance']
  ];
  // 取り込まない方がよい行（カードの引き落とし・口座間の移動・現金の出し入れ・税金や社会保険など私的な支払い）
  const SKIP_BANK = 'カード|クレジット|VISA|JCB|AMEX|マスターカード|MASTERCARD|セゾン|SAISON|ラクテンカード|楽天カード|三井住友カード|ミツイスミトモカ|イオンカード|エポス|EPOS|オリコ|ORICO|ビューカード|ライフカード|ジャックス|JACCS|ATM|カード引出|引出|ヒキダシ|振替|フリカエ|定期預金|積立|ツミタテ|利息|リソク';
  const SKIP_ALL = '国民健康保険|コクミンケンコウ|国民年金|コクミンネンキン|住民税|ジユウミンゼイ|ジュウミンゼイ|所得税|ショトクゼイ|小規模企業共済|IDECO|イデコ|ふるさと納税|さとふる|ふるなび|楽天ふるさと|生命保険|セイメイホケン';

  const ACC_MAP = {};
  ACCOUNTS.forEach((a) => (ACC_MAP[a.id] = a));
  const norm = (s) => (HC.util ? HC.util.norm(s) : String(s).toLowerCase());

  // 辞書を照合用に展開（見出し語は先頭の語）
  let INDEX = null;
  function index() {
    if (INDEX) return INDEX;
    INDEX = [];
    DICT.forEach((d, i) => {
      d[0].split('|').forEach((w) => {
        if (w[0] === '~') return; // 検索専用（ほかの語の一部に含まれやすい）
        const k = norm(w.replace(/（[^）]*）/g, ''));
        if (k.length >= 2) INDEX.push({ k, i, w });
      });
    });
    VENDORS.forEach((v) => {
      v[0].split('|').forEach((w) => {
        const k = norm(w);
        if (k.length >= 2) INDEX.push({ k, v: v[1], w });
      });
    });
    INDEX.sort((a, b) => b.k.length - a.k.length);
    return INDEX;
  }

  /** 摘要などの文字列から科目を推測。{ a: 科目id|null, note, word, x(経費にならない) } か null */
  function suggest(text, kind) {
    const t = norm(text);
    if (!t) return null;
    for (const e of index()) {
      if (t.indexOf(e.k) < 0) continue;
      if (e.v) {
        if (kind === 'in') continue;
        return { a: e.v, word: e.w, note: '' };
      }
      const d = DICT[e.i];
      const acc = d[1] && ACC_MAP[d[1]];
      if (kind === 'in' && acc && acc.kind !== 'income') continue;
      if (kind === 'out' && acc && acc.kind === 'income') continue;
      if (kind === 'in' && !acc) continue;
      return { a: d[1], word: d[0].split('|')[0], note: d[2], x: !d[1], flags: d[3] || '' };
    }
    return null;
  }

  /** 取り込まない方がよさそうな明細か（src='bank' のときはカードの引き落とし・ATMなども） */
  const reOf = (words) => new RegExp(words.split('|').map((w) => norm(w).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|'));
  let SKIP_RE = null;
  function looksSkip(text, src) {
    if (!SKIP_RE) SKIP_RE = { all: reOf(SKIP_ALL), bank: reOf(SKIP_BANK) };
    const t = norm(text);
    return SKIP_RE.all.test(t) || (src === 'bank' && SKIP_RE.bank.test(t));
  }

  /** 科目さがし：キーワードで辞書を検索（見出し語・別名・科目名・説明に含まれるか） */
  function search(q) {
    const t = norm(q);
    if (!t) return DICT.map((d, i) => i);
    const hits = [];
    DICT.forEach((d, i) => {
      const words = d[0].split('|').map((w) => norm(w.replace(/^~/, '')));
      const acc = d[1] ? ACC_MAP[d[1]].name : '経費にならない';
      let score = 0;
      if (words[0] === t) score = 100;
      else if (words.some((w) => w === t)) score = 90;
      else if (words.some((w) => w.indexOf(t) >= 0)) score = 60;
      else if (words.some((w) => t.indexOf(w) >= 0 && w.length >= 2)) score = 50;
      else if (norm(acc).indexOf(t) >= 0) score = 30;
      else if (norm(d[2]).indexOf(t) >= 0) score = 10;
      if (score) hits.push([score, i]);
    });
    hits.sort((a, b) => b[0] - a[0] || a[1] - b[1]);
    return hits.map((h) => h[1]);
  }

  HC.kamoku = { ACCOUNTS, ACC_MAP, FORM_ROWS, BLANK_ROWS, WHITE_MAIN, DICT, VENDORS, suggest, looksSkip, search };
})(typeof window !== 'undefined' ? window : globalThis);

/* ---- calc.js ---- */
/* ひとり帳簿 — 計算（家事按分・減価償却・在庫・年間の集計と決算書の数字） */
(function (g) {
  'use strict';
  const HC = (g.HC = g.HC || {});
  const U = HC.util;
  const K = HC.kamoku;

  // ---------------- 減価償却 ----------------
  // 根拠：国税庁 タックスアンサー No.2100・No.2106（令和8年4月1日現在法令等）

  /** 定額法の償却率（平成19年4月1日以後に取得したもの）：1 ÷ 耐用年数 を小数点以下3位に切り上げ */
  function slRate(life) {
    life = Math.floor(U.num(life));
    if (life < 2) return 1;
    return Math.ceil(1000 / life) / 1000;
  }

  /** 中古資産の耐用年数（簡便法）。法定耐用年数と経過年数（年・月）から。2年未満は2年 */
  function usedLife(legal, elapsedMonths) {
    const L = U.num(legal) * 12;
    const E = U.num(elapsedMonths);
    let m;
    if (E >= L) m = L * 0.2;
    else m = L - E + E * 0.2;
    return Math.max(2, Math.floor(m / 12));
  }

  /** 青色申告の少額減価償却資産の特例の上限（取得日で変わる）。対象外なら 0 */
  function smallLimit(acquired) {
    const d = String(acquired || '');
    if (!d) return 0;
    if (d < '2006-04') return 0;
    if (d < '2026-04') return 300000; // 令和8年3月31日まで：30万円未満
    if (d < '2029-04') return 400000; // 令和8年4月1日〜令和11年3月31日：40万円未満
    return 0;
  }

  /** 取得価額から、あてはまる扱いの候補 */
  function assetTreatments(cost, acquired, blue) {
    cost = U.num(cost);
    const out = [];
    if (cost < 100000) return [{ id: 'expense', label: '全額をその年の経費（消耗品費など）', ok: true }];
    const lim = smallLimit(acquired);
    if (cost < 200000) out.push({ id: 'lump3', label: '一括償却資産（3年で均等に経費）', ok: true });
    if (lim) {
      out.push({
        id: 'small',
        label: '少額減価償却資産の特例（全額をその年の経費）',
        ok: !!blue && cost < lim,
        why: !blue ? '青色申告の人だけ' : cost >= lim ? (lim / 10000) + '万円未満のものだけ' : ''
      });
    }
    out.push({ id: 'sl', label: '定額法で減価償却（耐用年数で分けて経費）', ok: true });
    return out;
  }

  /**
   * 1つの資産の年ごとの償却表
   * asset: { cost, start:'YYYY-MM'（事業に使い始めた年月）, life, method:'sl'|'lump3'|'small', biz:%（事業専用割合）,
   *          openYear?, openBook?（この年の期首の未償却残高がわかっている場合） }
   * 戻り値: [{ year, months, open, dep, biz, close }]（dep は償却費の合計、biz は必要経費に入れる額）
   */
  function schedule(asset, untilYear) {
    const cost = Math.floor(U.num(asset.cost));
    const m = /^(\d{4})-(\d{2})/.exec(asset.start || '');
    if (!cost || !m) return [];
    const y0 = +m[1], mon0 = +m[2];
    const bizRate = Math.min(100, Math.max(0, asset.biz == null || asset.biz === '' ? 100 : U.num(asset.biz)));
    const method = asset.method || 'sl';
    const last = untilYear || y0 + 60;
    const rows = [];
    const bizOf = (dep) => Math.floor((dep * bizRate) / 100);

    if (method === 'small') {
      rows.push({ year: y0, months: 13 - mon0, open: cost, dep: cost, biz: bizOf(cost), close: 0 });
      return rows.filter((r) => r.year <= last);
    }
    if (method === 'lump3') {
      const a = Math.ceil(cost / 3);
      const parts = [a, a, cost - a * 2];
      let open = cost;
      parts.forEach((dep, i) => {
        rows.push({ year: y0 + i, months: 12, open, dep, biz: bizOf(dep), close: open - dep });
        open -= dep;
      });
      return rows.filter((r) => r.year <= last);
    }
    // 定額法：取得価額 × 償却率 × 月数/12（1円未満切り上げ）。最後は備忘価額1円を残す
    const life = Math.max(2, Math.floor(U.num(asset.life) || 0));
    const rate = slRate(life);
    let open = cost;
    let year = y0;
    if (asset.openYear && asset.openBook != null && asset.openBook !== '' && +asset.openYear > y0) {
      year = +asset.openYear;
      open = Math.floor(U.num(asset.openBook));
    }
    for (; year <= last && open > 1; year++) {
      const months = year === y0 ? 13 - mon0 : 12;
      let dep = Math.ceil((cost * rate * months) / 12);
      if (dep > open - 1) dep = open - 1;
      rows.push({ year, months, open, dep, biz: bizOf(dep), close: open - dep });
      open -= dep;
    }
    return rows;
  }

  /** その年の行（なければ null） */
  function depOfYear(asset, year) {
    return schedule(asset, year).find((r) => r.year === year) || null;
  }

  /** 報酬の源泉徴収税額（100万円以下 10.21%、超える部分 20.42%。1円未満切り捨て） */
  function whTax(amount) {
    const x = Math.floor(U.num(amount));
    if (x <= 0) return 0;
    if (x <= 1000000) return Math.floor(x * 0.1021);
    return Math.floor((x - 1000000) * 0.2042) + 102100;
  }

  // ---------------- 家事按分 ----------------
  /** 取引の事業分の割合（%）。取引ごとの指定 > 科目の既定 > 100 */
  function rateOf(entry, acc) {
    if (entry.r != null && entry.r !== '') return Math.min(100, Math.max(0, U.num(entry.r)));
    if (acc && acc.rate != null && acc.rate !== '') return Math.min(100, Math.max(0, U.num(acc.rate)));
    return 100;
  }

  /** 家事按分の計算（ツール用）。total × 割合（%）、1円未満切り捨て */
  function apportion(total, pct) {
    return Math.floor((U.num(total) * Math.min(100, Math.max(0, U.num(pct)))) / 100);
  }

  // ---------------- 在庫（棚卸） ----------------
  /** 品目 [{ unit, qty }] から期末の棚卸高（1円未満切り捨て） */
  function stockValue(items) {
    return (items || []).reduce((s, it) => s + Math.floor(U.num(it.unit) * U.num(it.qty)), 0);
  }
  /** 1個あたりの原価（製作費 ÷ 製作数。1円未満切り捨て） */
  function unitCost(cost, made) {
    made = U.num(made);
    return made > 0 ? Math.floor(U.num(cost) / made) : 0;
  }

  // ---------------- 年間の集計 ----------------
  /**
   * data: { entries, accounts, assets, inv, partners, settings }
   *   entries: [{ id, d:'YYYY-MM-DD', k:'in'|'out', a, amt, p, memo, m, r, wh, fee }]
   *   accounts: [{ id, name, kind, std, rate, hidden }]
   *   assets: 上の schedule の asset と { id, name }
   *   inv: { [year]: { open, close, items } }
   * form: 'white'（収支内訳書）| 'blue'（青色申告決算書・10万円控除）| 'misc'（雑所得）
   */
  function summarize(data, year, form) {
    year = +year;
    const ys = String(year);
    const accs = data.accounts || K.ACCOUNTS;
    const accMap = {};
    accs.forEach((a) => (accMap[a.id] = a));
    const entries = (data.entries || []).filter((e) => String(e.d || '').slice(0, 4) === ys);
    const feeAcc = accMap.fees ? 'fees' : 'misc';

    const income = { sales: 0, other_income: 0, kaji: 0 };
    let purchase = 0;
    const exp = {}; // id → { paid, groups: { rate: sum } }
    const monthly = Array.from({ length: 12 }, () => ({ sales: 0, purchase: 0 }));
    const salesBy = {}, purchaseBy = {}, whBy = {}, rentBy = {};
    let wh = 0, fees = 0, count = 0;
    const addExp = (id, amt, rate) => {
      const o = (exp[id] = exp[id] || { paid: 0, groups: {} });
      o.paid += amt;
      o.groups[rate] = (o.groups[rate] || 0) + amt;
    };

    for (const e of entries) {
      const amt = Math.round(U.num(e.amt));
      if (!amt && !U.num(e.fee)) continue;
      count++;
      const mon = +e.d.slice(5, 7) - 1;
      const acc = accMap[e.a];
      const p = (e.p || '').trim() || '（取引先なし）';
      if (e.k === 'in') {
        const id = acc && acc.kind === 'income' ? acc.id : 'sales';
        income[id] = (income[id] || 0) + amt;
        if (id === 'sales') {
          monthly[mon].sales += amt;
          salesBy[p] = (salesBy[p] || 0) + amt;
        }
        const w = Math.round(U.num(e.wh));
        if (w) {
          wh += w;
          const o = (whBy[p] = whBy[p] || { revenue: 0, wh: 0 });
          o.revenue += amt;
          o.wh += w;
        }
        const f = Math.round(U.num(e.fee));
        if (f) {
          fees += f;
          addExp(feeAcc, f, 100);
        }
      } else {
        if (acc && acc.kind === 'cogs') {
          purchase += amt;
          monthly[mon].purchase += amt;
          purchaseBy[p] = (purchaseBy[p] || 0) + amt;
          continue;
        }
        const id = acc && acc.kind === 'expense' && acc.id !== 'depreciation' ? acc.id : acc && acc.id === 'depreciation' ? 'depreciation' : 'misc';
        const rate = rateOf(e, acc);
        addExp(id, amt, rate);
        if (id === 'rent') {
          const o = (rentBy[p] = rentBy[p] || { paid: 0, groups: {} });
          o.paid += amt;
          o.groups[rate] = (o.groups[rate] || 0) + amt;
        }
      }
    }

    // 経費の事業分：科目ごと・割合ごとに合計してから按分（1円未満切り捨て）
    const bizOf = (groups) => Object.keys(groups).reduce((s, r) => s + Math.floor((groups[r] * Number(r)) / 100), 0);
    const expBiz = {};
    Object.keys(exp).forEach((id) => (expBiz[id] = bizOf(exp[id].groups)));
    // 地代家賃は「地代家賃の内訳」の合計と一致させる（支払先ごとに按分）
    Object.keys(rentBy).forEach((k) => (rentBy[k].biz = bizOf(rentBy[k].groups)));
    if (exp.rent) expBiz.rent = Object.keys(rentBy).reduce((s, k) => s + rentBy[k].biz, 0);

    // 減価償却
    const depRows = [];
    let depTotal = 0, smallTotal = 0;
    (data.assets || []).forEach((a) => {
      const r = depOfYear(a, year);
      if (!r) return;
      depRows.push({ asset: a, row: r, rate: a.method === 'sl' || !a.method ? slRate(a.life) : null });
      depTotal += r.biz;
      if (a.method === 'small') smallTotal += Math.floor(U.num(a.cost));
    });
    if (depTotal) {
      expBiz.depreciation = (expBiz.depreciation || 0) + depTotal;
      exp.depreciation = exp.depreciation || { paid: 0, groups: {} };
    }

    // 在庫
    const inv = (data.inv || {})[ys] || {};
    const prev = (data.inv || {})[String(year - 1)] || {};
    const open = inv.open != null && inv.open !== '' ? Math.floor(U.num(inv.open)) : prev.close != null && prev.close !== '' ? Math.floor(U.num(prev.close)) : 0;
    const close = inv.items && inv.items.length ? stockValue(inv.items) : Math.floor(U.num(inv.close));
    const cogs = { open, purchase, subtotal: open + purchase, close, cost: open + purchase - close };

    const incomeTotal = income.sales + income.other_income + income.kaji;
    const gross = incomeTotal - cogs.cost;

    // 経費の行（決算書の並び）
    const rows = formRows(form, accs, expBiz, exp);
    const expTotal = rows.list.reduce((s, r) => s + r.amount, 0);
    const profit = gross - expTotal;
    const blueDeduction = form === 'blue' ? Math.max(0, Math.min(100000, profit)) : 0;

    const top = (obj, n) => {
      const list = Object.keys(obj).map((k) => ({ name: k, amount: typeof obj[k] === 'number' ? obj[k] : obj[k].paid }));
      list.sort((a, b) => b.amount - a.amount);
      const head = list.slice(0, n);
      const rest = list.slice(n).reduce((s, x) => s + x.amount, 0);
      return { head, rest, total: list.reduce((s, x) => s + x.amount, 0) };
    };

    return {
      year,
      form,
      count,
      income: Object.assign({ total: incomeTotal }, income),
      cogs,
      gross,
      rows: rows.list,
      merged: rows.merged,
      whiteOther: rows.whiteOther,
      expTotal,
      profit,
      blueDeduction,
      net: profit - blueDeduction,
      monthly,
      sales: top(salesBy, 4),
      purchases: top(purchaseBy, 4),
      rent: Object.keys(rentBy).map((k) => ({ name: k, paid: rentBy[k].paid, biz: rentBy[k].biz })),
      wh: { total: wh, by: Object.keys(whBy).map((k) => ({ name: k, revenue: whBy[k].revenue, wh: whBy[k].wh })).sort((a, b) => b.revenue - a.revenue) },
      fees,
      dep: { rows: depRows, total: depTotal, small: smallTotal },
      expPaid: Object.keys(exp).reduce((o, id) => ((o[id] = exp[id].paid), o), {}),
      expBiz
    };
  }

  /** 経費を決算書の並びに。空欄に入りきらない独自科目は、金額の小さい順に雑費へまとめる */
  function formRows(form, accs, expBiz, exp) {
    const order = K.FORM_ROWS[form === 'white' || form === 'misc' ? 'white' : 'blue'];
    const blanks = K.BLANK_ROWS[form === 'white' || form === 'misc' ? 'white' : 'blue'];
    const accMap = {};
    accs.forEach((a) => (accMap[a.id] = a));
    const nameOf = (id) => (accMap[id] ? accMap[id].name : (K.ACC_MAP[id] || { name: id }).name);
    const customs = accs
      .filter((a) => a.kind === 'expense' && !a.std && (expBiz[a.id] || 0) !== 0)
      .map((a) => ({ id: a.id, name: a.name, amount: expBiz[a.id] || 0 }))
      .sort((a, b) => b.amount - a.amount);
    // 標準科目でも accounts から消えた id の金額は雑費へ
    const known = new Set(order.concat(customs.map((c) => c.id)));
    let toMisc = 0;
    const merged = [];
    Object.keys(expBiz).forEach((id) => {
      if (!known.has(id) && id !== 'misc' && expBiz[id]) {
        toMisc += expBiz[id];
        merged.push({ id, name: nameOf(id), amount: expBiz[id] });
      }
    });
    const inBlank = customs.slice(0, blanks);
    customs.slice(blanks).forEach((c) => {
      toMisc += c.amount;
      merged.push(c);
    });
    const list = [];
    order.forEach((id) => {
      if (id === '*') {
        inBlank.forEach((c) => list.push({ id: c.id, name: c.name, amount: c.amount, blank: true, paid: exp[c.id] ? exp[c.id].paid : 0 }));
        return;
      }
      const amount = (expBiz[id] || 0) + (id === 'misc' ? toMisc : 0);
      list.push({ id, name: nameOf(id), amount, std: true, paid: exp[id] ? exp[id].paid + (id === 'misc' ? toMisc : 0) : id === 'misc' ? toMisc : 0 });
    });
    // 収支内訳書の「その他の経費」の小計
    const whiteOther = list.filter((r) => K.WHITE_MAIN.indexOf(r.id) < 0).reduce((s, r) => s + r.amount, 0);
    return { list, merged, whiteOther };
  }

  /** 気をつけたいこと（警告）の一覧 */
  function checks(sum, data, settings) {
    const out = [];
    const blue = settings && settings.form === 'blue';
    if (sum.cogs.purchase && !sum.cogs.close && !((data.inv || {})[String(sum.year)] || {}).checked) {
      out.push({ lv: 'warn', t: '仕入がありますが、年末の在庫（期末棚卸高）が0円です。売れ残りがあれば「年末の整理」で入力してください。' });
    }
    if (!blue && sum.dep.small) {
      out.push({ lv: 'warn', t: '「少額減価償却資産の特例」は青色申告の人だけが使えます。白色申告なら一括償却資産か定額法にしてください。' });
    }
    if (blue && sum.dep.small > 3000000) {
      out.push({ lv: 'warn', t: '少額減価償却資産の特例は、1年で合計300万円までです（いまの合計 ' + U.yen(sum.dep.small) + '円）。' });
    }
    if (sum.merged && sum.merged.length) {
      out.push({ lv: 'info', t: '決算書の空欄が足りないため、' + sum.merged.map((m) => m.name + '（' + U.yen(m.amount) + '円）').join('・') + ' を「雑費」にまとめました。' });
    }
    const miscRow = sum.rows.find((r) => r.id === 'misc');
    if (miscRow && sum.expTotal && miscRow.amount > sum.expTotal * 0.2 && miscRow.amount > 100000) {
      out.push({ lv: 'info', t: '雑費が経費全体の2割をこえています。中身に合う科目があれば、そちらに分けておくと説明しやすくなります。' });
    }
    if (sum.profit < 0) {
      out.push({ lv: 'info', t: '赤字です。青色申告なら、赤字（純損失）を翌年以後3年間くり越せます（確定申告書の第四表）。' });
    }
    const unbased = (data.accounts || []).filter((a) => a.kind === 'expense' && a.rate != null && a.rate !== '' && U.num(a.rate) < 100 && !(a.basis || '').trim() && sum.expPaid[a.id]);
    if (unbased.length) {
      out.push({ lv: 'info', t: '家事按分の根拠（割合の理由）のメモがない科目があります：' + unbased.map((a) => a.name).join('・') + '。「設定」の科目で書いておくと、聞かれたときに説明できます。' });
    }
    return out;
  }

  HC.calc = { whTax, slRate, usedLife, smallLimit, assetTreatments, schedule, depOfYear, rateOf, apportion, stockValue, unitCost, summarize, formRows, checks };
})(typeof window !== 'undefined' ? window : globalThis);

/* ---- tools.js ---- */
/* ひとり帳簿 — Web版の計算ツール（勘定科目さがし・減価償却・家事按分・在庫と売上原価） */
(function (g) {
  'use strict';
  const HC = (g.HC = g.HC || {});
  const U = HC.util;
  const K = HC.kamoku;
  const C = HC.calc;
  const { esc, yen } = U;

  // 入力値を localStorage に残す（再訪時に戻す）
  function persist(el, key) {
    const saved = HC.store.get('tool:' + key, {}) || {};
    el.querySelectorAll('[name]').forEach((inp) => {
      if (!(inp.name in saved)) return;
      if (inp.type === 'radio') inp.checked = inp.value === saved[inp.name];
      else if (inp.type === 'checkbox') inp.checked = !!saved[inp.name];
      else inp.value = saved[inp.name];
    });
    const save = U.debounce(() => {
      const o = {};
      el.querySelectorAll('[name]').forEach((inp) => {
        if (inp.type === 'radio') {
          if (inp.checked) o[inp.name] = inp.value;
        } else if (inp.type === 'checkbox') o[inp.name] = inp.checked;
        else o[inp.name] = inp.value;
      });
      HC.store.set('tool:' + key, o);
    }, 300);
    el.addEventListener('input', save);
    el.addEventListener('change', save);
  }
  const val = (el, name) => {
    const inp = el.querySelector('[name="' + name + '"]');
    if (!inp) return '';
    if (inp.type === 'radio') {
      const c = el.querySelector('[name="' + name + '"]:checked');
      return c ? c.value : '';
    }
    if (inp.type === 'checkbox') return inp.checked;
    return inp.value;
  };

  // ---------------- 勘定科目さがし ----------------
  function mountKamoku(el) {
    const input = el.querySelector('#km-q');
    const best = el.querySelector('#km-best');
    const count = el.querySelector('#km-count');
    const list = document.querySelector('#km-all');
    if (!input || !list) return;
    const items = Array.from(list.querySelectorAll('.km-item'));
    const groups = Array.from(list.querySelectorAll('.km-group'));
    function run() {
      const q = input.value.trim();
      const hits = K.search(q);
      const set = new Set(hits);
      items.forEach((it) => (it.hidden = !!q && !set.has(+it.dataset.i)));
      groups.forEach((gr) => (gr.hidden = !!q && !gr.querySelector('.km-item:not([hidden])')));
      count.textContent = q ? hits.length + '件' : K.DICT.length + '語';
      if (!q) {
        best.innerHTML = '';
        return;
      }
      if (!hits.length) {
        best.innerHTML = '<div class="km-best is-none"><b>「' + esc(q) + '」は辞書にありませんでした。</b><p>近い言葉（「〇〇代」「〇〇の購入」など）で探すか、下の科目の説明から近いものを選んでください。どれにも当てはまらない少額のものは「雑費」でも大丈夫です。迷ったら、毎年同じ科目に入れることが大切です。</p></div>';
        return;
      }
      const d = K.DICT[hits[0]];
      const acc = d[1] ? K.ACC_MAP[d[1]] : null;
      const flags = d[3] || '';
      best.innerHTML =
        '<div class="km-best' + (acc ? '' : ' is-x') + '"><span class="km-q">' + esc(d[0].split('|')[0].replace(/^~/, '')) + ' は</span>' +
        '<span class="km-a">' + (acc ? esc(acc.name) : '経費になりません') + '</span>' +
        '<p>' + esc(d[2]) + '</p>' +
        (flags.includes('h') ? '<p class="km-tip">🏠 私用と兼ねるなら<a href="' + base() + 'tools/anbun/">家事按分</a>して、事業で使う分だけを経費に。</p>' : '') +
        (flags.includes('f') ? '<p class="km-tip">💻 10万円以上なら<a href="' + base() + 'tools/genka/">減価償却</a>（固定資産）になります。</p>' : '') +
        (acc && acc.desc ? '<p class="km-acc">' + esc(acc.name) + 'とは：' + esc(acc.desc) + '</p>' : '') +
        '</div>';
    }
    input.addEventListener('input', U.debounce(run, 120));
    el.querySelectorAll('[data-q]').forEach((b) =>
      b.addEventListener('click', () => {
        input.value = b.dataset.q;
        run();
        input.focus();
      })
    );
    const qs = new URLSearchParams(location.search).get('q');
    if (qs) input.value = qs;
    run();
  }
  const base = () => document.documentElement.dataset.base || './';

  // ---------------- 減価償却 ----------------
  const LIVES = [
    ['パソコン（サーバー用以外）', 4], ['ソフトウェア（自分で使うもの）', 5], ['カメラ', 5], ['テレビ・オーディオ機器', 5], ['複写機・その他の事務機器', 5],
    ['冷蔵庫・洗濯機などの家電', 6], ['エアコン', 6], ['事務机・いす（金属製）', 15], ['事務机・いす（金属製以外）', 8], ['看板', 3],
    ['普通自動車（新車）', 6], ['軽自動車（新車）', 4], ['二輪・三輪自動車', 3], ['自転車', 2]
  ];
  function mountGenka(el) {
    const now = new Date();
    el.innerHTML =
      '<div class="tool-grid"><form class="card tool-form" onsubmit="return false">' +
      '<label class="f"><span class="f-label">取得価額（税込）<small>本体＋送料・設置費など</small></span><input class="in num in-xl" name="cost" inputmode="decimal" value="258000"></label>' +
      '<div class="grid g2"><label class="f"><span class="f-label">事業に使い始めた年月</span><input class="in" type="month" name="start" value="' + now.getFullYear() + '-' + U.pad(now.getMonth() + 1) + '"></label>' +
      '<label class="f"><span class="f-label">事業で使う割合</span><span class="pct"><input class="in num" name="biz" value="100" inputmode="decimal"><span>%</span></span></label></div>' +
      '<div class="f"><span class="f-label">申告の種類</span><div class="seg"><label><input type="radio" name="blue" value="0" checked><span>白色申告</span></label><label><input type="radio" name="blue" value="1"><span>青色申告</span></label></div></div>' +
      '<div class="grid g2"><label class="f"><span class="f-label">耐用年数（一覧から）</span><select class="in" name="pick"><option value="">選ぶ…</option>' + LIVES.map((l) => '<option value="' + l[1] + '">' + esc(l[0]) + '：' + l[1] + '年</option>').join('') + '</select></label>' +
      '<label class="f"><span class="f-label">耐用年数（年）</span><input class="in num" name="life" value="4" inputmode="numeric"></label></div>' +
      '<label class="tg"><input type="checkbox" name="used"><span class="tg-ui"></span>中古で買った</label>' +
      '<div class="grid g2 used-row" hidden><label class="f"><span class="f-label">新品の耐用年数</span><input class="in num" name="legal" value="6" inputmode="numeric"></label><label class="f"><span class="f-label">買う前に使われていた期間（月）</span><input class="in num" name="elapsed" value="30" inputmode="numeric"></label></div>' +
      '</form><div class="tool-result" id="gk-out" aria-live="polite"></div></div>';
    const out = el.querySelector('#gk-out');
    persist(el, 'genka');
    function run() {
      const cost = Math.floor(U.num(val(el, 'cost')));
      const start = val(el, 'start');
      const blue = val(el, 'blue') === '1';
      const biz = Math.min(100, Math.max(0, U.num(val(el, 'biz') === '' ? 100 : val(el, 'biz'))));
      const used = val(el, 'used');
      el.querySelector('.used-row').hidden = !used;
      let life = Math.max(2, Math.floor(U.num(val(el, 'life'))) || 2);
      let usedNote = '';
      if (used) {
        const legal = U.num(val(el, 'legal'));
        if (legal) {
          life = C.usedLife(legal, U.num(val(el, 'elapsed')));
          usedNote = '中古の耐用年数（簡便法）：' + life + '年';
        }
      }
      if (!cost || !/^\d{4}-\d{2}$/.test(start)) {
        out.innerHTML = '<div class="card"><p class="empty">取得価額と使い始めた年月を入れてください。</p></div>';
        return;
      }
      const y0 = +start.slice(0, 4);
      const opts = C.assetTreatments(cost, start, blue);
      if (opts[0].id === 'expense') {
        out.innerHTML =
          '<div class="card"><div class="big-kv"><span>' + y0 + '年の経費</span><b>' + yen(Math.floor((cost * biz) / 100)) + '<small>円</small></b></div>' +
          '<p class="callout is-tip"><b>10万円未満</b>なので、減価償却はいりません。買った年に全額を「消耗品費」などで経費にします' + (biz < 100 ? '（事業で使う割合 ' + biz + '% 分）' : '') + '。</p></div>';
        return;
      }
      const cards = opts.map((o) => {
        const sch = C.schedule({ cost, start, life, method: o.id, biz }, y0 + 60);
        const first = sch[0] ? sch[0].biz : 0;
        return { o, sch, first };
      });
      const sl = cards.find((c) => c.o.id === 'sl');
      out.innerHTML =
        '<div class="card"><div class="out-h">選べる方法と、' + y0 + '年に経費になる額</div><div class="mcards">' +
        cards
          .map(
            (c) =>
              '<div class="mcard' + (c.o.ok ? '' : ' is-na') + (c.o.ok && c.o.id !== 'sl' ? ' is-best' : '') + '">' + (c.o.ok && c.o.id !== 'sl' ? '<span class="mcard-badge">早く経費にできる</span>' : '') +
              '<div class="mcard-h">' + esc(c.o.label.replace(/（.*$/, '')) + '</div>' +
              '<div class="mcard-v">' + (c.o.ok ? yen(c.first) + '<small>円</small>' : '<span class="na">使えません</span>') + '</div>' +
              '<div class="mcard-n">' + esc(c.o.ok ? c.o.label.replace(/^[^（]*/, '') : c.o.why) + '</div></div>'
          )
          .join('') +
        '</div>' + (usedNote ? '<p class="hint">' + esc(usedNote) + '</p>' : '') + '</div>' +
        '<div class="card"><div class="out-h">定額法（耐用年数 ' + life + '年・償却率 ' + C.slRate(life).toFixed(3) + '）で償却した場合</div>' +
        '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>年</th><th class="num">月数</th><th class="num">償却費</th><th class="num">経費になる額</th><th class="num">年末の未償却残高</th></tr></thead><tbody>' +
        sl.sch.map((r) => '<tr><td>' + r.year + '年</td><td class="num">' + r.months + '</td><td class="num">' + yen(r.dep) + '</td><td class="num"><b>' + yen(r.biz) + '</b></td><td class="num">' + yen(r.close) + '</td></tr>').join('') +
        '</tbody></table></div>' +
        '<p class="formula">取得価額 × 償却率 × 使った月数 ÷ 12（1円未満切り上げ）。最後の年は1円（備忘価額）を残します。経費になる額 ＝ 償却費 × 事業で使う割合。</p></div>';
    }
    el.addEventListener('input', (ev) => {
      if (ev.target.name === 'pick' && ev.target.value) el.querySelector('[name="life"]').value = ev.target.value;
      run();
    });
    el.addEventListener('change', run);
    run();
  }

  // ---------------- 家事按分 ----------------
  const ANBUN_MODES = [
    ['area', '床面積で（家賃・電気代など）'],
    ['time', '使う時間で（スマホ・ネット・電気代など）'],
    ['dist', '走った距離で（車・ガソリン代）'],
    ['pct', '割合を直接入れる']
  ];
  function mountAnbun(el) {
    el.innerHTML =
      '<div class="tool-grid"><form class="card tool-form" onsubmit="return false">' +
      '<label class="f"><span class="f-label">何の費用？</span><input class="in" name="what" value="家賃" maxlength="20"></label>' +
      '<div class="grid g2"><label class="f"><span class="f-label">1か月の支払額</span><input class="in num in-xl" name="monthly" inputmode="decimal" value="85000"></label>' +
      '<label class="f"><span class="f-label">支払った月数</span><input class="in num" name="months" inputmode="numeric" value="12"></label></div>' +
      '<div class="f"><span class="f-label">割合の決め方</span><div class="seg seg-wrap">' + ANBUN_MODES.map((m, i) => '<label><input type="radio" name="mode" value="' + m[0] + '"' + (i ? '' : ' checked') + '><span>' + esc(m[1]) + '</span></label>').join('') + '</div></div>' +
      '<div class="grid g2" data-mode="area"><label class="f"><span class="f-label">仕事に使う広さ（㎡）</span><input class="in num" name="a1" value="9" inputmode="decimal"></label><label class="f"><span class="f-label">家全体の広さ（㎡）</span><input class="in num" name="a2" value="36" inputmode="decimal"></label></div>' +
      '<div class="grid g3" data-mode="time" hidden><label class="f"><span class="f-label">仕事に使う時間（1日）</span><input class="in num" name="t1" value="8" inputmode="decimal"></label><label class="f"><span class="f-label">仕事をする日（週）</span><input class="in num" name="t2" value="5" inputmode="decimal"></label><label class="f"><span class="f-label">使う時間の合計（1日）</span><input class="in num" name="t3" value="16" inputmode="decimal"></label></div>' +
      '<div class="grid g2" data-mode="dist" hidden><label class="f"><span class="f-label">仕事で走った距離（km）</span><input class="in num" name="d1" value="3000" inputmode="decimal"></label><label class="f"><span class="f-label">全体で走った距離（km）</span><input class="in num" name="d2" value="6000" inputmode="decimal"></label></div>' +
      '<label class="f" data-mode="pct" hidden><span class="f-label">事業で使う割合</span><span class="pct"><input class="in num" name="p" value="50" inputmode="decimal"><span>%</span></span></label>' +
      '</form><div class="tool-result" id="ab-out" aria-live="polite"></div></div>';
    const out = el.querySelector('#ab-out');
    persist(el, 'anbun');
    function run() {
      const mode = val(el, 'mode') || 'area';
      el.querySelectorAll('[data-mode]').forEach((b) => (b.hidden = b.dataset.mode !== mode));
      const what = val(el, 'what') || '費用';
      const monthly = U.num(val(el, 'monthly'));
      const months = U.num(val(el, 'months')) || 12;
      const total = Math.round(monthly * months);
      let pct = 0, basis = '';
      if (mode === 'area') {
        const a1 = U.num(val(el, 'a1')), a2 = U.num(val(el, 'a2'));
        pct = a2 ? (a1 / a2) * 100 : 0;
        basis = '仕事に使う広さ ' + a1 + '㎡ ÷ 全体 ' + a2 + '㎡';
      } else if (mode === 'time') {
        const t1 = U.num(val(el, 't1')), t2 = U.num(val(el, 't2')), t3 = U.num(val(el, 't3'));
        pct = t3 ? ((t1 * t2) / (t3 * 7)) * 100 : 0;
        basis = '仕事 ' + t1 + '時間×週' + t2 + '日 ÷ 使う時間 ' + t3 + '時間×7日';
      } else if (mode === 'dist') {
        const d1 = U.num(val(el, 'd1')), d2 = U.num(val(el, 'd2'));
        pct = d2 ? (d1 / d2) * 100 : 0;
        basis = '仕事で走った ' + yen(d1) + 'km ÷ 全体 ' + yen(d2) + 'km';
      } else {
        pct = U.num(val(el, 'p'));
        basis = '';
      }
      pct = Math.min(100, Math.max(0, Math.floor(pct * 10) / 10));
      const rate = Math.round(pct);
      const biz = C.apportion(total, rate);
      const memo = what + '：' + (basis ? basis + ' = 約' + rate + '%' : '事業で使う割合 ' + rate + '%') + '（' + yen(monthly) + '円 × ' + months + 'か月 = ' + yen(total) + '円 → 経費 ' + yen(biz) + '円）';
      out.innerHTML =
        '<div class="card"><div class="big-kv"><span>経費になる額（1年）</span><b>' + yen(biz) + '<small>円</small></b></div>' +
        '<table class="kv"><tr><th>1年の支払額</th><td>' + yen(total) + '円</td></tr><tr><th>事業で使う割合</th><td>' + rate + '%' + (Math.abs(pct - rate) > 0.05 ? '<small>（' + pct + '%を四捨五入）</small>' : '') + '</td></tr><tr class="is-minus"><th>私用の分（経費にしない）</th><td>' + yen(total - biz) + '円</td></tr></table></div>' +
        '<div class="card"><div class="out-h">割合の理由のメモ（帳簿に残しておく）</div><p class="memo-out" id="ab-memo">' + esc(memo) + '</p><button type="button" class="btn btn-sm btn-ghost" data-act="copy">メモをコピー</button>' +
        '<p class="hint">家事按分に決まった割合はありません。「なぜその割合なのか」を説明できることが大切です。家賃は床面積、電気代は床面積や使う時間、スマホ・ネットは仕事で使う時間、車は走った距離で分けるのが一般的です。</p></div>';
    }
    el.addEventListener('input', run);
    el.addEventListener('change', run);
    el.addEventListener('click', (ev) => {
      if (!ev.target.closest('[data-act="copy"]')) return;
      const t = el.querySelector('#ab-memo').textContent;
      (navigator.clipboard ? navigator.clipboard.writeText(t) : Promise.reject()).then(() => U.toast('コピーしました', 'ok')).catch(() => U.toast('コピーできませんでした。手で選択してください', 'error'));
    });
    run();
  }

  // ---------------- 在庫・売上原価 ----------------
  function mountZaiko(el) {
    const saved = HC.store.get('tool:zaiko', null);
    const st = saved && Array.isArray(saved.items)
      ? saved
      : { open: 18700, sales: '', items: [{ name: '新刊A', cost: 66000, made: 150, left: 38 }, { name: '既刊B', cost: 0, unit: 550, left: 12 }, { name: 'アクリルキーホルダー', cost: 33000, made: 100, left: 41 }] };
    function unitOf(it) {
      return U.num(it.made) > 0 ? C.unitCost(it.cost, it.made) : Math.floor(U.num(it.unit));
    }
    function render() {
      el.innerHTML =
        '<div class="card"><div class="grid g3">' +
        '<label class="f"><span class="f-label">年はじめの在庫（期首棚卸高）<small>去年の年末在庫</small></span><input class="in num" data-k="open" value="' + esc(st.open) + '" inputmode="decimal"></label>' +
        '<label class="f"><span class="f-label">今年の売上<small>（任意）利益のめやすを出す</small></span><input class="in num" data-k="sales" value="' + esc(st.sales) + '" inputmode="decimal"></label></div>' +
        '<h3 class="sub-h">今年作った・仕入れた品と、年末に残った数</h3>' +
        '<div class="tbl-wrap"><table class="tbl zk-tbl"><thead><tr><th>品名</th><th class="num">製作費・仕入額<small>今年の分</small></th><th class="num">作った数</th><th class="num">1個あたりの原価</th><th class="num">年末の在庫数</th><th class="num">在庫の金額</th><th></th></tr></thead><tbody>' +
        st.items
          .map((it, i) => {
            const unit = unitOf(it);
            return (
              '<tr><td><input class="in in-sm" data-i="' + i + '" data-f="name" value="' + esc(it.name) + '"></td>' +
              '<td><input class="in in-sm num" data-i="' + i + '" data-f="cost" value="' + esc(it.cost) + '" inputmode="decimal"></td>' +
              '<td><input class="in in-sm num" data-i="' + i + '" data-f="made" value="' + esc(it.made == null ? '' : it.made) + '" inputmode="decimal" placeholder="—"></td>' +
              '<td class="num">' + (U.num(it.made) > 0 ? yen(unit) : '<input class="in in-sm num" data-i="' + i + '" data-f="unit" value="' + esc(it.unit == null ? '' : it.unit) + '" inputmode="decimal" placeholder="単価">') + '</td>' +
              '<td><input class="in in-sm num" data-i="' + i + '" data-f="left" value="' + esc(it.left) + '" inputmode="decimal"></td>' +
              '<td class="num" data-sum="' + i + '">' + yen(Math.floor(unit * U.num(it.left))) + '</td>' +
              '<td class="row-ops"><button type="button" class="ib ib-del" data-del="' + i + '" title="削除">×</button></td></tr>'
            );
          })
          .join('') +
        '</tbody></table></div>' +
        '<div class="row-set"><button type="button" class="btn btn-sm btn-ghost" data-add>＋ 品目を足す</button></div>' +
        '<p class="hint">前の年から売っている品（既刊など）は、製作費を空欄にして「1個あたりの原価」に去年と同じ単価を入れます。</p></div>' +
        '<div class="card" id="zk-out"></div>';
      result();
    }
    function result() {
      const purchase = st.items.reduce((s, it) => s + Math.floor(U.num(it.cost)), 0);
      const close = st.items.reduce((s, it) => s + Math.floor(unitOf(it) * U.num(it.left)), 0);
      const open = Math.floor(U.num(st.open));
      const cost = open + purchase - close;
      const sales = U.num(st.sales);
      const out = el.querySelector('#zk-out');
      out.innerHTML =
        '<div class="big-kv"><span>売上原価（経費になる額）</span><b>' + yen(cost) + '<small>円</small></b></div>' +
        '<table class="kv"><tr><th>期首棚卸高（年はじめの在庫）</th><td>' + yen(open) + '円</td></tr><tr><th>＋ 仕入金額（今年の製作費・仕入）</th><td>' + yen(purchase) + '円</td></tr><tr class="is-minus"><th>− 期末棚卸高（年末の在庫）</th><td>' + yen(close) + '円</td></tr><tr class="is-total"><th>＝ 売上原価</th><td>' + yen(cost) + '円</td></tr>' +
        (sales ? '<tr><th>売上 − 売上原価（売上総利益）</th><td>' + yen(sales - cost) + '円</td></tr>' : '') + '</table>' +
        '<p class="callout"><b>ポイント</b>：作った（仕入れた）費用は、全部がその年の経費になるわけではありません。年末に売れ残った分（' + yen(close) + '円）は「在庫」として翌年に持ち越し、売れた年の経費になります。</p>';
      st.items.forEach((it, i) => {
        const c = el.querySelector('[data-sum="' + i + '"]');
        if (c) c.textContent = yen(Math.floor(unitOf(it) * U.num(it.left)));
      });
      HC.store.set('tool:zaiko', st);
    }
    el.addEventListener('input', (ev) => {
      const t = ev.target;
      if (t.dataset.k) st[t.dataset.k] = t.value;
      else if (t.dataset.i != null) {
        const it = st.items[+t.dataset.i];
        it[t.dataset.f] = t.value;
        if (t.dataset.f === 'made' || (t.dataset.f === 'cost' && U.num(it.made) > 0)) {
          const pos = t.selectionStart;
          render();
          const again = el.querySelector('[data-i="' + t.dataset.i + '"][data-f="' + t.dataset.f + '"]');
          if (again) {
            again.focus();
            try {
              again.setSelectionRange(pos, pos);
            } catch (e) { /* noop */ }
          }
          return;
        }
      }
      result();
    });
    el.addEventListener('click', (ev) => {
      if (ev.target.closest('[data-add]')) {
        st.items.push({ name: '', cost: '', made: '', left: '' });
        render();
      } else if (ev.target.closest('[data-del]')) {
        st.items.splice(+ev.target.closest('[data-del]').dataset.del, 1);
        render();
      }
    });
    render();
  }

  HC.tools = { mountKamoku, mountGenka, mountAnbun, mountZaiko };
})(typeof window !== 'undefined' ? window : globalThis);

/* ---- site.js ---- */
/* ひとり帳簿 — Web版サイトの起動処理（ツールのマウント・Pro導線・広告・アフィリエイト） */
(function (g) {
  'use strict';
  const HC = (g.HC = g.HC || {});
  const U = HC.util;
  const CFG = g.HC_CONFIG || {};

  // ---- アフィリエイト ----
  function renderAffiliates(scope) {
    (scope || document).querySelectorAll('[data-aff]').forEach((box) => {
      if (box.dataset.done) return;
      const list = (CFG.affiliates && CFG.affiliates[box.dataset.aff]) || [];
      if (!list.length) return;
      box.dataset.done = '1';
      const anyAd = list.some((a) => a.affiliate);
      box.classList.add('aff');
      box.innerHTML =
        (box.dataset.affTitle ? '<div class="aff-h">' + U.esc(box.dataset.affTitle) + (anyAd ? '<span class="pr">PR</span>' : '') + '</div>' : '') +
        '<div class="aff-list">' +
        list
          .map(
            (a) =>
              '<a class="aff-item" href="' + U.esc(a.url) + '" target="_blank" rel="' + (a.affiliate ? 'sponsored ' : '') + 'noopener">' +
              '<span class="aff-name">' + U.esc(a.name) + (a.affiliate && !box.dataset.affTitle ? '<span class="pr">PR</span>' : '') + '</span>' +
              '<span class="aff-desc">' + U.esc(a.desc) + '</span>' +
              '<span class="aff-cta">' + U.esc(a.cta || '詳しく見る') + ' →</span></a>'
          )
          .join('') +
        '</div>';
      box.addEventListener('click', (e) => {
        const a = e.target.closest('a.aff-item');
        if (a && g.gtag) g.gtag('event', 'affiliate_click', { link_url: a.href, aff_group: box.dataset.aff });
      });
    });
  }

  // ---- 広告（AdSense 設定時のみ） ----
  function renderAds() {
    if (!CFG.adsenseClient) return;
    document.querySelectorAll('.ad-slot').forEach((slot) => {
      const id = (CFG.adsenseSlots || {})[slot.dataset.slot || 'content'];
      if (!id) return;
      slot.innerHTML =
        '<span class="ad-label">スポンサーリンク</span><ins class="adsbygoogle" style="display:block" data-ad-client="' + U.esc(CFG.adsenseClient) +
        '" data-ad-slot="' + U.esc(id) + '" data-ad-format="auto" data-full-width-responsive="true"></ins>';
      try {
        (g.adsbygoogle = g.adsbygoogle || []).push({});
      } catch (e) { /* noop */ }
    });
  }

  // ---- Pro 導線（BOOTHのURLが決まるまでは「準備中」） ----
  function renderPro() {
    document.querySelectorAll('[data-pro-link]').forEach((a) => {
      if (CFG.proUrl) {
        a.href = CFG.proUrl;
        a.target = '_blank';
        a.rel = 'noopener';
      }
    });
    document.querySelectorAll('[data-pro-only]').forEach((el) => (el.hidden = !CFG.proUrl));
    document.querySelectorAll('[data-pro-soon]').forEach((el) => (el.hidden = !!CFG.proUrl));
    document.querySelectorAll('[data-trial-link]').forEach((a) => {
      if (CFG.trialUrl) a.href = CFG.trialUrl;
      else a.hidden = true;
    });
    document.querySelectorAll('[data-pro-price]').forEach((el) => (el.textContent = CFG.proPrice || ''));
  }

  function nav() {
    const btn = document.querySelector('.nav-toggle');
    const menu = document.getElementById('site-nav');
    if (!btn || !menu) return;
    btn.addEventListener('click', () => {
      const open = btn.getAttribute('aria-expanded') !== 'true';
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
      menu.classList.toggle('is-open', open);
    });
  }

  function sw() {
    if (!('serviceWorker' in navigator) || location.protocol !== 'https:') return;
    const base = document.documentElement.dataset.base || './';
    g.addEventListener('load', () => navigator.serviceWorker.register(base + 'sw.js').catch(() => {}));
  }

  function boot() {
    nav();
    renderPro();
    const mounts = { kamoku: 'mountKamoku', genka: 'mountGenka', anbun: 'mountAnbun', zaiko: 'mountZaiko' };
    document.querySelectorAll('[data-app]').forEach((el) => {
      const fn = mounts[el.dataset.app];
      if (fn) HC.tools[fn](el);
    });
    renderAffiliates();
    renderAds();
    sw();
  }

  HC.site = { renderAffiliates };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})(window);
