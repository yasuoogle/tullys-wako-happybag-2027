/**
 * Tully's Coffee 2027 Happy Bag 予約申込フォーム（和光理化学研究所店）
 * スプレッドシートの「拡張機能 > Apps Script」に貼り付け、ウェブアプリとしてデプロイします。
 * フォーム画面は GitHub Pages の index.html です（このスクリプトはデータの受け口）。
 * 「設定」シートの行の位置は動かさないでください（B4〜B13、A15〜F16 を読み込みます）。
 */

const ORDER_SHEET = '申込一覧';
const SETTING_SHEET = '設定';
const STATUS_DEFAULT = '受付';
const STATUS_CANCEL = 'キャンセル';
const ORDER_PREFIX = 'HB2027-';
const MAX_QTY_PER_ITEM = 99; // 入力ミス防止のための上限

/*
 * GitHub Pages のフォーム（index.html）から呼ばれる API です。
 *   GET  ?action=config  … 商品・残数・受付期間などの設定を返す
 *   POST （JSON 本文）     … 申込を記録して申込番号を返す
 */
function doGet(e) {
  try {
    return json_({ ok: true, config: publicConfig_(getConfig_()) });
  } catch (err) {
    return json_({ ok: false, error: '設定を読み込めませんでした。' });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.website) return json_({ ok: true, orderNo: '-', total: 0 }); // ボット対策（隠し項目）
    const r = submitOrder(data);
    return json_(Object.assign({ ok: true }, r));
  } catch (err) {
    return json_({
      ok: false,
      error: String(err && err.message || err).replace(/^Error:\s*/, ''),
      config: safeConfig_()
    });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function safeConfig_() {
  try { return publicConfig_(getConfig_()); } catch (e) { return null; }
}

/** 申込を記録する */
function submitOrder(data) {
  data = data || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const cfg = getConfig_();
    const now = Date.now();
    if (cfg.openAt && now < cfg.openAt) throw new Error('受付開始前です。' + formatDate_(cfg.openAt) + 'から受け付けます。');
    if (cfg.closeAt && now > cfg.closeAt) throw new Error('受付は終了しました。');

    const name = clean_(data.name, 50);
    const email = clean_(data.email, 100);
    const message = clean_(data.message, 500);

    if (!name) throw new Error('氏名を入力してください。');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('メールアドレスを正しく入力してください。');

    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(data.pickup || ''));
    if (!m) throw new Error('引取日を選んでください。');
    const pickup = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    const pickupKey = m[1] + m[2] + m[3];
    if (cfg.pickupFromKey && pickupKey < cfg.pickupFromKey) throw new Error('引取日は' + cfg.pickupFromText + '以降を選んでください。');
    if (cfg.pickupToKey && pickupKey > cfg.pickupToKey) throw new Error('引取日は' + cfg.pickupToText + 'までの日付を選んでください。');

    const qty = cfg.products.map(function (p, i) {
      const n = Number((data.qty || [])[i] || 0);
      if (!Number.isInteger(n) || n < 0 || n > MAX_QTY_PER_ITEM) throw new Error(p.name + 'の個数が正しくありません。');
      if (cfg.perPersonLimit != null && n > cfg.perPersonLimit) throw new Error(p.name + 'はお一人さま' + cfg.perPersonLimit + '個までです。');
      if (p.remaining != null && n > p.remaining) {
        throw new Error(p.remaining === 0 ? p.name + 'は完売しました。' : p.name + 'の残りは' + p.remaining + '個です。');
      }
      return n;
    });
    if (qty.reduce(function (a, b) { return a + b; }, 0) === 0) throw new Error('福袋の個数を1つ以上選んでください。');

    const total = cfg.products.reduce(function (sum, p, i) { return sum + p.price * qty[i]; }, 0);

    const sheet = SpreadsheetApp.getActive().getSheetByName(ORDER_SHEET);
    const orderNo = nextOrderNo_(sheet);
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setNumberFormat('@'); // 先頭の0が消えないよう文字列として保存
    // 列：申込番号 / 申込日時 / 氏名 / メール / 引取日 / 5500個数 / 10000個数 / 合計金額 / 状況 / 通信欄
    sheet.getRange(row, 1, 1, 10).setValues([[
      orderNo, new Date(), name, email, pickup, qty[0] || 0, qty[1] || 0, '', STATUS_DEFAULT, message
    ]]);
    sheet.getRange(row, 8).setFormula('=F' + row + '*設定!$B$15+G' + row + '*設定!$B$16');
    SpreadsheetApp.flush();

    let mailSent = false;
    if (cfg.sendMail) {
      try {
        MailApp.sendEmail({
          to: email,
          subject: '【' + cfg.storeName + '】Happy Bag ご予約を承りました（' + orderNo + '）',
          body: mailBody_(cfg, { orderNo: orderNo, name: name, qty: qty, total: total, pickup: fmtDay_(pickup), message: message }),
          name: cfg.storeName
        });
        mailSent = true;
      } catch (e) {
        console.error('確認メールの送信に失敗: ' + e);
      }
    }
    return { orderNo: orderNo, total: total, mailSent: mailSent };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 内部処理 ---------- */

function getConfig_() {
  const s = SpreadsheetApp.getActive().getSheetByName(SETTING_SHEET);
  const v = s.getRange('B4:B13').getValues().map(function (r) { return r[0]; });
  const from = toTime_(v[4]), to = toTime_(v[9]);
  const products = s.getRange('A15:F16').getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      const stock = r[2] === '' ? null : Number(r[2]);
      const ordered = Number(r[3]) || 0;
      return {
        name: String(r[0]), price: Number(r[1]), image: imageUrl_(r[5]),
        remaining: stock == null ? null : Math.max(0, stock - ordered)
      };
    });
  return {
    storeName: String(v[0] || ''),
    title: String(v[1] || ''),
    openAt: toTime_(v[2]),
    closeAt: toTime_(v[3]),
    pickupFrom: from ? fmtIso_(from) : '',          // date入力の min（yyyy-MM-dd）
    pickupTo: to ? fmtIso_(to) : '',                // date入力の max
    pickupFromKey: from ? fmtIso_(from).replace(/-/g, '') : '',
    pickupToKey: to ? fmtIso_(to).replace(/-/g, '') : '',
    pickupFromText: from ? fmtDay_(new Date(from)) : '',
    pickupToText: to ? fmtDay_(new Date(to)) : '',
    payment: String(v[5] || ''),
    perPersonLimit: v[6] === '' ? null : Number(v[6]),
    sendMail: String(v[7]).trim() === 'する',
    notes: String(v[8] || ''),
    products: products
  };
}

function publicConfig_(cfg) {
  const now = Date.now();
  let status = 'open';
  if (cfg.openAt && now < cfg.openAt) status = 'before';
  else if (cfg.closeAt && now > cfg.closeAt) status = 'closed';
  const o = Object.assign({}, cfg, {
    status: status,
    openAtText: cfg.openAt ? formatDate_(cfg.openAt) : '',
    closeAtText: cfg.closeAt ? formatDate_(cfg.closeAt) : '',
    maxQty: MAX_QTY_PER_ITEM
  });
  return o;
}

/** Googleドライブの共有リンクを <img> で表示できるURLに変換する */
function imageUrl_(v) {
  const u = String(v || '').trim();
  if (!u) return '';
  const m = /\/d\/([\w-]{20,})/.exec(u) || /[?&]id=([\w-]{20,})/.exec(u);
  if (/drive\.google\.com|docs\.google\.com/.test(u) && m) return 'https://lh3.googleusercontent.com/d/' + m[1];
  return /^https:\/\//.test(u) ? u : '';
}

function toTime_(v) {
  if (v instanceof Date) return v.getTime();
  if (v === '' || v == null) return null;
  const d = new Date(String(v).replace(/-/g, '/'));
  return isNaN(d.getTime()) ? null : d.getTime();
}

function fmtIso_(t) {
  return Utilities.formatDate(new Date(t), 'Asia/Tokyo', 'yyyy-MM-dd');
}

function fmtDay_(d) {
  const w = '日月火水木金土'.charAt(Number(Utilities.formatDate(d, 'Asia/Tokyo', 'u')) % 7);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy年M月d日') + '（' + w + '）';
}

function formatDate_(t) {
  return Utilities.formatDate(new Date(t), 'Asia/Tokyo', 'yyyy年M月d日 H:mm');
}

/** 改行以外の制御文字を除き、数式として解釈されないようにする */
function clean_(s, max) {
  let t = String(s == null ? '' : s).replace(/[\u0000-\u0009\u000B-\u001F\u007F]/g, '').trim().slice(0, max);
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}

/** 「HB2027-」＋6桁のランダムな申込番号（既存の番号と重複しないもの）を返す */
function nextOrderNo_(sheet) {
  const last = sheet.getLastRow();
  const used = {};
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, 1).getDisplayValues().forEach(function (r) { used[r[0]] = true; });
  }
  for (let i = 0; i < 50; i++) {
    const no = ORDER_PREFIX + String(Math.floor(Math.random() * 1000000)).padStart(6, '0');
    if (!used[no]) return no;
  }
  throw new Error('申込番号を発行できませんでした。もう一度お試しください。');
}

function mailBody_(cfg, o) {
  const lines = [];
  lines.push(o.name + ' 様', '');
  lines.push('このたびは Happy Bag をご予約いただき、ありがとうございます。', '以下の内容で承りました。', '');
  lines.push('■ 申込番号：' + o.orderNo);
  cfg.products.forEach(function (p, i) {
    if (o.qty[i]) lines.push('■ ' + p.name + '：' + o.qty[i] + '個');
  });
  lines.push('■ 合計金額：' + o.total.toLocaleString('ja-JP') + '円（税込）');
  lines.push('■ 引取日：' + o.pickup);
  if (o.message) lines.push('■ 通信欄：' + o.message.replace(/^'/, ''));
  lines.push('', '【お支払い】', cfg.payment);
  if (cfg.notes) lines.push('', '【ご注意】', cfg.notes);
  lines.push('', 'お受け取りの際に申込番号をお伝えください。', '', cfg.storeName);
  return lines.join('\n');
}
