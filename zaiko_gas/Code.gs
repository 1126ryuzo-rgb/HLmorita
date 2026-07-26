// ============================================================
// 民泊 消耗品 在庫・発注アシスタント  ―  バックエンド (GAS)
// スプレッドシートをDBに、Googleアカウントで権限を判定する
// ============================================================

const TZ = 'Asia/Tokyo';
const SH_ITEMS  = '品目マスタ';
const SH_ORDERS = '発注履歴';
const SH_LOG    = '監査ログ';
const SH_ROLE   = '権限マスタ';

// ─── 発注先マスタ（発注方法つき）───────────────────────────────
const SUPPLIERS = {
  'アスクル':         { method:'オンライン', color:'#2E75B6', contact:'https://www.askul.co.jp/（会員ログイン）', cash:false },
  'JTB商事':          { method:'メール',     color:'#38A169', contact:'order@example-jtb-shoji.co.jp', cash:false },
  '大場金物':         { method:'電話・現金', color:'#DB6D28', contact:'TEL 095-000-0000（担当：大場様）', cash:true },
  'ダイソー・ニトリ': { method:'店頭・現金', color:'#805AD5', contact:'最寄り店舗にて購入', cash:true },
};
const CATEGORIES = ['アメニティ','清掃用品','トイレ・キッチン消耗品','リネン・寝具','備品・設備'];

// ─── 権限（サーバー側で強制）──────────────────────────────────
const PERMS = {
  viewer:  { name:'閲覧者',           canEdit:false, canDraft:false, canConfirm:false },
  staff:   { name:'発注担当者',       canEdit:false, canDraft:true,  canConfirm:false },
  manager: { name:'承認者（責任者）', canEdit:true,  canDraft:true,  canConfirm:true },
};

// ─── 初期サンプルデータ ───────────────────────────────────────
// [コード, 品名, 分類, 発注先, 単位, 在庫, 発注点, 適正在庫, 発注ロット, 単価]
const SEED_ITEMS = [
  ['A-101','歯ブラシセット','アメニティ','JTB商事','セット',18,30,80,50,42],
  ['A-102','シャンプー（個包装）','アメニティ','JTB商事','個',45,60,150,100,28],
  ['A-103','ボディソープ（個包装）','アメニティ','JTB商事','個',52,60,150,100,28],
  ['A-104','使い捨てスリッパ','アメニティ','JTB商事','足',12,40,120,60,35],
  ['A-105','カミソリ','アメニティ','JTB商事','本',70,50,120,100,18],
  ['C-201','マルチクリーナー','清掃用品','アスクル','本',3,5,12,6,398],
  ['C-202','使い捨て手袋(100枚)','清掃用品','アスクル','箱',2,4,10,5,540],
  ['C-203','アルコール除菌スプレー','清掃用品','アスクル','本',6,8,18,12,420],
  ['C-204','クロス（マイクロファイバー）','清掃用品','ダイソー・ニトリ','枚',9,15,40,20,110],
  ['C-205','スポンジ','清掃用品','ダイソー・ニトリ','個',22,20,50,30,55],
  ['T-301','トイレットペーパー(12R)','トイレ・キッチン消耗品','アスクル','パック',4,6,16,8,298],
  ['T-302','ボックスティッシュ(5箱)','トイレ・キッチン消耗品','アスクル','パック',5,6,16,8,258],
  ['T-303','ゴミ袋 45L(100枚)','トイレ・キッチン消耗品','アスクル','箱',2,3,8,5,680],
  ['T-304','キッチン洗剤','トイレ・キッチン消耗品','アスクル','本',7,6,16,10,178],
  ['T-305','ラップ 30cm','トイレ・キッチン消耗品','ダイソー・ニトリ','本',11,10,28,15,110],
  ['L-401','フェイスタオル','リネン・寝具','ダイソー・ニトリ','枚',24,30,80,40,199],
  ['L-402','バスタオル','リネン・寝具','ダイソー・ニトリ','枚',16,24,60,30,499],
  ['L-403','枕カバー','リネン・寝具','ダイソー・ニトリ','枚',8,12,30,20,299],
  ['B-501','LED電球 E26','備品・設備','大場金物','個',3,4,10,5,780],
  ['B-502','単3乾電池(20本)','備品・設備','大場金物','パック',1,2,6,3,880],
  ['B-503','ガムテープ','備品・設備','大場金物','巻',5,4,12,6,220],
];

const ITEM_HEADER  = ['コード','品名','分類','発注先','単位','在庫','発注点','適正在庫','発注ロット','単価','最終更新'];
const ORDER_HEADER = ['発注ID','発注日','発注先','コード','品名','単位','発注数','単価','金額','状況','納品日'];
const LOG_HEADER   = ['日時','操作者','権限','種別','内容'];
const ROLE_HEADER  = ['メールアドレス','権限(viewer/staff/manager)','表示名'];

// ============================================================
// ルーティング
// ============================================================
function doGet(e) {
  ensureSetup_();
  const t = HtmlService.createTemplateFromFile('App');
  t.suppliersJson  = JSON.stringify(SUPPLIERS);
  t.categoriesJson = JSON.stringify(CATEGORIES);
  t.permsJson      = JSON.stringify(PERMS);
  return t.evaluate()
    .setTitle('民泊 消耗品 在庫・発注アシスタント')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// 初期化 / シート準備
// ============================================================
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet_(name, header) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#ffffff');
  }
  return sh;
}

// 手動セットアップ（メニューから実行推奨）：シート作成＋サンプル投入＋自分を承認者に
function setup() {
  ensureSetup_();
  const email = Session.getActiveUser().getEmail();
  if (email) setRole_(email, 'manager', email.split('@')[0]);
  SpreadsheetApp.getUi().alert('セットアップ完了', 'シートを作成し、あなた（' + email + '）を承認者に登録しました。\n「デプロイ」→「ウェブアプリ」で公開してください。', SpreadsheetApp.getUi().ButtonSet.OK);
}

// シートが無ければ作成し、品目マスタが空ならサンプル投入。初回利用者を承認者に。
function ensureSetup_() {
  const itemSh = getSheet_(SH_ITEMS, ITEM_HEADER);
  getSheet_(SH_ORDERS, ORDER_HEADER);
  getSheet_(SH_LOG, LOG_HEADER);
  const roleSh = getSheet_(SH_ROLE, ROLE_HEADER);

  if (itemSh.getLastRow() < 2) {
    const rows = SEED_ITEMS.map(r => r.concat([nowStr_()]));
    itemSh.getRange(2, 1, rows.length, ITEM_HEADER.length).setValues(rows);
  }
  // 権限マスタが空なら、最初に開いた人を承認者として登録
  if (roleSh.getLastRow() < 2) {
    const email = Session.getActiveUser().getEmail();
    if (email) roleSh.appendRow([email, 'manager', email.split('@')[0]]);
  }
}

// スプレッドシートを開いたときにメニューを追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('在庫アシスタント')
    .addItem('初期セットアップ / 自分を承認者に', 'setup')
    .addToMenu();
}

// ============================================================
// 権限
// ============================================================
function currentUser_() {
  const email = Session.getActiveUser().getEmail() || '';
  const role  = roleFor_(email);
  return { email: email, role: role, roleName: PERMS[role].name, perms: PERMS[role] };
}

function roleFor_(email) {
  if (!email) return 'viewer';
  const sh = getSheet_(SH_ROLE, ROLE_HEADER);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === email.toLowerCase()) {
      const r = String(rows[i][1]).trim();
      return PERMS[r] ? r : 'viewer';
    }
  }
  return 'viewer';
}

function setRole_(email, role, name) {
  const sh = getSheet_(SH_ROLE, ROLE_HEADER);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === email.toLowerCase()) {
      sh.getRange(i + 1, 1, 1, 3).setValues([[email, role, name || '']]);
      return;
    }
  }
  sh.appendRow([email, role, name || '']);
}

function requirePerm_(key) {
  const me = currentUser_();
  if (!me.perms[key]) {
    throw new Error('権限がありません（現在：' + me.roleName + '）。この操作には上位の権限が必要です。');
  }
  return me;
}

// ============================================================
// 状態の取得（フロントの唯一のデータ源）
// ============================================================
function getBootstrap() {
  return buildState_();
}

function buildState_() {
  return JSON.stringify({
    me:     currentUser_(),
    items:  readItems_(),
    orders: readOrders_(),
    logs:   readLogs_(),
  });
}

function readItems_() {
  const sh = getSheet_(SH_ITEMS, ITEM_HEADER);
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] && !r[1]) continue;
    out.push({ code:s_(r[0]), name:s_(r[1]), category:s_(r[2]), supplier:s_(r[3]), unit:s_(r[4]),
      stock:n_(r[5]), reorder:n_(r[6]), target:n_(r[7]), lot:n_(r[8]), price:n_(r[9]), updated:s_(r[10]) });
  }
  return out;
}

function readOrders_() {
  const sh = getSheet_(SH_ORDERS, ORDER_HEADER);
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    out.push({ id:s_(r[0]), date:s_(r[1]), supplier:s_(r[2]), code:s_(r[3]), name:s_(r[4]),
      unit:s_(r[5]), qty:n_(r[6]), price:n_(r[7]), amount:n_(r[8]), status:s_(r[9]), received:s_(r[10]) });
  }
  return out.reverse(); // 新しい順
}

function readLogs_() {
  const sh = getSheet_(SH_LOG, LOG_HEADER);
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    out.push({ time:s_(r[0]), user:s_(r[1]), role:s_(r[2]), type:s_(r[3]), detail:s_(r[4]) });
  }
  return out.reverse().slice(0, 300);
}

// ============================================================
// 在庫調整（担当者以上）
// ============================================================
function adjustStock(code, delta) {
  const me = requirePerm_('canDraft');
  const sh = getSheet_(SH_ITEMS, ITEM_HEADER);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (s_(rows[i][0]) === code) {
      const before = n_(rows[i][5]);
      const after = Math.max(0, before + Number(delta));
      sh.getRange(i + 1, 6).setValue(after);
      sh.getRange(i + 1, 11).setValue(nowStr_());
      log_(me, '在庫調整', s_(rows[i][1]) + '：' + before + '→' + after + s_(rows[i][4]));
      break;
    }
  }
  return buildState_();
}

// ============================================================
// マスタ編集（承認者のみ）
// ============================================================
function saveItemField(code, field, value) {
  const me = requirePerm_('canEdit');
  const colMap = { code:1, name:2, category:3, supplier:4, unit:5, stock:6, reorder:7, target:8, lot:9, price:10 };
  const numFields = { stock:1, reorder:1, target:1, lot:1, price:1 };
  const col = colMap[field]; if (!col) throw new Error('不正な項目です');
  const sh = getSheet_(SH_ITEMS, ITEM_HEADER);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (s_(rows[i][0]) === code) {
      const before = rows[i][col - 1];
      const val = numFields[field] ? (parseInt(value, 10) || 0) : String(value).trim();
      sh.getRange(i + 1, col).setValue(val);
      sh.getRange(i + 1, 11).setValue(nowStr_());
      log_(me, 'マスタ改定', s_(rows[i][1]) + '：' + field + ' ' + before + '→' + val);
      break;
    }
  }
  return buildState_();
}

function addItem() {
  const me = requirePerm_('canEdit');
  const sh = getSheet_(SH_ITEMS, ITEM_HEADER);
  const code = 'NEW-' + Utilities.formatDate(new Date(), TZ, 'HHmmss');
  sh.appendRow([code, '新規品目', CATEGORIES[0], Object.keys(SUPPLIERS)[0], '個', 0, 0, 0, 1, 0, nowStr_()]);
  log_(me, 'マスタ改定', '品目を追加（' + code + '）');
  return buildState_();
}

function deleteItem(code) {
  const me = requirePerm_('canEdit');
  const sh = getSheet_(SH_ITEMS, ITEM_HEADER);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (s_(rows[i][0]) === code) {
      log_(me, 'マスタ改定', '品目を削除：' + s_(rows[i][1]));
      sh.deleteRow(i + 1);
      break;
    }
  }
  return buildState_();
}

function importItems(rowsJson) {
  const me = requirePerm_('canEdit');
  const rows = JSON.parse(rowsJson);
  const sh = getSheet_(SH_ITEMS, ITEM_HEADER);
  const out = rows.map(c => [s_(c[0]), s_(c[1]), c[2] || CATEGORIES[0], s_(c[3]), c[4] || '個',
    n_(c[5]), n_(c[6]), n_(c[7]), n_(c[8]) || 1, n_(c[9]), nowStr_()]);
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, ITEM_HEADER.length).setValues(out);
  log_(me, 'データ移行', 'CSV取込：' + out.length + '件追加');
  return buildState_();
}

// ============================================================
// 発注確定（承認者のみ・ダブルチェック）
// ============================================================
function suggestQty_(it) {
  const need = Math.max((it.target || it.reorder + it.lot) - it.stock, it.lot);
  const lot = it.lot > 0 ? it.lot : 1;
  return Math.ceil(need / lot) * lot;
}

function confirmOrder(supplier) {
  const me = requirePerm_('canConfirm');
  const items = readItems_().filter(it => it.supplier === supplier && it.stock <= it.reorder);
  if (!items.length) throw new Error('発注対象の品目がありません。');
  const sh = getSheet_(SH_ORDERS, ORDER_HEADER);
  const date = todayStr_();
  let sub = 0;
  const newRows = items.map((it, idx) => {
    const qty = suggestQty_(it);
    const amt = qty * it.price; sub += amt;
    const id = 'O' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + '-' + idx;
    return [id, date, supplier, it.code, it.name, it.unit, qty, it.price, amt, '発注済', ''];
  });
  sh.getRange(sh.getLastRow() + 1, 1, newRows.length, ORDER_HEADER.length).setValues(newRows);
  log_(me, '発注確定', supplier + ' へ ' + items.length + '品目 発注（見込¥' + sub.toLocaleString() + '）');
  return buildState_();
}

// ============================================================
// 入荷確認（担当者以上）→ 在庫へ加算
// ============================================================
function receiveOrder(orderId) {
  const me = requirePerm_('canDraft');
  const osh = getSheet_(SH_ORDERS, ORDER_HEADER);
  const orows = osh.getDataRange().getValues();
  let target = null, orow = -1;
  for (let i = 1; i < orows.length; i++) {
    if (s_(orows[i][0]) === orderId) { target = orows[i]; orow = i + 1; break; }
  }
  if (!target) throw new Error('発注が見つかりません。');
  if (s_(target[9]) === '納品済') throw new Error('すでに納品済です。');

  // 在庫へ加算
  const code = s_(target[3]), qty = n_(target[6]);
  const ish = getSheet_(SH_ITEMS, ITEM_HEADER);
  const irows = ish.getDataRange().getValues();
  let newStock = '?';
  for (let i = 1; i < irows.length; i++) {
    if (s_(irows[i][0]) === code) {
      newStock = n_(irows[i][5]) + qty;
      ish.getRange(i + 1, 6).setValue(newStock);
      ish.getRange(i + 1, 11).setValue(nowStr_());
      break;
    }
  }
  osh.getRange(orow, 10).setValue('納品済');
  osh.getRange(orow, 11).setValue(todayStr_());
  log_(me, '入荷確認', s_(target[4]) + ' ' + qty + s_(target[5]) + ' 入荷 → 在庫' + newStock);
  return buildState_();
}

// ============================================================
// 監査ログ
// ============================================================
function log_(me, type, detail) {
  const sh = getSheet_(SH_LOG, LOG_HEADER);
  sh.appendRow([nowStr_(), me.email || '(不明)', me.roleName, type, detail]);
}

// ============================================================
// ユーティリティ
// ============================================================
function nowStr_()   { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function s_(v) { return v == null ? '' : String(v).trim(); }
function n_(v) { const x = parseInt(v, 10); return isNaN(x) ? 0 : x; }

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
