// ハウジングロビー 評価シート Webアプリ - バックエンド
// Google Apps Script

const SHEET_DATA = '評価データ';
const SHEET_MASTER = 'マスタ';

// ─── ルーティング ─────────────────────────────────────────────
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'input';
  let tmpl;
  if (page === 'dashboard') {
    tmpl = HtmlService.createTemplateFromFile('Dashboard');
    tmpl.pageTitle = '評価ダッシュボード';
  } else {
    tmpl = HtmlService.createTemplateFromFile('Input');
    tmpl.pageTitle = '評価入力フォーム';
  }
  // テンプレートにマスタデータを渡す
  tmpl.employees    = JSON.stringify(EMPLOYEES);
  tmpl.kpis         = JSON.stringify(KPIS);
  tmpl.kpiTemplate  = JSON.stringify(KPI_TEMPLATE);
  return tmpl.evaluate()
    .setTitle('ハウジングロビー 評価シート')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── 社員・KPIマスタ ──────────────────────────────────────────
function getEmployees() {
  return EMPLOYEES;
}

function getKPIsForEmployee(dept, role) {
  const key = dept + '_' + role;
  return KPIS[key] || KPI_TEMPLATE;
}

// ─── データ保存 ───────────────────────────────────────────────
function saveEvaluation(jsonStr) {
  try {
    const data = JSON.parse(jsonStr);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_DATA);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_DATA);
      sheet.appendRow(['ID', 'タイムスタンプ', '社員名', '部署', '役職', '四半期', 'データ(JSON)', '合計点']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    }

    const id = data.name + '_' + data.quarter;
    const totalScore = calcTotal(data.kpis);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        sheet.getRange(i + 1, 1, 1, 8).setValues([[
          id, new Date(), data.name, data.dept, data.role,
          data.quarter, JSON.stringify(data.kpis), totalScore
        ]]);
        return JSON.stringify({ success: true, updated: true, total: totalScore });
      }
    }
    sheet.appendRow([id, new Date(), data.name, data.dept, data.role,
      data.quarter, JSON.stringify(data.kpis), totalScore]);
    return JSON.stringify({ success: true, updated: false, total: totalScore });
  } catch (err) {
    return JSON.stringify({ success: false, error: err.toString() });
  }
}

function calcTotal(kpis) {
  let total = 0;
  kpis.forEach(function(k) {
    const s = parseFloat(k.scale_result);
    const w = parseFloat(k.weight);
    if (!isNaN(s) && !isNaN(w)) total += w * s / 100;
  });
  return Math.round(total * 10) / 10;
}

// ─── データ取得 ───────────────────────────────────────────────
function getEvaluations(quarter) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DATA);
    if (!sheet) return JSON.stringify([]);
    const rows = sheet.getDataRange().getValues();
    const results = [];
    for (let i = 1; i < rows.length; i++) {
      if (!quarter || rows[i][5] === quarter) {
        results.push({
          name: rows[i][2], dept: rows[i][3], role: rows[i][4],
          quarter: rows[i][5], kpis: JSON.parse(rows[i][6] || '[]'),
          total: rows[i][7], timestamp: rows[i][1]
        });
      }
    }
    return JSON.stringify(results);
  } catch (err) {
    return JSON.stringify([]);
  }
}

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DATA);
    if (!sheet) return JSON.stringify({ employees: EMPLOYEES, records: [] });
    const rows = sheet.getDataRange().getValues();
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      records.push({
        name: rows[i][2], dept: rows[i][3], role: rows[i][4],
        quarter: rows[i][5], total: parseFloat(rows[i][7]) || 0,
        kpis: JSON.parse(rows[i][6] || '[]'),
        timestamp: rows[i][1] ? rows[i][1].toString() : ''
      });
    }
    return JSON.stringify({ employees: EMPLOYEES, records: records });
  } catch (err) {
    return JSON.stringify({ employees: EMPLOYEES, records: [], error: err.toString() });
  }
}

function getEvaluationForEdit(name, quarter) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DATA);
    if (!sheet) return JSON.stringify(null);
    const id = name + '_' + quarter;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        return JSON.stringify({
          name: rows[i][2], dept: rows[i][3], role: rows[i][4],
          quarter: rows[i][5], kpis: JSON.parse(rows[i][6] || '[]')
        });
      }
    }
    return JSON.stringify(null);
  } catch (err) {
    return JSON.stringify(null);
  }
}

// ─── 社員マスタデータ ─────────────────────────────────────────
const EMPLOYEES = [
  { name: "細田",     dept: "リーシング", role: "課長代理" },
  { name: "本多",     dept: "リーシング", role: "一般" },
  { name: "西村",     dept: "リーシング", role: "一般" },
  { name: "村田",     dept: "リーシング", role: "一般" },
  { name: "吉田",     dept: "リーシング", role: "一般" },
  { name: "LS新人①", dept: "リーシング", role: "新人" },
  { name: "LS新人②", dept: "リーシング", role: "新人" },
  { name: "奥村",     dept: "カスタマー", role: "課長代理" },
  { name: "CS新人①", dept: "カスタマー", role: "新人" },
  { name: "CS新人②", dept: "カスタマー", role: "新人" },
  { name: "野方",     dept: "経営企画",   role: "課長" },
  { name: "KK新人①", dept: "経営企画",   role: "新人" },
  { name: "KK新人②", dept: "経営企画",   role: "新人" },
  { name: "鹿山",     dept: "民泊課",     role: "一般" },
  { name: "虞",       dept: "民泊課",     role: "一般" },
  { name: "MB新人①", dept: "民泊課",     role: "新人" },
  { name: "MB新人②", dept: "民泊課",     role: "新人" },
  { name: "鬼崎",     dept: "総務管理",   role: "課長" },
  { name: "池田",     dept: "総務管理",   role: "主任" },
  { name: "田尻",     dept: "総務管理",   role: "一般" },
  { name: "宮本",     dept: "総務管理",   role: "一般" },
  { name: "松尾",     dept: "総務管理",   role: "一般" },
  { name: "SZ新人①", dept: "総務管理",   role: "新人" },
  { name: "SZ新人②", dept: "総務管理",   role: "新人" },
  { name: "沖田",     dept: "管理推進",   role: "一般" },
  { name: "KS新人①", dept: "管理推進",   role: "新人" },
  { name: "KS新人②", dept: "管理推進",   role: "新人" }
];

// ─── KPIデータ ────────────────────────────────────────────────
const KPIS = {
  "リーシング_課長代理": [
    { name: "売上目標の達成（全社）",     weight: 35, unit: "%",  bonus: false,
      thresholds: "75未満=0 / 75=10 / 80=20 / 85=30 / 90=40 / 95=50 / 100=60 / 105=70 / 110=80 / 115=90 / 120〜=100" },
    { name: "家賃収入売上UP（課）",       weight: 40, unit: "件", bonus: false,
      thresholds: "12未満=0 / 12=10 / 15=20 / 18=30 / 21=40 / 24=50 / 27=60 / 30=70 / 33=80 / 36=90 / 39〜=100" },
    { name: "仲介手数料売上UP（課）",     weight: 25, unit: "件", bonus: false,
      thresholds: "9未満=0 / 9=10 / 12=20 / 15=30 / 18=40 / 21=50 / 24=60 / 27=70 / 30=80 / 33=90 / 36〜=100" },
    { name: "【加点】入居率アップ（個人）",        weight: null, unit: "件", bonus: true,
      thresholds: "自社仲介による自社管理物件 成約1件→1点加点" },
    { name: "【加点】テナント物件成約（個人）",    weight: null, unit: "件", bonus: true,
      thresholds: "自社管理テナント物件 成約1件→1点加点" }
  ],
  "カスタマー_課長代理": [
    { name: "売上目標の達成（全社）",         weight: 35, unit: "%",  bonus: false,
      thresholds: "75未満=0 / 75=10 / 80=20 / 85=30 / 90=40 / 95=50 / 100=60 / 105=70 / 110=80 / 115=90 / 120〜=100" },
    { name: "お困りごと残件数の削減（課）",    weight: 25, unit: "件", bonus: false,
      thresholds: "91〜=0 / 90=10 / 82=20 / 74=30 / 66=40 / 58=50 / 50=60 / 42=70 / 34=80 / 26=90 / 〜18=100" },
    { name: "顧客満足度向上（課）",           weight: 10, unit: "点", bonus: false,
      thresholds: "4.75未満=0 / 4.75=20 / 4.80=40 / 4.85=60 / +口コミ15=70 / +口コミ20=80 / +口コミ25=90 / +口コミ30=100" },
    { name: "アンケート回収率アップ（課）",   weight: 10, unit: "%",  bonus: false,
      thresholds: "10未満=0 / 10=20 / 15=30 / 20=40 / 25=50 / 30=60 / 35=70 / 40=80 / 45=90 / 50〜=100" },
    { name: "原状回復の対応（課）",           weight: 20, unit: "件", bonus: false,
      thresholds: "7〜=0 / 6=10 / 5=20 / 4=30 / 3=40 / 2=50 / 1=60 / +50万×3件=70 / +50万×5件=80 / +50万×7件=90 / +50万×9件=100" }
  ],
  "経営企画_課長": [
    { name: "売上目標の達成（全社）",            weight: 40, unit: "%",  bonus: false,
      thresholds: "75未満=0 / 75=10 / 80=20 / 85=30 / 90=40 / 95=50 / 100=60 / 105=70 / 110=80 / 115=90 / 120〜=100" },
    { name: "民泊売上目標の達成（課）",           weight: 40, unit: "%",  bonus: false,
      thresholds: "75未満=0 / 75=10 / 80=20 / 85=30 / 90=40 / 95=50 / 100=60 / 105=70 / 110=80 / 115=90 / 120〜=100" },
    { name: "民泊収益アップの新規施策実行（課）", weight: 20, unit: "点", bonus: false,
      thresholds: "承認数×10点 + 実行数×10点（各最大50点）。60点=目標達成" }
  ],
  "総務管理_課長": [
    { name: "売上目標の達成（全社）",  weight: 20, unit: "%",   bonus: false,
      thresholds: "75未満=0 / 75=10 / 80=20 / 85=30 / 90=40 / 95=50 / 100=60 / 105=70 / 110=80 / 115=90 / 120〜=100" },
    { name: "販管費削減目標の達成",    weight: 40, unit: "万円", bonus: false,
      thresholds: "未達=0 / 月10万改善=60 / 月20万改善=70 / 月30万改善=80 / 月40万改善=90 / 月50万以上=100" },
    { name: "人事採用目標",            weight: 10, unit: "人",  bonus: false,
      thresholds: "期ごとに目標設定" },
    { name: "その他課題（課長設定）",  weight: 30, unit: "—",   bonus: false,
      thresholds: "期ごとに設定" }
  ],
  "総務管理_主任": [
    { name: "売上目標の達成（全社）", weight: 30, unit: "%",  bonus: false,
      thresholds: "75未満=0 / 75=10 / 80=20 / 85=30 / 90=40 / 95=50 / 100=60 / 105=70 / 110=80 / 115=90 / 120〜=100" },
    { name: "加点・業務評価項目",     weight: 70, unit: "点", bonus: false,
      thresholds: "60点を持ち点。業績影響+3点 / 時短+2点 / 軽微+1点 / 金銭ミス-3点 / コンプラ違反-10点" },
    { name: "【減点】減点項目",        weight: null, unit: "点", bonus: true,
      thresholds: "金銭ミス/期限超過: −3点 / コンプラ違反/虚偽報告: −10点" }
  ],
  "管理推進_一般": [
    { name: "売上目標の達成（全社）",           weight: 20, unit: "%",  bonus: false,
      thresholds: "75未満=0 / 75=10 / 80=20 / 85=30 / 90=40 / 95=50 / 100=60 / 105=70 / 110=80 / 115=90 / 120〜=100" },
    { name: "（サブリース）オーナー交渉（個人）", weight: 50, unit: "pt", bonus: false,
      thresholds: "90未満=0 / 90=10 / 120=20 / 150=30 / 180=40 / 210=50 / 240=60 / 270=70 / 300=80 / 330=90 / 360〜=100" },
    { name: "（一般管理）オーナー提案（個人）",  weight: 30, unit: "—",  bonus: false,
      thresholds: "出来ない=0 / 担当者が出来る=60 / 評価売上300万=70 / 600万=80 / 900万=90 / 1200万〜=100" },
    { name: "【加点】新規管理物件取得（個人）",  weight: null, unit: "棟", bonus: true,
      thresholds: "1棟取得で+5点（会社紹介案件も同様）" },
    { name: "【加点】売買仲介契約締結（個人）",  weight: null, unit: "件", bonus: true,
      thresholds: "専任媒介契約締結で+5点" }
  ]
};

const KPI_TEMPLATE = [
  { name: "目標①",        weight: null, unit: "", bonus: false, thresholds: "期ごとに設定" },
  { name: "目標②",        weight: null, unit: "", bonus: false, thresholds: "期ごとに設定" },
  { name: "目標③",        weight: null, unit: "", bonus: false, thresholds: "期ごとに設定" },
  { name: "【加点】加点項目", weight: null, unit: "", bonus: true,  thresholds: "期ごとに設定" }
];
