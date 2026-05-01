#!/usr/bin/env python3
"""評価シート作成スクリプト - ハウジングロビー様"""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import base64, io

# ─── カラー・スタイル ──────────────────────────────────────
def fill(hex_c):
    return PatternFill(start_color=hex_c, end_color=hex_c, fill_type="solid")

def font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size, name="Meiryo UI")

def border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def align(h="center", wrap=True):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)

F = {
    "dblue":  fill("1F4E79"),
    "mblue":  fill("2E75B6"),
    "lblue":  fill("BDD7EE"),
    "yellow": fill("FFFF99"),
    "green":  fill("E2EFDA"),
    "orange": fill("FCE4D6"),
    "gray":   fill("F2F2F2"),
    "lgray":  fill("FAFAFA"),
    "q1":     fill("4472C4"),
    "q2":     fill("C55A11"),
    "q3":     fill("538135"),
    "q4":     fill("BF8F00"),
    "total":  fill("D6DCE4"),
    "semi":   fill("FFD966"),
    "white":  fill("FFFFFF"),
}

BORDER = border()

# ─── 列番号定数 ────────────────────────────────────────────
C_ROLE    =  1  # A 評価項目
C_PERIOD  =  2  # B 期限
C_TARGET  =  3  # C 目標状態
C_RULE    =  4  # D ルール
C_WEIGHT  =  5  # E 重み(%)
C_T0      =  6  # F 基準/0
C_T10     =  7  # G 基準/10
C_T20     =  8  # H 基準/20
C_T30     =  9  # I 基準/30
C_T40     = 10  # J 基準/40
C_T50     = 11  # K 基準/50
C_T60     = 12  # L 基準/60
C_T70     = 13  # M 基準/70
C_T80     = 14  # N 基準/80
C_T90     = 15  # O 基準/90
C_T100    = 16  # P 基準/100
C_UNIT    = 17  # Q 単位
C_RESULT  = 18  # R 結果(入力)
C_SCALE   = 19  # S 尺度結果(入力)
C_WSCORE  = 20  # T 重み×尺度

COL_HEADERS = [
    "評価項目", "期限", "目標状態（どのような状態）", "ルール",
    "重み", "0", "10", "20", "30", "40", "50", "60", "70", "80", "90", "100",
    "単位", "結果\n(入力)", "尺度結果\n(0-100)", "重み×尺度"
]

# ─── 四半期定義 ────────────────────────────────────────────
QUARTERS = [
    {"label": "Q1", "period": "11月〜1月", "fill_key": "q1"},
    {"label": "Q2", "period": "2月〜4月",  "fill_key": "q2"},
    {"label": "Q3", "period": "5月〜7月",  "fill_key": "q3"},
    {"label": "Q4", "period": "8月〜10月", "fill_key": "q4"},
]

MAX_KPI = 8  # 1四半期あたり最大KPI行数

# ─── KPIデータ ──────────────────────────────────────────────
# thresholds: [0点閾値, 10点閾値, ..., 100点閾値] の11要素
# reverse=True → 数値が小さいほど高得点
KPIS = {}

KPIS[("リーシング", "課長代理")] = [
    {
        "name": "売上目標の達成（全社）",
        "period": "期末日",
        "target": "全社売上予算比100%達成",
        "rule": "小数点第1位切り捨て",
        "weight": 35, "unit": "%", "reverse": False,
        "thresholds": ["75未満", 75, 80, 85, 90, 95, 100, 105, 110, 115, "120〜"],
        "t_num":      [None,     75, 80, 85, 90, 95, 100, 105, 110, 115, None],
        "bonus": False,
    },
    {
        "name": "家賃収入売上UP（課）",
        "period": "期末日",
        "target": "サブリース物件の契約件数達成",
        "rule": "契約日ベース",
        "weight": 40, "unit": "件", "reverse": False,
        "thresholds": ["12未満", 12, 15, 18, 21, 24, 27, 30, 33, 36, "39〜"],
        "t_num":      [None,     12, 15, 18, 21, 24, 27, 30, 33, 36, None],
        "bonus": False,
    },
    {
        "name": "仲介手数料売上UP（課）",
        "period": "期末日",
        "target": "自社成約件数達成",
        "rule": "契約日ベース",
        "weight": 25, "unit": "件", "reverse": False,
        "thresholds": ["9未満", 9, 12, 15, 18, 21, 24, 27, 30, 33, "36〜"],
        "t_num":      [None,    9, 12, 15, 18, 21, 24, 27, 30, 33, None],
        "bonus": False,
    },
    {
        "name": "【加点】入居率アップ（個人）",
        "period": "期末日",
        "target": "—",
        "rule": "自社仲介による自社管理物件 成約1件→1点加点",
        "weight": None, "unit": "件", "reverse": False,
        "thresholds": None, "t_num": None, "bonus": True,
    },
    {
        "name": "【加点】テナント物件成約（個人）",
        "period": "期末日",
        "target": "—",
        "rule": "自社管理テナント物件 成約1件→1点加点",
        "weight": None, "unit": "件", "reverse": False,
        "thresholds": None, "t_num": None, "bonus": True,
    },
]

KPIS[("カスタマー", "課長代理")] = [
    {
        "name": "売上目標の達成（全社）",
        "period": "期末日", "target": "全社売上予算比100%達成",
        "rule": "小数点第1位切り捨て",
        "weight": 35, "unit": "%", "reverse": False,
        "thresholds": ["75未満", 75, 80, 85, 90, 95, 100, 105, 110, 115, "120〜"],
        "t_num":      [None,     75, 80, 85, 90, 95, 100, 105, 110, 115, None],
        "bonus": False,
    },
    {
        "name": "お困りごと残件数の削減（課）",
        "period": "期末日", "target": "お困りごと残件数50件以下",
        "rule": "期末日18:00時点の残件数",
        "weight": 25, "unit": "件", "reverse": True,
        "thresholds": ["91〜", 90, 82, 74, 66, 58, 50, 42, 34, 26, "〜18"],
        "t_num":      [None,  90, 82, 74, 66, 58, 50, 42, 34, 26,  None],
        "bonus": False,
    },
    {
        "name": "顧客満足度向上（課）",
        "period": "期末日", "target": "アンケート評価4.85以上",
        "rule": "期末時点の総合評価平均点",
        "weight": 10, "unit": "点/件", "reverse": False,
        "thresholds": ["4.75未満", "—", "4.75", "—", "4.80", "—", "4.85",
                       "4.85+口コミ15", "4.85+口コミ20", "4.85+口コミ25", "4.85+口コミ30"],
        "t_num": None,
        "bonus": False,
    },
    {
        "name": "アンケート回収率アップ（課）",
        "period": "期末日", "target": "アンケート回収率30%以上",
        "rule": "完了件数のうち回収した割合",
        "weight": 10, "unit": "%", "reverse": False,
        "thresholds": ["10未満", "—", 10, 15, 20, 25, 30, 35, 40, 45, "50〜"],
        "t_num":      [None,    None, 10, 15, 20, 25, 30, 35, 40, 45, None],
        "bonus": False,
    },
    {
        "name": "原状回復の対応（課）",
        "period": "期末日", "target": "退去14日以内の提案、超過1日以下",
        "rule": "提案日で評価。再見積・交渉は除外",
        "weight": 20, "unit": "件", "reverse": True,
        "thresholds": ["7〜", 6, 5, 4, 3, 2, 1,
                       "60点+50万×3件", "60点+50万×5件", "60点+50万×7件", "60点+50万×9件"],
        "t_num": [None, 6, 5, 4, 3, 2, 1, None, None, None, None],
        "bonus": False,
    },
]

KPIS[("経営企画", "課長")] = [
    {
        "name": "売上目標の達成（全社）",
        "period": "期末日", "target": "全社売上予算比100%達成",
        "rule": "小数点第1位切り捨て",
        "weight": 40, "unit": "%", "reverse": False,
        "thresholds": ["75未満", 75, 80, 85, 90, 95, 100, 105, 110, 115, "120〜"],
        "t_num":      [None,     75, 80, 85, 90, 95, 100, 105, 110, 115, None],
        "bonus": False,
    },
    {
        "name": "民泊売上目標の達成（課）",
        "period": "期末日", "target": "民泊売上予算比100%達成",
        "rule": "小数点第1位切り捨て",
        "weight": 40, "unit": "%", "reverse": False,
        "thresholds": ["75未満", 75, 80, 85, 90, 95, 100, 105, 110, 115, "120〜"],
        "t_num":      [None,     75, 80, 85, 90, 95, 100, 105, 110, 115, None],
        "bonus": False,
    },
    {
        "name": "民泊収益アップの新規施策実行（課）",
        "period": "期末日", "target": "新規施策60点達成",
        "rule": "承認数×10点 + 実行数×10点（各最大50点）",
        "weight": 20, "unit": "点", "reverse": False,
        "thresholds": [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
        "t_num":      [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
        "bonus": False,
    },
]

KPIS[("総務管理", "課長")] = [
    {
        "name": "売上目標の達成（全社）",
        "period": "期末日", "target": "全社売上予算比100%達成",
        "rule": "小数点第1位切り捨て",
        "weight": 20, "unit": "%", "reverse": False,
        "thresholds": ["75未満", 75, 80, 85, 90, 95, 100, 105, 110, 115, "120〜"],
        "t_num":      [None,     75, 80, 85, 90, 95, 100, 105, 110, 115, None],
        "bonus": False,
    },
    {
        "name": "販管費削減目標の達成",
        "period": "期末日", "target": "月平均10万円の収支改善",
        "rule": "毎週収支改善MTG実施。外部支出削減・時短・助成金等",
        "weight": 40, "unit": "万円", "reverse": False,
        "thresholds": ["未達", "—", "—", "—", "—", "—", "月10万改善",
                       "月20万改善", "月30万改善", "月40万改善", "月50万以上"],
        "t_num": None,
        "bonus": False,
    },
    {
        "name": "人事採用目標",
        "period": "期末日", "target": "新卒採用目標の達成",
        "rule": "期ごとに目標設定",
        "weight": 10, "unit": "人", "reverse": False,
        "thresholds": None, "t_num": None, "bonus": False,
    },
    {
        "name": "その他課題（課長設定）",
        "period": "期末日", "target": "期ごとに設定",
        "rule": "期ごとに設定",
        "weight": 30, "unit": "—", "reverse": False,
        "thresholds": None, "t_num": None, "bonus": False,
    },
]

KPIS[("総務管理", "主任")] = [
    {
        "name": "売上目標の達成（全社）",
        "period": "期末日", "target": "全社売上予算比100%達成",
        "rule": "小数点第1位切り捨て",
        "weight": 30, "unit": "%", "reverse": False,
        "thresholds": ["75未満", 75, 80, 85, 90, 95, 100, 105, 110, 115, "120〜"],
        "t_num":      [None,     75, 80, 85, 90, 95, 100, 105, 110, 115, None],
        "bonus": False,
    },
    {
        "name": "加点・業務評価項目",
        "period": "期末日",
        "target": "社内環境改善提案・業務改善・引継ぎ掌握・上席評価",
        "rule": "60点を持ち点とし、上席判断で加点・減点\n業績影響:+3点 / 時短:+2点 / 軽微:+1点",
        "weight": 70, "unit": "点", "reverse": False,
        "thresholds": None, "t_num": None, "bonus": False,
    },
    {
        "name": "【減点】減点項目",
        "period": "期末日", "target": "—",
        "rule": "金銭ミス/期限超過: −3点\nコンプラ違反/虚偽報告: −10点",
        "weight": None, "unit": "点", "reverse": False,
        "thresholds": None, "t_num": None, "bonus": True,
    },
]

KPIS[("管理推進", "一般")] = [
    {
        "name": "売上目標の達成（全社）",
        "period": "期末日", "target": "全社売上予算比100%達成",
        "rule": "小数点第1位切り捨て",
        "weight": 20, "unit": "%", "reverse": False,
        "thresholds": ["75未満", 75, 80, 85, 90, 95, 100, 105, 110, 115, "120〜"],
        "t_num":      [None,     75, 80, 85, 90, 95, 100, 105, 110, 115, None],
        "bonus": False,
    },
    {
        "name": "（サブリース）オーナー交渉（個人）",
        "period": "期末日", "target": "オーナー交渉240ポイント達成",
        "rule": "契約締結時点でカウント",
        "weight": 50, "unit": "pt", "reverse": False,
        "thresholds": ["90未満", 90, 120, 150, 180, 210, 240, 270, 300, 330, "360〜"],
        "t_num":      [None,     90, 120, 150, 180, 210, 240, 270, 300, 330, None],
        "bonus": False,
    },
    {
        "name": "（一般管理）オーナー提案（個人）",
        "period": "期末日", "target": "担当者が提案・契約書作成を出来る状態",
        "rule": "尺度60=担当者が出来る。70以上=評価売上で判定",
        "weight": 30, "unit": "—", "reverse": False,
        "thresholds": ["出来ない", "—", "—", "—", "—", "—", "担当者が出来る",
                       "評価売上300万", "評価売上600万", "評価売上900万", "評価売上1200万〜"],
        "t_num": None,
        "bonus": False,
    },
    {
        "name": "【加点】新規管理物件取得（個人）",
        "period": "期末日", "target": "—",
        "rule": "1棟取得で+5点（会社紹介案件も同様）",
        "weight": None, "unit": "棟", "reverse": False,
        "thresholds": None, "t_num": None, "bonus": True,
    },
    {
        "name": "【加点】売買仲介契約締結（個人）",
        "period": "期末日", "target": "—",
        "rule": "専任媒介契約締結で+5点",
        "weight": None, "unit": "件", "reverse": False,
        "thresholds": None, "t_num": None, "bonus": True,
    },
]

# 未定義ロール用テンプレート（一般・新人）
KPI_TEMPLATE = [
    {"name": "目標①",    "period": "", "target": "", "rule": "", "weight": None,
     "unit": "", "reverse": False, "thresholds": None, "t_num": None, "bonus": False},
    {"name": "目標②",    "period": "", "target": "", "rule": "", "weight": None,
     "unit": "", "reverse": False, "thresholds": None, "t_num": None, "bonus": False},
    {"name": "目標③",    "period": "", "target": "", "rule": "", "weight": None,
     "unit": "", "reverse": False, "thresholds": None, "t_num": None, "bonus": False},
    {"name": "【加点】加点項目", "period": "", "target": "", "rule": "", "weight": None,
     "unit": "", "reverse": False, "thresholds": None, "t_num": None, "bonus": True},
]

# ─── 社員リスト ────────────────────────────────────────────
EMPLOYEES = [
    # リーシング
    {"name": "細田",         "dept": "リーシング", "role": "課長代理"},
    {"name": "本多",         "dept": "リーシング", "role": "一般"},
    {"name": "西村",         "dept": "リーシング", "role": "一般"},
    {"name": "村田",         "dept": "リーシング", "role": "一般"},
    {"name": "吉田",         "dept": "リーシング", "role": "一般"},
    {"name": "LS新人①",     "dept": "リーシング", "role": "新人"},
    {"name": "LS新人②",     "dept": "リーシング", "role": "新人"},
    # カスタマー
    {"name": "奥村",         "dept": "カスタマー", "role": "課長代理"},
    {"name": "CS新人①",     "dept": "カスタマー", "role": "新人"},
    {"name": "CS新人②",     "dept": "カスタマー", "role": "新人"},
    # 経営企画
    {"name": "野方",         "dept": "経営企画",   "role": "課長"},
    {"name": "KK新人①",     "dept": "経営企画",   "role": "新人"},
    {"name": "KK新人②",     "dept": "経営企画",   "role": "新人"},
    # 民泊課
    {"name": "鹿山",         "dept": "民泊課",     "role": "一般"},
    {"name": "虞",           "dept": "民泊課",     "role": "一般"},
    {"name": "MB新人①",     "dept": "民泊課",     "role": "新人"},
    {"name": "MB新人②",     "dept": "民泊課",     "role": "新人"},
    # 総務管理
    {"name": "鬼崎",         "dept": "総務管理",   "role": "課長"},
    {"name": "池田",         "dept": "総務管理",   "role": "主任"},
    {"name": "田尻",         "dept": "総務管理",   "role": "一般"},
    {"name": "宮本",         "dept": "総務管理",   "role": "一般"},
    {"name": "松尾",         "dept": "総務管理",   "role": "一般"},
    {"name": "SZ新人①",     "dept": "総務管理",   "role": "新人"},
    {"name": "SZ新人②",     "dept": "総務管理",   "role": "新人"},
    # 管理推進
    {"name": "沖田",         "dept": "管理推進",   "role": "一般"},
    {"name": "KS新人①",     "dept": "管理推進",   "role": "新人"},
    {"name": "KS新人②",     "dept": "管理推進",   "role": "新人"},
]

# ─── 行位置計算 ─────────────────────────────────────────────
# 固定レイアウト:
#   Row 1     : 社員情報ヘッダー
#   Row 2     : blank
#   Q1: rows  3-14  (header:3, col-hdr:4, KPI:5-12, total:13, blank:14)
#   Q2: rows 15-26
#   Q3: rows 27-38
#   Q4: rows 39-50
#   Row 51    : blank
#   Row 52    : 上半期合計
#   Row 53    : 下半期合計

def q_rows(q_idx):
    """q_idx=0(Q1)〜3(Q4) → (header_row, col_hdr_row, kpi_start, kpi_end, total_row, blank_row)"""
    base = 3 + q_idx * 12
    return (base, base+1, base+2, base+2+MAX_KPI-1, base+2+MAX_KPI, base+2+MAX_KPI+1)

TOTAL_ROW = {q: q_rows(q)[4] for q in range(4)}
# Q1総合: row 13, Q2: row 25, Q3: row 37, Q4: row 49

SEMI_H1_ROW = 52
SEMI_H2_ROW = 53
INFO_ROW = 1

# ─── ヘルパー ──────────────────────────────────────────────
def set_cell(ws, row, col, value, fill_key=None, bold=False, fcolor="000000",
             h_align="center", wrap=True, fsize=10, number_format=None):
    cell = ws.cell(row=row, column=col, value=value)
    if fill_key:
        cell.fill = F[fill_key]
    cell.font = font(bold=bold, color=fcolor, size=fsize)
    cell.alignment = align(h_align, wrap)
    cell.border = BORDER
    if number_format:
        cell.number_format = number_format
    return cell

def merge_row(ws, row, col_start, col_end, value, fill_key, bold=True, fcolor="FFFFFF", fsize=11):
    ws.merge_cells(start_row=row, start_column=col_start,
                   end_row=row, end_column=col_end)
    cell = ws.cell(row=row, column=col_start, value=value)
    cell.fill = F[fill_key]
    cell.font = font(bold=bold, color=fcolor, size=fsize)
    cell.alignment = align("center", True)
    cell.border = BORDER
    return cell

def add_dropdown(ws, col, row_start, row_end, formula):
    dv = DataValidation(
        type="list", formula1=formula, allow_blank=True,
        showErrorMessage=True, errorTitle="入力エラー",
        error="リストから選択してください"
    )
    ws.add_data_validation(dv)
    for r in range(row_start, row_end + 1):
        dv.add(ws.cell(row=r, column=col))

# ─── 個人シート作成 ────────────────────────────────────────
def write_individual_sheet(wb, emp):
    name  = emp["name"]
    dept  = emp["dept"]
    role  = emp["role"]
    sname = name[:31]
    ws    = wb.create_sheet(title=sname)

    # 列幅設定
    col_widths = {
        1: 22, 2: 10, 3: 22, 4: 22, 5: 7,
        6: 8, 7: 8, 8: 8, 9: 8, 10: 8, 11: 8,
        12: 8, 13: 8, 14: 8, 15: 8, 16: 8,
        17: 7, 18: 10, 19: 10, 20: 12,
    }
    for c, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    # ─ Row 1: 社員情報ヘッダー ─
    ws.row_dimensions[1].height = 28
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)
    c = ws.cell(row=1, column=1,
                value=f"【個人評価シート】  {dept}  {role}　{name}")
    c.fill = F["dblue"]
    c.font = font(bold=True, color="FFFFFF", size=13)
    c.alignment = align("left", False)
    c.border = BORDER

    # ─ 凡例（Row 2）─
    ws.row_dimensions[2].height = 18
    legend = [
        (C_RESULT, "黄色=入力欄", "yellow"),
        (C_SCALE,  "黄色=入力欄", "yellow"),
        (C_WSCORE, "緑色=自動計算", "green"),
    ]
    for col, txt, fk in legend:
        set_cell(ws, 2, col, txt, fk, fsize=8)

    # KPI取得
    kpis = KPIS.get((dept, role), KPI_TEMPLATE)

    # ─ 各四半期ブロック ─
    for q_idx, q in enumerate(QUARTERS):
        hdr_row, col_hdr_row, kpi_start, kpi_end, total_row, blank_row = q_rows(q_idx)

        ws.row_dimensions[hdr_row].height = 22
        # 四半期ヘッダー
        merge_row(ws, hdr_row, 1, 20,
                  f"【 {q['label']} 】  {q['period']}", q["fill_key"])

        ws.row_dimensions[col_hdr_row].height = 36
        # 列ヘッダー
        for ci, h in enumerate(COL_HEADERS, start=1):
            hc = set_cell(ws, col_hdr_row, ci, h, "gray", bold=True, fsize=8)

        # KPI行
        scale_rows = []
        for ki, kpi in enumerate(kpis[:MAX_KPI]):
            r = kpi_start + ki
            ws.row_dimensions[r].height = 40
            is_bonus = kpi.get("bonus", False)
            row_fill = "orange" if is_bonus else "lgray"

            set_cell(ws, r, C_ROLE,   kpi["name"],   row_fill, fsize=9, h_align="left")
            set_cell(ws, r, C_PERIOD, kpi["period"],  row_fill, fsize=9)
            set_cell(ws, r, C_TARGET, kpi["target"],  row_fill, fsize=8, h_align="left")
            set_cell(ws, r, C_RULE,   kpi["rule"],    row_fill, fsize=8, h_align="left")

            # 重み
            w = kpi["weight"]
            if w is not None:
                set_cell(ws, r, C_WEIGHT, w/100, row_fill, fsize=9,
                         number_format="0%")
            else:
                set_cell(ws, r, C_WEIGHT, "加点", row_fill, fsize=8, fcolor="8B0000")

            # 基準点（F-P）
            ths = kpi.get("thresholds")
            for ci_offset, tv in enumerate(ths if ths else [""]*11):
                set_cell(ws, r, C_T0 + ci_offset, tv, row_fill, fsize=8)

            # 単位
            set_cell(ws, r, C_UNIT, kpi["unit"], row_fill, fsize=9)

            # 結果入力（黄色）
            set_cell(ws, r, C_RESULT, None, "yellow", fsize=10)

            # 尺度結果（黄色）
            set_cell(ws, r, C_SCALE, None, "yellow", fsize=10)

            # 重み×尺度（緑・自動）
            if not is_bonus and w is not None:
                wr_col = get_column_letter(C_WEIGHT)
                sc_col = get_column_letter(C_SCALE)
                formula = (
                    f'=IF(OR({sc_col}{r}="",'
                    f'NOT(ISNUMBER({sc_col}{r}))),"—",'
                    f'{wr_col}{r}*{sc_col}{r}/100)'
                )
                set_cell(ws, r, C_WSCORE, formula, "green", fsize=10,
                         number_format="0.0")
                scale_rows.append(r)
            else:
                set_cell(ws, r, C_WSCORE, "—", "green", fsize=9, fcolor="888888")

        # 空きKPI行を埋める
        for ki in range(len(kpis), MAX_KPI):
            r = kpi_start + ki
            ws.row_dimensions[r].height = 28
            for ci in range(1, 21):
                set_cell(ws, r, ci, None, "white")

        # 合計行
        ws.row_dimensions[total_row].height = 24
        merge_row(ws, total_row, 1, C_WSCORE-1, f"{q['label']} 合計点", "total",
                  bold=True, fcolor="1F4E79", fsize=10)
        if scale_rows:
            first_r = kpi_start
            last_r  = kpi_start + MAX_KPI - 1
            wt_col  = get_column_letter(C_WSCORE)
            total_f = f'=SUMIF({wt_col}{first_r}:{wt_col}{last_r},"<>—")'
            tc = set_cell(ws, total_row, C_WSCORE, total_f, "total",
                          bold=True, fsize=11, number_format="0.0")
        else:
            set_cell(ws, total_row, C_WSCORE, "—", "total", fsize=10)

        # 空白区切り行
        ws.row_dimensions[blank_row].height = 6
        for ci in range(1, 21):
            c = ws.cell(row=blank_row, column=ci)
            c.fill = F["white"]

        # 尺度結果にドロップダウン追加
        add_dropdown(ws, C_SCALE, kpi_start, kpi_start + MAX_KPI - 1,
                     '"0,10,20,30,40,50,60,70,80,90,100"')

    # ─ 半期合計行 ─
    wt_col = get_column_letter(C_WSCORE)
    ws.row_dimensions[51].height = 8
    ws.row_dimensions[SEMI_H1_ROW].height = 22
    ws.row_dimensions[SEMI_H2_ROW].height = 22

    merge_row(ws, SEMI_H1_ROW, 1, C_WSCORE-1, "上半期合計（Q1＋Q2）", "semi",
              bold=True, fcolor="1F4E79", fsize=10)
    h1_f = f'={wt_col}{TOTAL_ROW[0]}+{wt_col}{TOTAL_ROW[1]}'
    set_cell(ws, SEMI_H1_ROW, C_WSCORE, h1_f, "semi", bold=True,
             fsize=11, number_format="0.0")

    merge_row(ws, SEMI_H2_ROW, 1, C_WSCORE-1, "下半期合計（Q3＋Q4）", "semi",
              bold=True, fcolor="1F4E79", fsize=10)
    h2_f = f'={wt_col}{TOTAL_ROW[2]}+{wt_col}{TOTAL_ROW[3]}'
    set_cell(ws, SEMI_H2_ROW, C_WSCORE, h2_f, "semi", bold=True,
             fsize=11, number_format="0.0")

    # 先頭行固定
    ws.freeze_panes = "A3"

    return ws

# ─── 分析シート作成 ────────────────────────────────────────
def write_analysis_sheet(wb, label, q_indices, sheet_name):
    ws = wb.create_sheet(title=sheet_name)

    # 列幅
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 12
    for ci in range(4, 4 + len(q_indices) + 3):
        ws.column_dimensions[get_column_letter(ci)].width = 12

    # タイトル
    end_col = 3 + len(q_indices) + 2
    ws.row_dimensions[1].height = 26
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    c = ws.cell(row=1, column=1, value=f"【分析表】{label}")
    c.fill = F["dblue"]
    c.font = font(bold=True, color="FFFFFF", size=13)
    c.alignment = align("left", False)
    c.border = BORDER

    # ヘッダー行
    ws.row_dimensions[2].height = 30
    headers = ["氏名", "部署", "役職"]
    q_labels_row = []
    for qi in q_indices:
        q = QUARTERS[qi]
        headers.append(f"{q['label']}\n{q['period']}")
        q_labels_row.append(qi)
    if len(q_indices) == 2:
        if q_indices == [0, 1]:
            headers.append("上半期合計")
        else:
            headers.append("下半期合計")
    headers.append("順位")

    q_fill_keys = [QUARTERS[qi]["fill_key"] for qi in q_indices]
    for ci, h in enumerate(headers, start=1):
        fk = q_fill_keys[ci - 4] if 4 <= ci <= 3 + len(q_indices) else "mblue"
        hc = ws.cell(row=2, column=ci, value=h)
        hc.fill = F[fk]
        hc.font = font(bold=True, color="FFFFFF", size=9)
        hc.alignment = align("center", True)
        hc.border = BORDER

    # データ行
    for ri, emp in enumerate(EMPLOYEES, start=3):
        ws.row_dimensions[ri].height = 22
        sname = emp["name"][:31]
        set_cell(ws, ri, 1, emp["name"],  "lgray", fsize=10)
        set_cell(ws, ri, 2, emp["dept"],  "lgray", fsize=9)
        set_cell(ws, ri, 3, emp["role"],  "lgray", fsize=9)

        score_cols = []
        for ci_off, qi in enumerate(q_indices):
            total_row_num = TOTAL_ROW[qi]
            wt_col = get_column_letter(C_WSCORE)
            ref = f"'{sname}'!{wt_col}{total_row_num}"
            col = 4 + ci_off
            set_cell(ws, ri, col, f"={ref}", "green", fsize=10, number_format="0.0")
            score_cols.append(get_column_letter(col))

        # 半期合計 or 単Q
        sum_col = 4 + len(q_indices)
        if len(q_indices) == 2:
            sum_f = "=" + "+".join(f"{sc}{ri}" for sc in score_cols)
            set_cell(ws, ri, sum_col, sum_f, "semi", bold=True,
                     fsize=10, number_format="0.0")
        else:
            set_cell(ws, ri, sum_col, "—", "lgray", fsize=9)

        # 順位（半期合計基準）
        rank_col  = sum_col + 1
        total_emp = len(EMPLOYEES)
        sum_letter = get_column_letter(sum_col)
        rank_f = (
            f'=IFERROR(RANK({sum_letter}{ri},'
            f'{sum_letter}3:{sum_letter}{2+total_emp},0),"")'
        )
        set_cell(ws, ri, rank_col, rank_f, "lgray", fsize=10)

    ws.freeze_panes = "A3"
    return ws

# ─── マスタシート ──────────────────────────────────────────
def write_master_sheet(wb):
    ws = wb.active
    ws.title = "マスタ"

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 16

    ws.row_dimensions[1].height = 28
    ws.merge_cells("A1:D1")
    c = ws.cell(row=1, column=1, value="【マスタ】社員一覧・評価制度概要")
    c.fill = F["dblue"]
    c.font = font(bold=True, color="FFFFFF", size=13)
    c.alignment = align("left", False)
    c.border = BORDER

    # 評価制度説明
    info = [
        ("", ""),
        ("■ 評価期間", ""),
        ("Q1", "11月〜1月"),
        ("Q2", "2月〜4月"),
        ("Q3", "5月〜7月"),
        ("Q4", "8月〜10月"),
        ("上半期", "Q1 ＋ Q2"),
        ("下半期", "Q3 ＋ Q4"),
        ("", ""),
        ("■ 全社売上連動重み", ""),
        ("部長",   "50%"),
        ("課長",   "40%"),
        ("課長代理", "35%"),
        ("主任",   "30%"),
        ("一般",   "20%"),
        ("", ""),
        ("■ 尺度スコアの入力方法", ""),
        ("・スコア列（S列）をクリック", "→ドロップダウンから0〜100を選択"),
        ("・基準点欄を参照して判定",   "数値が明確な場合は結果欄(R列)に入力後に判定"),
        ("", ""),
        ("■ 自動計算の仕組み", ""),
        ("重み×尺度結果（T列）", "= 重み(E列) × 尺度結果(S列) ÷ 100"),
        ("Q合計点（T列 合計行）", "= 各KPIのT列合計"),
        ("上半期合計", "= Q1合計 ＋ Q2合計"),
        ("下半期合計", "= Q3合計 ＋ Q4合計"),
    ]

    for ri, (k, v) in enumerate(info, start=2):
        ws.row_dimensions[ri].height = 20
        ck = ws.cell(row=ri, column=1, value=k)
        cv = ws.cell(row=ri, column=3, value=v)
        for cell in [ck, cv]:
            cell.border = BORDER
            cell.alignment = align("left", False)
            cell.font = font(size=10)
            if k.startswith("■"):
                cell.fill = F["lblue"]
                cell.font = font(bold=True, size=10, color="1F4E79")
            elif k in ("Q1","Q2","Q3","Q4"):
                qmap = {"Q1":"q1","Q2":"q2","Q3":"q3","Q4":"q4"}
                cell.fill = F[qmap[k]]
                cell.font = font(bold=True, color="FFFFFF", size=10)
            else:
                cell.fill = F["white"]

    # 社員一覧テーブル
    start_r = len(info) + 4
    ws.merge_cells(start_row=start_r, start_column=1,
                   end_row=start_r, end_column=4)
    ch = ws.cell(row=start_r, column=1, value="■ 社員一覧")
    ch.fill = F["mblue"]
    ch.font = font(bold=True, color="FFFFFF", size=11)
    ch.alignment = align("left", False)
    ch.border = BORDER

    hdr_r = start_r + 1
    for ci, h in enumerate(["No.", "氏名", "部署", "役職"], start=1):
        hc = ws.cell(row=hdr_r, column=ci, value=h)
        hc.fill = F["lblue"]
        hc.font = font(bold=True, size=10, color="1F4E79")
        hc.alignment = align("center", False)
        hc.border = BORDER

    for ei, emp in enumerate(EMPLOYEES, start=1):
        r = hdr_r + ei
        ws.row_dimensions[r].height = 20
        for ci, v in enumerate([ei, emp["name"], emp["dept"], emp["role"]], start=1):
            ec = ws.cell(row=r, column=ci, value=v)
            ec.fill = F["lgray"] if ei % 2 else F["white"]
            ec.alignment = align("center" if ci==1 else "left", False)
            ec.font = font(size=10)
            ec.border = BORDER

    ws.freeze_panes = "A2"

# ─── メイン ────────────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()

    print("マスタシート作成中...")
    write_master_sheet(wb)

    print("個人シート作成中...")
    for emp in EMPLOYEES:
        print(f"  {emp['name']} ({emp['dept']}・{emp['role']})")
        write_individual_sheet(wb, emp)

    print("分析シート作成中...")
    write_analysis_sheet(wb, "Q1分析（11月〜1月）", [0], "Q1分析")
    write_analysis_sheet(wb, "Q2分析（2月〜4月）",  [1], "Q2分析")
    write_analysis_sheet(wb, "Q3分析（5月〜7月）",  [2], "Q3分析")
    write_analysis_sheet(wb, "Q4分析（8月〜10月）", [3], "Q4分析")
    write_analysis_sheet(wb, "上半期分析（Q1＋Q2）", [0, 1], "上半期分析")
    write_analysis_sheet(wb, "下半期分析（Q3＋Q4）", [2, 3], "下半期分析")

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    data = out.read()

    with open("/tmp/eval_sheet.xlsx", "wb") as f:
        f.write(data)

    b64 = base64.b64encode(data).decode()
    with open("/tmp/eval_sheet_b64.txt", "w") as f:
        f.write(b64)

    print(f"\n完了: {len(data):,} bytes")
    print(f"シート数: {len(wb.sheetnames)}")
    print("シート一覧:", wb.sheetnames)
    return b64

if __name__ == "__main__":
    main()
