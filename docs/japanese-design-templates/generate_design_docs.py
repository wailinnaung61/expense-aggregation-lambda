#!/usr/bin/env python3
"""
日本語学習管理システム（JLMS）向け
要件定義書・基本設計書・詳細設計書 プロフェッショナルExcel生成スクリプト
"""

from openpyxl import Workbook
from openpyxl.styles import (
    Font, Fill, PatternFill, Border, Side, Alignment, NamedStyle, Protection
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import PieChart, Reference, BarChart
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy
from pathlib import Path

OUT_DIR = Path(__file__).resolve().parent

# ─── Color Palette (corporate navy / teal — not purple AI default) ───
NAVY = "1B3A4B"
TEAL = "0D7377"
TEAL_LIGHT = "14919B"
ACCENT = "C73E1D"
GOLD = "B8860B"
WHITE = "FFFFFF"
BG_LIGHT = "F7F9FA"
BG_ALT = "E8F1F2"
BG_HEADER = "1B3A4B"
BG_SUB = "0D7377"
BG_SECTION = "2C5F6E"
BG_WARN = "FFF3E0"
BG_OK = "E8F5E9"
BG_NG = "FFEBEE"
GRAY = "6B7280"
BORDER_C = "CBD5E1"
BLACK = "1A1A1A"

thin = Border(
    left=Side(style="thin", color=BORDER_C),
    right=Side(style="thin", color=BORDER_C),
    top=Side(style="thin", color=BORDER_C),
    bottom=Side(style="thin", color=BORDER_C),
)
medium = Border(
    left=Side(style="medium", color=NAVY),
    right=Side(style="medium", color=NAVY),
    top=Side(style="medium", color=NAVY),
    bottom=Side(style="medium", color=NAVY),
)
thick_bottom = Border(
    left=Side(style="thin", color=BORDER_C),
    right=Side(style="thin", color=BORDER_C),
    top=Side(style="thin", color=BORDER_C),
    bottom=Side(style="medium", color=TEAL),
)

fill_navy = PatternFill("solid", fgColor=NAVY)
fill_teal = PatternFill("solid", fgColor=TEAL)
fill_section = PatternFill("solid", fgColor=BG_SECTION)
fill_light = PatternFill("solid", fgColor=BG_LIGHT)
fill_alt = PatternFill("solid", fgColor=BG_ALT)
fill_white = PatternFill("solid", fgColor=WHITE)
fill_gold = PatternFill("solid", fgColor=GOLD)
fill_warn = PatternFill("solid", fgColor=BG_WARN)
fill_ok = PatternFill("solid", fgColor=BG_OK)
fill_accent = PatternFill("solid", fgColor=ACCENT)

font_title = Font(name="Yu Gothic", size=22, bold=True, color=WHITE)
font_subtitle = Font(name="Yu Gothic", size=14, bold=True, color=WHITE)
font_h1 = Font(name="Yu Gothic", size=14, bold=True, color=WHITE)
font_h2 = Font(name="Yu Gothic", size=11, bold=True, color=WHITE)
font_h3 = Font(name="Yu Gothic", size=10, bold=True, color=NAVY)
font_label = Font(name="Yu Gothic", size=9, bold=True, color=NAVY)
font_body = Font(name="Yu Gothic", size=9, color=BLACK)
font_body_w = Font(name="Yu Gothic", size=9, color=WHITE)
font_small = Font(name="Yu Gothic", size=8, color=GRAY)
font_cover_meta = Font(name="Yu Gothic", size=11, color=BLACK)
font_cover_val = Font(name="Yu Gothic", size=11, bold=True, color=NAVY)
font_num = Font(name="Consolas", size=9, color=BLACK)

align_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_r = Alignment(horizontal="right", vertical="center", wrap_text=True)
align_t = Alignment(horizontal="left", vertical="top", wrap_text=True)


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def style_range(ws, cells, fill=None, font=None, border=None, align=None):
    for row in ws[cells]:
        for cell in row:
            if fill:
                cell.fill = fill
            if font:
                cell.font = font
            if border:
                cell.border = border
            if align:
                cell.alignment = align


def merge_write(ws, start, end, value, fill=None, font=None, align=None, border=None):
    ws.merge_cells(f"{start}:{end}")
    cell = ws[start]
    cell.value = value
    if fill:
        cell.fill = fill
    if font:
        cell.font = font
    if align:
        cell.alignment = align
    # apply to all merged cells
    for row in ws[f"{start}:{end}"]:
        for c in row:
            if fill:
                c.fill = fill
            if border:
                c.border = border
            if align:
                c.alignment = align
            if font:
                c.font = font
    return cell


def header_bar(ws, row, cols, title, fill=fill_navy, font=font_h1, height=28):
    end = get_column_letter(cols)
    merge_write(ws, f"A{row}", f"{end}{row}", title, fill=fill, font=font, align=align_c, border=thin)
    ws.row_dimensions[row].height = height


def section_bar(ws, row, cols, title, fill=fill_teal):
    end = get_column_letter(cols)
    merge_write(ws, f"A{row}", f"{end}{row}", title, fill=fill, font=font_h2, align=align_l, border=thin)
    ws.row_dimensions[row].height = 22


def kv_row(ws, row, label, value, label_cols=2, value_end=8):
    # label in A-B, value in C-end
    merge_write(ws, f"A{row}", f"{get_column_letter(label_cols)}{row}", label,
                fill=fill_alt, font=font_label, align=align_c, border=thin)
    merge_write(ws, f"{get_column_letter(label_cols+1)}{row}", f"{get_column_letter(value_end)}{row}",
                value, fill=fill_white, font=font_body, align=align_l, border=thin)
    ws.row_dimensions[row].height = 20


def table_header(ws, row, headers, fill=fill_navy):
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.fill = fill
        cell.font = font_body_w
        cell.alignment = align_c
        cell.border = thin
    ws.row_dimensions[row].height = 22


def table_row(ws, row, values, alt=False, heights=18):
    fill = fill_alt if alt else fill_white
    for i, v in enumerate(values, 1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.fill = fill
        cell.font = font_body
        cell.alignment = align_l if i > 1 else align_c
        cell.border = thin
    ws.row_dimensions[row].height = heights


def write_cover(ws, doc_type, doc_id, version="1.0", status="正式版"):
    """Professional Japanese document cover sheet."""
    set_col_widths(ws, [14, 16, 16, 16, 16, 14, 14, 14])
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 12

    # Top brand bar
    merge_write(ws, "A2", "H2", "株式会社 サクラ教育テクノロジー ／ Sakura EduTech Inc.",
                fill=fill_navy, font=Font(name="Yu Gothic", size=10, color=WHITE), align=align_c)
    ws.row_dimensions[2].height = 24

    merge_write(ws, "A3", "H3", "", fill=fill_teal)
    ws.row_dimensions[3].height = 6

    # Document type badge
    merge_write(ws, "A5", "H5", "PROJECT DOCUMENTATION ｜ システム開発標準ドキュメント",
                fill=fill_light, font=font_small, align=align_c)
    ws.row_dimensions[5].height = 18

    merge_write(ws, "A7", "H7", doc_type,
                fill=fill_navy, font=font_title, align=align_c)
    ws.row_dimensions[7].height = 48

    merge_write(ws, "A8", "H8", "日本語学習管理システム（JLMS：Japanese Learning Management System）",
                fill=fill_section, font=font_subtitle, align=align_c)
    ws.row_dimensions[8].height = 32

    merge_write(ws, "A9", "H9", "対象フェーズ：要件定義 → 基本設計 → 詳細設計　｜　教育・研修用完成見本（実案件水準）",
                fill=fill_alt, font=font_small, align=align_c)

    # Meta table
    r = 12
    metas = [
        ("文書番号", doc_id, "版数", version),
        ("文書状態", status, "機密区分", "社外秘（Confidential）"),
        ("プロジェクト名", "JLMS 構築プロジェクト", "案件コード", "PRJ-JLMS-2026-001"),
        ("顧客／発注元", "株式会社グローバル人材育成機構", "開発ベンダー", "株式会社サクラ教育テクノロジー"),
        ("作成部署", "システム開発本部 設計グループ", "作成者", "設計責任者　山田 太郎"),
        ("作成日", "2026-04-01", "最終更新日", "2026-08-01"),
        ("承認者（顧客）", "情報システム部長　鈴木 花子", "承認日", "2026-08-03"),
        ("承認者（ベンダー）", "プロジェクトマネージャ　佐藤 一郎", "承認日", "2026-08-02"),
    ]
    for label1, val1, label2, val2 in metas:
        merge_write(ws, f"A{r}", f"B{r}", label1, fill=fill_alt, font=font_label, align=align_c, border=thin)
        merge_write(ws, f"C{r}", f"D{r}", val1, fill=fill_white, font=font_cover_val, align=align_l, border=thin)
        merge_write(ws, f"E{r}", f"F{r}", label2, fill=fill_alt, font=font_label, align=align_c, border=thin)
        merge_write(ws, f"G{r}", f"H{r}", val2, fill=fill_white, font=font_cover_val, align=align_l, border=thin)
        ws.row_dimensions[r].height = 24
        r += 1

    r += 1
    merge_write(ws, f"A{r}", f"H{r}", "改訂履歴（Revision History）",
                fill=fill_teal, font=font_h2, align=align_l, border=thin)
    r += 1
    headers = ["版", "改訂日", "改訂箇所", "改訂内容", "作成者", "確認者", "承認者", "備考"]
    # span across - use 8 cols with combined content in fewer conceptually
    table_header(ws, r, ["版", "改訂日", "改訂箇所", "改訂内容", "作成者", "確認者", "承認者", "備考"])
    r += 1
    revs = [
        ["0.1", "2026-04-01", "全体", "初版ドラフト作成", "山田", "佐藤", "—", "内部レビュー用"],
        ["0.5", "2026-05-15", "機能要件", "顧客レビュー指摘反映", "山田", "佐藤", "—", "中間版"],
        ["0.9", "2026-07-10", "非機能・IF", "性能・外部IF確定", "山田", "佐藤", "鈴木", "承認前最終"],
        ["1.0", "2026-08-01", "全体", "正式版として確定", "山田", "佐藤", "鈴木", "ベースライン"],
    ]
    for i, rev in enumerate(revs):
        table_row(ws, r, rev, alt=i % 2 == 1)
        r += 1

    r += 2
    merge_write(ws, f"A{r}", f"H{r}",
                "本資料は教育・研修目的の完成見本です。実案件の文書体系・表記・粒度を再現しています。転用時は案件固有情報に置換してください。",
                fill=fill_warn, font=font_small, align=align_l, border=thin)
    r += 2
    merge_write(ws, f"A{r}", f"H{r}",
                "© 2026 Sakura EduTech Inc. All Rights Reserved. ｜ Template: JLMS-DOC-STD-v1.0",
                fill=fill_navy, font=Font(name="Yu Gothic", size=8, color=WHITE), align=align_c)
    ws.print_title_rows = "1:3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def write_toc(ws, items, title="目次（Table of Contents）"):
    set_col_widths(ws, [8, 50, 14, 18, 20, 20, 14, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, title)
    table_header(ws, 3, ["No.", "章・節", "シート名", "ページ／位置", "担当", "状態", "必須", "備考"])
    for i, (no, chapter, sheet, page, owner, status, req, note) in enumerate(items):
        table_row(ws, 4 + i, [no, chapter, sheet, page, owner, status, req, note], alt=i % 2 == 1)
        # status coloring
        cell = ws.cell(row=4 + i, column=6)
        if status == "確定":
            cell.fill = fill_ok
        elif status == "レビュー中":
            cell.fill = fill_warn


# ═══════════════════════════════════════════════════════════════════
# 1. 要件定義書
# ═══════════════════════════════════════════════════════════════════
def build_requirements():
    wb = Workbook()

    # --- 表紙 ---
    ws = wb.active
    ws.title = "00_表紙"
    write_cover(ws, "要件定義書", "DOC-JLMS-RD-001", "1.0", "正式版（Approved）")

    # --- 目次 ---
    ws = wb.create_sheet("01_目次")
    toc = [
        ("1", "はじめに（目的・背景・範囲）", "02_はじめに", "—", "PM", "確定", "◎", ""),
        ("2", "用語定義・参照資料", "03_用語・参照", "—", "設計", "確定", "◎", ""),
        ("3", "現状業務・課題", "04_現状業務", "—", "BA", "確定", "◎", ""),
        ("4", "システム化方針・スコープ", "05_スコープ", "—", "PM", "確定", "◎", ""),
        ("5", "アクター・ユースケース", "06_ユースケース", "—", "設計", "確定", "◎", ""),
        ("6", "機能要件一覧", "07_機能要件", "—", "設計", "確定", "◎", "優先度付き"),
        ("7", "機能要件詳細", "08_機能詳細", "—", "設計", "確定", "◎", ""),
        ("8", "非機能要件", "09_非機能要件", "—", "アーキ", "確定", "◎", "ISO25010準拠"),
        ("9", "画面・帳票一覧", "10_画面帳票一覧", "—", "設計", "確定", "◎", ""),
        ("10", "外部インターフェース要件", "11_外部IF", "—", "設計", "確定", "◎", ""),
        ("11", "データ要件", "12_データ要件", "—", "DBA", "確定", "◎", ""),
        ("12", "移行・運用・制約", "13_移行運用制約", "—", "PM", "確定", "◎", ""),
        ("13", "要件トレーサビリティ", "14_トレーサビリティ", "—", "QA", "確定", "◎", ""),
        ("14", "承認記録", "15_承認", "—", "PM", "確定", "◎", ""),
    ]
    write_toc(ws, toc, "要件定義書　目次")

    # --- はじめに ---
    ws = wb.create_sheet("02_はじめに")
    set_col_widths(ws, [16, 18, 18, 18, 18, 16, 14, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "1. はじめに")
    section_bar(ws, 3, 8, "1.1 目的（Purpose）")
    merge_write(ws, "A4", "H6",
                "本要件定義書は、日本語学習管理システム（以下、JLMS）の構築にあたり、"
                "発注者と開発ベンダーが合意すべき業務要件・機能要件・非機能要件・制約事項を明確化し、"
                "以降の基本設計・詳細設計・実装・試験のベースラインとすることを目的とする。\n"
                "本システムは、企業・教育機関における日本語学習者の学習進捗管理、教材配信、"
                "習熟度測定、講師・管理者の業務効率化を実現する。",
                fill=fill_white, font=font_body, align=align_t, border=thin)
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 18

    section_bar(ws, 8, 8, "1.2 背景（Background）")
    merge_write(ws, "A9", "H11",
                "・海外拠点・外国人材の増加に伴い、社内日本語教育の標準化・可視化ニーズが高まっている。\n"
                "・現行はExcel・紙ベースの進捗管理であり、受講状況のリアルタイム把握・教材版管理が困難。\n"
                "・JLPT相当レベル判定と学習パスの自動提案により、学習効果と管理者工数削減を同時に実現する。\n"
                "・2026年度中の本番稼働を目標とし、パイロット（受講者200名）→全社展開（受講者2,000名）の段階導入とする。",
                fill=fill_white, font=font_body, align=align_t, border=thin)

    section_bar(ws, 13, 8, "1.3 適用範囲（Scope of this Document）")
    table_header(ws, 14, ["区分", "対象", "対象外", "備考", "", "", "", ""])
    # custom rows
    rows_scope = [
        ["業務", "学習管理、教材管理、試験、レポート、権限", "給与・人事マスタ本体の改修", "人事は参照のみ"],
        ["システム", "Webアプリ（学習者／講師／管理者）、バッチ、外部IF", "ネイティブモバイルアプリ（Phase2）", "Responsive Web対応"],
        ["データ", "受講者・教材・学習履歴・試験結果", "既存LMS全データの完全移行", "必要データのみCSV移行"],
        ["期間", "要件定義〜詳細設計完了までを本書で担保", "運用設計の詳細手順書", "運用は別紙"],
    ]
    for i, row in enumerate(rows_scope):
        table_row(ws, 15 + i, row, alt=i % 2 == 1, heights=28)

    section_bar(ws, 20, 8, "1.4 前提条件・制約条件")
    table_header(ws, 21, ["ID", "区分", "内容", "影響", "対応方針", "確定日", "責任者", "状態"])
    presup = [
        ["PR-01", "前提", "認証は顧客Azure AD（Entra ID）SAML/OIDC連携", "ログイン設計", "標準プロトコル採用", "2026-05-01", "アーキ", "確定"],
        ["PR-02", "前提", "クラウドはAWS（東京リージョン）を利用", "構成・コスト", "Well-Architected準拠", "2026-04-15", "インフラ", "確定"],
        ["CN-01", "制約", "個人情報は日本国内に保管（越境転送不可）", "DR設計", "東京＋大阪マルチAZ", "2026-05-01", "情シス", "確定"],
        ["CN-02", "制約", "初期同時接続ピーク 500セッション", "性能要件", "負荷試験で検証", "2026-06-01", "アーキ", "確定"],
        ["CN-03", "制約", "既存人事システムは参照APIのみ提供", "マスタ同期", "日次バッチ＋差分API", "2026-06-15", "設計", "確定"],
        ["AS-01", "仮定", "教材コンテンツは発注側が提供（著作権クリア済）", "教材管理", "納品フォーマット定義", "2026-05-20", "PM", "確定"],
    ]
    for i, row in enumerate(presup):
        table_row(ws, 22 + i, row, alt=i % 2 == 1)

    # --- 用語・参照 ---
    ws = wb.create_sheet("03_用語・参照")
    set_col_widths(ws, [10, 22, 40, 22, 14, 14, 14, 12])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "2. 用語定義・参照資料")
    section_bar(ws, 3, 8, "2.1 用語定義（Glossary）")
    table_header(ws, 4, ["No.", "用語", "定義", "英語／略称", "関連章", "出典", "備考", "確定"])
    terms = [
        ["1", "JLMS", "本プロジェクトで構築する日本語学習管理システム", "Japanese LMS", "全体", "本書", "", "○"],
        ["2", "学習者", "JLMS上で日本語を学習するエンドユーザ", "Learner", "UC", "本書", "社員・外部含む", "○"],
        ["3", "講師", "教材・課題・評価を担当するユーザ", "Instructor", "UC", "本書", "", "○"],
        ["4", "管理者", "組織・権限・マスタを管理するユーザ", "Admin", "UC", "本書", "情シス／教育担当", "○"],
        ["5", "学習パス", "レベルに応じた推奨教材・試験の順序付き計画", "Learning Path", "機能", "本書", "自動／手動", "○"],
        ["6", "習熟度", "技能別（読・書・聴・話）の到達度指標", "Proficiency", "機能", "本書", "0–100スコア", "○"],
        ["7", "JLPT相当", "日本語能力試験レベル（N5〜N1）へのマッピング", "JLPT-equiv.", "機能", "JLPT", "公式試験ではない", "○"],
        ["8", "テナント", "顧客組織単位の論理分離領域", "Tenant", "非機能", "本書", "マルチテナント", "○"],
        ["9", "コンテンツ包", "教材・問題・音声・字幕をまとめた配布単位", "Content Pack", "データ", "本書", "版管理対象", "○"],
        ["10", "SLO", "サービスレベル目標（可用性・応答時間等）", "Service Level Objective", "非機能", "SRE", "", "○"],
    ]
    for i, t in enumerate(terms):
        table_row(ws, 5 + i, t, alt=i % 2 == 1, heights=24)

    section_bar(ws, 17, 8, "2.2 参照資料（References）")
    table_header(ws, 18, ["ID", "資料名", "版", "発行元", "日付", "保管場所", "用途", "必須"])
    refs = [
        ["REF-01", "RFP／提案依頼書", "1.2", "発注元", "2026-02-01", "SharePoint/RFP", "要件根拠", "◎"],
        ["REF-02", "提案書・見積書", "A", "サクラEduTech", "2026-03-01", "契約フォルダ", "範囲根拠", "◎"],
        ["REF-03", "情報セキュリティポリシー", "3.0", "発注元", "2025-10-01", "情シス", "非機能", "◎"],
        ["REF-04", "個人情報取扱規程", "2.1", "発注元", "2025-04-01", "法務", "データ要件", "◎"],
        ["REF-05", "人事システム API仕様書", "4.3", "人事ベンダー", "2026-01-15", "IFフォルダ", "外部IF", "◎"],
        ["REF-06", "AWS Well-Architected Framework", "2025", "AWS", "—", "外部公開", "アーキ方針", "○"],
        ["REF-07", "JIS X 8341-3 アクセシビリティ", "2016", "JIS", "—", "外部公開", "UI要件", "○"],
        ["REF-08", "ISO/IEC 25010 品質モデル", "2011", "ISO", "—", "外部公開", "非機能分類", "○"],
    ]
    for i, t in enumerate(refs):
        table_row(ws, 19 + i, t, alt=i % 2 == 1)

    # --- 現状業務 ---
    ws = wb.create_sheet("04_現状業務")
    set_col_widths(ws, [10, 18, 28, 28, 18, 16, 14, 12])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "3. 現状業務・課題（As-Is）")
    section_bar(ws, 3, 8, "3.1 現状業務フロー概要")
    table_header(ws, 4, ["工程", "担当", "現行手段", "成果物", "所要時間", "課題ID", "頻度", "備考"])
    asis = [
        ["受講者登録", "教育担当", "Excel＋メール", "受講者名簿.xlsx", "2人日/月", "IS-01", "随時", "手入力ミス多"],
        ["教材配布", "講師", "共有フォルダ", "PDF/音声", "0.5人日/回", "IS-02", "週次", "版が混在"],
        ["学習実施", "学習者", "紙ワークブック", "宿題提出", "—", "IS-03", "毎日", "進捗不可視"],
        ["進捗確認", "講師", "口頭／Excel", "進捗表", "1人日/週", "IS-04", "週次", "集計遅延"],
        ["習熟度測定", "講師", "紙テスト採点", "点数表", "2人日/回", "IS-05", "月次", "技能別不可"],
        ["レポート", "管理者", "手集計", "経営報告", "3人日/月", "IS-06", "月次", "属人化"],
    ]
    for i, t in enumerate(asis):
        table_row(ws, 5 + i, t, alt=i % 2 == 1)

    section_bar(ws, 12, 8, "3.2 課題一覧（Issues）とTo-Be方針")
    table_header(ws, 13, ["課題ID", "課題", "影響度", "緊急度", "To-Be方針", "対応要件ID", "KPI", "優先"])
    issues = [
        ["IS-01", "受講者マスタ二重管理・入力ミス", "高", "高", "人事連携＋単一マスタ", "FR-USR-01", "入力ミス率0.1%以下", "P1"],
        ["IS-02", "教材版の混在・古い教材配信", "高", "中", "コンテンツ版管理・公開制御", "FR-CNT-02", "誤配信0件", "P1"],
        ["IS-03", "学習進捗がリアルタイムに見えない", "高", "高", "学習イベント自動収集", "FR-LRN-01", "進捗遅延検知24h以内", "P1"],
        ["IS-04", "講師の進捗確認工数過大", "中", "高", "ダッシュボード・アラート", "FR-RPT-01", "確認工数70%削減", "P1"],
        ["IS-05", "技能別習熟度が測れない", "高", "中", "技能別アセスメント", "FR-ASM-01", "技能別スコア可視化", "P2"],
        ["IS-06", "経営報告の手作業集計", "中", "中", "定型レポート自動生成", "FR-RPT-03", "集計工数90%削減", "P2"],
        ["IS-07", "学習者のモチベーション維持困難", "中", "中", "バッジ・ストリーク・通知", "FR-LRN-05", "継続率+15pt", "P3"],
    ]
    for i, t in enumerate(issues):
        table_row(ws, 14 + i, t, alt=i % 2 == 1, heights=22)

    # --- スコープ ---
    ws = wb.create_sheet("05_スコープ")
    set_col_widths(ws, [12, 20, 14, 14, 16, 16, 16, 20])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "4. システム化方針・スコープ（To-Be Scope）")
    section_bar(ws, 3, 8, "4.1 システム化方針")
    merge_write(ws, "A4", "H5",
                "①クラウドネイティブ（AWS）で迅速なスケールと運用自動化を実現する。\n"
                "②マルチテナント対応により将来の外部提供（SaaS化）を見据える。\n"
                "③学習データはイベント駆動で収集し、分析・パーソナライズの基盤とする。\n"
                "④アクセシビリティ（JIS X 8341-3 AA準拠目標）と多言語UI（日・英）を標準装備する。",
                fill=fill_white, font=font_body, align=align_t, border=thin)

    section_bar(ws, 7, 8, "4.2 スコープマトリクス（In / Out / Future）")
    table_header(ws, 8, ["領域", "機能群", "Phase1", "Phase2", "Out of Scope", "根拠", "依存", "備考"])
    scope = [
        ["学習", "教材閲覧・進捗・ブックマーク", "In", "—", "—", "RFP必須", "—", ""],
        ["学習", "学習パス自動提案", "In", "強化学習最適化", "—", "差別化", "習熟度", ""],
        ["試験", "四択・記述・聴解", "In", "スピーキングAI採点", "公式JLPT代行", "RFP", "音声IF", ""],
        ["教材", "版管理・公開・タグ", "In", "マーケットプレイス", "コンテンツ制作自体", "著作権", "—", ""],
        ["管理", "ユーザ・組織・権限", "In", "—", "給与連携", "セキュリティ", "Azure AD", ""],
        ["分析", "進捗ダッシュボード", "In", "予測離脱アラート", "BI全社統合", "KPI", "—", ""],
        ["通知", "メール・アプリ内", "In", "Slack/Teams", "SMS", "コスト", "SES", ""],
        ["モバイル", "Responsive Web", "In", "PWA/ネイティブ", "—", "利用シーン", "—", ""],
        ["課金", "—", "Out", "In（SaaS）", "Phase1対象外", "契約", "—", "将来"],
    ]
    for i, t in enumerate(scope):
        r = 9 + i
        table_row(ws, r, t, alt=i % 2 == 1)
        cell = ws.cell(row=r, column=3)
        if t[2] == "In":
            cell.fill = fill_ok
        elif t[2] == "Out":
            cell.fill = fill_warn

    section_bar(ws, 20, 8, "4.3 ステークホルダ")
    table_header(ws, 21, ["役割", "氏名／組織", "関心事", "影響度", "関与度", "コミュニケーション", "承認権限", "備考"])
    stakeholders = [
        ["スポンサー", "グローバル人材育成機構 役員", "投資対効果・期限", "高", "中", "月次Steering", "最終承認", ""],
        ["プロダクトオーナー", "教育企画部 部長", "学習効果・現場適合", "高", "高", "週次", "要件承認", ""],
        ["情シス", "情報システム部", "セキュリティ・運用", "高", "高", "週次", "非機能承認", ""],
        ["現場講師代表", "日本語教育センター", "使いやすさ・工数", "中", "高", "隔週ワークショップ", "画面確認", ""],
        ["学習者代表", "海外拠点パイロット", "学習体験", "中", "中", "UAT", "受入意見", ""],
        ["PM（ベンダー）", "サクラEduTech 佐藤", "納期・品質・範囲", "高", "高", "日次", "成果物承認", ""],
    ]
    for i, t in enumerate(stakeholders):
        table_row(ws, 22 + i, t, alt=i % 2 == 1, heights=22)

    # --- ユースケース ---
    ws = wb.create_sheet("06_ユースケース")
    set_col_widths(ws, [12, 14, 22, 28, 14, 12, 12, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "5. アクター・ユースケース")
    section_bar(ws, 3, 8, "5.1 アクター定義")
    table_header(ws, 4, ["アクターID", "名称", "種別", "説明", "認証", "主な端末", "人数規模", "備考"])
    actors = [
        ["ACT-01", "学習者", "人", "日本語を学習するユーザ", "Azure AD / 招待", "PC・スマホ", "2,000", ""],
        ["ACT-02", "講師", "人", "教材・評価・フィードバック担当", "Azure AD", "PC", "80", ""],
        ["ACT-03", "組織管理者", "人", "自組織のユーザ・レポート管理", "Azure AD", "PC", "20", ""],
        ["ACT-04", "システム管理者", "人", "テナント全体・マスタ・監査", "Azure AD＋MFA", "PC", "5", ""],
        ["ACT-05", "人事システム", "外部", "社員マスタ提供", "API Key / mTLS", "—", "—", "日次同期"],
        ["ACT-06", "メールサービス", "外部", "通知配信（AWS SES）", "IAM", "—", "—", ""],
        ["ACT-07", "オブジェクトストレージ", "外部", "教材バイナリ保管（S3）", "IAM", "—", "—", ""],
    ]
    for i, t in enumerate(actors):
        table_row(ws, 5 + i, t, alt=i % 2 == 1)

    section_bar(ws, 14, 8, "5.2 ユースケース一覧")
    table_header(ws, 15, ["UC-ID", "名称", "主アクター", "概要", "事前条件", "優先度", "Phase", "機能ID"])
    ucs = [
        ["UC-01", "ログイン／ログアウト", "学習者他", "IdP経由で認証しセッション確立", "アカウント有効", "P1", "1", "FR-AUTH-01"],
        ["UC-02", "学習パスを確認する", "学習者", "推奨教材・期限・進捗を表示", "パス割当済", "P1", "1", "FR-LRN-02"],
        ["UC-03", "教材を学習する", "学習者", "教材閲覧・動画視聴・理解度チェック", "公開教材", "P1", "1", "FR-LRN-01"],
        ["UC-04", "課題を提出する", "学習者", "記述・音声課題を提出", "課題公開中", "P1", "1", "FR-LRN-03"],
        ["UC-05", "試験を受験する", "学習者", "制限時間付きアセスメント実施", "受験資格", "P1", "1", "FR-ASM-01"],
        ["UC-06", "フィードバックを確認", "学習者", "採点結果・コメント閲覧", "採点完了", "P1", "1", "FR-ASM-03"],
        ["UC-07", "教材を公開する", "講師", "コンテンツ包を版指定で公開", "承認権限", "P1", "1", "FR-CNT-02"],
        ["UC-08", "提出物を評価する", "講師", "採点・コメント・差戻し", "提出あり", "P1", "1", "FR-ASM-02"],
        ["UC-09", "クラス進捗を監視", "講師", "ダッシュボードで遅延検知", "クラス担当", "P1", "1", "FR-RPT-01"],
        ["UC-10", "ユーザを招待・無効化", "組織管理者", "メンバーライフサイクル管理", "管理権限", "P1", "1", "FR-USR-01"],
        ["UC-11", "組織レポート出力", "組織管理者", "CSV/PDFで進捗レポート", "管理権限", "P2", "1", "FR-RPT-03"],
        ["UC-12", "人事マスタ同期", "人事システム", "差分ユーザを取り込む", "IF疎通", "P1", "1", "FR-IF-01"],
        ["UC-13", "監査ログ照会", "システム管理者", "操作履歴を検索・エクスポート", "監査権限", "P1", "1", "FR-AUD-01"],
        ["UC-14", "学習リマインド通知", "メールサービス", "未学習者へ自動通知", "通知設定ON", "P2", "1", "FR-NTF-01"],
    ]
    for i, t in enumerate(ucs):
        table_row(ws, 16 + i, t, alt=i % 2 == 1, heights=20)

    # --- 機能要件一覧 ---
    ws = wb.create_sheet("07_機能要件")
    set_col_widths(ws, [14, 12, 22, 36, 10, 10, 10, 12])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "6. 機能要件一覧（Functional Requirements）")
    merge_write(ws, "A2", "H2",
                "優先度: P1=Must（本番必須） / P2=Should（Phase1望ましい） / P3=Could（余裕があれば）　｜　MoSCoW準拠",
                fill=fill_light, font=font_small, align=align_l)
    table_header(ws, 3, ["要件ID", "カテゴリ", "要件名", "概要", "優先度", "Phase", "複雑度", "状態"])
    frs = [
        ["FR-AUTH-01", "認証", "シングルサインオン", "Entra ID連携によるSSOログイン／ログアウト", "P1", "1", "中", "確定"],
        ["FR-AUTH-02", "認証", "ロールベースアクセス制御", "学習者/講師/組織管理者/システム管理者のRBAC", "P1", "1", "中", "確定"],
        ["FR-USR-01", "ユーザ", "ユーザライフサイクル", "招待・有効化・無効化・属性編集・組織所属", "P1", "1", "中", "確定"],
        ["FR-USR-02", "ユーザ", "人事マスタ取込", "日次バッチおよび差分APIによる自動同期", "P1", "1", "高", "確定"],
        ["FR-CNT-01", "教材", "コンテンツ包管理", "教材メタデータ・ファイル・タグ・カテゴリ管理", "P1", "1", "中", "確定"],
        ["FR-CNT-02", "教材", "版管理・公開制御", "セマンティックバージョン、公開/非公開、予約公開", "P1", "1", "高", "確定"],
        ["FR-LRN-01", "学習", "教材学習・進捗記録", "閲覧/視聴イベント、完了条件、しおり", "P1", "1", "高", "確定"],
        ["FR-LRN-02", "学習", "学習パス", "レベル別パス割当、進捗率、期限管理", "P1", "1", "高", "確定"],
        ["FR-LRN-03", "学習", "課題提出", "テキスト/ファイル/音声提出、締切、再提出", "P1", "1", "中", "確定"],
        ["FR-LRN-04", "学習", "語彙・漢字デッキ", "SRS（間隔反復）による暗記学習", "P2", "1", "高", "確定"],
        ["FR-LRN-05", "学習", "ゲーミフィケーション", "ストリーク、バッジ、週間目標", "P3", "1", "低", "確定"],
        ["FR-ASM-01", "評価", "アセスメント実施", "四択/記述/聴解、制限時間、不正対策（離脱検知）", "P1", "1", "高", "確定"],
        ["FR-ASM-02", "評価", "手動採点", "講師による採点・コメント・差戻しワークフロー", "P1", "1", "中", "確定"],
        ["FR-ASM-03", "評価", "自動採点・結果表示", "客観問題の即時採点と技能別スコア表示", "P1", "1", "中", "確定"],
        ["FR-ASM-04", "評価", "JLPT相当レベル推定", "累積スコアからN5–N1相当を推定表示", "P2", "1", "高", "確定"],
        ["FR-RPT-01", "分析", "講師ダッシュボード", "クラス進捗、遅延者、提出状況の可視化", "P1", "1", "中", "確定"],
        ["FR-RPT-02", "分析", "学習者マイページ", "個人進捗、弱点技能、次のアクション", "P1", "1", "中", "確定"],
        ["FR-RPT-03", "分析", "組織レポート", "部署別進捗CSV/PDF、定期配信", "P2", "1", "中", "確定"],
        ["FR-NTF-01", "通知", "リマインド通知", "未学習・締切前・採点完了のメール/アプリ内通知", "P2", "1", "低", "確定"],
        ["FR-IF-01", "連携", "人事IF", "ユーザ差分取得・突合・エラーハンドリング", "P1", "1", "高", "確定"],
        ["FR-AUD-01", "監査", "監査ログ", "認証・権限変更・個人情報参照の記録と検索", "P1", "1", "中", "確定"],
        ["FR-SYS-01", "基盤", "マルチテナント", "テナント分離（論理）、ブランド設定", "P1", "1", "高", "確定"],
        ["FR-SYS-02", "基盤", "多言語UI", "日本語／英語切替（教材本文は別）", "P1", "1", "低", "確定"],
    ]
    for i, t in enumerate(frs):
        r = 4 + i
        table_row(ws, r, t, alt=i % 2 == 1)
        prio = ws.cell(row=r, column=5)
        if t[4] == "P1":
            prio.fill = PatternFill("solid", fgColor="FECACA")
            prio.font = Font(name="Yu Gothic", size=9, bold=True, color=ACCENT)
        elif t[4] == "P2":
            prio.fill = fill_warn
        else:
            prio.fill = fill_ok

    # summary counts
    r = 4 + len(frs) + 1
    section_bar(ws, r, 8, "6.1 要件サマリ")
    r += 1
    table_header(ws, r, ["優先度", "件数", "割合", "Phase1必須", "備考", "", "", ""])
    r += 1
    table_row(ws, r, ["P1 Must", "16", "70%", "Yes", "本番稼働の必須条件", "", "", ""], False)
    r += 1
    table_row(ws, r, ["P2 Should", "5", "22%", "Partial", "可能な限りPhase1", "", "", ""], True)
    r += 1
    table_row(ws, r, ["P3 Could", "2", "8%", "No", "時間があれば", "", "", ""], False)

    # --- 機能詳細 ---
    ws = wb.create_sheet("08_機能詳細")
    set_col_widths(ws, [14, 16, 40, 20, 14, 14, 14, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "7. 機能要件詳細（代表ユースケース記述）")

    def write_fr_detail(ws, start, fr_id, name, actor, trigger, pre, basic, alt, post, rules, conf):
        r = start
        merge_write(ws, f"A{r}", f"H{r}", f"{fr_id}　{name}", fill=fill_section, font=font_h2, align=align_l, border=thin)
        r += 1
        details = [
            ("主アクター", actor),
            ("トリガー", trigger),
            ("事前条件", pre),
            ("基本フロー", basic),
            ("代替／例外", alt),
            ("事後条件", post),
            ("ビジネスルール", rules),
            ("受入条件（要約）", conf),
        ]
        for label, val in details:
            merge_write(ws, f"A{r}", f"B{r}", label, fill=fill_alt, font=font_label, align=align_c, border=thin)
            merge_write(ws, f"C{r}", f"H{r}", val, fill=fill_white, font=font_body, align=align_t, border=thin)
            ws.row_dimensions[r].height = 48 if label in ("基本フロー", "代替／例外", "ビジネスルール", "受入条件（要約）") else 22
            r += 1
        return r + 1

    r = 3
    r = write_fr_detail(
        ws, r, "FR-LRN-01", "教材学習・進捗記録",
        "学習者（ACT-01）",
        "学習パスまたは教材一覧から教材を開く",
        "ユーザが有効／教材が公開済／テナントが有効",
        "1) 教材詳細を表示\n2) 学習を開始（started_at記録）\n3) ページ/動画進捗を定期保存（少なくとも30秒毎またはセクション遷移時）\n4) 完了条件充足で completed を記録\n5) 学習パス進捗率を再計算",
        "E1: セッション切れ→再認証後に再開位置復元\nE2: コンテンツ取得失敗→再試行導線＋エラーログ\nA1: オフライン下書き（将来）はPhase2",
        "学習イベントが保存され、ダッシュボードに反映される",
        "BR1: 完了条件は教材タイプ別に定義（視聴90%／全ページ／確認テスト合格）\nBR2: 同一教材の再学習は履歴を追記（上書きしない）\nBR3: 個人の学習履歴は本人・担当講師・管理者のみ参照可",
        "・進捗が30秒以内に永続化される\n・完了後にパス進捗が即時更新される\n・権限外ユーザは教材本文にアクセスできない",
    )
    r = write_fr_detail(
        ws, r, "FR-ASM-01", "アセスメント実施",
        "学習者（ACT-01）",
        "試験開始ボタン押下",
        "受験資格あり／実施期間内／未提出または再受験許可あり",
        "1) 注意事項・制限時間を表示し同意取得\n2) 問題を逐次表示（ランダム出題設定時はシャッフル）\n3) 回答を自動保存\n4) 時間切れまたは提出でロック\n5) 客観問題は自動採点、主観問題は採点待ちへ",
        "E1: 途中離脱（タブクローズ）→再開ポリシーに従い継続または無効\nE2: 通信断→ローカル一時保存後同期\nA1: 講師による特別再受験許可",
        "回答が提出され、スコアまたは採点待ち状態になる",
        "BR1: 制限時間超過後の回答は受け付けない\nBR2: 離脱検知回数しきい値超過でフラグ付与（不正疑い）\nBR3: 結果公開タイミングは試験設定に従う（即時／一斉公開）",
        "・制限時間が正確に適用される\n・提出後に回答改ざん不可\n・自動採点結果が期待正答と一致する",
    )
    r = write_fr_detail(
        ws, r, "FR-IF-01", "人事IF（ユーザ差分同期）",
        "人事システム（ACT-05）／バッチ",
        "日次スケジュール（JST 02:00）または手動実行",
        "API認証情報有効／対象テナント有効",
        "1) 差分APIから更新ユーザを取得\n2) バリデーション（必須項目・メール形式）\n3) 新規作成／更新／退職（無効化）を適用\n4) 結果サマリを監査ログ・運用通知へ",
        "E1: APIタイムアウト→リトライ3回（指数バックオフ）後に失敗終了\nE2: 部分失敗→成功分コミット、失敗行をエラー表へ\nE3: 大量変更（閾値超）→自動適用停止し承認待ち",
        "JLMSユーザマスタが人事側と整合（差分適用完了）",
        "BR1: メールアドレスをビジネスキーとする\nBR2: 退職日到来で自動無効化（即座に学習不可）\nBR3: 管理者ロールは自動付与しない（手動）",
        "・正常系で差分が全件適用される\n・エラー行が運用画面で確認できる\n・閾値超過時に自動適用されない",
    )

    # --- 非機能要件 ---
    ws = wb.create_sheet("09_非機能要件")
    set_col_widths(ws, [12, 14, 14, 36, 14, 14, 12, 12])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "8. 非機能要件（ISO/IEC 25010 準拠分類）")
    table_header(ws, 3, ["要件ID", "品質特性", "副特性", "要件内容（測定可能指標）", "目標値", "測定方法", "優先", "状態"])
    nfrs = [
        ["NFR-01", "性能効率性", "時間効率性", "主要画面（教材一覧・ダッシュボード）のp95応答時間", "≦ 2.0秒", "APM合成監視", "P1", "確定"],
        ["NFR-02", "性能効率性", "時間効率性", "試験提出APIのp95応答時間", "≦ 1.0秒", "負荷試験", "P1", "確定"],
        ["NFR-03", "性能効率性", "容量", "同時セッション数", "500（Peak）", "負荷試験", "P1", "確定"],
        ["NFR-04", "信頼性", "可用性", "サービス可用性（計画停止除く）", "99.9%/月", "外形監視", "P1", "確定"],
        ["NFR-05", "信頼性", "回復性", "障害時RPO / RTO", "RPO≦5分 / RTO≦1時間", "DR訓練", "P1", "確定"],
        ["NFR-06", "セキュリティ", "機密性", "保存データ暗号化（DB・S3）", "AES-256", "構成監査", "P1", "確定"],
        ["NFR-07", "セキュリティ", "機密性", "通信暗号化", "TLS1.2+", "SSLスキャン", "P1", "確定"],
        ["NFR-08", "セキュリティ", "責任追跡性", "監査ログ保管期間", "≦1年オンライン＋2年アーカイブ", "運用確認", "P1", "確定"],
        ["NFR-09", "セキュリティ", "真正性", "管理者操作はMFA必須", "100%", "IdP設定", "P1", "確定"],
        ["NFR-10", "セキュリティ", "脆弱性", "依存ライブラリの重大脆弱性", "Critical 0（7日以内修正）", "SCA/CI", "P1", "確定"],
        ["NFR-11", "使用性", "習得性", "主要タスク完了までの初回所要", "学習開始まで≦5分", "ユーザテスト", "P2", "確定"],
        ["NFR-12", "使用性", "アクセシビリティ", "JIS X 8341-3", "AA目標", "自動+手動検査", "P2", "確定"],
        ["NFR-13", "保守性", "解析性", "構造化ログ・相関ID付与率", "100%（API）", "ログ監査", "P1", "確定"],
        ["NFR-14", "保守性", "試験性", "ユニットテストカバレッジ（ドメイン）", "≧80%", "CI", "P1", "確定"],
        ["NFR-15", "移植性", "適応性", "ブラウザサポート", "Chrome/Edge/Safari直近2メジャー", "互換試験", "P1", "確定"],
        ["NFR-16", "互換性", "相互運用性", "人事APIとのインターフェース適合", "仕様準拠100%", "結合試験", "P1", "確定"],
        ["NFR-17", "信頼性", "障害許容性", "単一AZ障害時の継続稼働", "自動フェイルオーバー", "カオス/障害試験", "P1", "確定"],
        ["NFR-18", "性能効率性", "資源効率性", "バッチ（人事同期）実行時間", "≦30分/全件2,000ユーザ", "バッチ監視", "P1", "確定"],
        ["NFR-19", "セキュリティ", "機密性", "個人情報の国外転送", "禁止", "アーキレビュー", "P1", "確定"],
        ["NFR-20", "保守性", "モジュール性", "テナント設定変更の無停止反映", "デプロイ不要", "運用確認", "P2", "確定"],
    ]
    for i, t in enumerate(nfrs):
        table_row(ws, 4 + i, t, alt=i % 2 == 1, heights=28)

    # --- 画面帳票 ---
    ws = wb.create_sheet("10_画面帳票一覧")
    set_col_widths(ws, [12, 18, 14, 28, 12, 12, 12, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "9. 画面・帳票一覧")
    section_bar(ws, 3, 8, "9.1 画面一覧")
    table_header(ws, 4, ["画面ID", "画面名", "利用者", "目的", "優先", "認証", "新規/流用", "備考"])
    screens = [
        ["SCR-LOGIN", "ログイン", "全員", "SSO開始・エラー表示", "P1", "不要", "新規", "IdPリダイレクト"],
        ["SCR-HOME-L", "学習者ホーム", "学習者", "次の学習・通知・進捗", "P1", "要", "新規", ""],
        ["SCR-PATH", "学習パス詳細", "学習者", "パス構成と進捗確認", "P1", "要", "新規", ""],
        ["SCR-CONTENT", "教材ビューア", "学習者", "教材学習・進捗送信", "P1", "要", "新規", "動画対応"],
        ["SCR-ASSIGN", "課題提出", "学習者", "課題アップロード", "P1", "要", "新規", ""],
        ["SCR-EXAM", "試験実施", "学習者", "アセスメント受験", "P1", "要", "新規", "フルスクリーン推奨"],
        ["SCR-RESULT", "結果・FB", "学習者", "スコア・コメント確認", "P1", "要", "新規", ""],
        ["SCR-DECK", "語彙デッキ", "学習者", "SRS学習", "P2", "要", "新規", ""],
        ["SCR-HOME-I", "講師ホーム", "講師", "要対応一覧", "P1", "要", "新規", ""],
        ["SCR-GRADE", "採点ワークベンチ", "講師", "提出物採点", "P1", "要", "新規", ""],
        ["SCR-CNT-ADM", "教材管理", "講師/管理者", "コンテンツCRUD・公開", "P1", "要", "新規", ""],
        ["SCR-CLASS", "クラス進捗", "講師", "遅延者モニタ", "P1", "要", "新規", ""],
        ["SCR-USR", "ユーザ管理", "管理者", "招待・無効化・ロール", "P1", "要", "新規", ""],
        ["SCR-RPT", "組織レポート", "管理者", "集計・エクスポート", "P2", "要", "新規", ""],
        ["SCR-AUDIT", "監査ログ", "システム管理者", "ログ検索", "P1", "要", "新規", ""],
        ["SCR-TENANT", "テナント設定", "システム管理者", "ブランド・ポリシー", "P1", "要", "新規", ""],
        ["SCR-BATCH", "バッチ監視", "システム管理者", "同期ジョブ結果", "P1", "要", "新規", ""],
        ["SCR-ERR", "共通エラー", "全員", "404/403/5xx表示", "P1", "—", "新規", ""],
    ]
    for i, t in enumerate(screens):
        table_row(ws, 5 + i, t, alt=i % 2 == 1)

    section_bar(ws, 25, 8, "9.2 帳票・エクスポート一覧")
    table_header(ws, 26, ["帳票ID", "名称", "形式", "利用者", "生成契機", "保持", "個人情報", "備考"])
    reports = [
        ["RPT-01", "個人学習履歴", "PDF/CSV", "学習者/管理者", "オンデマンド", "生成都度", "含む", ""],
        ["RPT-02", "クラス進捗一覧", "CSV/XLSX", "講師", "オンデマンド", "生成都度", "含む", ""],
        ["RPT-03", "組織月次レポート", "PDF", "管理者", "月次自動＋手動", "13ヶ月", "含む", "メール配信可"],
        ["RPT-04", "試験結果明細", "CSV", "講師/管理者", "試験終了後", "試験設定に従う", "含む", ""],
        ["RPT-05", "監査ログエクスポート", "CSV", "システム管理者", "オンデマンド", "保管ポリシー", "含む", "暗号化ZIP"],
        ["RPT-06", "人事同期エラー", "CSV", "システム管理者", "バッチ失敗時", "90日", "含む", ""],
    ]
    for i, t in enumerate(reports):
        table_row(ws, 27 + i, t, alt=i % 2 == 1)

    # --- 外部IF ---
    ws = wb.create_sheet("11_外部IF")
    set_col_widths(ws, [12, 16, 14, 14, 28, 14, 12, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "10. 外部インターフェース要件")
    table_header(ws, 3, ["IF-ID", "名称", "相手先", "方式", "概要", "頻度/契機", "方向", "SLA"])
    ifs = [
        ["IF-01", "人事ユーザ差分", "人事システム", "REST/JSON", "更新ユーザの差分取得", "日次02:00／手動", "受信", "30分以内完了"],
        ["IF-02", "SSO認証", "Entra ID", "OIDC", "認証・属性取得", "ログイン時", "双方向", "IdP SLA準拠"],
        ["IF-03", "メール通知", "AWS SES", "API", "トランザクションメール送信", "イベント時", "送信", "遅延≦5分"],
        ["IF-04", "教材オブジェクト", "Amazon S3", "HTTPS/SDK", "教材バイナリの配置・取得", "随時", "双方向", "可用性99.9%"],
        ["IF-05", "署名付きURL", "CloudFront", "HTTPS", "教材配信（期限付きURL）", "閲覧時", "送信", "p95≦1s発行"],
    ]
    for i, t in enumerate(ifs):
        table_row(ws, 4 + i, t, alt=i % 2 == 1, heights=24)

    section_bar(ws, 11, 8, "10.1 IF-01 人事ユーザ差分　項目（抜粋）")
    table_header(ws, 12, ["項目", "JSONキー", "型", "必須", "説明", "検証", "JLMSマッピング", "備考"])
    if_items = [
        ["社員番号", "employeeId", "string", "○", "一意キー補助", "^[A-Z0-9-]{3,20}$", "users.emp_id", ""],
        ["メール", "email", "string", "○", "ビジネスキー", "RFC5322", "users.email", "小文字正規化"],
        ["氏名（姓）", "lastName", "string", "○", "表示用", "1–80文字", "users.last_name", ""],
        ["氏名（名）", "firstName", "string", "○", "表示用", "1–80文字", "users.first_name", ""],
        ["部署コード", "deptCode", "string", "○", "組織所属", "マスタ存在", "org_units.code", ""],
        ["在籍区分", "status", "enum", "○", "active/leave/retired", "列挙", "users.status", ""],
        ["退職日", "retiredOn", "date", "△", "ISO8601", "日付", "users.retired_on", "retired時必須"],
        ["更新日時", "updatedAt", "datetime", "○", "差分判定", "ISO8601", "—", "カーソル"],
    ]
    for i, t in enumerate(if_items):
        table_row(ws, 13 + i, t, alt=i % 2 == 1)

    # --- データ要件 ---
    ws = wb.create_sheet("12_データ要件")
    set_col_widths(ws, [14, 18, 28, 14, 12, 12, 14, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "11. データ要件")
    section_bar(ws, 3, 8, "11.1 論理エンティティ一覧")
    table_header(ws, 4, ["エンティティ", "説明", "主な属性", "個人情報", "保管期間", "増加量/年", "所有者", "備考"])
    ents = [
        ["Tenant", "契約組織", "name, plan, locale", "含まない", "契約+1年", "十数件", "SysAdmin", ""],
        ["User", "利用者", "email, name, role, status", "含む", "退職+5年", "〜2,000", "Admin", ""],
        ["OrgUnit", "組織", "code, name, parent", "含まない", "存続中", "〜200", "Admin", ""],
        ["ContentPack", "教材包", "title, version, status", "含まない", "最終公開+7年", "〜500", "Instructor", ""],
        ["LearningPath", "学習パス", "level, items[], due", "含まない", "運用中", "〜50", "Instructor", ""],
        ["Enrollment", "受講登録", "user, path, progress", "含む", "修了+5年", "〜4,000", "System", ""],
        ["LearningEvent", "学習イベント", "type, at, payload", "含む", "3年", "百万件級", "System", "分析用"],
        ["Assessment", "試験定義", "title, duration, policy", "含まない", "運用中", "〜100", "Instructor", ""],
        ["Attempt", "受験結果", "score, answers, flags", "含む", "5年", "〜2万", "System", ""],
        ["Notification", "通知", "channel, body, status", "含む", "1年", "〜50万", "System", ""],
        ["AuditLog", "監査ログ", "actor, action, target", "含む", "3年", "〜100万", "SysAdmin", "改ざん防止"],
    ]
    for i, t in enumerate(ents):
        table_row(ws, 5 + i, t, alt=i % 2 == 1, heights=22)

    section_bar(ws, 18, 8, "11.2 個人情報・データ分類")
    table_header(ws, 19, ["分類", "例", "暗号化", "アクセス制御", "マスキング", "持ち出し", "ログ", "備考"])
    classify = [
        ["極秘（個人情報）", "氏名, メール, 学習履歴", "必須", "RBAC+最小権限", "一覧は一部隠す", "原則禁止", "参照記録", "国外保存禁止"],
        ["社外秘", "教材・試験問題", "必須", "ロール別", "—", "承認後", "操作記録", "著作権"],
        ["社内一般", "公開設定、タグ", "推奨", "認証ユーザ", "—", "可", "—", ""],
        ["公開可", "ヘルプ、ログイン画面文言", "任意", "不要可", "—", "可", "—", ""],
    ]
    for i, t in enumerate(classify):
        table_row(ws, 20 + i, t, alt=i % 2 == 1, heights=22)

    # --- 移行運用制約 ---
    ws = wb.create_sheet("13_移行運用制約")
    set_col_widths(ws, [12, 14, 36, 16, 14, 14, 12, 12])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "12. 移行・運用・保守・制約")
    section_bar(ws, 3, 8, "12.1 移行要件")
    table_header(ws, 4, ["移行ID", "対象", "方針", "量", "ツール", "検証", "ロールバック", "担当"])
    mig = [
        ["MIG-01", "受講者マスタ", "人事IF本番同期で初期化", "2,000", "IF-01", "件数・サンプル突合", "無効化", "設計"],
        ["MIG-02", "現行進捗Excel", "CSV取込（任意項目）", "パイロット200", "取込ツール", "学習者確認", "再取込", "BA"],
        ["MIG-03", "教材PDF/音声", "S3へ手動＋メタ登録", "初期100包", "管理画面", "開通確認", "版戻し", "講師"],
        ["MIG-04", "権限・組織", "管理画面で初期設定", "少", "手順書", "チェックリスト", "再設定", "Admin"],
    ]
    for i, t in enumerate(mig):
        table_row(ws, 5 + i, t, alt=i % 2 == 1, heights=22)

    section_bar(ws, 11, 8, "12.2 運用・保守要件（概要）")
    table_header(ws, 12, ["項目", "要件", "目標/水準", "実施者", "時間帯", "エスカレーション", "関連NFR", "備考"])
    ops = [
        ["監視", "外形・APM・ログアラート", "重大5分以内検知", "運用", "24/365", "オンコール→PM", "NFR-04", ""],
        ["バックアップ", "DB継続バックアップ＋S3版", "RPO≦5分", "クラウド", "自動", "障害時DR手順", "NFR-05", ""],
        ["パッチ", "OS/ミドル/ライブラリ", "Critical 7日以内", "開発/運用", "計画窓", "緊急時臨時", "NFR-10", ""],
        ["ヘルプデスク", "問い合わせ一次受付", "営業日10–18時", "発注側", "JST", "ベンダーL2", "—", "別SLA"],
        ["リリース", "Blue/Greenまたはローリング", "月1–2回＋緊急", "開発", "メンテ告知", "承認ワークフロー", "—", ""],
    ]
    for i, t in enumerate(ops):
        table_row(ws, 13 + i, t, alt=i % 2 == 1, heights=22)

    section_bar(ws, 20, 8, "12.3 制約・リスク（残存）")
    table_header(ws, 21, ["ID", "内容", "影響", "確率", "対策", "残存リスク", "オーナー", "状態"])
    risks = [
        ["RSK-01", "人事API仕様変更", "同期失敗", "中", "契約で通知期間／契約試験", "低", "PM", "監視"],
        ["RSK-02", "教材著作権クリア遅延", "公開遅延", "中", "マイルストーン管理", "中", "発注側", "対応中"],
        ["RSK-03", "パイロット期間の要件追加", "スコープ増", "高", "変更管理委員会", "中", "PO", "プロセス"],
        ["RSK-04", "Peak超過（キャンペーン）", "性能劣化", "低", "オートスケール＋負荷試験", "低", "アーキ", "対策済"],
    ]
    for i, t in enumerate(risks):
        table_row(ws, 22 + i, t, alt=i % 2 == 1, heights=22)

    # --- トレーサビリティ ---
    ws = wb.create_sheet("14_トレーサビリティ")
    set_col_widths(ws, [12, 14, 14, 14, 14, 14, 14, 18])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "13. 要件トレーサビリティマトリクス（抜粋）")
    merge_write(ws, "A2", "H2", "課題 → 要件 → ユースケース → 画面 → 試験（詳細設計以降でケースIDを追記）",
                fill=fill_light, font=font_small, align=align_l)
    table_header(ws, 3, ["課題ID", "要件ID", "UC-ID", "画面ID", "IF/帳票", "試験区分", "優先度", "備考"])
    tm = [
        ["IS-01", "FR-USR-01", "UC-10", "SCR-USR", "—", "結合", "P1", ""],
        ["IS-01", "FR-USR-02", "UC-12", "SCR-BATCH", "IF-01", "結合", "P1", ""],
        ["IS-02", "FR-CNT-02", "UC-07", "SCR-CNT-ADM", "IF-04", "結合", "P1", ""],
        ["IS-03", "FR-LRN-01", "UC-03", "SCR-CONTENT", "—", "システム", "P1", ""],
        ["IS-04", "FR-RPT-01", "UC-09", "SCR-CLASS", "RPT-02", "システム", "P1", ""],
        ["IS-05", "FR-ASM-01", "UC-05", "SCR-EXAM", "—", "システム", "P1", ""],
        ["IS-05", "FR-ASM-03", "UC-06", "SCR-RESULT", "RPT-04", "システム", "P1", ""],
        ["IS-06", "FR-RPT-03", "UC-11", "SCR-RPT", "RPT-03", "システム", "P2", ""],
        ["IS-07", "FR-LRN-05", "UC-02", "SCR-HOME-L", "—", "システム", "P3", ""],
        ["—", "FR-AUTH-01", "UC-01", "SCR-LOGIN", "IF-02", "結合", "P1", ""],
        ["—", "FR-AUD-01", "UC-13", "SCR-AUDIT", "RPT-05", "セキュリティ", "P1", ""],
        ["—", "NFR-01", "—", "主要画面", "—", "性能", "P1", "非機能"],
        ["—", "NFR-04", "—", "全体", "—", "信頼性", "P1", "非機能"],
        ["—", "NFR-06", "—", "全体", "—", "セキュリティ", "P1", "非機能"],
    ]
    for i, t in enumerate(tm):
        table_row(ws, 4 + i, t, alt=i % 2 == 1)

    # --- 承認 ---
    ws = wb.create_sheet("15_承認")
    set_col_widths(ws, [14, 18, 18, 14, 14, 18, 14, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "14. 承認記録")
    section_bar(ws, 3, 8, "14.1 レビュー・承認履歴")
    table_header(ws, 4, ["日時", "種別", "参加者", "結果", "指摘件数", "期限", "完了確認", "署名"])
    approvals = [
        ["2026-05-20", "内部レビュー", "設計G、QA", "条件付き合格", "12", "2026-05-27", "完了", "山田"],
        ["2026-06-10", "顧客レビュー", "PO、情シス、講師代表", "条件付き合格", "8", "2026-06-24", "完了", "鈴木"],
        ["2026-07-15", "非機能レビュー", "情シス、アーキ", "合格", "2", "2026-07-22", "完了", "情シス"],
        ["2026-08-01", "正式承認", "Steering委員会", "承認", "0", "—", "完了", "スポンサー"],
    ]
    for i, t in enumerate(approvals):
        table_row(ws, 5 + i, t, alt=i % 2 == 1)

    section_bar(ws, 11, 8, "14.2 ベースライン宣言")
    merge_write(ws, "A12", "H14",
                "本書 版1.0 を JLMS プロジェクトの要件ベースラインとして確定する。\n"
                "以降の追加・変更は変更管理プロセス（RFC起票→影響分析→CCB承認）を経て改訂版として管理する。\n"
                "未決事項は本書発行時点で残存していない（全項目「確定」）。",
                fill=fill_ok, font=font_body, align=align_t, border=thin)

    path = OUT_DIR / "01_要件定義書_JLMS.xlsx"
    wb.save(path)
    return path



def main():
    from basic_design import build_basic_design
    from detailed_design import build_detailed_design

    print("Generating 要件定義書…")
    p1 = build_requirements()
    print(f"  -> {p1}")
    print("Generating 基本設計書…")
    p2 = build_basic_design()
    print(f"  -> {p2}")
    print("Generating 詳細設計書…")
    p3 = build_detailed_design()
    print(f"  -> {p3}")
    print("Done.")


if __name__ == "__main__":
    main()
