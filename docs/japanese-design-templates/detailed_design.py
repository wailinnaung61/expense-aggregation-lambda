#!/usr/bin/env python3
"""詳細設計書（内部設計書）生成"""

from openpyxl import Workbook

from generate_design_docs import (
    OUT_DIR, thin, fill_navy, fill_teal, fill_section, fill_light, fill_alt,
    fill_white, fill_warn, fill_ok, font_h2, font_body, font_small, font_label,
    font_num, align_l, align_t, align_c, set_col_widths, merge_write,
    header_bar, section_bar, table_header, table_row, write_cover, write_toc,
)


def build_detailed_design():
    wb = Workbook()
    ws = wb.active
    ws.title = "00_表紙"
    write_cover(ws, "詳細設計書（内部設計書）", "DOC-JLMS-DD-001", "1.0", "正式版（Approved）")

    ws = wb.create_sheet("01_目次")
    toc = [
        ("1", "はじめに・設計方針", "02_設計方針", "—", "設計", "確定", "◎", ""),
        ("2", "プログラム／クラス設計", "03_プログラム設計", "—", "設計", "確定", "◎", ""),
        ("3", "API詳細設計", "04_API詳細", "—", "設計", "確定", "◎", ""),
        ("4", "画面詳細設計", "05_画面詳細", "—", "設計", "確定", "◎", ""),
        ("5", "DB物理設計", "06_DB物理設計", "—", "DBA", "確定", "◎", ""),
        ("6", "テーブル定義（代表）", "07_テーブル定義", "—", "DBA", "確定", "◎", ""),
        ("7", "バッチ詳細設計", "08_バッチ詳細", "—", "設計", "確定", "◎", ""),
        ("8", "エラー・メッセージ設計", "09_エラーメッセージ", "—", "設計", "確定", "◎", ""),
        ("9", "シーケンス・状態遷移", "10_シーケンス状態", "—", "設計", "確定", "◎", ""),
        ("10", "試験項目対応", "11_試験対応", "—", "QA", "確定", "◎", ""),
    ]
    write_toc(ws, toc, "詳細設計書　目次")

    # ── 設計方針 ──
    ws = wb.create_sheet("02_設計方針")
    set_col_widths(ws, [14, 20, 36, 14, 14, 12, 12, 12])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "1. 詳細設計方針")
    section_bar(ws, 3, 8, "1.1 目的とインプット／アウトプット")
    merge_write(
        ws, "A4", "H6",
        "目的：基本設計書（DOC-JLMS-BD-001）を実装可能な粒度（処理ロジック、物理スキーマ、API契約、"
        "画面項目・イベント、エラーコード）まで落とし込む。\n"
        "インプット：要件定義書、基本設計書、OpenAPIスケルトン、コーディング規約。\n"
        "アウトプット：本書、OpenAPI完成版、DDL、メッセージカタログ、UT観点。",
        fill=fill_white, font=font_body, align=align_t, border=thin,
    )
    section_bar(ws, 8, 8, "1.2 パッケージ構成（バックエンド）")
    table_header(ws, 9, ["層", "パッケージ例", "責務", "依存方向", "禁止事項", "テスト", "備考", ""])
    pkgs = [
        ["Interface", "interfaces/http", "Controller, DTO, OpenAPI", "→Application", "DB直接アクセス", "e2e/契約", "", ""],
        ["Application", "application/*", "ユースケース、トランザクション境界", "→Domain/Port", "框架依存最小化", "UT", "", ""],
        ["Domain", "domain/*", "エンティティ、ドメインサービス、仕様", "なし（純正）", "インフラimport", "UT厚め", "", ""],
        ["Port", "application/ports", "リポジトリ／クライアントIF", "Domain型のみ", "具象依存", "モック", "", ""],
        ["Infrastructure", "infrastructure/*", "JPA/Prisma相当、AWS SDK", "→Port実装", "Controllerから直接呼出", "結合", "", ""],
    ]
    for i, t in enumerate(pkgs):
        table_row(ws, 10 + i, t, alt=i % 2 == 1, heights=24)

    section_bar(ws, 17, 8, "1.3 コーディング／設計規約（抜粋）")
    table_header(ws, 18, ["ID", "規約", "内容", "根拠", "例外", "静的検査", "状態", "備考"])
    rules = [
        ["CR-01", "命名", "DBはsnake_case、APIはcamelCase、コードは言語慣例", "一貫性", "外部IF準拠", "lint", "確定", ""],
        ["CR-02", "テナント", "リポジトリ必殺でtenant付与。生SQLはレビュー必須", "SEC-03", "なし", "CodeQL規則", "確定", ""],
        ["CR-03", "日時", "保存UTC、表示はテナントTZ", "NFR", "日付のみ項目", "レビュー", "確定", ""],
        ["CR-04", "ログ", "構造化JSON、秘密情報マスク、requestId必須", "NFR-13", "—", "ログ試験", "確定", ""],
        ["CR-05", "例外", "ドメイン例外→Problem Detailsへ変換", "API共通", "—", "UT", "確定", ""],
        ["CR-06", "冪等", "提出・同期・通知はIdempotency-Key対応", "信頼性", "参照系", "結合", "確定", ""],
    ]
    for i, t in enumerate(rules):
        table_row(ws, 19 + i, t, alt=i % 2 == 1, heights=22)

    # ── プログラム設計 ──
    ws = wb.create_sheet("03_プログラム設計")
    set_col_widths(ws, [18, 16, 28, 18, 14, 14, 12, 12])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "2. プログラム／クラス設計（代表）")
    section_bar(ws, 3, 8, "2.1 主要クラス一覧")
    table_header(ws, 4, ["クラス", "層", "責務", "主要メソッド", "依存", "要件", "複雑度", "備考"])
    classes = [
        ["StartExamHandler", "Application", "試験開始ユースケース", "execute(cmd)", "AttemptRepo, ExamRepo, Clock", "FR-ASM-01", "中", ""],
        ["SaveAnswerHandler", "Application", "回答保存", "execute(cmd)", "AttemptRepo", "FR-ASM-01", "低", "冪等"],
        ["SubmitExamHandler", "Application", "提出・自動採点起動", "execute(cmd)", "AttemptRepo, AutoGrader, EventBus", "FR-ASM-01/03", "高", ""],
        ["AutoGrader", "Domain", "客観問題採点", "grade(attempt)", "AnswerKey", "FR-ASM-03", "中", "純関数寄り"],
        ["RecordLearningEventsHandler", "Application", "イベント受付・進捗更新", "execute(batch)", "EventRepo, ProgressService", "FR-LRN-01", "高", ""],
        ["ProgressService", "Domain", "完了判定・パス進捗再計算", "recalculate(enrollment)", "CompletionPolicy", "FR-LRN-02", "高", ""],
        ["PublishContentHandler", "Application", "版公開", "execute(cmd)", "ContentRepo, Audit", "FR-CNT-02", "中", ""],
        ["HrSyncJob", "Application", "人事同期オーケストレーション", "run(cursor)", "HrClient, UserRepo", "FR-IF-01", "高", "BAT-01"],
        ["RbacAuthorizer", "Application", "認可判定", "authorize(user,action,resource)", "PolicyStore", "FR-AUTH-02", "中", ""],
        ["AuditLogger", "Infrastructure", "監査追記", "append(entry)", "AuditRepo", "FR-AUD-01", "低", ""],
        ["PresignService", "Infrastructure", "署名URL発行", "presignGet(key,ttl)", "S3Client", "IF-04", "低", ""],
        ["NotificationDispatcher", "Application", "通知送信", "dispatch(n)", "SesClient, NotifRepo", "FR-NTF-01", "中", ""],
    ]
    for i, t in enumerate(classes):
        table_row(ws, 5 + i, t, alt=i % 2 == 1, heights=22)

    section_bar(ws, 19, 8, "2.2 StartExamHandler 処理仕様")
    table_header(ws, 20, ["Step", "処理", "入出力", "例外", "トランザクション", "ログ", "監査", "備考"])
    steps = [
        ["1", "入力検証（examId, userId）", "cmd", "VAL_001", "—", "debug", "—", ""],
        ["2", "試験定義取得・公開状態確認", "Assessment", "ASM_404, ASM_409", "—", "info", "—", ""],
        ["3", "受験資格・期間チェック", "Enrollment/Policy", "ASM_403, ASM_422", "—", "info", "—", ""],
        ["4", "既存未提出Attemptの有無確認", "Attempt?", "—", "—", "debug", "—", "再開ポリシー"],
        ["5", "Attempt作成（startedAt, expiresAt）", "Attempt", "CONFLICT", "開始", "info", "exam.start", "UUID v7"],
        ["6", "出題セット確定（シャッフルseed保存）", "questionIds", "—", "継続", "debug", "—", "再現性"],
        ["7", "コミット・DTO返却", "AttemptDTO", "—", "終了", "info", "—", ""],
    ]
    for i, t in enumerate(steps):
        table_row(ws, 21 + i, t, alt=i % 2 == 1, heights=20)

    section_bar(ws, 30, 8, "2.3 ProgressService 完了判定ポリシー")
    table_header(ws, 31, ["教材タイプ", "完了条件", "パラメータ", "再計算契機", "根拠要件", "UT観点", "状態", "備考"])
    policies = [
        ["document", "全必須ページ閲覧", "requiredPageIds", "イベント受信時", "FR-LRN-01", "欠ページ", "確定", ""],
        ["video", "視聴位置≧閾値", "thresholdPct=90", "heartbeat", "FR-LRN-01", "89%/90%", "確定", ""],
        ["quiz", "確認テスト合格", "passScore", "提出時", "FR-LRN-01", "境界点", "確定", ""],
        ["mixed", "子要素すべて完了", "children[]", "子完了時", "FR-LRN-01", "部分完了", "確定", ""],
    ]
    for i, t in enumerate(policies):
        table_row(ws, 32 + i, t, alt=i % 2 == 1)

    # ── API詳細 ──
    ws = wb.create_sheet("04_API詳細")
    set_col_widths(ws, [12, 10, 28, 22, 14, 14, 12, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "3. API詳細設計（代表エンドポイント）")

    def api_block(ws, r, api_id, method, path, summary, auth, req_body, res_body, errors, notes):
        merge_write(ws, f"A{r}", f"H{r}", f"{api_id}  {method}  {path}  — {summary}",
                    fill=fill_section, font=font_h2, align=align_l, border=thin)
        r += 1
        rows = [
            ("認可", auth),
            ("リクエスト", req_body),
            ("レスポンス200", res_body),
            ("エラー", errors),
            ("実装ノート", notes),
        ]
        for label, val in rows:
            merge_write(ws, f"A{r}", f"B{r}", label, fill=fill_alt, font=font_label, align=align_c, border=thin)
            merge_write(ws, f"C{r}", f"H{r}", val, fill=fill_white, font=font_body, align=align_t, border=thin)
            ws.row_dimensions[r].height = 52 if label != "認可" else 22
            r += 1
        return r + 1

    r = 3
    r = api_block(
        ws, r, "API-003", "POST", "/api/v1/exams/{examId}/start", "試験開始",
        "role=Learner かつ 対象試験の受講資格あり",
        "Path: examId(uuid)\nHeaders: Authorization, X-Request-Id, Idempotency-Key(推奨)\nBody: { \"resume\": true }",
        "{ \"attemptId\", \"expiresAt\", \"durationSec\", \"questions\":[{id,type,prompt,choices?}], \"serverTime\" }",
        "401 AUTH_001 / 403 ASM_403 / 404 ASM_404 / 409 ASM_409(期間外・回数超過) / 422 ASM_422",
        "StartExamHandler。再開可の未提出Attemptがあればそれを返却（resume=true時）。",
    )
    r = api_block(
        ws, r, "API-004", "PUT", "/api/v1/exams/attempts/{attemptId}", "回答一時保存",
        "Attempt所有者本人",
        "Path: attemptId\nBody: { \"answers\":[{\"questionId\",\"value\"}], \"clientSavedAt\" }",
        "{ \"attemptId\", \"savedAt\", \"answeredCount\", \"remainingSec\" }",
        "403 / 404 / 409 ASM_LOCKED(提出済・時間切れ) / 422 VAL_001",
        "部分更新。時間切れ判定はサーバ時計。ロック後は409。",
    )
    r = api_block(
        ws, r, "API-005", "POST", "/api/v1/exams/attempts/{attemptId}/submit", "試験提出",
        "Attempt所有者本人",
        "Path: attemptId\nHeaders: Idempotency-Key(必須)\nBody: { \"finalAnswers\":[…] }",
        "{ \"attemptId\", \"status\":\"submitted|auto_graded\", \"score?\", \"availableAt\" }",
        "409 ASM_LOCKED / 409 IDEMPOTENT_REPLAY（同一キー） / 422",
        "SubmitExamHandler→AutoGrader。主観問題のみなら status=submitted。監査 exam.submit。",
    )
    r = api_block(
        ws, r, "API-002", "POST", "/api/v1/events", "学習イベント一括送信",
        "role=Learner",
        "Body: { \"events\":[{\"type\",\"contentVersionId\",\"at\",\"payload\"}] }  max 50",
        "{ \"accepted\", \"rejected\":[{\"index\",\"code\"}], \"progress\":{\"enrollmentId\",\"percent\"} }",
        "400 VAL_001 / 403 / 413（件数超過）",
        "RecordLearningEventsHandler。時計ずれ±10分許容。progress再計算は同期的に返却。",
    )

    section_bar(ws, r, 8, "3.1 ステータスコード対応方針")
    r += 1
    table_header(ws, r, ["HTTP", "意味", "利用場面", "クライアント指針", "ログレベル", "再試行", "例コード", "備考"])
    r += 1
    codes = [
        ["200", "成功", "取得・更新成功", "正常処理", "info", "—", "—", ""],
        ["201", "作成", "資源作成", "Location任意", "info", "—", "—", ""],
        ["400", "構文不正", "JSON不正", "修正して再送", "warn", "不可", "VAL_001", ""],
        ["401", "未認証", "トークン無効", "再ログイン", "info", "不可", "AUTH_001", ""],
        ["403", "権限なし", "越権", "導線変更", "warn", "不可", "AUTH_403", "監査"],
        ["404", "不存在", "ID不正/他テナント", "同一応答で隠蔽可", "info", "不可", "ASM_404", " enumeration対策"],
        ["409", "競合", "状態遷移不正", "最新取得", "info", "条件付き", "ASM_LOCKED", ""],
        ["422", "意味不正", "業務バリデーション", "フィールド提示", "info", "不可", "ASM_422", ""],
        ["429", "制限", "レートリミット", "Backoff", "warn", "可", "RATE_001", ""],
        ["500", "障害", "予期せぬ誤り", "再試行/問い合わせ", "error", "可", "SYS_500", ""],
    ]
    for i, t in enumerate(codes):
        table_row(ws, r + i, t, alt=i % 2 == 1)

    # ── 画面詳細 ──
    ws = wb.create_sheet("05_画面詳細")
    set_col_widths(ws, [10, 14, 12, 12, 28, 14, 14, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "4. 画面詳細設計（SCR-EXAM ほか）")
    section_bar(ws, 3, 8, "4.1 SCR-EXAM イベント設計")
    table_header(ws, 4, ["イベント", "トリガー", "前条件", "API", "成功時UI", "失敗時UI", "二重防止", "備考"])
    evs = [
        ["onEnter", "ルート遷移", "認証済", "API-003", "問題1表示・タイマ開始", "エラー画面/トースト", "Idempotency", "resume対応"],
        ["onSelect", "選択肢クリック", "未ロック", "—（ローカル）", "選択ハイライト", "—", "—", "dirtyフラグ"],
        ["onAutoSave", "30秒/設問移動", "dirty", "API-004", "「保存済」表示", "再試行バナー", "in-flightロック", ""],
        ["onTick", "1秒", "未提出", "—", "残り時間更新", "—", "—", "サーバ補正"],
        ["onTimeout", "残り0", "未提出", "API-005自動", "結果or待機画面", "エラー時再送キュー", "必須", "強制提出"],
        ["onSubmit", "提出ボタン", "確認OK", "API-005", "完了画面", "メッセージ表示", "ボタンdisable", ""],
        ["onLeave", "beforeunload", "受験中", "—", "警告ダイアログ", "—", "—", "ポリシー表示"],
    ]
    for i, t in enumerate(evs):
        table_row(ws, 5 + i, t, alt=i % 2 == 1, heights=24)

    section_bar(ws, 14, 8, "4.2 SCR-EXAM 表示状態")
    table_header(ws, 15, ["状態", "説明", "タイマ", "入力", "提出ボタン", "遷移先", "永続状態", "備考"])
    states = [
        ["Loading", "開始API待ち", "停止", "不可", "不可", "—", "—", ""],
        ["InProgress", "受験中", "動作", "可", "可", "—", "in_progress", ""],
        ["Saving", "保存中", "動作", "可", "不可", "—", "in_progress", ""],
        ["Submitting", "提出中", "停止", "不可", "不可", "—", "in_progress", ""],
        ["Submitted", "提出完了", "停止", "不可", "不可", "SCR-RESULT", "submitted", ""],
        ["Locked", "時間切れロック", "0", "不可", "自動提出中", "SCR-RESULT", "submitted", ""],
        ["Error", "回復不能", "—", "不可", "不可", "SCR-ERR/再試行", "—", ""],
    ]
    for i, t in enumerate(states):
        table_row(ws, 16 + i, t, alt=i % 2 == 1)

    section_bar(ws, 25, 8, "4.3 SCR-GRADE 項目定義（抜粋）")
    table_header(ws, 26, ["項目ID", "項目名", "型", "必須", "編集", "検証", "APIフィールド", "備考"])
    gitems = [
        ["G-01", "提出者", "text", "—", "否", "—", "learnerName", "マスキング設定可"],
        ["G-02", "提出日時", "datetime", "—", "否", "—", "submittedAt", "ローカル表示"],
        ["G-03", "課題タイトル", "text", "—", "否", "—", "assignmentTitle", ""],
        ["G-04", "提出本文/ファイル", "viewer", "—", "否", "—", "payload", "ウイルススキャン済のみ"],
        ["G-05", "得点", "number", "○", "可", "0–maxScore整数", "score", ""],
        ["G-06", "コメント", "textarea", "△", "可", "≦2000", "comment", "差戻し時必須"],
        ["G-07", "判定", "enum", "○", "可", "pass/fail/return", "decision", ""],
        ["G-08", "保存", "button", "—", "可", "—", "API-007", "楽観ロック"],
    ]
    for i, t in enumerate(gitems):
        table_row(ws, 27 + i, t, alt=i % 2 == 1)

    # ── DB物理設計 ──
    ws = wb.create_sheet("06_DB物理設計")
    set_col_widths(ws, [14, 18, 28, 14, 14, 12, 12, 12])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "5. DB物理設計")
    section_bar(ws, 3, 8, "5.1 物理方針")
    table_header(ws, 4, ["項目", "方針", "根拠", "備考", "", "", "", ""])
    phys = [
        ["RDBMS", "Aurora PostgreSQL 15", "BD決定", "", "", "", "", ""],
        ["文字コード", "UTF8", "日本語要件", "照合順序は C.UTF-8 または ICU ja-x-icu", "", "", "", ""],
        ["PK", "UUID (uuid v7) byte/string", "KEY-01", "DBデフォルト生成可", "", "", "", ""],
        ["監査列", "created_at/created_by/updated_at/updated_by/deleted_at/version", "共通", "トリガまたはアプリ付与", "", "", "", ""],
        ["分割", "learning_events は月次パーティション", "量", "3年でデタッチ", "", "", "", ""],
        ["インデックス", "tenant_id複合を原則先頭", "SEC/性能", "下記一覧", "", "", "", ""],
        ["RLS", "Phase1はアプリ強制、RLSはStageで試験導入", "運用", "将来SEC強化", "", "", "", ""],
    ]
    for i, t in enumerate(phys):
        table_row(ws, 5 + i, t, alt=i % 2 == 1, heights=22)

    section_bar(ws, 14, 8, "5.2 インデックス方針（抜粋）")
    table_header(ws, 15, ["テーブル", "インデックス", "列", "種別", "目的", "予想選択性", "保守", "備考"])
    idxs = [
        ["users", "ux_users_tenant_email", "tenant_id,email", "UNIQUE", "ビジネスキー", "高", "—", "削除済除外は部分IDX"],
        ["enrollments", "ix_enroll_user", "tenant_id,user_id", "BTREE", "マイパス", "中", "—", ""],
        ["enrollments", "ix_enroll_path_progress", "tenant_id,path_id,progress_pct", "BTREE", "遅延抽出", "中", "—", ""],
        ["learning_events", "ix_events_enrollment_at", "tenant_id,enrollment_id,at", "BTREE", "履歴", "中", "パーティション局所", ""],
        ["attempts", "ix_attempts_exam_user", "tenant_id,exam_id,user_id", "BTREE", "受験回数", "中", "—", ""],
        ["audit_logs", "ix_audit_at", "tenant_id,at DESC", "BTREE", "検索", "低", "—", ""],
        ["hr_sync_errors", "ix_hr_err_run", "run_id", "BTREE", "エラー一覧", "高", "—", ""],
        ["content_versions", "ux_content_ver", "content_pack_id,version", "UNIQUE", "版一意", "高", "—", ""],
    ]
    for i, t in enumerate(idxs):
        table_row(ws, 16 + i, t, alt=i % 2 == 1)

    # ── テーブル定義 ──
    ws = wb.create_sheet("07_テーブル定義")
    set_col_widths(ws, [6, 22, 12, 10, 10, 18, 14, 20])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "6. テーブル定義（代表3表）")

    def table_def(ws, r, name, desc, cols):
        merge_write(ws, f"A{r}", f"H{r}", f"テーブル: {name}　— {desc}",
                    fill=fill_teal, font=font_h2, align=align_l, border=thin)
        r += 1
        table_header(ws, r, ["#", "列名", "型", "NULL", "デフォルト", "制約/FK", "個人情報", "説明"])
        r += 1
        for i, c in enumerate(cols):
            table_row(ws, r, c, alt=i % 2 == 1)
            r += 1
        return r + 1

    r = 3
    r = table_def(ws, r, "attempts", "試験受験（FR-ASM-*）", [
        ["1", "id", "uuid", "NO", "gen", "PK", "否", "受験ID"],
        ["2", "tenant_id", "uuid", "NO", "—", "FK tenants", "否", "テナント"],
        ["3", "assessment_id", "uuid", "NO", "—", "FK assessments", "否", "試験"],
        ["4", "user_id", "uuid", "NO", "—", "FK users", "含", "受験者"],
        ["5", "status", "text", "NO", "—", "CHK", "否", "in_progress/submitted/graded"],
        ["6", "started_at", "timestamptz", "NO", "now", "—", "否", "開始"],
        ["7", "expires_at", "timestamptz", "NO", "—", "—", "否", "期限"],
        ["8", "submitted_at", "timestamptz", "YES", "—", "—", "否", "提出"],
        ["9", "score", "numeric(6,2)", "YES", "—", "—", "否", "得点"],
        ["10", "max_score", "numeric(6,2)", "NO", "—", "—", "否", "満点"],
        ["11", "shuffle_seed", "bigint", "NO", "—", "—", "否", "出題再現"],
        ["12", "integrity_flags", "jsonb", "NO", "{}", "—", "否", "離脱等フラグ"],
        ["13", "version", "int", "NO", "0", "楽観ロック", "否", ""],
        ["14", "created_at", "timestamptz", "NO", "now", "—", "否", ""],
        ["15", "updated_at", "timestamptz", "NO", "now", "—", "否", ""],
        ["16", "deleted_at", "timestamptz", "YES", "—", "—", "否", "論理削除"],
    ])
    r = table_def(ws, r, "learning_events", "学習イベント（FR-LRN-01）", [
        ["1", "id", "uuid", "NO", "gen", "PK", "否", ""],
        ["2", "tenant_id", "uuid", "NO", "—", "FK", "否", "パーティションキー補助"],
        ["3", "enrollment_id", "uuid", "NO", "—", "FK enrollments", "含", ""],
        ["4", "user_id", "uuid", "NO", "—", "FK users", "含", "冗余（検索用）"],
        ["5", "content_version_id", "uuid", "NO", "—", "FK", "否", ""],
        ["6", "event_type", "text", "NO", "—", "CHK", "否", "start/progress/complete…"],
        ["7", "at", "timestamptz", "NO", "—", "—", "否", "発生時刻"],
        ["8", "payload", "jsonb", "NO", "{}", "—", "条件付", "位置情報等"],
        ["9", "request_id", "text", "YES", "—", "—", "否", "相関"],
        ["10", "created_at", "timestamptz", "NO", "now", "—", "否", "受信時刻"],
    ])
    r = table_def(ws, r, "users", "ユーザ（FR-USR-*）", [
        ["1", "id", "uuid", "NO", "gen", "PK", "否", ""],
        ["2", "tenant_id", "uuid", "NO", "—", "FK", "否", ""],
        ["3", "email", "citext", "NO", "—", "UNIQUE(tenant,email)", "含", "ビジネスキー"],
        ["4", "emp_id", "text", "YES", "—", "—", "含", "社員番号"],
        ["5", "last_name", "text", "NO", "—", "—", "含", ""],
        ["6", "first_name", "text", "NO", "—", "—", "含", ""],
        ["7", "status", "text", "NO", "active", "CHK", "否", "active/leave/retired/disabled"],
        ["8", "retired_on", "date", "YES", "—", "—", "否", ""],
        ["9", "locale", "text", "NO", "ja-JP", "—", "否", ""],
        ["10", "last_login_at", "timestamptz", "YES", "—", "—", "否", ""],
        ["11", "version", "int", "NO", "0", "—", "否", ""],
        ["12", "created_at", "timestamptz", "NO", "now", "—", "否", ""],
        ["13", "updated_at", "timestamptz", "NO", "now", "—", "否", ""],
        ["14", "deleted_at", "timestamptz", "YES", "—", "—", "否", ""],
    ])

    # ── バッチ詳細 ──
    ws = wb.create_sheet("08_バッチ詳細")
    set_col_widths(ws, [8, 18, 28, 14, 14, 14, 12, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "7. バッチ詳細設計（BAT-01 人事差分同期）")
    section_bar(ws, 3, 8, "7.1 基本情報")
    table_header(ws, 4, ["項目", "内容", "", "", "", "", "", ""])
    meta = [
        ["バッチID", "BAT-01", "", "", "", "", "", ""],
        ["実装クラス", "HrSyncJob", "", "", "", "", "", ""],
        ["起動", "EventBridge cron(0 17 * * ? *) ※UTC=JST02:00", "", "", "", "", "", ""],
        ["同時実行", "テナント単位で排他（Redisロック）", "", "", "", "", "", ""],
        ["引数", "tenantId?, fromCursor?, dryRun?", "", "", "", "", "", ""],
        ["成果物", "hr_sync_runs / hr_sync_errors / CloudWatchメトリクス", "", "", "", "", "", ""],
    ]
    for i, t in enumerate(meta):
        table_row(ws, 5 + i, t, alt=i % 2 == 1)

    section_bar(ws, 13, 8, "7.2 処理ステップ詳細")
    table_header(ws, 14, ["Step", "処理", "詳細ロジック", "SQL/IF", "エラー処理", "メトリクス", "補償", "備考"])
    bsteps = [
        ["1", "ロック取得", "lock=hr-sync:{tenant}", "Redis SET NX EX", "取得失敗→スキップ", "lock_fail", "—", ""],
        ["2", "カーソル読込", "直近successのcursor", "SELECT hr_sync_runs", "無ければepoch", "—", "—", ""],
        ["3", "差分取得", "page size 100", "IF-01 GET", "再試行3・失敗終了", "api_latency", "—", ""],
        ["4", "閾値判定", "変更率>20%で停止", "計算", "SEV1アラート", "threshold_block", "手動承認後再実行", ""],
        ["5", "バリデーション", "必須・形式・部署存在", "メモリ+マスタ", "行エラーへ", "row_error", "—", ""],
        ["6", "適用", "UPSERT/無効化 チャンク50", "INSERT…ON CONFLICT", "チャンク失敗記録", "applied", "チャンク単位", ""],
        ["7", "ロールは触らない", "管理者ロール自動付与禁止", "—", "—", "—", "—", "BR3"],
        ["8", "サマリ確定", "runレコード success/failed", "UPDATE", "—", "duration", "ロック解放", "finally"],
        ["9", "通知", "失敗時SES/Slack", "IF-03", "通知失敗はログのみ", "notify", "—", ""],
    ]
    for i, t in enumerate(bsteps):
        table_row(ws, 15 + i, t, alt=i % 2 == 1, heights=28)

    section_bar(ws, 26, 8, "7.3 リラン手順（運用）")
    merge_write(
        ws, "A27", "H29",
        "1) 管理画面またはCLIで失敗run_idを確認 → 2) hr_sync_errors を是正（相手側 or マッピング）\n"
        "3) dryRun=true で対象件数確認 → 4) fromCursorを失敗時点に固定して本実行\n"
        "5) 成功確認後、学習者サンプルでログイン可否をスポットチェック",
        fill=fill_warn, font=font_body, align=align_t, border=thin,
    )

    # ── エラーメッセージ ──
    ws = wb.create_sheet("09_エラーメッセージ")
    set_col_widths(ws, [12, 10, 18, 28, 18, 14, 12, 12])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "8. エラー・メッセージ設計")
    section_bar(ws, 3, 8, "8.1 メッセージカタログ（抜粋）")
    table_header(ws, 4, ["コード", "HTTP", "種別", "メッセージ（ja）", "メッセージ（en）", "表示箇所", "ログ", "監査"])
    msgs = [
        ["AUTH_001", "401", "認証", "セッションが無効です。再度ログインしてください。", "Your session is invalid. Please sign in again.", "トースト/全面", "info", "任意"],
        ["AUTH_403", "403", "認可", "この操作を行う権限がありません。", "You do not have permission for this action.", "全面", "warn", "○"],
        ["VAL_001", "400", "検証", "入力内容を確認してください。", "Please check your input.", "フィールド", "info", "—"],
        ["ASM_404", "404", "業務", "試験が見つかりません。", "Exam not found.", "全面", "info", "—"],
        ["ASM_403", "403", "業務", "この試験を受験する権限がありません。", "You are not allowed to take this exam.", "全面", "warn", "○"],
        ["ASM_409", "409", "業務", "現在この試験は受験できません（期間または回数）。", "This exam is not available now.", "ダイアログ", "info", "—"],
        ["ASM_422", "422", "業務", "受験条件を満たしていません。", "Exam prerequisites are not met.", "ダイアログ", "info", "—"],
        ["ASM_LOCKED", "409", "業務", "提出済みまたは時間切れのため編集できません。", "Attempt is locked.", "トースト", "info", "—"],
        ["CNT_409", "409", "業務", "この版は公開できない状態です。", "This version cannot be published.", "ダイアログ", "info", "○"],
        ["IF_HR_001", "500", "外部", "人事システム連携に失敗しました。時間をおいて再実行します。", "HR sync failed.", "運用", "error", "○"],
        ["RATE_001", "429", "制限", "リクエストが集中しています。しばらくしてから再試行してください。", "Too many requests.", "トースト", "warn", "—"],
        ["SYS_500", "500", "システム", "システムエラーが発生しました。問い合わせ番号を控えてください。", "A system error occurred.", "全面", "error", "○"],
    ]
    for i, t in enumerate(msgs):
        table_row(ws, 5 + i, t, alt=i % 2 == 1, heights=28)

    section_bar(ws, 19, 8, "8.2 例外マッピング方針")
    merge_write(
        ws, "A20", "H22",
        "Domain例外は code を必須で持ち、HTTP層で Problem Details に変換する。"
        "detail には内部情報（SQL、スタック）を出さない。問い合わせ番号として requestId を表示する。\n"
        "フィールドエラーは errors[{field,code,message}] 配列で返す。",
        fill=fill_white, font=font_body, align=align_t, border=thin,
    )

    # ── シーケンス・状態 ──
    ws = wb.create_sheet("10_シーケンス状態")
    set_col_widths(ws, [10, 16, 16, 16, 16, 16, 14, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "9. シーケンス・状態遷移")
    section_bar(ws, 3, 8, "9.1 試験受験シーケンス（論理）")
    merge_write(
        ws, "A4", "H9",
        "Learner UI → API Gateway/CloudFront → ExamController → StartExamHandler → AttemptRepo/DB\n"
        "　└ 以降、UIは定期的に SaveAnswerHandler へ PUT\n"
        "Learner UI → SubmitExamHandler →（TX）Attemptロック → AutoGrader → EventBus\n"
        "　├ 客観のみ: status=graded, score確定 → 結果APIで閲覧\n"
        "　└ 主観あり: status=submitted → 講師 Grading → score確定 → NotificationDispatcher → SES\n"
        "監視: 各ハンドラで requestId 付き構造化ログ。提出は監査 exam.submit。",
        fill=fill_light, font=font_body, align=align_t, border=thin,
    )
    section_bar(ws, 11, 8, "9.2 Attempt 状態遷移")
    table_header(ws, 12, ["From", "Event", "To", "ガード", "アクション", "通知", "監査", "備考"])
    sm = [
        ["(none)", "start", "in_progress", "資格・期間OK", "Attempt作成", "—", "exam.start", ""],
        ["in_progress", "save", "in_progress", "未期限・未提出", "answers upsert", "—", "—", ""],
        ["in_progress", "submit", "submitted", "未提出", "ロック、自動採点起動", "条件付", "exam.submit", ""],
        ["in_progress", "timeout", "submitted", "expires_at超過", "強制提出", "—", "exam.timeout", ""],
        ["submitted", "auto_grade_done", "graded", "客観のみ/混在で客観確定", "score更新", "結果公開設定", "exam.graded", ""],
        ["submitted", "manual_grade", "graded", "講師確定", "score/comment", "学習者へ", "exam.manual_grade", ""],
        ["submitted", "return", "in_progress", "再提出許可", "ロック解除", "学習者へ", "exam.return", "課題系"],
        ["graded", "—", "graded", "終端", "—", "—", "—", "変更は訂正流程"],
    ]
    for i, t in enumerate(sm):
        table_row(ws, 13 + i, t, alt=i % 2 == 1, heights=22)

    section_bar(ws, 23, 8, "9.3 ContentVersion 状態遷移")
    table_header(ws, 24, ["From", "Event", "To", "ガード", "アクション", "通知", "監査", "備考"])
    sm2 = [
        ["draft", "save", "draft", "編集権限", "メタ更新", "—", "—", ""],
        ["draft", "submit_review", "in_review", "必須項目齐全", "レビュー依頼", "レビュア", "content.submit", "任意WF"],
        ["draft/in_review", "publish", "published", "承認済/権限", "公開時刻設定、CDN", "担当講師", "content.publish", ""],
        ["published", "unpublish", "archived", "権限", "配信停止", "—", "content.unpublish", ""],
        ["published", "new_version", "draft(新)", "—", "版番号+1のdraft作成", "—", "content.fork", "旧版は維持"],
    ]
    for i, t in enumerate(sm2):
        table_row(ws, 25 + i, t, alt=i % 2 == 1, heights=22)

    # ── 試験対応 ──
    ws = wb.create_sheet("11_試験対応")
    set_col_widths(ws, [12, 14, 14, 28, 14, 12, 12, 14])
    ws.sheet_view.showGridLines = False
    header_bar(ws, 1, 8, "10. 試験項目対応（詳細設計→テスト）")
    table_header(ws, 3, ["設計ID", "要件ID", "試験レベル", "観点", "期待結果", "優先", "自動化", "ケースID案"])
    tests = [
        ["API-003", "FR-ASM-01", "結合", "期間外開始", "409 ASM_409", "P1", "○", "TC-ASM-001"],
        ["API-005", "FR-ASM-01", "結合", "二重提出", "同一結果・冪等", "P1", "○", "TC-ASM-010"],
        ["API-005", "FR-ASM-03", "結合", "四択自動採点", "score一致", "P1", "○", "TC-ASM-020"],
        ["API-002", "FR-LRN-01", "結合", "90%視聴完了", "progress=100", "P1", "○", "TC-LRN-005"],
        ["ProgressService", "FR-LRN-01", "UT", "閾値境界89/90", "未完了/完了", "P1", "○", "TC-LRN-UT-01"],
        ["BAT-01", "FR-IF-01", "結合", "閾値超過停止", "適用0・アラート", "P1", "○", "TC-IF-003"],
        ["BAT-01", "FR-IF-01", "結合", "退職者無効化", "ログイン不可", "P1", "○", "TC-IF-007"],
        ["SEC-03", "FR-SYS-01", "セキュリティ", "他テナントID指定", "404/403・漏洩なし", "P1", "○", "TC-SEC-010"],
        ["SEC-01", "NFR-09", "セキュリティ", "管理者MFA無し", "アクセス拒否", "P1", "△", "TC-SEC-001"],
        ["SCR-EXAM", "FR-ASM-01", "システム", "タイムアウト強制提出", "Locked→結果", "P1", "△", "TC-UI-EX-09"],
        ["NFR-01", "NFR-01", "性能", "ホームp95", "≦2.0s @500ses", "P1", "○", "TC-PERF-001"],
        ["RPT-03", "FR-RPT-03", "システム", "月次PDF生成", "レイアウト/値妥当", "P2", "△", "TC-RPT-003"],
        ["API-007", "FR-ASM-02", "結合", "採点楽観ロック衝突", "409", "P1", "○", "TC-GRD-004"],
        ["AuditLogger", "FR-AUD-01", "結合", "権限変更記録", "監査検索でヒット", "P1", "○", "TC-AUD-002"],
    ]
    for i, t in enumerate(tests):
        table_row(ws, 4 + i, t, alt=i % 2 == 1, heights=22)

    section_bar(ws, 20, 8, "10.1 詳細設計ベースライン")
    merge_write(
        ws, "A21", "H23",
        "本書 版1.0 を実装・単体／結合試験のベースラインとする。API契約は OpenAPI 同梱版と一致させる。"
        "スキーマ変更はマイグレーション番号を本書改訂履歴に追記すること。",
        fill=fill_ok, font=font_body, align=align_t, border=thin,
    )

    path = OUT_DIR / "03_詳細設計書_JLMS.xlsx"
    wb.save(path)
    return path
