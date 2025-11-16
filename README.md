# 成績管理システム Phase2

学校の成績データを管理するためのデスクトップアプリケーション

## 概要

評定、観点別評価、欠課情報などの成績データをExcelファイルから取り込み、データベースで一元管理し、必要に応じて検索・出力できるシステムです。

## 技術スタック

- **UI**: PySide6 (Qt for Python)
- **データベース**: SQLite
- **データ処理**: pandas
- **言語**: Python 3.10+

## プロジェクト構造

```
phase2/
├── main.py                        # エントリーポイント
├── requirements.txt               # 依存パッケージ
│
├── ui/                            # プレゼンテーション層
│   ├── main_window.py            # メインウィンドウ
│   └── dialogs/                  # 各種ダイアログ
│       ├── period_import_dialog.py
│       ├── student_list_dialog.py
│       ├── course_list_dialog.py
│       ├── data_management_dialog.py
│       ├── log_viewer_dialog.py
│       └── column_mapping_dialog.py
│
├── services/                      # ビジネスロジック層
│   ├── data_import_service.py    # データ取り込み
│   ├── data_export_service.py    # データ出力
│   └── data_validation_service.py
│
├── infrastructure/                # インフラ層
│   ├── database_manager.py       # DB操作
│   ├── file_manager.py           # ファイル操作
│   ├── config_manager.py         # 設定管理
│   └── logger.py                 # ログ記録
│
├── models/                        # データモデル層
│   ├── grade.py                  # 評定モデル
│   ├── viewpoint.py              # 観点モデル
│   └── absence.py                # 欠課モデル
│
└── data/                          # データ保存
    ├── database.db               # SQLiteデータベース
    ├── imports/                  # 取り込みファイル
    └── exports/                  # 出力ファイル
```

## セットアップ

```bash
# 1. 依存パッケージのインストール
pip install -r requirements.txt

# 2. アプリケーションの起動
python main.py
```

## 主な機能

### データ取り込み
- Excelファイルから評定、観点別評価、欠課情報を取り込み
- カラムマッピング機能（Excelのカラム名とDB項目の対応付け）
- 複数シート一括取り込み
- データプレビュー機能

### データ管理
- 生徒一覧表示・検索
- 科目一覧表示・検索
- データの編集・削除
- 期間・年度による絞り込み

### データ出力
- Excel形式でのデータ出力
- 条件指定による抽出
- テンプレート機能

### システム機能
- 操作ログの記録・閲覧
- 設定の保存・読み込み
- ファイルバックアップ

## データベース構造

### grades（評定）
- 学籍番号、生徒名
- 科目番号、科目名、教科名
- 期間、年度
- 評定値、単位数、修得単位数

### viewpoint_evaluations（観点別評価）
- 学籍番号、生徒名
- 科目番号、科目名、教科名
- 期間、年度
- 観点1〜5（知識・技能、思考・判断・表現、主体的学習態度等）

### absences（欠課情報）
- 学籍番号、生徒名
- 科目番号、科目名、教科名
- 期間、年度
- 欠席回数、遅刻回数、総時数、欠課率

### activity_logs（操作ログ）
- タイムスタンプ
- 操作種別
- 詳細情報

## ライセンス

MIT License

## 作者

seyaytua
