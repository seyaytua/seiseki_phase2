from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QTabWidget, QTableWidget, QMenuBar,
                               QMenu, QMessageBox, QLabel, QStatusBar, QTableWidgetItem,
                               QSpinBox, QCheckBox, QComboBox, QHeaderView, QGroupBox,
                               QFrame)
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QFont
import json
from pathlib import Path
from datetime import datetime


class MainWindow(QMainWindow):
    """評価・評定入力アプリ メインウィンドウ（ワークフロー型）"""
    
    def __init__(self, db_manager, config_manager, file_manager, data_importer, logger):
        super().__init__()
        self.db_manager = db_manager
        self.config_manager = config_manager
        self.file_manager = file_manager
        self.data_importer = data_importer
        self.logger = logger
        
        self.load_settings()
        self.setup_ui()
    
    def load_settings(self):
        """設定ファイル読み込み"""
        settings_path = Path('config/settings.json')
        with open(settings_path, 'r', encoding='utf-8-sig') as f:
            self.settings = json.load(f)
    
    def setup_ui(self):
        """UI初期化"""
        self.setWindowTitle(self.settings['app_name'])
        self.setGeometry(100, 50, 1400, 900)
        
        # メニューバー
        self.create_menu_bar()
        
        # 中央ウィジェット
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout(central_widget)
        
        # タイトル
        title_label = QLabel("評価・評定データ処理ワークフロー")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # ワークフローエリア
        workflow_group = self.create_workflow_area()
        layout.addWidget(workflow_group)
        
        # 区切り線
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line)
        
        # フィルタエリア
        filter_layout = QHBoxLayout()
        
        filter_layout.addWidget(QLabel("年度:"))
        self.year_filter = QSpinBox()
        self.year_filter.setMinimum(2000)
        self.year_filter.setMaximum(2100)
        self.year_filter.setValue(datetime.now().year)
        self.year_filter.valueChanged.connect(self.refresh_current_tab)
        filter_layout.addWidget(self.year_filter)
        
        filter_layout.addWidget(QLabel("期間:"))
        self.period_filter = QComboBox()
        self.period_filter.addItem("全て")
        self.period_filter.addItems(self.settings['periods'])
        self.period_filter.currentTextChanged.connect(self.refresh_current_tab)
        filter_layout.addWidget(self.period_filter)
        
        filter_layout.addWidget(QLabel("表示件数:"))
        self.limit_spin = QSpinBox()
        self.limit_spin.setMinimum(10)
        self.limit_spin.setMaximum(10000)
        self.limit_spin.setValue(100)
        self.limit_spin.setSingleStep(100)
        filter_layout.addWidget(self.limit_spin)
        
        refresh_btn = QPushButton("データ更新")
        refresh_btn.clicked.connect(self.refresh_current_tab)
        filter_layout.addWidget(refresh_btn)
        
        export_btn = QPushButton("Excel出力")
        export_btn.clicked.connect(self.export_current_data)
        filter_layout.addWidget(export_btn)
        
        filter_layout.addStretch()
        layout.addLayout(filter_layout)
        
        # タブウィジェット
        self.tab_widget = QTabWidget()
        self.tables = {}
        
        # 各データタイプのタブを作成
        tab_order = ['評定', '観点', '欠課情報']
        for data_type in tab_order:
            table = QTableWidget()
            self.tables[data_type] = table
            self.tab_widget.addTab(table, data_type)
        
        # タブ変更時にデータを読み込む
        self.tab_widget.currentChanged.connect(self.on_tab_changed)
        
        layout.addWidget(self.tab_widget)
        
        # ステータスバー
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("準備完了 - ワークフローに従って処理を進めてください")
    
    def create_workflow_area(self):
        """ワークフローエリア作成"""
        workflow_group = QGroupBox("作業フロー")
        workflow_layout = QVBoxLayout()
        
        # ステップ1: 評価・評定取り込み
        step1_layout = QHBoxLayout()
        step1_label = QLabel("【ステップ1】")
        step1_label.setStyleSheet("font-weight: bold; color: #2E86AB;")
        step1_layout.addWidget(step1_label)
        
        step1_btn = QPushButton("評価・評定データ取り込み")
        step1_btn.setMinimumHeight(40)
        step1_btn.setStyleSheet("background-color: #A7C7E7; font-weight: bold;")
        step1_btn.clicked.connect(lambda: self.open_import_dialog('評定'))
        step1_layout.addWidget(step1_btn)
        
        step1_desc = QLabel("→ 評定データ（1-10の評価値）を取り込みます")
        step1_layout.addWidget(step1_desc)
        step1_layout.addStretch()
        
        workflow_layout.addLayout(step1_layout)
        
        # ステップ2: 観点評価取り込み
        step2_layout = QHBoxLayout()
        step2_label = QLabel("【ステップ2】")
        step2_label.setStyleSheet("font-weight: bold; color: #2E86AB;")
        step2_layout.addWidget(step2_label)
        
        step2_btn = QPushButton("観点評価データ取り込み")
        step2_btn.setMinimumHeight(40)
        step2_btn.setStyleSheet("background-color: #A7C7E7; font-weight: bold;")
        step2_btn.clicked.connect(lambda: self.open_import_dialog('観点'))
        step2_layout.addWidget(step2_btn)
        
        step2_desc = QLabel("→ 観点評価（観点1-5）を取り込みます")
        step2_layout.addWidget(step2_desc)
        step2_layout.addStretch()
        
        workflow_layout.addLayout(step2_layout)
        
        # ステップ3: 未入力者チェック
        step3_layout = QHBoxLayout()
        step3_label = QLabel("【ステップ3】")
        step3_label.setStyleSheet("font-weight: bold; color: #FF6B35;")
        step3_layout.addWidget(step3_label)
        
        step3_btn = QPushButton("未入力者チェック")
        step3_btn.setMinimumHeight(40)
        step3_btn.setStyleSheet("background-color: #FFD23F; font-weight: bold;")
        step3_btn.clicked.connect(self.check_missing_entries)
        step3_layout.addWidget(step3_btn)
        
        step3_desc = QLabel("→ 評価・観点の未入力者を確認します")
        step3_layout.addWidget(step3_desc)
        step3_layout.addStretch()
        
        workflow_layout.addLayout(step3_layout)
        
        # ステップ4: 欠課データ前処理
        step4_layout = QHBoxLayout()
        step4_label = QLabel("【ステップ4】")
        step4_label.setStyleSheet("font-weight: bold; color: #C1292E;")
        step4_layout.addWidget(step4_label)
        
        step4_btn = QPushButton("欠課データ前処理")
        step4_btn.setMinimumHeight(40)
        step4_btn.setStyleSheet("background-color: #F4A261; font-weight: bold;")
        step4_btn.clicked.connect(self.open_absence_preprocessor)
        step4_layout.addWidget(step4_btn)
        
        step4_desc = QLabel("→ 複数ファイルから欠課データを抽出・集計します")
        step4_layout.addWidget(step4_desc)
        step4_layout.addStretch()
        
        workflow_layout.addLayout(step4_layout)
        
        # ステップ5: 欠課情報取り込み
        step5_layout = QHBoxLayout()
        step5_label = QLabel("【ステップ5】")
        step5_label.setStyleSheet("font-weight: bold; color: #2E86AB;")
        step5_layout.addWidget(step5_label)
        
        step5_btn = QPushButton("欠課情報取り込み")
        step5_btn.setMinimumHeight(40)
        step5_btn.setStyleSheet("background-color: #A7C7E7; font-weight: bold;")
        step5_btn.clicked.connect(lambda: self.open_import_dialog('欠課情報'))
        step5_layout.addWidget(step5_btn)
        
        step5_desc = QLabel("→ 前処理済みの欠課データを取り込みます")
        step5_layout.addWidget(step5_desc)
        step5_layout.addStretch()
        
        workflow_layout.addLayout(step5_layout)
        
        workflow_group.setLayout(workflow_layout)
        return workflow_group
    
    def create_menu_bar(self):
        """メニューバー作成"""
        menubar = self.menuBar()
        
        # ファイルメニュー
        file_menu = menubar.addMenu("ファイル(&F)")
        
        # データベース選択
        select_db_action = QAction("データベース選択(&D)", self)
        select_db_action.triggered.connect(self.select_database)
        file_menu.addAction(select_db_action)
        
        file_menu.addSeparator()
        
        # データ更新
        refresh_action = QAction("データ更新(&R)", self)
        refresh_action.triggered.connect(self.refresh_current_tab)
        file_menu.addAction(refresh_action)
        
        file_menu.addSeparator()
        
        # Excel出力
        export_current_action = QAction("現在のデータをExcel出力(&E)", self)
        export_current_action.triggered.connect(self.export_current_data)
        file_menu.addAction(export_current_action)
        
        export_all_action = QAction("全データをExcel出力(&A)", self)
        export_all_action.triggered.connect(self.export_all_data)
        file_menu.addAction(export_all_action)
        
        file_menu.addSeparator()
        
        # データ削除
        clear_data_action = QAction("現在のデータを削除(&C)", self)
        clear_data_action.triggered.connect(self.clear_current_data)
        file_menu.addAction(clear_data_action)
        
        file_menu.addSeparator()
        
        # 終了
        exit_action = QAction("終了(&X)", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # ワークフローメニュー
        workflow_menu = menubar.addMenu("ワークフロー(&W)")
        
        step1_action = QAction("ステップ1: 評価・評定取り込み(&1)", self)
        step1_action.triggered.connect(lambda: self.open_import_dialog('評定'))
        workflow_menu.addAction(step1_action)
        
        step2_action = QAction("ステップ2: 観点評価取り込み(&2)", self)
        step2_action.triggered.connect(lambda: self.open_import_dialog('観点'))
        workflow_menu.addAction(step2_action)
        
        step3_action = QAction("ステップ3: 未入力者チェック(&3)", self)
        step3_action.triggered.connect(self.check_missing_entries)
        workflow_menu.addAction(step3_action)
        
        step4_action = QAction("ステップ4: 欠課データ前処理(&4)", self)
        step4_action.triggered.connect(self.open_absence_preprocessor)
        workflow_menu.addAction(step4_action)
        
        step5_action = QAction("ステップ5: 欠課情報取り込み(&5)", self)
        step5_action.triggered.connect(lambda: self.open_import_dialog('欠課情報'))
        workflow_menu.addAction(step5_action)
        
        # 設定メニュー
        settings_menu = menubar.addMenu("設定(&S)")
        
        # 必須カラム管理
        required_cols_action = QAction("必須カラム管理(&R)", self)
        required_cols_action.triggered.connect(self.open_required_columns_manager)
        settings_menu.addAction(required_cols_action)
        
        settings_menu.addSeparator()
        
        # DBカラム設定サブメニュー
        db_cols_menu = settings_menu.addMenu("DBカラム設定(&D)")
        
        for data_type in ['評定', '観点', '欠課情報']:
            action = QAction(data_type, self)
            action.triggered.connect(lambda checked, dt=data_type: self.open_db_column_manager(dt))
            db_cols_menu.addAction(action)
        
        settings_menu.addSeparator()
        
        # プリセット管理サブメニュー
        preset_menu = settings_menu.addMenu("プリセット管理(&P)")
        
        for data_type in ['評定', '観点', '欠課情報']:
            action = QAction(data_type, self)
            action.triggered.connect(lambda checked, dt=data_type: self.open_preset_manager(dt))
            preset_menu.addAction(action)
        
        # ヘルプメニュー
        help_menu = menubar.addMenu("ヘルプ(&H)")
        
        workflow_guide_action = QAction("ワークフローガイド(&W)", self)
        workflow_guide_action.triggered.connect(self.show_workflow_guide)
        help_menu.addAction(workflow_guide_action)
        
        about_action = QAction("バージョン情報(&A)", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
    
    def select_database(self):
        """データベース選択ダイアログを開く"""
        from ui.database_selector_dialog import DatabaseSelectorDialog
        
        current_db_path = self.settings['database']['path']
        
        dialog = DatabaseSelectorDialog(current_db_path, self)
        if dialog.exec():
            new_db_path = dialog.get_selected_path()
            
            if new_db_path != current_db_path:
                # 設定を更新
                self.settings['database']['path'] = new_db_path
                
                # 設定ファイルに保存
                settings_path = Path('config/settings.json')
                with open(settings_path, 'w', encoding='utf-8') as f:
                    json.dump(self.settings, f, ensure_ascii=False, indent=2)
                
                # データベースマネージャーを再初期化
                try:
                    self.db_manager.close()
                except:
                    pass
                
                from database.db_manager import DatabaseManager
                self.db_manager = DatabaseManager(new_db_path)
                
                # データ更新
                self.refresh_current_tab()
                
                QMessageBox.information(
                    self,
                    "データベース変更",
                    f"データベースを変更しました:\n{new_db_path}"
                )
    
    def open_import_dialog(self, data_type):
        """取り込みダイアログを開く"""
        try:
            from ui.period_import_dialog import PeriodImportDialog
            
            dialog = PeriodImportDialog(
                data_type,
                self.db_manager,
                self.config_manager,
                self.file_manager,
                self.data_importer,
                self
            )
            
            if dialog.exec():
                self.status_bar.showMessage(f"{data_type}の取り込みが完了しました", 5000)
                self.refresh_current_tab()
                
                # 次のステップを提案
                self.suggest_next_step(data_type)
        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            print(error_detail)
            QMessageBox.critical(
                self,
                "エラー",
                f"取り込みダイアログの表示に失敗しました:\n{str(e)}"
            )
    
    def suggest_next_step(self, completed_step):
        """次のステップを提案"""
        next_steps = {
            '評定': ('ステップ2', '観点評価データの取り込みを行いますか？'),
            '観点': ('ステップ3', '未入力者チェックを行いますか？'),
            '欠課情報': (None, '全ての取り込みが完了しました！')
        }
        
        if completed_step in next_steps:
            step_name, message = next_steps[completed_step]
            
            if step_name:
                reply = QMessageBox.question(
                    self,
                    "次のステップ",
                    f"{completed_step}の取り込みが完了しました。\n\n"
                    f"{step_name}に進みますか？\n{message}",
                    QMessageBox.Yes | QMessageBox.No
                )
                
                if reply == QMessageBox.Yes:
                    if completed_step == '評定':
                        self.open_import_dialog('観点')
                    elif completed_step == '観点':
                        self.check_missing_entries()
            else:
                QMessageBox.information(
                    self,
                    "完了",
                    message
                )
    
    def check_missing_entries(self):
        """未入力者チェック"""
        try:
            from ui.missing_entry_checker_dialog import MissingEntryCheckerDialog
            
            dialog = MissingEntryCheckerDialog(self.db_manager, self)
            dialog.exec()
            
            # チェック後、欠課データ前処理を提案
            reply = QMessageBox.question(
                self,
                "次のステップ",
                "未入力者チェックが完了しました。\n\n"
                "ステップ4に進みますか?\n"
                "欠課データの前処理を行います。",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.open_absence_preprocessor()
        
        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            print(error_detail)
            QMessageBox.critical(
                self,
                "エラー",
                f"未入力者チェックに失敗:\n{str(e)}"
            )
    
    def open_absence_preprocessor(self):
        """欠課データ前処理ダイアログを開く"""
        try:
            from ui.absence_preprocessor_dialog import AbsencePreprocessorDialog
            
            dialog = AbsencePreprocessorDialog(self.config_manager, self)
            if dialog.exec():
                # 前処理完了後、取り込みを提案
                reply = QMessageBox.question(
                    self,
                    "次のステップ",
                    "欠課データの前処理が完了しました。\n\n"
                    "ステップ5に進みますか？\n"
                    "欠課情報の取り込みを行います。",
                    QMessageBox.Yes | QMessageBox.No
                )
                
                if reply == QMessageBox.Yes:
                    self.open_import_dialog('欠課情報')
        
        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            print(error_detail)
            QMessageBox.critical(
                self,
                "エラー",
                f"欠課データ前処理に失敗:\n{str(e)}"
            )
    
    def on_tab_changed(self, index):
        """タブ変更時の処理"""
        self.refresh_current_tab()
    
    def refresh_current_tab(self):
        """現在のタブを更新"""
        current_index = self.tab_widget.currentIndex()
        if current_index < 0:
            return
        
        data_types = ['評定', '観点', '欠課情報']
        data_type = data_types[current_index]
        table = self.tables[data_type]
        
        self.status_bar.showMessage("データ読み込み中...")
        
        self.load_data_to_table(data_type, table)
    
    def load_data_to_table(self, data_type, table):
        """データベースからデータをテーブルに読み込む"""
        table_mapping = {
            '評定': 'grades',
            '観点': 'viewpoint_evaluations',
            '欠課情報': 'absences'
        }
        
        table_name = table_mapping.get(data_type)
        if not table_name:
            self.status_bar.showMessage(f"{data_type}: 未対応のデータタイプです")
            return
        
        try:
            # フィルタ条件
            year = self.year_filter.value()
            period = self.period_filter.currentText()
            limit = self.limit_spin.value()
            
            # WHERE句作成
            where_clauses = [f"year = {year}"]
            if period != "全て":
                where_clauses.append(f"period = '{period}'")
            
            where_str = " AND ".join(where_clauses)
            
            # 件数取得
            count_query = f"SELECT COUNT(*) as count FROM {table_name} WHERE {where_str}"
            count_result = self.db_manager.fetch_one(count_query)
            total_count = count_result['count'] if count_result else 0
            
            if total_count == 0:
                table.clear()
                table.setRowCount(0)
                table.setColumnCount(0)
                self.status_bar.showMessage(f"{data_type}: データがありません（年度: {year}, 期間: {period}）")
                return
            
            # データ取得
            query = f"SELECT * FROM {table_name} WHERE {where_str} LIMIT {limit}"
            rows = self.db_manager.fetch_all(query)
            
            if not rows:
                table.clear()
                table.setRowCount(0)
                table.setColumnCount(0)
                self.status_bar.showMessage(f"{data_type}: データがありません")
                return
            
            # テーブルに表示
            table.clear()
            table.setRowCount(len(rows))
            table.setColumnCount(len(rows[0].keys()))
            table.setHorizontalHeaderLabels(rows[0].keys())
            
            for i, row in enumerate(rows):
                for j, key in enumerate(row.keys()):
                    value = row[key]
                    item = QTableWidgetItem(str(value) if value is not None else '')
                    table.setItem(i, j, item)
            
            table.resizeColumnsToContents()
            
            status_msg = f"{data_type}: {len(rows)}件表示中（全{total_count}件） | 年度: {year}, 期間: {period}"
            self.status_bar.showMessage(status_msg)
            
        except Exception as e:
            self.status_bar.showMessage(f"エラー: {str(e)}")
            QMessageBox.warning(self, "エラー", f"データ読み込みエラー:\n{str(e)}")
    
    def export_current_data(self):
        """現在のデータをExcel出力"""
        current_index = self.tab_widget.currentIndex()
        if current_index < 0:
            QMessageBox.warning(self, "警告", "データタブを選択してください")
            return
        
        data_types = ['評定', '観点', '欠課情報']
        data_type = data_types[current_index]
        
        table_mapping = {
            '評定': 'grades',
            '観点': 'viewpoint_evaluations',
            '欠課情報': 'absences'
        }
        
        table_name = table_mapping.get(data_type)
        
        try:
            year = self.year_filter.value()
            period = self.period_filter.currentText()
            
            where_clauses = [f"year = {year}"]
            if period != "全て":
                where_clauses.append(f"period = '{period}'")
            
            where_str = " AND ".join(where_clauses)
            
            query = f"SELECT * FROM {table_name} WHERE {where_str}"
            rows = self.db_manager.fetch_all(query)
            
            if not rows or len(rows) == 0:
                QMessageBox.information(self, "情報", f"{data_type}にはデータがありません")
                return
            
            data = [dict(row) for row in rows]
            columns = list(data[0].keys())
            
            from utils.excel_exporter import ExcelExporter
            exporter = ExcelExporter()
            
            export_path = exporter.export_to_excel(
                data=data,
                columns=columns,
                filename=f"{data_type}_{year}_{period}.xlsx",
                sheet_name=data_type
            )
            
            if self.logger:
                self.logger.log_action(
                    action_type='export',
                    master_type=data_type,
                    file_path=export_path,
                    record_count=len(data),
                    status='success',
                    details=f'Excel出力: 年度{year}, 期間{period}'
                )
            
            reply = QMessageBox.information(
                self,
                "出力完了",
                f"{data_type}をExcelファイルに出力しました。\n\n"
                f"ファイル: {Path(export_path).name}\n"
                f"レコード数: {len(data)}件\n\n"
                f"出力先フォルダを開きますか？",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                import os
                export_dir = Path(export_path).parent
                
                if os.name == 'nt':
                    os.startfile(export_dir)
        
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"Excel出力に失敗:\n{str(e)}")
    
    def export_all_data(self):
        """全データをExcel出力"""
        table_mapping = {
            '評定': 'grades',
            '観点': 'viewpoint_evaluations',
            '欠課情報': 'absences'
        }
        
        reply = QMessageBox.question(
            self,
            "確認",
            "全てのデータを1つのExcelファイルに出力しますか？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        try:
            data_dict = {}
            total_records = 0
            
            for data_type, table_name in table_mapping.items():
                try:
                    query = f"SELECT * FROM {table_name}"
                    rows = self.db_manager.fetch_all(query)
                    
                    if rows and len(rows) > 0:
                        data = [dict(row) for row in rows]
                        columns = list(data[0].keys())
                        data_dict[data_type] = (data, columns)
                        total_records += len(data)
                except:
                    pass
            
            if not data_dict:
                QMessageBox.information(self, "情報", "出力するデータがありません")
                return
            
            from utils.excel_exporter import ExcelExporter
            exporter = ExcelExporter()
            
            export_path = exporter.export_multiple_sheets(
                data_dict=data_dict,
                filename="全評価データ.xlsx"
            )
            
            QMessageBox.information(
                self,
                "出力完了",
                f"全データをExcelファイルに出力しました。\n\n"
                f"ファイル: {Path(export_path).name}\n"
                f"総レコード数: {total_records}件"
            )
        
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"Excel出力に失敗:\n{str(e)}")
    
    def clear_current_data(self):
        """現在のデータを削除"""
        current_index = self.tab_widget.currentIndex()
        if current_index < 0:
            return
        
        data_types = ['評定', '観点', '欠課情報']
        data_type = data_types[current_index]
        
        table_mapping = {
            '評定': 'grades',
            '観点': 'viewpoint_evaluations',
            '欠課情報': 'absences'
        }
        
        table_name = table_mapping.get(data_type)
        
        try:
            year = self.year_filter.value()
            period = self.period_filter.currentText()
            
            where_clauses = [f"year = {year}"]
            if period != "全て":
                where_clauses.append(f"period = '{period}'")
            
            where_str = " AND ".join(where_clauses)
            
            count_query = f"SELECT COUNT(*) as count FROM {table_name} WHERE {where_str}"
            result = self.db_manager.fetch_one(count_query)
            count = result['count'] if result else 0
            
            if count == 0:
                QMessageBox.information(self, "情報", "データがありません")
                return
            
            reply = QMessageBox.warning(
                self,
                "確認",
                f"{data_type}のデータ（{count}件）を削除しますか？\n"
                f"年度: {year}, 期間: {period}\n\n"
                "この操作は取り消せません。",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                delete_query = f"DELETE FROM {table_name} WHERE {where_str}"
                self.db_manager.execute_query(delete_query)
                
                QMessageBox.information(self, "削除完了", f"{count}件のデータを削除しました。")
                self.refresh_current_tab()
        
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"削除に失敗:\n{str(e)}")
    
    def open_required_columns_manager(self):
        """必須カラム管理ダイアログを開く"""
        try:
            from ui.required_columns_manager_dialog import RequiredColumnsManagerDialog
            
            dialog = RequiredColumnsManagerDialog(self)
            dialog.exec()
        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            print(error_detail)
            QMessageBox.critical(
                self,
                "エラー",
                f"必須カラム管理ダイアログの表示に失敗しました:\n{str(e)}"
            )
    
    def open_db_column_manager(self, data_type):
        """DBカラム設定ダイアログを開く"""
        try:
            from ui.db_column_manager_dialog import DBColumnManagerDialog
            
            dialog = DBColumnManagerDialog(data_type, self)
            dialog.exec()
        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            print(error_detail)
            QMessageBox.critical(
                self,
                "エラー",
                f"DBカラム設定ダイアログの表示に失敗しました:\n{str(e)}"
            )
    
    def open_preset_manager(self, data_type):
        """プリセット管理ダイアログを開く"""
        try:
            from ui.preset_manager_dialog import PresetManagerDialog
            
            dialog = PresetManagerDialog(data_type, self.config_manager, self)
            dialog.exec()
        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            print(error_detail)
            QMessageBox.critical(
                self,
                "エラー",
                f"プリセット管理ダイアログの表示に失敗しました:\n{str(e)}"
            )
    
    def show_workflow_guide(self):
        """ワークフローガイド表示"""
        guide_text = """ ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 評価・評定データ処理ワークフロー ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 【ステップ1】評価・評定データ取り込み ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 目的: 評定データ（1-10の評価値）を取り込む 手順: 1. 「評価・評定データ取り込み」ボタンをクリック 2. 期間（前期/後期/通年）と年度を選択 3. Excelファイルを選択 4. カラムマッピングを設定 5. 取り込み実行 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 【ステップ2】観点評価データ取り込み ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 目的: 観点評価（観点1-5）を取り込む 手順: 1. 「観点評価データ取り込み」ボタンをクリック 2. ステップ1と同じ期間・年度を選択 3. Excelファイルを選択 4. カラムマッピングを設定 5. 取り込み実行 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 【ステップ3】未入力者チェック ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 目的: 評価・観点の未入力者を確認 手順: 1. 「未入力者チェック」ボタンをクリック 2. 年度と期間を選択 3. チェック実行 4. 未入力者リストを確認 5. 必要に応じてExcel出力 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 【ステップ4】欠課データ前処理 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 目的: 複数の欠課データファイルを統合・集計 手順: 1. 「欠課データ前処理」ボタンをクリック 2. 複数のExcelファイルを選択 3. ヘッダー行を確認（通常は0行目） 4. 前処理実行 5. 処理済みファイルが自動生成される ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 【ステップ5】欠課情報取り込み ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 目的: 前処理済みの欠課データを取り込む 手順: 1. 「欠課情報取り込み」ボタンをクリック 2. ステップ4で生成されたファイルを選択 3. カラムマッピングを設定 4. 取り込み実行 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ """
        
        msg = QMessageBox(self)
        msg.setWindowTitle("ワークフローガイド")
        msg.setText(guide_text)
        msg.setStyleSheet("QLabel{min-width: 600px; font-family: 'MS Gothic';}")
        msg.exec()
    
    def show_about(self):
        """バージョン情報表示"""
        from __version__ import APP_NAME, APP_VERSION, __release_date__, APP_DESCRIPTION
        
        QMessageBox.about(
            self,
            "バージョン情報",
            f"{APP_NAME}\n"
            f"Version: {APP_VERSION}\n"
            f"リリース日: {__release_date__}\n\n"
            f"{APP_DESCRIPTION}\n\n"
            f"使用中のデータベース:\n{self.settings['database']['path']}"
        )
    
    def closeEvent(self, event):
        """ウィンドウを閉じる時の処理"""
        try:
            self.db_manager.close()
        except:
            pass
        event.accept()
