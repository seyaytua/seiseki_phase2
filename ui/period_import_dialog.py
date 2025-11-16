from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                               QPushButton, QComboBox, QSpinBox, QFileDialog,
                               QMessageBox, QTableWidget, QTableWidgetItem,
                               QProgressDialog, QCheckBox, QGroupBox)
from PySide6.QtCore import Qt
from datetime import datetime
from pathlib import Path
import pandas as pd


class PeriodImportDialog(QDialog):
    """期間別データ取り込みダイアログ（プレビュー機能付き）"""
    
    def __init__(self, data_type, db_manager, config_manager, file_manager, data_importer, parent=None):
        super().__init__(parent)
        self.data_type = data_type
        self.db_manager = db_manager
        self.config_manager = config_manager
        self.file_manager = file_manager
        self.data_importer = data_importer
        
        self.file_path = None
        self.sheet_names = []
        self.column_mapping = {}
        self.log_viewer = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """UI初期化"""
        self.setWindowTitle(f"{self.data_type}データ取り込み")
        self.setMinimumWidth(1000)
        self.setMinimumHeight(800)
        
        layout = QVBoxLayout(self)
        
        # タイトルと説明
        title_layout = QVBoxLayout()
        title_label = QLabel(f"{self.data_type}データ取り込み")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #2c3e50;")
        title_layout.addWidget(title_label)
        
        desc_label = QLabel("Excelファイルからデータを取り込みます。シート、ヘッダー行、カラムマッピングを設定してください。")
        desc_label.setStyleSheet("color: #7f8c8d; margin-bottom: 10px;")
        title_layout.addWidget(desc_label)
        
        layout.addLayout(title_layout)
        
        # ログビューアーボタン
        log_btn_layout = QHBoxLayout()
        log_btn_layout.addStretch()
        
        self.log_viewer_btn = QPushButton("📋 処理ログを表示")
        self.log_viewer_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        self.log_viewer_btn.clicked.connect(self.show_log_viewer)
        log_btn_layout.addWidget(self.log_viewer_btn)
        
        layout.addLayout(log_btn_layout)
        
        # 期間・年度選択
        period_layout = QHBoxLayout()
        
        period_layout.addWidget(QLabel("期間:"))
        self.period_combo = QComboBox()
        self.period_combo.addItems(['前期', '後期', '通年'])
        period_layout.addWidget(self.period_combo)
        
        period_layout.addWidget(QLabel("年度:"))
        self.year_spin = QSpinBox()
        self.year_spin.setMinimum(2000)
        self.year_spin.setMaximum(2100)
        self.year_spin.setValue(datetime.now().year)
        period_layout.addWidget(self.year_spin)
        
        period_layout.addStretch()
        layout.addLayout(period_layout)
        
        # ファイル選択
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("ファイル:"))
        
        self.file_label = QLabel("未選択")
        self.file_label.setStyleSheet("color: #e74c3c; font-weight: bold;")
        file_layout.addWidget(self.file_label)
        
        select_file_btn = QPushButton("📁 ファイル選択")
        select_file_btn.clicked.connect(self.select_file)
        file_layout.addWidget(select_file_btn)
        
        layout.addLayout(file_layout)
        
        # ヘッダー行設定
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("ヘッダー行:"))
        
        self.header_spin = QSpinBox()
        self.header_spin.setMinimum(0)
        self.header_spin.setMaximum(10)
        self.header_spin.setValue(0)
        self.header_spin.setToolTip("0 = 1行目がヘッダー")
        self.header_spin.valueChanged.connect(self.update_preview)
        header_layout.addWidget(self.header_spin)
        
        header_layout.addWidget(QLabel("行目（0始まり）"))
        
        preview_btn = QPushButton("🔄 プレビュー更新")
        preview_btn.clicked.connect(self.update_preview)
        header_layout.addWidget(preview_btn)
        
        header_layout.addStretch()
        layout.addLayout(header_layout)
        
        # プレビューエリア
        preview_group = QGroupBox("📊 データプレビュー（先頭10行）")
        preview_layout = QVBoxLayout()
        
        self.preview_table = QTableWidget()
        self.preview_table.setMaximumHeight(200)
        self.preview_table.setAlternatingRowColors(True)
        preview_layout.addWidget(self.preview_table)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        # シート選択エリア
        sheet_group = QGroupBox("📑 シート選択")
        sheet_layout = QVBoxLayout()
        
        self.sheet_table = QTableWidget()
        self.sheet_table.setColumnCount(2)
        self.sheet_table.setHorizontalHeaderLabels(['選択', 'シート名'])
        self.sheet_table.horizontalHeader().setStretchLastSection(True)
        self.sheet_table.setMaximumHeight(150)
        self.sheet_table.setAlternatingRowColors(True)
        self.sheet_table.itemSelectionChanged.connect(self.on_sheet_selected)
        sheet_layout.addWidget(self.sheet_table)
        
        sheet_btn_layout = QHBoxLayout()
        select_all_btn = QPushButton("✅ 全選択")
        select_all_btn.clicked.connect(self.select_all_sheets)
        sheet_btn_layout.addWidget(select_all_btn)
        
        deselect_all_btn = QPushButton("❌ 全解除")
        deselect_all_btn.clicked.connect(self.deselect_all_sheets)
        sheet_btn_layout.addWidget(deselect_all_btn)
        
        sheet_btn_layout.addStretch()
        sheet_layout.addLayout(sheet_btn_layout)
        
        sheet_group.setLayout(sheet_layout)
        layout.addWidget(sheet_group)
        
        # カラムマッピングエリア
        mapping_group = QGroupBox("🔗 カラムマッピング")
        mapping_layout = QVBoxLayout()
        
        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(2)
        self.mapping_table.setHorizontalHeaderLabels(['Excelカラム', 'データベースカラム'])
        self.mapping_table.setMaximumHeight(200)
        self.mapping_table.setAlternatingRowColors(True)
        mapping_layout.addWidget(self.mapping_table)
        
        mapping_btn_layout = QHBoxLayout()
        
        load_mapping_btn = QPushButton("💾 保存済みマッピング読み込み")
        load_mapping_btn.clicked.connect(self.load_saved_mapping)
        mapping_btn_layout.addWidget(load_mapping_btn)
        
        edit_mapping_btn = QPushButton("✏️ マッピング編集")
        edit_mapping_btn.clicked.connect(self.edit_mapping)
        mapping_btn_layout.addWidget(edit_mapping_btn)
        
        mapping_btn_layout.addStretch()
        mapping_layout.addLayout(mapping_btn_layout)
        
        mapping_group.setLayout(mapping_layout)
        layout.addWidget(mapping_group)
        
        # オプション
        option_layout = QHBoxLayout()
        self.timestamp_check = QCheckBox("⏰ ファイル名にタイムスタンプを追加")
        self.timestamp_check.setChecked(True)
        option_layout.addWidget(self.timestamp_check)
        option_layout.addStretch()
        layout.addLayout(option_layout)
        
        # ボタン
        button_layout = QHBoxLayout()
        
        import_btn = QPushButton("▶️ 取り込み実行")
        import_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        import_btn.clicked.connect(self.execute_import)
        button_layout.addWidget(import_btn)
        
        cancel_btn = QPushButton("❌ キャンセル")
        cancel_btn.setStyleSheet("""
            QPushButton {
                padding: 10px;
            }
        """)
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
    
    def show_log_viewer(self):
        """ログビューアーを表示"""
        try:
            from ui.log_viewer_dialog import LogViewerDialog
            
            if not hasattr(self, 'log_viewer') or self.log_viewer is None:
                self.log_viewer = LogViewerDialog(self)
            
            self.log_viewer.show()
            self.log_viewer.raise_()
            self.log_viewer.activateWindow()
            
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"ログビューアーの起動に失敗しました:\n{str(e)}")
    
    def select_file(self):
        """ファイル選択"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Excelファイル選択",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.file_path = file_path
            file_name = Path(file_path).name
            self.file_label.setText(file_name)
            self.file_label.setStyleSheet("color: #27ae60; font-weight: bold;")
            
            # シート名読み込み
            self.load_sheet_names()
            
            # プレビュー表示
            self.update_preview()
    
    def load_sheet_names(self):
        """シート名読み込み"""
        try:
            excel_file = pd.ExcelFile(self.file_path)
            self.sheet_names = excel_file.sheet_names
            
            # テーブルに表示
            self.sheet_table.setRowCount(len(self.sheet_names))
            
            for i, sheet_name in enumerate(self.sheet_names):
                # チェックボックス
                check_item = QTableWidgetItem()
                check_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                check_item.setCheckState(Qt.Checked)
                self.sheet_table.setItem(i, 0, check_item)
                
                # シート名
                name_item = QTableWidgetItem(sheet_name)
                self.sheet_table.setItem(i, 1, name_item)
            
            self.sheet_table.resizeColumnsToContents()
            
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"シート読み込みエラー:\n{str(e)}")
    
    def on_sheet_selected(self):
        """シート選択時の処理"""
        self.update_preview()
    
    def update_preview(self):
        """プレビュー更新"""
        if not self.file_path:
            return
        
        try:
            # 選択されているシートを取得
            current_row = self.sheet_table.currentRow()
            if current_row >= 0:
                sheet_name = self.sheet_table.item(current_row, 1).text()
            else:
                # 最初のシートを使用
                sheet_name = self.sheet_names[0] if self.sheet_names else 0
            
            header_row = self.header_spin.value()
            
            # データ読み込み（先頭10行）
            df = pd.read_excel(
                self.file_path,
                sheet_name=sheet_name,
                header=header_row,
                nrows=10
            )
            
            # プレビューテーブルに表示
            self.preview_table.clear()
            self.preview_table.setRowCount(len(df))
            self.preview_table.setColumnCount(len(df.columns))
            self.preview_table.setHorizontalHeaderLabels([str(col) for col in df.columns])
            
            for i in range(len(df)):
                for j, col in enumerate(df.columns):
                    value = df.iloc[i, j]
                    item = QTableWidgetItem(str(value) if pd.notna(value) else '')
                    self.preview_table.setItem(i, j, item)
            
            self.preview_table.resizeColumnsToContents()
            
            # カラムマッピングテーブルも更新
            self.load_columns(df.columns.tolist())
            
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"プレビュー表示エラー:\n{str(e)}")
    
    def select_all_sheets(self):
        """全シート選択"""
        for i in range(self.sheet_table.rowCount()):
            self.sheet_table.item(i, 0).setCheckState(Qt.Checked)
    
    def deselect_all_sheets(self):
        """全シート解除"""
        for i in range(self.sheet_table.rowCount()):
            self.sheet_table.item(i, 0).setCheckState(Qt.Unchecked)
    
    def load_columns(self, columns):
        """カラム読み込み"""
        try:
            # テーブルに表示
            self.mapping_table.setRowCount(len(columns))
            
            for i, col in enumerate(columns):
                # Excelカラム
                excel_item = QTableWidgetItem(str(col))
                self.mapping_table.setItem(i, 0, excel_item)
                
                # データベースカラム（コンボボックス）
                db_combo = QComboBox()
                db_combo.addItem("")  # 空白オプション
                
                # データタイプに応じたカラム
                if self.data_type == '評定':
                    db_combo.addItems(['student_number', 'student_name', 'course_number', 
                                      'course_name', 'school_subject_name', 'grade_value', 
                                      'credits', 'acquisition_credits', 'remarks'])
                elif self.data_type == '観点':
                    db_combo.addItems(['student_number', 'student_name', 'course_number',
                                      'course_name', 'school_subject_name',
                                      'viewpoint_1', 'viewpoint_2', 'viewpoint_3',
                                      'viewpoint_4', 'viewpoint_5', 'remarks'])
                elif self.data_type == '欠課情報':
                    db_combo.addItems(['student_number', 'student_name', 'course_number',
                                      'course_name', 'school_subject_name', 'absent_count',
                                      'late_count', 'total_hours', 'absence_rate', 'remarks'])
                
                self.mapping_table.setCellWidget(i, 1, db_combo)
            
            self.mapping_table.resizeColumnsToContents()
            
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"カラム読み込みエラー:\n{str(e)}")
    
    def load_saved_mapping(self):
        """保存済みマッピング読み込み"""
        try:
            mappings = self.config_manager.get_column_mapping(self.data_type)
            
            if not mappings:
                QMessageBox.information(self, "情報", "保存済みマッピングがありません")
                return
            
            # マッピング適用
            for i in range(self.mapping_table.rowCount()):
                excel_col = self.mapping_table.item(i, 0).text()
                
                if excel_col in mappings:
                    db_col = mappings[excel_col]
                    combo = self.mapping_table.cellWidget(i, 1)
                    
                    index = combo.findText(db_col)
                    if index >= 0:
                        combo.setCurrentIndex(index)
            
            QMessageBox.information(self, "完了", "マッピングを読み込みました")
            
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"マッピング読み込みエラー:\n{str(e)}")
    
    def edit_mapping(self):
        """マッピング編集ダイアログを開く"""
        if not self.file_path:
            QMessageBox.warning(self, "警告", "先にファイルを選択してください")
            return
        
        try:
            from ui.column_mapping_dialog import ColumnMappingDialog
            
            # 現在のExcelカラムを取得
            current_row = self.sheet_table.currentRow()
            if current_row >= 0:
                sheet_name = self.sheet_table.item(current_row, 1).text()
            else:
                sheet_name = self.sheet_names[0] if self.sheet_names else 0
            
            header_row = self.header_spin.value()
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=header_row, nrows=0)
            excel_columns = df.columns.tolist()
            
            # マッピングダイアログを開く
            dialog = ColumnMappingDialog(
                self.data_type,
                self.config_manager,
                excel_columns,
                self
            )
            
            if dialog.exec():
                # マッピング再読み込み
                self.load_saved_mapping()
        
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"マッピング編集エラー:\n{str(e)}")
    
    def get_selected_sheets(self):
        """選択されたシート名を取得"""
        selected = []
        for i in range(self.sheet_table.rowCount()):
            if self.sheet_table.item(i, 0).checkState() == Qt.Checked:
                sheet_name = self.sheet_table.item(i, 1).text()
                selected.append(sheet_name)
        return selected
    
    def get_column_mapping(self):
        """カラムマッピング取得"""
        mapping = {}
        
        for i in range(self.mapping_table.rowCount()):
            excel_col = self.mapping_table.item(i, 0).text()
            combo = self.mapping_table.cellWidget(i, 1)
            db_col = combo.currentText()
            
            if db_col:  # 空白でない場合のみ
                mapping[excel_col] = db_col
        
        return mapping
    
    def execute_import(self):
        """取り込み実行"""
        # 選択シート取得
        selected_sheets = self.get_selected_sheets()
        
        if not selected_sheets:
            QMessageBox.warning(self, "警告", "シートを選択してください")
            return
        
        # カラムマッピング取得
        self.column_mapping = self.get_column_mapping()
        
        if not self.column_mapping:
            QMessageBox.warning(self, "警告", "カラムマッピングを設定してください")
            return
        
        # 必須カラムチェック
        required_columns = ['student_number', 'course_number']
        missing_columns = [col for col in required_columns if col not in self.column_mapping.values()]
        
        if missing_columns:
            QMessageBox.warning(
                self, 
                "警告", 
                f"必須カラムが不足しています:\n{', '.join(missing_columns)}"
            )
            return
        
        # 確認ダイアログ
        reply = QMessageBox.question(
            self,
            "確認",
            f"以下の内容で取り込みを実行しますか？\n\n"
            f"データタイプ: {self.data_type}\n"
            f"期間: {self.period_combo.currentText()}\n"
            f"年度: {self.year_spin.value()}\n"
            f"ファイル: {Path(self.file_path).name}\n"
            f"シート数: {len(selected_sheets)}\n"
            f"ヘッダー行: {self.header_spin.value()}",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        # タイムスタンプ追加確認
        add_timestamp = self.timestamp_check.isChecked()
        
        # 進捗ダイアログ
        progress = QProgressDialog("取り込み中...", "キャンセル", 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        
        def update_progress(current, total, message):
            if total > 0:
                percent = int((current / total) * 100)
                progress.setValue(percent)
                progress.setLabelText(message)
        
        try:
            # データ取り込み実行
            success = self.data_importer.import_data(
                file_path=self.file_path,
                data_type=self.data_type,
                period=self.period_combo.currentText(),
                year=self.year_spin.value(),
                column_mapping=self.column_mapping,
                sheet_names=selected_sheets,
                header_row=self.header_spin.value(),
                progress_callback=update_progress,
                add_timestamp=add_timestamp
            )
            
            progress.setValue(100)
            
            if success:
                # マッピング保存確認
                reply = QMessageBox.question(
                    self,
                    "マッピング保存",
                    "このカラムマッピングを保存しますか?",
                    QMessageBox.Yes | QMessageBox.No
                )
                
                if reply == QMessageBox.Yes:
                    self.config_manager.save_column_mapping(
                        self.data_type,
                        self.column_mapping
                    )
                
                QMessageBox.information(self, "完了", f"{self.data_type}の取り込みが完了しました")
                self.accept()
            
        except Exception as e:
            progress.close()
            import traceback
            error_detail = traceback.format_exc()
            print(error_detail)
            QMessageBox.critical(self, "エラー", f"取り込みエラー:\n{str(e)}")