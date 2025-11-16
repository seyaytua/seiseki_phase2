from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                               QPushButton, QListWidget, QSpinBox, QFileDialog,
                               QMessageBox, QProgressDialog, QGroupBox, QTextEdit,
                               QTableWidget, QTableWidgetItem, QCheckBox, QComboBox,
                               QListWidgetItem)
from PySide6.QtCore import Qt
from pathlib import Path
from utils.absence_processor import AbsenceProcessor
import json


class AbsencePreprocessorDialog(QDialog):
    """欠課データ前処理ダイアログ（マッピング機能付き）"""
    
    def __init__(self, config_manager, parent=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.processor = AbsenceProcessor()
        self.file_paths = []
        self.column_mapping = {}
        self.log_viewer = None
        
        # DBカラム情報を読み込む
        self.load_db_columns()
        
        self.setup_ui()
    
    def load_db_columns(self):
        """DBカラム設定を読み込む"""
        db_columns_path = Path('config/db_columns.json')
        
        self.db_columns = []
        self.db_columns_info = {}
        self.output_columns = []
        
        if db_columns_path.exists():
            try:
                with open(db_columns_path, 'r', encoding='utf-8') as f:
                    columns_config = json.load(f)
                    
                    if '欠課情報' in columns_config:
                        for col in columns_config['欠課情報']:
                            col_name = col['name']
                            self.db_columns.append(col_name)
                            self.db_columns_info[col_name] = col.get('description', '')
                            
                            if col_name not in ['absence_mark', 'absence_type']:
                                self.output_columns.append(col_name)
            except Exception as e:
                print(f"DBカラム設定の読み込みエラー: {e}")
        
        if not self.db_columns:
            self.db_columns = [
                'student_number', 'class_name', 'attendance_number', 'student_name',
                'absent_count', 'course_name', 'subject_category_number',
                'subject_number', 'course_number', 'absence_mark', 'absence_type'
            ]
            self.output_columns = [
                'student_number', 'class_name', 'attendance_number', 'student_name',
                'absent_count', 'course_name', 'subject_category_number',
                'subject_number', 'course_number'
            ]
    
    def setup_ui(self):
        """UI初期化"""
        self.setWindowTitle("欠課データ前処理")
        self.setMinimumWidth(1000)
        self.setMinimumHeight(800)
        
        layout = QVBoxLayout(self)
        
        # 説明
        info_label = QLabel(
            "【欠課データ前処理の手順】\n"
            "1. 複数のExcelファイルを追加\n"
            "2. ヘッダー行を確認・設定（プレビューで確認）\n"
            "3. カラムマッピングを設定（自動設定されます）\n"
            "4. 前処理を実行（欠課略号「/」または欠課区分「1」を集計）\n"
            "5. 出力するカラムを選択\n"
            "6. 処理結果をExcel出力"
        )
        info_label.setStyleSheet("padding: 10px; background-color: #E3F2FD; border-radius: 5px;")
        layout.addWidget(info_label)
        
        # ログビューアーボタン
        log_btn_layout = QHBoxLayout()
        log_btn_layout.addStretch()
        
        self.log_viewer_btn = QPushButton("📋 処理ログを表示")
        self.log_viewer_btn.setStyleSheet(""" QPushButton { background-color: #2196F3; color: white; padding: 8px 16px; border-radius: 4px; font-weight: bold; } QPushButton:hover { background-color: #1976D2; } """)
        self.log_viewer_btn.clicked.connect(self.show_log_viewer)
        log_btn_layout.addWidget(self.log_viewer_btn)
        
        layout.addLayout(log_btn_layout)
        
        # ファイルリストエリア
        file_group = QGroupBox("処理対象ファイル")
        file_layout = QVBoxLayout()
        
        self.file_list = QListWidget()
        self.file_list.currentRowChanged.connect(self.on_file_selected)
        file_layout.addWidget(self.file_list)
        
        file_btn_layout = QHBoxLayout()
        
        add_files_btn = QPushButton("ファイル追加")
        add_files_btn.clicked.connect(self.add_files)
        file_btn_layout.addWidget(add_files_btn)
        
        remove_file_btn = QPushButton("選択削除")
        remove_file_btn.clicked.connect(self.remove_selected_file)
        file_btn_layout.addWidget(remove_file_btn)
        
        clear_files_btn = QPushButton("全削除")
        clear_files_btn.clicked.connect(self.clear_files)
        file_btn_layout.addWidget(clear_files_btn)
        
        file_btn_layout.addStretch()
        file_layout.addLayout(file_btn_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # ヘッダー行設定とプレビュー
        header_preview_group = QGroupBox("ヘッダー行設定とプレビュー")
        header_preview_layout = QVBoxLayout()
        
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
        
        preview_btn = QPushButton("プレビュー更新")
        preview_btn.clicked.connect(self.update_preview)
        header_layout.addWidget(preview_btn)
        
        header_layout.addStretch()
        header_preview_layout.addLayout(header_layout)
        
        preview_label = QLabel("データプレビュー（選択ファイルの先頭10行）:")
        header_preview_layout.addWidget(preview_label)
        
        self.preview_table = QTableWidget()
        self.preview_table.setMaximumHeight(150)
        header_preview_layout.addWidget(self.preview_table)
        
        header_preview_group.setLayout(header_preview_layout)
        layout.addWidget(header_preview_group)
        
        # カラムマッピングエリア
        mapping_group = QGroupBox("カラムマッピング（自動設定）")
        mapping_layout = QVBoxLayout()
        
        mapping_info = QLabel(
            "Excelのカラムとデータベースカラムの対応を設定してください。\n"
            "必須: 生徒番号、講座番号、欠課略号または欠課区分"
        )
        mapping_info.setStyleSheet("color: #666; font-size: 9pt;")
        mapping_layout.addWidget(mapping_info)
        
        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(2)
        self.mapping_table.setHorizontalHeaderLabels(['Excelカラム', 'データベースカラム'])
        self.mapping_table.setMaximumHeight(150)
        mapping_layout.addWidget(self.mapping_table)
        
        mapping_btn_layout = QHBoxLayout()
        
        load_mapping_btn = QPushButton("保存済みマッピング読み込み")
        load_mapping_btn.clicked.connect(self.load_saved_mapping)
        mapping_btn_layout.addWidget(load_mapping_btn)
        
        edit_mapping_btn = QPushButton("マッピング編集")
        edit_mapping_btn.clicked.connect(self.edit_mapping)
        mapping_btn_layout.addWidget(edit_mapping_btn)
        
        mapping_btn_layout.addStretch()
        mapping_layout.addLayout(mapping_btn_layout)
        
        mapping_group.setLayout(mapping_layout)
        layout.addWidget(mapping_group)
        
        # 出力カラム選択エリア
        output_group = QGroupBox("出力カラム選択")
        output_layout = QVBoxLayout()
        
        output_info = QLabel(
            "Excel出力時に含めるカラムを選択してください:\n"
            "推奨: 生徒番号、組、番号、氏名、欠課、講座名、教科番号、科目番号、講座番号"
        )
        output_info.setStyleSheet("color: #666; font-size: 9pt;")
        output_layout.addWidget(output_info)
        
        self.output_columns_list = QListWidget()
        self.output_columns_list.setMaximumHeight(120)
        output_layout.addWidget(self.output_columns_list)
        
        self.initialize_output_columns()
        
        output_btn_layout = QHBoxLayout()
        
        select_all_output_btn = QPushButton("全選択")
        select_all_output_btn.clicked.connect(self.select_all_output_columns)
        output_btn_layout.addWidget(select_all_output_btn)
        
        deselect_all_output_btn = QPushButton("全解除")
        deselect_all_output_btn.clicked.connect(self.deselect_all_output_columns)
        output_btn_layout.addWidget(deselect_all_output_btn)
        
        default_output_btn = QPushButton("推奨カラムを選択")
        default_output_btn.clicked.connect(self.select_default_output_columns)
        output_btn_layout.addWidget(default_output_btn)
        
        output_btn_layout.addStretch()
        output_layout.addLayout(output_btn_layout)
        
        output_group.setLayout(output_layout)
        layout.addWidget(output_group)
        
        # 処理結果プレビューエリア
        result_group = QGroupBox("処理結果プレビュー")
        result_layout = QVBoxLayout()
        
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setMaximumHeight(120)
        result_layout.addWidget(self.result_text)
        
        result_group.setLayout(result_layout)
        layout.addWidget(result_group)
        
        # 実行ボタン
        button_layout = QHBoxLayout()
        
        process_btn = QPushButton("前処理実行")
        process_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        process_btn.clicked.connect(self.execute_preprocessing)
        button_layout.addWidget(process_btn)
        
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(self.reject)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
    
    def show_log_viewer(self):
        """ログビューアーを表示"""
        from ui.log_viewer_dialog import LogViewerDialog
        
        if self.log_viewer is None:
            self.log_viewer = LogViewerDialog(self)
        
        self.log_viewer.show()
        self.log_viewer.raise_()
        self.log_viewer.activateWindow()
    
    def initialize_output_columns(self):
        """出力カラムリストを初期化"""
        self.output_columns_list.clear()
        
        default_output = [
            'student_number', 'class_name', 'attendance_number', 'student_name',
            'absent_count', 'course_name', 'subject_category_number',
            'subject_number', 'course_number'
        ]
        
        for col_name in self.output_columns:
            description = self.db_columns_info.get(col_name, '')
            
            if description:
                display_text = f"{col_name} ({description})"
            else:
                display_text = col_name
            
            item = QListWidgetItem(display_text)
            item.setData(Qt.UserRole, col_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            
            if col_name in default_output:
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
            
            self.output_columns_list.addItem(item)
    
    def select_all_output_columns(self):
        """全出力カラムを選択"""
        for i in range(self.output_columns_list.count()):
            self.output_columns_list.item(i).setCheckState(Qt.Checked)
    
    def deselect_all_output_columns(self):
        """全出力カラムを解除"""
        for i in range(self.output_columns_list.count()):
            self.output_columns_list.item(i).setCheckState(Qt.Unchecked)
    
    def select_default_output_columns(self):
        """推奨カラムを選択"""
        default_output = [
            'student_number', 'class_name', 'attendance_number', 'student_name',
            'absent_count', 'course_name', 'subject_category_number',
            'subject_number', 'course_number'
        ]
        
        for i in range(self.output_columns_list.count()):
            item = self.output_columns_list.item(i)
            col_name = item.data(Qt.UserRole)
            
            if col_name in default_output:
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
    
    def get_selected_output_columns(self):
        """選択された出力カラムを取得"""
        selected = []
        for i in range(self.output_columns_list.count()):
            item = self.output_columns_list.item(i)
            if item.checkState() == Qt.Checked:
                col_name = item.data(Qt.UserRole)
                selected.append(col_name)
        return selected
    
    def add_files(self):
        """ファイル追加"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Excelファイル選択",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        
        if file_paths:
            for file_path in file_paths:
                if file_path not in self.file_paths:
                    self.file_paths.append(file_path)
                    self.file_list.addItem(Path(file_path).name)
            
            if len(self.file_paths) == len(file_paths):
                self.file_list.setCurrentRow(0)
    
    def remove_selected_file(self):
        """選択ファイル削除"""
        current_row = self.file_list.currentRow()
        if current_row >= 0:
            self.file_list.takeItem(current_row)
            del self.file_paths[current_row]
            
            if not self.file_paths:
                self.preview_table.clear()
                self.preview_table.setRowCount(0)
                self.preview_table.setColumnCount(0)
                self.mapping_table.setRowCount(0)
    
    def clear_files(self):
        """全ファイル削除"""
        self.file_list.clear()
        self.file_paths = []
        self.preview_table.clear()
        self.preview_table.setRowCount(0)
        self.preview_table.setColumnCount(0)
        self.mapping_table.setRowCount(0)
    
    def on_file_selected(self, row):
        """ファイル選択時の処理"""
        if row >= 0:
            self.update_preview()
    
    def update_preview(self):
        """プレビュー更新"""
        current_row = self.file_list.currentRow()
        if current_row < 0 or current_row >= len(self.file_paths):
            return
        
        try:
            import pandas as pd
            
            file_path = self.file_paths[current_row]
            header_row = self.header_spin.value()
            
            df = pd.read_excel(file_path, header=header_row, nrows=10)
            
            self.preview_table.clear()
            self.preview_table.setRowCount(len(df))
            self.preview_table.setColumnCount(len(df.columns))
            self.preview_table.setHorizontalHeaderLabels([str(col) for col in df.columns])
            
            for i in range(len(df)):
                for j, col in enumerate(df.columns):
                    value = df.iloc[i, j]
                    item = QTableWidgetItem(str(value) if pd.notna(value) else '')
                    
                    if pd.notna(value):
                        if '/' in str(value) or str(value) == '1':
                            item.setBackground(Qt.yellow)
                    
                    self.preview_table.setItem(i, j, item)
            
            self.preview_table.resizeColumnsToContents()
            self.load_columns(df.columns.tolist())
            
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"プレビュー表示エラー:\n{str(e)}")
    
    def load_columns(self, columns):
        """カラムマッピングテーブルを更新"""
        self.mapping_table.setRowCount(len(columns))
        
        auto_mapping = {
            '生徒番号': 'student_number',
            '組': 'class_name',
            '番号': 'attendance_number',
            '氏名': 'student_name',
            '講座名': 'course_name',
            '教科番号': 'subject_category_number',
            '科目番号': 'subject_number',
            '講座番号': 'course_number',
            '欠課略号': 'absence_mark',
            '欠課区分': 'absence_type'
        }
        
        for i, col in enumerate(columns):
            excel_item = QTableWidgetItem(str(col))
            self.mapping_table.setItem(i, 0, excel_item)
            
            db_combo = QComboBox()
            db_combo.addItem("", "")
            
            for db_col_name in self.db_columns:
                description = self.db_columns_info.get(db_col_name, '')
                if description:
                    db_combo.addItem(f"{db_col_name} ({description})", db_col_name)
                else:
                    db_combo.addItem(db_col_name, db_col_name)
            
            if str(col) in auto_mapping:
                mapped_col = auto_mapping[str(col)]
                for j in range(db_combo.count()):
                    if db_combo.itemData(j) == mapped_col:
                        db_combo.setCurrentIndex(j)
                        break
            
            self.mapping_table.setCellWidget(i, 1, db_combo)
        
        self.mapping_table.resizeColumnsToContents()
    
    def load_saved_mapping(self):
        """保存済みマッピング読み込み"""
        try:
            mappings = self.config_manager.get_column_mapping('欠課情報')
            
            if not mappings:
                QMessageBox.information(self, "情報", "保存済みマッピングがありません")
                return
            
            for i in range(self.mapping_table.rowCount()):
                excel_col = self.mapping_table.item(i, 0).text()
                
                if excel_col in mappings:
                    db_col = mappings[excel_col]
                    combo = self.mapping_table.cellWidget(i, 1)
                    
                    for j in range(combo.count()):
                        if combo.itemData(j) == db_col:
                            combo.setCurrentIndex(j)
                            break
            
            QMessageBox.information(self, "完了", "マッピングを読み込みました")
            
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"マッピング読み込みエラー:\n{str(e)}")
    
    def edit_mapping(self):
        """マッピング編集ダイアログを開く"""
        if not self.file_paths:
            QMessageBox.warning(self, "警告", "先にファイルを選択してください")
            return
        
        try:
            from ui.column_mapping_dialog import ColumnMappingDialog
            
            current_row = self.file_list.currentRow()
            if current_row < 0:
                current_row = 0
            
            file_path = self.file_paths[current_row]
            header_row = self.header_spin.value()
            
            import pandas as pd
            df = pd.read_excel(file_path, header=header_row, nrows=0)
            excel_columns = df.columns.tolist()
            
            dialog = ColumnMappingDialog('欠課情報', self.config_manager, excel_columns, self)
            
            if dialog.exec():
                self.load_saved_mapping()
        
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"マッピング編集エラー:\n{str(e)}")
    
    def get_column_mapping(self):
        """現在のカラムマッピングを取得"""
        mapping = {}
        
        for i in range(self.mapping_table.rowCount()):
            excel_col = self.mapping_table.item(i, 0).text()
            combo = self.mapping_table.cellWidget(i, 1)
            db_col = combo.currentData()
            
            if excel_col and db_col:
                mapping[excel_col] = db_col
        
        return mapping
    
    def execute_preprocessing(self):
        """前処理実行"""
        if not self.file_paths:
            QMessageBox.warning(self, "警告", "処理対象ファイルを追加してください")
            return
        
        self.column_mapping = self.get_column_mapping()
        
        required_columns = ['student_number', 'course_number']
        db_columns = list(self.column_mapping.values())
        
        missing_columns = [col for col in required_columns if col not in db_columns]
        
        if missing_columns:
            QMessageBox.warning(
                self,
                "警告",
                f"必須カラムがマッピングされていません:\n\n" +
                '\n'.join([f"• {col}" for col in missing_columns])
            )
            return
        
        has_absence_mark = 'absence_mark' in db_columns
        has_absence_type = 'absence_type' in db_columns
        
        if not has_absence_mark and not has_absence_type:
            QMessageBox.warning(
                self,
                "警告",
                "欠課判定用のカラムがマッピングされていません。\n\n"
                "以下のいずれかをマッピングしてください:\n"
                "• absence_mark (欠課略号)\n"
                "• absence_type (欠課区分)"
            )
            return
        
        reply = QMessageBox.question(
            self,
            "確認",
            f"{len(self.file_paths)}個のファイルを処理しますか？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        progress = QProgressDialog("前処理実行中...", "キャンセル", 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        
        def update_progress(current, total, message):
            if total > 0:
                percent = int((current / total) * 100)
                progress.setValue(percent)
                progress.setLabelText(message)
        
        try:
            header_row = self.header_spin.value()
            
            result_df = self.processor.process_multiple_files(
                self.file_paths,
                header_row=header_row,
                column_mapping=self.column_mapping,
                progress_callback=update_progress
            )
            
            debug_info = self.processor.get_debug_info()
            
            if result_df is None or len(result_df) == 0:
                progress.close()
                QMessageBox.warning(self, "警告", "欠課データが見つかりませんでした")
                return
            
            summary = self.processor.get_summary()
            
            preview_text = f""" 処理完了！ 【処理結果サマリー】 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 総レコード数: {summary['total_records']:,}件 ユニーク学生数: {summary['unique_students']}人 ユニーク講座数: {summary['unique_courses']}科目 総欠課数: {summary['total_absences']:,}回 平均欠課数: {summary['average_absences']}回/人・科目 欠課0の組み合わせ: {summary['zero_absence_count']:,}件 生徒あたり平均履修講座数: {summary['courses_per_student']}講座 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 【次のステップ】 出力カラムを確認して「Excel出力」を実行してください。 """
            
            self.result_text.setText(preview_text)
            progress.setValue(100)
            
            reply = QMessageBox.question(
                self,
                "マッピング保存",
                "このカラムマッピングを保存しますか？",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.config_manager.save_column_mapping('欠課情報', self.column_mapping)
            
            reply = QMessageBox.question(
                self,
                "Excel出力",
                f"処理結果をExcelファイルに出力しますか？\n\nレコード数: {summary['total_records']:,}件",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                selected_columns = self.get_selected_output_columns()
                
                if not selected_columns:
                    QMessageBox.warning(self, "警告", "出力するカラムを選択してください")
                    return
                
                output_path = self.processor.export_to_excel(selected_columns=selected_columns)
                
                QMessageBox.information(
                    self,
                    "出力完了",
                    f"Excelファイルに出力しました:\n{output_path}\n\n"
                    f"レコード数: {summary['total_records']:,}件\n"
                    f"出力カラム数: {len(selected_columns)}個"
                )
                
                reply = QMessageBox.question(
                    self,
                    "フォルダを開く",
                    "出力先フォルダを開きますか？",
                    QMessageBox.Yes | QMessageBox.No
                )
                
                if reply == QMessageBox.Yes:
                    import os
                    output_dir = Path(output_path).parent
                    if os.name == 'nt':
                        os.startfile(output_dir)
                
                self.accept()
        
        except Exception as e:
            progress.close()
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "エラー", f"前処理エラー:\n{str(e)}")