from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                               QPushButton, QTableWidget, QTableWidgetItem,
                               QComboBox, QMessageBox, QGroupBox, QLineEdit,
                               QInputDialog, QTextEdit)
from PySide6.QtCore import Qt
import json
from pathlib import Path


class ColumnMappingDialog(QDialog):
    """カラムマッピング編集ダイアログ"""
    
    def __init__(self, data_type, config_manager, excel_columns, parent=None):
        super().__init__(parent)
        self.data_type = data_type
        self.config_manager = config_manager
        self.excel_columns = excel_columns
        
        # DBカラム情報を読み込む
        self.load_db_columns()
        
        self.setup_ui()
        self.load_current_mapping()
    
    def load_db_columns(self):
        """DBカラム設定から利用可能なカラムを読み込む"""
        db_columns_path = Path('config/db_columns.json')
        
        self.db_columns = []
        self.db_columns_info = {}
        
        if db_columns_path.exists():
            try:
                with open(db_columns_path, 'r', encoding='utf-8') as f:
                    columns_config = json.load(f)
                    
                    if self.data_type in columns_config:
                        for col in columns_config[self.data_type]:
                            col_name = col['name']
                            self.db_columns.append(col_name)
                            self.db_columns_info[col_name] = {
                                'type': col.get('type', 'TEXT'),
                                'description': col.get('description', '')
                            }
            except Exception as e:
                print(f"DBカラム設定の読み込みエラー: {e}")
        
        # デフォルトのカラムリスト（設定ファイルがない場合）
        if not self.db_columns:
            if self.data_type == '評定':
                self.db_columns = [
                    'year', 'period', 'student_number', 'student_name',
                    'course_number', 'course_name', 'school_subject_name',
                    'grade_value', 'credits', 'acquisition_credits', 'remarks'
                ]
            elif self.data_type == '観点':
                self.db_columns = [
                    'year', 'period', 'student_number', 'student_name',
                    'course_number', 'course_name', 'school_subject_name',
                    'viewpoint_1', 'viewpoint_2', 'viewpoint_3',
                    'viewpoint_4', 'viewpoint_5', 'remarks'
                ]
            elif self.data_type == '欠課情報':
                self.db_columns = [
                    'year', 'period', 'student_number', 'student_name',
                    'course_number', 'course_name', 'school_subject_name',
                    'absent_hours', 'late_hours', 'total_hours',
                    'absence_rate', 'remarks'
                ]
    
    def setup_ui(self):
        """UI初期化"""
        self.setWindowTitle(f"{self.data_type} - カラムマッピング編集")
        self.setMinimumWidth(900)
        self.setMinimumHeight(700)
        
        layout = QVBoxLayout(self)
        
        # 説明
        info_label = QLabel(
            f"Excelファイルのカラムとデータベースのカラムの対応を設定してください。\n"
            f"このマッピングは保存され、次回以降の取り込みで使用されます。"
        )
        info_label.setStyleSheet("padding: 10px; background-color: #E3F2FD; border-radius: 5px;")
        layout.addWidget(info_label)
        
        # マッピングテーブル
        mapping_group = QGroupBox("カラムマッピング")
        mapping_layout = QVBoxLayout()
        
        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(3)
        self.mapping_table.setHorizontalHeaderLabels(['Excelカラム', 'データベースカラム', '操作'])
        self.mapping_table.horizontalHeader().setStretchLastSection(False)
        self.mapping_table.setColumnWidth(0, 250)
        self.mapping_table.setColumnWidth(1, 400)
        self.mapping_table.setColumnWidth(2, 100)
        mapping_layout.addWidget(self.mapping_table)
        
        # マッピング操作ボタン
        mapping_btn_layout = QHBoxLayout()
        
        add_row_btn = QPushButton("行を追加")
        add_row_btn.clicked.connect(self.add_mapping_row)
        mapping_btn_layout.addWidget(add_row_btn)
        
        clear_all_btn = QPushButton("すべてクリア")
        clear_all_btn.clicked.connect(self.clear_all_mappings)
        mapping_btn_layout.addWidget(clear_all_btn)
        
        mapping_btn_layout.addStretch()
        mapping_layout.addLayout(mapping_btn_layout)
        
        mapping_group.setLayout(mapping_layout)
        layout.addWidget(mapping_group)
        
        # データベースカラム情報
        db_info_group = QGroupBox("データベースカラム情報")
        db_info_layout = QVBoxLayout()
        
        self.db_info_text = QTextEdit()
        self.db_info_text.setReadOnly(True)
        self.db_info_text.setMaximumHeight(150)
        
        # カラム情報を整形して表示
        info_lines = [f"【{self.data_type} - 利用可能なカラム】\n"]
        
        # 必須カラムを取得
        required_columns = self.get_required_columns()
        
        for col_name in self.db_columns:
            col_info = self.db_columns_info.get(col_name, {})
            data_type = col_info.get('type', 'TEXT')
            description = col_info.get('description', '')
            
            required_mark = " [必須]" if col_name in required_columns else ""
            
            if description:
                info_lines.append(f"• {col_name} ({data_type}){required_mark}: {description}")
            else:
                info_lines.append(f"• {col_name} ({data_type}){required_mark}")
        
        self.db_info_text.setText('\n'.join(info_lines))
        db_info_layout.addWidget(self.db_info_text)
        
        db_info_group.setLayout(db_info_layout)
        layout.addWidget(db_info_group)
        
        # ボタン
        button_layout = QHBoxLayout()
        
        save_btn = QPushButton("保存")
        save_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        save_btn.clicked.connect(self.save_mapping)
        button_layout.addWidget(save_btn)
        
        cancel_btn = QPushButton("キャンセル")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        # テーブル初期化
        self.initialize_table()
    
    def get_required_columns(self):
        """必須カラムを取得"""
        required_path = Path('config/required_columns.json')
        
        if required_path.exists():
            try:
                with open(required_path, 'r', encoding='utf-8') as f:
                    required_config = json.load(f)
                    return required_config.get(self.data_type, [])
            except:
                pass
        
        # デフォルトの必須カラム
        return ['student_number', 'course_number']
    
    def initialize_table(self):
        """テーブル初期化"""
        self.mapping_table.setRowCount(len(self.excel_columns))
        
        for i, excel_col in enumerate(self.excel_columns):
            self.create_mapping_row(i, excel_col)
        
        self.mapping_table.resizeRowsToContents()
    
    def create_mapping_row(self, row, excel_col=""):
        """マッピング行作成"""
        # Excelカラム（編集可能）
        excel_item = QTableWidgetItem(excel_col)
        self.mapping_table.setItem(row, 0, excel_item)
        
        # データベースカラム（コンボボックス）
        db_combo = QComboBox()
        db_combo.addItem("", "")  # 空白オプション
        
        # DBカラムを追加（説明付き）
        for col_name in self.db_columns:
            col_info = self.db_columns_info.get(col_name, {})
            description = col_info.get('description', '')
            
            if description:
                display_text = f"{col_name} ({description})"
            else:
                display_text = col_name
            
            db_combo.addItem(display_text, col_name)
        
        self.mapping_table.setCellWidget(row, 1, db_combo)
        
        # 削除ボタン
        delete_btn = QPushButton("削除")
        delete_btn.clicked.connect(lambda: self.delete_mapping_row(row))
        self.mapping_table.setCellWidget(row, 2, delete_btn)
    
    def add_mapping_row(self):
        """マッピング行追加"""
        # Excelカラム名入力
        excel_col, ok = QInputDialog.getText(
            self,
            "Excelカラム追加",
            "Excelカラム名を入力してください:"
        )
        
        if ok and excel_col:
            # 重複チェック
            for i in range(self.mapping_table.rowCount()):
                existing_col = self.mapping_table.item(i, 0).text()
                if existing_col == excel_col:
                    QMessageBox.warning(
                        self,
                        "警告",
                        f"'{excel_col}' は既に存在します"
                    )
                    return
            
            # 新しい行を追加
            row = self.mapping_table.rowCount()
            self.mapping_table.insertRow(row)
            self.create_mapping_row(row, excel_col)
            
            # 削除ボタンの再接続（行番号が変わるため）
            self.reconnect_delete_buttons()
    
    def delete_mapping_row(self, row):
        """マッピング行削除"""
        reply = QMessageBox.question(
            self,
            "確認",
            "この行を削除しますか？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.mapping_table.removeRow(row)
            # 削除ボタンの再接続
            self.reconnect_delete_buttons()
    
    def reconnect_delete_buttons(self):
        """削除ボタンの再接続"""
        for i in range(self.mapping_table.rowCount()):
            delete_btn = self.mapping_table.cellWidget(i, 2)
            if delete_btn:
                # 既存の接続を切断
                try:
                    delete_btn.clicked.disconnect()
                except:
                    pass
                # 新しい接続
                delete_btn.clicked.connect(lambda checked, row=i: self.delete_mapping_row(row))
    
    def clear_all_mappings(self):
        """すべてのマッピングをクリア"""
        reply = QMessageBox.question(
            self,
            "確認",
            "すべてのマッピングをクリアしますか？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            for i in range(self.mapping_table.rowCount()):
                combo = self.mapping_table.cellWidget(i, 1)
                if combo:
                    combo.setCurrentIndex(0)  # 空白に設定
    
    def load_current_mapping(self):
        """現在のマッピング読み込み"""
        try:
            current_mapping = self.config_manager.get_column_mapping(self.data_type)
            
            if not current_mapping:
                return
            
            # マッピング適用
            for i in range(self.mapping_table.rowCount()):
                excel_col = self.mapping_table.item(i, 0).text()
                
                if excel_col in current_mapping:
                    db_col = current_mapping[excel_col]
                    combo = self.mapping_table.cellWidget(i, 1)
                    
                    # コンボボックスから該当する値を探す
                    for j in range(combo.count()):
                        if combo.itemData(j) == db_col:
                            combo.setCurrentIndex(j)
                            break
        
        except Exception as e:
            print(f"マッピング読み込みエラー: {e}")
    
    def get_mapping(self):
        """現在のマッピング取得"""
        mapping = {}
        
        for i in range(self.mapping_table.rowCount()):
            excel_col = self.mapping_table.item(i, 0).text().strip()
            combo = self.mapping_table.cellWidget(i, 1)
            db_col = combo.currentData()
            
            if excel_col and db_col:  # 両方が空白でない場合のみ
                mapping[excel_col] = db_col
        
        return mapping
    
    def validate_mapping(self, mapping):
        """マッピング検証"""
        if not mapping:
            QMessageBox.warning(self, "警告", "マッピングが設定されていません")
            return False
        
        # 必須カラムチェック
        required_columns = self.get_required_columns()
        db_columns = list(mapping.values())
        
        missing_columns = []
        for col in required_columns:
            if col not in db_columns:
                missing_columns.append(col)
        
        if missing_columns:
            QMessageBox.warning(
                self,
                "検証エラー",
                f"必須カラムがマッピングされていません:\n\n" +
                '\n'.join([f"• {col}" for col in missing_columns])
            )
            return False
        
        # 重複チェック
        db_columns_list = [v for v in mapping.values() if v]
        if len(db_columns_list) != len(set(db_columns_list)):
            QMessageBox.warning(
                self,
                "検証エラー",
                "同じデータベースカラムが複数のExcelカラムにマッピングされています"
            )
            return False
        
        return True
    
    def save_mapping(self):
        """マッピング保存"""
        try:
            mapping = self.get_mapping()
            
            # 検証
            if not self.validate_mapping(mapping):
                return
            
            # 保存
            success = self.config_manager.save_column_mapping(self.data_type, mapping)
            
            if success:
                QMessageBox.information(
                    self,
                    "保存完了",
                    f"カラムマッピングを保存しました。\n\n"
                    f"マッピング数: {len(mapping)}個"
                )
                self.accept()
            else:
                QMessageBox.warning(self, "エラー", "マッピングの保存に失敗しました")
        
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"マッピング保存エラー:\n{str(e)}")