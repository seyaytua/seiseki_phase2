from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                               QLabel, QTableWidget, QTableWidgetItem, QComboBox,
                               QMessageBox, QHeaderView, QInputDialog)
from PySide6.QtCore import Qt
import json
from pathlib import Path


class DBColumnManagerDialog(QDialog):
    """DBカラム設定ダイアログ"""
    
    def __init__(self, data_type, parent=None):
        super().__init__(parent)
        self.data_type = data_type
        self.config_path = Path('config/db_columns.json')
        self.columns_config = self.load_config()
        
        self.setup_ui()
        self.load_columns()
    
    def setup_ui(self):
        """UI初期化"""
        self.setWindowTitle(f"{self.data_type} - DBカラム設定")
        self.setGeometry(200, 150, 800, 600)
        
        layout = QVBoxLayout(self)
        
        info_label = QLabel(
            f"{self.data_type}のデータベースカラムを管理します。\n"
            "カラム名、データ型、説明を設定できます。"
        )
        info_label.setStyleSheet("padding: 10px; background-color: #E3F2FD; border-radius: 5px;")
        layout.addWidget(info_label)
        
        # カラムテーブル
        self.column_table = QTableWidget()
        self.column_table.setColumnCount(3)
        self.column_table.setHorizontalHeaderLabels(["カラム名", "データ型", "説明"])
        self.column_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.column_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.column_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.column_table.setColumnWidth(0, 200)
        
        layout.addWidget(self.column_table)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        add_btn = QPushButton("カラム追加")
        add_btn.clicked.connect(self.add_column)
        button_layout.addWidget(add_btn)
        
        delete_btn = QPushButton("選択行削除")
        delete_btn.clicked.connect(self.delete_column)
        button_layout.addWidget(delete_btn)
        
        button_layout.addStretch()
        
        save_btn = QPushButton("保存")
        save_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 8px;")
        save_btn.clicked.connect(self.save_config_data)
        button_layout.addWidget(save_btn)
        
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
    
    def load_config(self):
        """設定ファイル読み込み"""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"設定ファイルの読み込みに失敗:\n{str(e)}")
                return self.get_default_config()
        else:
            return self.get_default_config()
    
    def get_default_config(self):
        """デフォルト設定取得"""
        return {
            '評定': [
                {'name': 'student_number', 'type': 'TEXT', 'description': '生徒番号'},
                {'name': 'student_name', 'type': 'TEXT', 'description': '生徒氏名'},
                {'name': 'course_number', 'type': 'TEXT', 'description': '講座番号'},
                {'name': 'course_name', 'type': 'TEXT', 'description': '講座名'},
                {'name': 'school_subject_name', 'type': 'TEXT', 'description': '校内科目名'},
                {'name': 'grade_value', 'type': 'INTEGER', 'description': '評定値（1-10）'},
                {'name': 'credits', 'type': 'INTEGER', 'description': '単位数'},
                {'name': 'acquisition_credits', 'type': 'INTEGER', 'description': '修得単位数'},
                {'name': 'remarks', 'type': 'TEXT', 'description': '備考'},
            ],
            '観点': [
                {'name': 'student_number', 'type': 'TEXT', 'description': '生徒番号'},
                {'name': 'student_name', 'type': 'TEXT', 'description': '生徒氏名'},
                {'name': 'course_number', 'type': 'TEXT', 'description': '講座番号'},
                {'name': 'course_name', 'type': 'TEXT', 'description': '講座名'},
                {'name': 'school_subject_name', 'type': 'TEXT', 'description': '校内科目名'},
                {'name': 'viewpoint_1', 'type': 'TEXT', 'description': '観点1評価'},
                {'name': 'viewpoint_2', 'type': 'TEXT', 'description': '観点2評価'},
                {'name': 'viewpoint_3', 'type': 'TEXT', 'description': '観点3評価'},
                {'name': 'viewpoint_4', 'type': 'TEXT', 'description': '観点4評価'},
                {'name': 'viewpoint_5', 'type': 'TEXT', 'description': '観点5評価'},
                {'name': 'remarks', 'type': 'TEXT', 'description': '備考'},
            ],
            '欠課情報': [
                {'name': 'student_number', 'type': 'TEXT', 'description': '生徒番号'},
                {'name': 'student_name', 'type': 'TEXT', 'description': '生徒氏名'},
                {'name': 'course_number', 'type': 'TEXT', 'description': '講座番号'},
                {'name': 'course_name', 'type': 'TEXT', 'description': '講座名'},
                {'name': 'school_subject_name', 'type': 'TEXT', 'description': '校内科目名'},
                {'name': 'absent_hours', 'type': 'INTEGER', 'description': '欠課時数'},
                {'name': 'late_hours', 'type': 'INTEGER', 'description': '遅刻時数'},
                {'name': 'total_hours', 'type': 'INTEGER', 'description': '総時数'},
                {'name': 'absence_rate', 'type': 'REAL', 'description': '欠課率'},
                {'name': 'remarks', 'type': 'TEXT', 'description': '備考'},
            ]
        }
    
    def save_config_data(self):
        """設定を保存"""
        # テーブルから現在の設定を取得
        columns = []
        for i in range(self.column_table.rowCount()):
            name_item = self.column_table.item(i, 0)
            type_combo = self.column_table.cellWidget(i, 1)
            desc_item = self.column_table.item(i, 2)
            
            if name_item:
                columns.append({
                    'name': name_item.text(),
                    'type': type_combo.currentText() if type_combo else 'TEXT',
                    'description': desc_item.text() if desc_item else ''
                })
        
        self.columns_config[self.data_type] = columns
        
        try:
            # configディレクトリが存在することを確認
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.columns_config, f, ensure_ascii=False, indent=2)
            
            QMessageBox.information(
                self,
                "保存完了",
                f"{self.data_type}のDBカラム設定を保存しました。\n\n"
                f"カラム数: {len(columns)}個"
            )
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"保存に失敗しました:\n{str(e)}")
    
    def load_columns(self):
        """カラム一覧を表示"""
        columns = self.columns_config.get(self.data_type, [])
        
        self.column_table.setRowCount(len(columns))
        
        data_types = ['TEXT', 'INTEGER', 'REAL', 'DATE', 'DATETIME', 'BOOLEAN']
        
        for i, col in enumerate(columns):
            # カラム名
            name_item = QTableWidgetItem(col['name'])
            self.column_table.setItem(i, 0, name_item)
            
            # データ型
            type_combo = QComboBox()
            type_combo.addItems(data_types)
            type_combo.setCurrentText(col['type'])
            self.column_table.setCellWidget(i, 1, type_combo)
            
            # 説明
            desc_item = QTableWidgetItem(col.get('description', ''))
            self.column_table.setItem(i, 2, desc_item)
    
    def add_column(self):
        """カラム追加"""
        name, ok = QInputDialog.getText(self, "カラム追加", "カラム名:")
        if ok and name:
            row = self.column_table.rowCount()
            self.column_table.insertRow(row)
            
            # カラム名
            name_item = QTableWidgetItem(name)
            self.column_table.setItem(row, 0, name_item)
            
            # データ型
            type_combo = QComboBox()
            type_combo.addItems(['TEXT', 'INTEGER', 'REAL', 'DATE', 'DATETIME', 'BOOLEAN'])
            self.column_table.setCellWidget(row, 1, type_combo)
            
            # 説明
            desc_item = QTableWidgetItem("")
            self.column_table.setItem(row, 2, desc_item)
    
    def delete_column(self):
        """選択行削除"""
        current_row = self.column_table.currentRow()
        if current_row >= 0:
            reply = QMessageBox.question(
                self,
                "確認",
                "選択したカラムを削除しますか？",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.column_table.removeRow(current_row)
    
    def get_columns(self):
        """現在の設定を取得（カラム名のリスト）"""
        columns = []
        for i in range(self.column_table.rowCount()):
            name_item = self.column_table.item(i, 0)
            if name_item:
                columns.append(name_item.text())
        return columns