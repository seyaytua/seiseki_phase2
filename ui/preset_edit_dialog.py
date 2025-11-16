from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                               QLabel, QLineEdit, QSpinBox, QMessageBox,
                               QGroupBox, QTableWidget, QTableWidgetItem,
                               QHeaderView, QComboBox)
from PySide6.QtCore import Qt
import json
from pathlib import Path


class PresetEditDialog(QDialog):
    """プリセット編集ダイアログ"""
    
    def __init__(self, data_type, preset_name, preset_manager, parent=None):
        super().__init__(parent)
        self.data_type = data_type
        self.old_preset_name = preset_name
        self.preset_manager = preset_manager
        self.preset_name = ""
        
        self.load_db_columns_info()
        
        self.setup_ui()
        
        if preset_name:
            self.load_preset()
    
    def load_db_columns_info(self):
        """DBカラムの情報を読み込む"""
        config_path = Path('config/db_columns.json')
        
        self.db_columns = []
        self.db_columns_info = {}
        
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    columns_config = json.load(f)
                    
                    if self.data_type in columns_config:
                        for col in columns_config[self.data_type]:
                            self.db_columns.append(col['name'])
                            self.db_columns_info[col['name']] = col.get('description', '')
            except Exception as e:
                QMessageBox.warning(self, "警告", f"DBカラム情報の読み込みに失敗:\n{str(e)}")
    
    def setup_ui(self):
        """UI初期化"""
        title = "プリセット編集" if self.old_preset_name else "新規プリセット作成"
        self.setWindowTitle(f"{self.data_type} - {title}")
        self.setGeometry(200, 150, 900, 700)
        
        layout = QVBoxLayout(self)
        
        # 基本設定
        basic_group = QGroupBox("基本設定")
        basic_layout = QVBoxLayout()
        
        # プリセット名
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("プリセット名:"))
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("例: 標準設定")
        name_layout.addWidget(self.name_edit)
        basic_layout.addLayout(name_layout)
        
        # ヘッダ行
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("ヘッダ行:"))
        self.header_spin = QSpinBox()
        self.header_spin.setMinimum(0)
        self.header_spin.setMaximum(10)
        self.header_spin.setValue(0)
        self.header_spin.setToolTip("0 = 1行目がヘッダー")
        header_layout.addWidget(self.header_spin)
        header_layout.addWidget(QLabel("行目（0始まり）"))
        header_layout.addStretch()
        basic_layout.addLayout(header_layout)
        
        basic_group.setLayout(basic_layout)
        layout.addWidget(basic_group)
        
        # カラムマッピング設定
        mapping_group = QGroupBox("カラムマッピング設定")
        mapping_layout = QVBoxLayout()
        
        info_label = QLabel(
            "このプリセットで使用するカラムマッピングを設定します。\n"
            "Excelカラム名は実際の取り込み時に自動的にマッピングされます。"
        )
        info_label.setStyleSheet("color: #666; font-size: 9pt;")
        mapping_layout.addWidget(info_label)
        
        # マッピングテーブル
        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(3)
        self.mapping_table.setHorizontalHeaderLabels(["Excelカラム名（例）", "→", "DBカラム"])
        self.mapping_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.mapping_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.mapping_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        
        mapping_layout.addWidget(self.mapping_table)
        
        # ボタン
        mapping_btn_layout = QHBoxLayout()
        add_btn = QPushButton("マッピング追加")
        add_btn.clicked.connect(self.add_mapping)
        mapping_btn_layout.addWidget(add_btn)
        
        remove_btn = QPushButton("選択行削除")
        remove_btn.clicked.connect(self.remove_mapping)
        mapping_btn_layout.addWidget(remove_btn)
        mapping_btn_layout.addStretch()
        
        mapping_layout.addLayout(mapping_btn_layout)
        
        mapping_group.setLayout(mapping_layout)
        layout.addWidget(mapping_group)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        save_btn = QPushButton("保存")
        save_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        save_btn.clicked.connect(self.save_preset)
        button_layout.addWidget(save_btn)
        
        cancel_btn = QPushButton("キャンセル")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
    
    def load_preset(self):
        """既存のプリセットを読み込む"""
        preset_data = self.preset_manager.get_preset(self.old_preset_name)
        if preset_data:
            self.name_edit.setText(self.old_preset_name)
            self.header_spin.setValue(preset_data['header_row'])
            
            # カラムマッピングを表示
            if preset_data.get('column_mapping'):
                self.display_mapping(preset_data['column_mapping'])
    
    def display_mapping(self, mapping):
        """カラムマッピングを表示"""
        self.mapping_table.setRowCount(len(mapping))
        
        for i, (excel_col, db_col) in enumerate(mapping.items()):
            # Excelカラム名
            excel_item = QTableWidgetItem(excel_col)
            self.mapping_table.setItem(i, 0, excel_item)
            
            # 矢印
            arrow_item = QTableWidgetItem("→")
            arrow_item.setFlags(arrow_item.flags() & ~Qt.ItemIsEditable)
            arrow_item.setTextAlignment(Qt.AlignCenter)
            self.mapping_table.setItem(i, 1, arrow_item)
            
            # DBカラム（コンボボックス）
            combo = QComboBox()
            combo.addItem("-- 選択してください --", None)
            for db_col_name in self.db_columns:
                description = self.db_columns_info.get(db_col_name, '')
                if description:
                    combo.addItem(f"{db_col_name} ({description})", db_col_name)
                else:
                    combo.addItem(db_col_name, db_col_name)
            
            # 現在の値を選択
            for j in range(combo.count()):
                if combo.itemData(j) == db_col:
                    combo.setCurrentIndex(j)
                    break
            
            self.mapping_table.setCellWidget(i, 2, combo)
    
    def add_mapping(self):
        """マッピング行を追加"""
        row = self.mapping_table.rowCount()
        self.mapping_table.insertRow(row)
        
        # Excelカラム名（編集可能）
        excel_item = QTableWidgetItem("")
        excel_item.setPlaceholderText("Excelのカラム名")
        self.mapping_table.setItem(row, 0, excel_item)
        
        # 矢印
        arrow_item = QTableWidgetItem("→")
        arrow_item.setFlags(arrow_item.flags() & ~Qt.ItemIsEditable)
        arrow_item.setTextAlignment(Qt.AlignCenter)
        self.mapping_table.setItem(row, 1, arrow_item)
        
        # DBカラム（コンボボックス）
        combo = QComboBox()
        combo.addItem("-- 選択してください --", None)
        for db_col_name in self.db_columns:
            description = self.db_columns_info.get(db_col_name, '')
            if description:
                combo.addItem(f"{db_col_name} ({description})", db_col_name)
            else:
                combo.addItem(db_col_name, db_col_name)
        
        self.mapping_table.setCellWidget(row, 2, combo)
    
    def remove_mapping(self):
        """選択行を削除"""
        current_row = self.mapping_table.currentRow()
        if current_row >= 0:
            self.mapping_table.removeRow(current_row)
    
    def get_mapping(self):
        """現在のマッピングを取得"""
        mapping = {}
        
        for i in range(self.mapping_table.rowCount()):
            excel_item = self.mapping_table.item(i, 0)
            combo = self.mapping_table.cellWidget(i, 2)
            
            if excel_item and combo:
                excel_col = excel_item.text().strip()
                db_col = combo.currentData()
                
                if excel_col and db_col:
                    mapping[excel_col] = db_col
        
        return mapping
    
    def save_preset(self):
        """プリセットを保存"""
        preset_name = self.name_edit.text().strip()
        if not preset_name:
            QMessageBox.warning(self, "警告", "プリセット名を入力してください")
            return
        
        # 名前変更の場合、重複チェック
        if self.old_preset_name != preset_name:
            existing_preset = self.preset_manager.get_preset(preset_name)
            if existing_preset:
                QMessageBox.warning(self, "警告", "同じ名前のプリセットが既に存在します")
                return
        
        header_row = self.header_spin.value()
        mapping = self.get_mapping()
        
        if not mapping:
            reply = QMessageBox.question(
                self,
                "確認",
                "カラムマッピングが設定されていません。\n"
                "このまま保存しますか？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.No:
                return
        
        try:
            # 古いプリセット名を削除（名前変更の場合）
            if self.old_preset_name and self.old_preset_name != preset_name:
                old_preset = self.preset_manager.get_preset(self.old_preset_name)
                if old_preset:
                    # 削除処理はPresetManagerDialog側で行う
                    pass
            
            # 保存
            success = self.preset_manager.save_preset(
                preset_name,
                header_row,
                mapping
            )
            
            if success:
                self.preset_name = preset_name
                self.accept()
            else:
                QMessageBox.critical(self, "エラー", "保存に失敗しました")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"保存に失敗しました:\n{str(e)}")