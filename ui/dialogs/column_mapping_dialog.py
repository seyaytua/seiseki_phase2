"""
カラムマッピングダイアログ

Excelのカラム名とデータベースのカラム名を対応付ける
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget,
    QTableWidgetItem, QPushButton, QComboBox, QLabel,
    QMessageBox
)
from PySide6.QtCore import Qt

from infrastructure.config_manager import ConfigManager


class ColumnMappingDialog(QDialog):
    """カラムマッピングダイアログ"""
    
    def __init__(self,
                 data_type: str,
                 excel_columns: list,
                 config_manager: ConfigManager,
                 parent=None):
        super().__init__(parent)
        
        self.data_type = data_type
        self.excel_columns = excel_columns
        self.config_manager = config_manager
        self.mapping = {}
        
        self._init_ui()
        self._load_mapping()
    
    def _init_ui(self):
        """UI初期化"""
        self.setWindowTitle(f"カラムマッピング設定 - {self.data_type}")
        self.setMinimumSize(600, 500)
        
        layout = QVBoxLayout(self)
        
        # 説明
        info = QLabel(
            "Excelのカラム名とデータベースのカラム名を対応付けてください\n"
            "※必須カラム: 学籍番号、科目番号"
        )
        info.setWordWrap(True)
        layout.addWidget(info)
        
        # マッピングテーブル
        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(3)
        self.mapping_table.setHorizontalHeaderLabels([
            'Excelカラム名',
            '→',
            'データベースカラム名'
        ])
        self.mapping_table.setColumnWidth(0, 200)
        self.mapping_table.setColumnWidth(1, 30)
        self.mapping_table.setColumnWidth(2, 250)
        
        self._populate_table()
        
        layout.addWidget(self.mapping_table)
        
        # ボタン
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        btn_reset = QPushButton("リセット")
        btn_reset.clicked.connect(self._reset_mapping)
        btn_layout.addWidget(btn_reset)
        
        btn_cancel = QPushButton("キャンセル")
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)
        
        btn_save = QPushButton("保存")
        btn_save.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2ecc71;
            }
        """)
        btn_save.clicked.connect(self.save_mapping)
        btn_layout.addWidget(btn_save)
        
        layout.addLayout(btn_layout)
    
    def _populate_table(self):
        """テーブルにデータを設定"""
        db_columns = self.config_manager.get_db_columns(self.data_type)
        required_columns = self.config_manager.get_required_columns(self.data_type)
        
        self.mapping_table.setRowCount(len(self.excel_columns))
        
        for i, excel_col in enumerate(self.excel_columns):
            # Excelカラム名
            excel_item = QTableWidgetItem(str(excel_col))
            excel_item.setFlags(Qt.ItemIsEnabled)
            self.mapping_table.setItem(i, 0, excel_item)
            
            # 矢印
            arrow_item = QTableWidgetItem("→")
            arrow_item.setFlags(Qt.ItemIsEnabled)
            arrow_item.setTextAlignment(Qt.AlignCenter)
            self.mapping_table.setItem(i, 1, arrow_item)
            
            # データベースカラム選択
            combo = QComboBox()
            combo.addItem("")  # 空白オプション
            
            for db_col in db_columns:
                # 必須カラムには★マーク
                if db_col in required_columns:
                    combo.addItem(f"★ {db_col}", db_col)
                else:
                    combo.addItem(db_col, db_col)
            
            # 自動推測（カラム名が一致する場合）
            excel_col_lower = str(excel_col).lower()
            for db_col in db_columns:
                if db_col in excel_col_lower or excel_col_lower in db_col:
                    index = combo.findData(db_col)
                    if index >= 0:
                        combo.setCurrentIndex(index)
                        break
            
            self.mapping_table.setCellWidget(i, 2, combo)
    
    def _load_mapping(self):
        """保存されたマッピングを読み込み"""
        saved_mapping = self.config_manager.get_column_mapping(self.data_type)
        
        if not saved_mapping:
            return
        
        for i in range(self.mapping_table.rowCount()):
            excel_col = self.mapping_table.item(i, 0).text()
            
            if excel_col in saved_mapping:
                db_col = saved_mapping[excel_col]
                combo = self.mapping_table.cellWidget(i, 2)
                
                index = combo.findData(db_col)
                if index >= 0:
                    combo.setCurrentIndex(index)
    
    def _reset_mapping(self):
        """マッピングをリセット"""
        reply = QMessageBox.question(
            self, "確認",
            "マッピングをリセットしますか？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            for i in range(self.mapping_table.rowCount()):
                combo = self.mapping_table.cellWidget(i, 2)
                combo.setCurrentIndex(0)
    
    def save_mapping(self):
        """マッピングを保存"""
        mapping = {}
        required_columns = self.config_manager.get_required_columns(self.data_type)
        found_required = set()
        
        for i in range(self.mapping_table.rowCount()):
            excel_col = self.mapping_table.item(i, 0).text()
            combo = self.mapping_table.cellWidget(i, 2)
            db_col = combo.currentData()
            
            if db_col:
                mapping[excel_col] = db_col
                
                # 必須カラムチェック
                if db_col in required_columns:
                    found_required.add(db_col)
        
        # 必須カラムの確認
        missing = set(required_columns) - found_required
        if missing:
            QMessageBox.warning(
                self, "警告",
                f"必須カラムが不足しています:\n{', '.join(missing)}"
            )
            return
        
        # 保存
        self.mapping = mapping
        self.config_manager.save_column_mapping(self.data_type, mapping)
        
        self.accept()
    
    def get_mapping(self) -> dict:
        """マッピングを取得"""
        return self.mapping
