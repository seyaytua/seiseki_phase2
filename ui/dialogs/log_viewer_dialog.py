"""
ログビューアダイアログ

システムの操作履歴を表示
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget,
    QTableWidgetItem, QPushButton, QLineEdit, QLabel,
    QComboBox
)
from PySide6.QtCore import Qt

from infrastructure.logger import Logger


class LogViewerDialog(QDialog):
    """ログビューアダイアログ"""
    
    def __init__(self, logger: Logger, parent=None):
        super().__init__(parent)
        
        self.logger = logger
        self._init_ui()
        self.load_logs()
    
    def _init_ui(self):
        """UI初期化"""
        self.setWindowTitle("ログビューア")
        self.setMinimumSize(900, 600)
        
        layout = QVBoxLayout(self)
        
        # フィルターバー
        filter_layout = QHBoxLayout()
        
        filter_layout.addWidget(QLabel("操作種別:"))
        self.type_combo = QComboBox()
        self.type_combo.addItem("すべて", None)
        self.type_combo.currentIndexChanged.connect(self.load_logs)
        filter_layout.addWidget(self.type_combo)
        
        filter_layout.addWidget(QLabel("検索:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("キーワードで検索...")
        self.search_input.textChanged.connect(self.load_logs)
        filter_layout.addWidget(self.search_input)
        
        btn_clear = QPushButton("クリア")
        btn_clear.clicked.connect(self.clear_filter)
        filter_layout.addWidget(btn_clear)
        
        layout.addLayout(filter_layout)
        
        # テーブル
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels([
            '日時', '操作種別', '詳細', 'ユーザー'
        ])
        self.table.setColumnWidth(0, 180)
        self.table.setColumnWidth(1, 120)
        self.table.setColumnWidth(2, 400)
        self.table.setColumnWidth(3, 100)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.table)
        
        # 件数表示
        self.count_label = QLabel()
        layout.addWidget(self.count_label)
        
        # ボタン
        btn_layout = QHBoxLayout()
        
        btn_refresh = QPushButton("🔄 更新")
        btn_refresh.clicked.connect(self.load_logs)
        btn_layout.addWidget(btn_refresh)
        
        btn_layout.addStretch()
        
        btn_close = QPushButton("閉じる")
        btn_close.clicked.connect(self.accept)
        btn_layout.addWidget(btn_close)
        
        layout.addLayout(btn_layout)
        
        # 操作種別一覧を取得
        self._load_action_types()
    
    def _load_action_types(self):
        """操作種別一覧を読み込み"""
        action_types = self.logger.get_action_types()
        
        for action_type in action_types:
            self.type_combo.addItem(action_type, action_type)
    
    def load_logs(self):
        """ログ読み込み"""
        keyword = self.search_input.text().strip()
        action_type = self.type_combo.currentData()
        
        logs = self.logger.search_logs(
            keyword=keyword if keyword else None,
            action_type=action_type,
            limit=500
        )
        
        self.table.setRowCount(len(logs))
        
        for i, (log_id, timestamp, action_type, details, user_name) in enumerate(logs):
            self.table.setItem(i, 0, QTableWidgetItem(str(timestamp)))
            self.table.setItem(i, 1, QTableWidgetItem(action_type))
            self.table.setItem(i, 2, QTableWidgetItem(details or ''))
            self.table.setItem(i, 3, QTableWidgetItem(user_name or ''))
        
        self.count_label.setText(f"総件数: {len(logs)}件")
    
    def clear_filter(self):
        """フィルタークリア"""
        self.type_combo.setCurrentIndex(0)
        self.search_input.clear()
        self.load_logs()
