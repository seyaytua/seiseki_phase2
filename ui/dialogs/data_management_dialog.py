"""
データ管理ダイアログ

データの削除や出力を行う
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGroupBox,
    QPushButton, QComboBox, QSpinBox, QLabel, QMessageBox
)

from infrastructure.database_manager import DatabaseManager
from infrastructure.logger import Logger


class DataManagementDialog(QDialog):
    """データ管理ダイアログ"""
    
    def __init__(self,
                 db_manager: DatabaseManager,
                 logger: Logger,
                 parent=None):
        super().__init__(parent)
        
        self.db_manager = db_manager
        self.logger = logger
        
        self._init_ui()
    
    def _init_ui(self):
        """UI初期化"""
        self.setWindowTitle("データ管理")
        self.setMinimumSize(500, 400)
        
        layout = QVBoxLayout(self)
        
        # データ削除セクション
        layout.addWidget(self._create_delete_section())
        
        # 閉じるボタン
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        btn_close = QPushButton("閉じる")
        btn_close.clicked.connect(self.accept)
        btn_layout.addWidget(btn_close)
        
        layout.addLayout(btn_layout)
    
    def _create_delete_section(self) -> QGroupBox:
        """データ削除セクション"""
        group = QGroupBox("🗑️ データ削除")
        layout = QVBoxLayout()
        
        # 説明
        info = QLabel(
            "指定した期間・年度のデータを削除します\n"
            "※この操作は取り消せません"
        )
        info.setWordWrap(True)
        info.setStyleSheet("color: #e74c3c; font-weight: bold;")
        layout.addWidget(info)
        
        # データタイプ選択
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("データタイプ:"))
        
        self.type_combo = QComboBox()
        self.type_combo.addItems(['評定', '観点', '欠課情報'])
        type_layout.addWidget(self.type_combo)
        type_layout.addStretch()
        
        layout.addLayout(type_layout)
        
        # 期間選択
        period_layout = QHBoxLayout()
        period_layout.addWidget(QLabel("期間:"))
        
        self.period_combo = QComboBox()
        self.period_combo.addItems(['前期', '後期', '通年'])
        period_layout.addWidget(self.period_combo)
        period_layout.addStretch()
        
        layout.addLayout(period_layout)
        
        # 年度選択
        year_layout = QHBoxLayout()
        year_layout.addWidget(QLabel("年度:"))
        
        self.year_spin = QSpinBox()
        self.year_spin.setRange(2000, 2100)
        self.year_spin.setValue(2024)
        year_layout.addWidget(self.year_spin)
        year_layout.addStretch()
        
        layout.addLayout(year_layout)
        
        # 削除ボタン
        btn_delete = QPushButton("🗑️ 削除実行")
        btn_delete.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                padding: 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        btn_delete.clicked.connect(self.delete_data)
        layout.addWidget(btn_delete)
        
        group.setLayout(layout)
        return group
    
    def delete_data(self):
        """データ削除実行"""
        data_type = self.type_combo.currentText()
        period = self.period_combo.currentText()
        year = self.year_spin.value()
        
        # 確認ダイアログ
        confirm_msg = f"""
        以下のデータを削除します:
        
        データタイプ: {data_type}
        期間: {period}
        年度: {year}
        
        ※この操作は取り消せません
        
        本当に削除しますか？
        """
        
        reply = QMessageBox.warning(
            self, "確認",
            confirm_msg,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        # 削除実行
        try:
            count = self.db_manager.delete_data_by_period(
                data_type, period, year
            )
            
            # ログ記録
            self.logger.log_action(
                Logger.ACTION_DELETE,
                f"{data_type} - {period} {year}年度 - {count}件削除"
            )
            
            QMessageBox.information(
                self, "成功",
                f"{count}件のデータを削除しました"
            )
            
        except Exception as e:
            QMessageBox.critical(
                self, "エラー",
                f"削除エラー: {str(e)}"
            )
