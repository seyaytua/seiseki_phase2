"""
科目一覧ダイアログ

登録されている科目の一覧を表示
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget,
    QTableWidgetItem, QPushButton, QLineEdit, QLabel
)
from PySide6.QtCore import Qt

from infrastructure.database_manager import DatabaseManager


class CourseListDialog(QDialog):
    """科目一覧ダイアログ"""
    
    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        
        self.db_manager = db_manager
        self._init_ui()
        self.load_courses()
    
    def _init_ui(self):
        """UI初期化"""
        self.setWindowTitle("科目一覧")
        self.setMinimumSize(700, 500)
        
        layout = QVBoxLayout(self)
        
        # 検索バー
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("検索:"))
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("科目番号、科目名、教科名で検索...")
        self.search_input.textChanged.connect(self.search_courses)
        search_layout.addWidget(self.search_input)
        
        btn_clear = QPushButton("クリア")
        btn_clear.clicked.connect(self.clear_search)
        search_layout.addWidget(btn_clear)
        
        layout.addLayout(search_layout)
        
        # テーブル
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['科目番号', '科目名', '教科名'])
        self.table.setColumnWidth(0, 150)
        self.table.setColumnWidth(1, 250)
        self.table.setColumnWidth(2, 200)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.table)
        
        # 件数表示
        self.count_label = QLabel()
        layout.addWidget(self.count_label)
        
        # 閉じるボタン
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        btn_close = QPushButton("閉じる")
        btn_close.clicked.connect(self.accept)
        btn_layout.addWidget(btn_close)
        
        layout.addLayout(btn_layout)
    
    def load_courses(self, search_term: str = None):
        """科目一覧読み込み"""
        courses = self.db_manager.get_courses(search_term)
        
        self.table.setRowCount(len(courses))
        
        for i, (course_number, course_name, subject_name) in enumerate(courses):
            self.table.setItem(i, 0, QTableWidgetItem(course_number))
            self.table.setItem(i, 1, QTableWidgetItem(course_name or ''))
            self.table.setItem(i, 2, QTableWidgetItem(subject_name or ''))
        
        self.count_label.setText(f"総件数: {len(courses)}件")
    
    def search_courses(self):
        """科目検索"""
        search_term = self.search_input.text().strip()
        self.load_courses(search_term if search_term else None)
    
    def clear_search(self):
        """検索クリア"""
        self.search_input.clear()
        self.load_courses()
