from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                               QLabel, QLineEdit, QFileDialog, QMessageBox,
                               QRadioButton, QButtonGroup, QGroupBox)
from PySide6.QtCore import Qt
from pathlib import Path


class DatabaseSelectorDialog(QDialog):
    """データベース選択ダイアログ"""
    
    def __init__(self, current_db_path=None, parent=None):
        super().__init__(parent)
        self.selected_db_path = current_db_path
        
        self.setup_ui()
        
        if current_db_path:
            self.existing_path_edit.setText(str(current_db_path))
            self.existing_radio.setChecked(True)
    
    def setup_ui(self):
        """UI初期化"""
        self.setWindowTitle("データベース選択")
        self.setGeometry(300, 200, 700, 250)
        
        layout = QVBoxLayout(self)
        
        info_label = QLabel(
            "使用するデータベースを選択してください。\n"
            "マスタ管理アプリと同じデータベースを選択することで、\n"
            "生徒マスタや講座マスタなどの情報を共有できます。"
        )
        layout.addWidget(info_label)
        
        # 選択グループ
        select_group = QGroupBox("データベース選択")
        select_layout = QVBoxLayout()
        
        self.mode_group = QButtonGroup()
        
        # 既存選択
        existing_layout = QHBoxLayout()
        self.existing_radio = QRadioButton("既存のデータベースを選択")
        self.existing_radio.setChecked(True)
        self.mode_group.addButton(self.existing_radio)
        existing_layout.addWidget(self.existing_radio)
        
        self.existing_path_edit = QLineEdit()
        self.existing_path_edit.setPlaceholderText("既存のデータベースファイルを選択")
        existing_layout.addWidget(self.existing_path_edit)
        
        existing_browse_btn = QPushButton("参照")
        existing_browse_btn.clicked.connect(self.browse_existing_database)
        existing_layout.addWidget(existing_browse_btn)
        
        select_layout.addLayout(existing_layout)
        
        # 新規作成
        new_layout = QHBoxLayout()
        self.new_radio = QRadioButton("新規作成")
        self.mode_group.addButton(self.new_radio)
        new_layout.addWidget(self.new_radio)
        
        self.new_path_edit = QLineEdit()
        self.new_path_edit.setPlaceholderText("data/database.db")
        self.new_path_edit.setText("data/database.db")
        new_layout.addWidget(self.new_path_edit)
        
        new_browse_btn = QPushButton("参照")
        new_browse_btn.clicked.connect(self.browse_new_database)
        new_layout.addWidget(new_browse_btn)
        
        select_layout.addLayout(new_layout)
        
        select_group.setLayout(select_layout)
        layout.addWidget(select_group)
        
        layout.addStretch()
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(self.accept_selection)
        button_layout.addWidget(ok_btn)
        
        cancel_btn = QPushButton("キャンセル")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
    
    def browse_existing_database(self):
        """既存データベースファイルを選択"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "既存データベースファイルを選択",
            "",
            "Database Files (*.db)"
        )
        
        if file_path:
            self.existing_path_edit.setText(file_path)
            self.existing_radio.setChecked(True)
    
    def browse_new_database(self):
        """新規データベースファイルの保存先を選択"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "新規データベースファイル",
            "data/database.db",
            "Database Files (*.db)"
        )
        
        if file_path:
            self.new_path_edit.setText(file_path)
            self.new_radio.setChecked(True)
    
    def accept_selection(self):
        """選択を確定"""
        if self.existing_radio.isChecked():
            db_path = self.existing_path_edit.text().strip()
            if not db_path:
                QMessageBox.warning(self, "警告", "既存のデータベースファイルを選択してください")
                return
            
            if not Path(db_path).exists():
                QMessageBox.warning(self, "警告", f"ファイル「{db_path}」が見つかりません")
                return
        else:
            db_path = self.new_path_edit.text().strip()
            if not db_path:
                QMessageBox.warning(self, "警告", "データベースファイル名を入力してください")
                return
            
            if Path(db_path).exists():
                reply = QMessageBox.question(
                    self,
                    "確認",
                    f"ファイル「{db_path}」は既に存在します。\n"
                    "上書きしますか？",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.No:
                    return
        
        self.selected_db_path = db_path
        self.accept()
    
    def get_selected_path(self):
        """選択されたデータベースパスを取得"""
        return self.selected_db_path