from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                               QLabel, QListWidget, QMessageBox, QListWidgetItem,
                               QGroupBox)
from PySide6.QtCore import Qt
import json
from pathlib import Path


class RequiredColumnsManagerDialog(QDialog):
    """必須カラム管理ダイアログ"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.config_path = Path('config/required_columns.json')
        self.config = self.load_config()
        
        self.setup_ui()
        self.load_data_types()
    
    def setup_ui(self):
        """UI初期化"""
        self.setWindowTitle("必須カラム管理")
        self.setGeometry(250, 150, 900, 600)
        
        layout = QVBoxLayout(self)
        
        info_label = QLabel(
            "各データタイプの必須カラムを管理します。\n"
            "チェックを外すと、そのカラムは必須ではなくなります。"
        )
        info_label.setStyleSheet("padding: 10px; background-color: #E3F2FD; border-radius: 5px;")
        layout.addWidget(info_label)
        
        # メインエリア
        main_layout = QHBoxLayout()
        
        # 左側: データタイプ選択
        left_group = QGroupBox("データタイプ")
        left_layout = QVBoxLayout()
        
        self.type_list = QListWidget()
        self.type_list.currentItemChanged.connect(self.on_type_selected)
        left_layout.addWidget(self.type_list)
        
        left_group.setLayout(left_layout)
        main_layout.addWidget(left_group, 1)
        
        # 右側: カラム一覧
        right_group = QGroupBox("必須カラム")
        right_layout = QVBoxLayout()
        
        help_label = QLabel(
            "チェックを外すと必須カラムから除外されます"
        )
        help_label.setStyleSheet("color: #666; font-size: 9pt;")
        right_layout.addWidget(help_label)
        
        self.column_list = QListWidget()
        right_layout.addWidget(self.column_list)
        
        right_group.setLayout(right_layout)
        main_layout.addWidget(right_group, 2)
        
        layout.addLayout(main_layout)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        self.save_btn = QPushButton("保存")
        self.save_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 8px;")
        self.save_btn.clicked.connect(self.save_config_data)
        button_layout.addWidget(self.save_btn)
        
        button_layout.addStretch()
        
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
            "評定": ["student_number", "course_number"],
            "観点": ["student_number", "course_number"],
            "欠課情報": ["student_number", "course_number"]
        }
    
    def load_data_types(self):
        """データタイプ一覧を読み込む"""
        self.type_list.clear()
        
        data_types = ["評定", "観点", "欠課情報"]
        
        for data_type in data_types:
            item = QListWidgetItem(data_type)
            item.setData(Qt.UserRole, data_type)
            self.type_list.addItem(item)
        
        # 最初の項目を選択
        if self.type_list.count() > 0:
            self.type_list.setCurrentRow(0)
    
    def on_type_selected(self, current, previous):
        """データタイプ選択時の処理"""
        if not current:
            return
        
        data_type = current.data(Qt.UserRole)
        self.load_columns(data_type)
    
    def load_columns(self, data_type):
        """カラム一覧を読み込む"""
        self.column_list.clear()
        
        # カラム情報を読み込む
        columns_config_path = Path('config/db_columns.json')
        
        if not columns_config_path.exists():
            QMessageBox.warning(
                self,
                "警告",
                "DBカラム設定ファイルが見つかりません。\n"
                "先にDBカラム設定を行ってください。"
            )
            return
        
        try:
            with open(columns_config_path, 'r', encoding='utf-8') as f:
                columns_config = json.load(f)
            
            if data_type not in columns_config:
                QMessageBox.warning(
                    self,
                    "警告",
                    f"{data_type}のDBカラム設定が見つかりません。"
                )
                return
            
            # 必須カラムリスト
            required_columns = self.config.get(data_type, [])
            
            # カラム一覧を表示
            for col_info in columns_config[data_type]:
                col_name = col_info['name']
                description = col_info.get('description', '')
                
                # リストアイテム作成
                display_text = f"{col_name}"
                if description:
                    display_text += f" ({description})"
                
                item = QListWidgetItem(display_text)
                item.setData(Qt.UserRole, col_name)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                
                # 必須カラムかどうかでチェック状態を設定
                if col_name in required_columns:
                    item.setCheckState(Qt.Checked)
                else:
                    item.setCheckState(Qt.Unchecked)
                
                self.column_list.addItem(item)
        
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"カラム情報の取得に失敗:\n{str(e)}")
    
    def save_config_data(self):
        """設定を保存"""
        current_item = self.type_list.currentItem()
        if not current_item:
            return
        
        data_type = current_item.data(Qt.UserRole)
        
        # チェックされているカラムを取得
        required_columns = []
        for i in range(self.column_list.count()):
            item = self.column_list.item(i)
            if item.checkState() == Qt.Checked:
                col_name = item.data(Qt.UserRole)
                required_columns.append(col_name)
        
        # 設定を更新
        self.config[data_type] = required_columns
        
        # ファイルに保存
        try:
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            
            QMessageBox.information(
                self,
                "保存完了",
                f"{data_type}の必須カラム設定を保存しました。\n\n"
                f"必須カラム: {len(required_columns)}個"
            )
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"保存に失敗:\n{str(e)}")