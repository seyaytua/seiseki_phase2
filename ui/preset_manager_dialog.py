from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                               QLabel, QListWidget, QMessageBox, QInputDialog,
                               QListWidgetItem, QGroupBox)
from PySide6.QtCore import Qt
import json
from pathlib import Path


class PresetManagerDialog(QDialog):
    """プリセット管理ダイアログ"""
    
    def __init__(self, data_type, config_manager, parent=None):
        super().__init__(parent)
        self.data_type = data_type
        self.config_manager = config_manager
        self.presets_path = Path('config/presets.json')
        self.presets = self.load_presets()
        
        self.setup_ui()
        self.load_preset_list()
    
    def setup_ui(self):
        """UI初期化"""
        self.setWindowTitle(f"{self.data_type} - プリセット管理")
        self.setGeometry(200, 150, 800, 600)
        
        layout = QVBoxLayout(self)
        
        # 説明
        info_label = QLabel(
            f"{self.data_type}の取り込みプリセットを管理します。\n"
            "プリセットには、ヘッダ行の位置とカラムマッピング設定が保存されます。"
        )
        info_label.setStyleSheet("padding: 10px; background-color: #E3F2FD; border-radius: 5px;")
        layout.addWidget(info_label)
        
        # プリセットリスト
        preset_group = QGroupBox("保存済みプリセット")
        preset_layout = QVBoxLayout()
        
        self.preset_list = QListWidget()
        self.preset_list.itemDoubleClicked.connect(self.edit_preset)
        preset_layout.addWidget(self.preset_list)
        
        preset_group.setLayout(preset_layout)
        layout.addWidget(preset_group)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        new_btn = QPushButton("新規プリセット作成")
        new_btn.clicked.connect(self.create_new_preset)
        button_layout.addWidget(new_btn)
        
        edit_btn = QPushButton("編集")
        edit_btn.clicked.connect(self.edit_preset)
        button_layout.addWidget(edit_btn)
        
        rename_btn = QPushButton("名前変更")
        rename_btn.clicked.connect(self.rename_preset)
        button_layout.addWidget(rename_btn)
        
        delete_btn = QPushButton("削除")
        delete_btn.clicked.connect(self.delete_preset)
        button_layout.addWidget(delete_btn)
        
        button_layout.addStretch()
        
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
    
    def load_presets(self):
        """プリセットファイル読み込み"""
        if self.presets_path.exists():
            try:
                with open(self.presets_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"プリセットファイルの読み込みに失敗:\n{str(e)}")
                return {}
        else:
            return {}
    
    def save_presets(self):
        """プリセットファイル保存"""
        try:
            self.presets_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.presets_path, 'w', encoding='utf-8') as f:
                json.dump(self.presets, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"プリセットファイルの保存に失敗:\n{str(e)}")
            return False
    
    def load_preset_list(self):
        """プリセット一覧を読み込む"""
        self.preset_list.clear()
        
        data_type_presets = self.presets.get(self.data_type, {})
        
        if not data_type_presets:
            item = QListWidgetItem("（プリセットがありません）")
            item.setForeground(Qt.gray)
            self.preset_list.addItem(item)
            return
        
        for preset_name, preset_data in data_type_presets.items():
            mapping_count = len(preset_data.get('column_mapping', {}))
            item_text = (
                f"{preset_name}\n"
                f" ヘッダ行: {preset_data['header_row']}行目 | "
                f"マッピング: {mapping_count}カラム"
            )
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, preset_name)
            self.preset_list.addItem(item)
    
    def create_new_preset(self):
        """新規プリセット作成"""
        from ui.preset_edit_dialog import PresetEditDialog
        
        dialog = PresetEditDialog(self.data_type, None, self, self)
        if dialog.exec():
            self.load_preset_list()
            QMessageBox.information(
                self,
                "作成完了",
                f"新しいプリセット「{dialog.preset_name}」を作成しました。"
            )
    
    def edit_preset(self):
        """プリセットを編集"""
        current_item = self.preset_list.currentItem()
        if not current_item or current_item.data(Qt.UserRole) is None:
            QMessageBox.warning(self, "警告", "プリセットを選択してください")
            return
        
        preset_name = current_item.data(Qt.UserRole)
        
        from ui.preset_edit_dialog import PresetEditDialog
        
        dialog = PresetEditDialog(self.data_type, preset_name, self, self)
        if dialog.exec():
            self.load_preset_list()
            QMessageBox.information(
                self,
                "更新完了",
                f"プリセット「{dialog.preset_name}」を更新しました。"
            )
    
    def rename_preset(self):
        """プリセット名を変更"""
        current_item = self.preset_list.currentItem()
        if not current_item or current_item.data(Qt.UserRole) is None:
            QMessageBox.warning(self, "警告", "プリセットを選択してください")
            return
        
        old_name = current_item.data(Qt.UserRole)
        new_name, ok = QInputDialog.getText(
            self,
            "名前変更",
            "新しい名前:",
            text=old_name
        )
        
        if ok and new_name:
            if new_name in self.presets.get(self.data_type, {}):
                QMessageBox.warning(self, "警告", "同じ名前のプリセットが既に存在します")
                return
            
            # 名前変更
            if self.data_type not in self.presets:
                self.presets[self.data_type] = {}
            
            preset_data = self.presets[self.data_type].pop(old_name)
            self.presets[self.data_type][new_name] = preset_data
            
            if self.save_presets():
                self.load_preset_list()
                QMessageBox.information(self, "成功", "名前を変更しました")
    
    def delete_preset(self):
        """プリセットを削除"""
        current_item = self.preset_list.currentItem()
        if not current_item or current_item.data(Qt.UserRole) is None:
            QMessageBox.warning(self, "警告", "プリセットを選択してください")
            return
        
        preset_name = current_item.data(Qt.UserRole)
        
        reply = QMessageBox.question(
            self,
            "確認",
            f"プリセット「{preset_name}」を削除しますか？\n\n"
            "この操作は取り消せません。",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            if self.data_type in self.presets and preset_name in self.presets[self.data_type]:
                del self.presets[self.data_type][preset_name]
                
                if self.save_presets():
                    self.load_preset_list()
                    QMessageBox.information(self, "成功", "プリセットを削除しました")
    
    def get_preset(self, preset_name):
        """プリセットデータ取得"""
        return self.presets.get(self.data_type, {}).get(preset_name)
    
    def save_preset(self, preset_name, header_row, column_mapping):
        """プリセット保存"""
        if self.data_type not in self.presets:
            self.presets[self.data_type] = {}
        
        self.presets[self.data_type][preset_name] = {
            'header_row': header_row,
            'column_mapping': column_mapping
        }
        
        return self.save_presets()