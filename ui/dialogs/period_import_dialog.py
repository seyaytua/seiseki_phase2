"""
データ取り込みダイアログ

Excelファイルからデータを取り込むためのダイアログ
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel,
    QPushButton, QComboBox, QSpinBox, QTableWidget, QTableWidgetItem,
    QFileDialog, QProgressDialog, QMessageBox, QCheckBox, QLineEdit
)
from PySide6.QtCore import Qt
import pandas as pd

from infrastructure.database_manager import DatabaseManager
from infrastructure.config_manager import ConfigManager
from infrastructure.file_manager import FileManager
from services.data_import_service import DataImportService


class PeriodImportDialog(QDialog):
    """データ取り込みダイアログ"""
    
    def __init__(self,
                 data_type: str,
                 db_manager: DatabaseManager,
                 config_manager: ConfigManager,
                 file_manager: FileManager,
                 import_service: DataImportService,
                 parent=None):
        super().__init__(parent)
        
        self.data_type = data_type
        self.db_manager = db_manager
        self.config_manager = config_manager
        self.file_manager = file_manager
        self.import_service = import_service
        
        self.file_path = None
        self.sheet_names = []
        self.column_mapping = {}
        
        self._init_ui()
        self._load_saved_mapping()
    
    def _init_ui(self):
        """UI初期化"""
        self.setWindowTitle(f"{self.data_type}データ取り込み")
        self.setMinimumSize(1000, 700)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        
        # ファイル選択セクション
        layout.addWidget(self._create_file_section())
        
        # 期間・年度設定セクション
        layout.addWidget(self._create_period_section())
        
        # シート選択セクション
        layout.addWidget(self._create_sheet_section())
        
        # プレビューセクション
        layout.addWidget(self._create_preview_section())
        
        # ボタン
        layout.addWidget(self._create_button_section())
    
    def _create_file_section(self) -> QGroupBox:
        """ファイル選択セクション"""
        group = QGroupBox("📁 ファイル選択")
        layout = QHBoxLayout()
        
        self.file_label = QLabel("ファイルが選択されていません")
        layout.addWidget(self.file_label)
        
        btn_select = QPushButton("📂 ファイル選択")
        btn_select.clicked.connect(self.select_file)
        layout.addWidget(btn_select)
        
        group.setLayout(layout)
        return group
    
    def _create_period_section(self) -> QGroupBox:
        """期間・年度設定セクション"""
        group = QGroupBox("📅 期間・年度設定")
        layout = QHBoxLayout()
        
        # 期間選択
        layout.addWidget(QLabel("期間:"))
        self.period_combo = QComboBox()
        self.period_combo.addItems(['前期', '後期', '通年'])
        self.period_combo.setCurrentText(
            self.config_manager.get_setting('default_period', '前期')
        )
        layout.addWidget(self.period_combo)
        
        layout.addSpacing(20)
        
        # 年度選択
        layout.addWidget(QLabel("年度:"))
        self.year_spin = QSpinBox()
        self.year_spin.setRange(2000, 2100)
        self.year_spin.setValue(
            self.config_manager.get_setting('default_year', 2024)
        )
        layout.addWidget(self.year_spin)
        
        layout.addSpacing(20)
        
        # ヘッダー行
        layout.addWidget(QLabel("ヘッダー行:"))
        self.header_spin = QSpinBox()
        self.header_spin.setRange(0, 50)
        self.header_spin.setValue(
            self.config_manager.get_default_header_row(self.data_type)
        )
        self.header_spin.valueChanged.connect(self.update_preview)
        layout.addWidget(self.header_spin)
        
        layout.addStretch()
        
        # タイムスタンプチェックボックス
        self.timestamp_check = QCheckBox("ファイル名にタイムスタンプを追加")
        self.timestamp_check.setChecked(True)
        layout.addWidget(self.timestamp_check)
        
        group.setLayout(layout)
        return group
    
    def _create_sheet_section(self) -> QGroupBox:
        """シート選択セクション"""
        group = QGroupBox("📄 シート選択")
        layout = QVBoxLayout()
        
        # 説明
        info = QLabel("取り込むシートを選択してください（複数選択可能）")
        layout.addWidget(info)
        
        # シートテーブル
        self.sheet_table = QTableWidget()
        self.sheet_table.setColumnCount(2)
        self.sheet_table.setHorizontalHeaderLabels(['選択', 'シート名'])
        self.sheet_table.setColumnWidth(0, 50)
        self.sheet_table.setColumnWidth(1, 300)
        self.sheet_table.setMaximumHeight(150)
        layout.addWidget(self.sheet_table)
        
        # 選択ボタン
        btn_layout = QHBoxLayout()
        btn_select_all = QPushButton("すべて選択")
        btn_select_all.clicked.connect(self.select_all_sheets)
        btn_layout.addWidget(btn_select_all)
        
        btn_deselect_all = QPushButton("すべて解除")
        btn_deselect_all.clicked.connect(self.deselect_all_sheets)
        btn_layout.addWidget(btn_deselect_all)
        
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        group.setLayout(layout)
        return group
    
    def _create_preview_section(self) -> QGroupBox:
        """プレビューセクション"""
        group = QGroupBox("👁️ データプレビュー")
        layout = QVBoxLayout()
        
        # カラムマッピングボタン
        btn_mapping = QPushButton("🔧 カラムマッピング設定")
        btn_mapping.clicked.connect(self.show_column_mapping)
        layout.addWidget(btn_mapping)
        
        # プレビューテーブル
        self.preview_table = QTableWidget()
        self.preview_table.setMaximumHeight(200)
        layout.addWidget(self.preview_table)
        
        group.setLayout(layout)
        return group
    
    def _create_button_section(self) -> QWidget:
        """ボタンセクション"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        
        layout.addStretch()
        
        # キャンセルボタン
        btn_cancel = QPushButton("キャンセル")
        btn_cancel.clicked.connect(self.reject)
        layout.addWidget(btn_cancel)
        
        # 取り込み実行ボタン
        self.btn_import = QPushButton("✅ 取り込み実行")
        self.btn_import.setEnabled(False)
        self.btn_import.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 10px 20px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2ecc71;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)
        self.btn_import.clicked.connect(self.execute_import)
        layout.addWidget(self.btn_import)
        
        return widget
    
    def select_file(self):
        """ファイル選択"""
        last_dir = self.config_manager.get_setting('last_import_dir', '')
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Excelファイル選択",
            last_dir,
            "Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.file_path = file_path
            self.file_label.setText(file_path)
            
            # 最後のディレクトリを保存
            from pathlib import Path
            self.config_manager.save_setting(
                'last_import_dir',
                str(Path(file_path).parent)
            )
            
            # ファイル検証
            is_valid, message = self.import_service.validate_file(file_path)
            if not is_valid:
                QMessageBox.warning(self, "エラー", message)
                return
            
            # シート名読み込み
            self.load_sheet_names()
            
            # プレビュー更新
            self.update_preview()
            
            # 取り込みボタン有効化
            self.btn_import.setEnabled(True)
    
    def load_sheet_names(self):
        """シート名読み込み"""
        self.sheet_names = self.import_service.get_sheet_names(self.file_path)
        
        self.sheet_table.setRowCount(len(self.sheet_names))
        
        for i, sheet_name in enumerate(self.sheet_names):
            # チェックボックス
            check_item = QTableWidgetItem()
            check_item.setCheckState(Qt.Checked)
            check_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            self.sheet_table.setItem(i, 0, check_item)
            
            # シート名
            name_item = QTableWidgetItem(sheet_name)
            name_item.setFlags(Qt.ItemIsEnabled)
            self.sheet_table.setItem(i, 1, name_item)
    
    def select_all_sheets(self):
        """すべてのシートを選択"""
        for i in range(self.sheet_table.rowCount()):
            self.sheet_table.item(i, 0).setCheckState(Qt.Checked)
    
    def deselect_all_sheets(self):
        """すべてのシートの選択を解除"""
        for i in range(self.sheet_table.rowCount()):
            self.sheet_table.item(i, 0).setCheckState(Qt.Unchecked)
    
    def get_selected_sheets(self) -> list:
        """選択されたシート名を取得"""
        selected = []
        for i in range(self.sheet_table.rowCount()):
            if self.sheet_table.item(i, 0).checkState() == Qt.Checked:
                selected.append(self.sheet_table.item(i, 1).text())
        return selected
    
    def update_preview(self):
        """プレビュー更新"""
        if not self.file_path or not self.sheet_names:
            return
        
        # 最初のシートをプレビュー
        sheet_name = self.sheet_names[0]
        header_row = self.header_spin.value()
        
        df = self.import_service.preview_data(
            self.file_path,
            sheet_name,
            header_row,
            nrows=10
        )
        
        if df is not None:
            self._display_preview(df)
    
    def _display_preview(self, df: pd.DataFrame):
        """プレビュー表示"""
        self.preview_table.setRowCount(len(df))
        self.preview_table.setColumnCount(len(df.columns))
        self.preview_table.setHorizontalHeaderLabels(list(df.columns))
        
        for i in range(len(df)):
            for j, col in enumerate(df.columns):
                value = df.iloc[i, j]
                item = QTableWidgetItem(str(value) if pd.notna(value) else '')
                self.preview_table.setItem(i, j, item)
        
        # カラム幅調整
        self.preview_table.resizeColumnsToContents()
    
    def show_column_mapping(self):
        """カラムマッピング設定表示"""
        if not self.file_path or not self.sheet_names:
            QMessageBox.warning(
                self, "警告",
                "先にファイルを選択してください"
            )
            return
        
        try:
            from ui.dialogs.column_mapping_dialog import ColumnMappingDialog
            
            # 最初のシートのカラムを取得
            sheet_name = self.sheet_names[0]
            header_row = self.header_spin.value()
            
            df = self.import_service.preview_data(
                self.file_path,
                sheet_name,
                header_row,
                nrows=1
            )
            
            if df is None:
                return
            
            dialog = ColumnMappingDialog(
                data_type=self.data_type,
                excel_columns=list(df.columns),
                config_manager=self.config_manager,
                parent=self
            )
            
            if dialog.exec():
                self.column_mapping = dialog.get_mapping()
                QMessageBox.information(
                    self, "成功",
                    "カラムマッピングを設定しました"
                )
        except ImportError:
            # カラムマッピングダイアログが未実装の場合
            self._simple_column_mapping()
    
    def _simple_column_mapping(self):
        """簡易カラムマッピング（ダイアログ未実装時用）"""
        QMessageBox.information(
            self, "情報",
            "カラムマッピング機能は自動設定されます"
        )
    
    def _load_saved_mapping(self):
        """保存されたマッピングを読み込み"""
        self.column_mapping = self.config_manager.get_column_mapping(self.data_type)
    
    def execute_import(self):
        """取り込み実行"""
        # バリデーション
        selected_sheets = self.get_selected_sheets()
        if not selected_sheets:
            QMessageBox.warning(self, "警告", "シートを選択してください")
            return
        
        if not self.column_mapping:
            reply = QMessageBox.question(
                self, "確認",
                "カラムマッピングが設定されていません。\n"
                "自動マッピングで続行しますか？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.No:
                return
            
            # 自動マッピング（カラム名がそのまま使える場合）
            self.column_mapping = {}
        
        # 確認ダイアログ
        period = self.period_combo.currentText()
        year = self.year_spin.value()
        
        confirm_msg = f"""
        データタイプ: {self.data_type}
        期間: {period}
        年度: {year}
        シート数: {len(selected_sheets)}
        
        取り込みを実行しますか？
        """
        
        reply = QMessageBox.question(
            self, "確認",
            confirm_msg,
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        # 進捗ダイアログ
        progress = QProgressDialog(
            "データ取り込み中...",
            "キャンセル",
            0,
            len(selected_sheets),
            self
        )
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        
        def update_progress(current, total, message):
            progress.setValue(current)
            progress.setLabelText(message)
            if progress.wasCanceled():
                return False
            return True
        
        # インポート実行
        success, message, count = self.import_service.import_data(
            file_path=self.file_path,
            data_type=self.data_type,
            period=period,
            year=year,
            column_mapping=self.column_mapping,
            sheet_names=selected_sheets,
            header_row=self.header_spin.value(),
            progress_callback=update_progress,
            add_timestamp=self.timestamp_check.isChecked()
        )
        
        progress.close()
        
        # 結果表示
        if success:
            QMessageBox.information(self, "成功", message)
            self.accept()
        else:
            QMessageBox.critical(self, "エラー", message)
