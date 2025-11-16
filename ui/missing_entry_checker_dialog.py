from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                               QLabel, QTableWidget, QTableWidgetItem, QHeaderView,
                               QMessageBox, QComboBox, QSpinBox, QGroupBox,
                               QTabWidget)
from PySide6.QtCore import Qt
from pathlib import Path
from datetime import datetime


class MissingEntryCheckerDialog(QDialog):
    """未入力者チェックダイアログ"""
    
    def __init__(self, db_manager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        
        self.setup_ui()
    
    def setup_ui(self):
        """UI初期化"""
        self.setWindowTitle("未入力者チェック")
        self.setGeometry(150, 100, 1200, 700)
        
        layout = QVBoxLayout(self)
        
        # 説明
        info_label = QLabel(
            "評価・評定と観点評価の未入力者を確認します。\n"
            "受講者マスタと照合し、データが登録されていない生徒を抽出します。"
        )
        layout.addWidget(info_label)
        
        # 検索条件エリア
        condition_group = QGroupBox("検索条件")
        condition_layout = QHBoxLayout()
        
        condition_layout.addWidget(QLabel("年度:"))
        self.year_spin = QSpinBox()
        self.year_spin.setMinimum(2000)
        self.year_spin.setMaximum(2100)
        self.year_spin.setValue(datetime.now().year)
        condition_layout.addWidget(self.year_spin)
        
        condition_layout.addWidget(QLabel("期間:"))
        self.period_combo = QComboBox()
        self.period_combo.addItems(['前期', '後期', '通年'])
        condition_layout.addWidget(self.period_combo)
        
        check_btn = QPushButton("チェック実行")
        check_btn.clicked.connect(self.check_missing_entries)
        condition_layout.addWidget(check_btn)
        
        condition_layout.addStretch()
        condition_group.setLayout(condition_layout)
        layout.addWidget(condition_group)
        
        # タブウィジェット
        self.tab_widget = QTabWidget()
        
        # 評定未入力タブ
        self.grade_table = QTableWidget()
        self.tab_widget.addTab(self.grade_table, "評定未入力")
        
        # 観点未入力タブ
        self.viewpoint_table = QTableWidget()
        self.tab_widget.addTab(self.viewpoint_table, "観点未入力")
        
        layout.addWidget(self.tab_widget)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        export_btn = QPushButton("未入力リストをExcel出力")
        export_btn.clicked.connect(self.export_missing_list)
        button_layout.addWidget(export_btn)
        
        button_layout.addStretch()
        
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
    
    def check_missing_entries(self):
        """未入力者チェック実行"""
        year = self.year_spin.value()
        period = self.period_combo.currentText()
        
        try:
            # 評定未入力チェック
            missing_grades = self.check_missing_grades(year, period)
            self.display_missing_data(self.grade_table, missing_grades, "評定")
            
            # 観点未入力チェック
            missing_viewpoints = self.check_missing_viewpoints(year, period)
            self.display_missing_data(self.viewpoint_table, missing_viewpoints, "観点")
            
            # 結果サマリー
            total_missing = len(missing_grades) + len(missing_viewpoints)
            
            if total_missing == 0:
                QMessageBox.information(
                    self,
                    "チェック完了",
                    f"未入力者はいません。\n\n"
                    f"年度: {year}, 期間: {period}"
                )
            else:
                QMessageBox.warning(
                    self,
                    "チェック完了",
                    f"未入力者が見つかりました。\n\n"
                    f"評定未入力: {len(missing_grades)}件\n"
                    f"観点未入力: {len(missing_viewpoints)}件\n\n"
                    f"詳細は各タブで確認してください。"
                )
        
        except Exception as e:
            QMessageBox.critical(
                self,
                "エラー",
                f"チェック中にエラーが発生しました:\n{str(e)}"
            )
    
    def check_missing_grades(self, year, period):
        """評定未入力をチェック"""
        query = """ SELECT e.course_number, e.course_name, e.student_number, e.student_name FROM enrollments e LEFT JOIN grades g ON e.student_number = g.student_number AND e.course_number = g.course_number AND g.year = ? AND g.period = ? WHERE g.id IS NULL ORDER BY e.course_number, e.student_number """
        
        results = self.db_manager.fetch_all(query, (year, period))
        return [dict(row) for row in results]
    
    def check_missing_viewpoints(self, year, period):
        """観点未入力をチェック"""
        query = """ SELECT e.course_number, e.course_name, e.student_number, e.student_name FROM enrollments e LEFT JOIN viewpoint_evaluations v ON e.student_number = v.student_number AND e.course_number = v.course_number AND v.year = ? AND v.period = ? WHERE v.id IS NULL ORDER BY e.course_number, e.student_number """
        
        results = self.db_manager.fetch_all(query, (year, period))
        return [dict(row) for row in results]
    
    def display_missing_data(self, table, data, data_type):
        """未入力データを表示"""
        if not data:
            table.clear()
            table.setRowCount(1)
            table.setColumnCount(1)
            table.setHorizontalHeaderLabels(["結果"])
            table.setItem(0, 0, QTableWidgetItem(f"{data_type}の未入力者はいません"))
            return
        
        table.clear()
        table.setRowCount(len(data))
        table.setColumnCount(len(data[0].keys()))
        table.setHorizontalHeaderLabels(data[0].keys())
        
        for i, row in enumerate(data):
            for j, (key, value) in enumerate(row.items()):
                item = QTableWidgetItem(str(value) if value is not None else '')
                table.setItem(i, j, item)
        
        table.resizeColumnsToContents()
    
    def export_missing_list(self):
        """未入力リストをExcel出力"""
        year = self.year_spin.value()
        period = self.period_combo.currentText()
        
        try:
            missing_grades = self.check_missing_grades(year, period)
            missing_viewpoints = self.check_missing_viewpoints(year, period)
            
            if not missing_grades and not missing_viewpoints:
                QMessageBox.information(self, "情報", "出力するデータがありません")
                return
            
            from utils.excel_exporter import ExcelExporter
            exporter = ExcelExporter()
            
            data_dict = {}
            
            if missing_grades:
                columns = list(missing_grades[0].keys())
                data_dict['評定未入力'] = (missing_grades, columns)
            
            if missing_viewpoints:
                columns = list(missing_viewpoints[0].keys())
                data_dict['観点未入力'] = (missing_viewpoints, columns)
            
            export_path = exporter.export_multiple_sheets(
                data_dict=data_dict,
                filename=f"未入力者リスト_{year}_{period}.xlsx"
            )
            
            reply = QMessageBox.information(
                self,
                "出力完了",
                f"未入力リストを出力しました。\n\n"
                f"ファイル: {Path(export_path).name}\n\n"
                f"出力先フォルダを開きますか？",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                import os
                export_dir = Path(export_path).parent
                
                if os.name == 'nt':
                    os.startfile(export_dir)
        
        except Exception as e:
            QMessageBox.critical(
                self,
                "エラー",
                f"Excel出力に失敗:\n{str(e)}"
            )