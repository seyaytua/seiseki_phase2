"""
メインウィンドウ

アプリケーションのメイン画面
"""
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QGroupBox, QMessageBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

from infrastructure.database_manager import DatabaseManager
from infrastructure.file_manager import FileManager
from infrastructure.config_manager import ConfigManager
from infrastructure.logger import Logger
from services.data_import_service import DataImportService
from services.data_export_service import DataExportService


class MainWindow(QMainWindow):
    """メインウィンドウクラス"""
    
    def __init__(self,
                 db_manager: DatabaseManager,
                 file_manager: FileManager,
                 config_manager: ConfigManager,
                 logger: Logger,
                 import_service: DataImportService,
                 export_service: DataExportService):
        """
        初期化
        
        Args:
            db_manager: データベースマネージャー
            file_manager: ファイルマネージャー
            config_manager: 設定マネージャー
            logger: ログマネージャー
            import_service: インポートサービス
            export_service: エクスポートサービス
        """
        super().__init__()
        
        self.db_manager = db_manager
        self.file_manager = file_manager
        self.config_manager = config_manager
        self.logger = logger
        self.import_service = import_service
        self.export_service = export_service
        
        self._init_ui()
        self._load_window_geometry()
        
        # アプリ起動ログ
        self.logger.log_action(Logger.ACTION_APP_START, "アプリケーション起動")
    
    def _init_ui(self):
        """UI初期化"""
        self.setWindowTitle("成績管理システム Phase2")
        
        # 中央ウィジェット
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # メインレイアウト
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # タイトル
        title_label = QLabel("成績管理システム")
        title_font = QFont()
        title_font.setPointSize(24)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # サブタイトル
        subtitle_label = QLabel("Phase 2 - データ管理・分析システム")
        subtitle_font = QFont()
        subtitle_font.setPointSize(12)
        subtitle_label.setFont(subtitle_font)
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setStyleSheet("color: #666;")
        main_layout.addWidget(subtitle_label)
        
        main_layout.addSpacing(20)
        
        # データ取り込みセクション
        import_group = self._create_import_section()
        main_layout.addWidget(import_group)
        
        # データ閲覧セクション
        view_group = self._create_view_section()
        main_layout.addWidget(view_group)
        
        # システムセクション
        system_group = self._create_system_section()
        main_layout.addWidget(system_group)
        
        main_layout.addStretch()
        
        # ステータスバー
        self.statusBar().showMessage("準備完了")
    
    def _create_import_section(self) -> QGroupBox:
        """データ取り込みセクション作成"""
        group = QGroupBox("📁 データ取り込み")
        group.setStyleSheet("""
            QGroupBox {
                font-size: 14px;
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 5px;
                margin-top: 10px;
                padding: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)
        
        layout = QHBoxLayout()
        layout.setSpacing(10)
        
        # 評定取り込みボタン
        btn_import_grades = QPushButton("📊 評定データ取り込み")
        btn_import_grades.setMinimumHeight(60)
        btn_import_grades.setStyleSheet(self._get_button_style("#3498db"))
        btn_import_grades.clicked.connect(lambda: self.show_import_dialog('評定'))
        layout.addWidget(btn_import_grades)
        
        # 観点取り込みボタン
        btn_import_viewpoints = QPushButton("📝 観点別評価取り込み")
        btn_import_viewpoints.setMinimumHeight(60)
        btn_import_viewpoints.setStyleSheet(self._get_button_style("#9b59b6"))
        btn_import_viewpoints.clicked.connect(lambda: self.show_import_dialog('観点'))
        layout.addWidget(btn_import_viewpoints)
        
        # 欠課情報取り込みボタン
        btn_import_absences = QPushButton("⏰ 欠課情報取り込み")
        btn_import_absences.setMinimumHeight(60)
        btn_import_absences.setStyleSheet(self._get_button_style("#e74c3c"))
        btn_import_absences.clicked.connect(lambda: self.show_import_dialog('欠課情報'))
        layout.addWidget(btn_import_absences)
        
        group.setLayout(layout)
        return group
    
    def _create_view_section(self) -> QGroupBox:
        """データ閲覧セクション作成"""
        group = QGroupBox("👥 データ閲覧・管理")
        group.setStyleSheet("""
            QGroupBox {
                font-size: 14px;
                font-weight: bold;
                border: 2px solid #27ae60;
                border-radius: 5px;
                margin-top: 10px;
                padding: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)
        
        layout = QHBoxLayout()
        layout.setSpacing(10)
        
        # 生徒一覧ボタン
        btn_students = QPushButton("👤 生徒一覧")
        btn_students.setMinimumHeight(60)
        btn_students.setStyleSheet(self._get_button_style("#27ae60"))
        btn_students.clicked.connect(self.show_student_list)
        layout.addWidget(btn_students)
        
        # 科目一覧ボタン
        btn_courses = QPushButton("📚 科目一覧")
        btn_courses.setMinimumHeight(60)
        btn_courses.setStyleSheet(self._get_button_style("#16a085"))
        btn_courses.clicked.connect(self.show_course_list)
        layout.addWidget(btn_courses)
        
        # データ管理ボタン
        btn_data_mgmt = QPushButton("⚙️ データ管理")
        btn_data_mgmt.setMinimumHeight(60)
        btn_data_mgmt.setStyleSheet(self._get_button_style("#2c3e50"))
        btn_data_mgmt.clicked.connect(self.show_data_management)
        layout.addWidget(btn_data_mgmt)
        
        group.setLayout(layout)
        return group
    
    def _create_system_section(self) -> QGroupBox:
        """システムセクション作成"""
        group = QGroupBox("🔧 システム")
        group.setStyleSheet("""
            QGroupBox {
                font-size: 14px;
                font-weight: bold;
                border: 2px solid #95a5a6;
                border-radius: 5px;
                margin-top: 10px;
                padding: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)
        
        layout = QHBoxLayout()
        layout.setSpacing(10)
        
        # ログビューアボタン
        btn_logs = QPushButton("📋 ログビューア")
        btn_logs.setMinimumHeight(60)
        btn_logs.setStyleSheet(self._get_button_style("#7f8c8d"))
        btn_logs.clicked.connect(self.show_log_viewer)
        layout.addWidget(btn_logs)
        
        # 設定ボタン
        btn_settings = QPushButton("⚙️ 設定")
        btn_settings.setMinimumHeight(60)
        btn_settings.setStyleSheet(self._get_button_style("#34495e"))
        btn_settings.clicked.connect(self.show_settings)
        layout.addWidget(btn_settings)
        
        # ヘルプボタン
        btn_help = QPushButton("❓ ヘルプ")
        btn_help.setMinimumHeight(60)
        btn_help.setStyleSheet(self._get_button_style("#95a5a6"))
        btn_help.clicked.connect(self.show_help)
        layout.addWidget(btn_help)
        
        group.setLayout(layout)
        return group
    
    def _get_button_style(self, color: str) -> str:
        """ボタンスタイル取得"""
        return f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border: none;
                border-radius: 5px;
                font-size: 14px;
                font-weight: bold;
                padding: 10px;
            }}
            QPushButton:hover {{
                background-color: {color}dd;
            }}
            QPushButton:pressed {{
                background-color: {color}aa;
            }}
        """
    
    def show_import_dialog(self, data_type: str):
        """データ取り込みダイアログ表示"""
        try:
            from ui.dialogs.period_import_dialog import PeriodImportDialog
            
            dialog = PeriodImportDialog(
                data_type=data_type,
                db_manager=self.db_manager,
                config_manager=self.config_manager,
                file_manager=self.file_manager,
                import_service=self.import_service,
                parent=self
            )
            dialog.exec()
            
        except ImportError as e:
            QMessageBox.warning(
                self, "エラー",
                f"ダイアログの読み込みに失敗しました\n{str(e)}"
            )
    
    def show_student_list(self):
        """生徒一覧表示"""
        try:
            from ui.dialogs.student_list_dialog import StudentListDialog
            
            dialog = StudentListDialog(
                db_manager=self.db_manager,
                parent=self
            )
            dialog.exec()
            
        except ImportError as e:
            QMessageBox.information(
                self, "開発中",
                "この機能は現在開発中です"
            )
    
    def show_course_list(self):
        """科目一覧表示"""
        try:
            from ui.dialogs.course_list_dialog import CourseListDialog
            
            dialog = CourseListDialog(
                db_manager=self.db_manager,
                parent=self
            )
            dialog.exec()
            
        except ImportError:
            QMessageBox.information(
                self, "開発中",
                "この機能は現在開発中です"
            )
    
    def show_data_management(self):
        """データ管理画面表示"""
        try:
            from ui.dialogs.data_management_dialog import DataManagementDialog
            
            dialog = DataManagementDialog(
                db_manager=self.db_manager,
                logger=self.logger,
                parent=self
            )
            dialog.exec()
            
        except ImportError:
            QMessageBox.information(
                self, "開発中",
                "この機能は現在開発中です"
            )
    
    def show_log_viewer(self):
        """ログビューア表示"""
        try:
            from ui.dialogs.log_viewer_dialog import LogViewerDialog
            
            dialog = LogViewerDialog(
                logger=self.logger,
                parent=self
            )
            dialog.exec()
            
        except ImportError:
            QMessageBox.information(
                self, "開発中",
                "この機能は現在開発中です"
            )
    
    def show_settings(self):
        """設定画面表示"""
        QMessageBox.information(
            self, "設定",
            "設定画面は今後実装予定です"
        )
    
    def show_help(self):
        """ヘルプ表示"""
        help_text = """
        <h2>成績管理システム Phase2</h2>
        <h3>使い方</h3>
        <ul>
            <li><b>データ取り込み:</b> Excelファイルから成績データを取り込みます</li>
            <li><b>生徒一覧:</b> 登録されている生徒の一覧を表示します</li>
            <li><b>科目一覧:</b> 登録されている科目の一覧を表示します</li>
            <li><b>データ管理:</b> データの削除や出力を行います</li>
            <li><b>ログビューア:</b> システムの操作履歴を確認します</li>
        </ul>
        <p>詳細はREADME.mdを参照してください</p>
        """
        
        QMessageBox.information(
            self, "ヘルプ",
            help_text
        )
    
    def _load_window_geometry(self):
        """ウィンドウ位置・サイズを読み込み"""
        width, height, x, y = self.config_manager.get_window_geometry()
        self.resize(width, height)
        
        if x is not None and y is not None:
            self.move(x, y)
    
    def _save_window_geometry(self):
        """ウィンドウ位置・サイズを保存"""
        geometry = self.geometry()
        self.config_manager.save_window_geometry(
            geometry.width(),
            geometry.height(),
            geometry.x(),
            geometry.y()
        )
    
    def closeEvent(self, event):
        """ウィンドウクローズイベント"""
        # ウィンドウ位置・サイズ保存
        self._save_window_geometry()
        
        # アプリ終了ログ
        self.logger.log_action(Logger.ACTION_APP_EXIT, "アプリケーション終了")
        
        event.accept()
