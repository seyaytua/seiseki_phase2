"""
成績管理システム Phase2

メインエントリーポイント
"""
import sys
from PySide6.QtWidgets import QApplication

from infrastructure.database_manager import DatabaseManager
from infrastructure.file_manager import FileManager
from infrastructure.config_manager import ConfigManager
from infrastructure.logger import Logger
from services.data_import_service import DataImportService
from services.data_export_service import DataExportService
from ui.main_window import MainWindow


def main():
    """メイン関数"""
    # アプリケーション作成
    app = QApplication(sys.argv)
    app.setApplicationName("成績管理システム Phase2")
    app.setOrganizationName("School")
    
    try:
        # インフラ層の初期化
        db_manager = DatabaseManager("data/database.db")
        db_manager.connect()
        
        file_manager = FileManager("data")
        config_manager = ConfigManager("config.json")
        logger = Logger(db_manager)
        
        # サービス層の初期化
        import_service = DataImportService(db_manager, file_manager, logger)
        export_service = DataExportService(db_manager, file_manager, logger)
        
        # メインウィンドウ作成
        window = MainWindow(
            db_manager=db_manager,
            file_manager=file_manager,
            config_manager=config_manager,
            logger=logger,
            import_service=import_service,
            export_service=export_service
        )
        
        window.show()
        
        # アプリケーション実行
        sys.exit(app.exec())
        
    except Exception as e:
        print(f"起動エラー: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        # クリーンアップ
        if 'db_manager' in locals():
            db_manager.close()


if __name__ == "__main__":
    main()
