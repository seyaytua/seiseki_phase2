import sys
import traceback
from pathlib import Path
from PySide6.QtWidgets import QApplication

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from database.db_manager import DatabaseManager
from utils.config_manager import ConfigManager
from utils.file_manager import FileManager
from utils.logger import Logger  # ← ここを修正
from utils.data_importer import DataImporter
from ui.main_window import MainWindow


def main():
    """メインエントリーポイント"""
    db_manager = None
    try:
        app = QApplication(sys.argv)
        
        print("アプリケーション起動中...")
        
        # データベースパス設定
        data_dir = Path(__file__).parent / 'data'
        data_dir.mkdir(parents=True, exist_ok=True)
        db_path = data_dir / 'database.db'
        
        # データベース接続
        db_manager = DatabaseManager(db_path=str(db_path))
        db_manager.connect()
        
        print(f"使用中のデータベース: {db_manager.db_path}")
        
        # 各種マネージャー初期化
        config_manager = ConfigManager()
        file_manager = FileManager(config_manager)
        logger = Logger(db_manager)
        data_importer = DataImporter(db_manager, file_manager, logger)
        
        print("アプリケーション起動完了！")
        
        # メインウィンドウ作成
        main_window = MainWindow(
            db_manager=db_manager,
            config_manager=config_manager,
            file_manager=file_manager,
            data_importer=data_importer,
            logger=logger
        )
        
        main_window.show()
        
        # ログ記録 - ウィンドウ表示後
        logger.log_action("application_start", "アプリケーション起動")
        
        # アプリケーション実行
        exit_code = app.exec()
        
        # クリーンアップ - ログ記録してから接続を閉じる
        try:
            logger.log_action("application_exit", "アプリケーション終了")
        except Exception as log_error:
            print(f"終了ログ記録エラー: {log_error}")
        
        if db_manager:
            db_manager.close()
        
        return exit_code
        
    except Exception as e:
        print(f"\nエラー発生: {type(e).__name__}")
        print(f"エラー内容: {str(e)}")
        traceback.print_exc()
        
        if db_manager:
            try:
                db_manager.close()
            except:
                pass
        
        input("\nEnterキーを押して終了...")
        return 1


if __name__ == "__main__":
    sys.exit(main())