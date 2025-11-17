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
from utils.logger import Logger
from utils.data_importer import DataImporter
from ui.main_window import MainWindow


def get_database_path(app, config_manager):
    """データベースパスを取得（初回のみダイアログ表示）"""
    settings = config_manager.get_settings()
    
    # 設定から前回のDBパスを取得
    last_db_path = settings.get('database', {}).get('path')
    
    # 前回のDBパスが存在し、ファイルも存在する場合はそれを使用
    if last_db_path and Path(last_db_path).exists():
        print(f"前回使用したデータベースを使用: {last_db_path}")
        return last_db_path
    
    # 存在しない場合はダイアログを表示
    print("データベースを選択してください...")
    from ui.database_selector_dialog import DatabaseSelectorDialog
    
    # デフォルトパス
    default_path = Path(__file__).parent / 'data' / 'database.db'
    
    dialog = DatabaseSelectorDialog(str(default_path))
    if dialog.exec():
        selected_path = dialog.get_selected_path()
        
        # 選択したパスを設定に保存
        if 'database' not in settings:
            settings['database'] = {}
        settings['database']['path'] = selected_path
        settings['database']['description'] = "選択されたデータベース"
        config_manager.save_settings(settings)
        
        return selected_path
    else:
        # キャンセルされた場合はデフォルトパスを使用
        print("デフォルトデータベースを使用します")
        return str(default_path)


def main():
    """メインエントリーポイント"""
    db_manager = None
    try:
        app = QApplication(sys.argv)
        
        print("アプリケーション起動中...")
        
        # ConfigManager初期化
        config_manager = ConfigManager()
        
        # データベースパス取得（初回のみダイアログ表示）
        db_path = get_database_path(app, config_manager)
        
        # データディレクトリ作成
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        
        # データベース接続
        db_manager = DatabaseManager(db_path=str(db_path))
        db_manager.connect()
        
        print(f"使用中のデータベース: {db_manager.db_path}")
        
        # 各種マネージャー初期化
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
