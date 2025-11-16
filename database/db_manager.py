import sqlite3
from pathlib import Path


class DatabaseManager:
    """データベース管理クラス"""
    
    def __init__(self, db_path=None):
        """初期化"""
        if db_path is None:
            # デフォルトパス
            base_dir = Path(__file__).parent.parent
            data_dir = base_dir / 'data'
            data_dir.mkdir(parents=True, exist_ok=True)
            self.db_path = str(data_dir / 'database.db')
        else:
            self.db_path = db_path
        
        self.connection = None
    
    def connect(self):
        """データベース接続"""
        try:
            self.connection = sqlite3.connect(self.db_path)
            self.connection.row_factory = sqlite3.Row
            self.create_tables()
            return True
        except Exception as e:
            print(f"データベース接続エラー: {e}")
            raise
    
    def get_connection(self):
        """データベース接続を取得"""
        if self.connection is None:
            self.connect()
        return self.connection
    
    def close(self):
        """データベース切断"""
        if self.connection:
            try:
                self.connection.close()
                self.connection = None
            except Exception as e:
                print(f"データベース切断エラー: {e}")
    
    def create_tables(self):
        """テーブル作成"""
        try:
            # 評定テーブル
            self.execute_query(""" CREATE TABLE IF NOT EXISTS grades ( id INTEGER PRIMARY KEY AUTOINCREMENT, student_id TEXT NOT NULL, subject_code TEXT NOT NULL, period TEXT NOT NULL, year INTEGER NOT NULL, grade INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, UNIQUE(student_id, subject_code, period, year) ) """)
            
            # 観点テーブル
            self.execute_query(""" CREATE TABLE IF NOT EXISTS viewpoints ( id INTEGER PRIMARY KEY AUTOINCREMENT, student_id TEXT NOT NULL, subject_code TEXT NOT NULL, period TEXT NOT NULL, year INTEGER NOT NULL, viewpoint1 TEXT, viewpoint2 TEXT, viewpoint3 TEXT, viewpoint4 TEXT, viewpoint5 TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, UNIQUE(student_id, subject_code, period, year) ) """)
            
            # 欠課情報テーブル
            self.execute_query(""" CREATE TABLE IF NOT EXISTS absences ( id INTEGER PRIMARY KEY AUTOINCREMENT, student_id TEXT NOT NULL, subject_code TEXT NOT NULL, period TEXT NOT NULL, year INTEGER NOT NULL, absence_count INTEGER DEFAULT 0, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, UNIQUE(student_id, subject_code, period, year) ) """)
            
            # 生徒マスタ
            self.execute_query(""" CREATE TABLE IF NOT EXISTS students ( id INTEGER PRIMARY KEY AUTOINCREMENT, student_id TEXT UNIQUE NOT NULL, name TEXT, grade INTEGER, class TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ) """)
            
            # 科目マスタ
            self.execute_query(""" CREATE TABLE IF NOT EXISTS subjects ( id INTEGER PRIMARY KEY AUTOINCREMENT, subject_code TEXT UNIQUE NOT NULL, subject_name TEXT NOT NULL, category TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ) """)
            
        except Exception as e:
            print(f"テーブル作成エラー: {e}")
            raise
    
    def execute_query(self, query, params=None):
        """クエリ実行"""
        try:
            if not self.connection:
                raise Exception("データベースが接続されていません")
            
            cursor = self.connection.cursor()
            
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            
            self.connection.commit()
            return cursor
            
        except Exception as e:
            print(f"クエリ実行エラー: {e}")
            if self.connection:
                self.connection.rollback()
            raise
    
    def fetch_all(self, query, params=None):
        """全行取得"""
        try:
            cursor = self.execute_query(query, params)
            return cursor.fetchall()
        except Exception as e:
            print(f"データ取得エラー: {e}")
            return []
    
    def fetch_one(self, query, params=None):
        """1行取得"""
        try:
            cursor = self.execute_query(query, params)
            return cursor.fetchone()
        except Exception as e:
            print(f"データ取得エラー: {e}")
            return None
    
    def get_table_info(self, table_name):
        """テーブル情報取得"""
        try:
            query = f"PRAGMA table_info({table_name})"
            return self.fetch_all(query)
        except Exception as e:
            print(f"テーブル情報取得エラー: {e}")
            return []
    
    def table_exists(self, table_name):
        """テーブル存在確認"""
        try:
            query = """ SELECT name FROM sqlite_master WHERE type='table' AND name=? """
            result = self.fetch_one(query, (table_name,))
            return result is not None
        except Exception as e:
            print(f"テーブル存在確認エラー: {e}")
            return False