"""
データベース管理クラス

SQLiteデータベースの接続・操作を管理
"""
import sqlite3
from pathlib import Path
from typing import Optional, List, Tuple, Any
from datetime import datetime


class DatabaseManager:
    """データベース管理クラス"""
    
    def __init__(self, db_path: str = "data/database.db"):
        """
        初期化
        
        Args:
            db_path: データベースファイルのパス
        """
        self.db_path = Path(db_path)
        self.connection: Optional[sqlite3.Connection] = None
        self.cursor: Optional[sqlite3.Cursor] = None
        
        # データベースディレクトリの作成
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
    
    def connect(self):
        """データベースに接続"""
        try:
            self.connection = sqlite3.connect(str(self.db_path))
            self.cursor = self.connection.cursor()
            
            # 外部キー制約を有効化
            self.cursor.execute("PRAGMA foreign_keys = ON")
            
            # テーブル作成
            self._create_tables()
            
            print(f"データベース接続成功: {self.db_path}")
        except sqlite3.Error as e:
            print(f"データベース接続エラー: {e}")
            raise
    
    def _create_tables(self):
        """テーブル作成"""
        
        # 評定テーブル
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS grades (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_number TEXT NOT NULL,
                student_name TEXT,
                course_number TEXT NOT NULL,
                course_name TEXT,
                school_subject_name TEXT,
                period TEXT NOT NULL,
                year INTEGER NOT NULL,
                grade_value TEXT,
                credits REAL,
                acquisition_credits REAL,
                remarks TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(student_number, course_number, period, year)
            )
        """)
        
        # 観点別評価テーブル
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS viewpoint_evaluations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_number TEXT NOT NULL,
                student_name TEXT,
                course_number TEXT NOT NULL,
                course_name TEXT,
                school_subject_name TEXT,
                period TEXT NOT NULL,
                year INTEGER NOT NULL,
                viewpoint_1 TEXT,
                viewpoint_2 TEXT,
                viewpoint_3 TEXT,
                viewpoint_4 TEXT,
                viewpoint_5 TEXT,
                remarks TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(student_number, course_number, period, year)
            )
        """)
        
        # 欠課情報テーブル
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS absences (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_number TEXT NOT NULL,
                student_name TEXT,
                course_number TEXT NOT NULL,
                course_name TEXT,
                school_subject_name TEXT,
                period TEXT NOT NULL,
                year INTEGER NOT NULL,
                absent_count INTEGER DEFAULT 0,
                late_count INTEGER DEFAULT 0,
                total_hours INTEGER DEFAULT 0,
                absence_rate REAL DEFAULT 0.0,
                remarks TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(student_number, course_number, period, year)
            )
        """)
        
        # 操作ログテーブル
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS activity_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                action_type TEXT NOT NULL,
                details TEXT,
                user_name TEXT
            )
        """)
        
        self.connection.commit()
    
    def execute_query(self, query: str, params: Tuple = None) -> int:
        """
        クエリ実行（INSERT, UPDATE, DELETE）
        
        Args:
            query: SQLクエリ
            params: パラメータ
            
        Returns:
            int: 影響を受けた行数
        """
        try:
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            
            self.connection.commit()
            return self.cursor.rowcount
        except sqlite3.Error as e:
            print(f"クエリ実行エラー: {e}")
            self.connection.rollback()
            raise
    
    def fetch_all(self, query: str, params: Tuple = None) -> List[Tuple]:
        """
        全データ取得（SELECT）
        
        Args:
            query: SQLクエリ
            params: パラメータ
            
        Returns:
            List[Tuple]: 取得結果
        """
        try:
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            
            return self.cursor.fetchall()
        except sqlite3.Error as e:
            print(f"データ取得エラー: {e}")
            raise
    
    def fetch_one(self, query: str, params: Tuple = None) -> Optional[Tuple]:
        """
        1件取得（SELECT）
        
        Args:
            query: SQLクエリ
            params: パラメータ
            
        Returns:
            Optional[Tuple]: 取得結果
        """
        try:
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            
            return self.cursor.fetchone()
        except sqlite3.Error as e:
            print(f"データ取得エラー: {e}")
            raise
    
    def insert_grade(self, grade_data: dict) -> int:
        """
        評定データ挿入
        
        Args:
            grade_data: 評定データ辞書
            
        Returns:
            int: 挿入された行のID
        """
        query = """
            INSERT OR REPLACE INTO grades 
            (student_number, student_name, course_number, course_name, 
             school_subject_name, period, year, grade_value, credits, 
             acquisition_credits, remarks, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        params = (
            grade_data.get('student_number'),
            grade_data.get('student_name'),
            grade_data.get('course_number'),
            grade_data.get('course_name'),
            grade_data.get('school_subject_name'),
            grade_data.get('period'),
            grade_data.get('year'),
            grade_data.get('grade_value'),
            grade_data.get('credits'),
            grade_data.get('acquisition_credits'),
            grade_data.get('remarks'),
            datetime.now()
        )
        
        self.execute_query(query, params)
        return self.cursor.lastrowid
    
    def insert_viewpoint(self, viewpoint_data: dict) -> int:
        """
        観点別評価データ挿入
        
        Args:
            viewpoint_data: 観点別評価データ辞書
            
        Returns:
            int: 挿入された行のID
        """
        query = """
            INSERT OR REPLACE INTO viewpoint_evaluations 
            (student_number, student_name, course_number, course_name, 
             school_subject_name, period, year, viewpoint_1, viewpoint_2, 
             viewpoint_3, viewpoint_4, viewpoint_5, remarks, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        params = (
            viewpoint_data.get('student_number'),
            viewpoint_data.get('student_name'),
            viewpoint_data.get('course_number'),
            viewpoint_data.get('course_name'),
            viewpoint_data.get('school_subject_name'),
            viewpoint_data.get('period'),
            viewpoint_data.get('year'),
            viewpoint_data.get('viewpoint_1'),
            viewpoint_data.get('viewpoint_2'),
            viewpoint_data.get('viewpoint_3'),
            viewpoint_data.get('viewpoint_4'),
            viewpoint_data.get('viewpoint_5'),
            viewpoint_data.get('remarks'),
            datetime.now()
        )
        
        self.execute_query(query, params)
        return self.cursor.lastrowid
    
    def insert_absence(self, absence_data: dict) -> int:
        """
        欠課情報データ挿入
        
        Args:
            absence_data: 欠課情報データ辞書
            
        Returns:
            int: 挿入された行のID
        """
        query = """
            INSERT OR REPLACE INTO absences 
            (student_number, student_name, course_number, course_name, 
             school_subject_name, period, year, absent_count, late_count, 
             total_hours, absence_rate, remarks, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        params = (
            absence_data.get('student_number'),
            absence_data.get('student_name'),
            absence_data.get('course_number'),
            absence_data.get('course_name'),
            absence_data.get('school_subject_name'),
            absence_data.get('period'),
            absence_data.get('year'),
            absence_data.get('absent_count', 0),
            absence_data.get('late_count', 0),
            absence_data.get('total_hours', 0),
            absence_data.get('absence_rate', 0.0),
            absence_data.get('remarks'),
            datetime.now()
        )
        
        self.execute_query(query, params)
        return self.cursor.lastrowid
    
    def get_students(self, search_term: str = None) -> List[Tuple]:
        """
        生徒一覧取得
        
        Args:
            search_term: 検索キーワード
            
        Returns:
            List[Tuple]: 生徒一覧
        """
        query = """
            SELECT DISTINCT student_number, student_name 
            FROM grades 
            WHERE 1=1
        """
        params = []
        
        if search_term:
            query += " AND (student_number LIKE ? OR student_name LIKE ?)"
            params = [f"%{search_term}%", f"%{search_term}%"]
        
        query += " ORDER BY student_number"
        
        return self.fetch_all(query, tuple(params) if params else None)
    
    def get_courses(self, search_term: str = None) -> List[Tuple]:
        """
        科目一覧取得
        
        Args:
            search_term: 検索キーワード
            
        Returns:
            List[Tuple]: 科目一覧
        """
        query = """
            SELECT DISTINCT course_number, course_name, school_subject_name 
            FROM grades 
            WHERE 1=1
        """
        params = []
        
        if search_term:
            query += " AND (course_number LIKE ? OR course_name LIKE ? OR school_subject_name LIKE ?)"
            params = [f"%{search_term}%", f"%{search_term}%", f"%{search_term}%"]
        
        query += " ORDER BY course_number"
        
        return self.fetch_all(query, tuple(params) if params else None)
    
    def delete_data_by_period(self, data_type: str, period: str, year: int) -> int:
        """
        期間指定によるデータ削除
        
        Args:
            data_type: データタイプ（'評定', '観点', '欠課情報'）
            period: 期間
            year: 年度
            
        Returns:
            int: 削除された行数
        """
        table_map = {
            '評定': 'grades',
            '観点': 'viewpoint_evaluations',
            '欠課情報': 'absences'
        }
        
        table = table_map.get(data_type)
        if not table:
            raise ValueError(f"無効なデータタイプ: {data_type}")
        
        query = f"DELETE FROM {table} WHERE period = ? AND year = ?"
        return self.execute_query(query, (period, year))
    
    def close(self):
        """データベース接続を閉じる"""
        if self.connection:
            self.connection.close()
            print("データベース接続を閉じました")
