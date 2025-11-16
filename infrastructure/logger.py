"""
ログ記録クラス

操作ログの記録・閲覧を管理
"""
from datetime import datetime
from typing import List, Tuple, Optional
from infrastructure.database_manager import DatabaseManager


class Logger:
    """ログ記録クラス"""
    
    def __init__(self, db_manager: DatabaseManager):
        """
        初期化
        
        Args:
            db_manager: データベースマネージャー
        """
        self.db = db_manager
    
    def log_action(self, action_type: str, details: str, user_name: str = None):
        """
        アクション記録
        
        Args:
            action_type: 操作種別
            details: 詳細情報
            user_name: ユーザー名（オプション）
        """
        query = """
            INSERT INTO activity_logs 
            (action_type, details, user_name, timestamp)
            VALUES (?, ?, ?, ?)
        """
        
        try:
            self.db.execute_query(
                query, 
                (action_type, details, user_name, datetime.now())
            )
            print(f"ログ記録: {action_type} - {details}")
        except Exception as e:
            print(f"ログ記録エラー: {e}")
    
    def get_recent_logs(self, limit: int = 100) -> List[Tuple]:
        """
        最近のログ取得
        
        Args:
            limit: 取得件数
            
        Returns:
            List[Tuple]: ログデータのリスト
        """
        query = """
            SELECT id, timestamp, action_type, details, user_name
            FROM activity_logs
            ORDER BY timestamp DESC
            LIMIT ?
        """
        
        try:
            return self.db.fetch_all(query, (limit,))
        except Exception as e:
            print(f"ログ取得エラー: {e}")
            return []
    
    def search_logs(self, 
                   keyword: str = None, 
                   start_date: datetime = None, 
                   end_date: datetime = None,
                   action_type: str = None,
                   limit: int = 1000) -> List[Tuple]:
        """
        ログ検索
        
        Args:
            keyword: キーワード
            start_date: 開始日時
            end_date: 終了日時
            action_type: 操作種別
            limit: 取得件数上限
            
        Returns:
            List[Tuple]: ログデータのリスト
        """
        query = """
            SELECT id, timestamp, action_type, details, user_name
            FROM activity_logs
            WHERE 1=1
        """
        params = []
        
        # キーワード検索
        if keyword:
            query += " AND (action_type LIKE ? OR details LIKE ?)"
            params.extend([f"%{keyword}%", f"%{keyword}%"])
        
        # 日時範囲
        if start_date:
            query += " AND timestamp >= ?"
            params.append(start_date)
        
        if end_date:
            query += " AND timestamp <= ?"
            params.append(end_date)
        
        # 操作種別
        if action_type:
            query += " AND action_type = ?"
            params.append(action_type)
        
        query += " ORDER BY timestamp DESC LIMIT ?"
        params.append(limit)
        
        try:
            return self.db.fetch_all(query, tuple(params))
        except Exception as e:
            print(f"ログ検索エラー: {e}")
            return []
    
    def get_action_types(self) -> List[str]:
        """
        操作種別一覧取得
        
        Returns:
            List[str]: 操作種別のリスト
        """
        query = """
            SELECT DISTINCT action_type
            FROM activity_logs
            ORDER BY action_type
        """
        
        try:
            results = self.db.fetch_all(query)
            return [row[0] for row in results]
        except Exception as e:
            print(f"操作種別取得エラー: {e}")
            return []
    
    def get_log_statistics(self, days: int = 30) -> dict:
        """
        ログ統計取得
        
        Args:
            days: 対象日数
            
        Returns:
            dict: 統計情報
        """
        cutoff_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        cutoff_date = cutoff_date.replace(day=cutoff_date.day - days)
        
        # 総件数
        total_query = """
            SELECT COUNT(*) FROM activity_logs
            WHERE timestamp >= ?
        """
        total_result = self.db.fetch_one(total_query, (cutoff_date,))
        total_count = total_result[0] if total_result else 0
        
        # 操作種別ごとの件数
        type_query = """
            SELECT action_type, COUNT(*) as count
            FROM activity_logs
            WHERE timestamp >= ?
            GROUP BY action_type
            ORDER BY count DESC
        """
        type_results = self.db.fetch_all(type_query, (cutoff_date,))
        
        # 日別件数
        daily_query = """
            SELECT DATE(timestamp) as date, COUNT(*) as count
            FROM activity_logs
            WHERE timestamp >= ?
            GROUP BY DATE(timestamp)
            ORDER BY date DESC
        """
        daily_results = self.db.fetch_all(daily_query, (cutoff_date,))
        
        return {
            'total_count': total_count,
            'by_type': {row[0]: row[1] for row in type_results},
            'by_date': {row[0]: row[1] for row in daily_results}
        }
    
    def delete_old_logs(self, days: int = 90) -> int:
        """
        古いログを削除
        
        Args:
            days: 削除対象の日数（これより古いログを削除）
            
        Returns:
            int: 削除したログ数
        """
        cutoff_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        cutoff_date = cutoff_date.replace(day=cutoff_date.day - days)
        
        query = "DELETE FROM activity_logs WHERE timestamp < ?"
        
        try:
            count = self.db.execute_query(query, (cutoff_date,))
            print(f"{count}件の古いログを削除しました")
            return count
        except Exception as e:
            print(f"ログ削除エラー: {e}")
            return 0
    
    # ログタイプ定数
    ACTION_IMPORT = 'data_import'
    ACTION_EXPORT = 'data_export'
    ACTION_DELETE = 'data_delete'
    ACTION_SEARCH = 'data_search'
    ACTION_VIEW = 'data_view'
    ACTION_EDIT = 'data_edit'
    ACTION_APP_START = 'app_start'
    ACTION_APP_EXIT = 'app_exit'
    ACTION_ERROR = 'error'
