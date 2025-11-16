from datetime import datetime


class Logger:
    """ログ記録クラス"""
    
    def __init__(self, db_manager):
        """初期化"""
        self.db = db_manager
        self.create_log_table()
    
    def create_log_table(self):
        """ログテーブル作成"""
        try:
            query = """ CREATE TABLE IF NOT EXISTS action_logs ( id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT NOT NULL, action_type TEXT NOT NULL, details TEXT, user TEXT DEFAULT 'system', created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ) """
            self.db.execute_query(query)
        except Exception as e:
            print(f"ログテーブル作成エラー: {e}")
    
    def log_action(self, action_type, details="", user="system"):
        """アクション記録"""
        try:
            # データベース接続確認
            if not self.db.connection:
                print(f"警告: データベース接続が閉じられています。ログ記録をスキップします: {action_type}")
                return
            
            query = """ INSERT INTO action_logs (timestamp, action_type, details, user) VALUES (?, ?, ?, ?) """
            params = (datetime.now().isoformat(), action_type, details, user)
            self.db.execute_query(query, params)
            
        except Exception as e:
            print(f"ログ記録エラー: {e}")
    
    def get_logs(self, limit=100, action_type=None, start_date=None, end_date=None):
        """ログ取得"""
        try:
            query = "SELECT * FROM action_logs WHERE 1=1"
            params = []
            
            if action_type:
                query += " AND action_type = ?"
                params.append(action_type)
            
            if start_date:
                query += " AND timestamp >= ?"
                params.append(start_date)
            
            if end_date:
                query += " AND timestamp <= ?"
                params.append(end_date)
            
            query += " ORDER BY timestamp DESC LIMIT ?"
            params.append(limit)
            
            return self.db.fetch_all(query, tuple(params))
            
        except Exception as e:
            print(f"ログ取得エラー: {e}")
            return []
    
    def get_recent_logs(self, limit=50):
        """最近のログ取得"""
        return self.get_logs(limit=limit)
    
    def get_logs_by_type(self, action_type, limit=100):
        """タイプ別ログ取得"""
        return self.get_logs(limit=limit, action_type=action_type)
    
    def get_logs_by_date_range(self, start_date, end_date, limit=100):
        """期間別ログ取得"""
        return self.get_logs(limit=limit, start_date=start_date, end_date=end_date)
    
    def clear_old_logs(self, days=90):
        """古いログ削除"""
        try:
            cutoff_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            cutoff_date = cutoff_date.replace(day=cutoff_date.day - days)
            
            query = "DELETE FROM action_logs WHERE timestamp < ?"
            params = (cutoff_date.isoformat(),)
            
            self.db.execute_query(query, params)
            
            return True
            
        except Exception as e:
            print(f"ログ削除エラー: {e}")
            return False
    
    def get_log_count(self, action_type=None):
        """ログ件数取得"""
        try:
            if action_type:
                query = "SELECT COUNT(*) FROM action_logs WHERE action_type = ?"
                params = (action_type,)
            else:
                query = "SELECT COUNT(*) FROM action_logs"
                params = ()
            
            result = self.db.fetch_one(query, params)
            return result[0] if result else 0
            
        except Exception as e:
            print(f"ログ件数取得エラー: {e}")
            return 0
    
    def get_log_statistics(self):
        """ログ統計取得"""
        try:
            query = """ SELECT action_type, COUNT(*) as count, MAX(timestamp) as last_occurrence FROM action_logs GROUP BY action_type ORDER BY count DESC """
            
            return self.db.fetch_all(query)
            
        except Exception as e:
            print(f"ログ統計取得エラー: {e}")
            return []
    
    def export_logs(self, file_path, start_date=None, end_date=None):
        """ログをエクスポート"""
        try:
            import csv
            
            logs = self.get_logs(
                limit=10000,
                start_date=start_date,
                end_date=end_date
            )
            
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                if logs:
                    writer = csv.writer(f)
                    
                    # ヘッダー
                    headers = ['ID', 'タイムスタンプ', 'アクション', '詳細', 'ユーザー', '作成日時']
                    writer.writerow(headers)
                    
                    # データ
                    for log in logs:
                        writer.writerow(log)
            
            return True
            
        except Exception as e:
            print(f"ログエクスポートエラー: {e}")
            return False