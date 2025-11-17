import pandas as pd
from pathlib import Path
from datetime import datetime


class DataImporter:
    """データ取り込みクラス"""
    
    def __init__(self, db_manager, file_manager, logger):
        """初期化"""
        self.db = db_manager
        self.file_manager = file_manager
        self.logger = logger
    
    def import_data(self, file_path, data_type, period, year, column_mapping, sheet_names=None, header_row=0, progress_callback=None, add_timestamp=True):
        """データ取り込み"""
        try:
            # ファイルコピー
            copied_file = self.file_manager.copy_import_file(
                file_path, data_type, period, year, add_timestamp
            )
            
            # Excel読み込み
            excel_file = pd.ExcelFile(file_path)
            
            # シート名取得
            if sheet_names is None:
                sheet_names = excel_file.sheet_names
            
            total_sheets = len(sheet_names)
            total_rows = 0
            
            for i, sheet_name in enumerate(sheet_names):
                if progress_callback:
                    progress_callback(i, total_sheets, f"シート処理中: {sheet_name}")
                
                # シート読み込み（header_row指定）
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
                
                # カラム名変更
                df = df.rename(columns=column_mapping)
                
                # データ型に応じた処理
                if data_type == '評定':
                    rows = self.import_grades(df, period, year)
                elif data_type == '観点':
                    rows = self.import_viewpoints(df, period, year)
                elif data_type == '欠課情報':
                    rows = self.import_absences(df, period, year)
                else:
                    raise ValueError(f"未対応のデータ型: {data_type}")
                
                total_rows += rows
            
            # ログ記録
            self.logger.log_action(
                'data_import',
                f"{data_type} - {period} {year}年度 - {total_rows}件"
            )
            
            if progress_callback:
                progress_callback(total_sheets, total_sheets, "取り込み完了")
            
            return True
            
        except Exception as e:
            print(f"データ取り込みエラー: {e}")
            self.logger.log_action(
                'data_import_error',
                f"{data_type} - {str(e)}"
            )
            raise
    
    def import_grades(self, df, period, year):
        """評定データ取り込み"""
        try:
            required_columns = ['student_number', 'course_number']
            
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"必須カラムがありません: {col}")
            
            # データクリーニング
            df = df.dropna(subset=required_columns)
            
            # period と year を列として追加
            df = df.copy()
            df['period'] = period
            df['year'] = year
            
            # 必要な列のみ選択（テーブル定義と一致させる）
            columns_order = [
                'year', 'period', 'student_number', 'student_name',
                'course_number', 'course_name', 'school_subject_name',
                'grade_value', 'credits', 'acquisition_credits', 'remarks'
            ]
            
            # 存在しない列は None で埋める
            for col in columns_order:
                if col not in df.columns:
                    df[col] = None
            
            df_to_insert = df[columns_order]
            
            # 既存データを削除（INSERT OR REPLACE の代替）
            self.db.execute_query(
                "DELETE FROM grades WHERE period=? AND year=?",
                (period, year)
            )
            
            # 一括INSERT
            df_to_insert.to_sql(
                'grades',
                self.db.get_connection(),
                if_exists='append',
                index=False,
                method='multi'
            )
            
            return len(df_to_insert)
            
        except Exception as e:
            print(f"評定データ取り込みエラー: {e}")
            raise
    
    def import_viewpoints(self, df, period, year):
        """観点データ取り込み"""
        try:
            required_columns = ['student_number', 'course_number']
            
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"必須カラムがありません: {col}")
            
            # データクリーニング
            df = df.dropna(subset=required_columns)
            
            # period と year を列として追加
            df = df.copy()
            df['period'] = period
            df['year'] = year
            
            # 必要な列のみ選択（テーブル定義と一致させる）
            columns_order = [
                'year', 'period', 'student_number', 'student_name',
                'course_number', 'course_name', 'school_subject_name',
                'viewpoint_1', 'viewpoint_2', 'viewpoint_3',
                'viewpoint_4', 'viewpoint_5', 'remarks'
            ]
            
            # 存在しない列は None で埋める
            for col in columns_order:
                if col not in df.columns:
                    df[col] = None
            
            df_to_insert = df[columns_order]
            
            # 既存データを削除（INSERT OR REPLACE の代替）
            self.db.execute_query(
                "DELETE FROM viewpoint_evaluations WHERE period=? AND year=?",
                (period, year)
            )
            
            # 一括INSERT
            df_to_insert.to_sql(
                'viewpoint_evaluations',
                self.db.get_connection(),
                if_exists='append',
                index=False,
                method='multi'
            )
            
            return len(df_to_insert)
            
        except Exception as e:
            print(f"観点データ取り込みエラー: {e}")
            raise
    
    def import_absences(self, df, period, year):
        """欠課情報取り込み (db_columns.json に基づく)"""
        try:
            required_columns = ['student_number']
            
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"必須カラムがありません: {col}")
            
            # データクリーニング
            df = df.dropna(subset=required_columns)
            
            # period と year を列として追加
            df = df.copy()
            df['period'] = period
            df['year'] = year
            
            # 必要な列のみ選択（テーブル定義と一致させる）
            columns_order = [
                'student_number', 'class_name', 'attendance_number', 'student_name',
                'absent_count', 'course_name', 'subject_category_number', 'subject_number',
                'course_number', 'year', 'period', 'absence_mark', 'absence_type'
            ]
            
            # 存在しない列は None で埋める
            for col in columns_order:
                if col not in df.columns:
                    df[col] = None
            
            df_to_insert = df[columns_order]
            
            # 既存データを削除（INSERT OR REPLACE の代替）
            self.db.execute_query(
                "DELETE FROM absences WHERE period=? AND year=?",
                (period, year)
            )
            
            # 一括INSERT
            df_to_insert.to_sql(
                'absences',
                self.db.get_connection(),
                if_exists='append',
                index=False,
                method='multi'
            )
            
            return len(df_to_insert)
            
        except Exception as e:
            print(f"欠課情報取り込みエラー: {e}")
            raise
