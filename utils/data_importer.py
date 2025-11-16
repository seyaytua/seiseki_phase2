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
            
            # データ挿入
            count = 0
            for _, row in df.iterrows():
                query = """ INSERT OR REPLACE INTO grades (student_number, student_name, course_number, course_name, school_subject_name, period, year, grade_value, credits, acquisition_credits, remarks) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) """
                params = (
                    row.get('student_number'),
                    row.get('student_name'),
                    row.get('course_number'),
                    row.get('course_name'),
                    row.get('school_subject_name'),
                    period,
                    year,
                    row.get('grade_value'),
                    row.get('credits'),
                    row.get('acquisition_credits'),
                    row.get('remarks')
                )
                self.db.execute_query(query, params)
                count += 1
            
            return count
            
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
            
            # データ挿入
            count = 0
            for _, row in df.iterrows():
                query = """ INSERT OR REPLACE INTO viewpoint_evaluations (student_number, student_name, course_number, course_name, school_subject_name, period, year, viewpoint_1, viewpoint_2, viewpoint_3, viewpoint_4, viewpoint_5, remarks) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) """
                params = (
                    row.get('student_number'),
                    row.get('student_name'),
                    row.get('course_number'),
                    row.get('course_name'),
                    row.get('school_subject_name'),
                    period,
                    year,
                    row.get('viewpoint_1'),
                    row.get('viewpoint_2'),
                    row.get('viewpoint_3'),
                    row.get('viewpoint_4'),
                    row.get('viewpoint_5'),
                    row.get('remarks')
                )
                self.db.execute_query(query, params)
                count += 1
            
            return count
            
        except Exception as e:
            print(f"観点データ取り込みエラー: {e}")
            raise
    
    def import_absences(self, df, period, year):
        """欠課情報取り込み"""
        try:
            required_columns = ['student_number', 'course_number']
            
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"必須カラムがありません: {col}")
            
            # データクリーニング
            df = df.dropna(subset=required_columns)
            
            # データ挿入
            count = 0
            for _, row in df.iterrows():
                query = """ INSERT OR REPLACE INTO absences (student_number, student_name, course_number, course_name, school_subject_name, period, year, absent_hours, late_hours, total_hours, absence_rate, remarks) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) """
                params = (
                    row.get('student_number'),
                    row.get('student_name'),
                    row.get('course_number'),
                    row.get('course_name'),
                    row.get('school_subject_name'),
                    period,
                    year,
                    row.get('absent_hours'),
                    row.get('late_hours'),
                    row.get('total_hours'),
                    row.get('absence_rate'),
                    row.get('remarks')
                )
                self.db.execute_query(query, params)
                count += 1
            
            return count
            
        except Exception as e:
            print(f"欠課情報取り込みエラー: {e}")
            raise