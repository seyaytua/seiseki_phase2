"""
データエクスポートサービス

データベースからExcelファイルへのデータ出力
"""
import pandas as pd
from typing import Optional, List
from datetime import datetime

from infrastructure.database_manager import DatabaseManager
from infrastructure.file_manager import FileManager
from infrastructure.logger import Logger


class DataExportService:
    """データエクスポートサービス"""
    
    def __init__(self,
                 db_manager: DatabaseManager,
                 file_manager: FileManager,
                 logger: Logger):
        """
        初期化
        
        Args:
            db_manager: データベースマネージャー
            file_manager: ファイルマネージャー
            logger: ログマネージャー
        """
        self.db = db_manager
        self.file_manager = file_manager
        self.logger = logger
    
    def export_data(self,
                   data_type: str,
                   period: Optional[str] = None,
                   year: Optional[int] = None,
                   student_number: Optional[str] = None,
                   course_number: Optional[str] = None,
                   filename: str = "export.xlsx",
                   add_timestamp: bool = True) -> tuple[bool, str]:
        """
        データ出力メイン処理
        
        Args:
            data_type: データタイプ（'評定', '観点', '欠課情報'）
            period: 期間（オプション）
            year: 年度（オプション）
            student_number: 学籍番号（オプション）
            course_number: 科目番号（オプション）
            filename: 出力ファイル名
            add_timestamp: ファイル名にタイムスタンプを追加するか
            
        Returns:
            tuple[bool, str]: (成功/失敗, メッセージ)
        """
        try:
            # データ取得
            if data_type == '評定':
                df = self._get_grades_dataframe(period, year, student_number, course_number)
            elif data_type == '観点':
                df = self._get_viewpoints_dataframe(period, year, student_number, course_number)
            elif data_type == '欠課情報':
                df = self._get_absences_dataframe(period, year, student_number, course_number)
            else:
                return False, f"無効なデータタイプ: {data_type}"
            
            if df is None or len(df) == 0:
                return False, "出力するデータがありません"
            
            # 出力ファイルパス生成
            export_path = self.file_manager.create_export_file(filename, add_timestamp)
            
            # Excel出力
            df.to_excel(export_path, index=False, sheet_name=data_type)
            
            # ログ記録
            self.logger.log_action(
                Logger.ACTION_EXPORT,
                f"{data_type} - {len(df)}件 - {export_path.name}"
            )
            
            return True, f"{len(df)}件のデータを出力しました\n保存先: {export_path}"
            
        except Exception as e:
            error_msg = f"データ出力エラー: {str(e)}"
            print(error_msg)
            self.logger.log_action(Logger.ACTION_ERROR, error_msg)
            return False, error_msg
    
    def _get_grades_dataframe(self,
                             period: Optional[str],
                             year: Optional[int],
                             student_number: Optional[str],
                             course_number: Optional[str]) -> Optional[pd.DataFrame]:
        """
        評定データをDataFrameとして取得
        
        Args:
            period: 期間
            year: 年度
            student_number: 学籍番号
            course_number: 科目番号
            
        Returns:
            Optional[pd.DataFrame]: 評定データ
        """
        query = """
            SELECT 
                student_number AS 学籍番号,
                student_name AS 氏名,
                course_number AS 科目番号,
                course_name AS 科目名,
                school_subject_name AS 教科名,
                period AS 期間,
                year AS 年度,
                grade_value AS 評定,
                credits AS 単位数,
                acquisition_credits AS 修得単位数,
                remarks AS 備考
            FROM grades
            WHERE 1=1
        """
        params = []
        
        if period:
            query += " AND period = ?"
            params.append(period)
        
        if year:
            query += " AND year = ?"
            params.append(year)
        
        if student_number:
            query += " AND student_number = ?"
            params.append(student_number)
        
        if course_number:
            query += " AND course_number = ?"
            params.append(course_number)
        
        query += " ORDER BY year DESC, period, student_number, course_number"
        
        try:
            results = self.db.fetch_all(query, tuple(params) if params else None)
            if results:
                columns = ['学籍番号', '氏名', '科目番号', '科目名', '教科名', 
                          '期間', '年度', '評定', '単位数', '修得単位数', '備考']
                return pd.DataFrame(results, columns=columns)
            return None
        except Exception as e:
            print(f"データ取得エラー: {e}")
            return None
    
    def _get_viewpoints_dataframe(self,
                                 period: Optional[str],
                                 year: Optional[int],
                                 student_number: Optional[str],
                                 course_number: Optional[str]) -> Optional[pd.DataFrame]:
        """
        観点別評価データをDataFrameとして取得
        
        Args:
            period: 期間
            year: 年度
            student_number: 学籍番号
            course_number: 科目番号
            
        Returns:
            Optional[pd.DataFrame]: 観点別評価データ
        """
        query = """
            SELECT 
                student_number AS 学籍番号,
                student_name AS 氏名,
                course_number AS 科目番号,
                course_name AS 科目名,
                school_subject_name AS 教科名,
                period AS 期間,
                year AS 年度,
                viewpoint_1 AS 知識技能,
                viewpoint_2 AS 思考判断表現,
                viewpoint_3 AS 主体的学習態度,
                viewpoint_4 AS 観点4,
                viewpoint_5 AS 観点5,
                remarks AS 備考
            FROM viewpoint_evaluations
            WHERE 1=1
        """
        params = []
        
        if period:
            query += " AND period = ?"
            params.append(period)
        
        if year:
            query += " AND year = ?"
            params.append(year)
        
        if student_number:
            query += " AND student_number = ?"
            params.append(student_number)
        
        if course_number:
            query += " AND course_number = ?"
            params.append(course_number)
        
        query += " ORDER BY year DESC, period, student_number, course_number"
        
        try:
            results = self.db.fetch_all(query, tuple(params) if params else None)
            if results:
                columns = ['学籍番号', '氏名', '科目番号', '科目名', '教科名',
                          '期間', '年度', '知識技能', '思考判断表現', '主体的学習態度',
                          '観点4', '観点5', '備考']
                return pd.DataFrame(results, columns=columns)
            return None
        except Exception as e:
            print(f"データ取得エラー: {e}")
            return None
    
    def _get_absences_dataframe(self,
                               period: Optional[str],
                               year: Optional[int],
                               student_number: Optional[str],
                               course_number: Optional[str]) -> Optional[pd.DataFrame]:
        """
        欠課情報データをDataFrameとして取得
        
        Args:
            period: 期間
            year: 年度
            student_number: 学籍番号
            course_number: 科目番号
            
        Returns:
            Optional[pd.DataFrame]: 欠課情報データ
        """
        query = """
            SELECT 
                student_number AS 学籍番号,
                student_name AS 氏名,
                course_number AS 科目番号,
                course_name AS 科目名,
                school_subject_name AS 教科名,
                period AS 期間,
                year AS 年度,
                absent_count AS 欠席回数,
                late_count AS 遅刻回数,
                total_hours AS 総時数,
                absence_rate AS 欠課率,
                remarks AS 備考
            FROM absences
            WHERE 1=1
        """
        params = []
        
        if period:
            query += " AND period = ?"
            params.append(period)
        
        if year:
            query += " AND year = ?"
            params.append(year)
        
        if student_number:
            query += " AND student_number = ?"
            params.append(student_number)
        
        if course_number:
            query += " AND course_number = ?"
            params.append(course_number)
        
        query += " ORDER BY year DESC, period, student_number, course_number"
        
        try:
            results = self.db.fetch_all(query, tuple(params) if params else None)
            if results:
                columns = ['学籍番号', '氏名', '科目番号', '科目名', '教科名',
                          '期間', '年度', '欠席回数', '遅刻回数', '総時数', '欠課率', '備考']
                return pd.DataFrame(results, columns=columns)
            return None
        except Exception as e:
            print(f"データ取得エラー: {e}")
            return None
