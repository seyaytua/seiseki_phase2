"""
データインポートサービス

Excelファイルからのデータ取り込み処理
"""
import pandas as pd
from typing import Callable, Optional, List
from pathlib import Path

from models import Grade, ViewpointEvaluation, Absence
from infrastructure.database_manager import DatabaseManager
from infrastructure.file_manager import FileManager
from infrastructure.logger import Logger


class DataImportService:
    """データインポートサービス"""
    
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
    
    def import_data(self,
                   file_path: str,
                   data_type: str,
                   period: str,
                   year: int,
                   column_mapping: dict,
                   sheet_names: List[str],
                   header_row: int = 0,
                   progress_callback: Optional[Callable] = None,
                   add_timestamp: bool = True) -> tuple[bool, str, int]:
        """
        データ取り込みメイン処理
        
        Args:
            file_path: Excelファイルパス
            data_type: データタイプ（'評定', '観点', '欠課情報'）
            period: 期間
            year: 年度
            column_mapping: カラムマッピング
            sheet_names: 取り込むシート名のリスト
            header_row: ヘッダー行番号
            progress_callback: 進捗コールバック関数
            add_timestamp: ファイル名にタイムスタンプを追加するか
            
        Returns:
            tuple[bool, str, int]: (成功/失敗, メッセージ, 取り込み件数)
        """
        try:
            # ファイルバックアップ
            copied_file = self.file_manager.copy_import_file(
                file_path, data_type, period, year, add_timestamp
            )
            
            total_rows = 0
            
            # 各シートを処理
            for i, sheet_name in enumerate(sheet_names):
                if progress_callback:
                    progress_callback(i, len(sheet_names), f"処理中: {sheet_name}")
                
                # シート読み込み
                df = pd.read_excel(
                    file_path,
                    sheet_name=sheet_name,
                    header=header_row
                )
                
                # カラム名変更（マッピング適用）
                df = df.rename(columns=column_mapping)
                
                # データタイプ別処理
                if data_type == '評定':
                    rows = self._import_grades(df, period, year)
                elif data_type == '観点':
                    rows = self._import_viewpoints(df, period, year)
                elif data_type == '欠課情報':
                    rows = self._import_absences(df, period, year)
                else:
                    raise ValueError(f"無効なデータタイプ: {data_type}")
                
                total_rows += rows
            
            # 進捗完了
            if progress_callback:
                progress_callback(len(sheet_names), len(sheet_names), "完了")
            
            # ログ記録
            self.logger.log_action(
                Logger.ACTION_IMPORT,
                f"{data_type} - {period} {year}年度 - {total_rows}件"
            )
            
            return True, f"{total_rows}件のデータを取り込みました", total_rows
            
        except Exception as e:
            error_msg = f"データ取り込みエラー: {str(e)}"
            print(error_msg)
            self.logger.log_action(Logger.ACTION_ERROR, error_msg)
            return False, error_msg, 0
    
    def _import_grades(self, df: pd.DataFrame, period: str, year: int) -> int:
        """
        評定データ取り込み
        
        Args:
            df: DataFrame
            period: 期間
            year: 年度
            
        Returns:
            int: 取り込み件数
        """
        # 必須カラムチェック
        required_columns = ['student_number', 'course_number']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"必須カラム不足: {col}")
        
        # NaN除去
        df_clean = df.dropna(subset=required_columns)
        
        count = 0
        errors = []
        
        # 1行ずつ処理
        for idx, row in df_clean.iterrows():
            try:
                # Gradeオブジェクト作成
                grade = Grade.from_dataframe_row(row, period, year)
                
                # バリデーション
                is_valid, error_msg = grade.validate()
                if not is_valid:
                    errors.append(f"行{idx + 2}: {error_msg}")
                    continue
                
                # データベースに保存
                self.db.insert_grade(grade.to_dict())
                count += 1
                
            except Exception as e:
                errors.append(f"行{idx + 2}: {str(e)}")
        
        # エラー表示
        if errors:
            print(f"エラー {len(errors)}件:")
            for error in errors[:10]:  # 最初の10件のみ表示
                print(f"  {error}")
        
        return count
    
    def _import_viewpoints(self, df: pd.DataFrame, period: str, year: int) -> int:
        """
        観点別評価データ取り込み
        
        Args:
            df: DataFrame
            period: 期間
            year: 年度
            
        Returns:
            int: 取り込み件数
        """
        # 必須カラムチェック
        required_columns = ['student_number', 'course_number']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"必須カラム不足: {col}")
        
        # NaN除去
        df_clean = df.dropna(subset=required_columns)
        
        count = 0
        errors = []
        
        # 1行ずつ処理
        for idx, row in df_clean.iterrows():
            try:
                # ViewpointEvaluationオブジェクト作成
                viewpoint = ViewpointEvaluation.from_dataframe_row(row, period, year)
                
                # バリデーション
                is_valid, error_msg = viewpoint.validate()
                if not is_valid:
                    errors.append(f"行{idx + 2}: {error_msg}")
                    continue
                
                # データベースに保存
                self.db.insert_viewpoint(viewpoint.to_dict())
                count += 1
                
            except Exception as e:
                errors.append(f"行{idx + 2}: {str(e)}")
        
        # エラー表示
        if errors:
            print(f"エラー {len(errors)}件:")
            for error in errors[:10]:
                print(f"  {error}")
        
        return count
    
    def _import_absences(self, df: pd.DataFrame, period: str, year: int) -> int:
        """
        欠課情報データ取り込み
        
        Args:
            df: DataFrame
            period: 期間
            year: 年度
            
        Returns:
            int: 取り込み件数
        """
        # 必須カラムチェック
        required_columns = ['student_number', 'course_number']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"必須カラム不足: {col}")
        
        # NaN除去
        df_clean = df.dropna(subset=required_columns)
        
        count = 0
        errors = []
        
        # 1行ずつ処理
        for idx, row in df_clean.iterrows():
            try:
                # Absenceオブジェクト作成
                absence = Absence.from_dataframe_row(row, period, year)
                
                # バリデーション
                is_valid, error_msg = absence.validate()
                if not is_valid:
                    errors.append(f"行{idx + 2}: {error_msg}")
                    continue
                
                # データベースに保存
                self.db.insert_absence(absence.to_dict())
                count += 1
                
            except Exception as e:
                errors.append(f"行{idx + 2}: {str(e)}")
        
        # エラー表示
        if errors:
            print(f"エラー {len(errors)}件:")
            for error in errors[:10]:
                print(f"  {error}")
        
        return count
    
    def validate_file(self, file_path: str) -> tuple[bool, str]:
        """
        ファイルの事前検証
        
        Args:
            file_path: ファイルパス
            
        Returns:
            tuple[bool, str]: (検証結果, メッセージ)
        """
        # ファイル存在チェック
        if not Path(file_path).exists():
            return False, "ファイルが見つかりません"
        
        # Excel形式チェック
        if not file_path.endswith(('.xlsx', '.xls')):
            return False, "Excelファイルを選択してください"
        
        # ファイル読み込みテスト
        try:
            excel_file = pd.ExcelFile(file_path)
            if len(excel_file.sheet_names) == 0:
                return False, "シートが見つかりません"
        except Exception as e:
            return False, f"ファイル読み込みエラー: {str(e)}"
        
        return True, "OK"
    
    def get_sheet_names(self, file_path: str) -> List[str]:
        """
        シート名一覧取得
        
        Args:
            file_path: ファイルパス
            
        Returns:
            List[str]: シート名のリスト
        """
        try:
            excel_file = pd.ExcelFile(file_path)
            return excel_file.sheet_names
        except Exception as e:
            print(f"シート名取得エラー: {e}")
            return []
    
    def preview_data(self, 
                    file_path: str, 
                    sheet_name: str, 
                    header_row: int = 0,
                    nrows: int = 10) -> Optional[pd.DataFrame]:
        """
        データプレビュー
        
        Args:
            file_path: ファイルパス
            sheet_name: シート名
            header_row: ヘッダー行番号
            nrows: プレビュー行数
            
        Returns:
            Optional[pd.DataFrame]: プレビューデータ
        """
        try:
            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                header=header_row,
                nrows=nrows
            )
            return df
        except Exception as e:
            print(f"プレビュー取得エラー: {e}")
            return None
