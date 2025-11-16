"""
評定データモデル

評定（成績）データを表現するクラス
"""
from dataclasses import dataclass
from typing import Optional
from datetime import datetime


@dataclass
class Grade:
    """評定データモデル"""
    
    # 必須フィールド
    student_number: str          # 学籍番号
    course_number: str           # 科目番号
    period: str                  # 期間（前期/後期/通年）
    year: int                    # 年度
    
    # オプションフィールド
    id: Optional[int] = None
    student_name: Optional[str] = None
    course_name: Optional[str] = None
    school_subject_name: Optional[str] = None
    grade_value: Optional[str] = None
    credits: Optional[float] = None
    acquisition_credits: Optional[float] = None
    remarks: Optional[str] = None
    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None
    
    @classmethod
    def from_dataframe_row(cls, row, period: str, year: int):
        """
        DataFrameの行からGradeオブジェクトを作成
        
        Args:
            row: pandas DataFrameの行
            period: 期間
            year: 年度
            
        Returns:
            Grade: 評定データオブジェクト
        """
        return cls(
            student_number=str(row.get('student_number', '')),
            student_name=row.get('student_name'),
            course_number=str(row.get('course_number', '')),
            course_name=row.get('course_name'),
            school_subject_name=row.get('school_subject_name'),
            period=period,
            year=year,
            grade_value=row.get('grade_value'),
            credits=cls._to_float(row.get('credits')),
            acquisition_credits=cls._to_float(row.get('acquisition_credits')),
            remarks=row.get('remarks')
        )
    
    @classmethod
    def from_db_row(cls, row: tuple):
        """
        データベースの行からGradeオブジェクトを作成
        
        Args:
            row: データベースクエリ結果の行（タプル）
            
        Returns:
            Grade: 評定データオブジェクト
        """
        return cls(
            id=row[0],
            student_number=row[1],
            student_name=row[2],
            course_number=row[3],
            course_name=row[4],
            school_subject_name=row[5],
            period=row[6],
            year=row[7],
            grade_value=row[8],
            credits=row[9],
            acquisition_credits=row[10],
            remarks=row[11],
            created_at=row[12] if len(row) > 12 else None,
            updated_at=row[13] if len(row) > 13 else None
        )
    
    def validate(self) -> tuple[bool, str]:
        """
        データのバリデーション
        
        Returns:
            tuple[bool, str]: (検証結果, エラーメッセージ)
        """
        if not self.student_number:
            return False, "学籍番号は必須です"
        
        if not self.course_number:
            return False, "科目番号は必須です"
        
        if not self.period:
            return False, "期間は必須です"
        
        if not self.year or self.year < 2000:
            return False, "有効な年度を指定してください"
        
        if self.period not in ['前期', '後期', '通年']:
            return False, "期間は「前期」「後期」「通年」のいずれかを指定してください"
        
        # 単位数の検証
        if self.credits is not None and self.credits < 0:
            return False, "単位数は0以上である必要があります"
        
        if self.acquisition_credits is not None and self.acquisition_credits < 0:
            return False, "修得単位数は0以上である必要があります"
        
        return True, ""
    
    def to_dict(self) -> dict:
        """
        辞書形式に変換
        
        Returns:
            dict: 評定データの辞書
        """
        return {
            'id': self.id,
            'student_number': self.student_number,
            'student_name': self.student_name,
            'course_number': self.course_number,
            'course_name': self.course_name,
            'school_subject_name': self.school_subject_name,
            'period': self.period,
            'year': self.year,
            'grade_value': self.grade_value,
            'credits': self.credits,
            'acquisition_credits': self.acquisition_credits,
            'remarks': self.remarks,
            'created_at': self.created_at,
            'updated_at': self.updated_at
        }
    
    @staticmethod
    def _to_float(value) -> Optional[float]:
        """
        値をfloatに変換（エラーハンドリング付き）
        
        Args:
            value: 変換する値
            
        Returns:
            Optional[float]: 変換後の値（変換不可の場合はNone）
        """
        if value is None or value == '':
            return None
        try:
            return float(value)
        except (ValueError, TypeError):
            return None
