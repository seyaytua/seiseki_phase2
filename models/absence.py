"""
欠課情報データモデル

欠課情報データを表現するクラス
"""
from dataclasses import dataclass
from typing import Optional
from datetime import datetime


@dataclass
class Absence:
    """欠課情報データモデル"""
    
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
    absent_count: int = 0        # 欠席回数
    late_count: int = 0          # 遅刻回数
    total_hours: int = 0         # 総時数
    absence_rate: float = 0.0    # 欠課率
    remarks: Optional[str] = None
    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None
    
    @classmethod
    def from_dataframe_row(cls, row, period: str, year: int):
        """
        DataFrameの行からAbsenceオブジェクトを作成
        
        Args:
            row: pandas DataFrameの行
            period: 期間
            year: 年度
            
        Returns:
            Absence: 欠課情報データオブジェクト
        """
        absent_count = cls._to_int(row.get('absent_count'), 0)
        late_count = cls._to_int(row.get('late_count'), 0)
        total_hours = cls._to_int(row.get('total_hours'), 0)
        
        # 欠課率の計算
        absence_rate = 0.0
        if total_hours > 0:
            absence_rate = (absent_count / total_hours) * 100
        
        return cls(
            student_number=str(row.get('student_number', '')),
            student_name=row.get('student_name'),
            course_number=str(row.get('course_number', '')),
            course_name=row.get('course_name'),
            school_subject_name=row.get('school_subject_name'),
            period=period,
            year=year,
            absent_count=absent_count,
            late_count=late_count,
            total_hours=total_hours,
            absence_rate=absence_rate,
            remarks=row.get('remarks')
        )
    
    @classmethod
    def from_db_row(cls, row: tuple):
        """
        データベースの行からAbsenceオブジェクトを作成
        
        Args:
            row: データベースクエリ結果の行（タプル）
            
        Returns:
            Absence: 欠課情報データオブジェクト
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
            absent_count=row[8],
            late_count=row[9],
            total_hours=row[10],
            absence_rate=row[11],
            remarks=row[12],
            created_at=row[13] if len(row) > 13 else None,
            updated_at=row[14] if len(row) > 14 else None
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
        
        # 数値検証
        if self.absent_count < 0:
            return False, "欠席回数は0以上である必要があります"
        
        if self.late_count < 0:
            return False, "遅刻回数は0以上である必要があります"
        
        if self.total_hours < 0:
            return False, "総時数は0以上である必要があります"
        
        if self.total_hours > 0 and self.absent_count > self.total_hours:
            return False, "欠席回数は総時数を超えることはできません"
        
        return True, ""
    
    def calculate_absence_rate(self):
        """
        欠課率を計算して更新
        """
        if self.total_hours > 0:
            self.absence_rate = (self.absent_count / self.total_hours) * 100
        else:
            self.absence_rate = 0.0
    
    def to_dict(self) -> dict:
        """
        辞書形式に変換
        
        Returns:
            dict: 欠課情報データの辞書
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
            'absent_count': self.absent_count,
            'late_count': self.late_count,
            'total_hours': self.total_hours,
            'absence_rate': self.absence_rate,
            'remarks': self.remarks,
            'created_at': self.created_at,
            'updated_at': self.updated_at
        }
    
    @staticmethod
    def _to_int(value, default: int = 0) -> int:
        """
        値をintに変換（エラーハンドリング付き）
        
        Args:
            value: 変換する値
            default: 変換失敗時のデフォルト値
            
        Returns:
            int: 変換後の値
        """
        if value is None or value == '':
            return default
        try:
            return int(float(value))  # floatを経由してintに変換
        except (ValueError, TypeError):
            return default
