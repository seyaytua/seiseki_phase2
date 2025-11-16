"""
観点別評価データモデル

観点別評価データを表現するクラス
"""
from dataclasses import dataclass
from typing import Optional
from datetime import datetime


@dataclass
class ViewpointEvaluation:
    """観点別評価データモデル"""
    
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
    viewpoint_1: Optional[str] = None  # 知識・技能
    viewpoint_2: Optional[str] = None  # 思考・判断・表現
    viewpoint_3: Optional[str] = None  # 主体的に学習に取り組む態度
    viewpoint_4: Optional[str] = None  # 予備
    viewpoint_5: Optional[str] = None  # 予備
    remarks: Optional[str] = None
    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None
    
    @classmethod
    def from_dataframe_row(cls, row, period: str, year: int):
        """
        DataFrameの行からViewpointEvaluationオブジェクトを作成
        
        Args:
            row: pandas DataFrameの行
            period: 期間
            year: 年度
            
        Returns:
            ViewpointEvaluation: 観点別評価データオブジェクト
        """
        return cls(
            student_number=str(row.get('student_number', '')),
            student_name=row.get('student_name'),
            course_number=str(row.get('course_number', '')),
            course_name=row.get('course_name'),
            school_subject_name=row.get('school_subject_name'),
            period=period,
            year=year,
            viewpoint_1=row.get('viewpoint_1'),
            viewpoint_2=row.get('viewpoint_2'),
            viewpoint_3=row.get('viewpoint_3'),
            viewpoint_4=row.get('viewpoint_4'),
            viewpoint_5=row.get('viewpoint_5'),
            remarks=row.get('remarks')
        )
    
    @classmethod
    def from_db_row(cls, row: tuple):
        """
        データベースの行からViewpointEvaluationオブジェクトを作成
        
        Args:
            row: データベースクエリ結果の行（タプル）
            
        Returns:
            ViewpointEvaluation: 観点別評価データオブジェクト
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
            viewpoint_1=row[8],
            viewpoint_2=row[9],
            viewpoint_3=row[10],
            viewpoint_4=row[11],
            viewpoint_5=row[12],
            remarks=row[13],
            created_at=row[14] if len(row) > 14 else None,
            updated_at=row[15] if len(row) > 15 else None
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
        
        # 観点評価値の検証（A, B, C のいずれか）
        valid_values = ['A', 'B', 'C', 'a', 'b', 'c', None, '']
        for i, viewpoint in enumerate([self.viewpoint_1, self.viewpoint_2, 
                                       self.viewpoint_3, self.viewpoint_4, 
                                       self.viewpoint_5], 1):
            if viewpoint and viewpoint not in valid_values:
                return False, f"観点{i}の評価値は A, B, C のいずれかを指定してください"
        
        return True, ""
    
    def to_dict(self) -> dict:
        """
        辞書形式に変換
        
        Returns:
            dict: 観点別評価データの辞書
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
            'viewpoint_1': self.viewpoint_1,
            'viewpoint_2': self.viewpoint_2,
            'viewpoint_3': self.viewpoint_3,
            'viewpoint_4': self.viewpoint_4,
            'viewpoint_5': self.viewpoint_5,
            'remarks': self.remarks,
            'created_at': self.created_at,
            'updated_at': self.updated_at
        }
