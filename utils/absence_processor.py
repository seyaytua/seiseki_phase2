import pandas as pd
from pathlib import Path
from datetime import datetime


class AbsenceProcessor:
    """欠課データ前処理クラス"""
    
    def __init__(self):
        self.result_df = None
        self.debug_info = []
    
    def process_multiple_files(self, file_paths, header_row=0, column_mapping=None, progress_callback=None):
        """複数ファイルを処理して欠課データを集計"""
        all_data = []  # 全データ（欠課あり・なし含む）
        
        total_files = len(file_paths)
        print(f"\n{'='*70}")
        print(f"処理開始: {total_files}ファイル")
        print(f"{'='*70}")
        
        for idx, file_path in enumerate(file_paths):
            file_name = Path(file_path).name
            
            # 進捗コールバック
            if progress_callback:
                progress_callback(idx, total_files, f"処理中 ({idx+1}/{total_files}): {file_name}")
            
            print(f"\n[{idx+1}/{total_files}] {file_name}")
            print(f"{'-'*70}")
            
            try:
                # Excelファイル読み込み（全シート）
                excel_file = pd.ExcelFile(file_path)
                total_sheets = len(excel_file.sheet_names)
                
                print(f"シート数: {total_sheets}枚")
                
                file_data = []
                file_total_rows = 0
                file_absence_count = 0
                
                for sheet_idx, sheet_name in enumerate(excel_file.sheet_names):
                    # シート処理の進捗表示
                    print(f" [{sheet_idx+1:2d}/{total_sheets:2d}] {sheet_name:<30s} ... ", end='', flush=True)
                    
                    try:
                        # シート読み込み
                        df = pd.read_excel(
                            file_path,
                            sheet_name=sheet_name,
                            header=header_row
                        )
                        
                        # 空のシートはスキップ
                        if len(df) == 0:
                            print("スキップ (空シート)")
                            continue
                        
                        # カラムマッピング適用
                        if column_mapping:
                            df = df.rename(columns=column_mapping)
                        
                        # 必須カラムチェック
                        required_cols = ['student_number', 'course_number']
                        missing_cols = [col for col in required_cols if col not in df.columns]
                        
                        if missing_cols:
                            print(f"スキップ (カラム不足: {missing_cols})")
                            continue
                        
                        # NaNの行を除外
                        original_len = len(df)
                        df = df.dropna(subset=['student_number', 'course_number'])
                        
                        if len(df) == 0:
                            print(f"スキップ (有効データなし, 元{original_len}行)")
                            continue
                        
                        # ファイル名とシート名を追加
                        df['source_file'] = file_name
                        df['sheet_name'] = sheet_name
                        
                        # 欠課フラグを追加
                        df['is_absence'] = self.check_absence(df)
                        
                        absence_count = df['is_absence'].sum()
                        print(f"OK ({len(df):5d}行, 欠課{absence_count:4d}件)")
                        
                        file_data.append(df)
                        file_total_rows += len(df)
                        file_absence_count += absence_count
                    
                    except Exception as e:
                        print(f"エラー: {str(e)[:50]}")
                        continue
                
                # ファイル単位での結合
                if file_data:
                    file_df = pd.concat(file_data, ignore_index=True)
                    all_data.append(file_df)
                    print(f"\n ファイル合計: {file_total_rows:,}行, 欠課{file_absence_count:,}件")
                    self.debug_info.append(f"{file_name}: {file_total_rows:,}件読み込み (欠課{file_absence_count:,}件)")
                else:
                    print(f"\n ファイル合計: データなし")
                    self.debug_info.append(f"{file_name}: データなし")
            
            except Exception as e:
                print(f"\nファイル処理エラー: {e}")
                import traceback
                traceback.print_exc()
                self.debug_info.append(f"{file_name}: エラー - {str(e)}")
                continue
        
        # データが見つからない場合
        if not all_data:
            print(f"\n{'='*70}")
            print("エラー: 有効なデータが見つかりませんでした")
            print(f"{'='*70}")
            return None
        
        # 全データを結合
        print(f"\n{'='*70}")
        print("全データ結合中...")
        print(f"{'='*70}")
        
        combined_df = pd.concat(all_data, ignore_index=True)
        total_records = len(combined_df)
        total_absences = combined_df['is_absence'].sum()
        total_attendances = (~combined_df['is_absence']).sum()
        
        print(f"総データ件数: {total_records:,}件")
        print(f" 欠課データ: {total_absences:,}件 ({total_absences/total_records*100:.1f}%)")
        print(f" 出席データ: {total_attendances:,}件 ({total_attendances/total_records*100:.1f}%)")
        
        # 生徒×講座で集計
        print(f"\n集計処理中...")
        aggregated_df = self.aggregate_by_student_course(combined_df)
        
        self.result_df = aggregated_df
        
        print(f"\n{'='*70}")
        print("処理完了")
        print(f"{'='*70}")
        
        return self.result_df
    
    def check_absence(self, df):
        """欠課判定"""
        absence_condition = pd.Series([False] * len(df))
        
        # absence_markで判定
        if 'absence_mark' in df.columns:
            mark_condition = df['absence_mark'].astype(str).str.contains('/', na=False)
            absence_condition |= mark_condition
        
        # absence_typeで判定
        if 'absence_type' in df.columns:
            type_condition = df['absence_type'] == 1
            absence_condition |= type_condition
        
        return absence_condition
    
    def aggregate_by_student_course(self, df):
        """生徒×講座で集計（実際の履修組み合わせのみ）"""
        print(f"\n{'='*70}")
        print("集計処理詳細")
        print(f"{'='*70}")
        print(f"集計前の行数: {len(df):,}件")
        
        # 生徒×講座の組み合わせごとに集計
        grouped = df.groupby(['student_number', 'course_number']).agg({
            'is_absence': 'sum',  # 欠課数
            'class_name': 'first',
            'attendance_number': 'first',
            'student_name': 'first',
            'course_name': 'first',
            'subject_category_number': 'first',
            'subject_number': 'first'
        }).reset_index()
        
        # カラム名を変更
        grouped = grouped.rename(columns={'is_absence': 'absent_count'})
        
        # absent_countを整数型に変換
        grouped['absent_count'] = grouped['absent_count'].astype(int)
        
        print(f"集計後の行数: {len(grouped):,}件")
        print(f" ユニーク生徒数: {grouped['student_number'].nunique():,}人")
        print(f" ユニーク講座数: {grouped['course_number'].nunique():,}科目")
        print(f" 欠課0の組み合わせ: {len(grouped[grouped['absent_count'] == 0]):,}件")
        print(f" 欠課ありの組み合わせ: {len(grouped[grouped['absent_count'] > 0]):,}件")
        
        # 欠課数の分布を表示
        print(f"\n【欠課数の分布】")
        absence_dist = grouped['absent_count'].value_counts().sort_index()
        for count, freq in absence_dist.head(10).items():
            print(f" 欠課{count:2d}回: {freq:5,}件")
        
        if len(absence_dist) > 10:
            print(f" ... (他 {len(absence_dist)-10} パターン)")
        
        # 生徒あたりの平均履修講座数
        courses_per_student = grouped.groupby('student_number').size()
        print(f"\n【生徒あたりの履修講座数】")
        print(f" 平均: {courses_per_student.mean():.1f}講座")
        print(f" 最小: {courses_per_student.min()}講座")
        print(f" 最大: {courses_per_student.max()}講座")
        
        return grouped
    
    def get_summary(self):
        """処理結果のサマリーを取得"""
        if self.result_df is None or len(self.result_df) == 0:
            return {
                'total_records': 0,
                'unique_students': 0,
                'unique_courses': 0,
                'total_absences': 0,
                'average_absences': 0.0,
                'zero_absence_count': 0,
                'max_absences': 0,
                'courses_per_student': 0.0
            }
        
        courses_per_student = self.result_df.groupby('student_number').size().mean()
        
        return {
            'total_records': len(self.result_df),
            'unique_students': self.result_df['student_number'].nunique(),
            'unique_courses': self.result_df['course_number'].nunique(),
            'total_absences': int(self.result_df['absent_count'].sum()),
            'average_absences': round(self.result_df['absent_count'].mean(), 2),
            'zero_absence_count': len(self.result_df[self.result_df['absent_count'] == 0]),
            'max_absences': int(self.result_df['absent_count'].max()),
            'courses_per_student': round(courses_per_student, 1)
        }
    
    def get_debug_info(self):
        """デバッグ情報を取得"""
        return '\n'.join(self.debug_info)
    
    def export_to_excel(self, output_dir='output/preprocessed', selected_columns=None):
        """結果をExcel出力"""
        if self.result_df is None or len(self.result_df) == 0:
            raise ValueError("出力するデータがありません")
        
        # 出力ディレクトリ作成
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        # ファイル名生成
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'absence_preprocessed_{timestamp}.xlsx'
        filepath = output_path / filename
        
        # 出力するカラムを選択
        if selected_columns:
            # 選択されたカラムのみ出力
            available_columns = [col for col in selected_columns if col in self.result_df.columns]
            output_df = self.result_df[available_columns].copy()
        else:
            output_df = self.result_df.copy()
        
        # absent_countでソート（欠課が多い順）
        output_df = output_df.sort_values('absent_count', ascending=False)
        
        # Excel出力
        print(f"\n{'='*70}")
        print("Excel出力中...")
        print(f"{'='*70}")
        
        output_df.to_excel(filepath, index=False, sheet_name='欠課データ')
        
        print(f"出力完了: {filepath}")
        print(f" 出力レコード数: {len(output_df):,}件")
        print(f" 出力カラム数: {len(output_df.columns)}列")
        print(f" カラム: {', '.join(output_df.columns.tolist())}")
        
        return str(filepath)