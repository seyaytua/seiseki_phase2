"""
ファイル管理クラス

ファイル操作（コピー、移動、削除）を管理
"""
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional, List


class FileManager:
    """ファイル管理クラス"""
    
    def __init__(self, data_dir: str = "data"):
        """
        初期化
        
        Args:
            data_dir: データディレクトリのパス
        """
        self.data_dir = Path(data_dir)
        self.imports_dir = self.data_dir / 'imports'
        self.exports_dir = self.data_dir / 'exports'
        
        # ディレクトリ作成
        self._create_directories()
    
    def _create_directories(self):
        """必要なディレクトリを作成"""
        # インポートディレクトリ
        for data_type in ['評定', '観点', '欠課情報']:
            (self.imports_dir / data_type).mkdir(parents=True, exist_ok=True)
        
        # エクスポートディレクトリ
        self.exports_dir.mkdir(parents=True, exist_ok=True)
    
    def copy_import_file(self, 
                        source_path: str, 
                        data_type: str, 
                        period: str, 
                        year: int, 
                        add_timestamp: bool = True) -> Path:
        """
        インポートファイルをバックアップ
        
        Args:
            source_path: ソースファイルのパス
            data_type: データタイプ（'評定', '観点', '欠課情報'）
            period: 期間
            year: 年度
            add_timestamp: タイムスタンプを追加するか
            
        Returns:
            Path: コピー先のパス
        """
        source_file = Path(source_path)
        
        # バックアップ先ディレクトリ
        backup_dir = self.imports_dir / data_type / str(year) / period
        backup_dir.mkdir(parents=True, exist_ok=True)
        
        # ファイル名生成
        if add_timestamp:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{source_file.stem}_{timestamp}{source_file.suffix}"
        else:
            filename = source_file.name
        
        dest_path = backup_dir / filename
        
        # ファイルコピー
        try:
            shutil.copy2(source_path, dest_path)
            print(f"ファイルをバックアップしました: {dest_path}")
            return dest_path
        except Exception as e:
            print(f"ファイルコピーエラー: {e}")
            raise
    
    def get_import_files(self, 
                        data_type: str, 
                        year: Optional[int] = None, 
                        period: Optional[str] = None) -> List[Path]:
        """
        インポートファイル一覧取得
        
        Args:
            data_type: データタイプ
            year: 年度（オプション）
            period: 期間（オプション）
            
        Returns:
            List[Path]: ファイルパスのリスト
        """
        base_dir = self.imports_dir / data_type
        
        if not base_dir.exists():
            return []
        
        # 検索パターン構築
        if year and period:
            search_dir = base_dir / str(year) / period
        elif year:
            search_dir = base_dir / str(year)
        else:
            search_dir = base_dir
        
        if not search_dir.exists():
            return []
        
        # Excelファイルを検索
        files = []
        for ext in ['.xlsx', '.xls']:
            files.extend(search_dir.glob(f'**/*{ext}'))
        
        return sorted(files, key=lambda x: x.stat().st_mtime, reverse=True)
    
    def create_export_file(self, 
                          filename: str, 
                          add_timestamp: bool = True) -> Path:
        """
        エクスポートファイルのパスを生成
        
        Args:
            filename: ファイル名
            add_timestamp: タイムスタンプを追加するか
            
        Returns:
            Path: エクスポートファイルのパス
        """
        file_path = Path(filename)
        
        if add_timestamp:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            export_filename = f"{file_path.stem}_{timestamp}{file_path.suffix}"
        else:
            export_filename = file_path.name
        
        export_path = self.exports_dir / export_filename
        
        return export_path
    
    def get_export_files(self, limit: int = 100) -> List[Path]:
        """
        エクスポートファイル一覧取得
        
        Args:
            limit: 取得件数上限
            
        Returns:
            List[Path]: ファイルパスのリスト
        """
        if not self.exports_dir.exists():
            return []
        
        # Excelファイルを検索
        files = []
        for ext in ['.xlsx', '.xls']:
            files.extend(self.exports_dir.glob(f'*{ext}'))
        
        # 更新日時でソート（新しい順）
        files = sorted(files, key=lambda x: x.stat().st_mtime, reverse=True)
        
        return files[:limit]
    
    def delete_file(self, file_path: Path) -> bool:
        """
        ファイル削除
        
        Args:
            file_path: 削除するファイルのパス
            
        Returns:
            bool: 削除成功かどうか
        """
        try:
            if file_path.exists():
                file_path.unlink()
                print(f"ファイルを削除しました: {file_path}")
                return True
            else:
                print(f"ファイルが見つかりません: {file_path}")
                return False
        except Exception as e:
            print(f"ファイル削除エラー: {e}")
            return False
    
    def get_file_info(self, file_path: Path) -> dict:
        """
        ファイル情報取得
        
        Args:
            file_path: ファイルパス
            
        Returns:
            dict: ファイル情報
        """
        if not file_path.exists():
            return {}
        
        stat = file_path.stat()
        
        return {
            'name': file_path.name,
            'path': str(file_path),
            'size': stat.st_size,
            'size_mb': round(stat.st_size / (1024 * 1024), 2),
            'created_at': datetime.fromtimestamp(stat.st_ctime),
            'modified_at': datetime.fromtimestamp(stat.st_mtime)
        }
    
    def clean_old_exports(self, days: int = 30) -> int:
        """
        古いエクスポートファイルを削除
        
        Args:
            days: 削除対象の日数（これより古いファイルを削除）
            
        Returns:
            int: 削除したファイル数
        """
        if not self.exports_dir.exists():
            return 0
        
        count = 0
        cutoff_time = datetime.now().timestamp() - (days * 24 * 60 * 60)
        
        for file_path in self.exports_dir.glob('*.xlsx'):
            if file_path.stat().st_mtime < cutoff_time:
                if self.delete_file(file_path):
                    count += 1
        
        print(f"{count}個の古いファイルを削除しました")
        return count
