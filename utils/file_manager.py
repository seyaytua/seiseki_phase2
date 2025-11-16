import shutil
from pathlib import Path
from datetime import datetime


class FileManager:
    """ファイル管理クラス"""
    
    def __init__(self, config_manager):
        """初期化"""
        self.config_manager = config_manager
        
        # ベースディレクトリ
        self.base_dir = Path(__file__).parent.parent
        self.data_dir = self.base_dir / 'data'
        self.import_dir = self.data_dir / 'imports'
        self.export_dir = self.data_dir / 'exports'
        self.backup_dir = self.data_dir / 'backups'
        
        # ディレクトリ作成
        self.create_directories()
    
    def create_directories(self):
        """必要なディレクトリを作成"""
        directories = [
            self.data_dir,
            self.import_dir,
            self.export_dir,
            self.backup_dir
        ]
        
        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
    
    def get_import_path(self, data_type, period, year, filename, add_timestamp=True):
        """取り込みファイルパス生成"""
        # ディレクトリ作成
        type_dir = self.import_dir / data_type / str(year) / period
        type_dir.mkdir(parents=True, exist_ok=True)
        
        # ファイル名生成
        file_path = Path(filename)
        stem = file_path.stem
        suffix = file_path.suffix
        
        if add_timestamp:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            new_filename = f"{stem}_{timestamp}{suffix}"
        else:
            new_filename = filename
        
        return type_dir / new_filename
    
    def copy_import_file(self, source_path, data_type, period, year, add_timestamp=True):
        """取り込みファイルをコピー"""
        try:
            source = Path(source_path)
            
            if not source.exists():
                raise FileNotFoundError(f"ファイルが見つかりません: {source_path}")
            
            # コピー先パス生成
            dest_path = self.get_import_path(
                data_type, 
                period, 
                year, 
                source.name, 
                add_timestamp
            )
            
            # ファイルコピー
            shutil.copy2(source, dest_path)
            
            return dest_path
            
        except Exception as e:
            print(f"ファイルコピーエラー: {e}")
            raise
    
    def get_export_path(self, filename, add_timestamp=True):
        """エクスポートファイルパス生成"""
        file_path = Path(filename)
        stem = file_path.stem
        suffix = file_path.suffix
        
        if add_timestamp:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            new_filename = f"{stem}_{timestamp}{suffix}"
        else:
            new_filename = filename
        
        return self.export_dir / new_filename
    
    def create_backup(self, source_path, backup_type='manual'):
        """バックアップ作成"""
        try:
            source = Path(source_path)
            
            if not source.exists():
                raise FileNotFoundError(f"ファイルが見つかりません: {source_path}")
            
            # バックアップディレクトリ
            backup_subdir = self.backup_dir / backup_type
            backup_subdir.mkdir(parents=True, exist_ok=True)
            
            # バックアップファイル名
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_filename = f"{source.stem}_{timestamp}{source.suffix}"
            backup_path = backup_subdir / backup_filename
            
            # ファイルコピー
            shutil.copy2(source, backup_path)
            
            return backup_path
            
        except Exception as e:
            print(f"バックアップ作成エラー: {e}")
            raise
    
    def list_import_files(self, data_type=None, period=None, year=None):
        """取り込みファイル一覧取得"""
        files = []
        
        search_dir = self.import_dir
        
        if data_type:
            search_dir = search_dir / data_type
            
            if year:
                search_dir = search_dir / str(year)
                
                if period:
                    search_dir = search_dir / period
        
        if search_dir.exists():
            for file_path in search_dir.rglob('*'):
                if file_path.is_file():
                    files.append({
                        'path': file_path,
                        'name': file_path.name,
                        'size': file_path.stat().st_size,
                        'modified': datetime.fromtimestamp(file_path.stat().st_mtime)
                    })
        
        return sorted(files, key=lambda x: x['modified'], reverse=True)
    
    def list_backups(self, backup_type=None):
        """バックアップファイル一覧取得"""
        files = []
        
        search_dir = self.backup_dir
        
        if backup_type:
            search_dir = search_dir / backup_type
        
        if search_dir.exists():
            for file_path in search_dir.rglob('*'):
                if file_path.is_file():
                    files.append({
                        'path': file_path,
                        'name': file_path.name,
                        'size': file_path.stat().st_size,
                        'modified': datetime.fromtimestamp(file_path.stat().st_mtime)
                    })
        
        return sorted(files, key=lambda x: x['modified'], reverse=True)
    
    def delete_file(self, file_path):
        """ファイル削除"""
        try:
            path = Path(file_path)
            
            if path.exists():
                path.unlink()
                return True
            else:
                return False
                
        except Exception as e:
            print(f"ファイル削除エラー: {e}")
            raise
    
    def cleanup_old_backups(self, days=30, backup_type=None):
        """古いバックアップを削除"""
        try:
            deleted_count = 0
            cutoff_date = datetime.now().timestamp() - (days * 24 * 60 * 60)
            
            search_dir = self.backup_dir
            
            if backup_type:
                search_dir = search_dir / backup_type
            
            if search_dir.exists():
                for file_path in search_dir.rglob('*'):
                    if file_path.is_file():
                        if file_path.stat().st_mtime < cutoff_date:
                            file_path.unlink()
                            deleted_count += 1
            
            return deleted_count
            
        except Exception as e:
            print(f"バックアップクリーンアップエラー: {e}")
            raise
    
    def get_directory_size(self, directory):
        """ディレクトリサイズ取得"""
        try:
            total_size = 0
            dir_path = Path(directory)
            
            if dir_path.exists():
                for file_path in dir_path.rglob('*'):
                    if file_path.is_file():
                        total_size += file_path.stat().st_size
            
            return total_size
            
        except Exception as e:
            print(f"ディレクトリサイズ取得エラー: {e}")
            return 0
    
    def format_file_size(self, size_bytes):
        """ファイルサイズをフォーマット"""
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} PB"
    
    def get_storage_info(self):
        """ストレージ情報取得"""
        return {
            'import_size': self.get_directory_size(self.import_dir),
            'export_size': self.get_directory_size(self.export_dir),
            'backup_size': self.get_directory_size(self.backup_dir),
            'total_size': self.get_directory_size(self.data_dir)
        }