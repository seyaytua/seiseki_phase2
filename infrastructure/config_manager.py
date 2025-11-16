"""
設定管理クラス

アプリケーション設定の保存・読み込みを管理
"""
import json
from pathlib import Path
from typing import Any, Optional, Dict


class ConfigManager:
    """設定管理クラス"""
    
    def __init__(self, config_file: str = "config.json"):
        """
        初期化
        
        Args:
            config_file: 設定ファイルのパス
        """
        self.config_file = Path(config_file)
        self.config: Dict[str, Any] = {}
        
        # 設定ファイルの読み込み
        self.load_config()
    
    def load_config(self):
        """設定ファイルの読み込み"""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                print("設定ファイルを読み込みました")
            except Exception as e:
                print(f"設定ファイル読み込みエラー: {e}")
                self.config = self._get_default_config()
        else:
            # デフォルト設定
            self.config = self._get_default_config()
            self.save_config()
    
    def save_config(self):
        """設定ファイルの保存"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            print("設定ファイルを保存しました")
        except Exception as e:
            print(f"設定ファイル保存エラー: {e}")
    
    def _get_default_config(self) -> dict:
        """
        デフォルト設定を取得
        
        Returns:
            dict: デフォルト設定
        """
        return {
            "column_mappings": {
                "評定": {},
                "観点": {},
                "欠課情報": {}
            },
            "last_import_dir": "",
            "last_export_dir": "",
            "window_size": [1200, 800],
            "window_position": None,
            "recent_files": [],
            "default_year": 2024,
            "default_period": "前期"
        }
    
    def get_column_mapping(self, data_type: str) -> Dict[str, str]:
        """
        カラムマッピング取得
        
        Args:
            data_type: データタイプ（'評定', '観点', '欠課情報'）
            
        Returns:
            Dict[str, str]: カラムマッピング
        """
        return self.config.get('column_mappings', {}).get(data_type, {})
    
    def save_column_mapping(self, data_type: str, mapping: Dict[str, str]):
        """
        カラムマッピング保存
        
        Args:
            data_type: データタイプ
            mapping: カラムマッピング
        """
        if 'column_mappings' not in self.config:
            self.config['column_mappings'] = {}
        
        self.config['column_mappings'][data_type] = mapping
        self.save_config()
    
    def get_setting(self, key: str, default: Any = None) -> Any:
        """
        設定値取得
        
        Args:
            key: 設定キー
            default: デフォルト値
            
        Returns:
            Any: 設定値
        """
        return self.config.get(key, default)
    
    def save_setting(self, key: str, value: Any):
        """
        設定値保存
        
        Args:
            key: 設定キー
            value: 設定値
        """
        self.config[key] = value
        self.save_config()
    
    def get_window_geometry(self) -> tuple:
        """
        ウィンドウ位置・サイズ取得
        
        Returns:
            tuple: (幅, 高さ, x座標, y座標)
        """
        size = self.config.get('window_size', [1200, 800])
        position = self.config.get('window_position')
        
        if position:
            return (size[0], size[1], position[0], position[1])
        else:
            return (size[0], size[1], None, None)
    
    def save_window_geometry(self, width: int, height: int, x: int, y: int):
        """
        ウィンドウ位置・サイズ保存
        
        Args:
            width: 幅
            height: 高さ
            x: x座標
            y: y座標
        """
        self.config['window_size'] = [width, height]
        self.config['window_position'] = [x, y]
        self.save_config()
    
    def add_recent_file(self, file_path: str, max_count: int = 10):
        """
        最近使用したファイルを追加
        
        Args:
            file_path: ファイルパス
            max_count: 最大保存件数
        """
        recent = self.config.get('recent_files', [])
        
        # 既存のエントリを削除
        if file_path in recent:
            recent.remove(file_path)
        
        # 先頭に追加
        recent.insert(0, file_path)
        
        # 件数制限
        recent = recent[:max_count]
        
        self.config['recent_files'] = recent
        self.save_config()
    
    def get_recent_files(self) -> list:
        """
        最近使用したファイル一覧取得
        
        Returns:
            list: ファイルパスのリスト
        """
        return self.config.get('recent_files', [])
    
    def get_default_header_row(self, data_type: str) -> int:
        """
        デフォルトヘッダー行取得
        
        Args:
            data_type: データタイプ
            
        Returns:
            int: ヘッダー行番号（0始まり）
        """
        key = f"default_header_row_{data_type}"
        return self.config.get(key, 0)
    
    def save_default_header_row(self, data_type: str, row: int):
        """
        デフォルトヘッダー行保存
        
        Args:
            data_type: データタイプ
            row: ヘッダー行番号
        """
        key = f"default_header_row_{data_type}"
        self.config[key] = row
        self.save_config()
    
    def get_db_columns(self, data_type: str) -> list:
        """
        データベースカラム一覧取得
        
        Args:
            data_type: データタイプ
            
        Returns:
            list: カラム名のリスト
        """
        columns_map = {
            '評定': [
                'student_number', 'student_name', 'course_number', 'course_name',
                'school_subject_name', 'grade_value', 'credits', 
                'acquisition_credits', 'remarks'
            ],
            '観点': [
                'student_number', 'student_name', 'course_number', 'course_name',
                'school_subject_name', 'viewpoint_1', 'viewpoint_2', 'viewpoint_3',
                'viewpoint_4', 'viewpoint_5', 'remarks'
            ],
            '欠課情報': [
                'student_number', 'student_name', 'course_number', 'course_name',
                'school_subject_name', 'absent_count', 'late_count', 'total_hours',
                'absence_rate', 'remarks'
            ]
        }
        
        return columns_map.get(data_type, [])
    
    def get_required_columns(self, data_type: str) -> list:
        """
        必須カラム一覧取得
        
        Args:
            data_type: データタイプ
            
        Returns:
            list: 必須カラム名のリスト
        """
        # すべてのデータタイプで共通の必須カラム
        return ['student_number', 'course_number']
