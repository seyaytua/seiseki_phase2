import json
from pathlib import Path


class ConfigManager:
    """設定管理クラス"""
    
    def __init__(self, config_dir=None):
        """初期化"""
        if config_dir is None:
            # デフォルトの設定ディレクトリ
            self.config_dir = Path(__file__).parent.parent / 'config'
        else:
            self.config_dir = Path(config_dir)
        
        self.config_dir.mkdir(parents=True, exist_ok=True)
        
        # 設定ファイルパス
        self.config_file = self.config_dir / 'app_config.json'
        self.mapping_file = self.config_dir / 'column_mappings.json'
        
        # 設定読み込み
        self.config = self.load_config()
        self.mappings = self.load_mappings()
    
    def load_config(self):
        """設定ファイル読み込み"""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"設定ファイル読み込みエラー: {e}")
                return self.get_default_config()
        else:
            return self.get_default_config()
    
    def get_default_config(self):
        """デフォルト設定取得"""
        return {
            'window': {
                'width': 1200,
                'height': 800,
                'maximized': False
            },
            'paths': {
                'last_import_dir': '',
                'last_export_dir': ''
            },
            'import': {
                'add_timestamp': True,
                'auto_backup': True
            }
        }
    
    def save_config(self):
        """設定ファイル保存"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            print(f"設定ファイル保存エラー: {e}")
            return False
    
    def get_config(self, key, default=None):
        """設定値取得"""
        keys = key.split('.')
        value = self.config
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        
        return value
    
    def set_config(self, key, value):
        """設定値設定"""
        keys = key.split('.')
        config = self.config
        
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        
        config[keys[-1]] = value
        self.save_config()
    
    def load_mappings(self):
        """カラムマッピング読み込み"""
        if self.mapping_file.exists():
            try:
                with open(self.mapping_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"マッピングファイル読み込みエラー: {e}")
                return {}
        else:
            return {}
    
    def save_mappings(self):
        """カラムマッピング保存"""
        try:
            with open(self.mapping_file, 'w', encoding='utf-8') as f:
                json.dump(self.mappings, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            print(f"マッピングファイル保存エラー: {e}")
            return False
    
    def get_column_mapping(self, data_type):
        """カラムマッピング取得"""
        return self.mappings.get(data_type, {})
    
    def save_column_mapping(self, data_type, mapping):
        """カラムマッピング保存"""
        self.mappings[data_type] = mapping
        return self.save_mappings()
    
    def delete_column_mapping(self, data_type):
        """カラムマッピング削除"""
        if data_type in self.mappings:
            del self.mappings[data_type]
            return self.save_mappings()
        return False
    
    def get_all_mappings(self):
        """全カラムマッピング取得"""
        return self.mappings.copy()
    
    def get_window_geometry(self):
        """ウィンドウ位置・サイズ取得"""
        return {
            'width': self.get_config('window.width', 1200),
            'height': self.get_config('window.height', 800),
            'maximized': self.get_config('window.maximized', False)
        }
    
    def save_window_geometry(self, width, height, maximized):
        """ウィンドウ位置・サイズ保存"""
        self.set_config('window.width', width)
        self.set_config('window.height', height)
        self.set_config('window.maximized', maximized)
    
    def get_last_import_dir(self):
        """最後の取り込みディレクトリ取得"""
        return self.get_config('paths.last_import_dir', '')
    
    def set_last_import_dir(self, directory):
        """最後の取り込みディレクトリ設定"""
        self.set_config('paths.last_import_dir', directory)
    
    def get_last_export_dir(self):
        """最後のエクスポートディレクトリ取得"""
        return self.get_config('paths.last_export_dir', '')
    
    def set_last_export_dir(self, directory):
        """最後のエクスポートディレクトリ設定"""
        self.set_config('paths.last_export_dir', directory)    
    def get_settings(self):
        """settings.jsonファイルを読み込み"""
        settings_path = self.config_dir / 'settings.json'
        if settings_path.exists():
            try:
                with open(settings_path, 'r', encoding='utf-8-sig') as f:
                    return json.load(f)
            except Exception as e:
                print(f"settings.json読み込みエラー: {e}")
                return {}
        return {}
    
    def save_settings(self, settings):
        """settings.jsonファイルに保存"""
        settings_path = self.config_dir / 'settings.json'
        try:
            with open(settings_path, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            print(f"settings.json保存エラー: {e}")
            return False
