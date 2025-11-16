from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTextEdit, 
                               QPushButton, QLabel)
from PySide6.QtCore import Qt, Signal, QObject
from PySide6.QtGui import QTextCursor, QFont
import sys
from io import StringIO


class LogStream(QObject):
    """標準出力をキャプチャしてシグナルで送信"""
    text_written = Signal(str)
    
    def __init__(self):
        super().__init__()
        self.buffer = StringIO()
    
    def write(self, text):
        self.buffer.write(text)
        self.text_written.emit(text)
    
    def flush(self):
        pass


class LogViewerDialog(QDialog):
    """ログ表示ダイアログ"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.setup_log_capture()
    
    def setup_ui(self):
        """UI初期化"""
        self.setWindowTitle("処理ログ")
        self.setMinimumSize(800, 600)
        
        layout = QVBoxLayout(self)
        
        # 説明ラベル
        info_label = QLabel("処理の詳細ログをリアルタイムで表示します")
        info_label.setStyleSheet("padding: 5px; background-color: #E3F2FD;")
        layout.addWidget(info_label)
        
        # ログ表示エリア
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setStyleSheet(""" QTextEdit { background-color: #1E1E1E; color: #D4D4D4; border: 1px solid #3C3C3C; } """)
        layout.addWidget(self.log_text)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        clear_btn = QPushButton("ログクリア")
        clear_btn.clicked.connect(self.clear_log)
        button_layout.addWidget(clear_btn)
        
        save_btn = QPushButton("ログ保存")
        save_btn.clicked.connect(self.save_log)
        button_layout.addWidget(save_btn)
        
        button_layout.addStretch()
        
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(self.close)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
    
    def setup_log_capture(self):
        """標準出力のキャプチャを設定"""
        self.log_stream = LogStream()
        self.log_stream.text_written.connect(self.append_log)
        
        # 元の標準出力を保存
        self.original_stdout = sys.stdout
        
        # 標準出力をリダイレクト
        sys.stdout = self.log_stream
    
    def append_log(self, text):
        """ログを追加"""
        self.log_text.moveCursor(QTextCursor.End)
        self.log_text.insertPlainText(text)
        self.log_text.moveCursor(QTextCursor.End)
    
    def clear_log(self):
        """ログをクリア"""
        self.log_text.clear()
    
    def save_log(self):
        """ログをファイルに保存"""
        from PySide6.QtWidgets import QFileDialog
        from datetime import datetime
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_filename = f'process_log_{timestamp}.txt'
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "ログ保存",
            default_filename,
            "Text Files (*.txt)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.toPlainText())
                
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.information(self, "完了", f"ログを保存しました:\n{file_path}")
            
            except Exception as e:
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.warning(self, "エラー", f"ログ保存エラー:\n{str(e)}")
    
    def closeEvent(self, event):
        """ダイアログを閉じる時の処理"""
        # 標準出力を元に戻す
        sys.stdout = self.original_stdout
        event.accept()