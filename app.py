import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, 
                           QLabel, QFileDialog, QMessageBox, QLineEdit, QGroupBox, QHBoxLayout)
from PyQt6.QtCore import Qt
import pandas as pd
import os

class FileRenamerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("發票檔案重新命名工具")
        self.setMinimumSize(600, 400)
        
        # 初始化UI
        self.init_ui()
        
    def init_ui(self):
        # 主窗口佈局
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # 1. Excel 文件選擇
        excel_group = QGroupBox("1. 選擇 Excel 文件")
        excel_layout = QHBoxLayout()
        
        self.excel_path = QLineEdit()
        self.excel_path.setPlaceholderText("請選擇 Excel 文件...")
        excel_layout.addWidget(self.excel_path)
        
        btn_browse_excel = QPushButton("瀏覽...")
        btn_browse_excel.clicked.connect(self.browse_excel)
        excel_layout.addWidget(btn_browse_excel)
        
        excel_group.setLayout(excel_layout)
        
        # 2. PDF 目錄選擇
        pdf_group = QGroupBox("2. 選擇 PDF 目錄")
        pdf_layout = QHBoxLayout()
        
        self.pdf_dir = QLineEdit()
        self.pdf_dir.setPlaceholderText("請選擇包含 PDF 的目錄...")
        pdf_layout.addWidget(self.pdf_dir)
        
        btn_browse_pdf = QPushButton("瀏覽...")
        btn_browse_pdf.clicked.connect(self.browse_pdf_dir)
        pdf_layout.addWidget(btn_browse_pdf)
        
        pdf_group.setLayout(pdf_layout)
        
        # 3. 執行按鈕
        self.btn_rename = QPushButton("開始重新命名")
        self.btn_rename.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 16px;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        self.btn_rename.clicked.connect(self.process_files)
        self.btn_rename.setEnabled(False)
        
        # 4. 日誌區域
        self.log_label = QLabel("準備就緒...")
        self.log_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # 將所有元件添加到主佈局
        layout.addWidget(excel_group)
        layout.addWidget(pdf_group)
        layout.addWidget(self.btn_rename)
        layout.addWidget(self.log_label)
        layout.addStretch()
        
    def browse_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "選擇 Excel 文件", 
            "", 
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*)"
        )
        if file_path:
            self.excel_path.setText(file_path)
            self.check_ready()
    
    def browse_pdf_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "選擇 PDF 目錄")
        if dir_path:
            self.pdf_dir.setText(dir_path)
            self.check_ready()
    
    def check_ready(self):
        """檢查是否已選擇所有必要路徑"""
        excel_ready = os.path.isfile(self.excel_path.text())
        pdf_ready = os.path.isdir(self.pdf_dir.text())
        self.btn_rename.setEnabled(excel_ready and pdf_ready)
    
    def log(self, message):
        """更新日誌"""
        self.log_label.setText(message)
        QApplication.processEvents()  # 立即更新UI
    
    def process_files(self):
        """處理文件重新命名"""
        try:
            self.btn_rename.setEnabled(False)
            self.log("正在讀取 Excel 文件...")
            
            # 讀取 Excel 文件
            df = pd.read_excel(self.excel_path.text(), sheet_name=0, header=0)
            
            # 檢查必要的列
            if '新檔名' not in df.columns or '掃描檔名' not in df.columns:
                QMessageBox.critical(self, "錯誤", "Excel 文件中缺少必要的列：'新檔名' 或 '掃描檔名'")
                return
            
            # 過濾有效行
            valid_rows = df.dropna(subset=['掃描檔名'])
            if len(valid_rows) == 0:
                QMessageBox.warning(self, "警告", "沒有找到有效的 '掃描檔名' 數據")
                return
            
            success_count = 0
            error_count = 0
            
            self.log(f"準備處理 {len(valid_rows)} 個檔案...")
            
            for _, row in valid_rows.iterrows():
                try:
                    original_file = str(row['掃描檔名']).strip()
                    new_name = f"{row['新檔名']}.pdf"
                    
                    original_path = os.path.join(self.pdf_dir.text(), original_file)
                    new_path = os.path.join(self.pdf_dir.text(), new_name)
                    
                    if not os.path.exists(original_path):
                        self.log(f"找不到原始檔案: {original_file}")
                        error_count += 1
                        continue
                        
                    if os.path.exists(new_path):
                        self.log(f"檔案已存在，跳過: {new_name}")
                        error_count += 1
                        continue
                        
                    os.rename(original_path, new_path)
                    success_count += 1
                    self.log(f"已重新命名: {original_file} -> {new_name}")
                    
                except Exception as e:
                    self.log(f"處理 {original_file} 時出錯: {str(e)}")
                    error_count += 1
            
            # 顯示結果
            result_msg = f"處理完成！\n成功: {success_count} 個\n失敗: {error_count} 個"
            QMessageBox.information(self, "完成", result_msg)
            self.log("準備就緒...")
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"處理過程中發生錯誤:\n{str(e)}")
            self.log("發生錯誤，請重試...")
        finally:
            self.btn_rename.setEnabled(True)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = FileRenamerApp()
    window.show()
    sys.exit(app.exec())