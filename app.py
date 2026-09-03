import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, 
                             QLabel, QFileDialog, QMessageBox, QLineEdit, QGroupBox, 
                             QHBoxLayout, QTextEdit, QCheckBox)
from PyQt6.QtCore import Qt
import pandas as pd
import openpyxl  # 🌟 改用 openpyxl 進行原生刪除操作以保留公式與樣式
import os
from datetime import datetime

# 🌟 解析 Excel 欄位英文字母轉數字
def col2num(col_str):
    """將 Excel 欄位英文字母轉成以 1 為基準的數字 (例: 'A'->1, 'Z'->26, 'AA'->27)"""
    num = 0
    for c in col_str.upper().strip():
        if 'A' <= c <= 'Z':
            num = num * 26 + (ord(c) - ord('A') + 1)
    return num

def parse_keep_column_indices(range_string, total_cols):
    """
    輸入保留範圍 (如 'A-AF, AY')，回傳「需要被刪除」的欄位索引清單 (1-indexed)，從大到小排序。
    """
    keep_indices = set()
    parts = [p.strip() for p in range_string.split(',') if p.strip()]
    
    for part in parts:
        if '-' in part:
            sub_parts = part.split('-')
            if len(sub_parts) == 2:
                start_col = col2num(sub_parts[0])
                end_col = col2num(sub_parts[1])
                if start_col > 0 and end_col >= start_col:
                    for c in range(start_col, end_col + 1):
                        keep_indices.add(c)
        else:
            col = col2num(part)
            if col > 0:
                keep_indices.add(col)
                
    # 算出「要刪除」的欄位 (1 到 total_cols 中，不在 keep_indices 裡的欄位)
    delete_indices = [c for c in range(1, total_cols + 1) if c not in keep_indices]
    
    # 🌟 關鍵：刪除欄位時必須從「右至左 (大到小)」刪除，否則左邊刪掉後右邊欄位編號會移位
    delete_indices.sort(reverse=True)
    return delete_indices


class FileRenamerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("發票檔案重新命名與 Excel 裁切工具 (保留公式版)")
        self.setMinimumSize(700, 800)
        
        self.init_ui()
        
    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # 1. 選擇要重命名的目標檔案目錄
        dir_group = QGroupBox("1. 選擇目標檔案目錄")
        dir_layout = QHBoxLayout()
        
        self.pdf_dir = QLineEdit()
        self.pdf_dir.setPlaceholderText("請選擇包含檔案的目錄...")
        dir_layout.addWidget(self.pdf_dir)
        
        btn_browse_dir = QPushButton("瀏覽...")
        btn_browse_dir.clicked.connect(self.browse_pdf_dir)
        dir_layout.addWidget(btn_browse_dir)
        
        dir_group.setLayout(dir_layout)

        # 2. 步驟二：自動掃描與自動帶入規則設定
        scan_group = QGroupBox("2. 掃描檔名與規則預填設定")
        scan_v_layout = QVBoxLayout()
        
        rule1_h_layout = QHBoxLayout()
        self.chk_enable_rule = QCheckBox("啟用關鍵字自動替換帶入新檔名 (Optional)")
        self.chk_enable_rule.toggled.connect(self.toggle_rule_inputs)
        rule1_h_layout.addWidget(self.chk_enable_rule)
        
        rule_inputs_layout = QHBoxLayout()
        self.txt_keyword = QLineEdit()
        self.txt_keyword.setPlaceholderText("若檔名包含此字 (例: INV)")
        self.txt_keyword.setEnabled(False)
        
        self.txt_replace = QLineEdit()
        self.txt_replace.setPlaceholderText("自動替換/帶入為 (例: 發票)")
        self.txt_replace.setEnabled(False)
        
        rule_inputs_layout.addWidget(QLabel("關鍵字:"))
        rule_inputs_layout.addWidget(self.txt_keyword)
        rule_inputs_layout.addWidget(QLabel("帶入字:"))
        rule_inputs_layout.addWidget(self.txt_replace)

        today_date_str = datetime.now().strftime("%Y%m%d")
        self.chk_append_date = QCheckBox(f"新檔名後綴自動附加當天日期 (_{today_date_str}) (Optional)")

        self.btn_export_csv = QPushButton("🔍 掃描檔案並匯出 CSV 對照表")
        self.btn_export_csv.setStyleSheet("padding: 8px; font-weight: bold; background-color: #2196F3; color: white;")
        self.btn_export_csv.clicked.connect(self.scan_and_export_csv)
        self.btn_export_csv.setEnabled(False)
        
        scan_v_layout.addLayout(rule1_h_layout)
        scan_v_layout.addLayout(rule_inputs_layout)
        scan_v_layout.addWidget(self.chk_append_date)
        scan_v_layout.addWidget(self.btn_export_csv)
        scan_group.setLayout(scan_v_layout)
        
        # 3. 步驟三：選擇填寫好的 CSV / Excel 文件
        excel_group = QGroupBox("3. 選擇已填寫新檔名的 CSV / Excel 文件")
        excel_layout = QHBoxLayout()
        
        self.excel_path = QLineEdit()
        self.excel_path.setPlaceholderText("請選擇對照表 (CSV 或 Excel)...")
        excel_layout.addWidget(self.excel_path)
        
        btn_browse_excel = QPushButton("瀏覽...")
        btn_browse_excel.clicked.connect(self.browse_excel)
        excel_layout.addWidget(btn_browse_excel)
        
        excel_group.setLayout(excel_layout)

        # 4. 步驟四：裁切 Excel 欄位設定 (Optional)
        crop_group = QGroupBox("4. 裁切 Excel 欄位設定 (轉完檔名後執行 - 保留公式)")
        crop_v_layout = QVBoxLayout()
        
        crop_h_layout = QHBoxLayout()
        self.chk_enable_crop = QCheckBox("啟用轉檔完成後自動裁切 .xlsx 欄位 (保留原檔公式與樣式)")
        self.chk_enable_crop.toggled.connect(self.toggle_crop_inputs)
        crop_h_layout.addWidget(self.chk_enable_crop)
        
        crop_inputs_layout = QHBoxLayout()
        self.txt_crop_ranges = QLineEdit()
        self.txt_crop_ranges.setPlaceholderText("請輸入保留欄位範圍 (例: A-AF, AY)")
        self.txt_crop_ranges.setEnabled(False)
        
        crop_inputs_layout.addWidget(QLabel("保留欄位:"))
        crop_inputs_layout.addWidget(self.txt_crop_ranges)
        
        crop_v_layout.addLayout(crop_h_layout)
        crop_v_layout.addLayout(crop_inputs_layout)
        crop_group.setLayout(crop_v_layout)

        # 5. 日誌顯示區域
        log_group = QGroupBox("操作紀錄與執行日誌")
        log_layout = QVBoxLayout()
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        log_layout.addWidget(self.log_box)
        log_group.setLayout(log_layout)
        
        # 6. 執行按鈕
        self.btn_rename = QPushButton("開始執行 (重新命名 / 裁切)")
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
        
        # 組合介面元件
        layout.addWidget(dir_group)
        layout.addWidget(scan_group)
        layout.addWidget(excel_group)
        layout.addWidget(crop_group)
        layout.addWidget(log_group)
        layout.addWidget(self.btn_rename)
        
    def toggle_rule_inputs(self, checked):
        self.txt_keyword.setEnabled(checked)
        self.txt_replace.setEnabled(checked)

    def toggle_crop_inputs(self, checked):
        self.txt_crop_ranges.setEnabled(checked)

    def browse_pdf_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "選擇檔案目錄")
        if dir_path:
            self.pdf_dir.setText(dir_path)
            self.check_ready()

    def browse_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "選擇對照表", 
            "", 
            "表格文件 (*.csv *.xlsx *.xls);;所有文件 (*)"
        )
        if file_path:
            self.excel_path.setText(file_path)
            self.check_ready()
            
    def check_ready(self):
        excel_ready = os.path.isfile(self.excel_path.text())
        pdf_ready = os.path.isdir(self.pdf_dir.text())
        
        self.btn_export_csv.setEnabled(pdf_ready)
        self.btn_rename.setEnabled(excel_ready and pdf_ready)
    
    def log(self, message):
        self.log_box.append(message)
        QApplication.processEvents()

    def scan_and_export_csv(self):
        folder_path = self.pdf_dir.text()
        if not os.path.isdir(folder_path):
            QMessageBox.warning(self, "警告", "請先選擇有效的檔案目錄！")
            return
            
        try:
            files = [
                f for f in os.listdir(folder_path) 
                if os.path.isfile(os.path.join(folder_path, f)) and not f.endswith(".csv")
            ]
            
            if not files:
                QMessageBox.warning(self, "提示", "選定的資料夾內沒有任何檔案！")
                return

            today_str = datetime.now().strftime("%Y%m%d")
            default_csv_name = os.path.join(folder_path, f"檔案重命名對照表_{today_str}.csv")
            
            save_path, _ = QFileDialog.getSaveFileName(
                self, 
                "儲存檔名對照表 CSV", 
                default_csv_name, 
                "CSV 文件 (*.csv)"
            )
            
            if not save_path:
                return

            use_rule = self.chk_enable_rule.isChecked()
            keyword = self.txt_keyword.text().strip()
            replace_text = self.txt_replace.text().strip()
            append_date = self.chk_append_date.isChecked()

            data = []
            matched_count = 0

            for file_name in files:
                name_without_ext, ext = os.path.splitext(file_name)
                suggested_new_name = ""

                if use_rule and keyword and (keyword in name_without_ext):
                    suggested_new_name = name_without_ext.replace(keyword, replace_text)
                    matched_count += 1

                if append_date:
                    base_for_date = suggested_new_name if suggested_new_name else name_without_ext
                    suggested_new_name = f"{base_for_date}_{today_str}"

                data.append({
                    '掃描檔名': file_name,
                    '新檔名': suggested_new_name
                })
            
            df = pd.DataFrame(data)
            df.to_csv(save_path, index=False, encoding='utf-8-sig')
            
            self.log_box.clear()
            self.log(f"✅ 成功掃描 {len(files)} 個檔案！")
            if use_rule and keyword:
                self.log(f"🎯 符合關鍵字 '{keyword}' 且自動替換的檔案共 {matched_count} 個。")
            if append_date:
                self.log(f"📅 已為建議檔名自動附加日期後綴 (_{today_str})。")
            self.log(f"📄 已匯出 CSV 檔至: {save_path}")
            
            self.excel_path.setText(save_path)
            self.check_ready()

            reply = QMessageBox.question(
                self, '完成', 
                f"已成功匯出對照表！\n共找到 {len(files)} 個檔案。\n\n是否立即開啟此 CSV 檔案進行編輯？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
                QMessageBox.StandardButton.Yes
            )

            if reply == QMessageBox.StandardButton.Yes:
                os.startfile(save_path)

        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"匯出 CSV 時發生錯誤:\n{str(e)}")

    # 🌟 [全新改寫] 使用 openpyxl 直接對原檔刪除欄位（保留公式與格式）
    def crop_excel_files(self, folder_path, range_str):
        self.log("\n✂️ 開始執行 Excel 欄位裁切 (保留公式模式)...")
        
        # 搜尋資料夾內的所有 .xlsx 檔案 (排除對照表 CSV 和暫存檔)
        excel_files = [
            f for f in os.listdir(folder_path) 
            if f.lower().endswith('.xlsx') and not f.startswith('~$')
        ]
        
        if not excel_files:
            self.log("⚠️ 沒有找到任何 .xlsx 檔案可進行裁切。")
            return

        cropped_count = 0
        for e_file in excel_files:
            file_path = os.path.join(folder_path, e_file)
            try:
                # 載入 Excel 活頁簿 (保持公式原樣)
                wb = openpyxl.load_workbook(file_path)
                ws = wb.active  # 使用主要工作表
                
                total_cols = ws.max_column
                
                # 計算出需要刪除的欄位 (已預先由大到小排序)
                delete_indices = parse_keep_column_indices(range_str, total_cols)
                
                if not delete_indices:
                    self.log(f"⚠️ {e_file} 沒有需要刪除的欄位，跳過。")
                    continue

                # 🌟 從右向左逐一刪除不保留的欄位，確保公式參照不偏移
                for col_idx in delete_indices:
                    ws.delete_cols(col_idx)
                
                wb.save(file_path)
                wb.close()
                
                cropped_count += 1
                self.log(f"✂️ 已完成欄位裁切: {e_file} (已安全刪除未指定的欄位，保留原有公式)")

            except Exception as ex:
                self.log(f"❌ 裁切 {e_file} 時發生錯誤: {str(ex)}")
                
        self.log(f"🎉 欄位裁切作業完成！共處理 {cropped_count} 個 Excel 檔案。")

    def process_files(self):
        try:
            self.btn_rename.setEnabled(False)
            self.log("\n🚀 開始執行重新命名流程...")
            
            file_path = self.excel_path.text()
            folder_path = self.pdf_dir.text()
            
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8-sig')
            else:
                df = pd.read_excel(file_path, sheet_name=0, header=0)
            
            if '新檔名' not in df.columns or '掃描檔名' not in df.columns:
                QMessageBox.critical(self, "錯誤", "對照表中缺少必要的欄位：'新檔名' 或 '掃描檔名'")
                return
            
            valid_rows = df.dropna(subset=['新檔名'])
            valid_rows = valid_rows[valid_rows['新檔名'].astype(str).str.strip() != '']
            
            if len(valid_rows) == 0:
                QMessageBox.warning(self, "警告", "沒有在『新檔名』欄位中找到任何已填寫的資料！")
                return
            
            success_count = 0
            error_count = 0
            
            for _, row in valid_rows.iterrows():
                try:
                    original_file = str(row['掃描檔名']).strip()
                    new_base_name = str(row['新檔名']).strip()
                    
                    original_path = os.path.join(folder_path, original_file)
                    
                    if not os.path.exists(original_path):
                        self.log(f"❌ 找不到原始檔案: {original_file}")
                        error_count += 1
                        continue

                    _, ext = os.path.splitext(original_file)
                    
                    if new_base_name.lower().endswith(ext.lower()):
                        new_name = new_base_name
                    else:
                        new_name = f"{new_base_name}{ext}"

                    new_path = os.path.join(folder_path, new_name)
                    
                    if original_path == new_path:
                        continue
                        
                    if os.path.exists(new_path):
                        self.log(f"⚠️ 目標檔名已存在，跳過: {new_name}")
                        error_count += 1
                        continue
                        
                    os.rename(original_path, new_path)
                    success_count += 1
                    self.log(f"✅ 已重新命名: {original_file}  ➡️  {new_name}")
                    
                except Exception as e:
                    self.log(f"❌ 處理 {original_file} 時出錯: {str(e)}")
                    error_count += 1
            
            self.log(f"\n重新命名完成！成功: {success_count} 個，失敗/跳過: {error_count} 個。")

            # 🌟 自動判斷是否執行 Excel 欄位裁切
            if self.chk_enable_crop.isChecked():
                crop_ranges = self.txt_crop_ranges.text().strip()
                if crop_ranges:
                    self.crop_excel_files(folder_path, crop_ranges)
                else:
                    QMessageBox.warning(self, "提示", "已勾選欄位裁切，但未輸入欄位範圍！")

            result_msg = f"處理完成！\n成功改名: {success_count} 個\n失敗/跳過: {error_count} 個"
            QMessageBox.information(self, "完成", result_msg)
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"處理過程中發生錯誤:\n{str(e)}")
        finally:
            self.btn_rename.setEnabled(True)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = FileRenamerApp()
    window.show()
    sys.exit(app.exec())