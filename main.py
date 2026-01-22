import os
import time
import openpyxl
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Đảm bảo các file này đã có trên GitHub
from helpers import extract_gpa, check_semester_exists 
from config import * 
def main():
    options = Options()
    options.add_argument("--headless") 
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    excel_path = "Data_14DH.xlsx"
    
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb["14DHTH"]
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            url_xem_diem = row[4].value # Cột E
            if not url_xem_diem: continue
            
            driver.get(str(url_xem_diem).strip())
            time.sleep(2) 
            
            soup = BeautifulSoup(driver.page_source, "html.parser")
            gpa = extract_gpa(soup)
            status = "còn học" if check_semester_exists(soup, "HK2 (2025 - 2026)") else "nghỉ học"
            
            ws.cell(row=row_idx, column=7, value=gpa)
            ws.cell(row=row_idx, column=8, value=status)
            
            # XÓA dòng này nếu muốn chạy hết cả lớp, hiện tại chỉ test 2 người
            if row_idx > 3: break 

        wb.save(excel_path)
        print("💾 Đã lưu dữ liệu vào Excel thành công.")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
