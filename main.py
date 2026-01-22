import os
import time
import logging
import openpyxl
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Giả sử bạn để các hàm này trong file helpers.py
from helpers import extract_gpa, check_semester_exists 
# Và các biến trong config.py
from config import * def main():
    options = Options()
    options.add_argument("--headless") # Chạy ngầm không mở cửa sổ
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    
    # Tự động tải Driver phù hợp với server GitHub
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    # Đường dẫn file trên GitHub (dùng đường dẫn tương đối)
    excel_path = "Data_14DH.xlsx"
    
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb["14DHTH"]
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            url_xem_diem = row[4].value # Cột E
            if not url_xem_diem: continue
            
            driver.get(str(url_xem_diem).strip())
            time.sleep(2) # Đợi load trang
            
            soup = BeautifulSoup(driver.page_source, "html.parser")
            gpa = extract_gpa(soup)
            status = "còn học" if check_semester_exists(soup, "HK2 (2025 - 2026)") else "nghỉ học"
            
            # Ghi vào cột G và H
            ws.cell(row=row_idx, column=7, value=gpa)
            ws.cell(row=row_idx, column=8, value=status)
            
            # Chỉ chạy thử 2 người nếu muốn test nhanh
            if row_idx > 3: break 

        wb.save(excel_path)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
