import os
import time
import openpyxl
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Đảm bảo bạn đã có helpers.py và config.py trên repo
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
            # 1. KIỂM TRA CỘT G (GPA - index 6)
            gpa_hien_tai = row[6].value 
            
            # Nếu đã có điểm (không rỗng) thì bỏ qua
            if gpa_hien_tai is not None and str(gpa_hien_tai).strip() != "":
                print(f"⏩ Dòng {row_idx}: Đã có điểm ({gpa_hien_tai}), bỏ qua.")
                continue
                
            # 2. LẤY URL TỪ CỘT E (index 4)
            url_xem_diem = row[4].value
            if not url_xem_diem:
                continue
            
            print(f"🔍 Đang xử lý dòng {row_idx}...")
            driver.get(str(url_xem_diem).strip())
            time.sleep(2) 
            
            soup = BeautifulSoup(driver.page_source, "html.parser")
            gpa = extract_gpa(soup)
            status = "còn học" if check_semester_exists(soup, "HK2 (2025 - 2026)") else "nghỉ học"
            
            # 3. GHI DỮ LIỆU MỚI
            ws.cell(row=row_idx, column=7, value=gpa)     # Cột G
            ws.cell(row=row_idx, column=8, value=status)  # Cột H
            
            # Lưu sau mỗi dòng để đảm bảo không mất dữ liệu nếu rớt mạng
            wb.save(excel_path)
            print(f"✅ Đã điền: GPA={gpa}, {status}")

        print("🎉 HOÀN TẤT!")
    except Exception as e:
        print(f"❌ Lỗi: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
