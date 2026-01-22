import os
import time
import openpyxl
import random
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from helpers import extract_gpa, check_semester_exists

def main():
    # 1. Cấu hình trình duyệt
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.set_page_load_timeout(30) # Chống đơ khi web lag
    
    excel_path = "Data_14DH.xlsx"
    BATCH_LIMIT = 100  # Giới hạn 100 người mỗi lần chạy
    processed_count = 0
    
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb["14DHTH"]
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            # Nếu đã quét đủ 100 người mới trong đợt này thì dừng
            if processed_count >= BATCH_LIMIT:
                print(f"🎯 Đã hoàn thành hạn mức {BATCH_LIMIT} người của đợt này.")
                break
                
            status_hien_tai = row[7].value # Cột H
            url_xem_diem = row[4].value    # Cột E

            # logic: Trống trạng thái mới quét
            if status_hien_tai and str(status_hien_tai).strip():
                continue

            if not url_xem_diem:
                continue

            try:
                print(f"🔍 [{processed_count + 1}/{BATCH_LIMIT}] Đang quét dòng {row_idx}...")
                driver.get(str(url_xem_diem).strip())
                
                # Nghỉ ngẫu nhiên để tránh bị chặn IP
                time.sleep(random.uniform(2, 4)) 
                
                soup = BeautifulSoup(driver.page_source, "html.parser")
                gpa = extract_gpa(soup)
                
                # Kiểm tra học kỳ đúng ý bạn
                is_active = check_semester_exists(soup, "HK2 (2025 - 2026)")
                status_moi = "còn học" if is_active else "nghỉ học"
                
                # Ghi dữ liệu
                ws.cell(row=row_idx, column=7, value=gpa)
                ws.cell(row=row_idx, column=8, value=status_moi)
                
                processed_count += 1
                
                # Lưu sau mỗi dòng để đảm bảo an toàn dữ liệu
                wb.save(excel_path)
                print(f"✅ Xong: {gpa} | {status_moi}")

            except Exception as e:
                print(f"⚠️ Lỗi dòng {row_idx}: {e}")
                continue

        print(f"🏁 Đợt chạy kết thúc. Tổng cộng đã quét thêm {processed_count} sinh viên.")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
