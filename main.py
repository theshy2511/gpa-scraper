import os
import time
import openpyxl
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from helpers import extract_gpa, check_semester_exists

def main():
    # Cấu hình trình duyệt chạy ngầm (Headless)
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    # Giới hạn thời gian chờ load trang là 20 giây để tránh treo
    driver.set_page_load_timeout(20)
    
    excel_path = "Data_14DH.xlsx"
    
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb["14DHTH"]
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            # Cột H là index 7, Cột E là index 4
            status_hien_tai = row[7].value 
            url_xem_diem = row[4].value

            # 1. Nếu đã có trạng thái thì bỏ qua
            if status_hien_tai and str(status_hien_tai).strip():
                print(f"⏩ Dòng {row_idx}: Đã có dữ liệu, bỏ qua.")
                continue

            if not url_xem_diem:
                continue

            # 2. Truy cập web với xử lý lỗi timeout
            try:
                print(f"🔍 Đang quét dòng {row_idx}...")
                driver.get(str(url_xem_diem).strip())
                time.sleep(2) # Chờ render nhẹ
                
                soup = BeautifulSoup(driver.page_source, "html.parser")
                
                # 3. Lấy GPA và xét học kỳ
                gpa = extract_gpa(soup)
                con_hoc = check_semester_exists(soup, "HK2 (2025 - 2026)")
                status_moi = "còn học" if con_hoc else "nghỉ học"
                
                # 4. Ghi vào file (Cột G và H)
                ws.cell(row=row_idx, column=7, value=gpa)
                ws.cell(row=row_idx, column=8, value=status_moi)
                
                # Lưu ngay lập tức sau mỗi dòng
                wb.save(excel_path)
                print(f"✅ Xong dòng {row_idx}: {gpa} | {status_moi}")

            except Exception as e:
                print(f"⚠️ Lỗi tại dòng {row_idx} (Có thể do web lag): {e}")
                continue # Lỗi người này thì làm tiếp người sau

        print("🎉 Đã hoàn thành toàn bộ danh sách!")

    except Exception as e:
        print(f"❌ Lỗi hệ thống: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
