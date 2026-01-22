import re
from bs4 import BeautifulSoup

def extract_gpa(soup):
    try:
        # Lấy văn bản từ trang web
        text = soup.get_text(separator=' ', strip=True)
        
        # Tìm khu vực có chữ "Trung bình chung tích luỹ"
        if "Trung bình chung tích luỹ" in text:
            start_idx = text.find("Trung bình chung tích luỹ")
            # Cắt lấy đoạn văn bản ngắn ngay sau đó để tìm điểm
            sub_text = text[start_idx : start_idx + 150]
            
            # Tìm tất cả các số thập phân (ví dụ: 7.00, 3.22)
            numbers = re.findall(r"\d+\.\d+", sub_text)
            
            # Chuyển sang dạng số thực (float)
            float_numbers = [float(n) for n in numbers]
            
            # LỌC: Chỉ lấy số nào nhỏ hơn hoặc bằng 4.0 (Hệ 4)
            gpa_4_list = [n for n in float_numbers if n <= 4.0]
            
            if gpa_4_list:
                # Trả về số đầu tiên tìm thấy và ép định dạng 2 chữ số thập phân
                return "{:.2f}".format(gpa_4_list[0])
                
        return "0.00"
    except Exception:
        return "N/A"

def check_semester_exists(soup, semester_name):
    """Giữ nguyên logic xét học kỳ"""
    return semester_name in soup.get_text()
