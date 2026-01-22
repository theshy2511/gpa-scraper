import re
from bs4 import BeautifulSoup

def extract_gpa(soup):
    try:
        # Lấy toàn bộ văn bản từ trang xem điểm
        text = soup.get_text(separator=' ', strip=True)
        
        # Tìm chính xác con số (x.xx) đứng sau cụm từ bạn cung cấp
        # Cấu trúc: Trung bình chung tích luỹ: 3.22
        pattern = r"Trung bình chung tích luỹ[:\s]*(\d+\.\d+)"
        match = re.search(pattern, text)
        
        if match:
            return match.group(1)
        
        # Nếu không tìm thấy định dạng x.xx, thử tìm số nguyên (ví dụ: 3)
        pattern_int = r"Trung bình chung tích luỹ[:\s]*(\d+)"
        match_int = re.search(pattern_int, text)
        if match_int:
            return match_int.group(1)
            
        return "0.0"
    except Exception:
        return "N/A"

def check_semester_exists(soup, semester_name):
    """Xác định trạng thái còn học/nghỉ học cho sinh viên HUIT"""
    return semester_name in soup.get_text()
