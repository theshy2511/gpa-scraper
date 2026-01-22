import re
from bs4 import BeautifulSoup

def extract_gpa(soup):
    """
    Trích xuất điểm trung bình tích lũy từ mã nguồn HTML của trang điểm HUIT.
    """
    try:
        # Lấy toàn bộ văn bản từ trang web
        text = soup.get_text()
        
        # Tìm con số thập phân (ví dụ 3.22) đứng sau chữ 'tích lũy'
        # re.IGNORECASE giúp tìm kiếm không phân biệt chữ hoa chữ thường
        match = re.search(r"tích lũy[:\s]+(\d+\.\d+)", text, re.IGNORECASE)
        
        if match:
            return match.group(1)
        return "0.0" # Trả về 0.0 nếu không tìm thấy điểm
    except Exception:
        return "N/A"

def check_semester_exists(soup, semester_name):
    """
    Kiểm tra xem học kỳ cụ thể có tồn tại trong dữ liệu của sinh viên không.
    Dùng để xác định sinh viên năm 3 còn học hay đã nghỉ.
    """
    try:
        # Lấy toàn bộ nội dung văn bản trên trang
        content = soup.get_text()
        
        # Kiểm tra sự xuất hiện của học kỳ (ví dụ: HK2 (2025 - 2026))
        if semester_name in content:
            return True
        return False
    except Exception:
        return False
