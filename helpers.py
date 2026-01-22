import re
from bs4 import BeautifulSoup

def extract_gpa(soup):
    """Lấy điểm trung bình từ HTML"""
    try:
        text = soup.get_text()
        # Tìm con số đứng sau cụm từ liên quan đến tích lũy
        match = re.search(r"tích lũy[:\s]+(\d+\.\d+)", text)
        return match.group(1) if match else "0.0"
    except:
        return "N/A"

def check_semester_exists(soup, semester_name):
    """Kiểm tra sự xuất hiện của học kỳ cụ thể"""
    return semester_name in soup.get_text()
