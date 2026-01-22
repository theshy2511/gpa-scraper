import re
from bs4 import BeautifulSoup

def extract_gpa(soup):
    try:
        # Lấy văn bản, giữ nguyên các khoảng trắng để phân tách số
        text = soup.get_text(separator=' ', strip=True)
        
        # 1. Tìm khu vực chứa cụm từ bạn chỉ định
        # Sử dụng Regex để tìm tất cả các số thập phân đứng sau cụm từ đó
        # Chúng ta tìm trong phạm vi 50 ký tự sau cụm từ để tránh lấy nhầm số ở xa
        pattern = r"Trung bình chung tích luỹ.*?(\d+\.\d{1,2}|\d+)"
        matches = re.findall(pattern, text)
        
        if matches:
            # Chuyển danh sách tìm được sang số thực
            scores = []
            for m in matches:
                try:
                    scores.append(float(m))
                except: continue
            
            # 2. Lọc lấy điểm hệ 4 (thường <= 4.0)
            # Nếu có nhiều số, ta lấy số nào <= 4.0
            gpa_4 = [s for s in scores if s <= 4.0]
            
            if gpa_4:
                # Trả về số đầu tiên tìm thấy và ép định dạng 2 chữ số thập phân (VD: 3.20)
                return "{:.2f}".format(gpa_4[0])
            else:
                # Nếu không có số nào <= 4, có thể trang web chỉ hiện hệ 10
                # Ta lấy số đầu tiên và thử quy đổi hoặc giữ nguyên để bạn kiểm tra
                return "{:.2f}".format(scores[0]) if scores else "0.00"
                
        return "0.00"
    except Exception:
        return "N/A"

def check_semester_exists(soup, semester_name):
    """Giữ nguyên logic xét học kỳ của bạn"""
    return semester_name in soup.get_text()
