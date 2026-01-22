import re
from bs4 import BeautifulSoup

def extract_gpa(soup):
    try:
        # 1. Lấy toàn bộ text và xóa các ký tự trắng dư thừa
        full_text = " ".join(soup.get_text().split())
        
        # 2. Tìm vị trí của từ khóa (không phân biệt hoa thường, chấp nhận cả luỹ/lũy)
        # Sử dụng Regex để quét cụm từ một cách linh hoạt nhất
        keyword_pattern = r"Trung bình chung tích luy" # Tìm tương đối
        match_keyword = re.search(keyword_pattern, full_text, re.IGNORECASE)
        
        if match_keyword:
            # Lấy 100 ký tự sau từ khóa để tìm số
            start = match_keyword.start()
            sub_text = full_text[start : start + 100]
            
            # 3. Tìm tất cả các số (cả số nguyên và số thập phân, chấp nhận dấu phẩy hoặc dấu chấm)
            # Regex này bắt được: 3.22 hoặc 3,22 hoặc 3
            numbers = re.findall(r"\d+[.,]\d+|\d+", sub_text)
            
            # Chuyển đổi về float và lọc hệ 4
            valid_scores = []
            for n in numbers:
                val = float(n.replace(',', '.'))
                # Điểm hệ 4 thường nằm từ 0.0 đến 4.0
                if 0.0 < val <= 4.0:
                    valid_scores.append(val)
            
            if valid_scores:
                # Trả về số cuối cùng tìm được (thường hệ 4 nằm sau hệ 10)
                return "{:.2f}".format(valid_scores[-1])
        
        # Nếu vẫn không thấy, quét toàn bộ trang tìm số nào <= 4.0
        all_numbers = re.findall(r"\d+[.,]\d+", full_text)
        gpa_4_back = [float(n.replace(',', '.')) for n in all_numbers if 0.0 < float(n.replace(',', '.')) <= 4.0]
        
        if gpa_4_back:
            return "{:.2f}".format(gpa_4_back[-1])

        return "0.00"
    except Exception:
        return "N/A"

def check_semester_exists(soup, semester_name):
    return semester_name in soup.get_text()
