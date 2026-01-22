import os
from pathlib import Path
from datetime import datetime

# ==== PATHS ====
BASE_DIR = Path(__file__).parent
EXCEL_FILE = BASE_DIR / "Data_13DH.xlsx"  # File dữ liệu 13DH
CHECKPOINT_FILE = BASE_DIR / "checkpoint_13dh.json"  # File lưu tiến trình

# Output directories
OUTPUT_DIR = BASE_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"

# Create directories if not exist
OUTPUT_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)

# Output files
EXCEL_OUTPUT = OUTPUT_DIR / f"DuLieuSV_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
LOG_FILE = LOGS_DIR / "data_collection_13dh.log"

# ==== EXCEL SHEETS ====
SHEET_SINHVIEN = "13DHTH"  # Sheet cho lớp 13DH
SHEET_MONHOC = "MonHoc"  # Sheet danh sách môn 3TC (nếu có)

# ==== PRICES ====
LT_PRICE = 785000
TH_PRICE = 1000000

# ==== CAPTCHA API ====
CAPTCHA_API_URL = "https://api.capsolver.com/createTask"
CAPTCHA_API_KEY = os.getenv("CAPTCHA_API_KEY") or "CAP-1F1636B7F40F4F54E5902B2C743487D4B51CF9AED89DDB2442C908412FE42827"
CAPTCHA_MAX_RETRIES = 10  # Tăng lên 10 để đảm bảo success rate cao hơn

# ==== TEST MODE ====
TEST_MODE = True  # BẬT test mode - chỉ chạy 100 sinh viên (set False để chạy hết)
TEST_LIMIT = 100  # Chỉ chạy 100 sinh viên đầu tiên khi TEST_MODE = True

# ==== FORCE REPROCESS ====
FORCE_REPROCESS = False  # Set True để chạy lại TẤT CẢ sinh viên, False để tự động bỏ qua sinh viên đã có dữ liệu

# ==== BROWSER SETTINGS ====
HEADLESS = False  # Set True để ẩn browser
BROWSER_TIMEOUT = 30

# ==== DELAYS (tránh bị block) ====
DELAY_BETWEEN_STUDENTS = 5  # giây - tăng từ 3s lên 5s
DELAY_AFTER_CAPTCHA_FAIL = 1.5  # giây - giảm vì chỉ refresh CAPTCHA thay vì reload trang
DELAY_AFTER_SUCCESS = 1

# ==== V2 TOOL - GPA COLLECTION ====
EXCEL_FILE_V2 = BASE_DIR / "Data_13DH_V2.xlsx"  # File dữ liệu V2
CHECKPOINT_FILE_V2 = BASE_DIR / "checkpoint_gpa_v2.json"  # Checkpoint cho V2
SHEET_V2 = "13DHTH"  # Sheet cho V2 (giống file 13DH)

# Target semester (for priority checking)
TARGET_SEMESTER = "HK1 (2025 - 2026)"

# Dropout threshold
DROPOUT_THRESHOLD = 0.5  # 50% môn thiếu điểm/không đạt
