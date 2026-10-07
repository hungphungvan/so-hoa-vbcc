import os

from dotenv import load_dotenv
from openai import OpenAI

# Tải biến môi trường
load_dotenv()

API_KEY = os.getenv("CKEY_API_KEY")
BASE_URL = os.getenv("CKEY_BASE_URL", "https://api.xah.io/v1")
MODEL_NAME = os.getenv("GEMINI_MODEL", "dungcsnd113/gemini-pro-3.1")

if not API_KEY:
    raise ValueError("Chưa thiết lập CKEY_API_KEY trong file .env")

# Khởi tạo client dùng chung cho toàn dự án
api_client = OpenAI(
    api_key=API_KEY,
    base_url=BASE_URL,
    timeout=180.0 # cho phép chờ tối đa 3 phút mỗi request
)
