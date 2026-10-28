import os
import google.generativeai as genai
from dotenv import load_dotenv

# Tải biến môi trường từ file .env
load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY")

if API_KEY:
    genai.configure(api_key=API_KEY)
else:
    print("⚠️ Cảnh báo: Chưa tìm thấy GEMINI_API_KEY trong file .env!")

# Khởi tạo model
model = genai.GenerativeModel("models/gemini-3.5-flash")