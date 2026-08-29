# import requests

# url = "https://yageo-my.sharepoint.com/:x:/p/trang_thithuy_nguyen/IQByW3rBRH00TYXuvq6kpzd8AeBaDJRA2BgSH1xjvv55CF4?e=0O5Bu2&download=1"

# output_file = "PE3_DUNG_CU_STOCK_THANG_082026.xlsx"

# response = requests.get(url, allow_redirects=True)

# if response.status_code == 200:
#     with open(output_file, "wb") as f:
#         f.write(response.content)

#     print(f"Download thành công: {output_file}")
#     print(f"Size: {len(response.content) / 1024:.1f} KB")
# else:
#     print("Download thất bại")
#     print("Status code:", response.status_code)
import os
import requests
from pathlib import Path
from dotenv import load_dotenv
load_dotenv(Path(__file__).resolve().parents[1] / ".env")
API_UPLOAD_URL = os.getenv("API_UPLOAD_ACTION_FILE")  # thêm biến này vào .env

def upload_file_to_server(local_path):
    with open(local_path, "rb") as f:
        resp = requests.post(API_UPLOAD_URL, files={"file": (os.path.basename(local_path), f)}, timeout=60)
    resp.raise_for_status()
    return resp.json()["server_path"]

# # trong action_before_closed():
# for key, value in enumerate(action_content):
#     if value['type'] == 'link':
#         local_path = QtCore.QUrl(value['href']).toLocalFile()
#         try:
#             server_path = upload_file_to_server(local_path)
#         except Exception as e:
#             server_path = value['href']  # fallback nếu upload lỗi
#         link_list.append({
#             "file_name": value['text'],
#             "file_path": server_path
#         }_

upload_file_to_server(r"C:\Users\2173452100291\Downloads\Automated_Safety_Stock_Management.pptx")