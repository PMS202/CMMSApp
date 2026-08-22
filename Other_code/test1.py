import requests

url = "https://yageo-my.sharepoint.com/:x:/p/trang_thithuy_nguyen/IQByW3rBRH00TYXuvq6kpzd8AeBaDJRA2BgSH1xjvv55CF4?e=0O5Bu2&download=1"

output_file = "PE3_DUNG_CU_STOCK_THANG_082026.xlsx"

response = requests.get(url, allow_redirects=True)

if response.status_code == 200:
    with open(output_file, "wb") as f:
        f.write(response.content)

    print(f"Download thành công: {output_file}")
    print(f"Size: {len(response.content) / 1024:.1f} KB")
else:
    print("Download thất bại")
    print("Status code:", response.status_code)