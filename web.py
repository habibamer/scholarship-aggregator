import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import time

urls = {
    "bachelor": "https://www.for9a.com/en/opportunity/category/Scholarships/bachelor",
    "master": "https://www.for9a.com/en/opportunity/category/Scholarships/master",
    "phd": "https://www.for9a.com/en/opportunity/category/Scholarships/phd"
}

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

data_list = []
max_pages = 5  

for deg, base_url in urls.items():
    print(f"Fetching level: {deg}...")
    
    for page in range(1, max_pages + 1):
        target_url = f"{base_url}?page={page}"
        print(f"  --> Page {page}...")
        
        try:
            res = requests.get(target_url, headers=headers, timeout=10)
            if res.status_code != 200:
                print(f"      Stopped page {page} with status {res.status_code}")
                break
                
            soup = BeautifulSoup(res.content, "lxml")
            cards = soup.find_all("div", class_="p-4 flex flex-col flex-grow")
            
            if not cards:
                print("      No more cards found.")
                break
                
            for card in cards:
                title_tag = card.find("a", class_="editor_page")
                if not title_tag:
                    continue
                    
                title = title_tag.get_text(strip=True)
                href = title_tag.get("href", "")
                link = href if href.startswith("http") else f"https://www.for9a.com{href}"
                
                c_tag = card.find("span", class_="bg-gray-100")
                country = c_tag.get_text(strip=True) if c_tag else "International"
                
                item = [deg, title, country, link]
                if item not in data_list:
                    data_list.append(item)
                    
            time.sleep(1) 
            
        except Exception as e:
            print(f"      Error on page {page}: {e}")
            break


wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Scholarships"

cols = ["Degree", "Title", "Country", "Link"]
ws.append(cols)


fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
font = Font(bold=True, size=11)

for i in range(1, 5):
    c = ws.cell(row=1, column=i)
    c.fill = fill
    c.font = font
    c.alignment = Alignment(horizontal="center", vertical="center")

for row in data_list:
    ws.append(row)

align = Alignment(wrap_text=True, vertical="center")
for row in ws.iter_rows(min_row=2, max_row=len(data_list) + 1, min_col=1, max_col=4):
    for c in row:
        c.alignment = align

widths = {1: 15, 2: 50, 3: 20, 4: 45}
for idx, w in widths.items():
    ws.column_dimensions[get_column_letter(idx)].width = w

wb.save("scholarships.xlsx")
print(f"Done! Collected {len(data_list)} scholarships successfully.")