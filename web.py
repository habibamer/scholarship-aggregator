import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import time

# target urls
scholarship_urls = {
    "bachelor": "https://www.for9a.com/en/opportunity/category/Scholarships/bachelor",
    "master": "https://www.for9a.com/en/opportunity/category/Scholarships/master",
    "phd": "https://www.for9a.com/en/opportunity/category/Scholarships/phd"
}

# headers
browser_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

all_scholarships = []

# scrape loop
for level, target_url in scholarship_urls.items():
    print(f"Fetching: {level}...")
    page_data = requests.get(target_url, headers=browser_headers)
    
    if page_data.status_code == 200:
        soup = BeautifulSoup(page_data.content, "lxml")
        cards = soup.find_all("div", class_="p-4 flex flex-col flex-grow")
        
        for card in cards:
            title_tag = card.find("a", class_="editor_page")
            if not title_tag:
                continue
                
            title = title_tag.get_text(strip=True)
            href = title_tag.get("href", "")
            full_link = href if href.startswith("http") else f"https://www.for9a.com{href}"
            
            country_tag = card.find("span", class_="bg-gray-100")
            country = country_tag.get_text(strip=True) if country_tag else "International"
            
            item = [level, title, country, full_link]
            if item not in all_scholarships:
                all_scholarships.append(item)
    else:
        print(f"Failed {level}: {page_data.status_code}")
        
    time.sleep(1)

# excel setup
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Scholarships"

headers = ["Degree", "Title", "Country", "Link"]
ws.append(headers)

# header style
header_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
header_font = Font(bold=True, size=11)

for col_num in range(1, len(headers) + 1):
    cell = ws.cell(row=1, column=col_num)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

# add rows
for row_data in all_scholarships:
    ws.append(row_data)

# formatting
wrap_alignment = Alignment(wrap_text=True, vertical="center")

for row in ws.iter_rows(min_row=2, max_row=len(all_scholarships) + 1, min_col=1, max_col=4):
    for cell in row:
        cell.alignment = wrap_alignment

max_widths = {1: 15, 2: 50, 3: 20, 4: 45}

for col_idx, width in max_widths.items():
    col_letter = get_column_letter(col_idx)
    ws.column_dimensions[col_letter].width = width

# save file
excel_filename = "scholarships.xlsx"
wb.save(excel_filename)

print(f"Done! Saved {len(all_scholarships)} items in '{excel_filename}'.")