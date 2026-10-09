import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import time
import json
from datetime import datetime, timezone

# استخدام Session للحفاظ على أداء الطلبات والكوكيز
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
    "Connection": "keep-alive"
})

data_list = []

# ==========================================
# 1. كشط موقع فرصة (For9a)
# ==========================================
urls_for9a = {
    "bachelor": "https://www.for9a.com/en/opportunity/category/Scholarships/bachelor",
    "master": "https://www.for9a.com/en/opportunity/category/Scholarships/master",
    "phd": "https://www.for9a.com/en/opportunity/category/Scholarships/phd"
}

max_pages_for9a = 8

for deg, base_url in urls_for9a.items():
    print(f"[For9a] Fetching level: {deg}...")
    
    for page in range(1, max_pages_for9a + 1):
        target_url = f"{base_url}?page={page}"
        print(f"  --> Page {page}...")
        
        res = None
        for attempt in range(3):
            try:
                res = session.get(target_url, timeout=25)
                if res.status_code == 200:
                    break
            except Exception as e:
                print(f"      [For9a] Attempt {attempt+1} failed: {e}")
                time.sleep(3)

        if not res or res.status_code != 200:
            print(f"      [For9a] Skipped page {page} due to HTTP status: {res.status_code if res else 'No Response'}")
            continue
            
        soup = BeautifulSoup(res.content, "lxml")
        cards = soup.find_all("div", class_="p-4 flex flex-col flex-grow")
        
        if not cards:
            print("      No more cards found.")
            break
            
        for card in cards:
            card_text = card.get_text().lower()
            if "closed" in card_text or "مغلق" in card_text:
                continue

            title_tag = card.find("a", class_="editor_page")
            if not title_tag:
                continue
                
            title = title_tag.get_text(strip=True)
            href = title_tag.get("href", "")
            link = href if href.startswith("http") else f"https://www.for9a.com{href}"
            
            c_tag = card.find("span", class_="bg-gray-100")
            country = c_tag.get_text(strip=True) if c_tag else "International"
            
            deadline_tag = card.find("span", class_=lambda x: x and "bg-orange-50" in x)
            deadline = deadline_tag.get_text(strip=True) if deadline_tag else "N/A"
            
            item = [deg, title, country, deadline, link]
            if item not in data_list:
                data_list.append(item)
                
        time.sleep(2)

# ==========================================
# 2. كشط موقع Opportunity Desk
# ==========================================
urls_od = {
    "bachelor": "https://opportunitydesk.org/category/fellowships-and-scholarships/undergraduate/",
    "master": "https://opportunitydesk.org/category/fellowships-and-scholarships/masters-postgraduate/",
    "phd": "https://opportunitydesk.org/category/fellowships-and-scholarships/phd-post-doctoral/"
}

max_pages_od = 5

for deg, base_url in urls_od.items():
    print(f"\n[OpportunityDesk] Fetching level: {deg}...")
    
    for page in range(1, max_pages_od + 1):
        target_url = base_url if page == 1 else f"{base_url}page/{page}/"
        print(f"  --> Page {page}...")
        
        res = None
        for attempt in range(3):
            try:
                res = session.get(target_url, timeout=25)
                if res.status_code == 200:
                    break
            except Exception as e:
                print(f"      [OpportunityDesk] Attempt {attempt+1} failed: {e}")
                time.sleep(3)
                
        if not res or res.status_code != 200:
            print(f"      [OpportunityDesk] Skipped page {page} due to HTTP status: {res.status_code if res else 'No Response'}")
            continue
            
        soup = BeautifulSoup(res.content, "lxml")
        articles = soup.find_all("article")
        
        if not articles:
            articles = soup.find_all("div", class_=lambda x: x and "post" in x)
            
        if not articles:
            print("      No more articles found.")
            break
            
        for art in articles:
            art_text = art.get_text(strip=True).upper()
            
            is_closed = "CLOSED" in art_text
            closed_badge = art.find(lambda tag: tag.name in ["span", "div", "a"] and "CLOSED" in tag.get_text().upper())
            if closed_badge:
                is_closed = True

            if is_closed:
                continue

            title_tag = art.find(["h2", "h3"])
            if not title_tag:
                continue
                
            a_tag = title_tag.find("a")
            if not a_tag:
                continue
                
            title = a_tag.get_text(strip=True)
            link = a_tag.get("href", "")
            
            country = "International"
            deadline = "Check Official Website"
            
            item = [deg, title, country, deadline, link]
            if item not in data_list:
                data_list.append(item)
                
        time.sleep(2)

# ==========================================
# 3. إعادة ترتيب البيانات (تجميع حسب المرحلة)
# ==========================================
degree_order = {"bachelor": 1, "master": 2, "phd": 3}
data_list.sort(key=lambda item: degree_order.get(item[0].lower(), 4))

print(f"\nTotal collected active scholarships: {len(data_list)}")

# ==========================================
# 4. معالجة وإنشاء ملف Excel
# ==========================================
if not data_list:
    raise SystemExit("No active data collected")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Scholarships"

cols = ["Degree", "Title", "Country", "Deadline", "Link"]
ws.append(cols)

fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
font = Font(bold=True, size=11)

for i in range(1, 6):
    c = ws.cell(row=1, column=i)
    c.fill = fill
    c.font = font
    c.alignment = Alignment(horizontal="center", vertical="center")

for row in data_list:
    ws.append(row)

align = Alignment(wrap_text=True, vertical="center")
for row in ws.iter_rows(min_row=2, max_row=len(data_list) + 1, min_col=1, max_col=5):
    for c in row:
        c.alignment = align

widths = {1: 15, 2: 45, 3: 18, 4: 25, 5: 40}
for idx, w in widths.items():
    ws.column_dimensions[get_column_letter(idx)].width = w

excel_filename = "scholarships.xlsx"
try:
    wb.save(excel_filename)
    print(f"Successfully saved to {excel_filename}")
except PermissionError:
    wb.save("scholarships_new.xlsx")
    print("Saved to scholarships_new.xlsx due to permission error.")

with open("last_updated.json", "w") as f:
    json.dump({"updated": datetime.now(timezone.utc).isoformat()}, f)