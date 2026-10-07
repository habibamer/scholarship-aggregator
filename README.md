# Automated Scholarship Aggregator & Pipeline

An automated data pipeline and web directory that collects and presents scholarship opportunities for Bachelor, Master, and PhD programs, updated daily.

The project uses a Python script to fetch data, GitHub Actions to schedule daily runs, SheetJS for client-side Excel parsing, and Vanilla JavaScript for the interactive, multi-language interface.

- **Live Demo:** https://habibamer.github.io/scholarship-aggregator/
- **Author:** Habib Amer

## System Architecture & Workflow

1. **Scheduled Execution:** GitHub Actions triggers the scraper daily at 00:00 UTC (`scraper.yml`).
2. **Data Scraping:** A Python script (`web.py`) fetches scholarship listings from the source portal using `requests` and `BeautifulSoup`.
3. **File Generation:** The script writes the listings to an Excel spreadsheet (`scholarships.xlsx`) and logs an execution timestamp (`last_updated.json`).
4. **Automated Commit:** Updated files are committed back to the repository by a GitHub Actions bot (`github-actions[bot]`).
5. **Client Presentation:** The web app reads the Excel file in the browser via SheetJS and renders the interactive table on GitHub Pages.

## Technical Details

### 1. Data Collection (`web.py`)

- **Category scraping:** Iterates through degree categories (bachelor, master, phd) and paginates through the first 5 pages of each, using BeautifulSoup with `lxml`.
- **HTTP requests:** Sends a standard browser User-Agent header, with timeout controls and a delay (`time.sleep`) between requests to avoid rate limits.
- **Data export:** Uses `openpyxl` to build the workbook with header styling, background fill, text wrapping, and custom column widths (15, 50, 20, 45).
- **Execution log:** Writes a UTC timestamp to `last_updated.json` to record the latest run.

### 2. Automation Pipeline (`.github/workflows/scraper.yml`)

- **Scheduling:** Standard cron syntax (`0 0 * * *`) for daily runs, with `workflow_dispatch` enabled for manual execution.
- **Environment:** Runs on `ubuntu-latest` with Python 3.10 (`actions/setup-python@v5`) and installs `requests`, `beautifulsoup4`, `openpyxl`, and `lxml`.
- **Git integration:** Commits the updated files to the `main` branch automatically.

### 3. Frontend (`index.html`, `script.js`, `style.css`)

- **In-browser parsing:** Uses SheetJS (`xlsx.full.min.js`) to parse the binary Excel data (`arrayBuffer`) on the client, with no custom backend.
- **UI translation (i18n):** A translation dictionary for English, Arabic, Russian, and Spanish, switching document direction (`rtl` / `ltr`) accordingly.
- **Client-side filtering:** In-memory filtering by title keyword, degree, and a dynamically built country list, triggered by the Search button or the Enter key.
- **Cache control:** Appends a timestamp (`?t=...`) to `fetch()` calls so browsers do not serve an outdated dataset.

## Data Source & Known Limitations

- **Educational purpose:** Data is collected from for9a.com for educational purposes. Always verify details on the official provider page before applying.
- **Dependency on page structure:** The scraper relies on the source site's HTML structure, so a layout change there may require updating `web.py`.
- **Dataset scope:** Only the first 5 pages per degree category are fetched. The Excel file is overwritten on each run, so older listings that fall off those pages are removed.
- **No expiry check:** Application deadlines are not validated, so some listings may already be closed.
- **Localized interface only:** Language switching applies to the interface; scholarship titles and country names stay in their original English.

## Repository Structure

```
.
├── .github/
│   └── workflows/
│       └── scraper.yml        # Scheduled GitHub Actions workflow
├── index.html                 # Page markup and structure
├── style.css                  # Custom CSS layout
├── script.js                  # Excel parsing, filtering, and i18n logic
├── web.py                     # Python data collection script
├── scholarships.xlsx          # Generated Excel dataset
└── last_updated.json          # Last execution timestamp
```

## Local Setup & Testing

1. Clone the repository:

   ```bash
   git clone https://github.com/habibamer/scholarship-aggregator.git
   cd scholarship-aggregator
   ```

2. Install the Python dependencies:

   ```bash
   pip install requests beautifulsoup4 openpyxl lxml
   ```

3. Run the scraper:

   ```bash
   python web.py
   ```

4. Start a local server for the frontend (opening `index.html` directly via `file://` does not work, because browser security restrictions block `fetch()` for local files):

   ```bash
   python -m http.server 8000
   ```

5. Open 'http://localhost:8000' in your browser.

## Author

Developed and maintained by Habib Amer.
