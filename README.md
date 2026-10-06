Scholarship Aggregator and Directory

This is an interactive web interface which lets students find, sift through and view fully funded scholarship options available globally for bachelor, master and Ph.D. level studies.

Features

Web Scraping: Python scripts automatically extract scholarship data from online sources and save it directly into an Excel sheet.

Advanced Filters: Instant search and filter capability based on degree, destination or keywords.

Multi-Language Support: English, Arabic, Russian, Spanish languages supported by automatic directionality (RTL/LTR).

Dynamic Excel Data Import: Extracting information from a .xlsx file locally via SheetJS.

Responsive Design: Works across various desktop/tablet/mobile browsers.

No Dependencies: Very light-weight, free to host directly on GitHub Pages without any API key.

Technology Stack

Web Scraping: Python (requests, BeautifulSoup, Pandas)

Frontend: HTML5 / CSS3 / JavaScript (ES6+)

Data Processing: SheetJS (xlsx)

Hosting: GitHub Pages

Technical Details

Python script fetches scholarship links, cleans up titles and degrees, and outputs them into scholarships.xlsx.

JavaScript reads the Excel file on page load and renders the results dynamically in the HTML table.

Language switching updates all text on the fly and adjusts document direction (RTL for Arabic).

Repository Structure

index.html: HTML structure of the website

style.css: Style of the website with RTL/LTR design

script.js: Functionality of the website (filters, translations, excel processing)

scholarships.xlsx: The excel sheet that contains scholarships

Using and Maintaining

Local setup: Just open index.html file in the browser.

Updating data: Run the Python script to refresh scholarships.xlsx with new entries.
