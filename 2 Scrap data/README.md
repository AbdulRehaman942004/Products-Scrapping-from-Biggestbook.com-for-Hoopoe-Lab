# Product Scraper

This script scrapes product information from the BiggestBook website and updates the Excel file.

## Requirements

1. **Python 3.7+**
2. **Chrome Browser** (latest version)
3. **ChromeDriver** - Must match your Chrome version

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Install ChromeDriver:
   - Download from: https://chromedriver.chromium.org/downloads
   - Make sure the version matches your Chrome browser version
   - Add ChromeDriver to your PATH, or place it in the same folder as the script

## Usage

Run the script:
```bash
python scrape_products.py
```

## What it does

1. Reads the Excel file `ScrappedProducts.xlsx` from the parent folder
2. **Creates a backup** of the Excel file before starting (stored in `Backups` folder)
3. For each product link:
   - Accesses the product page
   - Extracts the unit from the price (e.g., "$1,053.27 /EA" → "EA")
   - Compares with the "Unit of Measure" column in Excel
   - If units match, scrapes:
     - **Image URL**: Product image source
     - **Product Name**: "Global Product Type" from Product Details section
     - **Description**: Full description text
   - If units don't match, writes "Unit not matched" in all columns
4. Saves progress every 10 rows
5. Updates the Excel file with scraped data
6. **Creates a backup** after completion

## Backup System

The script automatically creates backups to prevent data loss:
- Backups are stored in the `Backups` folder (created automatically in the parent directory)
- Backup files are named with timestamps: `ScrappedProducts_backup_YYYYMMDD_HHMMSS.xlsx`
- **Only the last 5 backups are kept** - older backups are automatically deleted
- A backup is created:
  - Before starting any scraping operation
  - After completing a scraping operation

## Notes

- The script skips rows that already have data (unless it's "Unit not matched")
- Progress is saved every 10 rows to prevent data loss
- There's a 2-second delay between requests to avoid overwhelming the server
- The script runs in headless mode (no browser window)

