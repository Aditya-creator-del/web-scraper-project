# Web Scraper: Product Price Extractor

A Python script that scrapes product data from a website and saves it as a
clean CSV and Excel file.

## Features
- Scrapes title, price, rating, stock status, and product URL
- Handles pagination automatically
- Cleans and deduplicates data with pandas
- Exports to `.csv` and formatted `.xlsx`
- Polite scraping with a delay between requests

## Tech Stack
Python, Requests, BeautifulSoup, pandas, openpyxl

## Setup
```
pip install -r requirements.txt
python price_scraper.py
```

## Output
- `products.csv`
- `products.xlsx`

## Customizing
Change `BASE_URL` and the CSS selectors in `parse_products()` to scrape
a different site (job listings, news headlines, etc.).

## Disclaimer
For educational purposes. Always check a site's robots.txt and terms of
service before scraping.
