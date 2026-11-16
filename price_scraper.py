"""
Project 1: Data Extraction
Scrapes product names, prices, ratings and availability from
books.toscrape.com (a sandbox site built for scraping practice)
and saves a clean CSV + Excel file.

Install:  pip install requests beautifulsoup4 pandas openpyxl
Run:      python price_scraper.py
"""

import re
import time
from datetime import datetime
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://books.toscrape.com/catalogue/page-1.html"
HEADERS = {"User-Agent": "Mozilla/5.0 (learning-project scraper)"}
DELAY_SECONDS = 1  # be polite: pause between requests
MAX_PAGES = 5      # raise this (or set None) to scrape everything

RATING_MAP = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}


def fetch(url: str) -> BeautifulSoup | None:
    """Download a page and return parsed HTML, or None on failure."""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        return BeautifulSoup(resp.text, "html.parser")
    except requests.RequestException as e:
        print(f"  ! Failed to fetch {url}: {e}")
        return None


def parse_products(soup: BeautifulSoup, page_url: str) -> list[dict]:
    """Extract one dict per product card on the page."""
    rows = []
    for card in soup.select("article.product_pod"):
        title = card.h3.a["title"]
        price_text = card.select_one("p.price_color").get_text(strip=True)
        price = float(re.sub(r"[^\d.]", "", price_text))  # '£51.77' -> 51.77
        rating_word = card.select_one("p.star-rating")["class"][1]
        in_stock = "In stock" in card.select_one("p.availability").get_text()
        link = urljoin(page_url, card.h3.a["href"])

        rows.append(
            {
                "title": title,
                "price_gbp": price,
                "rating": RATING_MAP.get(rating_word),
                "in_stock": in_stock,
                "url": link,
            }
        )
    return rows


def scrape(start_url: str, max_pages: int | None = None) -> pd.DataFrame:
    all_rows, url, page = [], start_url, 0

    while url and (max_pages is None or page < max_pages):
        page += 1
        print(f"Scraping page {page}: {url}")
        soup = fetch(url)
        if soup is None:
            break

        all_rows.extend(parse_products(soup, url))

        next_link = soup.select_one("li.next a")
        url = urljoin(url, next_link["href"]) if next_link else None
        time.sleep(DELAY_SECONDS)

    df = pd.DataFrame(all_rows)
    if df.empty:
        return df

    # Clean-up
    df = df.drop_duplicates(subset="url").reset_index(drop=True)
    df["title"] = df["title"].str.strip()
    df["scraped_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    return df.sort_values("price_gbp").reset_index(drop=True)


def save(df: pd.DataFrame, stem: str = "products") -> None:
    df.to_csv(f"{stem}.csv", index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(f"{stem}.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Products")
        ws = writer.sheets["Products"]
        ws.freeze_panes = "A2"
        for col in ws.columns:  # auto-fit column widths
            width = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(width + 2, 60)
    print(f"Saved {len(df)} rows -> {stem}.csv and {stem}.xlsx")


if __name__ == "__main__":
    data = scrape(BASE_URL, max_pages=MAX_PAGES)
    if data.empty:
        print("No data scraped.")
    else:
        save(data)
        print(data.head())
