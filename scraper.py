"""
Book Price Scraper & Excel Report Generator
-------------------------------------------
Scrapes every book from books.toscrape.com (a practice site made for
scraping), cleans the data with pandas, and exports a formatted Excel
report with a summary sheet.

Usage:
    pip install -r requirements.txt
    python scraper.py
"""

import time
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

START_URL = "https://books.toscrape.com/catalogue/page-1.html"
OUTPUT_FILE = "books_report.xlsx"
DELAY_SECONDS = 0.5  # be polite: pause between requests
HEADERS = {"User-Agent": "Mozilla/5.0 (portfolio project)"}
RATING_MAP = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}


def fetch(url, retries=3):
    """Download a page, retrying on network errors. Returns HTML or None."""
    for attempt in range(1, retries + 1):
        try:
            response = requests.get(url, headers=HEADERS, timeout=15)
            if response.status_code == 404:
                return None
            response.raise_for_status()
            response.encoding = "utf-8"  # avoids broken "Â£" characters
            return response.text
        except requests.RequestException as error:
            print(f"  Attempt {attempt}/{retries} failed: {error}")
            time.sleep(2)
    return None


def parse_page(html, page_url):
    """Extract book records and the next-page URL from one listing page."""
    soup = BeautifulSoup(html, "html.parser")
    books = []

    for item in soup.select("article.product_pod"):
        link = item.h3.a
        books.append(
            {
                "title": link["title"],
                "price_raw": item.select_one("p.price_color").get_text(strip=True),
                "rating_raw": item.select_one("p.star-rating")["class"][1],
                "availability": item.select_one("p.availability").get_text(strip=True),
                "url": urljoin(page_url, link["href"]),
            }
        )

    next_link = soup.select_one("li.next a")
    next_url = urljoin(page_url, next_link["href"]) if next_link else None
    return books, next_url


def scrape_all(start_url):
    """Follow the 'next' links until there are no more pages."""
    all_books = []
    url, page_number = start_url, 1

    while url:
        print(f"Scraping page {page_number}...")
        html = fetch(url)
        if html is None:
            print("  Could not load this page, stopping.")
            break
        books, url = parse_page(html, url)
        all_books.extend(books)
        page_number += 1
        time.sleep(DELAY_SECONDS)

    return all_books


def clean(records):
    """Turn raw records into a clean, analysis-ready DataFrame."""
    df = pd.DataFrame(records)

    df = df.drop_duplicates(subset="url")

    df["price"] = pd.to_numeric(
        df["price_raw"].str.replace(r"[^\d.]", "", regex=True), errors="coerce"
    )
    df["rating"] = df["rating_raw"].map(RATING_MAP)
    df["in_stock"] = df["availability"].str.contains("In stock", case=False)

    df = df.dropna(subset=["title", "price"])
    df = df[["title", "price", "rating", "in_stock", "url"]]
    df = df.sort_values("title").reset_index(drop=True)
    return df


def build_summary(df):
    """Create a small summary table for the second sheet."""
    cheapest = df.loc[df["price"].idxmin()]
    priciest = df.loc[df["price"].idxmax()]

    rows = [
        ("Total books", len(df)),
        ("Average price (£)", round(df["price"].mean(), 2)),
        ("Median price (£)", round(df["price"].median(), 2)),
        ("Cheapest book", f"{cheapest['title']} (£{cheapest['price']:.2f})"),
        ("Most expensive book", f"{priciest['title']} (£{priciest['price']:.2f})"),
        ("Books in stock", int(df["in_stock"].sum())),
        ("", ""),
        ("Books per rating", ""),
    ]
    for rating, count in df["rating"].value_counts().sort_index().items():
        avg = df.loc[df["rating"] == rating, "price"].mean()
        rows.append((f"  {int(rating)} star", f"{count} books, avg £{avg:.2f}"))

    return pd.DataFrame(rows, columns=["Metric", "Value"])


def style_sheet(ws, widths):
    """Apply header colors, column widths, filters and a frozen header row."""
    header_fill = PatternFill("solid", start_color="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "A2"


def export_excel(df, summary, path):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Books", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

        books_ws = writer.sheets["Books"]
        style_sheet(books_ws, [60, 10, 10, 10, 70])
        books_ws.auto_filter.ref = books_ws.dimensions
        for row in books_ws.iter_rows(min_row=2, min_col=2, max_col=2):
            row[0].number_format = '"£"#,##0.00'

        style_sheet(writer.sheets["Summary"], [28, 60])


def main():
    records = scrape_all(START_URL)
    if not records:
        print("No data collected. Check your internet connection.")
        return

    df = clean(records)
    summary = build_summary(df)
    export_excel(df, summary, OUTPUT_FILE)
    print(f"\nDone! {len(df)} books saved to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
