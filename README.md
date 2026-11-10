# Book Price Scraper & Excel Report Generator

A Python tool that scrapes every book from [books.toscrape.com](https://books.toscrape.com) (a practice site built for scraping), cleans the data, and produces a formatted Excel report.

## What it does

1. Crawls all listing pages automatically (follows "next" links)
2. Extracts title, price, star rating, stock status, and product URL
3. Cleans the data with pandas (removes duplicates, converts prices to numbers, maps ratings to 1-5)
4. Exports `books_report.xlsx` with two sheets:
   - **Books**: full data, filterable, with a frozen header and £ price formatting
   - **Summary**: totals, average/median price, cheapest and most expensive book, breakdown by rating

## Setup

```bash
pip install -r requirements.txt
python scraper.py
```

## Tech used

Python, requests, BeautifulSoup, pandas, openpyxl

## Features

- Retries on network errors
- Polite delay between requests
- Easy to adapt to other websites by changing the parsing function

## Need something like this for your own data?

Contact me: *add your Fiverr/Upwork/email link here*
