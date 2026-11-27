"""
Book Scraper - Portfolio Project
Author: Oscar J. Villa García
Description: Scrapes books.toscrape.com and exports data to Excel with formatting.
"""

import argparse
import logging
import math
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

log = logging.getLogger(__name__)

BASE_URL   = "https://books.toscrape.com/catalogue/"
START_URL  = "https://books.toscrape.com/catalogue/page-1.html"
MAX_PAGES  = 10
DELAY      = 0.5
MIN_DELAY  = 0.2
OUTPUT_DIR = Path(__file__).resolve().parent.parent / "output"
RATING_MAP = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

def get_soup(url):
    headers = {"User-Agent": "Mozilla/5.0 (portfolio-scraper/1.0)"}
    response = requests.get(url, headers=headers, timeout=10)
    response.raise_for_status()
    # The pages are UTF-8, so do not rely on the HTTP headers to say so.
    response.encoding = "utf-8"
    return BeautifulSoup(response.text, "html.parser")

def parse_books(soup):
    books = []
    for article in soup.select("article.product_pod"):
        title  = article.h3.a["title"]
        price  = float(article.select_one(".price_color").text.strip().lstrip("£"))
        rating = RATING_MAP.get(article.p["class"][1], 0)
        avail  = article.select_one(".availability").text.strip()
        link   = BASE_URL + article.h3.a["href"].replace("../", "")
        books.append({"Title": title, "Price (£)": price, "Rating": rating, "Availability": avail, "URL": link})
    return books

def scrape(max_pages=MAX_PAGES, delay=DELAY):
    all_books, url, page = [], START_URL, 1
    while url:
        log.info(f"Scraping page {page}")
        soup = get_soup(url)
        all_books.extend(parse_books(soup))
        if max_pages and page >= max_pages:
            break
        next_btn = soup.select_one("li.next a")
        url = BASE_URL + next_btn["href"] if next_btn else None
        page += 1
        time.sleep(delay)
    log.info(f"Done. {len(all_books)} books scraped.")
    return all_books

def export_to_excel(books, output_dir=OUTPUT_DIR):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(books)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = output_dir / f"books_report_{ts}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Books")
        ws = writer.book.active
        fill = PatternFill("solid", fgColor="1F4E79")
        font = Font(bold=True, color="FFFFFF", size=11)
        for col in range(1, len(df.columns)+1):
            cell = ws.cell(row=1, column=col)
            cell.fill = fill
            cell.font = font
        ws.freeze_panes = "A2"
        for col in range(1, len(df.columns)+1):
            ws.column_dimensions[get_column_letter(col)].width = 20
    log.info(f"Excel saved -> {out_path}")
    return out_path

def max_pages_value(text):
    value = int(text)
    if value < 0:
        raise argparse.ArgumentTypeError("must be 0 or a positive integer")
    return value

def delay_value(text):
    value = float(text)
    if not math.isfinite(value) or value < MIN_DELAY:
        raise argparse.ArgumentTypeError(f"must be at least {MIN_DELAY} seconds to avoid overloading the site")
    return value

def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="Scrape books.toscrape.com and export the results to Excel.")
    parser.add_argument("--max-pages", type=max_pages_value, default=MAX_PAGES,
                        help="pages to scrape, 0 means all pages (default: %(default)s)")
    parser.add_argument("--delay", type=delay_value, default=DELAY,
                        help=f"seconds to wait between pages, minimum {MIN_DELAY} (default: %(default)s)")
    parser.add_argument("--output", type=Path, default=OUTPUT_DIR,
                        help="folder for the Excel report (default: the output folder of the project)")
    return parser.parse_args(argv)

def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
    books = scrape(max_pages=args.max_pages, delay=args.delay)
    out_path = export_to_excel(books, args.output)
    print(f"Report ready: {out_path}")
    return out_path

if __name__ == "__main__":
    main()
