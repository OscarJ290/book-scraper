# Book Scraper

A Python script that collects book data from [books.toscrape.com](https://books.toscrape.com) and saves it to a formatted Excel report.

## What it does

The script reads the catalogue pages one by one and follows the link to the next page. For every book it collects the title, price, star rating, availability and URL. It then writes an Excel file with the full list of books, a summary and a bar chart.

books.toscrape.com is a practice site made for learning web scraping.

## Setup on Windows

You need Git and Python 3.12 or 3.14. These commands are for PowerShell. If you use Python 3.12, replace `3.14` with `3.12`.

```powershell
git clone https://github.com/OscarJ290/book-scraper.git
cd book-scraper
py -3.14 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Usage

Run the script with the default settings:

```powershell
python src/scraper.py
```

It accepts three optional arguments:

* `--max-pages` sets how many pages to scrape. The default is `10`. Use `0` to scrape all pages.
* `--delay` sets the pause between pages in seconds. The default is `0.5`. Values below `0.2` are rejected so the site is not overloaded.
* `--output` sets the folder for the Excel report. The default is the `output` folder in the project root. The folder is created if it does not exist.

Examples:

```powershell
python src/scraper.py --max-pages 2
```

```powershell
python src/scraper.py --max-pages 0 --delay 1
```

```powershell
python src/scraper.py --max-pages 5 --output C:\Reports\books
```

Each run creates a new file named `books_report_YYYYMMDD_HHMMSS.xlsx`. The log shows the progress and ends with the number of books saved and skipped.

If a book has incomplete HTML, the script logs a warning, skips that book and counts it. If a page fails to load because of a network or HTTP error, the script logs the error, stops requesting pages and exports the books it already has.

## The Excel report

### Books sheet

* One row per book with the columns Title, Price (£), Rating, Availability and URL
* A styled header row that stays visible while you scroll
* Alternating rows with a light grey fill
* Prices shown with two decimals
* Clickable links in the URL column

### Summary sheet

* Total number of books
* Minimum, maximum and average price
* Number of books in stock
* Number of pages visited
* Date and time of the run, taken from your computer clock
* A table with the number of books for each rating from 1 to 5
* A bar chart built from that table

If no books were collected, the three price values show `n/a`.

## Running the tests

With the virtual environment active:

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest
```

The tests do not use the network. A shared fixture replaces `requests.get` with a function that raises an error, and the scraping tests read handwritten HTML pages with invented books from `tests/fixtures`.

A GitHub Actions workflow runs the tests on every push and pull request with Python 3.12 and 3.14.

## Known limitations

* The script depends on the HTML of books.toscrape.com and will break if that HTML changes.
* It does not run JavaScript.
* It has no proxy support and no handling for blocked requests.
* It does not retry failed requests.
* A run that stops early because of a network error still writes the report with the books collected so far, but exits with code 1 and says so in the final message. If the first page fails, the report contains no books. A run that collects no books also exits with code 1.
* It reads only the catalogue pages, not the detail page of each book.
* Prices are in pounds, as the site shows them. There is no currency conversion.

## License

MIT. See the [LICENSE](LICENSE) file.
