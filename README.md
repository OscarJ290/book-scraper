# 📚 Book Scraper — Portfolio Project

A professional Python web scraper that extracts book data from [books.toscrape.com](https://books.toscrape.com) and generates a formatted Excel report with charts and summaries.

Built by **Oscar J. Villa García** · [ojviga@gmail.com](mailto:ojviga@gmail.com)

---

## ✨ Features

- 🔍 Scrapes title, price, rating, availability and URL across multiple pages
- 📊 Exports to a **styled Excel workbook** with:
  - Formatted data table with alternating rows and clickable links
  - Summary sheet with key metrics
  - Bar chart of books by rating
- ⚙️ Configurable: set max pages, delay between requests
- 🪵 Logging with timestamps at every step

---

## 🚀 Quick Start

```bash
git clone https://github.com/oscarjvilla290/book-scraper.git
cd book-scraper
pip install -r requirements.txt
python src/scraper.py
```

---

## 📁 Project Structure

```
book-scraper/
├── src/
│   └── scraper.py
├── output/
├── requirements.txt
└── README.md
```

---

## 🛠️ Tech Stack

- **Python 3.11+** · `requests` · `BeautifulSoup4` · `pandas` · `openpyxl`

---

## 📄 License

MIT
