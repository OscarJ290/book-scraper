import logging
from datetime import datetime
from pathlib import Path

import pytest
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

import scraper

FIXTURES = Path(__file__).parent / "fixtures"
PAGE_1_URL = scraper.BASE_URL + "page-1.html"
PAGE_2_URL = scraper.BASE_URL + "page-2.html"


def load_soup(name):
    html = (FIXTURES / name).read_text(encoding="utf-8")
    return BeautifulSoup(html, "html.parser")


@pytest.fixture
def requested(monkeypatch):
    """Serve the fixture pages through get_soup and record the requested URLs."""
    urls = []

    def fake_get_soup(url):
        urls.append(url)
        return load_soup(url.rsplit("/", 1)[-1])

    monkeypatch.setattr(scraper, "get_soup", fake_get_soup)
    return urls


@pytest.fixture
def sleeps(monkeypatch):
    """Record the pauses between pages instead of waiting."""
    calls = []
    monkeypatch.setattr(scraper.time, "sleep", calls.append)
    return calls


@pytest.fixture
def broken_page_2(monkeypatch):
    """Serve page 1 and raise a network error on page 2."""
    urls = []

    def fake_get_soup(url):
        urls.append(url)
        if url == PAGE_2_URL:
            raise requests.ConnectionError("connection lost")
        return load_soup("page-1.html")

    monkeypatch.setattr(scraper, "get_soup", fake_get_soup)
    return urls


def all_books():
    books_1, _ = scraper.parse_books(load_soup("page-1.html"))
    books_2, _ = scraper.parse_books(load_soup("page-2.html"))
    return books_1 + books_2


def rating_table(ws):
    """Return the rating table of the Summary sheet as a dict."""
    rows = list(ws.iter_rows(min_col=1, max_col=2, values_only=True))
    start = rows.index(("Rating", "Books")) + 1
    return dict(rows[start:start + 5])


def metrics(ws):
    return {row[0]: row[1] for row in ws.iter_rows(min_row=2, max_row=8, max_col=2, values_only=True)}


def test_network_is_blocked():
    with pytest.raises(RuntimeError, match="Network access is disabled"):
        scraper.get_soup(PAGE_1_URL)


def test_get_soup_decodes_utf8_even_if_headers_say_otherwise(monkeypatch):
    response = requests.Response()
    response.status_code = 200
    response.encoding = "ISO-8859-1"
    response._content = '<p class="price_color">£12.50</p>'.encode("utf-8")
    monkeypatch.setattr(requests, "get", lambda *args, **kwargs: response)

    soup = scraper.get_soup(PAGE_1_URL)

    assert soup.select_one(".price_color").text == "£12.50"


def test_parse_books_extracts_every_field():
    books, _ = scraper.parse_books(load_soup("page-1.html"))

    assert books[0] == {
        "Title": "The Paper Lighthouse",
        "Price (£)": 12.5,
        "Rating": 3,
        "Availability": "In stock",
        "URL": scraper.BASE_URL + "the-paper-lighthouse_1/index.html",
    }
    assert isinstance(books[0]["Price (£)"], float)
    assert isinstance(books[0]["Rating"], int)
    assert books[1]["Title"] == "A Café for Quiet Robots"
    assert books[1]["URL"] == scraper.BASE_URL + "a-cafe-for-quiet-robots_2/index.html"


def test_parse_books_skips_the_incomplete_book(caplog):
    books, skipped = scraper.parse_books(load_soup("page-1.html"))

    assert [book["Title"] for book in books] == [
        "The Paper Lighthouse",
        "A Café for Quiet Robots",
        "Seventeen Imaginary Rivers",
    ]
    assert skipped == 1
    assert "missing price for 'The Missing Label'" in caplog.text


@pytest.mark.parametrize("selector, change", [
    ("h3 a", lambda tag: tag.decompose()),
    ("h3 a", lambda tag: tag.attrs.pop("title")),
    ("h3 a", lambda tag: tag.attrs.pop("href")),
    (".price_color", lambda tag: tag.decompose()),
    (".price_color", lambda tag: setattr(tag, "string", "free")),
    ("p.star-rating", lambda tag: tag.decompose()),
    ("p.star-rating", lambda tag: tag.attrs.update({"class": ["star-rating"]})),
    (".availability", lambda tag: tag.decompose()),
])
def test_parse_books_skips_each_kind_of_incomplete_html(selector, change):
    soup = load_soup("page-2.html")
    change(soup.select_one("article.product_pod").select_one(selector))

    books, skipped = scraper.parse_books(soup)

    assert [book["Title"] for book in books] == ["The Slow Umbrella"]
    assert skipped == 1


def test_scrape_stops_at_max_pages(requested, sleeps):
    result = scraper.scrape(max_pages=1)

    assert requested == [PAGE_1_URL]
    assert result.pages == 1
    assert len(result.books) == 3
    assert sleeps == []


@pytest.mark.parametrize("max_pages", [0, 10])
def test_scrape_stops_when_there_is_no_next_link(requested, sleeps, max_pages):
    result = scraper.scrape(max_pages=max_pages, delay=0.3)

    assert requested == [PAGE_1_URL, PAGE_2_URL]
    assert result.pages == 2
    assert len(result.books) == 5
    assert result.skipped == 1
    assert sleeps == [0.3]


def test_network_error_keeps_the_books_already_collected(broken_page_2, sleeps, caplog):
    result = scraper.scrape(max_pages=0)

    assert broken_page_2 == [PAGE_1_URL, PAGE_2_URL]
    assert result.pages == 1
    assert len(result.books) == 3
    assert "Network error on page 2" in caplog.text


def test_http_error_on_the_first_page_returns_an_empty_result(monkeypatch):
    def fake_get_soup(url):
        raise requests.HTTPError("503 Server Error")

    monkeypatch.setattr(scraper, "get_soup", fake_get_soup)

    result = scraper.scrape()

    assert result.books == []
    assert result.pages == 0


def test_export_creates_the_books_sheet(tmp_path):
    books = all_books()

    path = scraper.export_to_excel(books, tmp_path, pages_visited=2)
    ws = load_workbook(path)["Books"]

    assert [cell.value for cell in ws[1]] == scraper.COLUMNS
    assert ws.max_row == len(books) + 1
    assert ws.freeze_panes == "A2"
    assert ws["A1"].fill.fgColor.rgb.endswith("1F4E79")
    assert ws["A1"].font.bold
    assert ws["B2"].value == 12.5
    assert ws["B2"].number_format == "0.00"
    assert ws["E2"].hyperlink.target == books[0]["URL"]
    assert ws["E6"].hyperlink.target == books[-1]["URL"]
    assert ws["A2"].fill.fill_type is None
    assert ws["A3"].fill.fgColor.rgb.endswith("F2F2F2")
    assert ws["E3"].fill.fgColor.rgb.endswith("F2F2F2")
    assert ws.column_dimensions["A"].width == 60


def test_export_creates_the_summary_sheet_with_a_chart(tmp_path):
    path = scraper.export_to_excel(all_books(), tmp_path, pages_visited=2)
    wb = load_workbook(path)
    ws = wb["Summary"]
    values = metrics(ws)

    assert wb.sheetnames == ["Books", "Summary"]
    assert values["Total books"] == 5
    assert values["Minimum price (£)"] == 7.05
    assert values["Maximum price (£)"] == 40.0
    assert values["Average price (£)"] == pytest.approx(20.0)
    assert values["Books in stock"] == 4
    assert values["Pages visited"] == 2
    assert isinstance(values["Run date and time"], datetime)
    assert rating_table(ws) == {1: 1, 2: 1, 3: 1, 4: 0, 5: 2}
    assert len(ws._charts) == 1


def test_export_with_no_books_still_creates_the_report(tmp_path):
    path = scraper.export_to_excel([], tmp_path)
    wb = load_workbook(path)
    values = metrics(wb["Summary"])

    assert [cell.value for cell in wb["Books"][1]] == scraper.COLUMNS
    assert values["Total books"] == 0
    assert values["Average price (£)"] == "n/a"
    assert rating_table(wb["Summary"]) == {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}


@pytest.mark.parametrize("delay", ["0.1", "0", "-1", "nan", "inf", "abc"])
def test_invalid_delay_is_rejected(delay, capsys):
    with pytest.raises(SystemExit) as error:
        scraper.parse_args(["--delay", delay])

    assert error.value.code == 2
    assert "--delay" in capsys.readouterr().err


def test_minimum_delay_is_accepted():
    assert scraper.parse_args(["--delay", "0.2"]).delay == 0.2


def test_negative_max_pages_is_rejected():
    with pytest.raises(SystemExit):
        scraper.parse_args(["--max-pages", "-1"])


def test_default_arguments():
    args = scraper.parse_args([])

    assert args.max_pages == 10
    assert args.delay == 0.5
    assert args.output == Path(scraper.__file__).resolve().parent.parent / "output"


def test_main_exports_partial_results_and_logs_the_totals(broken_page_2, sleeps, tmp_path, capsys, caplog):
    caplog.set_level(logging.INFO)

    path = scraper.main(["--max-pages", "0", "--output", str(tmp_path / "report")])

    ws = load_workbook(path)["Summary"]
    assert path.parent == tmp_path / "report"
    assert metrics(ws)["Total books"] == 3
    assert metrics(ws)["Pages visited"] == 1
    assert "Books saved: 3. Books skipped: 1." in caplog.text
    assert capsys.readouterr().out.strip() == f"Report ready: {path}"
