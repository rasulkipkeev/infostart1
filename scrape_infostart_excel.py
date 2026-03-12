import csv
import html
import logging
import random
import re
import time
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger(__name__)

BASE = "https://infostart.ru"
LIST_PATH = "/public/all/integraciya_i_obmen_dannymi"
SORT = "property_count_download"
OUT_DIR = Path(r"D:\prog\infostart")

# Динамические имена файлов на основе LIST_PATH
base_name = LIST_PATH.strip('/').split('/')[-1]
if not base_name:
    base_name = "infostart_data"
CSV_PATH = OUT_DIR / f"{base_name}.csv"
XLSX_PATH = OUT_DIR / f"{base_name}.xlsx"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
}

FIELDS = [
    "page",
    "position_on_page",
    "title",
    "card_url",
    "price",
    "rating",
    "date",
    "views",
    "downloads",
    "comments",
    "author",
    "preview",
    "tags",
    "source_page_url",
]


def page_url(page: int) -> str:
    return f"{BASE}{LIST_PATH}?sort={SORT}&PAGEN_1={page}"


def get_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(HEADERS)
    
    # Настройка повторных попыток (Retries)
    retry_strategy = Retry(
        total=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["HEAD", "GET", "OPTIONS"]
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    
    return session


def get_soup(session: requests.Session, url: str) -> BeautifulSoup:
    resp = session.get(url, timeout=60)
    resp.raise_for_status()
    resp.encoding = "cp1251"
    return BeautifulSoup(resp.text, "html.parser")


def extract_total_pages(soup: BeautifulSoup) -> int:
    pages = []
    for a in soup.find_all("a", href=True):
        href = html.unescape(a["href"])
        m = re.search(r"PAGEN_1=(\d+)", href)
        if m:
            pages.append(int(m.group(1)))
    if not pages:
        logger.warning("Pagination not found, assuming 1 page.")
        return 1
    return max(pages)


def normalize_text(value: str) -> str:
    return " ".join(value.split())


def extract_item(item: BeautifulSoup, page: int, pos: int) -> dict:
    title_link = item.select_one("div.publication-name a[href]")
    if not title_link:
        raise RuntimeError(f"Title link not found on page {page}, item {pos}")

    title = normalize_text(title_link.get_text(" ", strip=True))
    href = html.unescape(title_link["href"])
    card_url = href if href.startswith("http") else f"{BASE}{href}"

    price_node = item.select_one("p.price-block")
    price = normalize_text(price_node.get_text(" ", strip=True)) if price_node else ""

    rating_node = item.select_one("span.obj-rate-count-p")
    rating = normalize_text(rating_node.get_text(" ", strip=True)) if rating_node else ""
    if not rating:
        alt_rating = item.select_one("span.text-nowrap.rate-article")
        rating = normalize_text(alt_rating.get_text(" ", strip=True)) if alt_rating else ""

    preview_node = item.select_one("p.public-preview-text-wrap")
    preview = normalize_text(preview_node.get_text(" ", strip=True)) if preview_node else ""

    tags = []
    for a in item.select("p.public-tags-wrap a.public-tag"):
        txt = normalize_text(a.get_text(" ", strip=True))
        if txt:
            tags.append(txt)

    # Более безопасное извлечение мета-данных (на случай изменения их количества)
    meta_spans = item.select("p.desc-article span.text-nowrap")
    meta = [normalize_text(span.get_text(" ", strip=True)) for span in meta_spans]
    
    def safe_get(lst: list, idx: int) -> str:
        return lst[idx] if idx < len(lst) else ""

    date = safe_get(meta, 0)
    views = safe_get(meta, 1)
    downloads = safe_get(meta, 2)
    author = safe_get(meta, 3)
    comments = safe_get(meta, 4)

    if not downloads or not comments or not views:
        stats = [normalize_text(span.get_text(" ", strip=True)) for span in item.select("div.view-table-right span.text-nowrap")]
        if len(stats) >= 4:
            rating = rating or stats[0]
            downloads = downloads or stats[1]
            comments = comments or stats[2]
            views = views or stats[3]

    return {
        "page": page,
        "position_on_page": pos,
        "title": title,
        "card_url": card_url,
        "price": price,
        "rating": rating,
        "date": date,
        "views": views,
        "downloads": downloads,
        "comments": comments,
        "author": author,
        "preview": preview,
        "tags": " | ".join(tags),
        "source_page_url": page_url(page),
    }


def append_to_csv(row: dict) -> None:
    # Запись одной строки (дозапись)
    file_exists = CSV_PATH.exists()
    with CSV_PATH.open("a", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        if not file_exists:
            writer.writeheader()
        writer.writerow(row)


def col_letter(idx: int) -> str:
    letters = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def xlsx_cell(value: str) -> str:
    value = "" if value is None else str(value)
    return f'<c t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'


def write_xlsx(rows: list[dict]) -> None:
    if not rows:
        return
    widths = []
    for field in FIELDS:
        max_len = max([len(field)] + [len(str(row.get(field, ""))) for row in rows])
        widths.append(min(max(max_len + 2, 12), 80))

    sheet_rows = []
    header_cells = "".join(xlsx_cell(field) for field in FIELDS)
    sheet_rows.append(f'<row r="1">{header_cells}</row>')
    for idx, row in enumerate(rows, start=2):
        cells = "".join(xlsx_cell(row.get(field, "")) for field in FIELDS)
        sheet_rows.append(f'<row r="{idx}">{cells}</row>')

    cols_xml = "".join(
        f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
        for i, w in enumerate(widths, start=1)
    )
    last_col = col_letter(len(FIELDS))
    last_row = len(rows) + 1

    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:{last_col}{last_row}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>{cols_xml}</cols>
  <sheetData>{''.join(sheet_rows)}</sheetData>
</worksheet>'''

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="infostart" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>'''

    workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''

    with zipfile.ZipFile(XLSX_PATH, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr("xl/styles.xml", styles_xml)


def main() -> None:
    # Удаляем старый CSV, если он существует, чтобы дозапись начиналась с чистого листа
    if CSV_PATH.exists():
        CSV_PATH.unlink()
        logger.info(f"Удален старый файл: {CSV_PATH}")

    with get_session() as session:
        logger.info(f"Запрос первой страницы: {page_url(1)}")
        first = get_soup(session, page_url(1))
        total_pages = extract_total_pages(first)
        logger.info(f"Найдено страниц: {total_pages}")
        
        rows = []
        seen_urls = set()

        for page in range(1, total_pages + 1):
            try:
                soup = first if page == 1 else get_soup(session, page_url(page))
                items = soup.select("div.publication-item")
                if not items:
                    logger.warning(f"Пустая страница (нет публикаций): {page}")
                    continue

                page_rows = 0
                for pos, item in enumerate(items, start=1):
                    row = extract_item(item, page, pos)
                    if row["card_url"] in seen_urls:
                        continue
                    seen_urls.add(row["card_url"])
                    rows.append(row)
                    append_to_csv(row)
                    page_rows += 1
                
                logger.info(f"Страница {page}/{total_pages} обработана: {page_rows} новых элементов. Всего: {len(rows)}")
                
                # Рандомизированная задержка
                time.sleep(random.uniform(0.3, 1.0))
            except Exception as e:
                logger.error(f"Ошибка при обработке страницы {page}: {e}")

        if rows:
            logger.info(f"Генерация Excel файла: {XLSX_PATH}")
            write_xlsx(rows)
            logger.info(f"Успешно сохранено XLSX строк: {len(rows)}")
        else:
            logger.warning("Не собрано ни одной строки данных.")


if __name__ == "__main__":
    main()
