"""
Дополнение недостающих полей каталога данными с bosch-professional.com.

Как это работает:

- Карточка товара открывается по прямой ссылке
  ``https://www.bosch-professional.com/kz/ru/products/x-<артикул>`` —
  человекочитаемая часть перед артикулом сайтом игнорируется (проверено
  вручную), поэтому специальный поиск не нужен: ссылка строится по одному
  артикулу.
- HTML отдаётся сразу при обычном запросе (без выполнения JavaScript),
  поэтому подойдёт обычный `requests`.
- Со страницы товара реально можно достать: модель, тип товара (категория),
  краткое техническое описание (из таблицы характеристик) и один вес
  (трактуется как вес нетто самого инструмента). Код ТН ВЭД, EAN, страну
  происхождения, сертификаты сайт по отдельному товару не публикует — эти
  поля остаются только из Excel-источников.
- Перед тем как принять данные страницы, скрипт сверяет «Номер заказа» на
  странице с искомым артикулом: если не совпало (или страница не найдена) —
  ничего не берётся.
- Дополняются только пустые поля — уже заполненные (из каталога/прайс-листа/
  реестра ДТ) никогда не перезаписываются.

Прежде чем гонять это по всей базе, проверьте правила использования сайта
(robots.txt / условия использования) сами — при разработке этого модуля
сайт был недоступен из окружения, где писался код, и это не проверялось.
"""
from __future__ import annotations

import re
import time
from dataclasses import dataclass
from html.parser import HTMLParser

import requests

PRODUCT_URL_TEMPLATE = "https://www.bosch-professional.com/kz/ru/products/x-{article}"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36 catalog_agent/1.0"
)

BRAND_PREFIXES = ["PRO", "BLUE", "ADVANCED", "EASY", "UNEO", "GREEN"]

ORDER_NUMBER_RE = re.compile(r"Номер[ \t]+заказа[ \t]*:?[ \t]*\'?[ \t]*([0-9A-Za-zА-Яа-я][0-9A-Za-zА-Яа-я\. \t]{4,20})")
WEIGHT_RE = re.compile(r"([\d]+[.,]\d+)\s*кг")


def build_product_url(article: str) -> str:
    return PRODUCT_URL_TEMPLATE.format(article=article)


def _normalize_order_number(raw: str) -> str:
    return re.sub(r"[^0-9A-ZА-Я]", "", raw.upper())


_BLOCK_TAGS = {
    "p", "div", "li", "tr", "br", "h1", "h2", "h3", "h4", "h5", "h6",
    "section", "article", "ul", "ol", "table", "thead", "tbody", "header", "footer",
}
_SKIP_TAGS = {"script", "style", "noscript", "template"}


class _VisibleTextExtractor(HTMLParser):
    """Извлекает видимый текст из HTML, как это делает браузер при Ctrl+A/Ctrl+C —
    честный парсер тегов (stdlib), а не регулярки, чтобы не путать реальный
    текст вроде '< 1 мВт' с началом тега."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self._chunks: list[str] = []
        self._skip_depth = 0

    def handle_starttag(self, tag, attrs):
        if tag in _SKIP_TAGS:
            self._skip_depth += 1
        elif tag in _BLOCK_TAGS:
            self._chunks.append("\n")

    def handle_endtag(self, tag):
        if tag in _SKIP_TAGS and self._skip_depth:
            self._skip_depth -= 1
        elif tag in _BLOCK_TAGS:
            self._chunks.append("\n")

    def handle_data(self, data):
        if not self._skip_depth:
            self._chunks.append(data)

    def get_text(self) -> str:
        return "".join(self._chunks)


def _strip_html(html: str) -> str:
    parser = _VisibleTextExtractor()
    parser.feed(html)
    text = parser.get_text()
    text = re.sub(r"[ \t]+", " ", text)
    lines = [ln.strip() for ln in text.splitlines()]
    return "\n".join(ln for ln in lines if ln)


def _extract_specs_block(text: str) -> list[str]:
    """Строки между заголовком 'ТЕХНИЧЕСКИЕ ХАРАКТЕРИСТИКИ' и футером-сноской."""
    lines = text.splitlines()
    start = None
    for i, ln in enumerate(lines):
        if ln.strip().upper().startswith("ТЕХНИЧЕСКИЕ ХАРАКТЕРИСТИКИ"):
            start = i + 1
            break
    if start is None:
        return []
    block = []
    for ln in lines[start:]:
        s = ln.strip()
        if not s:
            continue
        if s.upper().startswith("ДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ"):
            continue
        if s.startswith("*"):
            break
        if s.upper().endswith("ДОПОЛНИТЕЛЬНЫЕ СВЕДЕНИЯ") or ": ДОПОЛНИТЕЛЬНЫЕ СВЕДЕНИЯ" in s.upper():
            break
        block.append(s)
    return block


def _pair_spec_lines(lines: list[str]) -> list[tuple[str, str]]:
    """label/value идут строго парами строк подряд."""
    pairs = []
    i = 0
    while i + 1 < len(lines):
        label = lines[i].rstrip("*").strip()
        value = lines[i + 1].strip()
        pairs.append((label, value))
        i += 2
    return pairs


@dataclass
class ScrapedProduct:
    article: str
    model: str | None = None
    category: str | None = None
    short_description: str | None = None
    weight_kg: float | None = None


def parse_product_page(html: str, article: str) -> ScrapedProduct | None:
    text = _strip_html(html)

    order_match = ORDER_NUMBER_RE.search(text)
    if not order_match:
        return None
    if _normalize_order_number(order_match.group(1)) != re.sub(r"[^0-9A-ZА-Я]", "", article.upper()):
        return None  # страница не про этот артикул — не берём ничего

    lines = [ln for ln in text.splitlines() if ln.strip()]
    model = None
    category = None
    for i, ln in enumerate(lines):
        s = ln.strip()
        upper_words = s.split()
        if upper_words and upper_words[0].upper() in BRAND_PREFIXES and len(s) < 60:
            model = " ".join(upper_words[1:]).strip()
            if i + 1 < len(lines):
                category = lines[i + 1].strip()
            break

    spec_lines = _extract_specs_block(text)
    pairs = _pair_spec_lines(spec_lines)

    weight_kg = None
    desc_parts = []
    for label, value in pairs:
        if weight_kg is None and label.lower().startswith("вес"):
            m = WEIGHT_RE.search(value)
            if m:
                weight_kg = float(m.group(1).replace(",", "."))
        desc_parts.append(f"{label} - {value}")

    short_description = "; ".join(desc_parts) if desc_parts else None

    return ScrapedProduct(
        article=article,
        model=model,
        category=category,
        short_description=short_description,
        weight_kg=weight_kg,
    )


def fetch_product(article: str, session: requests.Session, timeout: float = 15.0) -> ScrapedProduct | None:
    url = build_product_url(article)
    try:
        resp = session.get(url, timeout=timeout, headers={"User-Agent": USER_AGENT})
    except requests.RequestException:
        return None
    if resp.status_code != 200:
        return None
    return parse_product_page(resp.text, article)


ENRICHABLE_FIELDS = {
    "модель": lambda p: p.model,
    "Наименование": lambda p: p.category,
    "Краткое техническое описание": lambda p: p.short_description,
    "ВЕС нетто": lambda p: p.weight_kg,
}


def enrich_records(
    records: list[dict],
    fields: list[str] | None = None,
    delay: float = 0.5,
    limit: int | None = None,
    on_progress=None,
) -> dict:
    """Дополняет пустые поля записей данными с сайта. Изменяет records на месте.

    Возвращает отчёт: {"checked": int, "filled": int, "not_found": int, "no_new_data": int}.
    """
    target_fields = fields or list(ENRICHABLE_FIELDS.keys())
    candidates = [r for r in records if not all(r.get(f) for f in target_fields)]
    if limit is not None:
        candidates = candidates[:limit]

    report = {"checked": 0, "filled": 0, "not_found": 0, "no_new_data": 0}
    session = requests.Session()

    for rec in candidates:
        article = rec.get("артикул")
        if not article:
            continue
        report["checked"] += 1
        scraped = fetch_product(article, session)
        if on_progress:
            on_progress(article, scraped)
        if scraped is None:
            report["not_found"] += 1
            time.sleep(delay)
            continue

        filled_any = False
        for field in target_fields:
            if rec.get(field):
                continue
            value = ENRICHABLE_FIELDS[field](scraped)
            if value:
                rec[field] = value
                filled_any = True
        if filled_any:
            report["filled"] += 1
        else:
            report["no_new_data"] += 1

        time.sleep(delay)

    return report
