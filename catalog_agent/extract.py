"""
Сборка каталога товаров установленного образца из разрозненных Excel-источников.

Источники данных (каждый — обычный .xlsx, данные разбросаны по разным листам
и разным наборам колонок):

1. Каталог  — уже накопленная база в целевом формате (TEMPLATE_COLUMNS).
   Служит основой: хранит вручную выверенные поля (ТН ВЭД, сертификаты,
   товарный знак и т.п.), которые нельзя получить из прайс-листов.
2. Прайс-лист — многостраничный прайс поставщика. На каждом листе своя
   шапка (может быть на разных строках, с разным набором колонок), но всегда
   есть колонка "Артикул" и колонка с моделью/наименованием.
3. Реестр ДТ (таможенные декларации) — запасной источник кода ТН ВЭД,
   страны происхождения и единицы измерения по артикулу.

Ключ сопоставления везде один — артикул.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

import openpyxl

TEMPLATE_COLUMNS = [
    "артикул",
    "модель",
    "Наименование",
    "Краткое техническое описание",
    "Код ТНВЭД",
    "мнр",
    "оис",
    "Сертификат",
    "начало",
    "конец",
    "ВЕС нетто",
    "Вес брутто",
    "производитель",
    "товарный знак",
    "марка",
    "страна происхождения",
    "ЕД.измер",
]

# Синонимы заголовков в прайс-листах -> к какому полю шаблона они относятся.
# "Сегмент" явно приоритетнее для поля "Наименование", "Модель" - для поля "модель".
PRICELIST_LABEL_SYNONYMS = {
    "артикул": "артикул",
    "сегмент": "Наименование",
    "модель": "модель",
    "краткие технические данные": "Краткое техническое описание",
    "краткое техническое описание": "Краткое техническое описание",
    "страна производства": "страна происхождения",
    "вес нетто, кг": "ВЕС нетто",
    "вес нетто кг": "ВЕС нетто",
    "вес брутто, кг": "Вес брутто",
    "вес брутто кг": "Вес брутто",
    "тип запасной части": "Наименование",
    "наименование": "_наименование_или_модель",  # обрабатывается отдельно
}

CUSTOMS_LABEL_SYNONYMS = {
    "артикул": "артикул",
    "наименование": "Наименование",
    "код тнвэд": "Код ТНВЭД",
    "страна": "страна происхождения",
    "ед.изм": "ЕД.измер",
}


def norm_article(value) -> str | None:
    """Нормализует артикул: убирает пробелы, ведущий апостроф, приводит к верхнему регистру."""
    if value is None:
        return None
    s = str(value).strip().lstrip("'").strip()
    return s.upper() if s else None


def norm_label(value) -> str | None:
    """Приводит заголовок ячейки к сравнимому виду: без переносов строк,
    без второго (англ.) варианта после '/', без лишних пробелов, в нижнем регистре.
    Числовые ячейки (номера колонок в служебных строках) в расчёт не берутся.
    """
    if not isinstance(value, str):
        return None
    s = value.replace("\n", " ").strip()
    if not s:
        return None
    s = s.split("/", 1)[0].strip().rstrip(",").strip()
    return s.lower() if s else None


def clean_unit(value) -> str | None:
    """'ШТ, ШТ' -> 'ШТ' и т.п.: реестр ДТ иногда дублирует единицу измерения через запятую."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    parts = [p.strip() for p in s.split(",") if p.strip()]
    return parts[0] if parts else s


@dataclass
class SourceStats:
    path: str
    sheets_used: list[str] = field(default_factory=list)
    sheets_skipped: list[str] = field(default_factory=list)
    rows_read: int = 0
    unique_articles: int = 0


def find_header_row(rows: list[tuple]) -> int | None:
    """Ищет строку-шапку таблицы в первых `len(rows)` строках листа.

    Шапки в этих прайс-листах бывают продублированы (билингвальная строка,
    затем упрощённая русская строка, затем строка-нумерация колонок).
    Нужна именно последняя "текстовая" строка с колонкой "артикул" —
    следующая за ней строка уже данные. Берём последнее совпадение.
    """
    best = None
    for idx, row in enumerate(rows):
        labels = {norm_label(c) for c in row}
        if "артикул" in labels:
            best = idx
    return best


def read_pricelist(path: str | Path, stats: SourceStats | None = None) -> dict[str, dict]:
    """Читает многостраничный прайс-лист, возвращает {артикул: {поле_шаблона: значение}}."""
    path = str(path)
    if stats is None:
        stats = SourceStats(path=path)
    result: dict[str, dict] = {}

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        scan_limit = min(ws.max_row or 0, 60)
        if scan_limit == 0:
            stats.sheets_skipped.append(sheet_name)
            continue
        head_rows = list(ws.iter_rows(min_row=1, max_row=scan_limit, values_only=True))
        header_idx = find_header_row(head_rows)
        if header_idx is None:
            stats.sheets_skipped.append(sheet_name)
            continue

        headers = head_rows[header_idx]
        col_by_label: dict[str, int] = {}
        for col_idx, cell in enumerate(headers):
            label = norm_label(cell)
            if label:
                col_by_label.setdefault(label, col_idx)

        art_col = col_by_label.get("артикул")
        if art_col is None:
            stats.sheets_skipped.append(sheet_name)
            continue

        has_segment = "сегмент" in col_by_label
        model_col = col_by_label.get("модель")
        naming_fallback_col = col_by_label.get("наименование")  # напр. Service: "Наименование/Model Name"

        header_row_number = header_idx + 1  # 1-based номер строки в файле
        sheet_rows = 0
        for row in ws.iter_rows(min_row=header_row_number + 1, values_only=True):
            if art_col >= len(row):
                continue
            article = norm_article(row[art_col])
            if not article:
                continue

            record: dict = {}
            for label, col_idx in col_by_label.items():
                target_field = PRICELIST_LABEL_SYNONYMS.get(label)
                if not target_field or target_field == "_наименование_или_модель":
                    continue
                if col_idx >= len(row):
                    continue
                val = row[col_idx]
                if val is not None and val != "":
                    record[target_field] = val

            # "модель": явная колонка "Модель", иначе (Service-подобные листы) — "Наименование/Model Name"
            if model_col is not None and model_col < len(row) and row[model_col] not in (None, ""):
                record["модель"] = row[model_col]
            elif not has_segment and naming_fallback_col is not None and naming_fallback_col < len(row):
                val = row[naming_fallback_col]
                if val not in (None, ""):
                    record["модель"] = val

            record["артикул"] = article
            result[article] = record
            sheet_rows += 1

        if sheet_rows:
            stats.sheets_used.append(sheet_name)
            stats.rows_read += sheet_rows
        else:
            stats.sheets_skipped.append(sheet_name)

    stats.unique_articles = len(result)
    return result


def read_customs_register(path: str | Path, stats: SourceStats | None = None) -> dict[str, dict]:
    """Читает реестр таможенных деклараций (одна строка = один товар в одной ДТ).
    При нескольких декларациях на один артикул более поздняя запись перекрывает более раннюю
    (порядок в файле = хронологический порядок ДТ)."""
    path = str(path)
    if stats is None:
        stats = SourceStats(path=path)
    result: dict[str, dict] = {}

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows_iter = ws.iter_rows(values_only=True)
        try:
            header = next(rows_iter)
        except StopIteration:
            stats.sheets_skipped.append(sheet_name)
            continue

        col_by_label: dict[str, int] = {}
        for col_idx, cell in enumerate(header):
            label = norm_label(cell)
            if label:
                col_by_label.setdefault(label, col_idx)

        art_col = col_by_label.get("артикул")
        if art_col is None:
            stats.sheets_skipped.append(sheet_name)
            continue

        sheet_rows = 0
        for row in rows_iter:
            if art_col >= len(row):
                continue
            article = norm_article(row[art_col])
            if not article:
                continue
            record: dict = {}
            for label, col_idx in col_by_label.items():
                target_field = CUSTOMS_LABEL_SYNONYMS.get(label)
                if not target_field or col_idx >= len(row):
                    continue
                val = row[col_idx]
                if target_field == "ЕД.измер":
                    val = clean_unit(val)
                if val is not None and val != "":
                    record[target_field] = val
            record["артикул"] = article
            result[article] = record  # последняя по файлу запись побеждает
            sheet_rows += 1

        if sheet_rows:
            stats.sheets_used.append(sheet_name)
            stats.rows_read += sheet_rows

    stats.unique_articles = len(result)
    return result


def read_catalog(path: str | Path, stats: SourceStats | None = None) -> dict[str, dict]:
    """Читает существующий каталог (уже в целевом формате, возможно с листа
    учёта индекса pandas — первая колонка 'Unnamed: 0' игнорируется)."""
    path = str(path)
    if stats is None:
        stats = SourceStats(path=path)
    result: dict[str, dict] = {}

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows_iter = ws.iter_rows(values_only=True)
        try:
            header = next(rows_iter)
        except StopIteration:
            stats.sheets_skipped.append(sheet_name)
            continue

        col_by_label: dict[str, int] = {}
        for col_idx, cell in enumerate(header):
            if not isinstance(cell, str):
                continue
            key = cell.strip()
            if key in TEMPLATE_COLUMNS:
                col_by_label.setdefault(key, col_idx)

        if "артикул" not in col_by_label:
            stats.sheets_skipped.append(sheet_name)
            continue

        art_col = col_by_label["артикул"]
        sheet_rows = 0
        for row in rows_iter:
            if art_col >= len(row):
                continue
            article = norm_article(row[art_col])
            if not article or article == "#N/A":
                continue
            record: dict = {}
            for label, col_idx in col_by_label.items():
                if col_idx >= len(row):
                    continue
                val = row[col_idx]
                if val is not None and val != "" and val != "#N/A":
                    record[label] = val
            record["артикул"] = article
            result[article] = record
            sheet_rows += 1

        if sheet_rows:
            stats.sheets_used.append(sheet_name)
            stats.rows_read += sheet_rows

    stats.unique_articles = len(result)
    return result


def build_records(
    catalog: dict[str, dict],
    pricelist: dict[str, dict],
    customs: dict[str, dict],
    articles: Iterable[str] | None = None,
) -> tuple[list[dict], list[str]]:
    """Сливает три источника в записи целевого формата.

    Приоритет данных по каждому полю:
      - "мнр", "оис", "Сертификат", "начало", "конец", "производитель",
        "товарный знак", "марка" — только из существующего каталога
        (это данные, которые нельзя автоматически извлечь ни из прайса,
        ни из реестра ДТ, их проверяет и вносит человек);
      - "модель", "Наименование", "Краткое техническое описание",
        "ВЕС нетто", "Вес брутто", "страна происхождения" — из прайс-листа,
        если там есть значение (это самые свежие тех.данные), иначе из каталога;
      - "Код ТНВЭД", "ЕД.измер" — из каталога, если там уже есть,
        иначе как запасной вариант — из реестра ДТ.

    Возвращает (записи, артикулы_без_кода_тнвэд).
    """
    if articles is None:
        keys = set(catalog) | set(pricelist)
    else:
        keys = {norm_article(a) for a in articles if norm_article(a)}

    pricelist_priority_fields = [
        "модель",
        "Наименование",
        "Краткое техническое описание",
        "ВЕС нетто",
        "Вес брутто",
        "страна происхождения",
    ]
    catalog_only_fields = [
        "мнр",
        "оис",
        "Сертификат",
        "начало",
        "конец",
        "производитель",
        "товарный знак",
        "марка",
    ]

    records = []
    missing_tnved = []
    for art in sorted(keys):
        cat = catalog.get(art, {})
        pl = pricelist.get(art, {})
        cu = customs.get(art, {})

        out = {"артикул": art}
        for f in catalog_only_fields:
            if cat.get(f) not in (None, ""):
                out[f] = cat[f]

        for f in pricelist_priority_fields:
            if pl.get(f) not in (None, ""):
                out[f] = pl[f]
            elif cat.get(f) not in (None, ""):
                out[f] = cat[f]

        out["Код ТНВЭД"] = cat.get("Код ТНВЭД") or cu.get("Код ТНВЭД") or None
        out["ЕД.измер"] = cat.get("ЕД.измер") or cu.get("ЕД.измер") or None
        if not out.get("страна происхождения"):
            if cu.get("страна происхождения"):
                out["страна происхождения"] = cu["страна происхождения"]

        if not out.get("Код ТНВЭД"):
            missing_tnved.append(art)

        records.append({col: out.get(col) for col in TEMPLATE_COLUMNS})

    return records, missing_tnved


def write_output(records: list[dict], missing_tnved: list[str], output_path: str | Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Каталог"
    ws.append(TEMPLATE_COLUMNS)
    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(bold=True)
    ws.freeze_panes = "A2"

    for rec in records:
        ws.append([rec.get(col) for col in TEMPLATE_COLUMNS])

    for i, col in enumerate(TEMPLATE_COLUMNS, start=1):
        width = max(12, min(40, len(col) + 4))
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width

    if missing_tnved:
        ws2 = wb.create_sheet("Проверить вручную")
        ws2.append(["артикул", "Причина"])
        for cell in ws2[1]:
            cell.font = openpyxl.styles.Font(bold=True)
        for art in missing_tnved:
            ws2.append([art, "нет кода ТН ВЭД ни в каталоге, ни в реестре ДТ"])
        ws2.column_dimensions["A"].width = 20
        ws2.column_dimensions["B"].width = 45

    wb.save(str(output_path))
