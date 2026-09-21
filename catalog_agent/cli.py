"""CLI: собрать каталог установленного образца из прайс-листов, реестра ДТ и
уже существующего каталога.

Пример:
    python -m catalog_agent.cli \
        --catalog "Каталог.xlsx" \
        --pricelist "Pricelist_KZ.xlsx" \
        --customs "PT_20260708.xlsx" \
        --output "Каталог_обновлённый.xlsx"
"""
from __future__ import annotations

import argparse
import sys

from . import extract


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--catalog", help="существующий каталог в целевом формате (.xlsx)")
    parser.add_argument(
        "--pricelist", action="append", default=[], help="прайс-лист поставщика (.xlsx), можно указывать несколько раз"
    )
    parser.add_argument(
        "--customs", action="append", default=[], help="реестр таможенных деклараций (.xlsx), можно указывать несколько раз"
    )
    parser.add_argument("--output", required=True, help="куда сохранить результат (.xlsx)")
    parser.add_argument(
        "--articles",
        help="текстовый файл со списком артикулов (по одному на строку) — "
        "ограничить вывод только этими позициями; по умолчанию берутся все "
        "артикулы из каталога и прайс-листов",
    )
    args = parser.parse_args(argv)

    catalog: dict = {}
    if args.catalog:
        stats = extract.SourceStats(path=args.catalog)
        catalog = extract.read_catalog(args.catalog, stats)
        print(f"[каталог] {args.catalog}: листы {stats.sheets_used}, артикулов {stats.unique_articles}")

    pricelist: dict = {}
    for p in args.pricelist:
        stats = extract.SourceStats(path=p)
        part = extract.read_pricelist(p, stats)
        pricelist.update(part)
        print(
            f"[прайс-лист] {p}: использованы листы {stats.sheets_used}, "
            f"пропущены {stats.sheets_skipped}, артикулов {stats.unique_articles}"
        )

    customs: dict = {}
    for c in args.customs:
        stats = extract.SourceStats(path=c)
        part = extract.read_customs_register(c, stats)
        customs.update(part)
        print(f"[реестр ДТ] {c}: листы {stats.sheets_used}, артикулов {stats.unique_articles}")

    articles = None
    if args.articles:
        with open(args.articles, encoding="utf-8") as f:
            articles = [line.strip() for line in f if line.strip()]

    records, missing_tnved = extract.build_records(catalog, pricelist, customs, articles)
    extract.write_output(records, missing_tnved, args.output)

    print(f"\nГотово: {len(records)} записей -> {args.output}")
    if missing_tnved:
        print(f"Без кода ТН ВЭД (см. лист 'Проверить вручную'): {len(missing_tnved)} шт.")

    return 0


if __name__ == "__main__":
    sys.exit(main())
