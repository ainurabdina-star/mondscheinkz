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

from . import extract, web_enrich


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
    parser.add_argument(
        "--enrich-web",
        action="store_true",
        help="дополнить пустые поля (модель, Наименование, Краткое техническое "
        "описание, ВЕС нетто) данными с bosch-professional.com для позиций, "
        "где их нет ни в одном из Excel-источников. Никогда не перезаписывает "
        "уже имеющиеся значения. Проверьте правила использования сайта перед "
        "регулярным запуском.",
    )
    parser.add_argument(
        "--enrich-delay",
        type=float,
        default=0.5,
        help="пауза между запросами к сайту в секундах (по умолчанию 0.5)",
    )
    parser.add_argument(
        "--enrich-limit",
        type=int,
        default=None,
        help="ограничить число позиций, которые будут проверены на сайте "
        "(удобно для пробного запуска)",
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

    if args.enrich_web:
        print("\n[bosch-professional.com] проверяю позиции с пустыми полями…")

        def _progress(article, scraped):
            status = "найдено" if scraped else "не найдено / не подтверждено"
            print(f"  {article}: {status}")

        report = web_enrich.enrich_records(
            records, delay=args.enrich_delay, limit=args.enrich_limit, on_progress=_progress
        )
        print(
            f"[bosch-professional.com] проверено {report['checked']}, "
            f"дополнено {report['filled']}, не найдено {report['not_found']}, "
            f"без новых данных {report['no_new_data']}"
        )
        missing_tnved = [r["артикул"] for r in records if not r.get("Код ТНВЭД")]

    extract.write_output(records, missing_tnved, args.output)

    print(f"\nГотово: {len(records)} записей -> {args.output}")
    if missing_tnved:
        print(f"Без кода ТН ВЭД (см. лист 'Проверить вручную'): {len(missing_tnved)} шт.")

    return 0


if __name__ == "__main__":
    sys.exit(main())
