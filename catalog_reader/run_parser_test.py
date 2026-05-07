from __future__ import annotations

import argparse
from pathlib import Path

from exporters.excel_exporter import export_import_result_to_excel
from parsers.semlastik_pdf import parse_semlastik_pdf
from utils import ensure_dir


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Тестовый запуск парсера Semlastik PDF без интерфейса Streamlit."
    )

    parser.add_argument(
        "--input",
        required=True,
        help="Путь к PDF каталогу.",
    )

    parser.add_argument(
        "--brand",
        default="semlastik",
        help="Бренд поставщика.",
    )

    parser.add_argument(
        "--prefix",
        default="SEM",
        help="Наш внутренний prefix.",
    )

    parser.add_argument(
        "--start-page",
        type=int,
        default=16,
        help="С какой страницы PDF начинать обработку.",
    )

    parser.add_argument(
        "--max-pages",
        type=int,
        default=5,
        help="Сколько страниц обработать для теста.",
    )

    parser.add_argument(
        "--output-dir",
        default="output",
        help="Папка для результата.",
    )

    args = parser.parse_args()

    input_path = Path(args.input)

    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    ensure_dir(args.output_dir)

    print("Начинаю обработку")
    print(f"Файл: {input_path}")
    print(f"Бренд: {args.brand}")
    print(f"Prefix: {args.prefix}")
    print(f"Стартовая страница: {args.start_page}")
    print(f"Количество страниц: {args.max_pages}")
    print("-" * 60)

    result = parse_semlastik_pdf(
        file_path=input_path,
        brand=args.brand,
        prefix=args.prefix,
        start_page=args.start_page,
        max_pages=args.max_pages,
    )

    summary = result.summary()

    print("Результат:")
    print(f"Всего строк: {summary['total']}")
    print(f"Ready: {summary['ready']}")
    print(f"Needs review: {summary['needs_review']}")
    print(f"No OE: {summary['no_oe']}")
    print(f"Duplicates: {summary['duplicates']}")
    print(f"Errors: {summary['errors']}")
    print("-" * 60)

    output_path = export_import_result_to_excel(
        result=result,
        output_dir=args.output_dir,
    )

    print(f"Готово. Excel сохранен: {output_path}")


if __name__ == "__main__":
    main()

    #python run_parser_test.py --input "uploads/DAF_PRODUCT_CATALOGUE_2026.pdf" --start-page 16 --max-pages 5