from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List

import pandas as pd

from models import CatalogRow, ImportResult, RowStatus
from utils import ensure_dir, safe_filename


SITE_IMPORT_COLUMNS = [
    "brand",
    "code",
    "brand_from",
    "code_from",
    "load_image",
    "load_characteristics",
    "load_cross",
    "load_applicability",
]

@dataclass
class SiteImportExportResult:
    ready_path: Path
    review_path: Path | None
    ready_count: int
    skipped_count: int


def export_site_import_from_result(
    result: ImportResult,
    output_dir: str | Path = "output/site_import",
) -> SiteImportExportResult:
    """
    Формирует файл для импорта на сайт.

    Главный формат сайта:

    brand | code | brand_from | code_from

    Где:
    - brand = поставщик / производитель детали
    - code = артикул поставщика
    - brand_from = производитель оригинального номера / vehicle_brand
    - code_from = OE номер

    В сайт-файл попадают только готовые строки:
    - status = ready
    - есть article
    - есть oe_number
    - есть vehicle_brand

    Все строки без vehicle_brand или OE уходят в review-файл.
    """

    output_dir = ensure_dir(output_dir)

    ready_rows: List[dict] = []
    review_rows: List[dict] = []

    for row in result.rows:
        if row.status != RowStatus.READY:
            add_review_row(
                review_rows=review_rows,
                row=row,
                reason=f"Строка не ready: {row.status.value if isinstance(row.status, RowStatus) else row.status}",
            )
            continue

        if not row.article:
            add_review_row(
                review_rows=review_rows,
                row=row,
                reason="Нет article / code",
            )
            continue

        if not row.oe_numbers:
            add_review_row(
                review_rows=review_rows,
                row=row,
                reason="Нет OE номеров / code_from",
            )
            continue

        if not row.vehicle_brand:
            add_review_row(
                review_rows=review_rows,
                row=row,
                reason="Нет vehicle_brand / brand_from",
            )
            continue

        for oe_number in row.oe_numbers:
            if not oe_number:
                continue

            ready_rows.append(
                {
                    "brand": format_site_brand(row.brand),
                    "code": str(row.article).strip(),
                    "brand_from": format_site_brand_from(row.vehicle_brand),
                    "code_from": str(oe_number).strip(),
                    "load_image": 0,
                    "load_characteristics": 0,
                    "load_cross": 1,
                    "load_applicability": 0,
                }
            )

    ready_df = pd.DataFrame(ready_rows, columns=SITE_IMPORT_COLUMNS)
    ready_df = ready_df.drop_duplicates().fillna("").astype(str)

    review_df = pd.DataFrame(review_rows).fillna("").astype(str)

    base_name = build_site_import_base_name(result)

    ready_path = output_dir / f"{base_name}_site_import.xlsx"
    review_path = None

    write_site_import_excel(
        df=ready_df,
        output_path=ready_path,
    )

    if not review_df.empty:
        review_path = output_dir / f"{base_name}_site_import_review.xlsx"
        write_review_excel(
            df=review_df,
            output_path=review_path,
        )

    return SiteImportExportResult(
        ready_path=ready_path,
        review_path=review_path,
        ready_count=len(ready_df),
        skipped_count=len(review_df),
    )


def add_review_row(
    review_rows: List[dict],
    row: CatalogRow,
    reason: str,
) -> None:
    """
    Добавляет строку в review-файл.
    """

    review_rows.append(
        {
            "status": row.status.value if isinstance(row.status, RowStatus) else str(row.status),
            "prefix": row.prefix,
            "article": row.article,
            "brand": row.brand,
            "vehicle_brand": row.vehicle_brand,
            "oe_numbers": ", ".join(row.oe_numbers),
            "description": row.description,
            "product_group": row.product_group,
            "page": row.page,
            "reason": reason,
            "source_file": row.source_file,
            "raw_text": row.raw_text,
        }
    )


def write_site_import_excel(
    df: pd.DataFrame,
    output_path: Path,
) -> None:
    """
    Пишет Excel для сайта.

    Важно:
    - только 4 колонки
    - все значения как текст
    - ведущие нули в OE сохраняются
    """

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df.to_excel(
            writer,
            sheet_name="import",
            index=False,
        )

        workbook = writer.book
        worksheet = writer.sheets["import"]

        header_format = workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "valign": "top",
            }
        )

        text_format = workbook.add_format(
            {
                "num_format": "@",
                "valign": "top",
            }
        )

        for col_num, column_name in enumerate(SITE_IMPORT_COLUMNS):
            worksheet.write(0, col_num, column_name, header_format)
            worksheet.set_column(col_num, col_num, 22, text_format)

        if len(df.columns) > 0:
            worksheet.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)

        worksheet.freeze_panes(1, 0)


def write_review_excel(
    df: pd.DataFrame,
    output_path: Path,
) -> None:
    """
    Пишет отдельный review-файл.

    Он НЕ предназначен для сайта.
    Это файл для проверки строк, которые нельзя безопасно отправлять на сайт.
    """

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df.to_excel(
            writer,
            sheet_name="review",
            index=False,
        )

        workbook = writer.book
        worksheet = writer.sheets["review"]

        header_format = workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "valign": "top",
            }
        )

        text_format = workbook.add_format(
            {
                "num_format": "@",
                "valign": "top",
            }
        )

        wrap_format = workbook.add_format(
            {
                "text_wrap": True,
                "valign": "top",
            }
        )

        for col_num, column_name in enumerate(df.columns):
            worksheet.write(0, col_num, column_name, header_format)

            if column_name in {"raw_text", "reason", "description"}:
                worksheet.set_column(col_num, col_num, 45, wrap_format)
            else:
                worksheet.set_column(col_num, col_num, 22, text_format)

        if len(df.columns) > 0:
            worksheet.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)

        worksheet.freeze_panes(1, 0)


def build_site_import_base_name(result: ImportResult) -> str:
    brand = safe_filename(result.brand or "brand")
    prefix = safe_filename(result.prefix or "prefix")
    source_name = Path(result.source_file).stem if result.source_file else "catalog"
    source_name = safe_filename(source_name)

    return f"{brand}_{prefix}_{source_name}"


def format_site_brand(value: str) -> str:
    """
    Форматирует производителя детали для сайта.

    semlastik -> Semlastik
    SEMLASTIK -> Semlastik
    3G -> 3G
    BPW -> BPW
    """

    value = str(value or "").strip()

    if not value:
        return ""

    if any(char.isdigit() for char in value):
        return value.upper()

    if len(value) <= 3:
        return value.upper()

    return value.lower().title()


def format_site_brand_from(value: str) -> str:
    """
    Форматирует производителя оригинального номера.

    Для brand_from обычно лучше верхний регистр:
    MERCEDES, DAF, IVECO, DENNIS.
    """

    value = str(value or "").strip()

    if not value:
        return ""

    return value.upper()