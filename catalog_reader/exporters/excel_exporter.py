from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd

from models import CatalogRow, ImportResult, RowStatus
from utils import ensure_dir, safe_filename


def export_import_result_to_excel(
    result: ImportResult,
    output_dir: str | Path = "output",
) -> Path:
    """
    Сохраняет результат обработки каталога в Excel.

    Главные листы:
    - db_ready: чистый формат для будущей базы
    - ready_import_long: один OE номер = одна строка
    - ready_import: широкий формат oe1, oe2, oe3

    db_ready содержит только ready строки и только главные поля:
    prefix | article | brand | vehicle_brand | oe_number
    """

    output_dir = ensure_dir(output_dir)

    filename = build_output_filename(result)
    output_path = output_dir / filename

    all_rows = result.rows
    max_oe_count = get_max_oe_count(all_rows)

    sheets = {
        "summary": summary_to_dataframe(result),

        # Главный будущий формат для базы
        "db_ready": rows_to_db_ready_dataframe(result.ready_rows),

        # Подробный длинный формат
        "ready_import_long": rows_to_long_dataframe(result.ready_rows),

        # Старый широкий формат для удобного просмотра
        "ready_import": rows_to_dataframe(result.ready_rows, max_oe_count),

        # Проверочные листы
        "needs_review": rows_to_dataframe(result.review_rows, max_oe_count, include_raw_text=True),
        "no_oe": rows_to_dataframe(result.no_oe_rows, max_oe_count, include_raw_text=True),
        "duplicates": rows_to_dataframe(result.duplicate_rows, max_oe_count, include_raw_text=True),
        "errors": rows_to_dataframe(result.error_rows, max_oe_count, include_raw_text=True),

        # Полная выгрузка
        "all_oe_long": rows_to_long_dataframe(all_rows, include_raw_text=True),
        "all_rows": rows_to_dataframe(all_rows, max_oe_count, include_raw_text=True),
    }

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_format = workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "border": 1,
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

        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            worksheet = writer.sheets[sheet_name]

            for col_num, column_name in enumerate(df.columns):
                worksheet.write(0, col_num, column_name, header_format)

            if len(df.columns) > 0:
                worksheet.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)

            worksheet.freeze_panes(1, 0)

            for col_num, column_name in enumerate(df.columns):
                width = guess_column_width(column_name, df[column_name] if column_name in df else None)

                if (
                    column_name.startswith("oe")
                    or column_name in {
                        "article",
                        "prefix",
                        "brand",
                        "vehicle_brand",
                        "product_group",
                        "catalog_name",
                    }
                ):
                    worksheet.set_column(col_num, col_num, width, text_format)
                elif column_name in {"raw_text", "reason", "description", "type_model"}:
                    worksheet.set_column(col_num, col_num, width, wrap_format)
                else:
                    worksheet.set_column(col_num, col_num, width)

    return output_path


def build_output_filename(result: ImportResult) -> str:
    brand = safe_filename(result.brand or "brand")
    prefix = safe_filename(result.prefix or "prefix")
    source_name = Path(result.source_file).stem if result.source_file else "catalog"
    source_name = safe_filename(source_name)

    return f"{brand}_{prefix}_{source_name}_parsed.xlsx"


def summary_to_dataframe(result: ImportResult) -> pd.DataFrame:
    summary = result.summary()

    rows = [
        {"metric": "source_file", "value": result.source_file},
        {"metric": "brand", "value": result.brand},
        {"metric": "prefix", "value": result.prefix},
        {"metric": "total", "value": summary["total"]},
        {"metric": "ready", "value": summary["ready"]},
        {"metric": "needs_review", "value": summary["needs_review"]},
        {"metric": "no_oe", "value": summary["no_oe"]},
        {"metric": "duplicates", "value": summary["duplicates"]},
        {"metric": "errors", "value": summary["errors"]},
    ]

    return pd.DataFrame(rows)


def rows_to_db_ready_dataframe(rows: Iterable[CatalogRow]) -> pd.DataFrame:
    """
    Самый чистый формат для будущей базы.

    Один OE номер = одна строка.
    Только ready-данные.
    """

    data = []

    for row in rows:
        for oe_number in row.oe_numbers:
            data.append(
                {
                    "prefix": row.prefix,
                    "article": row.article,
                    "brand": row.brand,
                    "vehicle_brand": row.vehicle_brand,
                    "oe_number": oe_number,
                    "description": row.description,
                    "product_group": row.product_group,
                    "source_file": row.source_file,
                    "page": row.page,
                }
            )

    columns = [
        "prefix",
        "article",
        "brand",
        "vehicle_brand",
        "oe_number",
        "description",
        "product_group",
        "source_file",
        "page",
    ]

    df = pd.DataFrame(data)

    if df.empty:
        return pd.DataFrame(columns=columns)

    return df[columns].fillna("").astype(str)


def rows_to_dataframe(
    rows: Iterable[CatalogRow],
    max_oe_count: int,
    include_raw_text: bool = False,
) -> pd.DataFrame:
    """
    Широкий формат:
    одна строка = один article,
    OE номера идут колонками oe1, oe2, oe3...
    """

    data = []

    for row in rows:
        item = {
            "status": row.status.value if isinstance(row.status, RowStatus) else str(row.status),
            "prefix": row.prefix,
            "article": row.article,
            "brand": row.brand,
            "vehicle_brand": row.vehicle_brand,
            "product_group": row.product_group,
            "catalog_name": row.catalog_name,
            "description": row.description,
            "type_model": row.type_model,
            "page": row.page,
            "reason": row.reason,
            "source_file": row.source_file,
        }

        for index in range(max_oe_count):
            column_name = f"oe{index + 1}"
            item[column_name] = row.oe_numbers[index] if index < len(row.oe_numbers) else ""

        if include_raw_text:
            item["raw_text"] = row.raw_text

        data.append(item)

    oe_columns = [f"oe{i + 1}" for i in range(max_oe_count)]

    columns = [
        "status",
        "prefix",
        "article",
        "brand",
        "vehicle_brand",
        "product_group",
        *oe_columns,
        "description",
        "type_model",
        "page",
        "reason",
        "catalog_name",
        "source_file",
    ]

    if include_raw_text:
        columns.append("raw_text")

    df = pd.DataFrame(data)

    if df.empty:
        return pd.DataFrame(columns=columns)

    for column in columns:
        if column not in df.columns:
            df[column] = ""

    return df[columns].fillna("").astype(str)


def rows_to_long_dataframe(
    rows: Iterable[CatalogRow],
    include_raw_text: bool = False,
) -> pd.DataFrame:
    """
    Длинный формат:
    один OE номер = одна строка.
    """

    data = []

    for row in rows:
        status = row.status.value if isinstance(row.status, RowStatus) else str(row.status)

        if row.oe_numbers:
            for index, oe_number in enumerate(row.oe_numbers, start=1):
                item = {
                    "status": status,
                    "prefix": row.prefix,
                    "article": row.article,
                    "brand": row.brand,
                    "vehicle_brand": row.vehicle_brand,
                    "product_group": row.product_group,
                    "oe_number": oe_number,
                    "oe_order": index,
                    "description": row.description,
                    "type_model": row.type_model,
                    "page": row.page,
                    "reason": row.reason,
                    "catalog_name": row.catalog_name,
                    "source_file": row.source_file,
                }

                if include_raw_text:
                    item["raw_text"] = row.raw_text

                data.append(item)
        else:
            item = {
                "status": status,
                "prefix": row.prefix,
                "article": row.article,
                "brand": row.brand,
                "vehicle_brand": row.vehicle_brand,
                "product_group": row.product_group,
                "oe_number": "",
                "oe_order": "",
                "description": row.description,
                "type_model": row.type_model,
                "page": row.page,
                "reason": row.reason,
                "catalog_name": row.catalog_name,
                "source_file": row.source_file,
            }

            if include_raw_text:
                item["raw_text"] = row.raw_text

            data.append(item)

    columns = [
        "status",
        "prefix",
        "article",
        "brand",
        "vehicle_brand",
        "product_group",
        "oe_number",
        "oe_order",
        "description",
        "type_model",
        "page",
        "reason",
        "catalog_name",
        "source_file",
    ]

    if include_raw_text:
        columns.append("raw_text")

    df = pd.DataFrame(data)

    if df.empty:
        return pd.DataFrame(columns=columns)

    for column in columns:
        if column not in df.columns:
            df[column] = ""

    return df[columns].fillna("").astype(str)


def get_max_oe_count(rows: Iterable[CatalogRow]) -> int:
    max_count = 0

    for row in rows:
        max_count = max(max_count, len(row.oe_numbers))

    return max(max_count, 3)


def guess_column_width(column_name: str, series=None) -> int:
    default_widths = {
        "status": 16,
        "prefix": 10,
        "article": 18,
        "brand": 18,
        "vehicle_brand": 18,
        "product_group": 34,
        "catalog_name": 45,
        "description": 34,
        "type_model": 30,
        "page": 10,
        "reason": 42,
        "source_file": 45,
        "raw_text": 80,
        "metric": 20,
        "value": 50,
        "oe_number": 24,
        "oe_order": 10,
        "output_file": 45,
    }

    if column_name.startswith("oe"):
        return 18

    if column_name in default_widths:
        return default_widths[column_name]

    if series is not None and len(series) > 0:
        max_length = max(len(str(x)) for x in series.head(100))
        return min(max(max_length + 2, 12), 50)

    return 15