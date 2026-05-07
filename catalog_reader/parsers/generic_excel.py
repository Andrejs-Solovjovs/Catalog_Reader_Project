from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd

from models import CatalogRow, ImportResult, RowStatus
from utils import (
    clean_text,
    is_probably_oe_number,
    normalize_article,
    normalize_brand,
    normalize_oe,
    normalize_prefix,
    unique_keep_order,
)
from validators.rules import validate_rows


HEADER_SCAN_ROWS = 60


COMMON_VEHICLE_BRANDS = [
    "DAF",
    "DENNIS",
    "SETRA",
    "NEOPLAN",
    "IKARUS",
    "VANHOOL",
    "VAN HOOL",
    "BOVA",
    "SOLARIS",
    "MAN",
    "MERCEDES",
    "MERCEDES BENZ",
    "SCANIA",
    "VOLVO",
    "RENAULT",
    "IVECO",
    "FORD",
    "BMC",
    "TEMSA",
    "OTOKAR",
]


@dataclass
class ExcelColumnMap:
    sheet_name: str
    header_row_index: int

    article_cols: List[int]
    oe_cols: List[int]
    description_cols: List[int]
    vehicle_brand_cols: List[int]
    product_group_cols: List[int]


def parse_generic_excel_catalog(
    file_path: str | Path,
    brand: str,
    prefix: str,
) -> ImportResult:
    """
    Универсальный Excel/CSV парсер.

    Подходит для каталогов, где данные уже лежат таблицей.

    Ожидаемый смысл колонок:
    - article / part no / supplier article / product code
    - OE / OEM / reference / original no
    - description / name
    - vehicle_brand / make / application brand
    - product_group / category / group

    Главное:
    на выходе все равно получаем общий формат CatalogRow.
    """

    file_path = Path(file_path)
    brand = normalize_brand(brand)
    prefix = normalize_prefix(prefix)
    catalog_name = file_path.stem

    rows: List[CatalogRow] = []

    try:
        tables = read_excel_or_csv_tables(file_path)
    except Exception as exc:
        return ImportResult(
            source_file=file_path.name,
            brand=brand,
            prefix=prefix,
            rows=[
                CatalogRow(
                    prefix=prefix,
                    brand=brand,
                    catalog_name=catalog_name,
                    source_file=file_path.name,
                    status=RowStatus.ERROR,
                    reason=f"Не удалось прочитать Excel/CSV файл: {exc}",
                )
            ],
        )

    if not tables:
        return ImportResult(
            source_file=file_path.name,
            brand=brand,
            prefix=prefix,
            rows=[
                CatalogRow(
                    prefix=prefix,
                    brand=brand,
                    catalog_name=catalog_name,
                    source_file=file_path.name,
                    status=RowStatus.ERROR,
                    reason="В Excel/CSV файле не найдены листы или таблицы",
                )
            ],
        )

    for sheet_name, df in tables:
        sheet_rows = parse_excel_sheet(
            df=df,
            sheet_name=sheet_name,
            file_path=file_path,
            brand=brand,
            prefix=prefix,
            catalog_name=catalog_name,
        )

        rows.extend(sheet_rows)

    if not rows:
        rows.append(
            CatalogRow(
                prefix=prefix,
                brand=brand,
                catalog_name=catalog_name,
                source_file=file_path.name,
                status=RowStatus.ERROR,
                reason="Парсер не нашел строк article + OE. Возможно, нужен отдельный шаблон для этого Excel.",
            )
        )

    rows = validate_rows(rows)

    return ImportResult(
        source_file=file_path.name,
        brand=brand,
        prefix=prefix,
        rows=rows,
    )


def read_excel_or_csv_tables(file_path: Path) -> List[Tuple[str, pd.DataFrame]]:
    suffix = file_path.suffix.lower()

    if suffix == ".csv":
        df = pd.read_csv(
            file_path,
            dtype=str,
            encoding_errors="ignore",
        )

        return [("CSV", df)]

    excel_file = pd.ExcelFile(file_path)

    tables: List[Tuple[str, pd.DataFrame]] = []

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            dtype=str,
            header=None,
        )

        tables.append((sheet_name, df))

    return tables


def parse_excel_sheet(
    df: pd.DataFrame,
    sheet_name: str,
    file_path: Path,
    brand: str,
    prefix: str,
    catalog_name: str,
) -> List[CatalogRow]:
    if df is None or df.empty:
        return []

    df = df.fillna("").astype(str)

    column_map = find_excel_column_map(
        df=df,
        sheet_name=sheet_name,
    )

    if not column_map:
        return [
            CatalogRow(
                prefix=prefix,
                brand=brand,
                catalog_name=catalog_name,
                product_group=sheet_name,
                source_file=file_path.name,
                status=RowStatus.ERROR,
                reason=f"На листе '{sheet_name}' не найдены колонки article и OE",
                raw_text=sheet_preview_text(df),
            )
        ]

    return parse_rows_with_column_map(
        df=df,
        column_map=column_map,
        sheet_name=sheet_name,
        file_path=file_path,
        brand=brand,
        prefix=prefix,
        catalog_name=catalog_name,
    )


def find_excel_column_map(
    df: pd.DataFrame,
    sheet_name: str,
) -> Optional[ExcelColumnMap]:
    """
    Ищет строку заголовков.

    Нужен минимум:
    - одна колонка article
    - одна колонка OE
    """

    scan_rows = min(len(df), HEADER_SCAN_ROWS)

    best_map: Optional[ExcelColumnMap] = None
    best_score = 0

    for row_index in range(scan_rows):
        row_values = [
            str(value or "").strip()
            for value in df.iloc[row_index].tolist()
        ]

        article_cols: List[int] = []
        oe_cols: List[int] = []
        description_cols: List[int] = []
        vehicle_brand_cols: List[int] = []
        product_group_cols: List[int] = []

        for col_index, value in enumerate(row_values):
            header_type = classify_header(value)

            if header_type == "article":
                article_cols.append(col_index)
            elif header_type == "oe":
                oe_cols.append(col_index)
            elif header_type == "description":
                description_cols.append(col_index)
            elif header_type == "vehicle_brand":
                vehicle_brand_cols.append(col_index)
            elif header_type == "product_group":
                product_group_cols.append(col_index)

        if not article_cols or not oe_cols:
            continue

        score = (
            len(article_cols) * 10
            + len(oe_cols) * 20
            + len(description_cols) * 5
            + len(vehicle_brand_cols) * 5
            + len(product_group_cols) * 5
        )

        if score > best_score:
            best_score = score
            best_map = ExcelColumnMap(
                sheet_name=sheet_name,
                header_row_index=row_index,
                article_cols=article_cols,
                oe_cols=oe_cols,
                description_cols=description_cols,
                vehicle_brand_cols=vehicle_brand_cols,
                product_group_cols=product_group_cols,
            )

    return best_map


def classify_header(value: str) -> str:
    """
    Определяет смысл колонки по названию.

    Возвращает:
    - article
    - oe
    - description
    - vehicle_brand
    - product_group
    - ""
    """

    original = str(value or "").strip()
    header = normalize_header(original)

    if not header:
        return ""

    article_markers = [
        "ARTICLE",
        "ARTICLE NO",
        "ARTICLE NUMBER",
        "ARTIKEL",
        "ARTICUL",
        "SUPPLIER ARTICLE",
        "SUPPLIER CODE",
        "ITEM NO",
        "ITEM NUMBER",
        "PART NO",
        "PART NUMBER",
        "PRODUCT CODE",
        "PRODUCT NO",
        "SEM NO",
        "SEM NUMBER",
        "MANUFACTURER PART",
        "MANUFACTURER CODE",
        "КОД ТОВАРА",
        "АРТИКУЛ",
        "НОМЕР ДЕТАЛИ",
    ]

    oe_markers = [
        "OE",
        "OE NO",
        "OE NUMBER",
        "OEM",
        "OEM NO",
        "OEM NUMBER",
        "REF",
        "REF NO",
        "REF NUMBER",
        "REFERENCE",
        "REFERENCE NO",
        "ORIGINAL",
        "ORIGINAL NO",
        "ORIGINAL NUMBER",
        "GENUINE NO",
        "CROSS",
        "CROSS REFERENCE",
        "CROSS NO",
        "ОЕ",
        "OEM НОМЕР",
        "ОРИГИНАЛЬНЫЙ НОМЕР",
        "КРОСС",
        "КРОСС НОМЕР",
    ]

    description_markers = [
        "DESCRIPTION",
        "DESC",
        "NAME",
        "PRODUCT NAME",
        "ITEM NAME",
        "PART NAME",
        "НАИМЕНОВАНИЕ",
        "ОПИСАНИЕ",
        "НАЗВАНИЕ",
    ]

    vehicle_brand_markers = [
        "VEHICLE BRAND",
        "VEHICLE MAKE",
        "MAKE",
        "CAR BRAND",
        "TRUCK BRAND",
        "APPLICATION BRAND",
        "VEHICLE MANUFACTURER",
        "OE BRAND",
        "MARQUE",
        "МАРКА АВТО",
        "МАРКА",
        "ПРИМЕНЯЕМОСТЬ",
        "ПРОИЗВОДИТЕЛЬ ТС",
    ]

    product_group_markers = [
        "PRODUCT GROUP",
        "GROUP",
        "CATEGORY",
        "CATALOG GROUP",
        "ГРУППА",
        "КАТЕГОРИЯ",
        "РАЗДЕЛ",
    ]

    if header_matches(header, article_markers):
        return "article"

    if header_matches(header, oe_markers):
        return "oe"

    if header_matches(header, description_markers):
        return "description"

    if header_matches(header, vehicle_brand_markers):
        return "vehicle_brand"

    if header_matches(header, product_group_markers):
        return "product_group"

    return ""


def header_matches(header: str, markers: Iterable[str]) -> bool:
    for marker in markers:
        marker = normalize_header(marker)

        if header == marker:
            return True

        if marker and marker in header:
            return True

    return False


def parse_rows_with_column_map(
    df: pd.DataFrame,
    column_map: ExcelColumnMap,
    sheet_name: str,
    file_path: Path,
    brand: str,
    prefix: str,
    catalog_name: str,
) -> List[CatalogRow]:
    rows: List[CatalogRow] = []

    last_article = ""
    last_description = ""
    last_vehicle_brand = ""
    last_product_group = ""

    start_row = column_map.header_row_index + 1

    for row_index in range(start_row, len(df)):
        row_values = [
            str(value or "").strip()
            for value in df.iloc[row_index].tolist()
        ]

        if is_empty_row(row_values):
            continue

        raw_text = " | ".join(value for value in row_values if value)

        article = first_non_empty(row_values, column_map.article_cols)
        description = first_non_empty(row_values, column_map.description_cols)
        vehicle_brand = first_non_empty(row_values, column_map.vehicle_brand_cols)
        product_group = first_non_empty(row_values, column_map.product_group_cols)

        oe_numbers = collect_oe_numbers_from_row(
            row_values=row_values,
            oe_cols=column_map.oe_cols,
        )

        if article:
            last_article = article

        if description:
            last_description = description

        if vehicle_brand:
            last_vehicle_brand = vehicle_brand

        if product_group:
            last_product_group = product_group

        # Если Excel сделан так, что article указан один раз,
        # а OE номера идут ниже отдельными строками.
        if not article and oe_numbers and last_article:
            article = last_article

            if not description:
                description = last_description

            if not vehicle_brand:
                vehicle_brand = last_vehicle_brand

            if not product_group:
                product_group = last_product_group

        if not article and not oe_numbers:
            continue

        if not vehicle_brand:
            vehicle_brand = guess_vehicle_brand_from_text(
                text=" ".join([sheet_name, raw_text, catalog_name])
            )

        if not product_group:
            product_group = sheet_name

        row = CatalogRow(
            prefix=prefix,
            article=normalize_article(article),
            brand=brand,
            oe_numbers=oe_numbers,
            description=clean_text(description),
            type_model="",
            catalog_name=catalog_name,
            vehicle_brand=clean_text(vehicle_brand).upper(),
            product_group=clean_text(product_group),
            page=row_index + 1,
            status=RowStatus.NEEDS_REVIEW,
            reason="pending validation",
            raw_text=f"sheet={sheet_name}; row={row_index + 1}; {raw_text}",
            source_file=file_path.name,
        )

        rows.append(row)

    return rows


def collect_oe_numbers_from_row(
    row_values: List[str],
    oe_cols: List[int],
) -> List[str]:
    oe_numbers: List[str] = []

    for col_index in oe_cols:
        if col_index >= len(row_values):
            continue

        value = row_values[col_index]

        for oe in split_oe_cell(value):
            if is_probably_oe_number(oe):
                oe_numbers.append(oe)

    return unique_keep_order(oe_numbers)


def split_oe_cell(value: str) -> List[str]:
    """
    Разбивает ячейку с OE номерами.

    Примеры:
    "123; 456; 789" -> ["123", "456", "789"]
    "123\\n456" -> ["123", "456"]
    "SACHS 318890" -> ["SACHS 318890"]
    "654104/3" -> ["654104/3"]
    """

    value = normalize_oe(value)

    if not value:
        return []

    value = clean_oe_cell(value)

    parts = re.split(r"[\n\r,;|]+", value)

    candidates: List[str] = []

    for part in parts:
        part = normalize_oe(part)

        if not part:
            continue

        tokens = part.split()

        if should_split_oe_tokens(tokens):
            candidates.extend(tokens)
        else:
            candidates.append(part)

    cleaned: List[str] = []

    for candidate in candidates:
        candidate = normalize_oe(candidate)

        if not candidate:
            continue

        if looks_like_component_reference(candidate):
            continue

        if is_probably_oe_number(candidate):
            cleaned.append(candidate)

    return unique_keep_order(cleaned)


def clean_oe_cell(value: str) -> str:
    value = normalize_oe(value)

    # Убираем типовые подписи.
    value = re.sub(r"^(OE|OEM|REF|REFERENCE|ORIGINAL)\s*[:：\-]?\s*", "", value, flags=re.IGNORECASE)

    # WVA обычно технический номер колодки, не OE.
    value = re.sub(r"^WVA\s*[:：]?\s*\d+\s*", "", value, flags=re.IGNORECASE)

    # TYPE / TIP / TİP — это применяемость, а не OE.
    value = re.sub(r"\b(TYPE|TIP|TİP)\b.*$", "", value, flags=re.IGNORECASE)

    return normalize_oe(value)


def should_split_oe_tokens(tokens: List[str]) -> bool:
    """
    Разбиваем строку по пробелам только в безопасных случаях.

    Например:
    "652516 656274" можно разбить.
    "SACHS 318890" нельзя разбивать.
    "800 1/2" нельзя разбивать.
    """

    if len(tokens) <= 1:
        return False

    tokens = [normalize_oe(token) for token in tokens if normalize_oe(token)]

    if len(tokens) <= 1:
        return False

    # Не разбиваем "800 1/2", "685 2/3".
    if any("/" in token and len(token.replace("/", "")) <= 3 for token in tokens):
        return False

    # Не разбиваем "SACHS 318890".
    if any(not any(char.isdigit() for char in token) for token in tokens):
        return False

    for token in tokens:
        compact = token.replace(" ", "")

        if len(compact) < 4:
            return False

        if not re.fullmatch(r"\d+[A-Za-z]?", compact):
            return False

    return True


def looks_like_component_reference(value: str) -> bool:
    """
    Отсекаем ссылки на компоненты комплекта:
    1-8631
    2-8632
    1-
    2-
    """

    value = normalize_oe(value)
    compact = value.replace(" ", "")

    if re.fullmatch(r"\d+-+\d+", compact):
        return True

    if re.fullmatch(r"\d+-+", compact):
        return True

    if "=>" in compact:
        return True

    return False


def first_non_empty(
    row_values: List[str],
    columns: List[int],
) -> str:
    for col_index in columns:
        if col_index >= len(row_values):
            continue

        value = str(row_values[col_index] or "").strip()

        if value:
            return value

    return ""


def guess_vehicle_brand_from_text(text: str) -> str:
    normalized = normalize_header(text)

    candidates = sorted(COMMON_VEHICLE_BRANDS, key=len, reverse=True)

    for candidate in candidates:
        candidate_normalized = normalize_header(candidate)

        if candidate_normalized and candidate_normalized in normalized:
            return candidate.replace(" ", "")

    return ""


def is_empty_row(values: List[str]) -> bool:
    return not any(str(value or "").strip() for value in values)


def sheet_preview_text(df: pd.DataFrame, max_rows: int = 20, max_cols: int = 15) -> str:
    if df is None or df.empty:
        return ""

    preview = df.iloc[:max_rows, :max_cols].fillna("").astype(str)

    lines = []

    for _, row in preview.iterrows():
        values = [str(value).strip() for value in row.tolist() if str(value).strip()]

        if values:
            lines.append(" | ".join(values))

    return clean_text("\n".join(lines))


def normalize_header(value: str) -> str:
    value = str(value or "").upper()
    value = value.replace("&", " AND ")
    value = value.replace("_", " ")
    value = value.replace("-", " ")
    value = re.sub(r"[^A-ZА-ЯЁ0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()