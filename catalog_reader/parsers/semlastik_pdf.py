from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

import fitz  # PyMuPDF

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


SEMLASTIK_VEHICLE_BRANDS = [
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

PRODUCT_GROUP_NAMES = [
    "Engine and Transmission Group",
    "Steering Group",
    "Suspension Group",
    "Cab Group",
    "Axle / Differential ve Shaft Group",
    "Axle / Differential and Shaft Group",
    "Clutch and Brake Group",
    "Fifth Wheel Group",
    "External Parts Group",
]


@dataclass
class PdfWord:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    block_no: int
    line_no: int
    word_no: int

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass
class HeaderRow:
    y: float
    sem_x: float
    ref_x: float
    text: str


def parse_semlastik_pdf(
    file_path: str | Path,
    brand: str = "semlastik",
    prefix: str = "SEM",
    start_page: int = 1,
    max_pages: Optional[int] = None,
) -> ImportResult:
    """
    Парсер Semlastik PDF.

    В Semlastik:
    - SEM NO. = артикул поставщика
    - REF NO. = OE / оригинальные номера

    brand = поставщик, например SEMLASTIK
    vehicle_brand = DAF / DENNIS / SETRA / NEOPLAN и т.д.
    """

    file_path = Path(file_path)
    brand = normalize_brand(brand)
    prefix = normalize_prefix(prefix)
    catalog_name = file_path.stem

    rows: List[CatalogRow] = []

    try:
        doc = fitz.open(file_path)
    except Exception as exc:
        error_row = CatalogRow(
            prefix=prefix,
            brand=brand,
            catalog_name=catalog_name,
            source_file=file_path.name,
            status=RowStatus.ERROR,
            reason=f"Не удалось открыть PDF: {exc}",
        )

        return ImportResult(
            source_file=file_path.name,
            brand=brand,
            prefix=prefix,
            rows=[error_row],
        )

    try:
        total_pages = len(doc)
        end_page = total_pages

        if max_pages is not None:
            end_page = min(total_pages, start_page - 1 + max_pages)

        for page_index in range(start_page - 1, end_page):
            page_number = page_index + 1

            try:
                page = doc[page_index]
                page_rows = parse_page(
                    page=page,
                    page_number=page_number,
                    source_file=file_path.name,
                    catalog_name=catalog_name,
                    brand=brand,
                    prefix=prefix,
                )
                rows.extend(page_rows)

            except Exception as exc:
                rows.append(
                    CatalogRow(
                        prefix=prefix,
                        brand=brand,
                        catalog_name=catalog_name,
                        page=page_number,
                        source_file=file_path.name,
                        status=RowStatus.ERROR,
                        reason=f"Ошибка обработки страницы: {exc}",
                    )
                )

    finally:
        doc.close()

    rows = validate_rows(rows)

    return ImportResult(
        source_file=file_path.name,
        brand=brand,
        prefix=prefix,
        rows=rows,
    )


def parse_page(
    page,
    page_number: int,
    source_file: str,
    catalog_name: str,
    brand: str,
    prefix: str,
) -> List[CatalogRow]:
    words = extract_words(page)

    if not words:
        return []

    page_text = build_raw_text(words)

    # Не парсим KEY LIST / CROSS LIST как товарные страницы.
    if is_non_product_list_page(page_text):
        return []

    # Пропускаем страницы без товарных признаков.
    if not looks_like_product_page(page_text):
        return []

    vehicle_brand = extract_vehicle_brand(page_text)
    product_group = extract_product_group(page_text)

    header_rows = find_header_rows(words)

    if not header_rows:
        return [
            CatalogRow(
                prefix=prefix,
                brand=brand,
                catalog_name=catalog_name,
                vehicle_brand=vehicle_brand,
                product_group=product_group,
                page=page_number,
                source_file=source_file,
                raw_text=page_text[:3000],
                status=RowStatus.ERROR,
                reason="Похоже на товарную страницу, но не найдены колонки SEM NO. / REF NO.",
            )
        ]

    rows: List[CatalogRow] = []

    for index, header in enumerate(header_rows):
        block_top = header.y
        block_bottom = (
            header_rows[index + 1].y
            if index + 1 < len(header_rows)
            else page.rect.height - 55
        )

        block_words = [
            word for word in words
            if block_top <= word.y0 < block_bottom
        ]

        if not block_words:
            continue

        row = parse_product_block(
            block_words=block_words,
            header=header,
            block_top=block_top,
            block_bottom=block_bottom,
            page_number=page_number,
            source_file=source_file,
            catalog_name=catalog_name,
            brand=brand,
            prefix=prefix,
            vehicle_brand=vehicle_brand,
            product_group=product_group,
        )

        if row:
            rows.append(row)

    return rows


def parse_product_block(
    block_words: List[PdfWord],
    header: HeaderRow,
    block_top: float,
    block_bottom: float,
    page_number: int,
    source_file: str,
    catalog_name: str,
    brand: str,
    prefix: str,
    vehicle_brand: str,
    product_group: str,
) -> Optional[CatalogRow]:
    description_word = find_description_word(
        words=block_words,
        sem_x=header.sem_x,
    )

    description_y = description_word.y0 if description_word else None

    article = extract_article(
        words=block_words,
        sem_x=header.sem_x,
        header_y=header.y,
        description_y=description_y,
    )

    oe_numbers = extract_oe_numbers(
        words=block_words,
        ref_x=header.ref_x,
        header_y=header.y,
        description_y=description_y,
    )

    description = extract_description(
        words=block_words,
        description_word=description_word,
        block_bottom=block_bottom,
    )

    raw_text = build_raw_text(block_words)

    if not article and not oe_numbers:
        return None

    return CatalogRow(
        prefix=prefix,
        article=article,
        brand=brand,
        oe_numbers=oe_numbers,
        description=description,
        type_model="",
        catalog_name=catalog_name,
        vehicle_brand=vehicle_brand,
        product_group=product_group,
        page=page_number,
        status=RowStatus.NEEDS_REVIEW,
        reason="pending validation",
        raw_text=raw_text,
        source_file=source_file,
    )


def extract_words(page) -> List[PdfWord]:
    result: List[PdfWord] = []

    for item in page.get_text("words"):
        result.append(
            PdfWord(
                x0=float(item[0]),
                y0=float(item[1]),
                x1=float(item[2]),
                y1=float(item[3]),
                text=str(item[4]),
                block_no=int(item[5]),
                line_no=int(item[6]),
                word_no=int(item[7]),
            )
        )

    return result


def find_header_rows(words: List[PdfWord]) -> List[HeaderRow]:
    """
    Ищем строки-заголовки, где на одной линии есть SEM и REF.

    В Semlastik порядок может быть разный:
    - SEM NO. | REF NO. | Type/Model
    - Type/Model | REF NO. | SEM NO.
    """

    headers: List[HeaderRow] = []

    sem_words = [
        word for word in words
        if normalize_token(word.text) == "SEM"
    ]

    for sem_word in sem_words:
        line_words = get_same_line_words(words, sem_word.y0, tolerance=4)

        ref_words = [
            word for word in line_words
            if normalize_token(word.text) == "REF"
        ]

        if not ref_words:
            continue

        ref_word = ref_words[0]

        line_text = " ".join(
            word.text for word in sorted(line_words, key=lambda item: item.x0)
        )

        headers.append(
            HeaderRow(
                y=sem_word.y0,
                sem_x=sem_word.x0,
                ref_x=ref_word.x0,
                text=line_text,
            )
        )

    headers = sorted(headers, key=lambda item: item.y)

    # Убираем дубли, если PDF отдал один и тот же заголовок несколько раз.
    unique_headers: List[HeaderRow] = []

    for header in headers:
        if not unique_headers:
            unique_headers.append(header)
            continue

        if abs(header.y - unique_headers[-1].y) > 10:
            unique_headers.append(header)

    return unique_headers


def extract_article(
    words: List[PdfWord],
    sem_x: float,
    header_y: float,
    description_y: Optional[float],
) -> str:
    y_min = header_y + 8
    y_max = min(description_y or header_y + 65, header_y + 75)

    x_min = max(0, sem_x - 65)
    x_max = sem_x + 120

    lines = extract_lines_in_window(
        words=words,
        x_min=x_min,
        x_max=x_max,
        y_min=y_min,
        y_max=y_max,
    )

    if not lines:
        return ""

    article = normalize_article(lines[0])

    return article


def extract_oe_numbers(
    words: List[PdfWord],
    ref_x: float,
    header_y: float,
    description_y: Optional[float],
) -> List[str]:
    y_min = header_y + 8
    y_max = min(description_y or header_y + 80, header_y + 95)

    x_min = max(0, ref_x - 75)
    x_max = ref_x + 135

    lines = extract_lines_in_window(
        words=words,
        x_min=x_min,
        x_max=x_max,
        y_min=y_min,
        y_max=y_max,
    )

    oe_numbers: List[str] = []

    for line in lines:
        for oe in split_ref_text_to_oe_numbers(line):
            if is_probably_oe_number(oe):
                oe_numbers.append(oe)

    return unique_keep_order(oe_numbers)


def split_ref_text_to_oe_numbers(text: str) -> List[str]:
    """
    Разбивает REF NO. строку на отдельные OE номера.

    Примеры:
    "652516 656274" -> ["652516", "656274"]
    "800 1/2" -> ["800 1/2"]
    "SACHS 318890" -> ["SACHS 318890"]
    "8283000639 TYPE 20 LONG STROK" -> ["8283000639"]
    """

    text = normalize_oe(text)

    if not text:
        return []

    text = clean_ref_line(text)

    if not text:
        return []

    candidates: List[str] = []

    comma_parts = re.split(r"[,;]+", text)

    for part in comma_parts:
        part = normalize_oe(part)

        if not part:
            continue

        tokens = part.split()

        if should_split_ref_tokens(tokens):
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


def clean_ref_line(text: str) -> str:
    """
    Убирает из REF NO. строки элементы, которые относятся к Type/Model.

    Например:
    8283000639 TYPE 20 LONG STROK -> 8283000639
    WVA:29030 8285388435 -> 8285388435
    """

    text = normalize_oe(text)

    # WVA обычно технический номер колодки, не OE.
    text = re.sub(r"^WVA\s*[:：]?\s*\d+\s*", "", text, flags=re.IGNORECASE)

    # TYPE / TIP / TİP — это применяемость, а не OE.
    text = re.sub(r"\b(TYPE|TIP|TİP)\b.*$", "", text, flags=re.IGNORECASE)

    return normalize_oe(text)


def should_split_ref_tokens(tokens: List[str]) -> bool:
    """
    Решает, надо ли строку REF NO. разбивать по пробелам.

    Разбиваем только безопасные случаи:
    - несколько отдельных числовых/OE токенов
    - каждый токен длиной >= 4
    - нет дробных номеров типа 800 1/2
    - нет брендов типа SACHS 318890
    """

    if len(tokens) <= 1:
        return False

    normalized_tokens = [normalize_oe(token) for token in tokens if normalize_oe(token)]

    if len(normalized_tokens) <= 1:
        return False

    # Не разбиваем номера вида "800 1/2", "685 2/3".
    if any("/" in token and len(token.replace("/", "")) <= 3 for token in normalized_tokens):
        return False

    # Не разбиваем строки вида "SACHS 318890".
    if any(not any(char.isdigit() for char in token) for token in normalized_tokens):
        return False

    # Разбиваем только токены вида 652516, 1440841S, 0009939996.
    for token in normalized_tokens:
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


def extract_description(
    words: List[PdfWord],
    description_word: Optional[PdfWord],
    block_bottom: float,
) -> str:
    if not description_word:
        return ""

    x_min = max(0, description_word.x0 - 90)
    x_max = description_word.x0 + 190

    y_min = description_word.y1 + 1
    y_max = min(description_word.y1 + 50, block_bottom)

    lines = extract_lines_in_window(
        words=words,
        x_min=x_min,
        x_max=x_max,
        y_min=y_min,
        y_max=y_max,
        remove_labels=True,
    )

    cleaned_lines = []

    for line in lines:
        line = clean_description_line(line)

        if line:
            cleaned_lines.append(line)

    return clean_text(" ".join(cleaned_lines))


def find_description_word(
    words: List[PdfWord],
    sem_x: float,
) -> Optional[PdfWord]:
    candidates = [
        word for word in words
        if normalize_token(word.text) == "DESCRIPTION"
    ]

    if not candidates:
        return None

    # Берем Description, который ближе всего к колонке SEM NO.
    return min(candidates, key=lambda word: abs(word.x0 - sem_x))


def extract_lines_in_window(
    words: List[PdfWord],
    x_min: float,
    x_max: float,
    y_min: float,
    y_max: float,
    remove_labels: bool = True,
) -> List[str]:
    selected = [
        word for word in words
        if y_min <= word.y0 < y_max
        and x_min <= word.cx <= x_max
    ]

    if remove_labels:
        selected = [
            word for word in selected
            if normalize_token(word.text)
            not in {
                "SEM",
                "NO",
                "REF",
                "TYPEMODEL",
                "DESCRIPTION",
            }
        ]

    lines: List[List[PdfWord]] = []

    for word in sorted(selected, key=lambda item: (item.y0, item.x0)):
        placed = False

        for line in lines:
            if abs(line[0].y0 - word.y0) < 4:
                line.append(word)
                placed = True
                break

        if not placed:
            lines.append([word])

    result: List[str] = []

    for line in lines:
        text = " ".join(
            word.text for word in sorted(line, key=lambda item: item.x0)
        ).strip()

        if text:
            result.append(text)

    return result


def get_same_line_words(
    words: List[PdfWord],
    y: float,
    tolerance: float = 4,
) -> List[PdfWord]:
    return [
        word for word in words
        if abs(word.y0 - y) <= tolerance
    ]


def clean_description_line(line: str) -> str:
    """
    Убираем из description технические размеры.

    Например:
    Tie Rod L=590 M=24x1,5 K=30 KG=5.340
    станет:
    Tie Rod
    """

    stop_prefixes = (
        "KG=",
        "M=",
        "D=",
        "L=",
        "H=",
        "AX=",
        "K=",
        "B=",
        "A=",
        "LD=",
        "UD=",
        "N=",
        "PK=",
        "W=",
        "T=",
    )

    result = []

    for token in line.split():
        upper = token.upper()

        if any(upper.startswith(prefix) for prefix in stop_prefixes):
            break

        result.append(token)

    return clean_text(" ".join(result))


def build_raw_text(words: List[PdfWord]) -> str:
    """
    Собираем текст блока/страницы по координатам.
    Это нужно для ручной проверки.
    """

    if not words:
        return ""

    lines: List[List[PdfWord]] = []

    for word in sorted(words, key=lambda item: (item.y0, item.x0)):
        placed = False

        for line in lines:
            if abs(line[0].y0 - word.y0) < 4:
                line.append(word)
                placed = True
                break

        if not placed:
            lines.append([word])

    result_lines = []

    for line in lines:
        line_text = " ".join(
            word.text for word in sorted(line, key=lambda item: item.x0)
        )
        result_lines.append(line_text)

    return clean_text("\n".join(result_lines))


def looks_like_product_page(text: str) -> bool:
    upper = text.upper()

    # Товарная страница должна содержать эти признаки.
    return (
        "SEM NO" in upper
        and "REF NO" in upper
        and "DESCRIPTION" in upper
    )


def is_non_product_list_page(text: str) -> bool:
    """
    KEY LIST и CROSS LIST нельзя парсить как обычные товары.
    Это справочные таблицы, а не карточки деталей.
    """

    upper = text.upper()

    if "KEY LIST" in upper:
        return True

    if "CROSS LIST" in upper:
        return True

    if re.search(r"REF\.?\s*NO\s+SEM\s+NO\s+PAGE", upper):
        return True

    if re.search(r"SEM\s+NO\s+BRAND\s+PAGE", upper):
        return True

    return False


def extract_vehicle_brand(text: str) -> str:
    """
    Определяет vehicle brand на странице:
    DENNIS / SETRA / NEOPLAN / IKARUS / VANHOOL / BOVA / SOLARIS / DAF ...

    В Semlastik это не поставщик.
    Поставщик остается brand = SEMLASTIK.
    """

    lines = [
        clean_text(line)
        for line in text.splitlines()
        if clean_text(line)
    ]

    # Обычно марка находится в верхней части страницы.
    top_lines = lines[:12]

    candidates = sorted(SEMLASTIK_VEHICLE_BRANDS, key=len, reverse=True)

    for line in top_lines:
        normalized_line = normalize_brand_line(line)

        for candidate in candidates:
            normalized_candidate = normalize_brand_line(candidate)

            if normalized_line == normalized_candidate:
                return candidate.replace(" ", "")

            if normalized_line.startswith(normalized_candidate + " "):
                return candidate.replace(" ", "")

            if normalized_line.endswith(" " + normalized_candidate):
                return candidate.replace(" ", "")

    return ""


def extract_product_group(text: str) -> str:
    """
    Определяет товарную группу страницы:
    Steering Group / Suspension Group / Cab Group и т.д.
    """

    upper = text.upper()

    for group_name in PRODUCT_GROUP_NAMES:
        if group_name.upper() in upper:
            return group_name

    return ""


def normalize_brand_line(value: str) -> str:
    value = str(value or "").upper()
    value = value.replace("|", " ")
    value = re.sub(r"[^A-Z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_token(value: str) -> str:
    value = str(value or "").upper()
    value = value.replace("/", "")
    value = re.sub(r"[^A-Z0-9]+", "", value)
    return value