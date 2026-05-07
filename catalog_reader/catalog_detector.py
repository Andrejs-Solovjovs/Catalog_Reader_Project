from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

import pandas as pd

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

from utils import clean_text


@dataclass
class BrandDetectionCandidate:
    brand_name: str
    prefix: str
    score: int
    source: str
    matched_value: str


@dataclass
class CatalogDetectionResult:
    """
    Результат автоопределения каталога.
    """

    detected: bool = False

    brand_name: str = ""
    prefix: str = ""

    confidence: int = 0
    source: str = ""
    matched_value: str = ""

    file_type: str = ""
    parser_hint: str = "auto"

    candidates: List[BrandDetectionCandidate] = field(default_factory=list)
    raw_preview: str = ""


def detect_catalog_supplier(
    file_path: str | Path,
    brand_registry=None,
) -> CatalogDetectionResult:
    """
    Автоопределяет поставщика / производителя детали по самому каталогу.

    Важно:
    - brand = поставщик, например SEMLASTIK, FEBI, 3G
    - prefix берется из справочника брендов
    - vehicle_brand сюда НЕ относится, он извлекается из товарных страниц
    """

    file_path = Path(file_path)
    suffix = file_path.suffix.lower()

    result = CatalogDetectionResult(
        file_type=suffix.replace(".", ""),
        parser_hint="auto",
    )

    if not file_path.exists():
        result.raw_preview = ""
        return result

    preview_text = build_detection_text(file_path)
    result.raw_preview = preview_text[:5000]

    result.parser_hint = detect_parser_hint(
        file_path=file_path,
        preview_text=preview_text,
    )

    if not brand_registry:
        return result

    candidates = find_brand_candidates(
        text=preview_text,
        brand_registry=brand_registry,
    )

    result.candidates = candidates

    if not candidates:
        return result

    best = sorted(candidates, key=lambda item: item.score, reverse=True)[0]

    result.detected = True
    result.brand_name = best.brand_name
    result.prefix = best.prefix
    result.confidence = best.score
    result.source = best.source
    result.matched_value = best.matched_value

    return result


def build_detection_text(file_path: Path) -> str:
    """
    Собирает текст, по которому будем искать поставщика.

    Источники:
    - имя файла
    - первые страницы PDF
    - первые строки Excel
    """

    suffix = file_path.suffix.lower()

    parts = [
        file_path.name,
        file_path.stem,
    ]

    if suffix == ".pdf":
        parts.append(read_pdf_preview_text(file_path))

    elif suffix in {".xlsx", ".xls", ".xlsm", ".csv"}:
        parts.append(read_excel_preview_text(file_path))

    return clean_text("\n".join(part for part in parts if part))


def read_pdf_preview_text(
    file_path: Path,
    max_pages: int = 5,
) -> str:
    """
    Читает первые страницы PDF.

    Этого обычно достаточно, чтобы определить поставщика по обложке,
    логотипу или описанию компании.
    """

    if fitz is None:
        return ""

    try:
        doc = fitz.open(file_path)
    except Exception:
        return ""

    texts = []

    try:
        pages_count = min(len(doc), max_pages)

        for page_index in range(pages_count):
            try:
                page = doc[page_index]
                text = page.get_text("text") or ""
                texts.append(text)
            except Exception:
                continue

    finally:
        doc.close()

    return clean_text("\n".join(texts))


def read_excel_preview_text(
    file_path: Path,
    max_sheets: int = 5,
    max_rows: int = 30,
    max_cols: int = 20,
) -> str:
    """
    Читает первые строки Excel/CSV.

    Для Excel-каталогов поставщик часто есть:
    - в имени файла
    - в названии листа
    - в первых строках
    - в заголовках колонок
    """

    suffix = file_path.suffix.lower()

    try:
        if suffix == ".csv":
            df = pd.read_csv(
                file_path,
                nrows=max_rows,
                dtype=str,
                encoding_errors="ignore",
            )
            return dataframe_preview_to_text("CSV", df, max_rows=max_rows, max_cols=max_cols)

        excel_file = pd.ExcelFile(file_path)

        texts = []

        for sheet_name in excel_file.sheet_names[:max_sheets]:
            try:
                df = pd.read_excel(
                    excel_file,
                    sheet_name=sheet_name,
                    nrows=max_rows,
                    dtype=str,
                    header=None,
                )

                texts.append(
                    dataframe_preview_to_text(
                        sheet_name=sheet_name,
                        df=df,
                        max_rows=max_rows,
                        max_cols=max_cols,
                    )
                )
            except Exception:
                continue

        return clean_text("\n".join(texts))

    except Exception:
        return ""


def dataframe_preview_to_text(
    sheet_name: str,
    df: pd.DataFrame,
    max_rows: int,
    max_cols: int,
) -> str:
    values = [f"SHEET: {sheet_name}"]

    if df is None or df.empty:
        return "\n".join(values)

    df = df.iloc[:max_rows, :max_cols].fillna("")

    for _, row in df.iterrows():
        row_values = [
            str(value).strip()
            for value in row.tolist()
            if str(value).strip()
        ]

        if row_values:
            values.append(" | ".join(row_values))

    return clean_text("\n".join(values))


def find_brand_candidates(
    text: str,
    brand_registry,
) -> List[BrandDetectionCandidate]:
    """
    Ищет бренды из справочника в тексте каталога.

    Чем точнее совпадение, тем выше score.
    """

    text = str(text or "")
    normalized_text = normalize_detection_text(text)

    candidates: List[BrandDetectionCandidate] = []

    for brand in getattr(brand_registry, "brands", []):
        searchable_values = [
            brand.name,
            brand.pref,
            *(brand.synonyms or []),
        ]

        for value in searchable_values:
            value = str(value or "").strip()

            if not value:
                continue

            normalized_value = normalize_detection_text(value)

            if not normalized_value:
                continue

            score = score_brand_match(
                normalized_text=normalized_text,
                normalized_value=normalized_value,
                original_value=value,
                brand_name=brand.name,
                prefix=brand.pref,
            )

            if score <= 0:
                continue

            candidates.append(
                BrandDetectionCandidate(
                    brand_name=brand.name,
                    prefix=brand.pref,
                    score=score,
                    source="file_content",
                    matched_value=value,
                )
            )

    return merge_brand_candidates(candidates)


def score_brand_match(
    normalized_text: str,
    normalized_value: str,
    original_value: str,
    brand_name: str,
    prefix: str,
) -> int:
    """
    Оценка совпадения.

    100 = очень уверенно
    70-90 = похоже
    ниже 70 лучше не автоподставлять без проверки
    """

    if not normalized_text or not normalized_value:
        return 0

    # Очень короткие значения типа "3G" опасны:
    # они могут случайно встречаться в размерах или кодах.
    is_short = len(normalized_value) <= 3

    if is_short:
        if has_strict_token_match(normalized_text, normalized_value):
            return 85

        return 0

    if normalized_value in normalized_text:
        # Длинное название бренда найдено в тексте.
        if len(normalized_value) >= 8:
            return 100

        return 90

    # SEM LASTIK и SEMLASTIK должны считаться совпадением.
    compact_text = normalized_text.replace(" ", "")
    compact_value = normalized_value.replace(" ", "")

    if compact_value and compact_value in compact_text:
        if len(compact_value) >= 8:
            return 98

        return 88

    return 0


def has_strict_token_match(
    normalized_text: str,
    normalized_value: str,
) -> bool:
    """
    Строго ищет короткий бренд как отдельный токен.

    Например:
    3G должно совпасть с " 3G ",
    но не должно совпасть внутри случайной строки.
    """

    pattern = rf"(^|[^A-Z0-9]){re.escape(normalized_value)}([^A-Z0-9]|$)"
    return re.search(pattern, normalized_text) is not None


def merge_brand_candidates(
    candidates: List[BrandDetectionCandidate],
) -> List[BrandDetectionCandidate]:
    """
    Убирает дубли кандидатов, оставляя лучший score.
    """

    by_key = {}

    for candidate in candidates:
        key = (candidate.brand_name, candidate.prefix)

        if key not in by_key:
            by_key[key] = candidate
            continue

        if candidate.score > by_key[key].score:
            by_key[key] = candidate

    return sorted(by_key.values(), key=lambda item: item.score, reverse=True)


def detect_parser_hint(
    file_path: Path,
    preview_text: str,
) -> str:
    """
    Подсказывает, какой внутренний шаблон больше подходит.

    Для пользователя это все равно будет один обработчик auto.
    """

    suffix = file_path.suffix.lower()
    upper = str(preview_text or "").upper()

    if suffix == ".pdf":
        if "SEM NO" in upper and "REF NO" in upper and "DESCRIPTION" in upper:
            return "semlastik_pdf"

        return "generic_pdf"

    if suffix in {".xlsx", ".xls", ".xlsm", ".csv"}:
        return "generic_excel"

    return "unknown"


def normalize_detection_text(value: str) -> str:
    value = str(value or "").upper()
    value = value.replace("&", " AND ")
    value = value.replace("_", " ")
    value = value.replace("-", " ")
    value = re.sub(r"[^A-ZА-ЯЁ0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()