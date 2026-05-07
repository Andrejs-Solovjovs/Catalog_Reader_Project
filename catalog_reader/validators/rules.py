from __future__ import annotations

import re
from typing import Dict, Iterable, List, Tuple

from models import CatalogRow, RowStatus
from utils import is_probably_oe_number, normalize_article, normalize_oe, unique_keep_order


KIT_COMPONENT_PATTERN = re.compile(r"^\s*\d+\s*-\s*.*=>", re.IGNORECASE)


def validate_row(row: CatalogRow) -> CatalogRow:
    """
    Проверяет одну строку каталога и выставляет status/reason.

    Главное правило:
    если есть сомнение — не ставим ready, а отправляем в needs_review.
    """

    row.article = normalize_article(row.article)
    row.oe_numbers = unique_keep_order([normalize_oe(x) for x in row.oe_numbers])

    if not row.article:
        row.status = RowStatus.ERROR
        row.reason = "Не найден артикул поставщика"
        return row

    if not row.oe_numbers:
        row.status = RowStatus.NO_OE
        row.reason = "Не найдены OE номера"
        return row

    bad_oe = [oe for oe in row.oe_numbers if not is_probably_oe_number(oe)]

    if bad_oe:
        row.status = RowStatus.NEEDS_REVIEW
        row.reason = f"Есть подозрительные OE номера: {', '.join(bad_oe)}"
        return row

    if contains_kit_components(row.raw_text):
        row.status = RowStatus.NEEDS_REVIEW
        row.reason = "Похоже на repair kit / состав комплекта, нужна ручная проверка"
        return row

    if looks_like_article_inside_oe(row.article, row.oe_numbers):
        row.status = RowStatus.NEEDS_REVIEW
        row.reason = "Артикул поставщика похож на один из OE номеров, нужна проверка"
        return row

    row.status = RowStatus.READY
    row.reason = "ok"
    return row


def validate_rows(rows: Iterable[CatalogRow]) -> List[CatalogRow]:
    """
    Проверяет список строк и дополнительно отмечает дубли.
    """

    validated = [validate_row(row) for row in rows]
    mark_duplicates(validated)

    return validated


def mark_duplicates(rows: List[CatalogRow]) -> None:
    """
    Отмечает полные дубли article + OE.

    Важно:
    один и тот же OE у разных артикулов не всегда ошибка,
    поэтому duplicate ставим только на полное совпадение.
    """

    seen: Dict[Tuple[str, str], CatalogRow] = {}

    for row in rows:
        if row.status not in {RowStatus.READY, RowStatus.NEEDS_REVIEW}:
            continue

        for oe in row.oe_numbers:
            key = (row.article, oe)

            if key in seen:
                row.status = RowStatus.DUPLICATE
                row.reason = f"Дубль связи article + OE: {row.article} + {oe}"
                break

            seen[key] = row


def contains_kit_components(text: str) -> bool:
    """
    В Semlastik часто встречаются строки:
    1-1396202=>7995
    2-1328885=>11533

    Это похоже на состав комплекта, а не на обычные OE номера.
    Такие строки лучше отправлять на ручную проверку.
    """

    text = str(text or "")

    for line in text.splitlines():
        if KIT_COMPONENT_PATTERN.search(line):
            return True

    return False


def looks_like_article_inside_oe(article: str, oe_numbers: List[str]) -> bool:
    """
    Проверка на подозрительную ситуацию:
    артикул поставщика оказался среди OE номеров.

    Иногда это нормально, но для MVP лучше отправить в ручную проверку.
    """

    article = normalize_article(article)

    if not article:
        return False

    return article in oe_numbers


def explain_status(row: CatalogRow) -> str:
    """
    Человеческое объяснение статуса.
    Можно использовать в интерфейсе.
    """

    if row.status == RowStatus.READY:
        return "Готово к выгрузке"

    if row.status == RowStatus.NEEDS_REVIEW:
        return "Нужна ручная проверка"

    if row.status == RowStatus.ERROR:
        return "Ошибка обработки"

    if row.status == RowStatus.DUPLICATE:
        return "Дубль"

    if row.status == RowStatus.NO_OE:
        return "Нет OE номеров"

    return "Неизвестный статус"