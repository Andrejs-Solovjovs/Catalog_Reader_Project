from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable, List


def normalize_brand(value: str) -> str:
    """
    Приводим бренд к единому виду.
    Например: Sem Lastik, SEMLASTIK, semlastik -> semlastik
    """
    value = str(value or "").strip().lower()
    value = value.replace(" ", "")
    return value


def normalize_prefix(value: str) -> str:
    """
    Prefix лучше хранить в верхнем регистре.
    Например: sem -> SEM
    """
    return str(value or "").strip().upper()


def normalize_article(value: str) -> str:
    """
    Артикул поставщика сохраняем как текст.

    Важно:
    - не удаляем ведущие нули
    - не превращаем в число
    - сохраняем точки и буквы
    """
    value = str(value or "").strip()
    value = " ".join(value.split())
    return value


def normalize_oe(value: str) -> str:
    """
    OE номер сохраняем как текст.

    Важно:
    0300911 должно остаться 0300911,
    1440841S должно остаться 1440841S.
    """
    value = str(value or "").strip()
    value = " ".join(value.split())
    return value


def clean_text(value: str) -> str:
    """
    Базовая очистка текста.
    """
    value = str(value or "").replace("\u00a0", " ")
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def unique_keep_order(items: Iterable[str]) -> List[str]:
    """
    Убираем дубли, но сохраняем порядок.
    """
    result = []

    for item in items:
        item = str(item or "").strip()

        if not item:
            continue

        if item not in result:
            result.append(item)

    return result


def is_probably_oe_number(value: str) -> bool:
    """
    Проверяем, похоже ли значение на OE/reference номер.

    Здесь правило специально НЕ слишком жесткое:
    OE может быть:
    - 1239183
    - 0300911
    - 1440841S
    - ACHF060
    - SACHS 311500

    Но мы не хотим брать мусор типа:
    - KG=0,55
    - M=12x1,5
    - D=Ø60
    - Description
    """
    value = normalize_oe(value)

    if not value:
        return False

    upper = value.upper()

    bad_fragments = [
        "KG=",
        "M=",
        "D=",
        "L=",
        "H=",
        "AX=",
        "Ø",
        "DESCRIPTION",
        "TYPE/MODEL",
        "SEM NO",
        "REF NO",
    ]

    if any(fragment in upper for fragment in bad_fragments):
        return False

    # Слишком короткие значения почти всегда мусор.
    if len(value.replace(" ", "")) < 4:
        return False

    # Должна быть хотя бы одна цифра.
    if not any(char.isdigit() for char in value):
        return False

    # Разрешаем буквы, цифры, пробел, точку, дефис, слэш.
    if not re.fullmatch(r"[A-Za-z0-9 .\-/]+", value):
        return False

    return True


def split_possible_oe_lines(text: str) -> List[str]:
    """
    Разбивает блок текста на возможные OE номера.
    """
    text = clean_text(text)

    candidates = []

    for line in text.splitlines():
        line = line.strip()

        if not line:
            continue

        # Иногда несколько номеров могут быть в одной строке через пробелы/запятые.
        parts = re.split(r"[,;]+", line)

        for part in parts:
            part = normalize_oe(part)

            if is_probably_oe_number(part):
                candidates.append(part)

    return unique_keep_order(candidates)


def ensure_dir(path: str | Path) -> Path:
    """
    Создает папку, если ее нет.
    """
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)
    return path


def safe_filename(value: str) -> str:
    """
    Безопасное имя файла без странных символов.
    """
    value = str(value or "").strip()
    value = re.sub(r"[^\w\-.]+", "_", value, flags=re.UNICODE)
    value = re.sub(r"_+", "_", value)
    return value.strip("_") or "file"