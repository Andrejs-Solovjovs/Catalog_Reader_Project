from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from models import ImportResult
from utils import ensure_dir


HISTORY_PATH = Path("data") / "import_history.json"
MAX_HISTORY_ITEMS = 200


def compute_file_hash(file_path: str | Path) -> str:
    """
    Считает SHA256 файла.

    Это нужно, чтобы понять:
    - этот файл уже обрабатывали или нет
    - даже если файл переименовали, хэш останется тем же
    """

    file_path = Path(file_path)
    sha256 = hashlib.sha256()

    with open(file_path, "rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            sha256.update(chunk)

    return sha256.hexdigest()


def load_import_history(history_path: str | Path = HISTORY_PATH) -> List[Dict[str, Any]]:
    """
    Загружает историю обработок.

    Если файла истории нет — возвращает пустой список.
    """

    history_path = Path(history_path)

    if not history_path.exists():
        return []

    try:
        with open(history_path, "r", encoding="utf-8") as file:
            data = json.load(file)

        if not isinstance(data, list):
            return []

        return data

    except Exception:
        return []


def save_import_history(
    history: List[Dict[str, Any]],
    history_path: str | Path = HISTORY_PATH,
) -> None:
    """
    Сохраняет историю обработок в JSON.
    """

    history_path = Path(history_path)
    ensure_dir(history_path.parent)

    history = history[:MAX_HISTORY_ITEMS]

    with open(history_path, "w", encoding="utf-8") as file:
        json.dump(history, file, ensure_ascii=False, indent=2)


def find_import_by_file_hash(
    file_hash: str,
    history_path: str | Path = HISTORY_PATH,
) -> Optional[Dict[str, Any]]:
    """
    Ищет в истории обработку с таким же file_hash.
    """

    file_hash = str(file_hash or "").strip()

    if not file_hash:
        return None

    history = load_import_history(history_path)

    for item in history:
        if str(item.get("file_hash") or "") == file_hash:
            return item

    return None


def find_import_by_file_path(
    file_path: str | Path,
    history_path: str | Path = HISTORY_PATH,
) -> Optional[Dict[str, Any]]:
    """
    Считает хэш файла и ищет его в истории.
    """

    file_path = Path(file_path)

    if not file_path.exists():
        return None

    file_hash = compute_file_hash(file_path)

    return find_import_by_file_hash(
        file_hash=file_hash,
        history_path=history_path,
    )


def add_import_history_item(
    result: ImportResult,
    output_path: str | Path,
    source_path: str | Path | None = None,
    parser_name: str = "auto",
    detected_brand: str = "",
    detection_source: str = "",
    history_path: str | Path = HISTORY_PATH,
) -> Dict[str, Any]:
    """
    Добавляет одну запись в историю после обработки каталога.
    """

    history = load_import_history(history_path)
    summary = result.summary()

    output_path = Path(output_path)

    file_hash = ""

    if source_path:
        source_path = Path(source_path)

        if source_path.exists():
            file_hash = compute_file_hash(source_path)

    item = {
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "source_file": result.source_file,
        "brand": result.brand,
        "prefix": result.prefix,
        "parser_name": parser_name,
        "detected_brand": detected_brand,
        "detection_source": detection_source,
        "file_hash": file_hash,
        "output_path": str(output_path),
        "output_file": output_path.name,
        "total": summary.get("total", 0),
        "ready": summary.get("ready", 0),
        "needs_review": summary.get("needs_review", 0),
        "no_oe": summary.get("no_oe", 0),
        "duplicates": summary.get("duplicates", 0),
        "errors": summary.get("errors", 0),
    }

    history.insert(0, item)
    save_import_history(history, history_path)

    return item


def clear_import_history(history_path: str | Path = HISTORY_PATH) -> None:
    """
    Очищает историю обработок.
    Файлы Excel при этом НЕ удаляются.
    """

    save_import_history([], history_path)


def get_last_import(history_path: str | Path = HISTORY_PATH) -> Optional[Dict[str, Any]]:
    """
    Возвращает последнюю обработку.
    """

    history = load_import_history(history_path)

    if not history:
        return None

    return history[0]

def delete_import_history_item(
    index: int,
    history_path: str | Path = HISTORY_PATH,
) -> bool:
    """
    Удаляет одну запись из истории по индексу.

    Важно:
    удаляется только запись из истории.
    Excel-файл из output/ НЕ удаляется.
    """

    history = load_import_history(history_path)

    if index < 0 or index >= len(history):
        return False

    del history[index]
    save_import_history(history, history_path)

    return True