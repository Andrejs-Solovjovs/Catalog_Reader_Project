from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional


DEFAULT_BRANDS_JSON_PATH = Path("data") / "exchange_export_brands_js_.json"


@dataclass
class BrandInfo:
    id: str
    name: str
    pref: str
    sup: str = ""
    mfa: str = ""
    visible: bool = True
    virtual: bool = False
    synonyms: List[str] = None

    def display_name(self) -> str:
        return f"{self.pref} - {self.name}"


class BrandRegistry:
    """
    Справочник брендов и внутренних prefix.

    Данные берутся из exchange_export_brands_js_.json.
    """

    def __init__(self, json_path: str | Path = DEFAULT_BRANDS_JSON_PATH):
        self.json_path = Path(json_path)
        self.brands: List[BrandInfo] = []
        self.by_normalized_name: Dict[str, BrandInfo] = {}
        self.by_pref: Dict[str, BrandInfo] = {}

        self.load()

    def load(self) -> None:
        if not self.json_path.exists():
            self.brands = []
            self.by_normalized_name = {}
            self.by_pref = {}
            return

        with open(self.json_path, "r", encoding="utf-8") as file:
            data = json.load(file)

        raw_brands = data.get("Brands", [])

        brands: List[BrandInfo] = []

        for item in raw_brands:
            name = str(item.get("Name") or "").strip()
            pref = str(item.get("Pref") or "").strip().upper()

            if not name or not pref:
                continue

            synonyms = item.get("Synonyms") or []

            if not isinstance(synonyms, list):
                synonyms = []

            brand = BrandInfo(
                id=str(item.get("ID") or ""),
                name=name,
                pref=pref,
                sup=str(item.get("SUP") or ""),
                mfa=str(item.get("MFA") or ""),
                visible=str(item.get("Visible") or "0") == "1",
                virtual=str(item.get("Virtual") or "0") == "1",
                synonyms=[str(x).strip() for x in synonyms if str(x).strip()],
            )

            brands.append(brand)

        self.brands = brands
        self.rebuild_indexes()

    def rebuild_indexes(self) -> None:
        self.by_normalized_name = {}
        self.by_pref = {}

        for brand in self.brands:
            self.by_normalized_name[normalize_brand_key(brand.name)] = brand
            self.by_pref[brand.pref.upper()] = brand

            for synonym in brand.synonyms or []:
                self.by_normalized_name[normalize_brand_key(synonym)] = brand

    def find_by_name(self, name: str) -> Optional[BrandInfo]:
        key = normalize_brand_key(name)
        return self.by_normalized_name.get(key)

    def find_by_pref(self, pref: str) -> Optional[BrandInfo]:
        return self.by_pref.get(str(pref or "").strip().upper())

    def search(self, query: str, limit: int = 30) -> List[BrandInfo]:
        """
        Поиск по названию, synonym и prefix.
        """

        query = str(query or "").strip()

        if not query:
            return self.brands[:limit]

        query_key = normalize_brand_key(query)
        query_upper = query.upper()

        result: List[BrandInfo] = []

        for brand in self.brands:
            searchable_values = [
                brand.name,
                brand.pref,
                *(brand.synonyms or []),
            ]

            normalized_values = [normalize_brand_key(value) for value in searchable_values]
            upper_values = [str(value).upper() for value in searchable_values]

            matched = False

            if brand.pref.upper() == query_upper:
                matched = True
            elif any(query_key in value for value in normalized_values):
                matched = True
            elif any(query_upper in value for value in upper_values):
                matched = True

            if matched:
                result.append(brand)

            if len(result) >= limit:
                break

        return result

    def get_options(self) -> List[str]:
        """
        Опции для selectbox в Streamlit.
        """
        return [brand.display_name() for brand in self.brands]

    def find_by_display_name(self, display_name: str) -> Optional[BrandInfo]:
        """
        Находит бренд из строки вида:
        CZJ - SEMLASTIK
        """

        display_name = str(display_name or "").strip()

        if " - " in display_name:
            pref = display_name.split(" - ", 1)[0].strip()
            found = self.find_by_pref(pref)

            if found:
                return found

        return self.find_by_name(display_name)


def normalize_brand_key(value: str) -> str:
    """
    Нормализация для поиска бренда.

    SEM LASTIK, Semlastik, SEMLASTIK -> SEMLASTIK
    """
    value = str(value or "").upper()
    value = value.replace("&", "AND")
    value = re.sub(r"[^A-ZА-ЯЁ0-9]+", "", value)
    return value


def get_brand_registry() -> BrandRegistry:
    return BrandRegistry(DEFAULT_BRANDS_JSON_PATH)


if __name__ == "__main__":
    registry = get_brand_registry()

    print(f"Загружено брендов: {len(registry.brands)}")

    for query in ["SEMLASTIK", "FEBI", "DAF", "BOSCH"]:
        brand = registry.find_by_name(query)

        if brand:
            print(f"{query}: {brand.pref} - {brand.name}")
        else:
            print(f"{query}: не найден")