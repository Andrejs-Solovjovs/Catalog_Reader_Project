from __future__ import annotations

from enum import Enum
from typing import List, Optional

from pydantic import BaseModel, Field, field_validator


class RowStatus(str, Enum):
    READY = "ready"
    NEEDS_REVIEW = "needs_review"
    ERROR = "error"
    DUPLICATE = "duplicate"
    NO_OE = "no_oe"


class CatalogRow(BaseModel):
    """
    Одна найденная деталь из каталога.

    article = артикул поставщика, например SEM NO.
    oe_numbers = оригинальные номера, например REF NO.

    brand = поставщик, например SEMLASTIK.
    vehicle_brand = марка техники/авто из каталога, например DAF, DENNIS, SETRA.
    """

    prefix: str = Field(default="")
    article: str = Field(default="")
    brand: str = Field(default="")

    oe_numbers: List[str] = Field(default_factory=list)

    description: str = Field(default="")
    type_model: str = Field(default="")

    # Дополнительные поля для многофункционального каталога
    catalog_name: str = Field(default="")
    vehicle_brand: str = Field(default="")
    product_group: str = Field(default="")

    # page = страница PDF / каталога
    page: Optional[int] = None

    status: RowStatus = RowStatus.NEEDS_REVIEW
    reason: str = ""

    raw_text: str = Field(default="")
    source_file: str = Field(default="")

    @field_validator(
        "prefix",
        "article",
        "brand",
        "description",
        "type_model",
        "catalog_name",
        "vehicle_brand",
        "product_group",
        "reason",
        "raw_text",
        "source_file",
    )
    @classmethod
    def clean_text_fields(cls, value: str) -> str:
        if value is None:
            return ""
        return str(value).strip()

    @field_validator("oe_numbers")
    @classmethod
    def clean_oe_numbers(cls, value: List[str]) -> List[str]:
        cleaned = []

        for item in value or []:
            item = str(item).strip()

            if not item:
                continue

            # Убираем лишние пробелы внутри номера, но НЕ удаляем ведущие нули.
            item = " ".join(item.split())

            if item not in cleaned:
                cleaned.append(item)

        return cleaned

    def oe_as_string(self) -> str:
        return ", ".join(self.oe_numbers)


class ImportResult(BaseModel):
    """
    Общий результат обработки одного файла.
    """

    source_file: str
    brand: str
    prefix: str

    rows: List[CatalogRow] = Field(default_factory=list)

    @property
    def ready_rows(self) -> List[CatalogRow]:
        return [row for row in self.rows if row.status == RowStatus.READY]

    @property
    def review_rows(self) -> List[CatalogRow]:
        return [row for row in self.rows if row.status == RowStatus.NEEDS_REVIEW]

    @property
    def error_rows(self) -> List[CatalogRow]:
        return [row for row in self.rows if row.status == RowStatus.ERROR]

    @property
    def duplicate_rows(self) -> List[CatalogRow]:
        return [row for row in self.rows if row.status == RowStatus.DUPLICATE]

    @property
    def no_oe_rows(self) -> List[CatalogRow]:
        return [row for row in self.rows if row.status == RowStatus.NO_OE]

    def summary(self) -> dict:
        return {
            "total": len(self.rows),
            "ready": len(self.ready_rows),
            "needs_review": len(self.review_rows),
            "errors": len(self.error_rows),
            "duplicates": len(self.duplicate_rows),
            "no_oe": len(self.no_oe_rows),
        }