from __future__ import annotations

from pathlib import Path
from typing import Optional

import pandas as pd
import streamlit as st

from catalog_detector import CatalogDetectionResult, detect_catalog_supplier
from exporters.excel_exporter import export_import_result_to_excel
from exporters.site_import_exporter import export_site_import_from_result
from import_history import (
    add_import_history_item,
    clear_import_history,
    compute_file_hash,
    delete_import_history_item,
    find_import_by_file_hash,
    load_import_history,
)
from models import CatalogRow, ImportResult, RowStatus
from parsers.semlastik_pdf import parse_semlastik_pdf
from utils import ensure_dir, normalize_brand, normalize_prefix, safe_filename

try:
    from brand_registry import BrandRegistry
except Exception:
    BrandRegistry = None

try:
    from parsers.generic_excel import parse_generic_excel_catalog
except Exception:
    parse_generic_excel_catalog = None


UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("output")
BRANDS_JSON_PATH = Path("data") / "exchange_export_brands_js_.json"


def main() -> None:
    st.set_page_config(
        page_title="Catalog Reader MVP",
        page_icon="📄",
        layout="wide",
    )

    ensure_dir(UPLOAD_DIR)
    ensure_dir(OUTPUT_DIR)
    ensure_dir("data")

    st.title("📄 Catalog Reader MVP")
    st.caption("MVP для чтения каталогов поставщиков и извлечения OE номеров")

    registry = load_brand_registry()

    with st.sidebar:
        st.header("Настройки импорта")

        manual_brand_input = st.text_input(
            "Бренд поставщика",
            value=st.session_state.get("manual_brand_input", ""),
            key="manual_brand_input",
            help="Можно оставить пустым. Приложение попробует определить поставщика по самому каталогу.",
        )

        sidebar_brand = None

        if manual_brand_input and registry:
            sidebar_brand = resolve_brand_from_registry(
                registry=registry,
                brand_input=manual_brand_input,
            )

        if sidebar_brand:
            st.success(f"Найден бренд: {sidebar_brand.pref} - {sidebar_brand.name}")

            manual_prefix_input = sidebar_brand.pref

            st.text_input(
                "Наш prefix",
                value=manual_prefix_input,
                disabled=True,
                help="Prefix подтянут автоматически из справочника брендов.",
            )

            if sidebar_brand.synonyms:
                st.caption("Синонимы: " + ", ".join(sidebar_brand.synonyms[:5]))

        else:
            if manual_brand_input:
                st.warning(
                    "Бренд не найден в справочнике или справочник не загружен. "
                    "Проверь файл data/exchange_export_brands_js_.json."
                )

            manual_prefix_input = st.text_input(
                "Prefix вручную",
                value=st.session_state.get("manual_prefix_input", ""),
                key="manual_prefix_input",
                help="Нужно только если бренд не найден в справочнике.",
            )

        st.divider()

        st.caption(
            "Обычно достаточно загрузить каталог. "
            "Поставщик и prefix будут определены автоматически по файлу и справочнику брендов."
        )

    tab_import, tab_site_import, tab_history = st.tabs(
        [
            "Обработка каталога",
            "Импорт на сайт",
            "История обработок",
        ]
    )

    with tab_import:
        show_import_tab(
            registry=registry,
            manual_brand_input=manual_brand_input,
            manual_prefix_input=manual_prefix_input,
        )

    with tab_site_import:
        show_site_import_tab()

    with tab_history:
        show_history_tab()

def show_import_tab(
    registry,
    manual_brand_input: str,
    manual_prefix_input: str,
) -> None:
    uploaded_file = st.file_uploader(
        "Загрузи каталог PDF или Excel",
        type=["pdf", "xlsx", "xls", "xlsm", "csv"],
    )

    if uploaded_file is None:
        st.info("Загрузи каталог, чтобы начать обработку.")
        return

    saved_path = save_uploaded_file(uploaded_file)

    st.success(f"Файл загружен: {saved_path.name}")

    file_hash = compute_file_hash(saved_path)
    previous_import = find_import_by_file_hash(file_hash)

    if previous_import:
        show_duplicate_file_warning(previous_import)

    detection = detect_catalog_supplier(
        file_path=saved_path,
        brand_registry=registry,
    )

    brand, prefix, brand_source = resolve_brand_and_prefix(
        registry=registry,
        detection=detection,
        manual_brand_input=manual_brand_input,
        manual_prefix_input=manual_prefix_input,
    )

    show_detection_panel(
        detection=detection,
        brand=brand,
        prefix=prefix,
        brand_source=brand_source,
    )

    if not brand or not prefix:
        st.error("Не удалось определить бренд или prefix. Укажи их вручную слева.")
        return

    col1, col2, col3 = st.columns(3)

    with col1:
        st.write("**Поставщик / brand:**", normalize_brand(brand))

    with col2:
        st.write("**Prefix:**", normalize_prefix(prefix))

    with col3:
        st.write("**Обработчик:**", detection.parser_hint or "auto")

    if previous_import:
        force_reprocess = st.checkbox(
            "Все равно обработать этот файл заново",
            value=False,
        )

        if not force_reprocess:
            st.info("Файл уже есть в истории. Можно скачать прошлый результат или включить повторную обработку.")
            return

    if st.button("Обработать каталог", type="primary"):
        with st.spinner("Обрабатываю каталог..."):
            result = run_auto_parser(
                file_path=saved_path,
                brand=brand,
                prefix=prefix,
                parser_hint=detection.parser_hint,
            )

            output_path = export_import_result_to_excel(
                result=result,
                output_dir=OUTPUT_DIR,
            )

            add_import_history_item(
                result=result,
                output_path=output_path,
                source_path=saved_path,
                parser_name=detection.parser_hint or "auto",
                detected_brand=detection.brand_name,
                detection_source=detection.source,
            )

            st.session_state["last_result"] = result
            st.session_state["last_output_path"] = str(output_path)

            st.success("Обработка завершена и добавлена в историю.")

    if "last_result" in st.session_state:
        show_result(st.session_state["last_result"])

    if "last_output_path" in st.session_state:
        show_download_button(Path(st.session_state["last_output_path"]))


def show_duplicate_file_warning(previous_import: dict) -> None:
    st.warning(
        "Этот файл уже обрабатывался ранее. "
        "Можно скачать прошлый результат или обработать файл заново."
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.write("**Дата:**", previous_import.get("created_at", ""))

    with col2:
        st.write("**Ready:**", previous_import.get("ready", 0))
        st.write("**Needs review:**", previous_import.get("needs_review", 0))

    with col3:
        output_path = Path(previous_import.get("output_path", ""))

        if output_path.exists():
            with open(output_path, "rb") as file:
                st.download_button(
                    label="Скачать прошлый Excel",
                    data=file,
                    file_name=output_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.info("Прошлый Excel файл не найден на диске.")


def resolve_brand_and_prefix(
    registry,
    detection: CatalogDetectionResult,
    manual_brand_input: str,
    manual_prefix_input: str,
) -> tuple[str, str, str]:
    """
    Приоритет такой:

    1. Если пользователь ввел бренд вручную — используем его.
       Prefix пробуем найти в справочнике.
    2. Если вручную ничего нет — используем автоопределение по каталогу.
    3. Если ничего не найдено — просим prefix вручную.
    """

    manual_brand_input = str(manual_brand_input or "").strip()
    manual_prefix_input = str(manual_prefix_input or "").strip()

    if manual_brand_input:
        found = resolve_brand_from_registry(registry, manual_brand_input)

        if found:
            return found.name, found.pref, "manual_input_registry"

        return manual_brand_input, manual_prefix_input, "manual_input"

    if detection and detection.detected:
        return detection.brand_name, detection.prefix, "auto_detection"

    return "", manual_prefix_input, "not_detected"


def show_detection_panel(
    detection: CatalogDetectionResult,
    brand: str,
    prefix: str,
    brand_source: str,
) -> None:
    st.subheader("Определение поставщика")

    if detection.detected:
        st.success(
            f"Поставщик определен автоматически: {detection.prefix} - {detection.brand_name}"
        )

        col1, col2, col3 = st.columns(3)

        with col1:
            st.write("**Confidence:**", detection.confidence)

        with col2:
            st.write("**Найдено по:**", detection.matched_value)

        with col3:
            st.write("**Шаблон:**", detection.parser_hint)

    else:
        st.warning("Поставщик не определен автоматически.")

    st.caption(f"Итоговый источник бренда: {brand_source}")

    with st.expander("Показать детали автоопределения"):
        st.write("**Итоговый brand:**", brand)
        st.write("**Итоговый prefix:**", prefix)
        st.write("**Тип файла:**", detection.file_type)
        st.write("**Parser hint:**", detection.parser_hint)

        if detection.candidates:
            candidate_rows = [
                {
                    "brand_name": item.brand_name,
                    "prefix": item.prefix,
                    "score": item.score,
                    "matched_value": item.matched_value,
                    "source": item.source,
                }
                for item in detection.candidates[:20]
            ]

            st.dataframe(
                pd.DataFrame(candidate_rows),
                use_container_width=True,
                height=250,
            )

        if detection.raw_preview:
            st.text_area(
                "Фрагмент текста, по которому определяли поставщика",
                value=detection.raw_preview[:3000],
                height=250,
            )

def show_site_import_tab() -> None:
    """
    Формирует файл для импорта на сайт.

    Формат сайта:
    brand | code | brand_from | code_from

    Где:
    brand      = производитель детали / поставщик
    code       = артикул поставщика
    brand_from = производитель оригинального номера / vehicle_brand
    code_from  = OE номер
    """

    st.subheader("Импорт на сайт")

    st.write(
        "Эта вкладка формирует отдельный Excel-файл для сайта в формате:"
    )

    st.code(
        "brand | code | brand_from | code_from",
        language="text",
    )

    st.caption(
        "Файл создается только из последнего обработанного каталога. "
        "В сайт-файл попадают только строки со статусом ready, где есть article, OE номер и vehicle_brand."
    )

    if "last_result" not in st.session_state:
        st.info("Сначала обработай каталог во вкладке “Обработка каталога”.")
        return

    result = st.session_state["last_result"]
    summary = result.summary()

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Всего строк", summary["total"])

    with col2:
        st.metric("Ready", summary["ready"])

    with col3:
        st.metric("Needs review", summary["needs_review"])

    with col4:
        st.metric("No OE", summary["no_oe"])

    site_preview_df = result_to_site_import_preview_dataframe(result)

    if site_preview_df.empty:
        st.warning(
            "Нет строк, которые можно безопасно отправить на сайт. "
            "Проверь, что в результате есть ready-строки с vehicle_brand и OE номерами."
        )
        return

    st.write("**Предпросмотр файла для сайта**")

    st.dataframe(
        site_preview_df.head(200),
        use_container_width=True,
        height=400,
    )

    st.caption(f"Всего строк для сайта: {len(site_preview_df)}")

    if st.button("Сформировать файл для сайта", type="primary"):
        export_result = export_site_import_from_result(result)

        st.session_state["last_site_import_path"] = str(export_result.ready_path)

        if export_result.review_path:
            st.session_state["last_site_import_review_path"] = str(export_result.review_path)
        else:
            st.session_state.pop("last_site_import_review_path", None)

        st.success(
            f"Файл для сайта сформирован. Строк: {export_result.ready_count}. "
            f"Отправлено в review: {export_result.skipped_count}."
        )

    if "last_site_import_path" in st.session_state:
        site_import_path = Path(st.session_state["last_site_import_path"])

        if site_import_path.exists():
            with open(site_import_path, "rb") as file:
                st.download_button(
                    label="Скачать файл для сайта",
                    data=file,
                    file_name=site_import_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    if "last_site_import_review_path" in st.session_state:
        review_path = Path(st.session_state["last_site_import_review_path"])

        if review_path.exists():
            st.warning(
                "Часть строк не попала в файл для сайта. "
                "Скачай review-файл и проверь причины."
            )

            with open(review_path, "rb") as file:
                st.download_button(
                    label="Скачать review-файл",
                    data=file,
                    file_name=review_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

def result_to_site_import_preview_dataframe(result: ImportResult) -> pd.DataFrame:
    """
    Делает предпросмотр формата для сайта:

    brand | code | brand_from | code_from
    """

    rows = []

    for row in result.ready_rows:
        if not row.article:
            continue

        if not row.vehicle_brand:
            continue

        if not row.oe_numbers:
            continue

        for oe_number in row.oe_numbers:
            if not oe_number:
                continue

            rows.append(
                {
                    "brand": format_preview_site_brand(row.brand),
                    "code": str(row.article).strip(),
                    "brand_from": str(row.vehicle_brand).strip().upper(),
                    "code_from": str(oe_number).strip(),
                    "load_image": "0",
                    "load_characteristics": "0",
                    "load_cross": "1",
                    "load_applicability": "0",
                }
            )

    columns = [
        "brand",
        "code",
        "brand_from",
        "code_from",
        "load_image",
        "load_characteristics",
        "load_cross",
        "load_applicability",
    ]

    df = pd.DataFrame(rows)

    if df.empty:
        return pd.DataFrame(columns=columns)

    df = df.drop_duplicates()

    return df[columns].fillna("").astype(str)


def format_preview_site_brand(value: str) -> str:
    """
    Формат brand для предпросмотра сайта.

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

def show_history_tab() -> None:
    st.subheader("История обработок")

    history = load_import_history()

    if not history:
        st.info("История пока пустая.")
        return

    col1, col2 = st.columns([4, 1])

    with col1:
        st.caption(f"Всего записей в истории: {len(history)}")

    with col2:
        if st.button("Очистить историю", use_container_width=True):
            st.session_state["confirm_clear_history"] = True
            st.rerun()

    if st.session_state.get("confirm_clear_history"):
        st.warning("Точно очистить всю историю? Excel-файлы из папки output НЕ будут удалены.")

        confirm_col1, confirm_col2 = st.columns(2)

        with confirm_col1:
            if st.button("Да, очистить историю", type="primary", use_container_width=True):
                clear_import_history()
                st.session_state.pop("confirm_clear_history", None)
                st.rerun()

        with confirm_col2:
            if st.button("Нет, оставить", use_container_width=True):
                st.session_state.pop("confirm_clear_history", None)
                st.rerun()

    history_df = pd.DataFrame(history)

    visible_columns = [
        "created_at",
        "source_file",
        "brand",
        "prefix",
        "parser_name",
        "detected_brand",
        "total",
        "ready",
        "needs_review",
        "no_oe",
        "duplicates",
        "errors",
        "output_file",
    ]

    for column in visible_columns:
        if column not in history_df.columns:
            history_df[column] = ""

    st.dataframe(
        history_df[visible_columns],
        use_container_width=True,
        height=350,
    )

    st.divider()
    st.write("**Скачать или удалить результаты из истории**")

    for index, item in enumerate(history[:30], start=1):
        history_index = index - 1

        source_file = item.get("source_file", "")
        created_at = item.get("created_at", "")
        brand = item.get("brand", "")
        prefix = item.get("prefix", "")
        ready = item.get("ready", 0)
        needs_review = item.get("needs_review", 0)
        errors = item.get("errors", 0)
        output_path = Path(item.get("output_path", ""))

        title = f"{index}. {source_file}"

        with st.expander(title, expanded=index == 1):
            col1, col2, col3 = st.columns(3)

            with col1:
                st.write(f"**Дата:** {created_at}")
                st.write(f"**Brand:** {brand}")
                st.write(f"**Prefix:** {prefix}")

            with col2:
                st.write(f"**Ready:** {ready}")
                st.write(f"**Needs review:** {needs_review}")
                st.write(f"**Errors:** {errors}")

            with col3:
                if output_path.exists():
                    with open(output_path, "rb") as file:
                        st.download_button(
                            label="Скачать Excel",
                            data=file,
                            file_name=output_path.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"history_download_{history_index}_{output_path.name}",
                            use_container_width=True,
                        )
                else:
                    st.warning("Excel файл не найден на диске.")

                if st.button(
                    "Удалить запись",
                    key=f"history_delete_request_{history_index}",
                    use_container_width=True,
                ):
                    st.session_state["history_delete_pending_index"] = history_index
                    st.rerun()

            if st.session_state.get("history_delete_pending_index") == history_index:
                st.warning(
                    "Точно удалить эту запись из истории? "
                    "Excel-файл из папки output НЕ будет удален."
                )

                confirm_col1, confirm_col2 = st.columns(2)

                with confirm_col1:
                    if st.button(
                        "Да, удалить запись",
                        type="primary",
                        key=f"history_delete_confirm_yes_{history_index}",
                        use_container_width=True,
                    ):
                        deleted = delete_import_history_item(history_index)

                        st.session_state.pop("history_delete_pending_index", None)

                        if deleted:
                            st.success("Запись удалена из истории.")
                        else:
                            st.error("Не удалось удалить запись.")

                        st.rerun()

                with confirm_col2:
                    if st.button(
                        "Нет, оставить",
                        key=f"history_delete_confirm_no_{history_index}",
                        use_container_width=True,
                    ):
                        st.session_state.pop("history_delete_pending_index", None)
                        st.rerun()

def load_brand_registry():
    """
    Загружает справочник брендов.

    Файл должен лежать здесь:
    catalog_reader/data/exchange_export_brands_js_.json
    """

    if BrandRegistry is None:
        st.sidebar.error("brand_registry.py не найден или не импортируется.")
        return None

    if not BRANDS_JSON_PATH.exists():
        st.sidebar.error(
            f"Файл справочника брендов не найден: {BRANDS_JSON_PATH}"
        )
        return None

    try:
        registry = BrandRegistry(BRANDS_JSON_PATH)

        if not registry.brands:
            st.sidebar.error("Справочник брендов загружен, но список Brands пустой.")
            return None

        st.sidebar.caption(f"Брендов в справочнике: {len(registry.brands)}")
        return registry

    except Exception as exc:
        st.sidebar.error(f"Ошибка загрузки справочника брендов: {exc}")
        return None

def resolve_brand_from_registry(registry, brand_input: str):
    """
    Ищет бренд в справочнике.

    Сначала точное совпадение по Name / Synonyms.
    Потом поиск по похожим значениям.
    """

    brand_input = str(brand_input or "").strip()

    if not registry or not brand_input:
        return None

    exact = registry.find_by_name(brand_input)

    if exact:
        return exact

    matches = registry.search(brand_input, limit=10)

    if not matches:
        return None

    if len(matches) == 1:
        return matches[0]

    options = [brand.display_name() for brand in matches]

    selected_display = st.selectbox(
        "Выбери бренд из справочника",
        options=options,
        index=0,
    )

    return registry.find_by_display_name(selected_display)


def save_uploaded_file(uploaded_file) -> Path:
    safe_name = safe_filename(uploaded_file.name)
    path = UPLOAD_DIR / safe_name

    with open(path, "wb") as file:
        file.write(uploaded_file.getbuffer())

    return path


def run_auto_parser(
    file_path: Path,
    brand: str,
    prefix: str,
    parser_hint: str = "auto",
) -> ImportResult:
    """
    Один общий обработчик для пользователя.

    Внутри есть маршрутизация:
    - semlastik_pdf
    - generic_excel
    - другие шаблоны позже
    """

    suffix = file_path.suffix.lower()

    if suffix == ".pdf":
        return parse_semlastik_pdf(
            file_path=file_path,
            brand=brand,
            prefix=prefix,
            start_page=1,
            max_pages=None,
        )

    if suffix in {".xlsx", ".xls", ".xlsm", ".csv"}:
        if parse_generic_excel_catalog is None:
            return ImportResult(
                source_file=file_path.name,
                brand=normalize_brand(brand),
                prefix=normalize_prefix(prefix),
                rows=[
                    CatalogRow(
                        prefix=normalize_prefix(prefix),
                        brand=normalize_brand(brand),
                        catalog_name=file_path.stem,
                        source_file=file_path.name,
                        status=RowStatus.ERROR,
                        reason="Excel-парсер еще не создан. Следующим шагом добавим parsers/generic_excel.py",
                    )
                ],
            )

        return parse_generic_excel_catalog(
            file_path=file_path,
            brand=brand,
            prefix=prefix,
        )

    return ImportResult(
        source_file=file_path.name,
        brand=normalize_brand(brand),
        prefix=normalize_prefix(prefix),
        rows=[
            CatalogRow(
                prefix=normalize_prefix(prefix),
                brand=normalize_brand(brand),
                catalog_name=file_path.stem,
                source_file=file_path.name,
                status=RowStatus.ERROR,
                reason=f"Формат файла пока не поддерживается: {suffix}",
            )
        ],
    )


def show_result(result: ImportResult) -> None:
    st.divider()
    st.subheader("Результат обработки")

    summary = result.summary()

    col1, col2, col3, col4, col5, col6 = st.columns(6)

    col1.metric("Всего строк", summary["total"])
    col2.metric("Ready", summary["ready"])
    col3.metric("Needs review", summary["needs_review"])
    col4.metric("No OE", summary["no_oe"])
    col5.metric("Duplicates", summary["duplicates"])
    col6.metric("Errors", summary["errors"])

    wide_df = result_to_wide_dataframe(result)
    long_df = result_to_long_dataframe(result)

    if wide_df.empty:
        st.warning("Парсер не нашел товарные строки.")
        return

    tab_long, tab_wide, tab_review = st.tabs(
        [
            "Главный формат: один OE = одна строка",
            "Широкий формат: article + oe1/oe2/oe3",
            "Проверка / raw text",
        ]
    )

    with tab_long:
        show_status_filtered_dataframe(
            df=long_df,
            title="Главный формат для будущей базы",
            height=600,
        )

    with tab_wide:
        show_status_filtered_dataframe(
            df=wide_df,
            title="Широкий формат",
            height=600,
        )

    with tab_review:
        review_columns = [
            "status",
            "prefix",
            "article",
            "brand",
            "vehicle_brand",
            "product_group",
            "description",
            "page",
            "reason",
            "raw_text",
        ]

        review_df = wide_df.copy()

        for column in review_columns:
            if column not in review_df.columns:
                review_df[column] = ""

        review_df = review_df[review_columns]

        show_status_filtered_dataframe(
            df=review_df,
            title="Данные для ручной проверки",
            height=650,
        )


def show_status_filtered_dataframe(
    df: pd.DataFrame,
    title: str,
    height: int = 600,
) -> None:
    st.write(f"**{title}**")

    if df.empty:
        st.info("Нет данных для отображения.")
        return

    status_values = sorted(df["status"].dropna().unique()) if "status" in df.columns else []

    if status_values:
        status_filter = st.multiselect(
            "Фильтр по статусу",
            options=status_values,
            default=status_values,
            key=f"status_filter_{title}",
        )

        filtered_df = df[df["status"].isin(status_filter)]
    else:
        filtered_df = df

    st.dataframe(
        filtered_df,
        use_container_width=True,
        height=height,
    )


def result_to_wide_dataframe(result: ImportResult) -> pd.DataFrame:
    """
    Одна строка = один article.
    OE номера идут колонками oe1, oe2, oe3...
    """

    rows = []

    max_oe_count = max((len(row.oe_numbers) for row in result.rows), default=0)
    max_oe_count = max(max_oe_count, 3)

    for row in result.rows:
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
            item[f"oe{index + 1}"] = row.oe_numbers[index] if index < len(row.oe_numbers) else ""

        item["raw_text"] = row.raw_text

        rows.append(item)

    if not rows:
        return pd.DataFrame()

    columns = [
        "status",
        "prefix",
        "article",
        "brand",
        "vehicle_brand",
        "product_group",
        *[f"oe{i + 1}" for i in range(max_oe_count)],
        "description",
        "type_model",
        "page",
        "reason",
        "catalog_name",
        "source_file",
        "raw_text",
    ]

    return pd.DataFrame(rows)[columns].fillna("").astype(str)


def result_to_long_dataframe(result: ImportResult) -> pd.DataFrame:
    """
    Один OE номер = одна строка.
    Это главный формат для будущей базы.
    """

    rows = []

    for row in result.rows:
        status = row.status.value if isinstance(row.status, RowStatus) else str(row.status)

        if row.oe_numbers:
            for index, oe_number in enumerate(row.oe_numbers, start=1):
                rows.append(
                    {
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
                )
        else:
            rows.append(
                {
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
            )

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

    if not rows:
        return pd.DataFrame(columns=columns)

    return pd.DataFrame(rows)[columns].fillna("").astype(str)


def show_download_button(output_path: Path) -> None:
    if not output_path.exists():
        st.error("Excel файл не найден.")
        return

    with open(output_path, "rb") as file:
        st.download_button(
            label="Скачать Excel результат",
            data=file,
            file_name=output_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()