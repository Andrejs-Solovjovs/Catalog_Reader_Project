# Catalog Reader MVP

MVP-приложение для чтения каталогов поставщиков автозапчастей и извлечения оригинальных OE номеров.

Приложение принимает каталоги в формате:

- PDF
- XLSX
- XLS
- XLSM
- CSV

На выходе формируется Excel-файл с единым форматом данных:

```text
prefix | article | brand | vehicle_brand | oe_number