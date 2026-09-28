"""Читает прайс NANFU и пишет одну таблицу цен в out/prices.xlsx."""

import hashlib
import json
import re
import sys
from datetime import date
from pathlib import Path

import xlrd
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUT = ROOT / "out"
SUPPLIERS_PATH = DATA / "suppliers.json"
SEEN_PATH = DATA / "seen.json"

COLUMNS = [
    "Код товара",
    "Штрихкод",
    "Наименование",
    "Единица измерения",
    "Цена",
    "НДС %",
    "Цена включает НДС",
    "Поставщик",
    "Дата обновления",
    "Тип цены",
    "Исходный файл",
]

MONTHS = {
    "январ": 1,
    "феврал": 2,
    "март": 3,
    "апрел": 4,
    "ма": 5,
    "июн": 6,
    "июл": 7,
    "август": 8,
    "сентябр": 9,
    "октябр": 10,
    "ноябр": 11,
    "декабр": 12,
}


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    digest.update(path.read_bytes())
    return digest.hexdigest()


def load_json(path: Path, default):
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def save_json(path: Path, payload) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def cell_text(value) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_price(value):
    text = cell_text(value).replace("\u00a0", "").replace(" ", "").replace(",", ".")
    text = re.sub(r"[^\d.\-]", "", text)
    if not text:
        return None
    return round(float(text), 2)


def month_date(title: str):
    low = title.lower()
    year = re.search(r"20\d{2}", low)
    if not year:
        return ""
    for stem, month in MONTHS.items():
        if stem in low:
            return date(int(year.group()), month, 1).strftime("%d.%m.%Y")
    return ""


def cell_date(value, datemode) -> str:
    if isinstance(value, float) and value > 20000:
        return xlrd.xldate_as_datetime(value, datemode).strftime("%d.%m.%Y")
    return month_date(cell_text(value))


def find_header(sheet):
    for row_index in range(min(sheet.nrows, 15)):
        labels = [cell_text(sheet.cell_value(row_index, col)) for col in range(sheet.ncols)]
        joined = " | ".join(labels).lower()
        if "штрихкод" in joined and "номенклатура" in joined:
            return row_index, labels
    raise ValueError("В файле нет строки заголовков со штрихкодом и номенклатурой")


def convert(path: Path) -> tuple[list[dict], list[str]]:
    settings = load_json(SUPPLIERS_PATH, {}).get("NANFU", {})
    price_column = settings.get("price_column") or "Промо цена"
    vat_rate = settings.get("vat_rate")
    vat_included = settings.get("vat_included")

    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    header_row, labels = find_header(sheet)
    title_value = sheet.cell_value(0, 0) if sheet.nrows else ""
    title = cell_text(title_value)

    def column_index(name: str) -> int:
        for index, label in enumerate(labels):
            if label.lower() == name.lower():
                return index
        raise ValueError(f"Нет колонки «{name}». В файле: {', '.join(x for x in labels if x)}")

    barcode_col = column_index("Штрихкод (шт.)")
    name_col = column_index("Номенклатура")
    price_col = column_index(price_column)
    updated = cell_date(title_value, book.datemode) or month_date(path.name)

    rows = []
    notes = [
        f"Файл: {path.name}",
        f"Шаблон: NANFU, рабочая цена: {price_column}",
        "НДС не задан: цена записана как в прайсе, без пересчёта.",
        f"Дата из названия месяца: {updated or 'не найдена'}",
    ]
    if vat_rate is None:
        notes.append("Чтобы задать НДС, укажите vat_rate и vat_included в data/suppliers.json для NANFU.")

    for row_index in range(header_row + 1, sheet.nrows):
        name = cell_text(sheet.cell_value(row_index, name_col))
        price = parse_price(sheet.cell_value(row_index, price_col))
        if not name or price is None:
            notes.append(f"Пропущена строка {row_index + 1}: нет наименования или цены")
            continue
        rows.append(
            {
                "Код товара": "",
                "Штрихкод": cell_text(sheet.cell_value(row_index, barcode_col)),
                "Наименование": name,
                "Единица измерения": "",
                "Цена": price,
                "НДС %": "" if vat_rate is None else vat_rate,
                "Цена включает НДС": "" if vat_included is None else ("да" if vat_included else "нет"),
                "Поставщик": "NANFU",
                "Дата обновления": updated,
                "Тип цены": price_column,
                "Исходный файл": path.name,
            }
        )
    notes.append(f"Записано строк: {len(rows)}")
    return rows, notes


def write_table(rows: list[dict], destination: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "Цены"
    sheet.append(COLUMNS)
    for row in rows:
        sheet.append([row[column] for column in COLUMNS])
    for cell in sheet["B"][1:]:
        cell.number_format = "@"
    sheet.column_dimensions["C"].width = 55
    destination.parent.mkdir(parents=True, exist_ok=True)
    book.save(destination)


def main() -> int:
    source = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(r"C:\Users\1\Desktop\Прайсы\Прайс NANFU сентябрь 26.xls")
    if not source.exists():
        print(f"Файл не найден: {source}")
        return 1

    digest = file_hash(source)
    seen = load_json(SEEN_PATH, {})
    if digest in seen:
        print(f"Этот файл уже обработан: {seen[digest]}")
        print("Повторная копия пропущена.")
        return 0

    rows, notes = convert(source)
    destination = OUT / "prices.xlsx"
    write_table(rows, destination)
    log_path = OUT / "prices.log"
    log_path.write_text("\n".join(notes) + "\n", encoding="utf-8")
    seen[digest] = source.name
    save_json(SEEN_PATH, seen)

    print(f"Готово: {destination}")
    print(f"Строк: {len(rows)}")
    print(f"Лог: {log_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
