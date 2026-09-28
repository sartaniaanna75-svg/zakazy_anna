"""Приводит известные прайсы к одной таблице out/prices.xlsx.

Запуск:
    python src/convert_prices.py "C:\\Users\\1\\Desktop\\Прайсы"
"""

import hashlib
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path

import xlrd
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "out"
DATA = ROOT / "data"
TEMPLATES_PATH = DATA / "templates.json"

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


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    digest.update(path.read_bytes())
    return digest.hexdigest()


def as_text(value) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    return str(value).replace("\xa0", " ").strip()


def parse_price(value):
    text = as_text(value).replace(" ", "").replace(",", ".")
    text = re.sub(r"[^\d.\-]", "", text)
    if not text or text in {".", "-", "-."}:
        return None
    number = float(text)
    if number <= 0:
        return None
    return round(number, 2)


def barcode_text(value) -> str:
    raw = as_text(value)
    if raw in {"", "##"}:
        return ""
    parts = re.findall(r"\d{8,14}", raw)
    if not parts:
        return raw
    trimmed = []
    for part in parts:
        short = part.lstrip("0")
        if short and any(other.lstrip("0") == short and other != part for other in parts):
            continue
        trimmed.append(part)
    chosen = trimmed or parts
    ean13 = [item for item in chosen if len(item) == 13]
    return ean13[0] if ean13 else chosen[-1]


def find_date(rows) -> str:
    blob = " ".join(as_text(cell) for row in rows[:25] for cell in row)
    match = re.search(r"(\d{2})\.(\d{2})\.(20\d{2})", blob)
    if match:
        return f"{match.group(1)}.{match.group(2)}.{match.group(3)}"
    return ""


def sheet_blob(rows, limit=30) -> str:
    return "\n".join(" | ".join(as_text(cell) for cell in row) for row in rows[:limit])


def load_rows(path: Path):
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        book = load_workbook(path, data_only=True, read_only=True)
        sheet = book.active
        rows = [list(row) for row in sheet.iter_rows(values_only=True)]
        book.close()
        return rows
    if suffix == ".xls":
        book = xlrd.open_workbook(path)
        sheet = book.sheet_by_index(0)
        rows = []
        for row_index in range(sheet.nrows):
            values = []
            for col_index in range(sheet.ncols):
                cell = sheet.cell(row_index, col_index)
                if cell.ctype == xlrd.XL_CELL_DATE:
                    values.append(xlrd.xldate_as_datetime(cell.value, book.datemode))
                else:
                    values.append(cell.value)
            rows.append(values)
        return rows
    return []


def blank_row(name, price, source, supplier, updated, price_type, code="", barcode="", unit=""):
    return {
        "Код товара": code,
        "Штрихкод": barcode,
        "Наименование": name,
        "Единица измерения": unit,
        "Цена": price,
        "НДС %": "",
        "Цена включает НДС": "",
        "Поставщик": supplier,
        "Дата обновления": updated,
        "Тип цены": price_type,
        "Исходный файл": source.name,
    }


def column_of(row, needle: str):
    needle_low = needle.lower()
    hits = []
    for index, cell in enumerate(row):
        label = as_text(cell).lower()
        if needle_low in label:
            hits.append((index, label))
    if needle_low == "код":
        hits = [(index, label) for index, label in hits if "штрих" not in label]
    return hits[0][0] if hits else None


def find_header(rows, must_contain):
    needles = [item.lower() for item in must_contain]
    for index, row in enumerate(rows[:40]):
        joined = " | ".join(as_text(cell).lower() for cell in row)
        if all(needle in joined for needle in needles):
            return index
    return None


def read_nanfu(rows, source):
    header = find_header(rows, ["штрихкод", "номенклатура", "промо цена"])
    if header is None:
        return None
    labels = [as_text(cell) for cell in rows[header]]
    barcode_col = labels.index("Штрихкод (шт.)")
    name_col = labels.index("Номенклатура")
    price_col = labels.index("Промо цена")
    updated = ""
    if rows and isinstance(rows[0][0], datetime):
        updated = rows[0][0].strftime("%d.%m.%Y")
    elif rows and isinstance(rows[0][0], float) and rows[0][0] > 20000:
        updated = datetime(1899, 12, 30).fromordinal(datetime(1899, 12, 30).toordinal() + int(rows[0][0]))
        # Excel serial is handled below if datetime conversion failed
    items = []
    for row in rows[header + 1 :]:
        price = parse_price(row[price_col] if price_col < len(row) else "")
        name = as_text(row[name_col] if name_col < len(row) else "")
        if not name or price is None:
            continue
        items.append(blank_row(name, price, source, "NANFU", updated if isinstance(updated, str) else "", "Промо цена", barcode=barcode_text(row[barcode_col])))
    if rows and isinstance(rows[0][0], float) and rows[0][0] > 20000:
        updated = xlrd.xldate_as_datetime(rows[0][0], 0).strftime("%d.%m.%Y")
        for item in items:
            item["Дата обновления"] = updated
    return items, f"NANFU, рабочая цена «Промо цена», строк {len(items)}. НДС не задан."


def read_discount_table(rows, source, supplier, price_type):
    header = find_header(rows, ["наименование товара", "регулярная цена", "цена"])
    if header is None:
        return None
    labels = rows[header]
    name_col = column_of(labels, "Наименование")
    price_col = None
    regular_col = None
    barcode_col = column_of(labels, "Штрихкод")
    code_col = column_of(labels, "Артикул")
    for index, cell in enumerate(labels):
        label = as_text(cell).lower()
        if label == "цена":
            price_col = index
        if "регулярная" in label:
            regular_col = index
    items = []
    for row in rows[header + 1 :]:
        if price_col is None or price_col >= len(row):
            continue
        price = parse_price(row[price_col])
        name = as_text(row[name_col]) if name_col is not None and name_col < len(row) else ""
        if not name or price is None:
            continue
        code = as_text(row[code_col]) if code_col is not None and code_col < len(row) else ""
        barcode = barcode_text(row[barcode_col]) if barcode_col is not None and barcode_col < len(row) else ""
        items.append(blank_row(name, price, source, supplier, find_date(rows), price_type, code=code, barcode=barcode))
    return items, f"{supplier}: цена после скидки, не регулярная. Строк {len(items)}."


def supplier_from_action(rows) -> str:
    for row in rows[:20]:
        for index, cell in enumerate(row):
            if "название акции" in as_text(cell).lower():
                for next_cell in row[index + 1 :]:
                    text = as_text(next_cell)
                    if text:
                        return text.split(" Сегменты")[0].strip()
    return ""


def read_hierarchy(rows, source):
    header = find_header(rows, ["ценовая группа"])
    if header is None:
        return None
    name_col = column_of(rows[header], "Ценовая группа")
    barcode_col = column_of(rows[header], "Штрихкод")
    unit_col = None
    price_col = None
    price_type = "Цена"
    window = rows[header : header + 3]
    for row in window:
        for index, cell in enumerate(row):
            label = as_text(cell).lower()
            if label == "цена":
                price_col = index
            if label in {"ед.", "ед"}:
                unit_col = index
    if price_col is None:
        for index, cell in enumerate(rows[header]):
            label = as_text(cell)
            if not label or index == name_col:
                continue
            low = label.lower()
            if low in {"штрихкод", "заказ"} or "ценовая группа" in low:
                continue
            price_col = index
            price_type = label
    supplier = "Прайс"
    blob = sheet_blob(rows, 12)
    if "аи-трейд" in blob.lower():
        supplier = 'ООО "АИ-ТРЕЙД КАВКАЗ"'
    elif "опт2" in blob.lower():
        supplier = "Опт2"
    elif "спец-предложение" in blob.lower():
        supplier = "Спец-предложение"
    elif "прайс-лист" in blob.lower():
        supplier = "ОПТ 1"
    items = []
    for row in rows[header + 1 :]:
        if price_col is None or price_col >= len(row):
            continue
        price = parse_price(row[price_col])
        name = as_text(row[name_col]) if name_col is not None and name_col < len(row) else ""
        if not name or price is None:
            continue
        unit = as_text(row[unit_col]) if unit_col is not None and unit_col < len(row) else ""
        barcode = barcode_text(row[barcode_col]) if barcode_col is not None and barcode_col < len(row) else ""
        items.append(blank_row(name, price, source, supplier, find_date(rows), price_type, barcode=barcode, unit=unit))
    return items, f"{supplier}, тип цены «{price_type}», групп без цены пропущено. Строк {len(items)}."


def read_gorokhov(rows, source):
    if not rows or "$БЛАНК_ЗАКАЗА" not in as_text(rows[0][0]):
        return None
    header = None
    for index, row in enumerate(rows[:30]):
        if as_text(row[0]) == "Шапка":
            header = index
            break
    if header is None:
        return None
    labels = rows[header]
    name_col = column_of(labels, "Наименование")
    code_col = column_of(labels, "Код")
    unit_col = column_of(labels, "Ед.")
    price_col = column_of(labels, "Цена")
    barcode_col = column_of(labels, "Штрихкод")
    items = []
    for row in rows[header + 1 :]:
        if as_text(row[0]) != "Строка":
            continue
        price = parse_price(row[price_col])
        name = as_text(row[name_col])
        if not name or price is None:
            continue
        items.append(
            blank_row(
                name,
                price,
                source,
                "Косметика 26",
                find_date(rows),
                "Цена",
                code=as_text(row[code_col]),
                barcode=barcode_text(row[barcode_col]),
                unit=as_text(row[unit_col]),
            )
        )
    return items, f"Бланк Косметика 26, только строки типа «Строка». Строк {len(items)}."


def read_cb_order(rows, source):
    header = find_header(rows, ["штрихкод", "номенклатура", "цпрайс"])
    if header is None:
        return None
    labels = rows[header]
    name_col = column_of(labels, "Номенклатура")
    code_col = column_of(labels, "Код")
    barcode_col = column_of(labels, "Штрихкод")
    pack_col = column_of(labels, "Шт в упаковке")
    price_col = column_of(labels, "ЦПрайс")
    items = []
    skipped = 0
    for row in rows[header + 1 :]:
        name = as_text(row[name_col]) if name_col < len(row) else ""
        price = parse_price(row[price_col]) if price_col < len(row) else None
        pack = as_text(row[pack_col]) if pack_col is not None and pack_col < len(row) else ""
        barcode = barcode_text(row[barcode_col]) if barcode_col < len(row) else ""
        if name.lower() == "итого" or not name or price is None or (not pack and not barcode):
            skipped += 1
            continue
        items.append(
            blank_row(
                name,
                price,
                source,
                "ЦБ",
                find_date([source.name]),
                "ЦПрайс",
                code=as_text(row[code_col]),
                barcode=barcode,
            )
        )
    file_date = ""
    match = re.search(r"(\d{2})\.(\d{2})", source.name)
    if match:
        file_date = f"{match.group(1)}.{match.group(2)}.2026"
        for item in items:
            item["Дата обновления"] = file_date
    return items, f"Бланк заказа ЦБ, разделы без товара пропущены ({skipped}). Строк {len(items)}."


def read_salt(rows, source):
    blob = sheet_blob(rows, 10).lower()
    if "кубанская соляная" not in blob:
        return None
    header = find_header(rows, ["наименование", "цена"])
    if header is None:
        return None
    name_col = column_of(rows[header], "Наименование")
    price_col = column_of(rows[header], "Цена")
    items = []
    for row in rows[header + 1 :]:
        if price_col >= len(row):
            continue
        price = parse_price(row[price_col])
        name = as_text(row[name_col]) if name_col < len(row) else ""
        if not name or price is None:
            continue
        items.append(blank_row(name, price, source, 'ООО "Кубанская Соляная Компания"', find_date(rows) or "23.09.2026", "Цена"))
    return items, f"Кубанская соляная компания, категории без цены пропущены. Строк {len(items)}. НДС внутри названия не разбирался."


def detect(rows, source: Path):
    blob = sheet_blob(rows, 20).lower()
    if "стоимостная оценка склада" in blob:
        return [], "Это оценка склада, не прайс поставщика. В закупку не включено."
    if source.suffix.lower() == ".pdf":
        return [], "PDF пропущен: рабочая копия — соседний xls."
    for reader in (
        read_gorokhov,
        read_nanfu,
        read_salt,
        read_cb_order,
    ):
        result = reader(rows, source)
        if result is not None:
            return result
    if "заявка на проведение акции" in blob:
        supplier = supplier_from_action(rows) or source.stem
        return read_discount_table(rows, source, supplier, "Цена акции")
    if "промо цена" in blob:
        return read_nanfu(rows, source)
    if "наименование товара" in blob and "процент скидки" in blob:
        supplier = "Bagi" if "баги" in source.name.lower() else source.stem
        return read_discount_table(rows, source, supplier, "Цена со скидкой")
    if "ценовая группа" in blob:
        return read_hierarchy(rows, source)
    return None


def remember_unknown(source: Path, rows) -> str:
    templates = {}
    if TEMPLATES_PATH.exists():
        templates = json.loads(TEMPLATES_PATH.read_text(encoding="utf-8"))
    header = next((row for row in rows[:25] if sum(bool(as_text(cell)) for cell in row) >= 3), [])
    templates[source.name] = [as_text(cell) for cell in header if as_text(cell)]
    TEMPLATES_PATH.write_text(json.dumps(templates, ensure_ascii=False, indent=2), encoding="utf-8")
    return "Неизвестный формат. Заголовки записаны в data/templates.json, файл в таблицу не попал."


def convert_file(path: Path):
    if path.suffix.lower() not in {".xls", ".xlsx"}:
        return [], f"{path.name}: формат {path.suffix} пропущен."
    rows = load_rows(path)
    result = detect(rows, path)
    if result is None:
        return [], f"{path.name}: {remember_unknown(path, rows)}"
    items, note = result
    return items, f"{path.name}: {note}"


def write_table(rows, destination: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "Цены"
    sheet.append(COLUMNS)
    for row in rows:
        sheet.append([row[column] for column in COLUMNS])
    for cell in sheet["B"][1:]:
        cell.number_format = "@"
    sheet.column_dimensions["C"].width = 60
    destination.parent.mkdir(parents=True, exist_ok=True)
    book.save(destination)


def load_imported() -> dict:
    path = DATA / "imported.json"
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def convert_folders(folders: list[Path], existing_rows: list[dict] | None = None):
    """Разбирает новые файлы. Тот же хеш второй раз не записывает."""
    imported = load_imported()
    notes = []
    fresh = []
    seen_now = {}
    for folder in folders:
        if not folder.exists():
            notes.append(f"Папка не найдена: {folder}")
            continue
        files = sorted(path for path in folder.iterdir() if path.is_file())
        for path in files:
            digest = file_hash(path)
            if digest in seen_now:
                notes.append(f"{path.name}: тот же файл, что {seen_now[digest]}. Пропущен.")
                continue
            seen_now[digest] = path.name
            if digest in imported:
                notes.append(f"{path.name}: уже загружен ({imported[digest]}). Повтор не записан.")
                continue
            items, note = convert_file(path)
            imported[digest] = path.name
            fresh.extend(items)
            notes.append(note)
    (DATA / "imported.json").write_text(json.dumps(imported, ensure_ascii=False, indent=2), encoding="utf-8")
    all_rows = list(existing_rows or []) + fresh
    return all_rows, notes, fresh


def main() -> int:
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(r"C:\Users\1\Desktop\Прайсы")
    all_rows, notes, fresh = convert_folders([folder])
    destination = OUT / "prices.xlsx"
    write_table(all_rows, destination)
    log_path = OUT / "prices.log"
    log_path.write_text("\n".join(notes) + f"\n\nНовых строк: {len(fresh)}. Всего: {len(all_rows)}\n", encoding="utf-8")
    print(f"Готово: {destination}")
    print(f"Новых строк: {len(fresh)}. Всего: {len(all_rows)}")
    print(f"Лог: {log_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
