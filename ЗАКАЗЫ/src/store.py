"""Хранение справочника, связок, цен и отказов поставщиков."""

import json
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUT = ROOT / "out"
PRICES_PATH = DATA / "prices.json"
CATALOG_PATH = DATA / "catalog.json"
MATCHES_PATH = DATA / "matches.json"
DECISIONS_PATH = DATA / "decisions.json"

UNIT_ALIASES = {
    "шт": "шт",
    "штука": "шт",
    "штук": "шт",
    "упак": "упак",
    "уп": "упак",
    "упаковка": "упак",
    "кг": "кг",
    "л": "л",
}


def load_json(path: Path, default):
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def save_json(path: Path, payload) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_prices() -> list:
    return load_json(PRICES_PATH, [])


def save_prices(rows: list) -> None:
    save_json(PRICES_PATH, rows)


def load_catalog() -> list:
    return load_json(CATALOG_PATH, [])


def save_catalog(rows: list) -> None:
    save_json(CATALOG_PATH, rows)


def load_matches() -> dict:
    return load_json(MATCHES_PATH, {})


def save_matches(matches: dict) -> None:
    save_json(MATCHES_PATH, matches)


def load_decisions() -> dict:
    data = load_json(DECISIONS_PATH, {"confirmed": {}, "absent": {}})
    data.setdefault("confirmed", {})
    data.setdefault("absent", {})
    return data


def save_decisions(data: dict) -> None:
    save_json(DECISIONS_PATH, data)


def offer_key(row: dict) -> str:
    return "|".join([row.get("Поставщик") or "", row.get("Код товара") or "", row.get("Наименование") or ""])


def norm_unit(value: str) -> str:
    text = (value or "").strip().lower().replace(".", "")
    return UNIT_ALIASES.get(text, text)


def parse_row_date(value: str) -> datetime:
    try:
        return datetime.strptime(value or "", "%d.%m.%Y")
    except ValueError:
        return datetime.min


def today_text() -> str:
    return date.today().strftime("%d.%m.%Y")


def unit_problem(catalog_unit: str, offer_unit: str) -> str:
    catalog = norm_unit(catalog_unit)
    offer = norm_unit(offer_unit)
    if catalog and offer and catalog != offer:
        return "единица не совпала"
    if catalog and not offer:
        return "единица не указана"
    return ""
