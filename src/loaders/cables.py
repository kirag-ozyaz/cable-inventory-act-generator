"""Загрузка списка н/в кабелей [3] — пять районов в одну таблицу."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from src.loaders._utils import drop_empty_rows, find_header_row, normalize_column_name
from src.loaders.paths import CABLE_DISTRICT_SHEETS, cables_path

# Единые имена столбцов после нормализации заголовков листов
CABLE_COLUMNS: dict[str, str] = {
    "наименование": "наименование",
    "инв №": "инв_номер",
    "инв номер": "инв_номер",
    "тп": "тп",
    "№ рубильника": "рубильник",
    "потребители": "потребители",
    "почт.адрес": "почт_адрес",
    "принадл": "принадл",
    "марка": "марка",
    "сеч": "сечение",
    "длина": "длина",
    "год ввода": "год_ввода",
    "поврежден": "поврежден",
    "временная запитка": "временная_запитка",
}


def _canonical_columns(raw_columns: pd.Index) -> list[str]:
    result: list[str] = []
    used: set[str] = set()
    for col in raw_columns:
        key = normalize_column_name(col)
        name = CABLE_COLUMNS.get(key)
        if name is None:
            if key:
                name = key.replace(" ", "_")
            else:
                name = "unnamed"
        base = name
        n = 1
        while name in used:
            n += 1
            name = f"{base}_{n}"
        used.add(name)
        result.append(name)
    return result


def _load_district_sheet(file_path: Path, district: int, sheet_name: str) -> pd.DataFrame:
    preview = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        nrows=6,
        engine="openpyxl",
    )
    header_row = find_header_row(preview)

    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=header_row,
        engine="openpyxl",
    )
    df.columns = _canonical_columns(df.columns)
    df = drop_empty_rows(df)

    if "инв_номер" not in df.columns:
        raise KeyError(f"Лист {sheet_name!r}: нет столбца инв. номера")

    df["инв_номер"] = pd.to_numeric(df["инв_номер"], errors="coerce")
    df = df[df["инв_номер"].notna()].copy()
    df["инв_номер"] = df["инв_номер"].astype("Int64")
    df["сетевой_район"] = district

    # Наименование кабеля — первый столбец листа (часто без заголовка)
    first_col = df.columns[0]
    if first_col not in ("наименование", "инв_номер"):
        df = df.rename(columns={first_col: "наименование"})

    return df


def load_cables(path: Path | None = None) -> pd.DataFrame:
    """
    Список н_в кабелей … xlsx: листы сетевых районов 1–5 → одна таблица.

    Добавляется колонка ``сетевой_район`` (1–5). Лист «списанные» не загружается.
    """
    file_path = path or cables_path()
    parts: list[pd.DataFrame] = []

    for district, sheet_name in CABLE_DISTRICT_SHEETS:
        parts.append(_load_district_sheet(file_path, district, sheet_name))

    if not parts:
        return pd.DataFrame()

    combined = pd.concat(parts, ignore_index=True, sort=False)
    combined = combined.dropna(axis=1, how="all")
    combined = combined.loc[:, ~combined.columns.str.startswith("unnamed")]

    # сетевой_район — сразу после инв_номер
    cols = list(combined.columns)
    cols.remove("сетевой_район")
    if "инв_номер" in cols:
        inv_idx = cols.index("инв_номер") + 1
        cols = cols[:inv_idx] + ["сетевой_район"] + cols[inv_idx:]
    else:
        cols = ["сетевой_район"] + cols

    return combined[cols].reset_index(drop=True)
