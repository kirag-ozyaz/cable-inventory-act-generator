"""Загрузка графика инвентаризации [1]."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from src.loaders.paths import schedule_path


def load_schedule(path: Path | None = None) -> pd.DataFrame:
    """
    График_инв_УльГЭС … xlsx, лист «Общий график».

    Заголовки — 9-я строка файла (header=8), служебная строка «СтолбецN» отбрасывается.
    """
    file_path = path or schedule_path()
    df = pd.read_excel(file_path, sheet_name=0, header=8, engine="openpyxl")

    inv_col = "Инвентарный номер АО УльГЭС"
    if inv_col not in df.columns:
        raise KeyError(f"В графике нет столбца {inv_col!r}")

    df[inv_col] = pd.to_numeric(df[inv_col], errors="coerce")
    df = df[df[inv_col].notna()].copy()
    df[inv_col] = df[inv_col].astype("Int64")

    return df.reset_index(drop=True)
