"""Загрузка подстанций с координатами [2]."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from src.loaders.paths import substations_path


def load_substations(path: Path | None = None) -> pd.DataFrame:
    """подстанции с координатами.xlsx, лист «Лист1»."""
    file_path = path or substations_path()
    df = pd.read_excel(file_path, sheet_name=0, engine="openpyxl")

    rename = {
        "Unnamed: 11": "широта",
        "Unnamed: 12": "долгота",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})

    return df.reset_index(drop=True)
