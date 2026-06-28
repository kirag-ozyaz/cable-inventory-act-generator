"""Вспомогательные функции чтения Excel."""

from __future__ import annotations

import re

import pandas as pd


def normalize_column_name(name: object) -> str:
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    text = str(name).strip().lower().replace("\n", " ")
    return re.sub(r"\s+", " ", text)


def find_header_row(preview: pd.DataFrame, marker: str = "инв") -> int:
    """Номер строки с заголовком (0-based) — ищем ячейку, содержащую marker."""
    for idx, row in preview.iterrows():
        for value in row:
            if marker in normalize_column_name(value):
                return int(idx)
    raise ValueError(f"Строка заголовка с «{marker}» не найдена")


def drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    mask = df.notna().any(axis=1)
    return df.loc[mask].reset_index(drop=True)
