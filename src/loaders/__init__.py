"""Загрузка исходных Excel-файлов в pandas."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from src.loaders.cables import load_cables
from src.loaders.schedule import load_schedule
from src.loaders.substations import load_substations


@dataclass
class SourceData:
    """Три справочника: [1] график, [2] подстанции, [3] кабели."""

    schedule: pd.DataFrame
    substations: pd.DataFrame
    cables: pd.DataFrame


def load_all() -> SourceData:
    return SourceData(
        schedule=load_schedule(),
        substations=load_substations(),
        cables=load_cables(),
    )


__all__ = [
    "SourceData",
    "load_all",
    "load_cables",
    "load_schedule",
    "load_substations",
]
