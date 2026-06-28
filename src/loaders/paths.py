"""Пути к исходным файлам в каталоге Data/."""

from __future__ import annotations

from pathlib import Path

DATA_DIR = Path(__file__).resolve().parents[2] / "Data"

# Листы списка кабелей [3]: (номер сетевого района, имя вкладки в Excel)
CABLE_DISTRICT_SHEETS: list[tuple[int, str]] = [
    (1, "сет р-он 1"),
    (2, "сет р 2"),
    (3, "сет р-он 3 "),
    (4, "сет р-он 4"),
    (5, "сет р-он 5"),
]


def _find_one(pattern: str) -> Path:
    matches = sorted(DATA_DIR.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"Не найден файл по шаблону {pattern!r} в {DATA_DIR}")
    real = [m for m in matches if not m.name.startswith("~$")]
    if not real:
        raise FileNotFoundError(f"Нет доступных файлов по шаблону {pattern!r} в {DATA_DIR}")
    return real[0]


def schedule_path() -> Path:
    return _find_one("*23062026.xlsx")


def substations_path() -> Path:
    return _find_one("*координат*.xlsx")


def cables_path() -> Path:
    return _find_one("*25062026.xlsx")


def template_path() -> Path:
    return _find_one("*+* 1.xlsx")
