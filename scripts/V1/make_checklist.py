"""Тестовый скрипт: создание чек-листа по инвентарному номеру.

Результат: ``output/чек-лист_{инв}.xlsx`` (до 10 объектов в файле).
При более 10 объектах — ``чек-лист_{инв}_1.xlsx``, ``…_2.xlsx``, …

Примеры:
    python scripts/v1/make_checklist.py
    python scripts/v1/make_checklist.py 7260
    python scripts/v1/make_checklist.py 7260 --act 4167 --date 20.06.2026
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from src.generator import ChecklistError, generate_checklist  # noqa: E402


def _parse_date(value: str):
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise argparse.ArgumentTypeError(f"Неверная дата: {value!r} (ожидается ГГГГ-ММ-ДД или ДД.ММ.ГГГГ)")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Создать чек-лист по инвентарному номеру")
    parser.add_argument(
        "inv",
        nargs="?",
        type=int,
        default=7260,
        help="инвентарный номер (по умолчанию 7260 — эталонный пример)",
    )
    parser.add_argument("--act", type=int, help="номер акта (ячейка E3)")
    parser.add_argument("--date", type=_parse_date, help="дата акта (ячейка G3)")
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=ROOT / "output",
        help="каталог для сохранения (по умолчанию output/)",
    )
    args = parser.parse_args(argv)

    try:
        paths = generate_checklist(
            args.inv,
            output_dir=args.output_dir,
            act_number=args.act,
            act_date=args.date,
        )
    except ChecklistError as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1

    for path in paths:
        try:
            shown = path.relative_to(ROOT)
        except ValueError:
            shown = path
        print(f"Создан: {shown}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
