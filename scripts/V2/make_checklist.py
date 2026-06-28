"""Создание одного чек-листа с сохранением оформления шаблона (zip-merge).

Как ``make_checklist.py``, но без перезаписи ``styles.xml`` openpyxl —
границы и область печати сохраняются как в ``templates/чек-лист_акт_шаблон.xlsx``.

Примеры:
    python scripts/v2/make_checklist.py 7260
    python scripts/v2/make_checklist.py 7260 --act 4167 --date 2026-06-20
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from src.generator import ChecklistError  # noqa: E402
from src.v2 import generate_checklist  # noqa: E402


def _parse_date(value: str):
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise argparse.ArgumentTypeError(f"Неверная дата: {value!r} (ожидается ГГГГ-ММ-ДД или ДД.ММ.ГГГГ)")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Создать чек-лист (zip-merge, оформление как в шаблоне)",
    )
    parser.add_argument(
        "inv",
        nargs="?",
        type=int,
        default=7260,
        help="инвентарный номер (по умолчанию 7260)",
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
