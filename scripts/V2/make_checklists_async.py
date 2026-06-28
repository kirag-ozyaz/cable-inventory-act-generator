"""Асинхронное создание чек-листов с сохранением оформления шаблона (zip-merge).

Как ``make_checklists_async.py``, но финальный xlsx собирается из zip шаблона:
``styles.xml``, drawings и printerSettings не перезаписываются openpyxl.

Примеры:
    python scripts/v2/make_checklists_async.py
    python scripts/v2/make_checklists_async.py --inv 7260 425
    python scripts/v2/make_checklists_async.py --inv 7260 --act 4167 --date 20.06.2026
    python scripts/v2/make_checklists_async.py --workers 8 -o output/
"""

from __future__ import annotations

import argparse
import asyncio
import contextlib
import io
import sys
import warnings
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

# openpyxl шумит этим предупреждением при каждой загрузке шаблона (по разу
# на инв. №), из-за чего ломается прогресс-бар. Глушим только его.
warnings.filterwarnings(
    "ignore",
    message="DrawingML support is incomplete",
    category=UserWarning,
    module="openpyxl",
)

from src.generator import (  # noqa: E402
    DEFAULT_TEMPLATE,
    INV_COL,
    ChecklistError,
    _schedule_column,
    extract_tp_name,
)
from src.v2 import generate_checklist  # noqa: E402
from src.loaders import SourceData, load_all  # noqa: E402
from src.loaders.paths import cables_path, schedule_path, substations_path  # noqa: E402


@dataclass
class _SourceSummary:
    """Краткая сводка по одному загруженному справочнику для вывода preflight."""

    label: str
    path: Path
    rows: int
    inv_count: int | None = None
    note: str | None = None


def _parse_date(value: str):
    """Парсит дату акта из аргумента CLI (``ГГГГ-ММ-ДД`` или ``ДД.ММ.ГГГГ``)."""
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise argparse.ArgumentTypeError(f"Неверная дата: {value!r} (ожидается ГГГГ-ММ-ДД или ДД.ММ.ГГГГ)")


def _inv_numbers(series) -> set[int]:
    """Уникальные целочисленные инв. № из столбца pandas (без NaN)."""
    numeric = series.dropna()
    return {int(v) for v in numeric}


def _schedule_inv(data: SourceData) -> set[int]:
    """Множество инв. № из графика инвентаризации."""
    if INV_COL not in data.schedule.columns:
        raise ChecklistError(f"В графике нет столбца {INV_COL!r}")
    return _inv_numbers(data.schedule[INV_COL])


def _cables_inv(data: SourceData) -> set[int]:
    """Множество инв. № из списка н/в кабелей."""
    if "инв_номер" not in data.cables.columns:
        raise ChecklistError("В списке кабелей нет столбца 'инв_номер'")
    return _inv_numbers(data.cables["инв_номер"])


def _plan_jobs(
    data: SourceData,
    requested: list[int] | None,
) -> tuple[list[int], list[int], list[int], list[int]]:
    """Формирует план обработки и списки расхождений между справочниками.

    Returns:
        ``(к обработке, в списке без графика, в графике без списка, нигде)``.
        Последний элемент заполняется только при явном ``--inv``.
    """
    schedule = _schedule_inv(data)
    cables = _cables_inv(data)

    if requested is None:
        to_process = sorted(schedule & cables)
        missing_schedule = sorted(cables - schedule)
        missing_cables = sorted(schedule - cables)
        return to_process, missing_schedule, missing_cables, []

    wanted = set(requested)
    to_process = sorted(wanted & schedule & cables)
    missing_schedule = sorted(wanted & cables - schedule)
    missing_cables = sorted(wanted & schedule - cables)
    not_found = sorted(wanted - schedule - cables)
    return to_process, missing_schedule, missing_cables, not_found


def _tp_names_for_processing(data: SourceData, inv_numbers: list[int]) -> set[str]:
    """Собирает коды ТП/РП из наименований кабелей по инв. № к обработке."""
    if not inv_numbers:
        return set()

    inv_set = set(inv_numbers)
    tp_names: set[str] = set()

    name_col = _schedule_column(data.schedule, "диспетчер")
    schedule_rows = data.schedule[data.schedule[INV_COL].isin(inv_set)]
    for value in schedule_rows[name_col].astype(str):
        tp = extract_tp_name(value)
        if tp:
            tp_names.add(tp)

    cable_rows = data.cables[data.cables["инв_номер"].isin(inv_set)]
    if "наименование" in cable_rows.columns:
        for value in cable_rows["наименование"].dropna().astype(str):
            tp = extract_tp_name(value)
            if tp:
                tp_names.add(tp)

    return tp_names


def _substation_issues(
    data: SourceData,
    tp_names: set[str],
) -> tuple[list[str], list[str]]:
    """Сверяет ТП/РП со справочником подстанций.

    Returns:
        ``(нет в справочнике, есть в справочнике без координат)``.
    """
    if not tp_names:
        return [], []

    substations = data.substations
    names = substations["Name"].astype(str)
    has_coords = "широта" in substations.columns and "долгота" in substations.columns

    not_in_catalog: list[str] = []
    no_coords: list[str] = []

    for tp in sorted(tp_names):
        row = substations[names.str.upper() == tp.upper()]
        if row.empty:
            row = substations[names.str.contains(tp, case=False, na=False)]
        if row.empty:
            not_in_catalog.append(tp)
            continue
        if not has_coords:
            no_coords.append(tp)
            continue
        rec = row.iloc[0]
        if pd.isna(rec.get("широта")) or pd.isna(rec.get("долгота")):
            no_coords.append(tp)

    return not_in_catalog, no_coords


def _format_inv_discrepancies(
    *,
    missing_schedule: list[int],
    missing_cables: list[int],
    not_found: list[int],
) -> list[str]:
    """Строки отчёта о несогласованных инв. № между справочниками."""
    if not missing_schedule and not missing_cables and not not_found:
        return []

    lines = [
        "Отчёт о необработанных инв. №:",
        "  Для перечисленных номеров чек-лист создать нельзя — "
        "нет согласованных данных в обоих справочниках.",
    ]

    if missing_schedule:
        nums = ", ".join(str(n) for n in missing_schedule)
        lines.append(
            f"  Есть в «Списке н/в кабелей», но нет в «Графике инвентаризации» "
            f"({len(missing_schedule)}): {nums}"
        )

    if missing_cables:
        nums = ", ".join(str(n) for n in missing_cables)
        lines.append(
            f"  Есть в «Графике инвентаризации», но нет в «Списке н/в кабелей» "
            f"({len(missing_cables)}): {nums}"
        )

    if not_found:
        nums = ", ".join(str(n) for n in not_found)
        lines.append(
            f"  Нет ни в графике, ни в списке кабелей "
            f"(указаны в --inv, в справочниках не найдены) ({len(not_found)}): {nums}"
        )

    return lines


def _format_substation_discrepancies(
    *,
    not_in_catalog: list[str],
    no_coords: list[str],
) -> list[str]:
    """Строки отчёта о ТП/РП без координат."""
    if not not_in_catalog and not no_coords:
        return []

    lines = [
        "Координаты начала КЛ (справочник подстанций):",
        "  Чек-листы создаются в любом случае; "
        "для перечисленных ТП/РП координаты в чек-листе останутся пустыми.",
    ]

    if not_in_catalog:
        items = ", ".join(not_in_catalog)
        lines.append(
            f"  Упоминаются в кабелях, но отсутствуют в "
            f"«подстанции с координатами» ({len(not_in_catalog)}): {items}"
        )

    if no_coords:
        items = ", ".join(no_coords)
        lines.append(
            f"  Есть в справочнике, но широта/долгота не заполнены "
            f"({len(no_coords)}): {items}"
        )

    return lines


def _write_discrepancies_report(
    output_dir: Path,
    *,
    missing_schedule: list[int],
    missing_cables: list[int],
    not_found: list[int],
    tp_not_in_catalog: list[str],
    tp_no_coords: list[str],
) -> Path | None:
    """Сохраняет отчёт о расхождениях в ``output_dir/расхождения.txt``."""
    sections: list[str] = []

    inv_lines = _format_inv_discrepancies(
        missing_schedule=missing_schedule,
        missing_cables=missing_cables,
        not_found=not_found,
    )
    if inv_lines:
        sections.append("\n".join(inv_lines))

    sub_lines = _format_substation_discrepancies(
        not_in_catalog=tp_not_in_catalog,
        no_coords=tp_no_coords,
    )
    if sub_lines:
        sections.append("\n".join(sub_lines))

    if not sections:
        return None

    report_path = output_dir / "расхождения.txt"
    report_path.write_text("\n\n".join(sections) + "\n", encoding="utf-8")
    return report_path


def _print_substation_report(
    *,
    not_in_catalog: list[str],
    no_coords: list[str],
) -> bool:
    """Печатает отчёт о ТП/РП без координат; ``True``, если отчёт был выведен."""
    lines = _format_substation_discrepancies(
        not_in_catalog=not_in_catalog,
        no_coords=no_coords,
    )
    if not lines:
        return False

    print()
    for line in lines:
        print(line)

    return True


def _require_file(label: str, path: Path) -> None:
    """Проверяет существование файла; иначе ``ChecklistError``."""
    if not path.is_file():
        raise ChecklistError(f"{label}: файл не найден: {path}")


def _preflight(
    requested: list[int] | None,
) -> tuple[
    SourceData,
    list[int],
    list[int],
    list[int],
    list[int],
    list[str],
    list[str],
]:
    """Preflight: загрузка справочников, проверка структуры, отчёты о расхождениях.

    Returns:
        ``(data, к обработке, в списке без графика, в графике без списка, нигде,
        ТП/РП нет в справочнике, ТП/РП без координат)``.
    """
    print("Проверка исходных данных...")

    schedule_file = schedule_path()
    cables_file = cables_path()
    substations_file = substations_path()

    _require_file("График инвентаризации", schedule_file)
    _require_file("Список кабелей", cables_file)
    _require_file("Подстанции", substations_file)
    _require_file("Шаблон", DEFAULT_TEMPLATE)

    print("  Загрузка справочников...")
    data = load_all()

    summaries: list[_SourceSummary] = []

    schedule = data.schedule
    if schedule.empty:
        raise ChecklistError("График инвентаризации: нет строк с инв. №")
    _schedule_column(schedule, "диспетчер")
    _schedule_column(schedule, "адрес", "местонахожд")
    schedule_inv = _schedule_inv(data)
    summaries.append(
        _SourceSummary(
            "График инвентаризации",
            schedule_file,
            len(schedule),
            len(schedule_inv),
        )
    )

    cables = data.cables
    if cables.empty:
        raise ChecklistError("Список кабелей: нет данных")
    for col in ("инв_номер", "марка", "сечение", "длина"):
        if col not in cables.columns:
            raise ChecklistError(f"Список кабелей: нет столбца {col!r}")
    cables_inv = _cables_inv(data)
    summaries.append(
        _SourceSummary(
            "Список кабелей",
            cables_file,
            len(cables),
            len(cables_inv),
        )
    )

    substations = data.substations
    if substations.empty:
        raise ChecklistError("Подстанции: нет данных")
    if "Name" not in substations.columns:
        raise ChecklistError("Подстанции: нет столбца Name")

    if "широта" in substations.columns and "долгота" in substations.columns:
        coords_filled = int(
            (substations["широта"].notna() & substations["долгота"].notna()).sum()
        )
        substations_note = f"{coords_filled} с координатами"
    else:
        substations_note = "столбцы координат отсутствуют"

    summaries.append(
        _SourceSummary(
            "Подстанции",
            substations_file,
            len(substations),
            note=substations_note,
        )
    )

    for item in summaries:
        extra = ""
        if item.inv_count is not None:
            extra += f", {item.inv_count} инв. №"
        if item.note:
            extra += f", {item.note}"
        print(f"  OK {item.label}: {_show_path(item.path)} — {item.rows} строк{extra}")
    print(f"  OK Шаблон: {_show_path(DEFAULT_TEMPLATE)}")

    to_process, missing_schedule, missing_cables, not_found = _plan_jobs(data, requested)

    print()
    print("Сопоставление инв. №:")
    print(f"  К обработке: {len(to_process)}")
    _print_missing_report(
        missing_schedule=missing_schedule,
        missing_cables=missing_cables,
        not_found=not_found,
    )

    tp_names = _tp_names_for_processing(data, to_process)
    tp_not_in_catalog, tp_no_coords = _substation_issues(data, tp_names)
    _print_substation_report(
        not_in_catalog=tp_not_in_catalog,
        no_coords=tp_no_coords,
    )

    return (
        data,
        to_process,
        missing_schedule,
        missing_cables,
        not_found,
        tp_not_in_catalog,
        tp_no_coords,
    )


def _show_path(path: Path) -> Path:
    """Путь относительно корня проекта для компактного вывода."""
    try:
        return path.relative_to(ROOT)
    except ValueError:
        return path


def _render_progress(done: int, total: int, *, width: int = 30) -> None:
    """Однострочный прогресс-бар в stderr (перерисовывается через \\r)."""
    if total <= 0:
        return
    filled = int(width * done / total)
    bar = "#" * filled + "-" * (width - filled)
    pct = done / total * 100
    end = "\n" if done >= total else ""
    print(f"\r[{bar}] {done}/{total} ({pct:5.1f}%)", end=end, file=sys.stderr, flush=True)


def _print_generator_notes(text: str) -> None:
    """Компактно печатает сообщения генератора (с группировкой повторов)."""
    lines = [line for line in text.splitlines() if line.strip()]
    if not lines:
        return

    print()
    print("Замечания генератора:")
    counts = Counter(lines)
    for line, count in counts.items():
        suffix = f" (x{count})" if count > 1 else ""
        print(f"  {line}{suffix}")


def _print_missing_report(
    *,
    missing_schedule: list[int],
    missing_cables: list[int],
    not_found: list[int],
) -> bool:
    """Печатает отчёт о инв. № без пары в обоих справочниках; ``True``, если выведен."""
    lines = _format_inv_discrepancies(
        missing_schedule=missing_schedule,
        missing_cables=missing_cables,
        not_found=not_found,
    )
    if not lines:
        return False

    print()
    for line in lines:
        print(line)

    return True


async def _generate_one(
    inv: int,
    *,
    executor: ThreadPoolExecutor,
    output_dir: Path,
    act_number: int | None,
    act_date,
    data: SourceData,
) -> tuple[int, list[Path] | BaseException]:
    """Генерирует чек-лист для одного инв. № в пуле потоков.

    Returns:
        ``(инв_№, пути_файлов)`` или ``(инв_№, исключение)`` при ошибке.
    """
    loop = asyncio.get_running_loop()

    def _run() -> list[Path]:
        return generate_checklist(
            inv,
            output_dir=output_dir,
            act_number=act_number,
            act_date=act_date,
            data=data,
        )

    try:
        paths = await loop.run_in_executor(executor, _run)
    except BaseException as exc:
        return inv, exc
    return inv, paths


async def _run_batch(
    inv_numbers: list[int],
    *,
    workers: int,
    output_dir: Path,
    act_number: int | None,
    act_date,
    data: SourceData,
    show_progress: bool,
) -> tuple[list[tuple[int, list[Path] | BaseException]], bool]:
    """Параллельно создаёт чек-листы для списка инв. №.

    Returns:
        ``(результаты, прервано_пользователем)``.
    """
    if not inv_numbers:
        return [], False

    total = len(inv_numbers)
    worker_count = max(1, min(workers, total))
    results: list[tuple[int, list[Path] | BaseException]] = []
    interrupted = False

    executor = ThreadPoolExecutor(max_workers=worker_count)
    try:
        tasks = {
            asyncio.create_task(
                _generate_one(
                    inv,
                    executor=executor,
                    output_dir=output_dir,
                    act_number=act_number,
                    act_date=act_date,
                    data=data,
                )
            ): inv
            for inv in inv_numbers
        }
        pending = set(tasks.keys())

        if show_progress:
            _render_progress(0, total)

        done_count = 0
        while pending:
            try:
                finished, pending = await asyncio.wait(
                    pending,
                    return_when=asyncio.FIRST_COMPLETED,
                )
            except KeyboardInterrupt:
                interrupted = True
                break

            for task in finished:
                try:
                    results.append(await task)
                except asyncio.CancelledError:
                    continue
                except BaseException as exc:
                    results.append((tasks[task], exc))
                done_count += 1
                if show_progress:
                    _render_progress(done_count, total)

        if interrupted:
            for task in pending:
                task.cancel()
            if pending:
                await asyncio.gather(*pending, return_exceptions=True)
    finally:
        if interrupted:
            executor.shutdown(wait=False, cancel_futures=True)
        else:
            executor.shutdown(wait=True)

    return results, interrupted


def main(argv: list[str] | None = None) -> int:
    """Точка входа CLI: preflight → асинхронная генерация → итоговый отчёт."""
    parser = argparse.ArgumentParser(
        description="Асинхронно создать чек-листы (zip-merge, оформление как в шаблоне)",
    )
    parser.add_argument(
        "--inv",
        nargs="*",
        type=int,
        metavar="N",
        help="инвентарные номера; если не указаны — все из графика",
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
    parser.add_argument(
        "--workers",
        type=int,
        default=10,
        help="число параллельных задач (по умолчанию 10)",
    )
    parser.add_argument(
        "--no-progress",
        action="store_true",
        help="не показывать прогресс-бар",
    )
    args = parser.parse_args(argv)
    args.output_dir.mkdir(parents=True, exist_ok=True)

    try:
        (
            data,
            to_process,
            missing_schedule,
            missing_cables,
            not_found,
            tp_not_in_catalog,
            tp_no_coords,
        ) = _preflight(args.inv)
    except (ChecklistError, FileNotFoundError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1

    report_path = _write_discrepancies_report(
        args.output_dir,
        missing_schedule=missing_schedule,
        missing_cables=missing_cables,
        not_found=not_found,
        tp_not_in_catalog=tp_not_in_catalog,
        tp_no_coords=tp_no_coords,
    )
    if report_path is not None:
        print(f"Отчёт о расхождениях: {_show_path(report_path)}")

    if not to_process:
        print("Нет инвентарных номеров для обработки.")
        return 1

    print()
    print(f"Генерация: {len(to_process)} номер(ов), workers={args.workers}")

    show_progress = not args.no_progress and sys.stderr.isatty()

    generation_errors: list[tuple[int, BaseException]] = []
    created_paths: list[Path] = []
    interrupted = False
    results: list[tuple[int, list[Path] | BaseException]] = []

    notes_buffer = io.StringIO()

    if to_process:
        try:
            with contextlib.redirect_stdout(notes_buffer):
                results, interrupted = asyncio.run(
                    _run_batch(
                        to_process,
                        workers=args.workers,
                        output_dir=args.output_dir,
                        act_number=args.act,
                        act_date=args.date,
                        data=data,
                        show_progress=show_progress,
                    )
                )
        except KeyboardInterrupt:
            interrupted = True
            results = []

        if interrupted:
            if show_progress:
                print(file=sys.stderr)
            print("\nПрервано пользователем (Ctrl+C).", file=sys.stderr)

        for inv, outcome in sorted(results, key=lambda item: item[0]):
            if isinstance(outcome, BaseException):
                generation_errors.append((inv, outcome))
                print(f"Ошибка {inv}: {outcome}", file=sys.stderr)
                continue
            created_paths.extend(outcome)

        for path in created_paths:
            print(f"Создан: {_show_path(path)}")

    created_count = len(created_paths)

    _print_generator_notes(notes_buffer.getvalue())

    skipped = len(missing_schedule) + len(missing_cables) + len(not_found)
    print()
    if interrupted:
        remaining = len(to_process) - len({inv for inv, _ in results}) if to_process else 0
        print(
            f"Итого: создано файлов {created_count}, "
            f"ошибок генерации {len(generation_errors)}, "
            f"не завершено {max(remaining, 0)}"
        )
        return 130

    print(
        f"Итого: создано файлов {created_count}, "
        f"ошибок генерации {len(generation_errors)}, "
        f"пропущено из-за расхождений {skipped}"
    )

    if generation_errors or missing_schedule or missing_cables or not_found:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
