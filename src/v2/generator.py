"""Генератор чек-листов v2 — сохранение оформления шаблона (zip-merge).

openpyxl используется только для расчёта значений и структуры листов
(вставка строк, формулы). Финальный xlsx собирается из zip шаблона:
``styles.xml``, drawings, printerSettings и атрибуты ``s`` ячеек
остаются из шаблона — как после сохранения файла в Excel.
"""

from __future__ import annotations

import io
import re
import shutil
import zipfile
from datetime import date, datetime
from pathlib import Path

import openpyxl

from src.generator import (
    ACT_OBJECT_END,
    ACT_OBJECT_START,
    CHECK_OBJECT_TEMPLATE_ROW,
    CHECK_TAIL_END,
    CHECK_TAIL_START,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_TEMPLATE,
    MAX_OBJECTS_PER_FILE,
    ChecklistError,
    _merge_content_types,
    _output_path,
    _with_copy_suffix,
    build_generation_plan,
    collect_objects,
    copy_people_to_output,
    eligible_inv_numbers,
    fill_workbook,
    schedule_processing_order,
)
from src.loaders import SourceData, load_all

_CELL_TAG = re.compile(
    r'(<c r=")([A-Z]+)(\d+)(")((?:(?!/>)[^>])*)(?:/>|(>.*?</c>))',
    re.DOTALL,
)


def _col_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def _template_cell_styles(sheet_xml: str) -> dict[str, str]:
    """Адрес ячейки → индекс стиля ``s`` из XML шаблона."""
    styles: dict[str, str] = {}
    for m in _CELL_TAG.finditer(sheet_xml):
        addr = f"{m.group(2)}{m.group(3)}"
        attrs = m.group(5)
        sm = re.search(r'\ss="(\d+)"', attrs)
        if sm:
            styles[addr] = sm.group(1)
    return styles


def _worksheet_opening_tag(sheet_xml: str) -> str | None:
    m = re.match(r"(<worksheet\b[^>]+>)", sheet_xml)
    return m.group(1) if m else None


def _check_template_source_row(count: int, row: int) -> int | None:
    """Строка листа «чек-лист» → номер строки-образца в шаблоне."""
    offset = count - 1
    obj_end = CHECK_OBJECT_TEMPLATE_ROW + count - 1
    if CHECK_OBJECT_TEMPLATE_ROW <= row <= obj_end:
        return CHECK_OBJECT_TEMPLATE_ROW
    tail_start = CHECK_TAIL_START + offset
    tail_end = CHECK_TAIL_END + offset
    if tail_start <= row <= tail_end:
        return row - offset
    spacer_begin = obj_end + 1
    spacer_end = tail_start - 1
    if spacer_begin <= row <= spacer_end:
        return CHECK_OBJECT_TEMPLATE_ROW + 1 + (row - obj_end - 1)
    return None


def _template_row_cells(sheet_xml: str, row: int) -> list[tuple[str, str]]:
    """Ячейки строки шаблона: (колонка, полный XML тег ``<c>``)."""
    cells: list[tuple[str, str]] = []
    for m in _CELL_TAG.finditer(sheet_xml):
        if int(m.group(3)) == row:
            cells.append((m.group(2), m.group(0)))
    return cells


def _remap_cell_row(cell_xml: str, dst_row: int) -> str:
    return re.sub(
        r'r="([A-Z]+)\d+"',
        lambda m: f'r="{m.group(1)}{dst_row}"',
        cell_xml,
    )


def _cells_in_row_body(row_body: str) -> dict[str, str]:
    """Колонка → XML тег ``<c>`` внутри ``<row>``."""
    return {m.group(2): m.group(0) for m in _CELL_TAG.finditer(row_body)}


def _sorted_row_body(cells: dict[str, str]) -> str:
    return "".join(cells[col] for col in sorted(cells, key=_col_index))


def _replace_row_body(sheet_xml: str, row: int, new_body: str) -> str:
    def repl(match: re.Match[str]) -> str:
        return f"{match.group(1)}{new_body}{match.group(3)}"

    return re.sub(
        rf'(<row r="{row}"(?:\s[^>]*)?>)(.*?)(</row>)',
        repl,
        sheet_xml,
        count=1,
        flags=re.DOTALL,
    )


def _supplement_check_sheet_cells(
    filled_xml: str,
    template_xml: str,
    *,
    object_count: int,
) -> str:
    """Добавляет из шаблона ячейки с границами, которые openpyxl не сохраняет."""
    offset = object_count - 1
    tail_end = CHECK_TAIL_END + offset
    patched = filled_xml
    for dst_row in range(CHECK_OBJECT_TEMPLATE_ROW, tail_end + 1):
        src_row = _check_template_source_row(object_count, dst_row)
        if src_row is None:
            continue
        row_match = re.search(
            rf'(<row r="{dst_row}"(?:\s[^>]*)?>)(.*?)(</row>)',
            patched,
            re.DOTALL,
        )
        if not row_match:
            continue
        merged = _cells_in_row_body(row_match.group(2))
        for col, cell_xml in _template_row_cells(template_xml, src_row):
            if col not in merged:
                merged[col] = _remap_cell_row(cell_xml, dst_row)
        patched = _replace_row_body(patched, dst_row, _sorted_row_body(merged))
    return patched


def _apply_template_styles(
    filled_xml: str,
    template_xml: str,
    *,
    fallback_row: int | None = None,
    object_row_range: tuple[int, int] | None = None,
    object_style_row: int | None = None,
    check_object_count: int | None = None,
) -> str:
    """Подставляет индексы стилей из шаблона в лист, сохранённый openpyxl."""
    tpl_styles = _template_cell_styles(template_xml)

    def style_for(col: str, row: int) -> str | None:
        if check_object_count is not None:
            src_row = _check_template_source_row(check_object_count, row)
            if src_row is not None:
                return tpl_styles.get(f"{col}{src_row}")
        elif object_row_range and object_style_row is not None:
            row_from, row_to = object_row_range
            if row_from <= row <= row_to:
                style = tpl_styles.get(f"{col}{object_style_row}")
                if style is not None:
                    return style
        addr = f"{col}{row}"
        if addr in tpl_styles:
            return tpl_styles[addr]
        if fallback_row is not None:
            return tpl_styles.get(f"{col}{fallback_row}")
        return None

    def repl(match: re.Match[str]) -> str:
        prefix, col, row_s, _quote, attrs, body = match.groups()
        row = int(row_s)
        style = style_for(col, row)
        if style is None:
            return match.group(0)
        attrs = re.sub(r'\ss="\d+"', "", attrs)
        if body:
            return f'{prefix}{col}{row_s}"{attrs} s="{style}"{body}'
        return f'{prefix}{col}{row_s}"{attrs} s="{style}"/>'

    return _CELL_TAG.sub(repl, filled_xml)


def _apply_template_worksheet_root(filled_xml: str, template_xml: str) -> str:
    """Корневой ``<worksheet>`` с namespace'ами как в шаблоне."""
    opening = _worksheet_opening_tag(template_xml)
    if not opening:
        return filled_xml
    return re.sub(r"<worksheet\b[^>]+>", opening, filled_xml, count=1)


def _patch_sheet_from_template(
    filled_xml: str,
    template_xml: str,
    *,
    fallback_row: int | None = None,
    object_row_range: tuple[int, int] | None = None,
    object_style_row: int | None = None,
    check_object_count: int | None = None,
) -> bytes:
    patched = _apply_template_worksheet_root(filled_xml, template_xml)
    if check_object_count is not None:
        patched = _supplement_check_sheet_cells(
            patched,
            template_xml,
            object_count=check_object_count,
        )
    patched = _apply_template_styles(
        patched,
        template_xml,
        fallback_row=fallback_row,
        object_row_range=object_row_range,
        object_style_row=object_style_row,
        check_object_count=check_object_count,
    )
    return patched.encode("utf-8")


def _merge_filled_into_template(
    template_path: Path,
    filled_buffer: io.BytesIO,
    *,
    object_count: int,
) -> dict[str, bytes]:
    """Собирает итоговый zip: оформление шаблона + данные из openpyxl."""
    with zipfile.ZipFile(template_path) as ztpl, zipfile.ZipFile(filled_buffer) as zfilled:
        tpl = {info.filename: ztpl.read(info.filename) for info in ztpl.infolist()}
        filled_names = {info.filename for info in zfilled.infolist()}
        filled = {name: zfilled.read(name) for name in filled_names}

    out = dict(tpl)

    sheet_options = {
        "xl/worksheets/sheet1.xml": {
            "object_row_range": (ACT_OBJECT_START, ACT_OBJECT_END),
            "object_style_row": ACT_OBJECT_START,
        },
        "xl/worksheets/sheet2.xml": {
            "check_object_count": object_count,
        },
    }
    for sheet_name, opts in sheet_options.items():
        if sheet_name not in filled:
            continue
        out[sheet_name] = _patch_sheet_from_template(
            filled[sheet_name].decode("utf-8"),
            tpl[sheet_name].decode("utf-8"),
            **opts,
        )

    for name in ("xl/workbook.xml", "xl/sharedStrings.xml", "xl/calcChain.xml"):
        if name in filled:
            out[name] = filled[name]

    if "[Content_Types].xml" in filled:
        out["[Content_Types].xml"] = _merge_content_types(
            tpl["[Content_Types].xml"].decode("utf-8"),
            filled["[Content_Types].xml"].decode("utf-8"),
        ).encode("utf-8")

    return out


def _write_checklist_file(
    template_path: Path,
    desired_path: Path,
    *,
    fill_fn,
    object_count: int,
) -> Path:
    """Копирует шаблон, заполняет через openpyxl, собирает xlsx без перезаписи ``styles.xml``."""
    copy_index = 0
    while copy_index < 100:
        output_path = _with_copy_suffix(desired_path, copy_index)
        try:
            if copy_index > 0:
                print(
                    f"Файл занят: {desired_path.name} — сохраняем как {output_path.name}",
                )
            shutil.copy2(template_path, output_path)
            wb = openpyxl.load_workbook(output_path)
            fill_fn(wb)

            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            package = _merge_filled_into_template(
                template_path,
                buffer,
                object_count=object_count,
            )
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for name, data in package.items():
                    zout.writestr(name, data)
            return output_path
        except PermissionError:
            copy_index += 1
    raise ChecklistError(f"Не удалось сохранить файл (занят): {desired_path.name}")


def generate_checklist(
    inv: int,
    *,
    output_dir: Path | None = None,
    template: Path | None = None,
    act_number: int | str | None = None,
    act_date: date | datetime | None = None,
    data: SourceData | None = None,
    generation_plan: dict | None = None,
    people_file: Path | None = None,
) -> list[Path]:
    """Создаёт чек-лист(ы) с сохранением оформления шаблона."""
    template_path = template or DEFAULT_TEMPLATE
    if not template_path.is_file():
        raise ChecklistError(f"Шаблон не найден: {template_path}")

    out_dir = output_dir or DEFAULT_OUTPUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    copy_people_to_output(out_dir, source=people_file)

    source_data = data or load_all()

    if act_number is None or act_date is None:
        if generation_plan is None:
            inv_order = schedule_processing_order(
                source_data.schedule,
                eligible_inv_numbers(source_data),
            )
            generation_plan = build_generation_plan(
                source_data,
                inv_order,
                act_override=act_number,
                date_override=act_date if isinstance(act_date, date) else None,
                people_file=people_file,
            )
        meta = generation_plan.get(inv)
        if meta is None:
            raise ChecklistError(f"Инв. № {inv} отсутствует в плане генерации")
        if act_number is None:
            act_number = meta.act_number
        if act_date is None:
            act_date = meta.act_date

    all_objects = collect_objects(source_data, inv)
    total_objects = len(all_objects)

    if total_objects > MAX_OBJECTS_PER_FILE:
        total_parts = (total_objects + MAX_OBJECTS_PER_FILE - 1) // MAX_OBJECTS_PER_FILE
        print(
            f"Инв. № {inv}: найдено {total_objects} объектов — "
            f"будет создано {total_parts} файлов (до {MAX_OBJECTS_PER_FILE} объектов в каждом)"
        )
    else:
        total_parts = 1

    created: list[Path] = []

    for part in range(1, total_parts + 1):
        start = (part - 1) * MAX_OBJECTS_PER_FILE
        chunk = all_objects[start : start + MAX_OBJECTS_PER_FILE]
        desired_path = _output_path(
            out_dir,
            inv,
            part,
            total_parts,
            act_number=act_number,
            act_date=act_date,
        )
        output_path = _write_checklist_file(
            template_path,
            desired_path,
            object_count=len(chunk),
            fill_fn=lambda wb: fill_workbook(
                wb,
                chunk,
                act_number=act_number,
                act_date=act_date,
                part=part,
                total_parts=total_parts,
                global_offset=start,
                total_objects=total_objects,
            ),
        )
        created.append(output_path)

    return created
