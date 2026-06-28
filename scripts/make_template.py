"""Создание пустого шаблона из «Чек-Лист + акт 1.xlsx».

Копирует исходный файл (форматирование, два листа, фигуры) и готовит
шаблон для Python-генератора:
  - очищает ячейки-входы под конкретный кабель;
  - удаляет все формулы (в т.ч. внешние ссылки [1]/[2]/[3]);
  - убирает внешние книги и calcChain.

Запуск:
    python scripts/make_template.py
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.loaders.paths import template_path  # noqa: E402

OUTPUT = ROOT / "templates" / "чек-лист_акт_шаблон.xlsx"

SHEET_FILES = ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml")

# Ячейки без формул, но с примерными данными — тоже очищаем.
CELLS_TO_CLEAR: dict[str, list[str]] = {
    "xl/worksheets/sheet1.xml": ["J1", "E3", "G3"],  # инв. №, № акта, дата
    "xl/worksheets/sheet2.xml": ["D9", "H9", "I9"],  # год ввода, координаты конца
}

# Части xlsx, которые не нужны без формул.
SKIP_FILES = (
    "xl/calcChain.xml",
    "xl/metadata.xml",
)

SKIP_PREFIXES = (
    "xl/externalLinks/",
)


def _clear_cell(xml: str, addr: str) -> str:
    """Очищает ячейку, сохраняя стиль (атрибут s)."""
    pattern = re.compile(
        r'<c r="' + re.escape(addr) + r'"([^>]*?)(?:/>|>.*?</c>)',
        re.DOTALL,
    )

    def repl(match: re.Match[str]) -> str:
        attrs = re.sub(r'\s+t="[^"]*"', "", match.group(1))
        return f'<c r="{addr}"{attrs}/>'

    new_xml, n = pattern.subn(repl, xml, count=1)
    if n == 0:
        raise ValueError(f"Ячейка {addr} не найдена в листе")
    return new_xml


def _strip_formulas(xml: str) -> str:
    """Удаляет формулы и кэшированные значения, оставляет пустые ячейки со стилем."""
    # (?:(?!/>)[^>])* — атрибуты, но не самозакрывающийся тег <c .../>
    pattern = re.compile(
        r'<c r="([^"]+)"((?:(?!/>)[^>])*)>(.*?)</c>',
        re.DOTALL,
    )

    def repl(match: re.Match[str]) -> str:
        addr, attrs, inner = match.group(1), match.group(2), match.group(3)
        if "<f" not in inner:
            return match.group(0)
        attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
        attrs = re.sub(r'\s+cm="[^"]*"', "", attrs)
        return f'<c r="{addr}"{attrs}/>'

    return pattern.sub(repl, xml)


def _clean_workbook_xml(xml: str) -> str:
    return re.sub(r"<externalReferences>.*?</externalReferences>", "", xml, flags=re.DOTALL)


def _clean_workbook_rels(xml: str) -> str:
    xml = re.sub(
        r'<Relationship[^>]+Type="[^"]*externalLink"[^>]*/>',
        "",
        xml,
    )
    xml = re.sub(
        r'<Relationship[^>]+Target="calcChain\.xml"[^>]*/>',
        "",
        xml,
    )
    xml = re.sub(
        r'<Relationship[^>]+Target="metadata\.xml"[^>]*/>',
        "",
        xml,
    )
    return xml


def _clean_content_types(xml: str) -> str:
    xml = re.sub(
        r'<Override[^>]+externalLink[^>]*/>',
        "",
        xml,
    )
    xml = re.sub(
        r'<Override[^>]+calcChain[^>]*/>',
        "",
        xml,
    )
    xml = re.sub(
        r'<Override[^>]+sheetMetadata[^>]*/>',
        "",
        xml,
    )
    return xml


def _should_skip(filename: str) -> bool:
    if filename in SKIP_FILES:
        return True
    return any(filename.startswith(prefix) for prefix in SKIP_PREFIXES)


def build_template(source: Path, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(source) as zin:
        items = [(info, zin.read(info.filename)) for info in zin.infolist()]

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            name = info.filename
            if _should_skip(name):
                continue

            if name in SHEET_FILES:
                xml = data.decode("utf-8")
                xml = _strip_formulas(xml)
                for addr in CELLS_TO_CLEAR.get(name, []):
                    xml = _clear_cell(xml, addr)
                data = xml.encode("utf-8")
            elif name == "xl/workbook.xml":
                data = _clean_workbook_xml(data.decode("utf-8")).encode("utf-8")
            elif name == "xl/_rels/workbook.xml.rels":
                data = _clean_workbook_rels(data.decode("utf-8")).encode("utf-8")
            elif name == "[Content_Types].xml":
                data = _clean_content_types(data.decode("utf-8")).encode("utf-8")

            zout.writestr(info, data)


def main() -> None:
    source = template_path()
    build_template(source, OUTPUT)
    print(f"Шаблон создан: {OUTPUT.relative_to(ROOT)}")
    print(f"Источник:      {source.name}")
    print("Формулы и внешние связи удалены.")


if __name__ == "__main__":
    main()
