"""Сравнение двух xlsx: сгенерированного и эталона из Office 2021."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
GENERATED = ROOT / "output" / "чек-лист_7260.xlsx"
REFERENCE = ROOT / "output" / "чек-лист_7260_2021.xlsx"


def list_entries(path: Path) -> dict[str, int]:
    with zipfile.ZipFile(path) as z:
        return {info.filename: info.file_size for info in z.infolist()}


def sheet_meta(path: Path, sheet: str) -> dict:
    with zipfile.ZipFile(path) as z:
        data = z.read(f"xl/worksheets/{sheet}.xml").decode("utf-8", errors="replace")
    root = re.match(r"<worksheet\b[^>]*>", data)
    page_setup = re.search(r"<pageSetup[^>]*/>", data)
    drawing = re.search(r'<drawing[^>]*/>', data)
    styles_idx = [int(x) for x in re.findall(r'\ss="(\d+)"', data)]
    return {
        "root": root.group(0)[:220] if root else None,
        "has_xmlns_r": "xmlns:r=" in data,
        "pageSetup": page_setup.group(0) if page_setup else None,
        "drawing": drawing.group(0) if drawing else None,
        "style_min": min(styles_idx) if styles_idx else None,
        "style_max": max(styles_idx) if styles_idx else None,
        "style_count": len(styles_idx),
        "xml_len": len(data),
    }


def styles_meta(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        data = z.read("xl/styles.xml").decode("utf-8", errors="replace")
    def count(tag: str) -> int:
        m = re.search(rf"<{tag} count=\"(\d+)\"", data)
        return int(m.group(1)) if m else 0
    return {
        "xml_len": len(data),
        "fonts": count("fonts"),
        "fills": count("fills"),
        "borders": count("borders"),
        "cellXfs": count("cellXfs"),
        "theme_refs": data.count("theme"),
    }


def workbook_meta(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as z:
        data = z.read("xl/workbook.xml").decode("utf-8", errors="replace")
    return re.findall(r'<definedName[^>]*Print[^<]*</definedName>', data)


def diff_sets(a: set[str], b: set[str], label_a: str, label_b: str) -> None:
    only_a = sorted(a - b)
    only_b = sorted(b - a)
    if only_a:
        print(f"  только в {label_a} ({len(only_a)}):")
        for x in only_a:
            print(f"    - {x}")
    if only_b:
        print(f"  только в {label_b} ({len(only_b)}):")
        for x in only_b:
            print(f"    - {x}")
    if not only_a and not only_b:
        print("  одинаковый набор файлов в zip")


def main() -> None:
    for p in (GENERATED, REFERENCE):
        if not p.is_file():
            print(f"НЕТ ФАЙЛА: {p}")
            return

    gen_entries = list_entries(GENERATED)
    ref_entries = list_entries(REFERENCE)
    print("=== Файлы ===")
    print(f"Сгенерированный: {GENERATED.name} ({GENERATED.stat().st_size:,} байт)")
    print(f"Эталон 2021:     {REFERENCE.name} ({REFERENCE.stat().st_size:,} байт)")
    print()

    print("=== Состав zip ===")
    diff_sets(set(gen_entries), set(ref_entries), "сгенер.", "эталон")
    common = sorted(set(gen_entries) & set(ref_entries))
    size_diffs = [
        (n, gen_entries[n], ref_entries[n])
        for n in common
        if gen_entries[n] != ref_entries[n]
    ]
    if size_diffs:
        print(f"  размер отличается ({len(size_diffs)} файлов):")
        for n, gs, rs in size_diffs:
            print(f"    {n}: {gs:,} vs {rs:,} ({gs - rs:+,})")
    print()

    for label, path in (("Сгенерированный", GENERATED), ("Эталон 2021", REFERENCE)):
        print(f"=== {label}: assets ===")
        with zipfile.ZipFile(path) as z:
            assets = sorted(
                n for n in z.namelist()
                if any(k in n for k in ("drawing", "printer", "rels/sheet", "externalLink"))
            )
            print(" ", assets)
        print()

    for sheet in ("sheet1", "sheet2"):
        print(f"=== Лист {sheet} ===")
        g = sheet_meta(GENERATED, sheet)
        r = sheet_meta(REFERENCE, sheet)
        for key in g:
            gv, rv = g[key], r[key]
            mark = "  <--" if gv != rv else ""
            print(f"  {key}: {gv!r} | {rv!r}{mark}")
        print()

    print("=== styles.xml ===")
    gs = styles_meta(GENERATED)
    rs = styles_meta(REFERENCE)
    for key in gs:
        print(f"  {key}: {gs[key]} | {rs[key]}{'  <--' if gs[key] != rs[key] else ''}")
    print()

    print("=== Области печати (workbook.xml) ===")
    gpa = workbook_meta(GENERATED)
    rpa = workbook_meta(REFERENCE)
    for i, (g, r) in enumerate(zip(gpa, rpa)):
        print(f"  [{i}] gen: {g}")
        print(f"      ref: {r}{'  <--' if g != r else ''}")
    if len(gpa) != len(rpa):
        print(f"  count: {len(gpa)} vs {len(rpa)}")
    print()

    # Binary compare key preserved parts
    print("=== Бинарное совпадение с эталоном ===")
    with zipfile.ZipFile(GENERATED) as zg, zipfile.ZipFile(REFERENCE) as zr:
        for name in (
            "xl/drawings/drawing1.xml",
            "xl/printerSettings/printerSettings1.bin",
            "xl/printerSettings/printerSettings2.bin",
            "xl/styles.xml",
        ):
            if name in zg.namelist() and name in zr.namelist():
                same = zg.read(name) == zr.read(name)
                print(f"  {name}: {'OK' if same else 'РАЗЛИЧАЕТСЯ'}")
            else:
                print(f"  {name}: отсутствует в одном из файлов")


if __name__ == "__main__":
    main()
