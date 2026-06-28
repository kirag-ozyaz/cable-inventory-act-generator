"""SKS — загрузка справочников и генерация актов инвентаризации."""

from __future__ import annotations

from src.loaders import load_all


def main() -> None:
    data = load_all()

    print("График [1]:", data.schedule.shape)
    print("  столбцы:", ", ".join(data.schedule.columns[:5]), "…")

    print("Подстанции [2]:", data.substations.shape)

    print("Кабели [3]:", data.cables.shape)
    print("  по районам:")
    print(data.cables["сетевой_район"].value_counts().sort_index().to_string())
    print()
    print(data.cables.head(3).to_string())


if __name__ == "__main__":
    main()
