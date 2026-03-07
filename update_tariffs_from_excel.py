
from __future__ import annotations

from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

WORKBOOK_CANDIDATES = [
    BASE_DIR / "source_commissions.xlsx",
    BASE_DIR / "source_logistics.xlsx",
]

SHEET_ALIASES = {
    "commissions_fbs_fbo": ["Комиссия_FBS и FBO", "Комиссии"],
    "last_mile": ["Тарифы (Последняя миля)", "Тарифы (услуга Последняя миля)"],
    "return_last_mile": ["Тарифы (возврат Последняя миля)"],
    "zero_mile": ["Тарифы (Доставка до СЦ)", "Тарифы (услуга Нулевая миля)"],
    "return_zero_mile": ["Тарифы (Возврат Доставка до СЦ)"],
    "zones": ["Зоны (услуга Последняя миля)"],
}


def available_workbooks() -> list[Path]:
    books = [path for path in WORKBOOK_CANDIDATES if path.exists()]
    if not books:
        raise FileNotFoundError(
            "Не найдены source_commissions.xlsx / source_logistics.xlsx в корне проекта."
        )
    return books


def read_from_any_workbook(aliases: list[str]) -> pd.DataFrame:
    for workbook in available_workbooks():
        xls = pd.ExcelFile(workbook)
        for sheet in aliases:
            if sheet in xls.sheet_names:
                return pd.read_excel(workbook, sheet_name=sheet)
    raise ValueError(f"Не найден ни один лист из списка: {aliases}")


def main():
    DATA_DIR.mkdir(exist_ok=True)

    commissions = read_from_any_workbook(SHEET_ALIASES["commissions_fbs_fbo"]).copy()
    commissions.columns = ["commission_rate", "template", "type", "subcategory", "category"]
    commissions.to_csv(DATA_DIR / "commissions_fbs_fbo.csv", index=False, encoding="utf-8-sig")

    last_mile = read_from_any_workbook(SHEET_ALIASES["last_mile"]).copy()
    last_mile.columns = ["origin_zone", "destination_zone", "volume_break", "tariff", "price", "extra_liter_price"]
    last_mile.to_csv(DATA_DIR / "last_mile.csv", index=False, encoding="utf-8-sig")

    return_last_mile = read_from_any_workbook(SHEET_ALIASES["return_last_mile"]).copy()
    return_last_mile.columns = ["origin_zone", "destination_zone", "volume_break", "tariff", "price", "extra_liter_price"]
    return_last_mile.to_csv(DATA_DIR / "return_last_mile.csv", index=False, encoding="utf-8-sig")

    zero_mile = read_from_any_workbook(SHEET_ALIASES["zero_mile"]).copy()
    zero_mile.columns = ["volume_break", "price"]
    zero_mile.to_csv(DATA_DIR / "zero_mile.csv", index=False, encoding="utf-8-sig")

    return_zero_mile = read_from_any_workbook(SHEET_ALIASES["return_zero_mile"]).copy()
    return_zero_mile.columns = ["volume_break", "price"]
    return_zero_mile.to_csv(DATA_DIR / "return_zero_mile.csv", index=False, encoding="utf-8-sig")

    zones = read_from_any_workbook(SHEET_ALIASES["zones"]).copy()
    zones.columns = ["region", "zone"]
    zones.to_csv(DATA_DIR / "zones.csv", index=False, encoding="utf-8-sig")

    print("CSV обновлены в папке data/")


if __name__ == "__main__":
    main()
