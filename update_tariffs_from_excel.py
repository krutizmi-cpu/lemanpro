
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

SOURCE_COMMISSIONS_XLSX = BASE_DIR / "source_commissions.xlsx"
SOURCE_LOGISTICS_XLSX = BASE_DIR / "source_logistics.xlsx"

def export_sheet(src: Path, sheet: str, dst: Path):
    df = pd.read_excel(src, sheet_name=sheet)
    df.to_csv(dst, index=False)
    print(f"OK: {dst.name} <= {src.name} / {sheet}")

def main():
    if not SOURCE_COMMISSIONS_XLSX.exists():
        raise FileNotFoundError(f"Положите файл комиссий как {SOURCE_COMMISSIONS_XLSX.name}")
    if not SOURCE_LOGISTICS_XLSX.exists():
        raise FileNotFoundError(f"Положите файл логистики как {SOURCE_LOGISTICS_XLSX.name}")

    DATA_DIR.mkdir(exist_ok=True)

    export_sheet(SOURCE_COMMISSIONS_XLSX, "Комиссия_FBS и FBO", DATA_DIR / "commissions_fbs_fbo.csv")
    export_sheet(SOURCE_COMMISSIONS_XLSX, "Тарифы (Последняя миля)", DATA_DIR / "last_mile.csv")
    export_sheet(SOURCE_COMMISSIONS_XLSX, "Тарифы (возврат Последняя миля)", DATA_DIR / "return_last_mile.csv")
    export_sheet(SOURCE_COMMISSIONS_XLSX, "Тарифы (Возврат Доставка до СЦ)", DATA_DIR / "return_zero_mile.csv")
    export_sheet(SOURCE_COMMISSIONS_XLSX, "Зоны (услуга Последняя миля)", DATA_DIR / "zones.csv")

    # Для нулевой мили берём логистический файл, который вы уже дали по текущей версии.
    export_sheet(SOURCE_LOGISTICS_XLSX, "Тарифы (услуга Нулевая миля)", DATA_DIR / "zero_mile.csv")

    print("\nГотово. Обновлённые CSV лежат в папке data/")

if __name__ == "__main__":
    main()
