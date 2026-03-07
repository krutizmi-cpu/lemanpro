
import math
import sqlite3
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Лемана Про — юнит-экономика FBS / FBO",
    layout="wide",
    page_icon="📦",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = BASE_DIR / "lemanpro_products.db"


# =========================
# DB
# =========================
def init_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS products (
            sku TEXT PRIMARY KEY,
            name TEXT,
            template TEXT,
            item_type TEXT,
            subcategory TEXT,
            category TEXT,
            length_cm REAL DEFAULT 0,
            width_cm REAL DEFAULT 0,
            height_cm REAL DEFAULT 0,
            weight_kg REAL DEFAULT 0,
            cost_price REAL DEFAULT 0,
            current_price REAL DEFAULT 0,
            promo_price REAL DEFAULT 0,
            region TEXT,
            manual_commission REAL
        )
        """
    )
    conn.commit()
    return conn


# =========================
# Utils
# =========================
def to_float(value, default=0.0):
    if value is None:
        return default
    try:
        if isinstance(value, str):
            value = value.replace("\xa0", "").replace(" ", "").replace(",", ".").strip()
            if value == "":
                return default
        return float(value)
    except Exception:
        return default


def normalize_dimension(raw, unit: str) -> float:
    value = to_float(raw, 0.0)
    unit = (unit or "").strip().lower()
    if unit in ("мм", "mm"):
        return value / 10.0
    if unit in ("м", "meter", "метр", "метры"):
        return value * 100.0
    return value


def normalize_weight(raw, unit: str) -> float:
    value = to_float(raw, 0.0)
    unit = (unit or "").strip().lower()
    if unit in ("г", "гр", "g", "gr"):
        return value / 1000.0
    return value


def first_existing(row: pd.Series, names: List[str], default=None):
    row_map = {str(k).strip().lower(): row[k] for k in row.index}
    for n in names:
        v = row_map.get(n.strip().lower())
        if pd.notna(v):
            return v
    return default


def clean_text(x) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    return str(x).strip()


def liters_from_dimensions(length_cm: float, width_cm: float, height_cm: float) -> float:
    if min(length_cm, width_cm, height_cm) <= 0:
        return 0.0
    return (length_cm * width_cm * height_cm) / 1000.0


def parse_bracket(text: str) -> Tuple[float, Optional[float]]:
    s = clean_text(text).lower().replace(",", ".")
    nums = []
    current = ""
    for ch in s:
        if ch.isdigit() or ch == ".":
            current += ch
        else:
            if current:
                nums.append(float(current))
                current = ""
    if current:
        nums.append(float(current))

    if len(nums) >= 2:
        return nums[0], nums[1]
    if len(nums) == 1:
        # На случай "от 120" или "120+"
        if "+" in s or "от" in s:
            return nums[0], None
        return 0.0, nums[0]
    return 0.0, None


def value_in_bracket(value: float, low: float, high: Optional[float]) -> bool:
    if high is None:
        return value >= low
    return low <= value <= high


def normalize_dest_label(label: str) -> str:
    s = clean_text(label)
    low = s.lower()
    if low in {"москва и мо", "москва", "московская область"}:
        return "Москва и МО"
    if low in {"спб и ло", "спб", "санкт-петербург", "ленинградская область"}:
        return "СПБ и ЛО"
    return s


def load_csv_checked(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Не найден файл: {path}")
    return pd.read_csv(path)


# =========================
# Tariffs from repo files
# =========================
@st.cache_data(show_spinner=False)
def load_commissions_from_repo(data_dir: Path) -> Dict[str, Dict[str, float]]:
    df = load_csv_checked(data_dir / "commissions_fbs_fbo.csv")
    df.columns = [clean_text(c) for c in df.columns]

    result = {
        "template": {},
        "item_type": {},
        "subcategory": {},
        "category": {},
    }

    for _, row in df.iterrows():
        commission_raw = to_float(first_existing(row, ["Комиссия"]), 0.0)
        commission = commission_raw if commission_raw <= 1 else commission_raw / 100.0

        template = clean_text(first_existing(row, ["Шаблон товара"]))
        item_type = clean_text(first_existing(row, ["Тип товара"]))
        subcategory = clean_text(first_existing(row, ["Подкатегория товара", "Подкатегория товара"]))
        category = clean_text(first_existing(row, ["Категория"]))

        if template:
            result["template"][template.lower()] = commission
        if item_type:
            result["item_type"][item_type.lower()] = commission
        if subcategory:
            result["subcategory"][subcategory.lower()] = commission
        if category:
            result["category"][category.lower()] = commission

    return result


@st.cache_data(show_spinner=False)
def load_logistics_from_repo(data_dir: Path):
    zero_df = load_csv_checked(data_dir / "zero_mile.csv")
    return_zero_df = load_csv_checked(data_dir / "return_zero_mile.csv")
    last_mile_df = load_csv_checked(data_dir / "last_mile.csv")
    return_last_mile_df = load_csv_checked(data_dir / "return_last_mile.csv")
    zones_df = load_csv_checked(data_dir / "zones.csv")

    def build_simple_table(df):
        rows = []
        for _, row in df.iterrows():
            bracket = clean_text(first_existing(row, ["Объемный брейк отправления, в л", "Объемный брейк отправления, в л."]))
            tariff = to_float(first_existing(row, ["Тариф, с НДС"]), 0.0)
            low, high = parse_bracket(bracket)
            rows.append(
                {
                    "low": low,
                    "high": high,
                    "base_tariff": tariff,
                    "extra_per_l": 0.0,
                    "raw_bracket": bracket,
                }
            )
        return rows

    def build_last_mile_table(df):
        rows = []
        for _, row in df.iterrows():
            origin_zone = int(to_float(first_existing(row, ["Зона откуда"]), 0))
            dest_label = normalize_dest_label(first_existing(row, ["Зона куда"]))
            bracket = clean_text(first_existing(row, ["Весовой брейк отправления, в л.", "Весовой брейк отправления, в л"]))
            tariff = to_float(first_existing(row, ["Тариф, с НДС"]), 0.0)
            extra_per_l = to_float(first_existing(row, ["+1 л, с НДС"]), 0.0)
            low, high = parse_bracket(bracket)
            rows.append(
                {
                    "origin_zone": origin_zone,
                    "dest_label": dest_label,
                    "low": low,
                    "high": high,
                    "base_tariff": tariff,
                    "extra_per_l": extra_per_l,
                    "raw_bracket": bracket,
                }
            )
        return rows

    zone_map = {}
    for _, row in zones_df.iterrows():
        region = clean_text(first_existing(row, ["Край/Область"]))
        zone = int(to_float(first_existing(row, ["Зона"]), 0))
        if region:
            zone_map[region.lower()] = zone

    return {
        "zero_mile": build_simple_table(zero_df),
        "return_zero_mile": build_simple_table(return_zero_df),
        "last_mile": build_last_mile_table(last_mile_df),
        "return_last_mile": build_last_mile_table(return_last_mile_df),
        "zone_map": zone_map,
    }


def get_simple_tariff(brackets: List[dict], liters: float) -> float:
    liters = max(liters, 0.0)
    for row in brackets:
        if value_in_bracket(liters, row["low"], row["high"]):
            if row["high"] is None and row["extra_per_l"] > 0:
                extra = max(liters - row["low"], 0.0)
                return row["base_tariff"] + math.ceil(extra) * row["extra_per_l"]
            return row["base_tariff"]
    return 0.0


def get_last_mile_tariff(table: List[dict], origin_zone: int, destination_label: str, liters: float) -> float:
    destination_label = normalize_dest_label(destination_label)
    liters = max(liters, 0.0)
    candidates = [r for r in table if r["origin_zone"] == origin_zone and r["dest_label"] == destination_label]
    for row in candidates:
        if value_in_bracket(liters, row["low"], row["high"]):
            if row["high"] is None and row["extra_per_l"] > 0:
                extra = max(liters - row["low"], 0.0)
                return row["base_tariff"] + math.ceil(extra) * row["extra_per_l"]
            return row["base_tariff"]
    return 0.0


def resolve_destination_label(region: str, zone_map: Dict[str, int]) -> Tuple[str, Optional[int]]:
    region_clean = clean_text(region)
    if not region_clean:
        return "Москва и МО", None

    direct = region_clean.lower()
    if direct in {"москва и мо", "москва", "московская область"}:
        return "Москва и МО", 0
    if direct in {"спб и ло", "санкт-петербург", "ленинградская область", "спб"}:
        return "СПБ и ЛО", 0

    zone = zone_map.get(direct)
    if zone is None:
        return "Москва и МО", None
    return str(zone), zone


# =========================
# Commission lookup
# =========================
def get_commission(row: pd.Series, commission_dict: Dict[str, Dict[str, float]]) -> Tuple[float, str]:
    manual = to_float(row.get("manual_commission"), -1)
    if manual >= 0:
        return (manual / 100.0 if manual > 1 else manual), "manual"

    template = clean_text(row.get("template")).lower()
    item_type = clean_text(row.get("item_type")).lower()
    subcategory = clean_text(row.get("subcategory")).lower()
    category = clean_text(row.get("category")).lower()

    if template and template in commission_dict["template"]:
        return commission_dict["template"][template], "template"
    if item_type and item_type in commission_dict["item_type"]:
        return commission_dict["item_type"][item_type], "item_type"
    if subcategory and subcategory in commission_dict["subcategory"]:
        return commission_dict["subcategory"][subcategory], "subcategory"
    if category and category in commission_dict["category"]:
        return commission_dict["category"][category], "category"
    return 0.0, "not_found"


# =========================
# Taxes
# =========================
def calc_tax(price: float, profit_before_tax: float, regime: str) -> Tuple[float, float, float]:
    price = max(price, 0.0)
    profit_before_tax = float(profit_before_tax)

    regimes = {
        "ОСНО (налог на прибыль 25%)": ("profit", 0.25),
        "УСН Доходы (6%)": ("revenue", 0.06),
        "УСН Доходы-Расходы (15%)": ("profit", 0.15),
        "АУСН Доходы (8%)": ("revenue", 0.08),
        "АУСН Доходы-Расходы (20%)": ("profit", 0.20),
        "Без налога": ("none", 0.0),
    }

    mode, rate = regimes.get(regime, ("none", 0.0))

    if mode == "revenue":
        tax = price * rate
    elif mode == "profit":
        tax = max(profit_before_tax, 0.0) * rate
    else:
        tax = 0.0

    profit_after_tax = profit_before_tax - tax
    margin_after_tax = (profit_after_tax / price * 100) if price > 0 else 0.0
    return round(tax, 2), round(profit_after_tax, 2), round(margin_after_tax, 2)


# =========================
# Core math
# =========================
def calculate_unit_metrics(
    row: pd.Series,
    scheme: str,
    tax_regime: str,
    commission_dict: Dict[str, Dict[str, float]],
    logistics_dict: dict,
    origin_zone: int,
    acquiring_pct: float,
    payout_pct: float,
    marketing_pct: float,
    other_mp_pct: float,
    packing_rub: float,
    other_fixed_rub: float,
    fbo_inbound_per_unit: float,
    buyout_pct: float,
    client_return_pct: float,
    target_margin_pct: float,
) -> dict:
    commission_pct, commission_source = get_commission(row, commission_dict)

    cost_price = to_float(row.get("cost_price"), 0.0)
    current_price = to_float(row.get("current_price"), 0.0)
    promo_price = to_float(row.get("promo_price"), 0.0)

    length_cm = to_float(row.get("length_cm"), 0.0)
    width_cm = to_float(row.get("width_cm"), 0.0)
    height_cm = to_float(row.get("height_cm"), 0.0)
    weight_kg = to_float(row.get("weight_kg"), 0.0)

    liters = liters_from_dimensions(length_cm, width_cm, height_cm)
    chargeable_liters = liters if liters > 0 else weight_kg

    region = clean_text(row.get("region"))
    destination_label, region_zone = resolve_destination_label(region, logistics_dict["zone_map"])
    destination_zone = region_zone if region_zone is not None else 0

    zero_mile = get_simple_tariff(logistics_dict["zero_mile"], chargeable_liters)
    return_zero_mile = get_simple_tariff(logistics_dict["return_zero_mile"], chargeable_liters)
    last_mile = get_last_mile_tariff(logistics_dict["last_mile"], origin_zone, destination_label, chargeable_liters)
    return_last_mile = get_last_mile_tariff(logistics_dict["return_last_mile"], origin_zone, destination_label, chargeable_liters)

    buyout_rate = max(min(buyout_pct / 100.0, 0.9999), 0.0001)
    cancel_rate = 1.0 - buyout_rate
    client_return_rate = max(min(client_return_pct / 100.0, 1.0), 0.0)

    if scheme == "FBS":
        expected_zero_mile = zero_mile / buyout_rate
        expected_return_zero_mile = return_zero_mile * (cancel_rate / buyout_rate)
    else:
        expected_zero_mile = max(fbo_inbound_per_unit, 0.0)
        expected_return_zero_mile = 0.0

    expected_last_mile = last_mile / buyout_rate
    expected_return_last_mile = (
        return_last_mile * (cancel_rate / buyout_rate)
        + return_last_mile * client_return_rate
    )

    logistics_total = expected_zero_mile + expected_return_zero_mile + expected_last_mile + expected_return_last_mile
    fixed_costs = cost_price + logistics_total + packing_rub + other_fixed_rub
    variable_pct = commission_pct + acquiring_pct / 100.0 + payout_pct / 100.0 + marketing_pct / 100.0 + other_mp_pct / 100.0

    def metrics_for_price(price: float):
        price = max(price, 0.0)
        variable_costs = price * variable_pct
        profit_before_tax = price - fixed_costs - variable_costs
        margin_before_tax = (profit_before_tax / price * 100) if price > 0 else 0.0
        tax, profit_after_tax, margin_after_tax = calc_tax(price, profit_before_tax, tax_regime)
        markup_pct = ((price / cost_price - 1) * 100) if cost_price > 0 else 0.0
        return {
            "price": round(price, 2),
            "variable_costs": round(variable_costs, 2),
            "profit_before_tax": round(profit_before_tax, 2),
            "margin_before_tax": round(margin_before_tax, 2),
            "tax": tax,
            "profit_after_tax": profit_after_tax,
            "margin_after_tax": margin_after_tax,
            "markup_pct": round(markup_pct, 2),
        }

    denom = 1.0 - variable_pct - target_margin_pct / 100.0
    recommended_price = fixed_costs / denom if denom > 0 else 0.0

    base_price = promo_price if promo_price > 0 else current_price
    current_metrics = metrics_for_price(base_price)
    recommended_metrics = metrics_for_price(recommended_price)

    bad_flag = ""
    if commission_source == "not_found":
        bad_flag = "Нет комиссии"
    elif denom <= 0:
        bad_flag = "Целевая маржа недостижима"
    elif current_metrics["profit_after_tax"] < 0:
        bad_flag = "Убыточно"
    elif current_metrics["margin_before_tax"] < target_margin_pct:
        bad_flag = "Ниже цели"

    return {
        "SKU": clean_text(row.get("sku")),
        "Название": clean_text(row.get("name")),
        "Схема": scheme,
        "Регион": region,
        "Зона откуда": origin_zone,
        "Зона получателя": destination_zone,
        "Объем, л": round(chargeable_liters, 3),
        "Вес, кг": round(weight_kg, 3),
        "Себестоимость, руб": round(cost_price, 2),
        "Текущая цена, руб": round(current_price, 2),
        "Цена акции, руб": round(promo_price, 2),
        "Цена расчета, руб": current_metrics["price"],
        "Комиссия, %": round(commission_pct * 100, 2),
        "Источник комиссии": commission_source,
        "Нулевая миля, руб": round(expected_zero_mile, 2),
        "Возврат нулевой мили, руб": round(expected_return_zero_mile, 2),
        "Последняя миля, руб": round(expected_last_mile, 2),
        "Возврат последней мили, руб": round(expected_return_last_mile, 2),
        "Логистика итого, руб": round(logistics_total, 2),
        "Переменные расходы, руб": current_metrics["variable_costs"],
        "Прибыль до налога (текущая), руб": current_metrics["profit_before_tax"],
        "Маржа до налога (текущая), %": current_metrics["margin_before_tax"],
        "Налог (текущая), руб": current_metrics["tax"],
        "Прибыль после налога (текущая), руб": current_metrics["profit_after_tax"],
        "Маржа после налога (текущая), %": current_metrics["margin_after_tax"],
        "Наценка (текущая), %": current_metrics["markup_pct"],
        "Рекоменд. цена, руб": recommended_metrics["price"],
        "Прибыль до налога (рекоменд.), руб": recommended_metrics["profit_before_tax"],
        "Маржа до налога (рекоменд.), %": recommended_metrics["margin_before_tax"],
        "Налог (рекоменд.), руб": recommended_metrics["tax"],
        "Прибыль после налога (рекоменд.), руб": recommended_metrics["profit_after_tax"],
        "Маржа после налога (рекоменд.), %": recommended_metrics["margin_after_tax"],
        "Наценка (рекоменд.), %": recommended_metrics["markup_pct"],
        "Флаг": bad_flag,
    }


# =========================
# UI
# =========================
conn = init_db()

st.title("Лемана Про — единый калькулятор юнит-экономики")
st.caption("Версия с тарифами и комиссиями из файлов репозитория: data/*.csv")

with st.sidebar:
    st.header("Параметры модели")

    scheme = st.radio("Схема работы", ["FBS", "FBO"], horizontal=True)

    tax_regime = st.selectbox(
        "Налоговый режим",
        [
            "ОСНО (налог на прибыль 25%)",
            "УСН Доходы (6%)",
            "УСН Доходы-Расходы (15%)",
            "АУСН Доходы (8%)",
            "АУСН Доходы-Расходы (20%)",
            "Без налога",
        ],
    )

    st.divider()
    st.subheader("Процентные расходы")
    target_margin_pct = st.slider("Целевая маржа до налога, %", 0, 80, 20)
    acquiring_pct = st.number_input("Эквайринг, %", min_value=0.0, max_value=10.0, value=1.5, step=0.1)
    payout_pct = st.number_input("Ранняя выплата / фин. услуги, %", min_value=0.0, max_value=10.0, value=0.0, step=0.1)
    marketing_pct = st.number_input("Маркетинг / реклама, %", min_value=0.0, max_value=50.0, value=5.0, step=0.1)
    other_mp_pct = st.number_input("Прочие % маркетплейса, %", min_value=0.0, max_value=50.0, value=0.0, step=0.1)

    st.divider()
    st.subheader("Фиксированные расходы")
    packing_rub = st.number_input("Упаковка на ед., руб", min_value=0.0, max_value=5000.0, value=0.0, step=10.0)
    other_fixed_rub = st.number_input("Прочие фикс. расходы на ед., руб", min_value=0.0, max_value=5000.0, value=0.0, step=10.0)
    if scheme == "FBO":
        fbo_inbound_per_unit = st.number_input(
            "Входящая логистика FBO на ед., руб",
            min_value=0.0,
            max_value=10000.0,
            value=0.0,
            step=10.0,
        )
    else:
        fbo_inbound_per_unit = 0.0

    st.divider()
    st.subheader("Поведенческие параметры")
    buyout_pct = st.slider("Выкуп, %", 1, 100, 95)
    client_return_pct = st.slider("Возвраты после выкупа, %", 0, 100, 5)

    st.divider()
    st.subheader("Логистика")
    origin_zone = st.number_input("Зона откуда", min_value=1, max_value=10, value=9, step=1)

st.markdown("### 1. Проверка тарифных файлов в репозитории")
required_files = [
    "commissions_fbs_fbo.csv",
    "zero_mile.csv",
    "return_zero_mile.csv",
    "last_mile.csv",
    "return_last_mile.csv",
    "zones.csv",
]
missing = [f for f in required_files if not (DATA_DIR / f).exists()]
if missing:
    st.error("Не хватает файлов в папке data: " + ", ".join(missing))
    st.stop()

try:
    commission_dict = load_commissions_from_repo(DATA_DIR)
    logistics_dict = load_logistics_from_repo(DATA_DIR)
except Exception as e:
    st.error(f"Не удалось прочитать файлы из data/: {e}")
    st.stop()

st.success("Тарифы и комиссии успешно загружены из папки data/.")

with st.expander("Какие файлы сейчас подключены", expanded=False):
    st.write("\n".join(f"- {f}" for f in required_files))

st.markdown("### 2. Каталог товаров")
dim_col, wt_col = st.columns(2)
with dim_col:
    dim_unit = st.selectbox("Единица размеров в файле каталога", ["см", "мм", "м"], index=0)
with wt_col:
    wt_unit = st.selectbox("Единица веса в файле каталога", ["кг", "г"], index=0)

catalog_file = st.file_uploader(
    "Загрузить каталог Excel",
    type=["xlsx"],
    key="catalog_file",
    help="Поддерживаются: SKU/Артикул, Название/Наименование, Шаблон товара, Тип товара, Подкатегория товара, Категория, Длина, Ширина, Высота, Вес, Себестоимость, Текущая цена, Цена акции, Регион.",
)

if catalog_file:
    try:
        df = pd.read_excel(catalog_file)
        save_rows = []
        for _, row in df.iterrows():
            sku = clean_text(first_existing(row, ["SKU", "Артикул", "Артикул продавца"]))
            name = clean_text(first_existing(row, ["Название", "Наименование", "Товар"]))
            if not sku:
                continue

            save_rows.append(
                (
                    sku,
                    name,
                    clean_text(first_existing(row, ["Шаблон товара", "Шаблон"])),
                    clean_text(first_existing(row, ["Тип товара"])),
                    clean_text(first_existing(row, ["Подкатегория товара", "Подкатегория"])),
                    clean_text(first_existing(row, ["Категория"])),
                    normalize_dimension(first_existing(row, ["Длина", "Длина, см", "Длина, мм"]), dim_unit),
                    normalize_dimension(first_existing(row, ["Ширина", "Ширина, см", "Ширина, мм"]), dim_unit),
                    normalize_dimension(first_existing(row, ["Высота", "Высота, см", "Высота, мм"]), dim_unit),
                    normalize_weight(first_existing(row, ["Вес", "Вес, кг", "Вес, г"]), wt_unit),
                    to_float(first_existing(row, ["Себестоимость", "Себес", "Закупка"]), 0.0),
                    to_float(first_existing(row, ["Текущая цена", "Цена", "Цена без акции"]), 0.0),
                    to_float(first_existing(row, ["Цена акции", "Акционная цена", "Цена со скидкой"]), 0.0),
                    clean_text(first_existing(row, ["Регион", "Область", "Край/Область"])),
                    None,
                )
            )

        if save_rows:
            conn.executemany(
                """
                INSERT INTO products (
                    sku, name, template, item_type, subcategory, category,
                    length_cm, width_cm, height_cm, weight_kg,
                    cost_price, current_price, promo_price, region, manual_commission
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(sku) DO UPDATE SET
                    name=excluded.name,
                    template=excluded.template,
                    item_type=excluded.item_type,
                    subcategory=excluded.subcategory,
                    category=excluded.category,
                    length_cm=excluded.length_cm,
                    width_cm=excluded.width_cm,
                    height_cm=excluded.height_cm,
                    weight_kg=excluded.weight_kg,
                    cost_price=excluded.cost_price,
                    current_price=excluded.current_price,
                    promo_price=excluded.promo_price,
                    region=excluded.region
                """,
                save_rows,
            )
            conn.commit()
            st.success(f"Загружено / обновлено SKU: {len(save_rows)}")
    except Exception as e:
        st.error(f"Ошибка загрузки каталога: {e}")

catalog_df = pd.read_sql_query("SELECT * FROM products ORDER BY sku", conn)

if catalog_df.empty:
    st.warning("Каталог пока пуст. Загрузите Excel с товарами.")
    st.stop()

st.markdown("#### Ручная корректировка каталога и комиссий")
editable = catalog_df.copy()
editable["manual_commission"] = editable["manual_commission"].fillna("")
edited = st.data_editor(
    editable,
    use_container_width=True,
    num_rows="dynamic",
    hide_index=True,
    column_config={
        "manual_commission": st.column_config.NumberColumn("Комиссия вручную, %", min_value=0.0, max_value=100.0, step=0.1),
        "current_price": st.column_config.NumberColumn("Текущая цена, руб", min_value=0.0, step=1.0),
        "promo_price": st.column_config.NumberColumn("Цена акции, руб", min_value=0.0, step=1.0),
        "cost_price": st.column_config.NumberColumn("Себестоимость, руб", min_value=0.0, step=1.0),
        "length_cm": st.column_config.NumberColumn("Длина, см", min_value=0.0, step=0.1),
        "width_cm": st.column_config.NumberColumn("Ширина, см", min_value=0.0, step=0.1),
        "height_cm": st.column_config.NumberColumn("Высота, см", min_value=0.0, step=0.1),
        "weight_kg": st.column_config.NumberColumn("Вес, кг", min_value=0.0, step=0.001),
    },
)

if st.button("Сохранить изменения в каталог", type="secondary"):
    rows = []
    for _, row in edited.iterrows():
        rows.append(
            (
                clean_text(row["name"]),
                clean_text(row["template"]),
                clean_text(row["item_type"]),
                clean_text(row["subcategory"]),
                clean_text(row["category"]),
                to_float(row["length_cm"], 0.0),
                to_float(row["width_cm"], 0.0),
                to_float(row["height_cm"], 0.0),
                to_float(row["weight_kg"], 0.0),
                to_float(row["cost_price"], 0.0),
                to_float(row["current_price"], 0.0),
                to_float(row["promo_price"], 0.0),
                clean_text(row["region"]),
                None if clean_text(row["manual_commission"]) == "" else to_float(row["manual_commission"], 0.0),
                clean_text(row["sku"]),
            )
        )
    conn.executemany(
        """
        UPDATE products
        SET name=?, template=?, item_type=?, subcategory=?, category=?,
            length_cm=?, width_cm=?, height_cm=?, weight_kg=?,
            cost_price=?, current_price=?, promo_price=?, region=?, manual_commission=?
        WHERE sku=?
        """,
        rows,
    )
    conn.commit()
    st.success("Изменения сохранены.")
    catalog_df = pd.read_sql_query("SELECT * FROM products ORDER BY sku", conn)

st.markdown("### 3. Расчёт")

if st.button("Рассчитать юнит-экономику", type="primary"):
    current_catalog = pd.read_sql_query("SELECT * FROM products ORDER BY sku", conn)

    results = []
    for _, row in current_catalog.iterrows():
        results.append(
            calculate_unit_metrics(
                row=row,
                scheme=scheme,
                tax_regime=tax_regime,
                commission_dict=commission_dict,
                logistics_dict=logistics_dict,
                origin_zone=int(origin_zone),
                acquiring_pct=acquiring_pct,
                payout_pct=payout_pct,
                marketing_pct=marketing_pct,
                other_mp_pct=other_mp_pct,
                packing_rub=packing_rub,
                other_fixed_rub=other_fixed_rub,
                fbo_inbound_per_unit=fbo_inbound_per_unit,
                buyout_pct=buyout_pct,
                client_return_pct=client_return_pct,
                target_margin_pct=target_margin_pct,
            )
        )

    result_df = pd.DataFrame(results)

    total_sku = len(result_df)
    profitable_count = int((result_df["Прибыль после налога (текущая), руб"] > 0).sum())
    loss_count = int((result_df["Прибыль после налога (текущая), руб"] <= 0).sum())
    avg_margin = round(result_df["Маржа после налога (текущая), %"].mean(), 2) if total_sku else 0.0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("SKU в расчёте", total_sku)
    k2.metric("Плюсовых SKU", profitable_count)
    k3.metric("Убыточных / нулевых SKU", loss_count)
    k4.metric("Средняя маржа после налога, %", avg_margin)

    st.markdown("#### Результат")

    def color_flags(val):
        if val in ("Убыточно", "Нет комиссии", "Целевая маржа недостижима"):
            return "background-color: #ffdddd"
        if val == "Ниже цели":
            return "background-color: #fff2cc"
        return ""

    styled = result_df.style.applymap(color_flags, subset=["Флаг"])
    st.dataframe(styled, use_container_width=True)

    st.download_button(
        "Скачать результат CSV",
        data=result_df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"lemanpro_unit_economics_{scheme.lower()}.csv",
        mime="text/csv",
    )

    template_cols = [
        "SKU",
        "Название",
        "Себестоимость, руб",
        "Текущая цена, руб",
        "Цена акции, руб",
        "Комиссия, %",
        "Источник комиссии",
        "Логистика итого, руб",
        "Прибыль после налога (текущая), руб",
        "Маржа после налога (текущая), %",
        "Рекоменд. цена, руб",
        "Наценка (рекоменд.), %",
        "Флаг",
    ]
    notes_df = result_df[template_cols].copy()
    notes_df["Комментарий для сотрудника"] = notes_df["Флаг"].map(
        {
            "Убыточно": "Проверь цену, комиссию и логистику: SKU сейчас убыточен.",
            "Ниже цели": "SKU прибыльный, но не дотягивает до целевой маржи.",
            "Нет комиссии": "Нужно вручную проставить категорию или комиссию.",
            "Целевая маржа недостижима": "Переменные расходы слишком высокие для выбранной целевой маржи.",
        }
    ).fillna("SKU в норме.")

    st.download_button(
        "Скачать шаблон для сотрудников CSV",
        data=notes_df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"lemanpro_template_for_team_{scheme.lower()}.csv",
        mime="text/csv",
    )
else:
    st.info("Загрузите каталог и нажмите «Рассчитать юнит-экономику».")
