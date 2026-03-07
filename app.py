
from __future__ import annotations

import math
import sqlite3
from dataclasses import dataclass
from difflib import get_close_matches
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Лемана Про — юнит-экономика FBS/FBO",
    page_icon="📦",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = BASE_DIR / "lemanpro_state.db"


# -----------------------------
# Storage
# -----------------------------
def init_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS product_catalog (
            sku TEXT PRIMARY KEY,
            name TEXT,
            template TEXT,
            type_name TEXT,
            subcategory TEXT,
            category TEXT,
            length_cm REAL,
            width_cm REAL,
            height_cm REAL,
            weight_kg REAL,
            cost_price REAL,
            current_price REAL,
            promo_price REAL
        )
        """
    )
    conn.commit()
    return conn


# -----------------------------
# Helpers
# -----------------------------
def norm_text(value) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().lower().replace("ё", "е").split())


def safe_float(value, default: float = 0.0) -> float:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(" ", "").replace(",", ".")
    if text == "":
        return default
    try:
        return float(text)
    except ValueError:
        return default


def parse_break(text: str) -> Tuple[float, float]:
    raw = norm_text(text)
    nums = []
    token = ""
    for ch in raw:
        if ch.isdigit() or ch in ".,":  # keep decimal
            token += ch
        else:
            if token:
                nums.append(float(token.replace(",", ".")))
                token = ""
    if token:
        nums.append(float(token.replace(",", ".")))
    if len(nums) >= 2:
        low, high = nums[0], nums[1]
    elif len(nums) == 1:
        low = 0.0
        high = nums[0]
    else:
        low = 0.0
        high = 999999.0
    return low, high


def normalize_dimensions(df: pd.DataFrame, dim_unit: str, weight_unit: str) -> pd.DataFrame:
    factor_dim = 0.1 if dim_unit == "мм" else 1.0
    factor_weight = 0.001 if weight_unit == "г" else 1.0

    for col in ["Длина", "Ширина", "Высота"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: safe_float(x) * factor_dim)
    if "Вес" in df.columns:
        df["Вес"] = df["Вес"].apply(lambda x: safe_float(x) * factor_weight)
    return df


def first_present(row: pd.Series, names: List[str]) -> str:
    for name in names:
        if name in row.index:
            value = row.get(name)
            if pd.notna(value) and str(value).strip():
                return str(value).strip()
    return ""


def flag_text(profit_after_tax: float, margin_after_tax_pct: float, commission_source: str, volume_l: float) -> str:
    flags = []
    if profit_after_tax < 0:
        flags.append("убыток")
    if margin_after_tax_pct < 0:
        flags.append("маржа < 0%")
    elif margin_after_tax_pct < 5:
        flags.append("маржа < 5%")
    if "fallback" in commission_source:
        flags.append("проверь комиссию")
    if volume_l > 120:
        flags.append("объем > 120 л")
    return ", ".join(flags) if flags else "ok"


# -----------------------------
# Data layer
# -----------------------------
@st.cache_data(show_spinner=False)
def load_data() -> Dict[str, pd.DataFrame]:
    commissions = pd.read_csv(DATA_DIR / "commissions_fbs_fbo.csv")
    last_mile = pd.read_csv(DATA_DIR / "last_mile.csv")
    return_last_mile = pd.read_csv(DATA_DIR / "return_last_mile.csv")
    zero_mile = pd.read_csv(DATA_DIR / "zero_mile.csv")
    return_zero_mile = pd.read_csv(DATA_DIR / "return_zero_mile.csv")
    zones = pd.read_csv(DATA_DIR / "zones.csv")

    for df in [last_mile, return_last_mile]:
        df["origin_zone"] = df["origin_zone"].astype(str).str.strip()
        df["destination_zone"] = df["destination_zone"].astype(str).str.strip()
        df["volume_from"], df["volume_to"] = zip(*df["volume_break"].map(parse_break))

    for df in [zero_mile, return_zero_mile]:
        df["volume_from"], df["volume_to"] = zip(*df["volume_break"].map(parse_break))

    commissions["commission_rate"] = commissions["commission_rate"].astype(float) * 100.0
    for col in ["template", "type", "subcategory", "category"]:
        commissions[f"{col}_norm"] = commissions[col].map(norm_text)

    zones["zone"] = zones["zone"].astype(str)
    zones["region_norm"] = zones["region"].map(norm_text)

    return {
        "commissions": commissions,
        "last_mile": last_mile,
        "return_last_mile": return_last_mile,
        "zero_mile": zero_mile,
        "return_zero_mile": return_zero_mile,
        "zones": zones,
    }


@dataclass
class CommissionMatch:
    rate: float
    source: str


class TariffEngine:
    def __init__(self, data: Dict[str, pd.DataFrame]):
        self.data = data
        self.commissions = data["commissions"]
        self.last_mile = data["last_mile"]
        self.return_last_mile = data["return_last_mile"]
        self.zero_mile = data["zero_mile"]
        self.return_zero_mile = data["return_zero_mile"]
        self.zones = data["zones"]

        self.template_index = {
            row["template_norm"]: float(row["commission_rate"])
            for _, row in self.commissions.dropna(subset=["template_norm"]).iterrows()
            if row["template_norm"]
        }
        self.type_index = {
            row["type_norm"]: float(row["commission_rate"])
            for _, row in self.commissions.dropna(subset=["type_norm"]).iterrows()
            if row["type_norm"]
        }
        self.subcategory_index = {
            row["subcategory_norm"]: float(row["commission_rate"])
            for _, row in self.commissions.dropna(subset=["subcategory_norm"]).iterrows()
            if row["subcategory_norm"]
        }
        self.category_index = {
            row["category_norm"]: float(row["commission_rate"])
            for _, row in self.commissions.dropna(subset=["category_norm"]).iterrows()
            if row["category_norm"]
        }

    def find_commission(self, template: str, type_name: str, subcategory: str, category: str, product_name: str = "") -> CommissionMatch:
        candidates = [
            ("template", norm_text(template), self.template_index),
            ("type", norm_text(type_name), self.type_index),
            ("subcategory", norm_text(subcategory), self.subcategory_index),
            ("category", norm_text(category), self.category_index),
        ]
        for source, key, idx in candidates:
            if key and key in idx:
                return CommissionMatch(idx[key], source)

        search_name = norm_text(product_name or template)
        if search_name:
            template_keys = list(self.template_index.keys())
            matches = get_close_matches(search_name, template_keys, n=1, cutoff=0.82)
            if matches:
                return CommissionMatch(self.template_index[matches[0]], "template_fallback_fuzzy")

        return CommissionMatch(0.0, "fallback_zero")

    def get_zone_code(self, region_name: str) -> Optional[str]:
        key = norm_text(region_name)
        if not key:
            return None
        exact = self.zones.loc[self.zones["region_norm"] == key]
        if not exact.empty:
            return str(exact.iloc[0]["zone"])
        variants = list(self.zones["region_norm"].dropna().unique())
        fuzzy = get_close_matches(key, variants, n=1, cutoff=0.84)
        if fuzzy:
            row = self.zones.loc[self.zones["region_norm"] == fuzzy[0]].iloc[0]
            return str(row["zone"])
        return None

    @staticmethod
    def _select_bracket(df: pd.DataFrame, volume_l: float) -> pd.Series:
        if df.empty:
            raise ValueError("Пустая тарифная таблица")
        hit = df[(df["volume_from"] <= volume_l) & (volume_l <= df["volume_to"])]
        if hit.empty:
            above = df.sort_values("volume_to")
            return above.iloc[-1]
        return hit.sort_values("volume_to").iloc[0]

    def zero_mile_cost(self, volume_l: float, return_trip: bool = False) -> float:
        table = self.return_zero_mile if return_trip else self.zero_mile
        row = self._select_bracket(table, volume_l)
        return float(row["price"])

    def last_mile_cost(self, origin_zone: str, destination_zone: str, volume_l: float, return_trip: bool = False) -> float:
        table = self.return_last_mile if return_trip else self.last_mile
        zone_df = table[
            (table["origin_zone"].astype(str) == str(origin_zone))
            & (table["destination_zone"].astype(str) == str(destination_zone))
        ].copy()
        if zone_df.empty:
            raise ValueError(f"Не найден тариф {origin_zone=} -> {destination_zone=}")
        row = self._select_bracket(zone_df, volume_l)
        base_price = float(row["price"])
        extra_per_liter = float(row.get("extra_liter_price", 0.0) or 0.0)
        volume_to = float(row["volume_to"])
        extra = max(volume_l - volume_to, 0.0) * extra_per_liter
        return base_price + extra


# -----------------------------
# Finance
# -----------------------------
TAX_OPTIONS = {
    "ОСНО (налог на прибыль 25%)": ("profit", 0.25),
    "УСН Доходы (6%)": ("revenue", 0.06),
    "УСН Доходы-Расходы (15%)": ("profit", 0.15),
    "АУСН Доходы (8%)": ("revenue", 0.08),
    "УСН + НДС 5%": ("revenue", 0.05),
    "УСН + НДС 7%": ("revenue", 0.07),
}


def calc_tax(expected_revenue: float, profit_before_tax: float, regime: str) -> float:
    mode, rate = TAX_OPTIONS.get(regime, ("profit", 0.0))
    if mode == "revenue":
        return max(expected_revenue, 0.0) * rate
    return max(profit_before_tax, 0.0) * rate


def solve_price_for_target_margin(
    target_margin_pct: float,
    tax_regime: str,
    cost_price: float,
    logistics_expected: float,
    fixed_costs: float,
    variable_rate_pct: float,
    sale_realization_rate: float,
) -> float:
    target = target_margin_pct / 100.0
    variable_rate = variable_rate_pct / 100.0

    def margin_after_tax(realized_price: float) -> float:
        expected_revenue = realized_price * sale_realization_rate
        variable_costs = expected_revenue * variable_rate
        profit_before_tax = expected_revenue - cost_price - logistics_expected - fixed_costs - variable_costs
        tax = calc_tax(expected_revenue, profit_before_tax, tax_regime)
        profit_after_tax = profit_before_tax - tax
        if expected_revenue <= 0:
            return -1.0
        return profit_after_tax / expected_revenue

    low, high = 0.0, max(cost_price + logistics_expected + fixed_costs + 1000.0, 1000.0)
    while margin_after_tax(high) < target and high < 1_000_000:
        high *= 1.5

    for _ in range(70):
        mid = (low + high) / 2
        if margin_after_tax(mid) >= target:
            high = mid
        else:
            low = mid
    return round(high, 2)


def compute_unit_economics(
    scheme: str,
    current_price: float,
    promo_price: float,
    promo_share_pct: float,
    cost_price: float,
    commission_pct: float,
    acquiring_pct: float,
    marketing_pct: float,
    early_payout_pct: float,
    extra_costs_rub: float,
    target_margin_pct: float,
    tax_regime: str,
    buyout_pct: float,
    return_after_buyout_pct: float,
    zero_mile_out: float,
    zero_mile_back: float,
    last_mile_out: float,
    last_mile_back: float,
    fbo_inbound_cost: float,
    marketplace_discount_pct: float,
) -> Dict[str, float]:
    promo_share = promo_share_pct / 100.0
    buyout_rate = buyout_pct / 100.0
    return_after_buyout_rate = return_after_buyout_pct / 100.0
    marketplace_discount_rate = marketplace_discount_pct / 100.0

    effective_current = max(current_price * (1 - marketplace_discount_rate), 0.0)
    effective_promo = max((promo_price if promo_price > 0 else current_price) * (1 - marketplace_discount_rate), 0.0)

    realized_price = effective_current * (1 - promo_share) + effective_promo * promo_share
    sale_realization_rate = max(buyout_rate * (1 - return_after_buyout_rate), 0.0)

    if scheme == "FBS":
        logistics_expected = (
            zero_mile_out
            + last_mile_out
            + (1 - buyout_rate) * (last_mile_back + zero_mile_back)
            + buyout_rate * return_after_buyout_rate * (last_mile_back + zero_mile_back)
        )
    else:
        logistics_expected = (
            fbo_inbound_cost
            + last_mile_out
            + (1 - buyout_rate) * last_mile_back
            + buyout_rate * return_after_buyout_rate * last_mile_back
        )

    expected_revenue = realized_price * sale_realization_rate
    variable_rate_pct = commission_pct + acquiring_pct + marketing_pct + early_payout_pct
    variable_costs = expected_revenue * variable_rate_pct / 100.0

    fixed_costs = extra_costs_rub
    profit_before_tax = expected_revenue - cost_price - logistics_expected - fixed_costs - variable_costs
    tax = calc_tax(expected_revenue, profit_before_tax, tax_regime)
    profit_after_tax = profit_before_tax - tax

    margin_before_tax_pct = (profit_before_tax / expected_revenue * 100.0) if expected_revenue > 0 else 0.0
    margin_after_tax_pct = (profit_after_tax / expected_revenue * 100.0) if expected_revenue > 0 else 0.0
    markup_pct = (profit_after_tax / cost_price * 100.0) if cost_price > 0 else 0.0

    recommended_realized = solve_price_for_target_margin(
        target_margin_pct=target_margin_pct,
        tax_regime=tax_regime,
        cost_price=cost_price,
        logistics_expected=logistics_expected,
        fixed_costs=fixed_costs,
        variable_rate_pct=variable_rate_pct,
        sale_realization_rate=sale_realization_rate,
    )
    promo_discount_pct = 0.0
    if current_price > 0 and promo_price > 0 and promo_price < current_price:
        promo_discount_pct = (1 - promo_price / current_price) * 100.0

    recommended_current = recommended_realized / (1 - marketplace_discount_rate) if marketplace_discount_rate < 1 else 0.0
    if promo_discount_pct > 0:
        recommended_promo = recommended_current * (1 - promo_discount_pct / 100.0)
    else:
        recommended_promo = recommended_current

    recommended_metrics = compute_recommended_metrics(
        recommended_realized=recommended_realized,
        cost_price=cost_price,
        logistics_expected=logistics_expected,
        fixed_costs=fixed_costs,
        variable_rate_pct=variable_rate_pct,
        sale_realization_rate=sale_realization_rate,
        tax_regime=tax_regime,
    )

    return {
        "effective_current_price": round(effective_current, 2),
        "effective_promo_price": round(effective_promo, 2),
        "realized_price": round(realized_price, 2),
        "sale_realization_rate_pct": round(sale_realization_rate * 100.0, 2),
        "expected_revenue": round(expected_revenue, 2),
        "logistics_expected": round(logistics_expected, 2),
        "variable_costs": round(variable_costs, 2),
        "profit_before_tax": round(profit_before_tax, 2),
        "tax": round(tax, 2),
        "profit_after_tax": round(profit_after_tax, 2),
        "margin_before_tax_pct": round(margin_before_tax_pct, 2),
        "margin_after_tax_pct": round(margin_after_tax_pct, 2),
        "markup_pct": round(markup_pct, 2),
        "recommended_realized_price": round(recommended_realized, 2),
        "recommended_current_price": round(recommended_current, 2),
        "recommended_promo_price": round(recommended_promo, 2),
        **recommended_metrics,
    }


def compute_recommended_metrics(
    recommended_realized: float,
    cost_price: float,
    logistics_expected: float,
    fixed_costs: float,
    variable_rate_pct: float,
    sale_realization_rate: float,
    tax_regime: str,
) -> Dict[str, float]:
    expected_revenue = recommended_realized * sale_realization_rate
    variable_costs = expected_revenue * variable_rate_pct / 100.0
    profit_before_tax = expected_revenue - cost_price - logistics_expected - fixed_costs - variable_costs
    tax = calc_tax(expected_revenue, profit_before_tax, tax_regime)
    profit_after_tax = profit_before_tax - tax
    margin_after_tax = (profit_after_tax / expected_revenue * 100.0) if expected_revenue > 0 else 0.0
    return {
        "recommended_profit_after_tax": round(profit_after_tax, 2),
        "recommended_margin_after_tax_pct": round(margin_after_tax, 2),
    }


# -----------------------------
# UI
# -----------------------------
def build_import_template() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "SKU": "SKU-001",
                "Название": "Кирпич строительный красный",
                "Шаблон товара": "200032_Кирпич",
                "Тип товара": "Блоки, кирпич и бетонные изделия",
                "Подкатегория товара": "Конструкционные материалы и изделия",
                "Категория": "Стройматериалы (конструкции, системы, оборудование)",
                "Длина": 25,
                "Ширина": 12,
                "Высота": 6.5,
                "Вес": 3.2,
                "Себестоимость": 19,
                "Текущая цена": 69,
                "Цена акции": 59,
            }
        ]
    )


def save_catalog_to_db(conn: sqlite3.Connection, df: pd.DataFrame) -> None:
    payload = []
    for _, row in df.iterrows():
        payload.append(
            (
                first_present(row, ["SKU", "Артикул", "sku"]),
                first_present(row, ["Название", "Наименование", "name"]),
                first_present(row, ["Шаблон товара", "Шаблон", "template"]),
                first_present(row, ["Тип товара", "type"]),
                first_present(row, ["Подкатегория товара", "subcategory"]),
                first_present(row, ["Категория", "category"]),
                safe_float(row.get("Длина", 0)),
                safe_float(row.get("Ширина", 0)),
                safe_float(row.get("Высота", 0)),
                safe_float(row.get("Вес", 0)),
                safe_float(row.get("Себестоимость", 0)),
                safe_float(row.get("Текущая цена", 0)),
                safe_float(row.get("Цена акции", 0)),
            )
        )

    conn.executemany(
        """
        INSERT OR REPLACE INTO product_catalog (
            sku, name, template, type_name, subcategory, category,
            length_cm, width_cm, height_cm, weight_kg, cost_price, current_price, promo_price
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        payload,
    )
    conn.commit()


def read_catalog(conn: sqlite3.Connection) -> pd.DataFrame:
    return pd.read_sql("SELECT * FROM product_catalog ORDER BY sku", conn)


def prepare_editor_df(df: pd.DataFrame, engine: TariffEngine) -> pd.DataFrame:
    work = df.copy()
    work["volume_l"] = (work["length_cm"].fillna(0) * work["width_cm"].fillna(0) * work["height_cm"].fillna(0)) / 1000.0

    matches = work.apply(
        lambda row: engine.find_commission(
            template=row.get("template", ""),
            type_name=row.get("type_name", ""),
            subcategory=row.get("subcategory", ""),
            category=row.get("category", ""),
            product_name=row.get("name", ""),
        ),
        axis=1,
    )
    work["commission_auto_pct"] = [m.rate for m in matches]
    work["commission_source"] = [m.source for m in matches]
    work["commission_override_pct"] = work["commission_auto_pct"]

    view = work.rename(
        columns={
            "sku": "SKU",
            "name": "Название",
            "template": "Шаблон товара",
            "type_name": "Тип товара",
            "subcategory": "Подкатегория товара",
            "category": "Категория",
            "length_cm": "Длина, см",
            "width_cm": "Ширина, см",
            "height_cm": "Высота, см",
            "weight_kg": "Вес, кг",
            "cost_price": "Себестоимость, руб",
            "current_price": "Текущая цена, руб",
            "promo_price": "Цена акции, руб",
            "volume_l": "Объем, л",
            "commission_auto_pct": "Комиссия авто, %",
            "commission_override_pct": "Комиссия вручную, %",
            "commission_source": "Источник комиссии",
        }
    )
    return view[
        [
            "SKU",
            "Название",
            "Шаблон товара",
            "Тип товара",
            "Подкатегория товара",
            "Категория",
            "Длина, см",
            "Ширина, см",
            "Высота, см",
            "Вес, кг",
            "Объем, л",
            "Себестоимость, руб",
            "Текущая цена, руб",
            "Цена акции, руб",
            "Комиссия авто, %",
            "Комиссия вручную, %",
            "Источник комиссии",
        ]
    ]


def calc_destination_tariffs(engine: TariffEngine, origin_zone: str, volume_l: float, zone_mix: Dict[str, float]) -> Dict[str, float]:
    lm_out = lm_back = 0.0
    for destination, share in zone_mix.items():
        if share <= 0:
            continue
        lm_out += engine.last_mile_cost(origin_zone, destination, volume_l, return_trip=False) * share
        lm_back += engine.last_mile_cost(origin_zone, destination, volume_l, return_trip=True) * share
    return {
        "last_mile_out": lm_out,
        "last_mile_back": lm_back,
    }


def main() -> None:
    conn = init_db()
    data = load_data()
    engine = TariffEngine(data)

    st.title("Лемана Про — юнит-экономика FBS / FBO")
    st.caption("Единый стандарт расчета: отдельные тарифы в data/, FBS и FBO в одной модели, ручная правка комиссии и прозрачная математика.")

    with st.sidebar:
        st.subheader("Параметры модели")
        scheme = st.radio("Схема", ["FBS", "FBO"], horizontal=True)
        tax_regime = st.selectbox("Налогообложение", list(TAX_OPTIONS.keys()))
        target_margin_pct = st.slider("Целевая маржа после налога, %", 0, 60, 20)

        st.divider()
        st.subheader("Коммерческие проценты")
        acquiring_pct = st.number_input("Эквайринг, %", 0.0, 10.0, 1.5, 0.1)
        marketing_pct = st.number_input("Маркетинг, %", 0.0, 30.0, 3.0, 0.1)
        early_payout_pct = st.number_input("Ранняя выплата, %", 0.0, 10.0, 0.0, 0.1)
        marketplace_discount_pct = st.number_input("Скидка площадки / субсидия, %", 0.0, 50.0, 0.0, 0.1)
        promo_share_pct = st.slider("Доля продаж по акции, %", 0, 100, 30)

        st.divider()
        st.subheader("Качество продаж")
        buyout_pct = st.slider("Выкуп, %", 1, 100, 92)
        return_after_buyout_pct = st.slider("Возвраты после выкупа, %", 0, 50, 4)

        st.divider()
        st.subheader("Прочие расходы")
        extra_costs_rub = st.number_input("Доп. расходы на 1 ед., руб", 0.0, 100000.0, 0.0, 10.0)
        fbo_inbound_cost = st.number_input("FBO: входящая логистика до склада, руб/ед.", 0.0, 100000.0, 0.0, 10.0)

        st.divider()
        st.subheader("Логистика FBS")
        warehouse_label = st.selectbox(
            "Склад / регион пикапа",
            ["Москва / МО (зона 9)"],
            index=0,
            help="Для вашего кейса FBS зафиксирован склад в Москве. В тарифных файлах это используется как origin zone = 9.",
        )
        origin_zone = "9"

        zone_mode = st.radio("Последняя миля", ["Средняя модель", "Одна зона"], index=0)
        if zone_mode == "Средняя модель":
            moscow_mo = st.slider("Москва и МО, %", 0, 100, 70)
            spb_lo = st.slider("СПБ и ЛО, %", 0, 100, 20)
            regions_total = max(0, 100 - moscow_mo - spb_lo)
            st.caption(f"Остаток автоматически уходит в регионы: {regions_total}%")
            default_zone = st.selectbox("Зона для регионов", [f"Зона {i}" for i in range(1, 11)], index=2)
            zone_mix = {
                "Москва и МО": moscow_mo / 100.0,
                "СПБ и ЛО": spb_lo / 100.0,
                default_zone: regions_total / 100.0,
            }
        else:
            single_zone = st.selectbox("Одна зона доставки", ["Москва и МО", "СПБ и ЛО"] + [f"Зона {i}" for i in range(1, 11)], index=0)
            zone_mix = {single_zone: 1.0}

    # Catalog management
    with st.expander("1. Каталог товаров", expanded=True):
        c1, c2, c3 = st.columns([1, 1, 2])
        with c1:
            dim_unit = st.selectbox("Единицы размеров", ["см", "мм"], index=0)
        with c2:
            weight_unit = st.selectbox("Единицы веса", ["кг", "г"], index=0)
        with c3:
            template_df = build_import_template()
            st.download_button(
                "Скачать шаблон импорта",
                data=template_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="lemanpro_import_template.csv",
                mime="text/csv",
            )

        upload = st.file_uploader(
            "Загрузите Excel/CSV: SKU, Название, Шаблон товара, Тип товара, Подкатегория товара, Категория, Длина, Ширина, Высота, Вес, Себестоимость, Текущая цена, Цена акции",
            type=["xlsx", "xls", "csv"],
        )
        if upload is not None:
            if upload.name.lower().endswith(".csv"):
                raw_df = pd.read_csv(upload)
            else:
                raw_df = pd.read_excel(upload)
            raw_df = normalize_dimensions(raw_df.copy(), dim_unit=dim_unit, weight_unit=weight_unit)

            if st.button("Сохранить каталог в базу", type="primary"):
                save_catalog_to_db(conn, raw_df)
                st.success("Каталог сохранен.")

        catalog = read_catalog(conn)
        if catalog.empty:
            st.info("Каталог пока пуст. Загрузите файл или начните с шаблона.")
            return

        st.dataframe(catalog, use_container_width=True, hide_index=True)

    # Editable calculation base
    st.subheader("2. Параметры SKU и ручная правка комиссии")
    editor_df = prepare_editor_df(catalog, engine)
    edited = st.data_editor(
        editor_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Комиссия авто, %": st.column_config.NumberColumn(format="%.2f", disabled=True),
            "Комиссия вручную, %": st.column_config.NumberColumn(format="%.2f", help="Можно править прямо здесь."),
            "Объем, л": st.column_config.NumberColumn(format="%.2f", disabled=True),
            "Источник комиссии": st.column_config.TextColumn(disabled=True),
        },
        key="sku_editor",
    )

    run_calc = st.button("3. Рассчитать юнит-экономику", type="primary")
    if not run_calc:
        return

    results = []
    for _, row in edited.iterrows():
        volume_l = max(safe_float(row["Объем, л"]), 0.0)
        if volume_l <= 0:
            volume_l = (
                safe_float(row["Длина, см"]) * safe_float(row["Ширина, см"]) * safe_float(row["Высота, см"])
            ) / 1000.0

        commission_pct = safe_float(row["Комиссия вручную, %"])
        if commission_pct == 0 and safe_float(row["Комиссия авто, %"]) > 0:
            commission_pct = safe_float(row["Комиссия авто, %"])

        zero_out = engine.zero_mile_cost(volume_l, return_trip=False) if scheme == "FBS" else 0.0
        zero_back = engine.zero_mile_cost(volume_l, return_trip=True) if scheme == "FBS" else 0.0
        zone_tariffs = calc_destination_tariffs(engine, origin_zone, volume_l, zone_mix)

        calc = compute_unit_economics(
            scheme=scheme,
            current_price=safe_float(row["Текущая цена, руб"]),
            promo_price=safe_float(row["Цена акции, руб"]),
            promo_share_pct=promo_share_pct,
            cost_price=safe_float(row["Себестоимость, руб"]),
            commission_pct=commission_pct,
            acquiring_pct=acquiring_pct,
            marketing_pct=marketing_pct,
            early_payout_pct=early_payout_pct,
            extra_costs_rub=extra_costs_rub,
            target_margin_pct=target_margin_pct,
            tax_regime=tax_regime,
            buyout_pct=buyout_pct,
            return_after_buyout_pct=return_after_buyout_pct,
            zero_mile_out=zero_out,
            zero_mile_back=zero_back,
            last_mile_out=zone_tariffs["last_mile_out"],
            last_mile_back=zone_tariffs["last_mile_back"],
            fbo_inbound_cost=fbo_inbound_cost,
            marketplace_discount_pct=marketplace_discount_pct,
        )

        results.append(
            {
                "SKU": row["SKU"],
                "Название": row["Название"],
                "Схема": scheme,
                "Шаблон товара": row["Шаблон товара"],
                "Тип товара": row["Тип товара"],
                "Подкатегория товара": row["Подкатегория товара"],
                "Категория": row["Категория"],
                "Объем, л": round(volume_l, 2),
                "Себестоимость, руб": round(safe_float(row["Себестоимость, руб"]), 2),
                "Текущая цена, руб": round(safe_float(row["Текущая цена, руб"]), 2),
                "Цена акции, руб": round(safe_float(row["Цена акции, руб"]), 2),
                "Комиссия, %": round(commission_pct, 2),
                "Источник комиссии": row["Источник комиссии"],
                "Нулевая миля, руб": round(zero_out, 2),
                "Возврат до СЦ, руб": round(zero_back, 2),
                "Последняя миля, руб": round(zone_tariffs["last_mile_out"], 2),
                "Возврат последней мили, руб": round(zone_tariffs["last_mile_back"], 2),
                "Ожидаемая логистика, руб": calc["logistics_expected"],
                "Эффективная текущая цена, руб": calc["effective_current_price"],
                "Эффективная акционная цена, руб": calc["effective_promo_price"],
                "Средняя реализованная цена, руб": calc["realized_price"],
                "Доля успешной реализации, %": calc["sale_realization_rate_pct"],
                "Ожидаемая выручка, руб": calc["expected_revenue"],
                "Переменные расходы, руб": calc["variable_costs"],
                "Прибыль до налога, руб": calc["profit_before_tax"],
                "Налог, руб": calc["tax"],
                "Прибыль после налога, руб": calc["profit_after_tax"],
                "Маржа до налога, %": calc["margin_before_tax_pct"],
                "Маржа после налога, %": calc["margin_after_tax_pct"],
                "Наценка, %": calc["markup_pct"],
                "Реком. средняя цена, руб": calc["recommended_realized_price"],
                "Реком. текущая цена, руб": calc["recommended_current_price"],
                "Реком. цена акции, руб": calc["recommended_promo_price"],
                "Прибыль по реком. цене, руб": calc["recommended_profit_after_tax"],
                "Маржа по реком. цене, %": calc["recommended_margin_after_tax_pct"],
                "Флаг": flag_text(calc["profit_after_tax"], calc["margin_after_tax_pct"], str(row["Источник комиссии"]), volume_l),
            }
        )

    result_df = pd.DataFrame(results)
    bad_mask = result_df["Флаг"] != "ok"

    total_revenue = result_df["Ожидаемая выручка, руб"].sum()
    total_profit = result_df["Прибыль после налога, руб"].sum()
    total_profit_recommended = result_df["Прибыль по реком. цене, руб"].sum()
    weighted_margin = (total_profit / total_revenue * 100.0) if total_revenue > 0 else 0.0
    avg_commission = result_df["Комиссия, %"].mean() if not result_df.empty else 0.0

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("SKU", f"{len(result_df):,}".replace(",", " "))
    k2.metric("Проблемных SKU", int(bad_mask.sum()))
    k3.metric("Ожидаемая прибыль", f"{total_profit:,.0f} ₽".replace(",", " "))
    k4.metric("Маржа портфеля", f"{weighted_margin:.1f}%")
    k5.metric("Средняя комиссия", f"{avg_commission:.1f}%")

    st.subheader("4. Результат")
    st.dataframe(
        result_df.style.apply(
            lambda row: ["background-color: #ffe5e5" if row["Флаг"] != "ok" else "" for _ in row],
            axis=1,
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("5. Выгрузки")
    employee_template = result_df[
        [
            "SKU",
            "Название",
            "Текущая цена, руб",
            "Цена акции, руб",
            "Реком. текущая цена, руб",
            "Реком. цена акции, руб",
            "Комиссия, %",
            "Ожидаемая логистика, руб",
            "Маржа после налога, %",
            "Маржа по реком. цене, %",
            "Флаг",
        ]
    ].copy()

    st.download_button(
        "Скачать полный результат CSV",
        data=result_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="lemanpro_unit_economics.csv",
        mime="text/csv",
    )
    st.download_button(
        "Скачать шаблон для сотрудников CSV",
        data=employee_template.to_csv(index=False).encode("utf-8-sig"),
        file_name="lemanpro_staff_template.csv",
        mime="text/csv",
    )

    with st.expander("Как считается логистика и математика"):
        st.markdown(
            """
            **FBS**
            - Нулевая миля всегда считается на каждую отправку.
            - Последняя миля всегда считается на каждую отправку.
            - При невыкупе добавляется возврат последней мили + возврат до СЦ.
            - При возврате после выкупа добавляется возврат последней мили + возврат до СЦ.

            **FBO**
            - Входящая логистика на склад задается вручную на 1 единицу.
            - Последняя миля и возврат последней мили считаются по тарифным таблицам.
            - Возврат до СЦ в FBO не добавляется.

            **Финансовая логика**
            - Ожидаемая выручка = средняя реализованная цена × доля успешной реализации.
            - Переменные проценты применяются к ожидаемой выручке.
            - Маржа считается после налога от ожидаемой выручки.
            - Рекомендованная цена ищется численно так, чтобы выйти на целевую маржу после налога.
            """
        )


if __name__ == "__main__":
    main()
