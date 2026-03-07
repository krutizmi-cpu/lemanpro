
import io
import math
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


st.set_page_config(page_title="Лемана Про — простая юнит-экономика", layout="wide", page_icon="📦")


# =========================
# Helpers
# =========================

def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def norm_text(x) -> str:
    s = clean_text(x).lower().replace("ё", "е")
    s = re.sub(r"[^a-zа-я0-9\s\-_/]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def to_float(x, default=0.0) -> float:
    if pd.isna(x):
        return default
    s = str(x).strip().replace(" ", "").replace("%", "").replace(",", ".")
    if not s:
        return default
    try:
        return float(s)
    except Exception:
        return default


def pick_col(df: pd.DataFrame, variants: List[str]) -> Optional[str]:
    cols = {norm_text(c): c for c in df.columns}
    for variant in variants:
        key = norm_text(variant)
        if key in cols:
            return cols[key]
    for c in df.columns:
        c_norm = norm_text(c)
        if any(norm_text(v) in c_norm for v in variants):
            return c
    return None


def parse_range_text(text: str) -> Tuple[float, Optional[float]]:
    s = norm_text(text).replace("л.", " л")
    nums = re.findall(r"\d+(?:[.,]\d+)?", s.replace(",", "."))
    if not nums:
        return 0.0, None
    vals = [float(x) for x in nums]
    if len(vals) == 1:
        return vals[0], None
    lo, hi = vals[0], vals[1]
    if hi < lo:
        lo, hi = hi, lo
    return lo, hi


def volume_l(length_cm: float, width_cm: float, height_cm: float) -> float:
    return max(length_cm, 0) * max(width_cm, 0) * max(height_cm, 0) / 1000.0


def round_price(x: float, step: int) -> float:
    if step <= 1:
        return round(x, 2)
    return math.ceil(x / step) * step


# =========================
# Tax model
# =========================

TAX_OPTIONS = {
    "Без налога": ("none", 0.0),
    "УСН Доходы 6%": ("revenue", 0.06),
    "УСН Доходы-Расходы 15%": ("profit", 0.15),
    "ОСНО 25% от прибыли": ("profit", 0.25),
}


def calc_tax(revenue: float, costs_before_tax: float, regime: str) -> float:
    mode, rate = TAX_OPTIONS.get(regime, ("none", 0.0))
    profit_before_tax = revenue - costs_before_tax
    if mode == "none":
        return 0.0
    if mode == "revenue":
        return max(revenue * rate, 0.0)
    if mode == "profit":
        return max(profit_before_tax, 0.0) * rate
    return 0.0


# =========================
# Parsers for uploaded files
# =========================

@dataclass
class CommissionRule:
    commission_pct: float
    template: str
    type_name: str
    subcategory: str
    category: str
    search_blob: str


@st.cache_data(show_spinner=False)
def load_commission_rules(file_bytes: bytes) -> List[CommissionRule]:
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet = None
    for s in xl.sheet_names:
        s_norm = norm_text(s)
        if "комиссия" in s_norm and ("fbs" in s_norm or "fbo" in s_norm or "комиссия" == s_norm):
            sheet = s
            break
    if sheet is None:
        sheet = xl.sheet_names[0]

    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
    col_comm = pick_col(df, ["Комиссия"])
    col_template = pick_col(df, ["Шаблон товара"])
    col_type = pick_col(df, ["Тип товара", "Категория товаров"])
    col_sub = pick_col(df, ["Подкатегория товара", "Подкатегория товаров"])
    col_cat = pick_col(df, ["Категория"])

    rules: List[CommissionRule] = []
    for _, row in df.iterrows():
        commission_raw = to_float(row.get(col_comm, 0))
        if commission_raw <= 0:
            continue
        commission_pct = commission_raw * 100 if commission_raw <= 1 else commission_raw
        template = clean_text(row.get(col_template, ""))
        type_name = clean_text(row.get(col_type, ""))
        subcategory = clean_text(row.get(col_sub, ""))
        category = clean_text(row.get(col_cat, ""))

        blob = " | ".join([template, type_name, subcategory, category]).strip(" |")
        if not blob:
            continue

        rules.append(
            CommissionRule(
                commission_pct=commission_pct,
                template=template,
                type_name=type_name,
                subcategory=subcategory,
                category=category,
                search_blob=norm_text(blob),
            )
        )
    return rules


@st.cache_data(show_spinner=False)
def load_logistics_tables(file_bytes: bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))

    sheet_last = None
    sheet_zero = None

    for s in xl.sheet_names:
        s_norm = norm_text(s)
        if "последняя миля" in s_norm and "возврат" not in s_norm:
            sheet_last = s
        if ("доставка до сц" in s_norm or "нулевая миля" in s_norm) and "возврат" not in s_norm:
            sheet_zero = s

    if sheet_last is None:
        raise ValueError("Не найден лист с тарифами последней мили.")
    if sheet_zero is None:
        raise ValueError("Не найден лист с тарифами нулевой мили / доставки до СЦ.")

    last_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_last)
    zero_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_zero)

    # zero mile
    zero_range_col = pick_col(zero_df, ["Объемный брейк отправления, в л"])
    zero_tariff_col = pick_col(zero_df, ["Тариф, с НДС"])
    zero_rows = []
    for _, row in zero_df.iterrows():
        lo, hi = parse_range_text(clean_text(row.get(zero_range_col, "")))
        tariff = to_float(row.get(zero_tariff_col, 0))
        if hi is None:
            hi = float("inf")
        if tariff > 0:
            zero_rows.append((lo, hi, tariff))
    zero_rows = sorted(zero_rows, key=lambda x: (x[0], x[1]))

    # last mile
    col_from = pick_col(last_df, ["Зона откуда"])
    col_to = pick_col(last_df, ["Зона куда"])
    col_break = pick_col(last_df, ["Весовой брейк отправления, в л", "Весовой брейк отправления, в л."])
    col_tariff = pick_col(last_df, ["Тариф, с НДС"])
    col_plus = pick_col(last_df, ["+1 л, с НДС"])

    last_rows = []
    for _, row in last_df.iterrows():
        zone_from = clean_text(row.get(col_from, ""))
        zone_to = clean_text(row.get(col_to, ""))
        lo, hi = parse_range_text(clean_text(row.get(col_break, "")))
        tariff = to_float(row.get(col_tariff, 0))
        plus_per_l = to_float(row.get(col_plus, 0))
        if hi is None:
            hi = float("inf")
        if zone_to and tariff > 0:
            last_rows.append((zone_from, zone_to, lo, hi, tariff, plus_per_l))

    return zero_rows, last_rows


def get_zero_mile_cost(v_l: float, zero_rows) -> float:
    for lo, hi, tariff in zero_rows:
        if v_l >= lo and v_l <= hi:
            return tariff
    if zero_rows:
        lo, hi, tariff = zero_rows[-1]
        extra = max(v_l - hi, 0) if math.isfinite(hi) else 0
        return tariff + extra * 0
    return 0.0


def get_last_mile_cost(v_l: float, zone_to: str, origin_zone: int, last_rows) -> float:
    zone_to_norm = norm_text(zone_to)
    candidates = []
    for z_from, z_to, lo, hi, tariff, plus_per_l in last_rows:
        same_from = str(origin_zone) == str(z_from).strip() if clean_text(z_from) else True
        if same_from and norm_text(z_to) == zone_to_norm:
            candidates.append((lo, hi, tariff, plus_per_l))
    if not candidates:
        for z_from, z_to, lo, hi, tariff, plus_per_l in last_rows:
            if norm_text(z_to) == zone_to_norm:
                candidates.append((lo, hi, tariff, plus_per_l))
    candidates = sorted(candidates, key=lambda x: (x[0], x[1]))
    for lo, hi, tariff, plus_per_l in candidates:
        if v_l >= lo and v_l <= hi:
            return tariff
    if candidates:
        lo, hi, tariff, plus_per_l = candidates[-1]
        if math.isfinite(hi) and plus_per_l > 0 and v_l > hi:
            return tariff + (v_l - hi) * plus_per_l
        return tariff
    return 0.0


# =========================
# Category detection
# =========================

def detect_category_by_rules(name: str, rules: List[CommissionRule]) -> Optional[CommissionRule]:
    name_norm = norm_text(name)
    if not name_norm:
        return None

    scored = []
    name_words = set(name_norm.split())

    for rule in rules:
        score = 0
        blob = rule.search_blob
        if not blob:
            continue

        template_words = set(blob.split())

        # exact/substring
        if rule.template and norm_text(rule.template) in name_norm:
            score += 50
        if rule.type_name and norm_text(rule.type_name) in name_norm:
            score += 30
        if rule.subcategory and norm_text(rule.subcategory) in name_norm:
            score += 20
        if rule.category and norm_text(rule.category) in name_norm:
            score += 10

        # token overlap
        overlap = len(name_words & template_words)
        score += overlap * 2

        if score > 0:
            scored.append((score, rule))

    if not scored:
        return None
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[0][1]


def detect_category_ai(name: str, rules: List[CommissionRule], api_key: str) -> Optional[CommissionRule]:
    if not api_key or OpenAI is None or not rules:
        return None
    variants = []
    for i, r in enumerate(rules[:400], start=1):
        label = " | ".join(x for x in [r.template, r.type_name, r.subcategory, r.category] if x)
        variants.append(f"{i}. {label} -> {r.commission_pct:.2f}%")
    prompt = (
        "Ты классификатор товаров для Лемана Про.\n"
        "Выбери один самый подходящий вариант категории для товара.\n"
        "Ответь только номером варианта.\n\n"
        f"Товар: {name}\n\n"
        "Варианты:\n" + "\n".join(variants)
    )
    try:
        client = OpenAI(api_key=api_key)
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Ты выбираешь один лучший вариант категории товара."},
                {"role": "user", "content": prompt},
            ],
            temperature=0,
            max_tokens=10,
        )
        answer = clean_text(resp.choices[0].message.content)
        m = re.search(r"\d+", answer)
        if not m:
            return None
        idx = int(m.group(0)) - 1
        if 0 <= idx < len(rules[:400]):
            return rules[idx]
    except Exception:
        return None
    return None


# =========================
# Price solver
# =========================

def economics_at_price(
    price: float,
    commission_pct: float,
    expected_logistics: float,
    cost: float,
    acquiring_pct: float,
    marketing_pct: float,
    extra_cost: float,
    tax_regime: str,
):
    pct_costs = (commission_pct + acquiring_pct + marketing_pct) / 100.0
    revenue = price
    variable_pct_costs = revenue * pct_costs
    costs_before_tax = cost + expected_logistics + extra_cost + variable_pct_costs
    tax = calc_tax(revenue, costs_before_tax, tax_regime)
    profit_after_tax = revenue - costs_before_tax - tax
    margin_after_tax = (profit_after_tax / revenue * 100.0) if revenue > 0 else 0.0
    markup = ((price / cost - 1.0) * 100.0) if cost > 0 else 0.0
    return {
        "revenue": revenue,
        "variable_pct_costs": variable_pct_costs,
        "costs_before_tax": costs_before_tax,
        "tax": tax,
        "profit_after_tax": profit_after_tax,
        "margin_after_tax": margin_after_tax,
        "markup_pct": markup,
    }


def solve_recommended_price(
    target_margin_pct: float,
    commission_pct: float,
    expected_logistics: float,
    cost: float,
    acquiring_pct: float,
    marketing_pct: float,
    extra_cost: float,
    tax_regime: str,
    rounding_step: int,
) -> float:
    # binary search on price that gives target margin after tax
    low = max(cost + expected_logistics + extra_cost, 1.0)
    high = max(low * 5, 100.0)
    for _ in range(40):
        e = economics_at_price(high, commission_pct, expected_logistics, cost, acquiring_pct, marketing_pct, extra_cost, tax_regime)
        if e["margin_after_tax"] >= target_margin_pct:
            break
        high *= 1.8

    for _ in range(80):
        mid = (low + high) / 2
        e = economics_at_price(mid, commission_pct, expected_logistics, cost, acquiring_pct, marketing_pct, extra_cost, tax_regime)
        if e["margin_after_tax"] >= target_margin_pct:
            high = mid
        else:
            low = mid
    return round_price(high, rounding_step)


# =========================
# UI
# =========================

st.title("Лемана Про — простая юнит-экономика")
st.caption("Загрузите каталог, при необходимости загрузите файл тарифов/комиссий, и получите одну рекомендованную цену продажи.")

with st.sidebar:
    st.subheader("Параметры расчета")
    tax_regime = st.selectbox("Налог", list(TAX_OPTIONS.keys()), index=0)
    target_margin_pct = st.slider("Целевая маржа после налога, %", 0, 60, 20)
    acquiring_pct = st.number_input("Эквайринг, %", min_value=0.0, max_value=10.0, value=1.5, step=0.1)
    marketing_pct = st.number_input("Маркетинг / реклама, %", min_value=0.0, max_value=30.0, value=0.0, step=0.1)
    extra_cost = st.number_input("Прочие расходы на 1 шт., руб", min_value=0.0, max_value=100000.0, value=0.0, step=10.0)
    rounding_step = st.selectbox("Округление цены, руб", [1, 5, 10, 50, 100], index=2)

    st.divider()
    st.subheader("FBS — настройки логистики")
    origin_zone = st.number_input("Зона откуда (Москва склад = 9)", min_value=1, max_value=20, value=9, step=1)

    c1, c2, c3 = st.columns(3)
    with c1:
        share_moscow = st.number_input("Москва/МО, %", min_value=0.0, max_value=100.0, value=70.0, step=1.0)
    with c2:
        share_spb = st.number_input("СПБ/ЛО, %", min_value=0.0, max_value=100.0, value=20.0, step=1.0)
    with c3:
        share_regions = st.number_input("Регионы, %", min_value=0.0, max_value=100.0, value=10.0, step=1.0)

    st.divider()
    st.subheader("Категоризация")
    use_ai = st.checkbox("Использовать OpenAI, если ключ есть", value=False)
    openai_key = st.text_input("OpenAI API key", type="password", value=st.secrets.get("OPENAI_API_KEY", "") if hasattr(st, "secrets") else "")

zone_sum = share_moscow + share_spb + share_regions
if zone_sum <= 0:
    st.error("Сумма долей зон должна быть больше 0%.")
    st.stop()

zone_weights = {
    "Москва и МО": share_moscow / zone_sum,
    "СПБ и ЛО": share_spb / zone_sum,
    "Регионы": share_regions / zone_sum,
}

st.subheader("1. Загрузите файлы")
catalog_file = st.file_uploader("Каталог товаров Excel: артикул, наименование, длина, ширина, высота, вес, себестоимость", type=["xlsx", "xls"])
commission_file = st.file_uploader("Файл комиссий Excel", type=["xlsx", "xls"])
logistics_file = st.file_uploader("Файл логистики Excel", type=["xlsx", "xls"])

st.info(
    "Минимум нужен каталог. "
    "Если не загрузить тарифы и комиссии, приложение не сможет корректно посчитать результат."
)

if catalog_file and commission_file and logistics_file:
    try:
        rules = load_commission_rules(commission_file.getvalue())
        zero_rows, last_rows = load_logistics_tables(logistics_file.getvalue())
    except Exception as e:
        st.error(f"Ошибка чтения файлов комиссий/логистики: {e}")
        st.stop()

    try:
        catalog = pd.read_excel(catalog_file)
    except Exception as e:
        st.error(f"Ошибка чтения каталога: {e}")
        st.stop()

    col_sku = pick_col(catalog, ["Артикул", "SKU", "sku"])
    col_name = pick_col(catalog, ["Наименование", "Название", "Товар"])
    col_len = pick_col(catalog, ["Длина"])
    col_wid = pick_col(catalog, ["Ширина"])
    col_hei = pick_col(catalog, ["Высота"])
    col_wgt = pick_col(catalog, ["Вес"])
    col_cost = pick_col(catalog, ["Себестоимость", "Закупка", "Себес"])

    missing = [x for x in [
        ("Артикул", col_sku), ("Наименование", col_name), ("Длина", col_len),
        ("Ширина", col_wid), ("Высота", col_hei), ("Вес", col_wgt), ("Себестоимость", col_cost)
    ] if x[1] is None]

    if missing:
        st.error("В каталоге не найдены обязательные колонки: " + ", ".join(name for name, _ in missing))
        st.stop()

    results = []
    progress = st.progress(0)

    for idx, row in catalog.iterrows():
        sku = clean_text(row.get(col_sku, ""))
        name = clean_text(row.get(col_name, ""))
        length_cm = to_float(row.get(col_len, 0))
        width_cm = to_float(row.get(col_wid, 0))
        height_cm = to_float(row.get(col_hei, 0))
        weight_kg = to_float(row.get(col_wgt, 0))
        cost = to_float(row.get(col_cost, 0))

        vol = volume_l(length_cm, width_cm, height_cm)

        rule = detect_category_by_rules(name, rules)
        if use_ai and openai_key:
            ai_rule = detect_category_ai(name, rules, openai_key)
            if ai_rule is not None:
                rule = ai_rule

        commission_pct = rule.commission_pct if rule else 0.0
        template = rule.template if rule else "Не определено"
        category = rule.category if rule else "Не определено"

        zero_mile = get_zero_mile_cost(vol, zero_rows)
        lm_moscow = get_last_mile_cost(vol, "Москва и МО", origin_zone, last_rows)
        lm_spb = get_last_mile_cost(vol, "СПБ и ЛО", origin_zone, last_rows)
        lm_regions = get_last_mile_cost(vol, "Регионы", origin_zone, last_rows)

        expected_last_mile = (
            zone_weights["Москва и МО"] * lm_moscow
            + zone_weights["СПБ и ЛО"] * lm_spb
            + zone_weights["Регионы"] * lm_regions
        )

        expected_logistics = zero_mile + expected_last_mile

        rec_price = solve_recommended_price(
            target_margin_pct=target_margin_pct,
            commission_pct=commission_pct,
            expected_logistics=expected_logistics,
            cost=cost,
            acquiring_pct=acquiring_pct,
            marketing_pct=marketing_pct,
            extra_cost=extra_cost,
            tax_regime=tax_regime,
            rounding_step=rounding_step,
        )

        econ = economics_at_price(
            price=rec_price,
            commission_pct=commission_pct,
            expected_logistics=expected_logistics,
            cost=cost,
            acquiring_pct=acquiring_pct,
            marketing_pct=marketing_pct,
            extra_cost=extra_cost,
            tax_regime=tax_regime,
        )

        results.append({
            "Артикул": sku,
            "Наименование": name,
            "Шаблон товара": template,
            "Категория": category,
            "Комиссия, %": round(commission_pct, 2),
            "Длина, см": round(length_cm, 2),
            "Ширина, см": round(width_cm, 2),
            "Высота, см": round(height_cm, 2),
            "Вес, кг": round(weight_kg, 3),
            "Объем, л": round(vol, 3),
            "Нулевая миля, руб": round(zero_mile, 2),
            "Последняя миля Москва/МО, руб": round(lm_moscow, 2),
            "Последняя миля СПБ/ЛО, руб": round(lm_spb, 2),
            "Последняя миля Регионы, руб": round(lm_regions, 2),
            "Средняя последняя миля, руб": round(expected_last_mile, 2),
            "Средняя логистика, руб": round(expected_logistics, 2),
            "Себестоимость, руб": round(cost, 2),
            "Рекомендованная цена, руб": round(rec_price, 2),
            "Прибыль после налога, руб": round(econ["profit_after_tax"], 2),
            "Маржа после налога, %": round(econ["margin_after_tax"], 2),
            "Наценка, %": round(econ["markup_pct"], 2),
        })

        progress.progress((idx + 1) / max(len(catalog), 1))

    progress.empty()

    res_df = pd.DataFrame(results)

    st.subheader("2. Результат")
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.metric("SKU", len(res_df))
    with k2:
        st.metric("Средняя комиссия, %", round(res_df["Комиссия, %"].mean(), 2) if len(res_df) else 0)
    with k3:
        st.metric("Средняя логистика, руб", round(res_df["Средняя логистика, руб"].mean(), 2) if len(res_df) else 0)
    with k4:
        st.metric("Средняя рекомендованная цена, руб", round(res_df["Рекомендованная цена, руб"].mean(), 2) if len(res_df) else 0)

    st.dataframe(res_df, use_container_width=True, height=600)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        res_df.to_excel(writer, sheet_name="Результат", index=False)
    st.download_button(
        "Скачать результат в Excel",
        data=output.getvalue(),
        file_name="lemanpro_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Как считает приложение"):
        st.markdown(
            """
            **Логика простая:**
            1. По наименованию товара определяется шаблон/категория из файла комиссий.  
            2. По габаритам считается объем в литрах.  
            3. По объему считается:
               - нулевая миля,
               - последняя миля по Москве/МО,
               - последняя миля по СПБ/ЛО,
               - последняя миля по регионам.  
            4. Из долей зон строится **одна средняя логистика**.  
            5. Система подбирает **одну рекомендованную цену**, при которой достигается целевая маржа.
            """
        )
else:
    st.warning("Загрузите 3 файла: каталог, комиссии, логистику.")
