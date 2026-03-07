import os
import re
import math
import tempfile
from difflib import SequenceMatcher
from io import BytesIO

import pandas as pd
import streamlit as st

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

st.set_page_config(page_title="Лемана Про — Юнит-экономика", layout="wide", page_icon="📦")

DEFAULT_RATES_PATHS = [
    "lemanpro_rates.xlsx",
    "data/lemanpro_rates.xlsx",
    "Коммерческие комиссии (февраль) (3).xlsx",
]

STOPWORDS = {
    "для", "и", "или", "в", "во", "на", "по", "с", "со", "из", "к", "от", "до", "под", "над",
    "the", "and", "or", "of", "a", "an", "to", "with", "без", "за", "при"
}


# ---------- helpers ----------
def normalize_text(text: str) -> str:
    text = str(text or "").lower().strip()
    text = text.replace("ё", "е")
    text = re.sub(r"[_/\\|]+", " ", text)
    text = re.sub(r"[^a-zа-я0-9\s-]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def tokenize(text: str):
    tokens = re.findall(r"[a-zа-я0-9]+", normalize_text(text))
    return [t for t in tokens if len(t) > 2 and t not in STOPWORDS]


def safe_float(value, default=0.0):
    try:
        if value is None or (isinstance(value, float) and math.isnan(value)):
            return default
        return float(str(value).replace(",", ".").strip())
    except Exception:
        return default


def normalize_dimension(value, unit: str):
    v = safe_float(value, 0.0)
    unit = str(unit).strip().lower()
    if unit in ("мм", "mm"):
        return v / 10.0
    return v


def normalize_weight(value, unit: str):
    v = safe_float(value, 0.0)
    unit = str(unit).strip().lower()
    if unit in ("г", "gr", "g", "гр"):
        return v / 1000.0
    return v


def find_existing_rates_file():
    for path in DEFAULT_RATES_PATHS:
        if os.path.exists(path):
            return path
    return None


def make_template_excel_bytes() -> bytes:
    df = pd.DataFrame([
        {
            "SKU": "ART-001",
            "Наименование": "Смеситель для кухни хром",
            "Длина": 35,
            "Ширина": 18,
            "Высота": 8,
            "Вес": 1.2,
            "Себестоимость": 2450,
            "Текущая цена": 3990,
        },
        {
            "SKU": "ART-002",
            "Наименование": "Лампа настольная черная",
            "Длина": 42,
            "Ширина": 16,
            "Высота": 16,
            "Вес": 2.1,
            "Себестоимость": 1800,
            "Текущая цена": 0,
        },
    ])
    info = pd.DataFrame({
        "Поле": [
            "SKU", "Наименование", "Длина", "Ширина", "Высота", "Вес", "Себестоимость", "Текущая цена"
        ],
        "Обязательно": ["Да", "Да", "Да", "Да", "Да", "Да", "Да", "Нет"],
        "Комментарий": [
            "Артикул продавца",
            "Название товара для определения категории",
            "Габарит товара или упаковки",
            "Габарит товара или упаковки",
            "Габарит товара или упаковки",
            "Вес товара или упаковки",
            "Полная себестоимость единицы",
            "Необязательное поле для сравнения с текущей ценой",
        ],
    })
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Шаблон", index=False)
        info.to_excel(writer, sheet_name="Описание", index=False)
    bio.seek(0)
    return bio.getvalue()


@st.cache_data(show_spinner=False)
def load_standard_ratebook_from_path(path: str):
    xls = pd.ExcelFile(path)
    sheets = xls.sheet_names

    required = {
        "Комиссия_FBS и FBO",
        "Тарифы (Последняя миля)",
        "Тарифы (возврат Последняя миля)",
        "Тарифы (Доставка до СЦ)",
        "Тарифы (Возврат Доставка до СЦ)",
    }
    missing = required - set(sheets)
    if missing:
        raise ValueError(f"В файле не хватает листов: {', '.join(sorted(missing))}")

    commission = pd.read_excel(path, sheet_name="Комиссия_FBS и FBO")
    last_mile = pd.read_excel(path, sheet_name="Тарифы (Последняя миля)")
    return_last_mile = pd.read_excel(path, sheet_name="Тарифы (возврат Последняя миля)")
    to_sc = pd.read_excel(path, sheet_name="Тарифы (Доставка до СЦ)")
    return_to_sc = pd.read_excel(path, sheet_name="Тарифы (Возврат Доставка до СЦ)")

    commission = commission.rename(columns={
        commission.columns[0]: "commission",
        commission.columns[1]: "template",
        commission.columns[2]: "type",
        commission.columns[3]: "subcategory",
        commission.columns[4]: "category",
    }).copy()
    commission["commission"] = pd.to_numeric(commission["commission"], errors="coerce").fillna(0.0)
    for col in ["template", "type", "subcategory", "category"]:
        commission[col] = commission[col].fillna("").astype(str)
    commission["search_text"] = (
        commission["template"] + " " + commission["type"] + " " + commission["subcategory"] + " " + commission["category"]
    ).map(normalize_text)
    commission["tokens"] = commission["search_text"].map(tokenize)

    def prep_last_mile(df: pd.DataFrame):
        df = df.rename(columns={
            df.columns[0]: "zone_from",
            df.columns[1]: "zone_to",
            df.columns[2]: "break_text",
            df.columns[4]: "base_tariff",
            df.columns[5]: "per_liter",
        }).copy()
        df["zone_from"] = df["zone_from"].astype(str).str.strip()
        df["zone_to"] = df["zone_to"].astype(str).str.strip()
        df["base_tariff"] = pd.to_numeric(df["base_tariff"], errors="coerce").fillna(0.0)
        df["per_liter"] = pd.to_numeric(df["per_liter"], errors="coerce").fillna(0.0)
        bounds = df["break_text"].astype(str).str.extract(r"(\d+(?:[\.,]\d+)?)\s*до\s*(\d+(?:[\.,]\d+)?)")
        df["break_from"] = bounds[0].str.replace(",", ".", regex=False).astype(float)
        df["break_to"] = bounds[1].str.replace(",", ".", regex=False).astype(float)
        return df

    def prep_sc(df: pd.DataFrame):
        df = df.rename(columns={
            df.columns[0]: "break_text",
            df.columns[1]: "base_tariff",
        }).copy()
        df["base_tariff"] = pd.to_numeric(df["base_tariff"], errors="coerce").fillna(0.0)
        bounds = df["break_text"].astype(str).str.extract(r"(\d+(?:[\.,]\d+)?)\s*до\s*(\d+(?:[\.,]\d+)?)")
        df["break_from"] = bounds[0].str.replace(",", ".", regex=False).astype(float)
        df["break_to"] = bounds[1].str.replace(",", ".", regex=False).astype(float)
        return df

    return {
        "commission": commission,
        "last_mile": prep_last_mile(last_mile),
        "return_last_mile": prep_last_mile(return_last_mile),
        "to_sc": prep_sc(to_sc),
        "return_to_sc": prep_sc(return_to_sc),
        "source_path": path,
    }


@st.cache_data(show_spinner=False)
def load_standard_ratebook_from_bytes(file_bytes: bytes):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(file_bytes)
    tmp.flush()
    tmp.close()
    return load_standard_ratebook_from_path(tmp.name)


def find_break_tariff(volume_l: float, df: pd.DataFrame, zone_from: str = None, zone_to: str = None):
    use = df.copy()
    if zone_from is not None and "zone_from" in use.columns:
        use = use[use["zone_from"].astype(str) == str(zone_from)]
    if zone_to is not None and "zone_to" in use.columns:
        use = use[use["zone_to"].astype(str).str.strip() == str(zone_to).strip()]
    if use.empty:
        return 0.0

    matched = use[(use["break_from"] <= volume_l) & (volume_l <= use["break_to"])]
    if matched.empty:
        matched = use.sort_values(["break_to"]).tail(1)
    row = matched.iloc[0]
    extra_l = max(0.0, volume_l - safe_float(row.get("break_to", 0.0), 0.0))
    return round(safe_float(row["base_tariff"], 0.0) + extra_l * safe_float(row.get("per_liter", 0.0), 0.0), 2)


def classify_by_rules(name: str, commission_df: pd.DataFrame):
    name_norm = normalize_text(name)
    name_tokens = set(tokenize(name_norm))
    best_idx = None
    best_score = -1.0

    for idx, row in commission_df.iterrows():
        row_tokens = set(row["tokens"]) if isinstance(row["tokens"], list) else set()
        overlap = len(name_tokens & row_tokens)
        seq = SequenceMatcher(None, name_norm, row["search_text"]).ratio()
        template_name = normalize_text(str(row["template"]).split("_", 1)[-1])
        prefix_bonus = 1.5 if template_name and template_name in name_norm else 0.0
        score = overlap * 1.8 + seq + prefix_bonus
        if score > best_score:
            best_score = score
            best_idx = idx

    if best_idx is None:
        return None, 0.0
    return commission_df.loc[best_idx], best_score


def classify_with_ai(name: str, commission_df: pd.DataFrame, api_key: str):
    if not api_key or OpenAI is None:
        return None
    candidates = commission_df[["template", "type", "subcategory", "category", "commission"]].head(250)
    lines = []
    for _, row in candidates.iterrows():
        lines.append(
            f"Шаблон: {row['template']} | Тип: {row['type']} | Подкатегория: {row['subcategory']} | Категория: {row['category']} | Комиссия: {row['commission']:.4f}"
        )
    prompt = "\n".join(lines)
    try:
        client = OpenAI(api_key=api_key)
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0,
            max_tokens=120,
            messages=[
                {"role": "system", "content": "Ты классификатор товаров Лемана Про. Выбери одну наиболее подходящую строку. Ответ только названием шаблона товара из списка."},
                {"role": "user", "content": f"Товар: {name}\n\nДоступные варианты:\n{prompt}"},
            ],
        )
        answer = (resp.choices[0].message.content or "").strip()
        match = commission_df[commission_df["template"].astype(str).str.strip() == answer]
        if not match.empty:
            return match.iloc[0]
    except Exception:
        return None
    return None


def calc_tax(revenue: float, total_costs: float, regime: str):
    profit_before = revenue - total_costs
    regimes = {
        "ОСНО (25% от прибыли)": ("profit", 0.25),
        "УСН Доходы (6%)": ("revenue", 0.06),
        "УСН Доходы-Расходы (15%)": ("profit", 0.15),
        "АУСН (8% от дохода)": ("revenue", 0.08),
        "УСН НДС 5%": ("revenue", 0.05),
        "УСН НДС 7%": ("revenue", 0.07),
    }
    mode, rate = regimes.get(regime, ("profit", 0.0))
    tax = revenue * rate if mode == "revenue" else max(profit_before, 0) * rate
    profit_after = profit_before - tax
    margin_after = (profit_after / revenue * 100) if revenue > 0 else 0.0
    return round(tax, 2), round(profit_after, 2), round(margin_after, 2)


def recommended_price(target_margin_pct, cost, fixed_costs, percent_costs_pct):
    denom = 1 - target_margin_pct / 100 - percent_costs_pct / 100
    if denom <= 0:
        return 0.0
    return (cost + fixed_costs) / denom


def round_price(value: float, step: int):
    if value <= 0:
        return 0.0
    return math.ceil(value / step) * step


def read_products(uploaded_file):
    if uploaded_file is None:
        return pd.DataFrame()

    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    rename_map = {}
    for col in df.columns:
        c = normalize_text(col)
        if c in {"sku", "артикул", "артикул товара", "seller sku", "vendorcode", "vendor code"}:
            rename_map[col] = "sku"
        elif c in {"наименование", "название", "наименование товара", "товар", "name", "item name"}:
            rename_map[col] = "name"
        elif c in {"длина", "длина см", "length"}:
            rename_map[col] = "length"
        elif c in {"ширина", "ширина см", "width"}:
            rename_map[col] = "width"
        elif c in {"высота", "высота см", "height"}:
            rename_map[col] = "height"
        elif c in {"вес", "вес кг", "weight"}:
            rename_map[col] = "weight"
        elif c in {"себестоимость", "себес", "cost", "закупка"}:
            rename_map[col] = "cost"
        elif c in {"цена", "текущая цена", "price", "цена продажи"}:
            rename_map[col] = "current_price"

    df = df.rename(columns=rename_map)
    required = ["sku", "name", "length", "width", "height", "weight", "cost"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            "В файле товаров не хватает колонок: " + ", ".join(missing) + 
            ". Скачайте шаблон ниже и заполните его по образцу."
        )

    keep = required + (["current_price"] if "current_price" in df.columns else [])
    return df[keep].copy()


# ---------- UI ----------
st.title("Лемана Про — калькулятор юнит-экономики")
st.caption("Менеджер загружает только файл товаров. Комиссии и логистика берутся из стандартного файла Лемана Про.")

with st.sidebar:
    st.subheader("Параметры")
    scheme = st.selectbox("Схема", ["FBS", "FBO"], index=0)
    tax_regime = st.selectbox("Налог", [
        "ОСНО (25% от прибыли)",
        "УСН Доходы (6%)",
        "УСН Доходы-Расходы (15%)",
        "АУСН (8% от дохода)",
        "УСН НДС 5%",
        "УСН НДС 7%",
    ])
    target_margin = st.slider("Целевая маржа после налога, %", 0, 60, 20)
    acquiring_pct = st.number_input("Эквайринг, %", 0.0, 10.0, 1.5, 0.1)
    marketing_pct = st.number_input("Маркетинг, %", 0.0, 30.0, 0.0, 0.1)
    early_payout_pct = st.number_input("Ранняя выплата, %", 0.0, 10.0, 0.0, 0.1)
    extra_cost_per_unit = st.number_input("Прочие расходы на ед., руб", 0.0, 100000.0, 0.0, 10.0)
    price_round_step = st.selectbox("Округление цены", [1, 10, 50, 100], index=2)

    st.divider()
    st.subheader("Логистика FBS")
    warehouse_zone = st.selectbox("Зона откуда", ["9"], index=0, help="Для склада в Москве используем зону 9.")
    msk_share = st.number_input("Доля Москва и МО, %", 0.0, 100.0, 70.0, 1.0)
    spb_share = st.number_input("Доля СПБ и ЛО, %", 0.0, 100.0, 20.0, 1.0)
    region_share = st.number_input("Доля Регионы, %", 0.0, 100.0, 10.0, 1.0)
    buyout_rate = st.number_input("Выкуп, %", 0.0, 100.0, 95.0, 1.0)
    return_after_buyout = st.number_input("Возвраты после выкупа, %", 0.0, 100.0, 3.0, 1.0)

    st.divider()
    st.subheader("Единицы измерения")
    dim_unit = st.selectbox("Габариты в файле товаров", ["см", "мм"], index=0)
    wt_unit = st.selectbox("Вес в файле товаров", ["кг", "г"], index=0)

    st.divider()
    st.subheader("AI")
    use_ai = st.checkbox("Использовать AI для сложных товаров", value=False)
    openai_key = st.text_input("OpenAI API Key", type="password") if use_ai else ""

ratebook_upload = st.file_uploader(
    "Стандартный файл Лемана Про с комиссиями и логистикой",
    type=["xlsx"],
    help="Если файл уже лежит в репозитории как lemanpro_rates.xlsx, менеджеру ничего сюда загружать не нужно."
)

ratebook = None
source_note = None
try:
    if ratebook_upload is not None:
        ratebook = load_standard_ratebook_from_bytes(ratebook_upload.getvalue())
        source_note = f"Файл загружен вручную: {ratebook_upload.name}"
    else:
        existing = find_existing_rates_file()
        if existing:
            ratebook = load_standard_ratebook_from_path(existing)
            source_note = f"Файл найден в проекте: {existing}"
except Exception as e:
    st.error(f"Не удалось прочитать стандартный файл Лемана Про: {e}")

if source_note:
    st.success(source_note)

st.subheader("Файл товаров")
col_t1, col_t2 = st.columns([2, 1])
with col_t1:
    product_file = st.file_uploader(
        "Загрузите файл товаров",
        type=["xlsx", "xls", "csv"],
        help="Нужны поля: SKU, Наименование, Длина, Ширина, Высота, Вес, Себестоимость. Текущая цена — необязательно."
    )
with col_t2:
    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    st.download_button(
        "Скачать шаблон Excel",
        data=make_template_excel_bytes(),
        file_name="lemanpro_products_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

products = pd.DataFrame()
if product_file is not None:
    try:
        products = read_products(product_file)
        st.write("Предпросмотр товаров")
        st.dataframe(products.head(20), use_container_width=True)
    except Exception as e:
        st.error(str(e))

if st.button("Рассчитать", type="primary", disabled=(ratebook is None or products.empty)):
    commission_df = ratebook["commission"]
    last_mile_df = ratebook["last_mile"]
    return_last_mile_df = ratebook["return_last_mile"]
    to_sc_df = ratebook["to_sc"]
    return_to_sc_df = ratebook["return_to_sc"]

    total_share = msk_share + spb_share + region_share
    if total_share <= 0:
        st.error("Сумма долей зон должна быть больше 0%.")
        st.stop()

    msk_w = msk_share / total_share
    spb_w = spb_share / total_share
    reg_w = region_share / total_share
    buyout = buyout_rate / 100
    cancel_or_not_buyout = 1 - buyout
    post_buyout_return = return_after_buyout / 100

    results = []
    progress = st.progress(0)

    for i, row in products.reset_index(drop=True).iterrows():
        name = str(row["name"])
        sku = str(row["sku"])
        length_cm = normalize_dimension(row["length"], dim_unit)
        width_cm = normalize_dimension(row["width"], dim_unit)
        height_cm = normalize_dimension(row["height"], dim_unit)
        weight_kg = normalize_weight(row["weight"], wt_unit)
        cost = safe_float(row["cost"])
        current_price = safe_float(row.get("current_price", 0.0), 0.0)
        volume_l = max(length_cm * width_cm * height_cm / 1000.0, 0.0)

        matched_row, score = classify_by_rules(name, commission_df)
        if use_ai and openai_key and (matched_row is None or score < 1.4):
            ai_match = classify_with_ai(name, commission_df, openai_key)
            if ai_match is not None:
                matched_row = ai_match
                score = max(score, 9.9)

        template = matched_row["template"] if matched_row is not None else "Не определено"
        item_type = matched_row["type"] if matched_row is not None else ""
        subcat = matched_row["subcategory"] if matched_row is not None else ""
        category = matched_row["category"] if matched_row is not None else ""
        commission_pct = safe_float(matched_row["commission"] * 100 if matched_row is not None else 0.0)

        to_sc = find_break_tariff(volume_l, to_sc_df)
        return_to_sc = find_break_tariff(volume_l, return_to_sc_df)
        lm_msk = find_break_tariff(volume_l, last_mile_df, zone_from=warehouse_zone, zone_to="Москва и МО")
        lm_spb = find_break_tariff(volume_l, last_mile_df, zone_from=warehouse_zone, zone_to="СПБ и ЛО")
        lm_reg = find_break_tariff(volume_l, last_mile_df, zone_from=warehouse_zone, zone_to="Регионы")
        ret_lm_msk = find_break_tariff(volume_l, return_last_mile_df, zone_from=warehouse_zone, zone_to="Москва и МО")
        ret_lm_spb = find_break_tariff(volume_l, return_last_mile_df, zone_from=warehouse_zone, zone_to="СПБ и ЛО")
        ret_lm_reg = find_break_tariff(volume_l, return_last_mile_df, zone_from=warehouse_zone, zone_to="Регионы")

        avg_last_mile = lm_msk * msk_w + lm_spb * spb_w + lm_reg * reg_w
        avg_return_last_mile = ret_lm_msk * msk_w + ret_lm_spb * spb_w + ret_lm_reg * reg_w

        if scheme == "FBS":
            expected_logistics = (
                to_sc + avg_last_mile
                + cancel_or_not_buyout * (avg_return_last_mile + return_to_sc)
                + buyout * post_buyout_return * (avg_return_last_mile + return_to_sc)
            )
        else:
            expected_logistics = avg_last_mile

        percent_costs_pct = commission_pct + acquiring_pct + marketing_pct + early_payout_pct
        rec_price_raw = recommended_price(target_margin, cost, expected_logistics + extra_cost_per_unit, percent_costs_pct)
        rec_price = round_price(rec_price_raw, price_round_step)

        pct_costs_rub = rec_price * percent_costs_pct / 100
        total_costs_before_tax = cost + expected_logistics + extra_cost_per_unit + pct_costs_rub
        tax, profit_after_tax, margin_after_tax = calc_tax(rec_price, total_costs_before_tax, tax_regime)
        profit_before_tax = rec_price - total_costs_before_tax
        margin_before_tax = (profit_before_tax / rec_price * 100) if rec_price > 0 else 0.0
        markup_base = cost + expected_logistics + extra_cost_per_unit
        markup_pct = ((rec_price / markup_base) - 1) * 100 if markup_base > 0 else 0.0

        current_profit_after_tax = None
        current_margin_after_tax = None
        if current_price > 0:
            current_pct_costs = current_price * percent_costs_pct / 100
            current_total_costs = cost + expected_logistics + extra_cost_per_unit + current_pct_costs
            _, current_profit_after_tax, current_margin_after_tax = calc_tax(current_price, current_total_costs, tax_regime)

        results.append({
            "SKU": sku,
            "Наименование": name,
            "Схема": scheme,
            "Шаблон": template,
            "Тип": item_type,
            "Подкатегория": subcat,
            "Категория": category,
            "Совпадение": round(score, 2),
            "Комиссия, %": round(commission_pct, 2),
            "Длина, см": round(length_cm, 2),
            "Ширина, см": round(width_cm, 2),
            "Высота, см": round(height_cm, 2),
            "Вес, кг": round(weight_kg, 3),
            "Объем, л": round(volume_l, 2),
            "Доставка до СЦ, руб": round(to_sc, 2),
            "Возврат до СЦ, руб": round(return_to_sc, 2),
            "Последняя миля Москва, руб": round(lm_msk, 2),
            "Последняя миля СПБ, руб": round(lm_spb, 2),
            "Последняя миля Регионы, руб": round(lm_reg, 2),
            "Средняя последняя миля, руб": round(avg_last_mile, 2),
            "Средняя возвратная миля, руб": round(avg_return_last_mile, 2),
            "Ожидаемая логистика, руб": round(expected_logistics, 2),
            "Себестоимость, руб": round(cost, 2),
            "Прочие расходы, руб": round(extra_cost_per_unit, 2),
            "Текущая цена, руб": round(current_price, 2) if current_price > 0 else None,
            "Текущая прибыль после налога, руб": round(current_profit_after_tax, 2) if current_profit_after_tax is not None else None,
            "Текущая маржа после налога, %": round(current_margin_after_tax, 2) if current_margin_after_tax is not None else None,
            "Рекомендованная цена, руб": round(rec_price, 2),
            "Процентные расходы, %": round(percent_costs_pct, 2),
            "Процентные расходы, руб": round(pct_costs_rub, 2),
            "Прибыль до налога, руб": round(profit_before_tax, 2),
            "Маржа до налога, %": round(margin_before_tax, 2),
            "Налог, руб": round(tax, 2),
            "Прибыль после налога, руб": round(profit_after_tax, 2),
            "Маржа после налога, %": round(margin_after_tax, 2),
            "Наценка, %": round(markup_pct, 2),
            "Флаг": "Проверить" if (score < 1.2 or margin_after_tax < target_margin - 2) else "ОК",
        })
        progress.progress((i + 1) / len(products))

    result_df = pd.DataFrame(results)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("SKU", len(result_df))
    k2.metric("Средняя комиссия, %", f"{result_df['Комиссия, %'].mean():.2f}")
    k3.metric("Средняя логистика, руб", f"{result_df['Ожидаемая логистика, руб'].mean():.0f}")
    k4.metric("Средняя рекомендованная цена, руб", f"{result_df['Рекомендованная цена, руб'].mean():.0f}")

    st.subheader("Результат")
    st.dataframe(result_df, use_container_width=True, height=650)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="Результат", index=False)
        pd.DataFrame({
            "Параметр": [
                "Схема", "Налог", "Целевая маржа, %", "Эквайринг, %", "Маркетинг, %", "Ранняя выплата, %",
                "Прочие расходы, руб", "Склад FBS / Зона откуда", "Москва и МО, %", "СПБ и ЛО, %", "Регионы, %",
                "Выкуп, %", "Возвраты после выкупа, %", "Файл тарифов"
            ],
            "Значение": [
                scheme, tax_regime, target_margin, acquiring_pct, marketing_pct, early_payout_pct,
                extra_cost_per_unit, warehouse_zone, msk_share, spb_share, region_share,
                buyout_rate, return_after_buyout, source_note or ""
            ]
        }).to_excel(writer, sheet_name="Параметры", index=False)
    output.seek(0)

    st.download_button(
        "Скачать результат Excel",
        data=output.getvalue(),
        file_name="lemanpro_unit_economics.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
