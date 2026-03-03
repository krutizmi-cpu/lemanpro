import streamlit as st
import sqlite3
import pandas as pd
from openai import OpenAI

st.set_page_config(
    page_title="Лемана Про — Юнит-экономика FBS",
    layout="wide",
    page_icon="📦"
)

DB_PATH = "products_storage.db"

def init_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sku TEXT UNIQUE,
            name TEXT,
            length_cm REAL,
            width_cm REAL,
            height_cm REAL,
            weight_kg REAL,
            cost REAL DEFAULT 0
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS ai_cache (
            name TEXT,
            client TEXT,
            category TEXT,
            PRIMARY KEY (name, client)
        )
    """)
    conn.commit()
    return conn

def normalize_value(raw, unit):
    try:
        v = float(str(raw).replace(",", ".").strip())
    except (ValueError, TypeError):
        return 0.0
    u = str(unit).strip().lower()
    if u in ("мм", "mm"): return v / 10.0
    if u in ("г", "g", "гр", "gr"): return v / 1000.0
    return v

def get_ai_category(name: str, categories: list, conn, client_key: str) -> str:
    c = conn.cursor()
    row = c.execute(
        "SELECT category FROM ai_cache WHERE name=? AND client=?",
        (name, client_key)
    ).fetchone()
    if row: return row[0]
    api_key = st.session_state.get("openai_key", "")
    if not api_key or not categories: return categories[0] if categories else "Неизвестно"
    try:
        client = OpenAI(api_key=api_key)
        cats_str = "\n".join(f"- {cat}" for cat in categories)
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": f"Ты классификатор товаров для маркетплейса {client_key}. Выбери ОДНУ категорию из списка. Ответь ТОЛЬКО её названием."},
                {"role": "user", "content": f"Товар: {name}\nКатегории:\n{cats_str}"}
            ],
            max_tokens=60,
            temperature=0
        )
        category = resp.choices[0].message.content.strip()
        if category not in categories: category = categories[0]
    except Exception:
        category = categories[0] if categories else "Неизвестно"
    c.execute("INSERT OR REPLACE INTO ai_cache (name, client, category) VALUES (?,?,?)", (name, client_key, category))
    conn.commit()
    return category

def calc_tax(revenue: float, cost_total: float, regime: str):
    profit_before = revenue - cost_total
    rates = {
        "ОСНО (25% от прибыли)": ("profit", 0.25),
        "УСН Доходы (6%)": ("revenue", 0.06),
        "УСН Доходы-Расходы (15%)": ("profit", 0.15),
        "АУСН (8% от дохода)": ("revenue", 0.08),
        "УСН с НДС 5%": ("revenue", 0.05),
        "УСН с НДС 7%": ("revenue", 0.07),
    }
    mode, rate = rates.get(regime, ("profit", 0.0))
    if mode == "revenue": tax = revenue * rate
    else: tax = max(profit_before * rate, 0)
    profit_after = profit_before - tax
    margin_after = (profit_after / revenue * 100) if revenue > 0 else 0
    return round(tax, 2), round(profit_after, 2), round(margin_after, 1)

# --- Лемана Про Logistics ---
CATEGORY_COMMISSIONS = {
    "Блоки, кирпич, бетон (6%)": 6,
    "Арматура / крепёжные элементы (6%)": 6,
    "Сухие смеси / цемент (6%)": 6,
    "Кровельные покрытия (9%)": 9,
    "Теплоизоляция (10%)": 10,
    "Гидроизоляция / пароизоляция (10%)": 10,
    "Сэндвич-панели (10%)": 10,
}

LAST_MILE = {
    "Внутри зоны": {1: 143, 3: 150, 5: 157, 10: 201, 15: 218, 20: 253, 30: 311, 50: 524, 80: 593, 100: 615, 120: 682},
    "СПБ и ЛО": {1: 139, 3: 154, 5: 172, 10: 216, 15: 253, 20: 288, 30: 343, 50: 442, 80: 558, 100: 660, 120: 762},
    "Регион": {1: 142, 3: 177, 5: 206, 10: 239, 15: 276, 20: 306, 30: 368, 50: 552, 80: 888, 100: 1101, 120: 1257}
}

def get_last_mile_tariff(zone, weight_kg):
    table = LAST_MILE.get(zone, LAST_MILE["Регион"])
    thresholds = sorted(table.keys())
    for t in thresholds:
        if weight_kg <= t: return table[t]
    return table[max(thresholds)]

# --- Main App ---
conn = init_db()

if "openai_key" not in st.session_state:
    st.session_state["openai_key"] = st.secrets.get("OPENAI_API_KEY", "")

st.header("Лемана Про — Юнит-экономика (FBS)")

with st.sidebar:
    st.subheader("⚙️ Параметры расчёта")
    tax_regime = st.selectbox("Система налогообложения", [
        "ОСНО (25% от прибыли)", "УСН Доходы (6%)", "УСН Доходы-Расходы (15%)",
        "АУСН (8% от дохода)", "УСН с НДС 5%", "УСН с НДС 7%"
    ])
    st.divider()
    st.subheader("📊 Параметры менеджера")
    target_margin = st.slider("Целевая маржа, %", 0, 50, 20)
    acquiring = st.number_input("Эквайринг, %", 0.0, 5.0, 1.5)
    early_payout = st.number_input("Ранняя выплата, %", 0.0, 5.0, 0.0)
    marketing = st.number_input("Маркетинг, %", 0.0, 20.0, 5.0)
    extra_costs = st.number_input("Доп. расходы на ед., руб", 0, 1000, 0)
    extra_logistics = st.number_input("Доп. логистика, руб", 0, 1000, 0)
    st.divider()
    lp_zone = st.selectbox("Зона доставки", list(LAST_MILE.keys()))

# Catalog Management
with st.expander("Блок 1. Каталог товаров", expanded=True):
    col1, col2 = st.columns(2)
    with col1: dim_unit = st.selectbox("Размеры", ["см", "мм"])
    with col2: wt_unit = st.selectbox("Вес", ["кг", "г"])
    
    uploaded = st.file_uploader("Загрузить Excel (SKU, Название, Длина, Ширина, Высота, Вес, Себестоимость)", type=["xlsx"])
    if uploaded:
        df = pd.read_excel(uploaded)
        if st.button("Сохранить в базу"):
            for _, row in df.iterrows():
                try:
                    sku = str(row.get('SKU', row.get('Артикул', ''))).strip()
                    name = str(row.get('Название', row.get('Наименование', ''))).strip()
                    if not sku or not name: continue
                    l = normalize_value(row.get('Длина', 0), dim_unit)
                    w = normalize_value(row.get('Ширина', 0), dim_unit)
                    h = normalize_value(row.get('Высота', 0), dim_unit)
                    wt = normalize_value(row.get('Вес', 0), wt_unit)
                    cost = float(str(row.get('Себестоимость', 0)).replace(',', '.'))
                    conn.execute("INSERT OR REPLACE INTO products (sku, name, length_cm, width_cm, height_cm, weight_kg, cost) VALUES (?,?,?,?,?,?,?)",
                                 (sku, name, l, w, h, wt, cost))
                except: continue
            conn.commit()
            st.success("Каталог обновлен")

    all_p = pd.read_sql("SELECT * FROM products", conn)
    st.dataframe(all_p)

# Calculation
st.subheader("Блок 2. Расчёт юнит-экономики")
if not all_p.empty and st.button("Рассчитать для всего каталога"):
    results = []
    cat_list = list(CATEGORY_COMMISSIONS.keys())

    for _, p in all_p.iterrows():
        logistics_lp = get_last_mile_tariff(lp_zone, p['weight_kg'])
        logistics_total = logistics_lp + extra_logistics
        cat = get_ai_category(p['name'], cat_list, conn, "lemanpro")
        comm = CATEGORY_COMMISSIONS.get(cat, 0.0)
        
        k_percent = comm + acquiring + early_payout + marketing
        denom = 1 - (k_percent / 100) - (target_margin / 100)
        rrc = (p['cost'] + logistics_total + extra_costs) / denom if denom > 0 else 0
        
        if rrc > 0:
            percent_costs = rrc * (k_percent / 100)
            profit_before = rrc - p['cost'] - logistics_total - extra_costs - percent_costs
            margin_before = (profit_before / rrc * 100) if rrc > 0 else 0
            tax, profit_after, margin_after = calc_tax(rrc, p['cost'] + logistics_total + extra_costs + percent_costs, tax_regime)
        else:
            profit_before = margin_before = tax = profit_after = margin_after = 0.0

        results.append({
            "SKU": p['sku'], "Название": p['name'], "Вес, кг": round(p['weight_kg'], 3),
            "Зона": lp_zone, "Последняя миля, руб": logistics_lp,
            "Категория": cat, "Комиссия, %": comm,
            "РРЦ, руб": round(rrc, 0), "Прибыль, руб": profit_after, "Маржа %": margin_after
        })

    res_df = pd.DataFrame(results)
    st.dataframe(res_df, use_container_width=True)
    st.download_button("Скачать (CSV)", res_df.to_csv(index=False).encode("utf-8"), "lemanpro_results.csv", mime="text/csv")
else:
    st.info("Загрузите каталог товаров для расчёта.")
