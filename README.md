# Лемана Про — repo-версия калькулятора

## Что внутри
- `app.py` — основной Streamlit app
- `data/commissions_fbs_fbo.csv` — комиссии
- `data/zero_mile.csv` — нулевая миля
- `data/return_zero_mile.csv` — возврат нулевой мили
- `data/last_mile.csv` — последняя миля
- `data/return_last_mile.csv` — возврат последней мили
- `data/zones.csv` — карта регионов в зоны
- `update_tariffs_from_excel.py` — скрипт, который обновляет CSV из свежих Excel Лемана Про

## Как запустить
```bash
pip install streamlit pandas openpyxl
streamlit run app.py
```

## Как устроено обновление тарифов
1. Берёте новый Excel по комиссиям от Лемана Про.
2. Кладёте его рядом с `app.py` под именем:
   - `source_commissions.xlsx`
3. Берёте новый Excel по логистике.
4. Кладёте его рядом с `app.py` под именем:
   - `source_logistics.xlsx`
5. Запускаете:
```bash
python update_tariffs_from_excel.py
```
6. Скрипт перезапишет CSV в папке `data/`.
7. После этого просто перезапускаете Streamlit app.

## Какие листы ожидаются
### В `source_commissions.xlsx`
- `Комиссия_FBS и FBO`
- `Тарифы (Последняя миля)`
- `Тарифы (возврат Последняя миля)`
- `Тарифы (Возврат Доставка до СЦ)`
- `Зоны (услуга Последняя миля)`

### В `source_logistics.xlsx`
- `Тарифы (услуга Нулевая миля)`

## Что менять вручную не нужно
Обычно руками править `app.py` не надо. Когда Лемана Про меняет ставки:
- заменяете исходные Excel,
- запускаете `update_tariffs_from_excel.py`,
- коммитите обновлённые CSV в репозиторий.

## Рекомендуемый процесс в Git
```bash
git add data/*.csv source_commissions.xlsx source_logistics.xlsx update_tariffs_from_excel.py app.py
git commit -m "Обновлены тарифы Лемана Про"
```

## Важная оговорка
Если Лемана Про поменяет не только цифры, но и названия листов или колонок, тогда нужно будет немного поправить `update_tariffs_from_excel.py`.
