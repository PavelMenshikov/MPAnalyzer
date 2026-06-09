import streamlit as st
import pandas as pd
import os
import io
import time
import logging
import json
from datetime import datetime
from curl_cffi import requests as curl_requests

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("mp_analyzer")
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="MP Analyzer — Wildberries Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

CSV_PATH = os.path.join(os.path.dirname(__file__), "mpstats_data.csv")
WB_SEARCH_URL = "https://search.wb.ru/exactmatch/ru/common/v4/search"
LIMIT = 100


@st.cache_data
def load_mpstats():
    if not os.path.exists(CSV_PATH):
        return None
    df = pd.read_csv(CSV_PATH, delimiter=";", dtype={"SKU": str})
    df["SKU"] = df["SKU"].str.strip()
    return df


def search_wb(query: str, limit: int = LIMIT, max_retries: int = 3):
    params = {
        "query": query,
        "resultset": "catalog",
        "limit": limit,
        "sort": "popular",
        "dest": "-1257786",
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
        "Referer": "https://www.wildberries.ru/",
    }
    last_resp = None
    for attempt in range(max_retries):
        if attempt > 0:
            wait = 5 * (2 ** attempt)
            logger.warning(f"Повтор {attempt+1}/{max_retries} через {wait}с...")
            time.sleep(wait)
        logger.info(f"Запрос к WB API: query='{query}'")
        last_resp = curl_requests.get(
            WB_SEARCH_URL,
            params=params,
            headers=headers,
            timeout=15,
            impersonate="chrome124",
        )
        logger.info(f"Статус: {last_resp.status_code}")
        if last_resp.status_code == 200:
            try:
                data = last_resp.json()
            except Exception as e:
                logger.error(f"JSON parse error: {e}")
                continue
            products = data.get("products", [])
            logger.info(f"Найдено товаров: {len(products)}")
            if products:
                logger.info(f"Первый: SKU={products[0].get('id')}, {products[0].get('name', '')[:50]}")
            return products
        elif last_resp.status_code == 429:
            logger.warning("429 Too Many Requests")
            continue
        else:
            logger.warning(f"Неожиданный статус: {last_resp.status_code}, тело: {last_resp.text[:200]}")
            continue
    logger.error("Все попытки исчерпаны")
    return []


def parse_search_products(products: list):
    rows = []
    for p in products:
        sku = str(p.get("id", ""))
        price_kop = p.get("salePriceU") or p.get("priceU")

        if not price_kop:
            top_price = p.get("price", {}) or {}
            price_kop = top_price.get("total") or top_price.get("basic")
        if not price_kop:
            sizes = p.get("sizes") or []
            if sizes:
                price_kop = sizes[0].get("price", {}).get("product") or sizes[0].get("price", {}).get("basic")

        price = round(price_kop / 100, 2) if price_kop else None

        rows.append(
            {
                "SKU": sku,
                "Название": p.get("name", ""),
                "Бренд": p.get("brand", ""),
                "Цена, ₽": price,
                "Рейтинг": p.get("rating"),
                "Отзывы": p.get("feedbacks", 0),
                "Поставщик": p.get("supplier", ""),
                "Ссылка": f"https://www.wildberries.ru/catalog/{sku}/detail.aspx",
            }
        )
    return pd.DataFrame(rows)


def merge_with_mpstats(wb_df: pd.DataFrame, mpstats_df: pd.DataFrame):
    merged = wb_df.merge(mpstats_df, on="SKU", how="left", suffixes=("", "_mp"))
    dup_cols = [c for c in merged.columns if c.endswith("_mp")]
    merged.drop(columns=dup_cols, inplace=True, errors="ignore")
    return merged


def color_row_conditions(df: pd.DataFrame):
    colors = {}
    revenue_col = "Revenue"
    lost_col = "Lost profit"
    days_col = "Days in website"

    revenue_series = pd.to_numeric(
        df[revenue_col].astype(str).str.replace(",", ".", regex=False).str.split("|").str[0],
        errors="coerce",
    )
    lost_series = (
        pd.to_numeric(
            df[lost_col].astype(str).str.replace(",", ".", regex=False).str.split("|").str[0],
            errors="coerce",
        )
        if lost_col in df.columns
        else pd.Series([None] * len(df))
    )
    days_series = (
        pd.to_numeric(df[days_col], errors="coerce") if days_col in df.columns else pd.Series([None] * len(df))
    )

    top5_revenue = (
        revenue_series.nlargest(5).iloc[-1]
        if revenue_series.notna().sum() >= 5
        else revenue_series.max()
    )

    for idx in df.index:
        rev = revenue_series.get(idx, None)
        lost = lost_series.get(idx, None)
        days = days_series.get(idx, None)

        reasons = []

        if pd.notna(rev) and rev >= top5_revenue:
            reasons.append("green")
        if pd.notna(rev) and pd.notna(lost) and lost > rev * 2:
            reasons.append("red")
        if pd.notna(days) and days < 60 and idx < 20:
            reasons.append("yellow")

        if reasons:
            colors[idx] = reasons

    return colors


def build_excel(
    wb_df: pd.DataFrame,
    dashboard_df: pd.DataFrame,
    query: str,
):
    xls = openpyxl.Workbook()

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)

    colors = color_row_conditions(dashboard_df)

    # Sheet 1: Выдача WB
    ws1 = xls.active
    ws1.title = "Выдача WB"

    for col_idx, col_name in enumerate(wb_df.columns, 1):
        cell = ws1.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for row_idx, (_, row) in enumerate(wb_df.iterrows(), 2):
        for col_idx, col_name in enumerate(wb_df.columns, 1):
            val = row[col_name]
            if pd.isna(val) or val is None:
                val = ""
            cell = ws1.cell(row=row_idx, column=col_idx, value=val)
            cell.alignment = Alignment(horizontal="left")

    for col_idx in range(1, len(wb_df.columns) + 1):
        ws1.column_dimensions[get_column_letter(col_idx)].width = 22

    # Sheet 2: Dashboard
    ws2 = xls.create_sheet(title="Dashboard")
    for col_idx, col_name in enumerate(dashboard_df.columns, 1):
        cell = ws2.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for row_idx, (_, row) in enumerate(dashboard_df.iterrows(), 2):
        for col_idx, col_name in enumerate(dashboard_df.columns, 1):
            val = row[col_name]
            if pd.isna(val) or val is None:
                val = ""
            cell = ws2.cell(row=row_idx, column=col_idx, value=val)
            cell.alignment = Alignment(horizontal="left")

        row_colors = colors.get(row_idx - 2, [])
        if row_colors:
            if "green" in row_colors and "yellow" in row_colors:
                fill = green_fill
            elif "red" in row_colors:
                fill = red_fill
            elif "green" in row_colors:
                fill = green_fill
            elif "yellow" in row_colors:
                fill = yellow_fill
            else:
                fill = None
            if fill:
                for col_idx in range(1, len(dashboard_df.columns) + 1):
                    ws2.cell(row=row_idx, column=col_idx).fill = fill

    for col_idx in range(1, len(dashboard_df.columns) + 1):
        ws2.column_dimensions[get_column_letter(col_idx)].width = 22

    buf = io.BytesIO()
    xls.save(buf)
    buf.seek(0)
    return buf


def main():
    st.sidebar.title("MP Analyzer")
    st.sidebar.caption("Wildberries Dashboard Tool")

    st.sidebar.divider()
    uploaded_file = st.sidebar.file_uploader(
        "📂 Загрузить MPStats CSV",
        type=["csv"],
        help="Формат: CSV с разделителем ';'. Если не загружен — используется демо-файл (Проекторы, 307 товаров)",
    )
    if uploaded_file is not None:
        try:
            mpstats_df = pd.read_csv(uploaded_file, delimiter=";", dtype={"SKU": str})
            mpstats_df["SKU"] = mpstats_df["SKU"].str.strip()
            st.sidebar.success(f"✅ Загружено: {len(mpstats_df)} товаров")
        except Exception as e:
            st.sidebar.error(f"Ошибка загрузки CSV: {e}")
            mpstats_df = load_mpstats()
    else:
        mpstats_df = load_mpstats()

    menu = st.sidebar.radio(
        "Навигация",
        [
            "📋 Часть 1 — Анализ ТЗ",
            "✏️ Часть 2 — Идеальное ТЗ",
            "🚀 Часть 3 — Реализация",
            "ℹ️ О проекте",
        ],
    )

    if menu == "📋 Часть 1 — Анализ ТЗ":
        show_part1()
    elif menu == "✏️ Часть 2 — Идеальное ТЗ":
        show_part2()
    elif menu == "🚀 Часть 3 — Реализация":
        show_part3(mpstats_df)
    elif menu == "ℹ️ О проекте":
        show_about(mpstats_df)


def show_part1():
    st.title("📋 Часть 1. Анализ технического задания")
    st.markdown(
        """
### Исходное ТЗ (от руководителя):

> «Напиши скрипт для Google Таблицы. Пользователь вводит поисковый запрос, нажимает кнопку, скрипт лезет в WB, вытаскивает **все** товары по запросу, берет **все данные** которые есть, соединяет данные с выгрузкой из MPStats, **форматирует как надо** и записывает в итоговую таблицу. Если что-то непонятно — **придумай сам**.»

---

### Найденные проблемы (8 шт.)
"""
    )

    problems = [
        (
            "1️⃣ Не указан эндпоинт WB API",
            "**Что не определено:** Не сказано, какой именно API Wildberries использовать — поисковый, карточки товара, каталог?\n\n"
            "**Что произойдёт:** Claude может выбрать неподходящий эндпоинт (например, устаревший / непубличный) или начать парсить HTML, что приведёт к неработающему коду.\n\n"
            "**Как сформулировать:** *«Используй эндпоинт search.wb.ru/exactmatch/ru/common/v4/search для получения поисковой выдачи и card.wb.ru/cards/detail для получения детальных данных по каждому артикулу.»*",
        ),
        (
            "2️⃣ Не определён лимит товаров («все товары»)",
            "**Что не определено:** Сколько товаров вытаскивать? По запросу может быть 10 000+ результатов.\n\n"
            "**Что произойдёт:** Claude либо будет делать бесконечный цикл пагинации (упадёт по таймауту), либо вытащит 10+ страниц, создав 1000+ запросов.\n\n"
            "**Как сформулировать:** *«Вытащи топ-100 артикулов из поисковой выдачи Wildberries.»*",
        ),
        (
            "3️⃣ Не указано, откуда брать поисковый запрос",
            "**Что не определено:** Где пользователь вводит запрос? В ячейке A1? Во всплывающем окне? В отдельном поле?\n\n"
            "**Что произойдёт:** Claude сделает на своё усмотрение, и интерфейс может не совпасть с ожиданиями.\n\n"
            "**Как сформулировать:** *«Пользователь вводит поисковый запрос в ячейку A1 на листе \"Выдача WB\". Скрипт читает значение из A1.»*",
        ),
        (
            "4️⃣ Не определены конкретные поля для выгрузки",
            "**Что не определено:** «Все данные которые есть» — это какие именно поля? WB API возвращает десятки полей (цена, бренд, рейтинг, размеры, цвета, характеристики...).\n\n"
            "**Что произойдёт:** Claude вытащит произвольный набор полей, часть из них будет неинформативной, часть — избыточной.\n\n"
            "**Как сформулировать:** *«Из WB API вытащи следующие поля: SKU (id), Название (name), Бренд (brand), Цена (salePriceU / 100), Рейтинг (rating), Количество отзывов (feedbacks), Поставщик (supplier), Ссылка на товар.»*",
        ),
        (
            "5️⃣ Не описана логика объединения с MPStats",
            "**Что не определено:** По какому полю объединять? Что делать, если SKU из WB нет в MPStats или наоборот? Какие поля из MPStats брать?\n\n"
            "**Что произойдёт:** Claude придумает логику наугад — может сделать INNER JOIN (потеряв товары), LEFT JOIN (создав пустые строки) или просто склеить данные в хаотичном порядке.\n\n"
            "**Как сформулировать:** *«Объединение по полю SKU (артикул). Если артикул есть в выдаче WB, но отсутствует в MPStats — поля из MPStats оставить пустыми. Если артикул есть в MPStats, но его нет в выдаче — не включать в итоговый дашборд.»*",
        ),
        (
            "6️⃣ «Форматирует как надо» — полная неопределённость",
            "**Что не определено:** Какие строки каким цветом красить? По какому условию? Какой шрифт? Какие границы?\n\n"
            "**Что произойдёт:** Claude выберет произвольные цвета и условия, которые могут не соответствовать бизнес-логике.\n\n"
            "**Как сформулировать:** *«Условное форматирование: (1) Топ-5 по выручке — зелёный фон строки. (2) Упущенная прибыль превышает выручку более чем в 2 раза — красный фон строки. (3) Товар на сайте менее 60 дней и входит в топ-20 поисковой выдачи — жёлтый фон строки.»*",
        ),
        (
            "7️⃣ «Придумай сам» — риск несовместимости с ожиданиями",
            "**Что не определено:** Все пробелы, которые Claude должен «додумать».\n\n"
            "**Что произойдёт:** Claude примет N решений наугад. Чем больше неопределённостей, тем выше шанс, что результат не совпадёт с ожиданиями. Руководитель получит не то, что хотел, и потратит время на переделки.\n\n"
            "**Как сформулировать:** *Устранить все неопределённости: явно описать каждый шаг, каждый формат, каждое исключение. Claude не должен ни о чём «догадываться» самостоятельно.*",
        ),
        (
            "8️⃣ Не указана обработка ошибок",
            "**Что не определено:** Что делать, если API недоступен? Если введён пустой запрос? Если SKU не найден? Если некорректные данные?\n\n"
            "**Что произойдёт:** При первой же ошибке скрипт упадёт, а пользователь увидит непонятную ошибку вместо результата.\n\n"
            "**Как сформулировать:** *«Если артикул не найден через WB API — в строке указать \"не найдено\", скрипт не прерывать. Если API недоступен — показать понятное сообщение пользователю. Если цены нет — оставить поле пустым. Пустые поля в ответе API (бренд, характеристики) не вызывают ошибку.»*",
        ),
    ]

    for title, body in problems:
        with st.expander(title, expanded=True):
            st.markdown(body)

    st.divider()
    st.success(
        "**Итого:** 8 конкретных проблем, каждая из которых приводит к неработающему "
        "или непредсказуемому результату. Корректное ТЗ должно исключать все "
        "неопределённости."
    )


def show_part2():
    st.title("✏️ Часть 2. Корректное техническое задание")

    st.markdown(
        """
### Цель

Разработать скрипт, который по поисковому запросу пользователя получает топ-100 товаров из Wildberries, 
обогащает их данными из CSV-выгрузки MPStats и формирует итоговый дашборд с условным форматированием.

---

### 1. Эндпоинты WB API

| Назначение | Эндпоинт | Параметры |
|---|---|---|
| Поисковая выдача | `https://search.wb.ru/exactmatch/ru/common/v4/search` | `query={запрос}`, `resultset=catalog`, `limit=100`, `sort=popular` |
| Данные по артикулам | `https://card.wb.ru/cards/detail` | `nm={список SKU через запятую}` |

### 2. Входные данные

- **Поисковый запрос** — берётся из ячейки **A1** листа «Выдача WB».
- **Данные MPStats** — CSV-файл с колонками: `Thumb`, `Name`, `Purchase After Return`, `SKU`, `URL`, `Brand`, `Country`, `Revenue`, `Lost profit`, `Sales`, `Price with WB wallet`, `Comments Valuation`, `Comments`, `Days in website`, `First Date`.

### 3. Поля из WB API (для листа «Выдача WB»)

| Поле в таблице | Поле в API | Примечание |
|---|---|---|---|
| SKU | `id` | Артикул |
| Название | `name` | |
| Бренд | `brand` | |
| Цена, ₽ | `salePriceU / 100` или `sizes[0].price.product / 100` | Копейки → рубли |
| Рейтинг | `rating` | |
| Отзывы | `feedbacks` | |
| Поставщик | `supplier` | |
| Ссылка | `https://www.wildberries.ru/catalog/{SKU}/detail.aspx` | Формируется из SKU |

### 4. Логика объединения (Dashboard)

1. Взять топ-100 SKU из поисковой выдачи.
2. Для каждого SKU получить детальные данные через card API.
3. **Объединение с MPStats по полю SKU**:
   - LEFT JOIN: все поля из «Выдача WB» + поля из MPStats.
   - Если SKU есть в выдаче, но нет в CSV — поля MPStats = пусто.
   - Если SKU есть в CSV, но нет в выдаче — строка **не включается** в дашборд.
4. Отсутствующие значения оставлять пустыми (не заполнять 'N/A', '0' и т.д.).

### 5. Условное форматирование (Dashboard)

| Условие | Цвет строки | Пояснение |
|---|---|---|
| Топ-5 по выручке (Revenue) | 🟢 Зелёный | Определяется динамически, не привязано к строкам |
| Упущенная прибыль (Lost profit) > Выручка × 2 | 🔴 Красный | Товар-убыток |
| Товар на сайте < 60 дней (Days in website) И входит в топ-20 поисковой выдачи (по Sales) | 🟡 Жёлтый | Перспективный новинка |

### 6. Обработка ошибок (edge cases)

| Ситуация | Действие |
|---|---|
| Артикул не найден в WB API | В строке указать «не найдено», продолжить со следующим |
| API недоступен / таймаут | Повторить запрос 1 раз; если не помогло — пропустить, вывести предупреждение |
| Пустой поисковый запрос | Показать сообщение «Введите поисковый запрос» |
| Цена отсутствует | Оставить поле пустым |
| Пустые поля в ответе API (бренд, характеристики) | Оставить пустыми, ошибку не вызывать |
| Некорректный CSV (нет SKU, не те колонки) | Показать понятное сообщение об ошибке |

### 7. Запрещённые действия

- ❌ Хардкодить данные вместо получения из API
- ❌ Привязывать условное форматирование к конкретным номерам строк
- ❌ Дублировать строки при повторном запуске — таблица полностью очищается перед заполнением
- ❌ Использовать непроверенные/недокументированные API с нестабильным форматом ответа
- ❌ Прерывать выполнение при единичной ошибке по одному товару

### 8. Критерий готовности

- [ ] Скрипт запускается на реальных данных
- [ ] Топ-100 выдачи WB корректно отображаются в листе «Выдача WB»
- [ ] Dashboard содержит объединённые данные с корректным LEFT JOIN
- [ ] Условное форматирование работает динамически (не привязано к строкам)
- [ ] Кнопка запуска находится в меню таблицы
- [ ] Повторный запуск не дублирует данные
- [ ] Все edge cases обработаны без падения скрипта
"""
    )


def show_part3(mpstats_df: pd.DataFrame | None):
    st.title("🚀 Часть 3. Реализация")
    st.markdown(
        "**Инструмент:** Python + Streamlit + openpyxl — выбран из-за гибкости, "
        "возможности сделать веб-интерфейс и одновременной генерации Excel-файла "
        "с полным форматированием. Streamlit позволяет быстро создавать интерактивные "
        "дашборды, а openpyxl даёт полный контроль над Excel-выводом.\n\n"
        "**API Wildberries:** используется эндпоинт поисковой выдачи "
        "`search.wb.ru/exactmatch/ru/common/v4/search` через библиотеку "
        "`curl_cffi` (имитация TLS Chrome для обхода WAF). Результаты "
        "кешируются, чтобы не дёргать API повторно.\n\n"
        "**Цена** определяется по цепочке: `salePriceU` → `priceU` → "
        "`price.total` → `sizes[0].price.product` (копейки → рубли).\n\n"
        "**Ссылка на товар** формируется из SKU: "
        "`https://www.wildberries.ru/catalog/{SKU}/detail.aspx`"
    )

    tab1, tab2, tab3 = st.tabs(["🔍 Выдача WB", "📊 Dashboard", "📥 Экспорт"])

    query = st.sidebar.text_input(
        "🔎 Поисковый запрос (введите и нажмите Enter)",
        placeholder="например, проектор для фильмов",
    )

    if "wb_data" not in st.session_state:
        st.session_state.wb_data = None
        st.session_state.dashboard_data = None
        st.session_state.last_query = ""
    if "wb_cache" not in st.session_state:
        st.session_state.wb_cache = {}

    with tab1:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.subheader("Поисковая выдача Wildberries")
        with col2:
            search_btn = st.button("🚀 Выполнить поиск", type="primary", use_container_width=True)

        if mpstats_df is None:
            st.warning("⚠️ Файл MPStats не найден. Загрузка данных MPStats недоступна.")

        if search_btn:
            if not query.strip():
                st.error("❌ Введите поисковый запрос")
            else:
                cache_key = query.strip().lower()
                cached = st.session_state.wb_cache.get(cache_key)

                if cached is not None:
                    search_results = cached
                    st.info("📦 Результаты из кеша (повторный запрос к WB API не выполнялся)")
                else:
                    with st.spinner("Поиск товаров на Wildberries..."):
                        try:
                            search_results = search_wb(cache_key, LIMIT)
                            st.session_state.wb_cache[cache_key] = search_results
                        except Exception as e:
                            st.error(f"❌ Ошибка API Wildberries: {e}")
                            search_results = []

                if not search_results:
                    st.warning("Ничего не найдено по запросу. Попробуйте другой запрос.")
                else:
                    wb_df = parse_search_products(search_results)

                    st.session_state.wb_data = wb_df
                    st.session_state.last_query = query

                    if mpstats_df is not None:
                        dashboard_df = merge_with_mpstats(wb_df, mpstats_df)
                    else:
                        dashboard_df = wb_df.copy()
                    st.session_state.dashboard_data = dashboard_df

                    st.success(f"✅ Найдено {len(wb_df)} товаров")

        if st.session_state.wb_data is not None:
            df = st.session_state.wb_data
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.caption(f"Всего товаров: {len(df)}")
        else:
            st.info("💡 Введите поисковый запрос в боковой панели и нажмите «Выполнить поиск»")

    with tab2:
        st.subheader("Dashboard — объединение с MPStats")

        if st.session_state.dashboard_data is not None:
            df = st.session_state.dashboard_data
            colors = color_row_conditions(df)

            has_csv_data = "Revenue" in df.columns and df["Revenue"].notna().any()

            if not has_csv_data and len(df) > 0:
                st.warning(
                    "⚠️ Нет данных из MPStats для этого поискового запроса. "
                    "Условное форматирование (цвета) не применится.\n\n"
                    "Загрузите свой CSV-файл MPStats через боковую панель "
                    "(`Загрузить MPStats CSV`) для получения аналитики по "
                    "выручке, упущенной прибыли и дням на сайте."
                )

            df_display = df.copy()
            sort_order = []
            for idx in df_display.index:
                c = colors.get(idx, [])
                if "red" in c:
                    sort_order.append(0)
                elif "green" in c:
                    sort_order.append(1)
                elif "yellow" in c:
                    sort_order.append(2)
                else:
                    sort_order.append(3)
            df_display["_sort"] = sort_order
            df_display = df_display.sort_values("_sort").drop(columns=["_sort"])

            def highlight_row(row):
                idx = row.name
                if idx in colors:
                    c = colors[idx]
                    if "red" in c:
                        return ["background-color: #FFC7CE"] * len(row)
                    elif "yellow" in c:
                        return ["background-color: #FFEB9C"] * len(row)
                    elif "green" in c:
                        return ["background-color: #C6EFCE"] * len(row)
                return [""] * len(row)

            styled = df_display.style.apply(highlight_row, axis=1)
            st.dataframe(styled, use_container_width=True, hide_index=True)

            green_count = sum(1 for v in colors.values() if "green" in v)
            red_count = sum(1 for v in colors.values() if "red" in v)
            yellow_count = sum(1 for v in colors.values() if "yellow" in v)

            col_leg = st.columns(5)
            with col_leg[0]:
                st.markdown(f"🟢 **Зелёный** — топ-5 по выручке ({green_count})")
            with col_leg[1]:
                st.markdown(f"🔴 **Красный** — упущ. прибыль > выручка × 2 ({red_count})")
            with col_leg[2]:
                st.markdown(f"🟡 **Жёлтый** — новинка < 60 дн. в топ-20 ({yellow_count})")
            with col_leg[3]:
                st.markdown(f"📊 **Всего:** {len(df)}")
            with col_leg[4]:
                st.markdown(f"📂 **Совпадений с CSV:** {df['Revenue'].notna().sum() if has_csv_data else 0}")
        else:
            st.info("💡 Сначала выполните поиск на вкладке «Выдача WB»")

    with tab3:
        st.subheader("📥 Экспорт в Excel")

        if st.session_state.dashboard_data is not None:
            st.markdown(
                "Будет создан Excel-файл с двумя листами:\n"
                "- **Выдача WB** — результаты поиска\n"
                "- **Dashboard** — объединённые данные с условным форматированием"
            )

            export_btn = st.download_button(
                label="📥 Скачать Excel",
                data=build_excel(
                    st.session_state.wb_data,
                    st.session_state.dashboard_data,
                    st.session_state.last_query,
                ),
                file_name=f"wb_dashboard_{st.session_state.last_query[:20].replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        else:
            st.info("💡 Сначала выполните поиск на вкладке «Выдача WB»")


def show_about(mpstats_df: pd.DataFrame | None):
    st.title("ℹ️ О проекте")
    st.markdown(
        """
### Как это работает

1. **Загрузите** свой CSV-файл MPStats (через боковую панель)
2. **Введите** поисковый запрос (например, «проектор»)
3. **Получите** топ-100 товаров из Wildberries, обогащённых аналитикой MPStats
4. **Экспортируйте** в Excel с цветовой маркировкой

Если CSV не загружен — используется демо-файл (Проекторы, 307 товаров). Загрузите свой
CSV-экспорт из MPStats для любой ниши.

---

### Гибридный подход: Streamlit + Excel

**Почему этот подход — лучший для тестового задания и реальной работы:**

| Критерий | Обычный Google Apps Script | Streamlit + Excel |
|---|---|---|
| Интерфейс | Только таблица | Интерактивный дашборд |
| Удобство разработки | GAS-отладка — боль | Python + IDE |
| Формат сдачи | Ссылка на таблицу | Ссылка на приложение + Excel |
| Возможность доработки | Ограниченная | Полная (Python экосистема) |
| Визуализация | Нет | Встроенная в Streamlit |

### Данные MPStats
"""
    )
    if mpstats_df is not None:
        st.dataframe(mpstats_df.head(10), use_container_width=True, hide_index=True)
        st.caption(f"Всего записей в CSV: {len(mpstats_df)}")
        col_names = ", ".join(mpstats_df.columns.tolist())
        st.caption(f"Колонки: {col_names}")

    st.markdown(
        """
### Использованные технологии

- **Python 3.10+** — основной язык
- **Streamlit** — веб-интерфейс
- **pandas** — обработка данных
- **openpyxl** — генерация Excel с форматированием
- **curl_cffi** — HTTP-запросы к API Wildberries (обход WAF через имитацию TLS Chrome)
- **Wildberries API** — публичный эндпоинт поисковой выдачи (`search.wb.ru/exactmatch/ru/common/v4/search`)

### Запуск приложения

```bash
cd mp_analyzer
pip install -r requirements.txt
streamlit run app.py
```
"""
    )


if __name__ == "__main__":
    main()
