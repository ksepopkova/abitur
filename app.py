st.write("DEBUG URL params:", dict(st.query_params))
import streamlit as st
import pandas as pd
import re
import io
import uuid
import json
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from yookassa import Configuration, Payment

st.set_page_config(page_title="Подбор вузов", layout="wide")

OBL = {
    "Русский язык": 0, "Математика": 1, "Обществознание": 2,
    "История": 3, "Иностранный язык": 4, "Биология": 5,
    "Химия": 6, "Физика": 7, "Информатика": 8,
    "География": 9, "Литература": 10, "ДВИ": 11,
}
VYB = {
    "Математика": 12, "Обществознание": 13, "История": 14,
    "Иностранный язык": 15, "Биология": 16, "Химия": 17,
    "Физика": 18, "Информатика": 19, "География": 20, "Литература": 21,
}

@st.cache_data
def load_data():
    sheets = {}
    for sheet in ["Москва", "Питер", "Регионы"]:
        df = pd.read_excel("База вузов 2026.xlsx", sheet_name=sheet, header=1)
        df["_лист"] = sheet
        sheets[sheet] = df
    full = pd.concat(sheets.values(), ignore_index=True)
    full.iloc[:, 25] = full.iloc[:, 25].astype(str).str.strip()
    return full

def get_city_group(city_val):
    s = str(city_val).strip()
    if any(x in s for x in ["Москва", "МО", "Московская"]):
        return "Москва и Московская область"
    if any(x in s for x in ["Петербург", "Ленинград", "Пушкин", "Гатчина"]):
        return "Санкт-Петербург и Ленинградская область"
    return s

@st.cache_data
def get_city_options(df):
    cities = df.iloc[:, 22].dropna().unique()
    groups = sorted(set(get_city_group(c) for c in cities))
    priority = ["Москва и Московская область", "Санкт-Петербург и Ленинградская область"]
    return priority + [g for g in groups if g not in priority]

@st.cache_data
def get_vuz_by_city(df, city_group):
    mask = df.iloc[:, 22].apply(lambda x: get_city_group(x) == city_group)
    vuzы = df[mask].iloc[:, 23].dropna().unique()
    return sorted(set(str(v).strip() for v in vuzы if str(v).strip() not in ("", "nan")))

@st.cache_data
def get_all_codes(df):
    codes = df.iloc[:, 25].dropna().unique()
    return sorted(set(str(c).strip() for c in codes if str(c).strip() not in ("", "nan")))

def clean_str(val):
    if pd.isna(val): return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s

def to_num(val):
    s = str(val).strip()
    if s.upper() == "NEW": return "NEW"
    if s == "-": return "-"
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except:
        return None

def cell_has_value(row, col_idx):
    val = row.iloc[col_idx]
    if pd.isna(val): return False
    s = str(val).strip()
    if s in ('', 'nan'): return False
    try: return float(s) > 0
    except: return True

def cell_value(row, col_idx):
    try: return float(row.iloc[col_idx])
    except: return 0

def subject_passes(row, col_idx, student_score):
    if not cell_has_value(row, col_idx): return False
    threshold = cell_value(row, col_idx)
    if threshold == 0: return True
    return student_score >= threshold

def has_budget_places(val):
    s = str(val).strip()
    if s in ('-', 'nan', ''): return False
    m = re.match(r'(\d+)\s*\((\d+)\)', s)
    if m and int(m.group(2)) == 0: return False
    try: return float(s) > 0
    except: return True

def calc_achievements(row, gto, attestat):
    total = 0
    if gto == "Золото": total += cell_value(row, 32)
    elif gto == "Серебро": total += cell_value(row, 33)
    elif gto == "Бронза": total += cell_value(row, 34)
    if attestat: total += cell_value(row, 35)
    return min(total, 10)

def calc_student_score(row, subjects):
    score = subjects.get("Русский язык", 0)
    obl_set = set()
    for subj, col in OBL.items():
        if subj in ("Русский язык", "ДВИ"): continue
        if cell_has_value(row, col):
            score += subjects.get(subj, 0)
            obl_set.add(subj)
    best_vyb = 0
    for subj, col in VYB.items():
        if subj in obl_set: continue
        if cell_has_value(row, col):
            s = subjects.get(subj, 0)
            if s > best_vyb: best_vyb = s
    score += best_vyb
    return score

def check_row(row, subjects):
    rus = subjects.get("Русский язык", 0)
    if not cell_has_value(row, OBL["Русский язык"]): return None
    if rus < cell_value(row, OBL["Русский язык"]): return None
    obl_required = [(s, c) for s, c in OBL.items()
                    if s not in ("Русский язык", "ДВИ") and cell_has_value(row, c)]
    for subj, col in obl_required:
        if not subject_passes(row, col, subjects.get(subj, 0)): return None
    dvi_required = cell_has_value(row, OBL["ДВИ"])
    obl_set = {s for s, _ in obl_required}
    vyb_required = [(s, c) for s, c in VYB.items()
                    if s not in obl_set and cell_has_value(row, c)]
    if vyb_required:
        if not any(subjects.get(s, 0) > 0 and subject_passes(row, c, subjects.get(s, 0))
                   for s, c in vyb_required): return None
    return "with_dvi" if dvi_required else "no_dvi"

def get_chance(student_score, pb, sb):
    try:
        pb_f, sb_f = float(pb), float(sb)
        if pb_f <= 1 and sb_f <= 1: return "new"
        if student_score >= sb_f + 10: return "podstrahovka"
        if student_score >= sb_f: return "realistic"
        if student_score > pb_f + 5: return "probable"
        if student_score >= pb_f - 15: return "risky"
        return "unlikely"
    except:
        return "new"

CHANCE_ORDER = {
    "probable": 0, "realistic": 1, "podstrahovka": 2,
    "risky": 3, "unlikely": 4, "new": 5, "no_dvi_score": 6,
}
CHANCE_LABEL = {
    "podstrahovka": "🟢 Уверенно",
    "realistic":    "🔵 Реалистично",
    "probable":     "🟡 Вероятно",
    "risky":        "🔴 Рискованно",
    "unlikely":     "⚫ Маловероятно",
    "new":          "⬜ Нет данных",
    "no_dvi_score": "⬜ Нет оценки — не указан балл за ДВИ",
}
PRIORITY_LABEL = {
    "podstrahovka": "3–5",
    "realistic":    "2–3",
    "probable":     "1–2",
    "risky":        "1*",
    "unlikely":     "—",
    "new":          "1*",
    "no_dvi_score": "—",
}

def build_result_row(row, subjects, gto, attestat, dvi_score=None):
    pb, sb = row.iloc[28], row.iloc[29]
    student_score = calc_student_score(row, subjects)
    achievements = calc_achievements(row, gto, attestat)
    total_score = student_score + achievements
    dvi_required = cell_has_value(row, OBL["ДВИ"])
    if dvi_required:
        if dvi_score and dvi_score > 0:
            total_score += dvi_score
            chance = get_chance(total_score, pb, sb)
        else:
            chance = "no_dvi_score"
    else:
        chance = get_chance(total_score, pb, sb)
    return {
        "Город":               clean_str(row.iloc[22]),
        "Вуз":                 clean_str(row.iloc[23]),
        "Факультет":           clean_str(row.iloc[24]),
        "Код и специальность": clean_str(row.iloc[25]),
        "Профиль":             clean_str(row.iloc[26]),
        "Мест":                to_num(row.iloc[27]),
        "Проходной балл":      to_num(pb),
        "Средний балл":        to_num(sb),
        "Ваш балл (ЕГЭ)":     student_score,
        "Достижения":          achievements if achievements > 0 else "",
        "Конкурсный балл":     total_score,
        "Шансы":               chance,
        "Рек. приоритет":      PRIORITY_LABEL[chance],
        "ГТО золото":          to_num(row.iloc[32]),
        "ГТО серебро":         to_num(row.iloc[33]),
        "ГТО бронза":          to_num(row.iloc[34]),
        "Аттестат":            to_num(row.iloc[35]),
    }

def filter_rows_flow2(df, subjects, show_dvi, selected_city_groups, gto, attestat, dvi_score=None):
    results = []
    for _, row in df.iterrows():
        city_raw = str(row.iloc[22]).strip()
        if get_city_group(city_raw) not in selected_city_groups: continue
        if not has_budget_places(row.iloc[27]): continue
        status = check_row(row, subjects)
        if status is None: continue
        if status == "with_dvi" and not show_dvi: continue
        results.append(build_result_row(row, subjects, gto, attestat, dvi_score))
    return pd.DataFrame(results)

def filter_rows_flow1(df, subjects, selected_vuz, selected_codes, gto, attestat, selected_cities=None, dvi_score=None):
    expanded_codes = set(selected_codes)
    for code in selected_codes:
        parts = code.split(' ')[0]
        if len(parts) >= 7:
            prefix = parts[:5]
            suffix = parts[5:]
            if suffix == '.00':
                for _, row in df.iterrows():
                    if clean_str(row.iloc[23]) not in selected_vuz: continue
                    row_code = clean_str(row.iloc[25])
                    if row_code.split(' ')[0].startswith(prefix):
                        expanded_codes.add(row_code)
            else:
                for _, row in df.iterrows():
                    if clean_str(row.iloc[23]) not in selected_vuz: continue
                    row_code = clean_str(row.iloc[25])
                    if row_code.split(' ')[0] == prefix + '.00':
                        expanded_codes.add(row_code)
    results = []
    for _, row in df.iterrows():
        city_raw = str(row.iloc[22]).strip()
        city_group = get_city_group(city_raw)
        vuz = clean_str(row.iloc[23])
        code = clean_str(row.iloc[25])
        if selected_cities and city_group not in selected_cities: continue
        if vuz not in selected_vuz: continue
        if code not in expanded_codes: continue
        if not has_budget_places(row.iloc[27]): continue
        status = check_row(row, subjects)
        if status is None: continue
        results.append(build_result_row(row, subjects, gto, attestat, dvi_score))
    return pd.DataFrame(results)

def count_slots(selected_codes):
    slots = set()
    for code in selected_codes:
        parts = code.split(' ')[0]
        if len(parts) >= 7:
            prefix = parts[:5]
            suffix = parts[5:]
            if suffix == '.00':
                slots.add(prefix + '.00')
            else:
                has_multi = any(c.split(' ')[0] == prefix + '.00' for c in selected_codes)
                slots.add(prefix + '.00' if has_multi else code)
        else:
            slots.add(code)
    return len(slots)
def save_payment_data(order_id, result_df, search_params, user_email, flow, payment_id=None):
    data = {
        "order_id": order_id,
        "payment_id": payment_id,
        "user_email": user_email,
        "flow": flow,
        "search_params": search_params,
        "result": result_df.to_dict(),
        "created_at": datetime.now().isoformat()
    }
    filename = f"payment_{order_id}.json"
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)

def load_payment_data(payment_id):
    """Загружаем данные по payment_id"""
    filename = f"payment_{payment_id}.json"
    if os.path.exists(filename):
        with open(filename, "r", encoding="utf-8") as f:
            return json.load(f)
    return None

def check_payment_status(payment_id):
    """Проверяем статус платежа через API ЮКассы"""
    try:
        payment = Payment.find_one(payment_id)
        return payment.status == "succeeded"
    except:
        return False
def create_payment(amount, description, return_url, order_id=None):
    if order_id is None:
        order_id = str(uuid.uuid4())
    payment = Payment.create({
        "amount": {"value": str(amount), "currency": "RUB"},
        "confirmation": {"type": "redirect", "return_url": return_url},
        "capture": True,
        "description": description,
        "metadata": {"order_id": order_id}
    })
    return payment.confirmation.confirmation_url, payment.id

def show_disclaimers():
    with st.expander("📖 Как читать таблицу и расставлять приоритеты"):
        st.markdown("""
**Расшифровка шансов:**
- 🟢 **Уверенно** — ваш балл выше среднего на 10+. Чаще всего это подстраховочный вариант на который вы точно пройдёте. Рекомендуемый приоритет: **3–5+**
- 🔵 **Реалистично** — ваш балл выше среднего. Хорошие шансы. Рекомендуемый приоритет: **2–3**
- 🟡 **Вероятно** — ваш балл ниже среднего, но выше проходного на 5+. Шансы есть. Рекомендуемый приоритет: **1–2**
- 🔴 **Рискованно** — ваш балл близко к проходному (не более чем на 5 выше или 15 ниже). Ставьте приоритет 1 только если очень хочется и есть варианты с уверенным поступлением
- ⚫ **Маловероятно** — ваш балл ниже проходного на 15+. Шансы крайне малы
- ⬜ **Нет данных** — новая специальность, статистики нет. Ставьте приоритет 1 только если очень хочется и есть варианты с уверенным поступлением

**Про приоритеты:**
В одном вузе можно выбрать не более 5 кодов специальностей, но количество профилей не ограничено — поэтому приоритетов может быть больше пяти.

- 1–2 приоритет: амбициозные, но могут не сработать
- 3–4 приоритет: реалистичные, стабильные
- 4+ приоритет: уверенное зачисление
        """)
    st.warning("""
🔴 **Самое важное — согласие на зачисление**

Зачисление не происходит без подачи согласия. Его нужно подать до **12:00 (мск) 5 августа** в один из вузов — только один одновременно.

Эта таблица поможет выбрать куда подать заявление и документы, но участвовать в конкурсе вы будете только в одном вузе.
    """)
    st.markdown("""
---
**Следующий шаг** — отслеживание позиции в конкурсных списках на сайтах вузов или на Госуслугах. Не бойтесь раздутых списков — не все абитуриенты в них ваши реальные конкуренты.

**Актуальность данных.** Данные актуальны на приёмную кампанию 2026 года. Результаты основаны на статистике прошлого года и носят рекомендательный характер.

**Сроки подачи документов.** Документы в вузы нужно подать до **25 июля**. Срок для ДВИ короче — даты и формат уточняйте на сайте вуза.

**Индивидуальные достижения.** Если у вас есть достижения которые принимает вуз, но они не учтены в нашей таблице — прибавьте их самостоятельно. Помните: суммарно не более 10 баллов. Полный список на сайте вуза.

*Таблица только для поступающих на общих основаниях на основном этапе. Квотники и олимпиадники — вам полезнее персональная консультация.*
    """)
    st.markdown(
        "<p style='font-size:11px; color:#aaaaaa;'>Данный сервис носит исключительно информационный характер и не является официальной консультацией. Результаты подбора основаны на статистических данных прошлых лет и не гарантируют поступление. Приёмная кампания зависит от множества факторов которые невозможно предсказать заранее — статистика прошлых лет обычно хорошо отражает реальность, но никто не застрахован от неожиданных скачков конкурса и изменения проходных баллов. Сервис не несёт ответственности за решения принятые на основе предоставленной информации.</p>",
        unsafe_allow_html=True
    )
def send_email(to_email, result_df, search_params):
    """Отправляем таблицу результатов на email"""
    try:
        # Создаём Excel файл в памяти
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            result_df.to_excel(writer, index=False, sheet_name="Результаты")
            # Лист с параметрами запроса
            params_df = pd.DataFrame([search_params])
            params_df.to_excel(writer, index=False, sheet_name="Запрос")
            workbook = writer.book
            worksheet = writer.sheets["Результаты"]
            num_fmt = workbook.add_format({"num_format": "0"})
            for col_name in ["Мест", "Проходной балл", "Средний балл",
                              "Ваш балл (ЕГЭ)", "Конкурсный балл",
                              "ГТО золото", "ГТО серебро", "ГТО бронза", "Аттестат"]:
                if col_name in result_df.columns:
                    col_idx = result_df.columns.get_loc(col_name)
                    worksheet.set_column(col_idx, col_idx, 12, num_fmt)
        buf.seek(0)

        # Создаём письмо
        msg = MIMEMultipart()
        msg["From"] = st.secrets["EMAIL_FROM"]
        msg["To"] = to_email
        msg["Subject"] = "Ваша таблица подбора вузов — Vuzline"

        body = """
Здравствуйте!

Ваша персональная таблица подбора вузов готова. Она прикреплена к этому письму.

В таблице два листа:
- Результаты — все подходящие специальности с оценкой шансов
- Запрос — параметры вашего поиска

Если у вас возникнут вопросы или вы хотите разобрать результаты подробнее — запишитесь на персональную консультацию со скидкой 500 руб. по промокоду VUZLINE500.
Чтобы узнать подробности или записаться, пишите менеджеру в телеграм-аккауунт @vuzline_webinar.

Удачи с поступлением!
Команда Vuzline
vuzline.ru
        """
        msg.attach(MIMEText(body, "plain", "utf-8"))

        # Прикрепляем Excel
        attachment = MIMEBase("application", "octet-stream")
        attachment.set_payload(buf.read())
        encoders.encode_base64(attachment)
        attachment.add_header(
            "Content-Disposition",
            "attachment",
            filename="vuzline_результаты.xlsx"
        )
        msg.attach(attachment)

        # Отправляем
        with smtplib.SMTP_SSL("smtp.yandex.ru", 465) as server:
            server.login(st.secrets["EMAIL_FROM"], st.secrets["EMAIL_PASSWORD"])
            server.send_message(msg)

        return True
    except Exception as e:
        st.error(f"Ошибка отправки письма: {e}")
        return False
def show_results(result, flow=1, paid=False):
    if len(result) == 0:
        st.warning("По вашему запросу ничего не найдено.")
        return

    st.success(f"Найдено {len(result)} специальностей")
    result = result.copy()
    result["chance_order"] = result["Шансы"].map(CHANCE_ORDER)
    result = result.sort_values(["Город", "Вуз", "chance_order"]).drop("chance_order", axis=1)
    result["Шансы"] = result["Шансы"].map(CHANCE_LABEL)

    if flow == 1:
        vuz_counts = result.groupby("Вуз")["Код и специальность"].nunique()
        overloaded = vuz_counts[vuz_counts > 5]
        if len(overloaded) > 0:
            for vuz, count in overloaded.items():
                st.warning(f"⚠️ {vuz}: найдено {count} кодов специальностей — при подаче документов выберите не более 5.")

    if flow == 2:
        CHANCE_PRIORITY = {
            "🟡 Вероятно": 0, "🔵 Реалистично": 1, "🟢 Уверенно": 2,
            "🔴 Рискованно": 3, "⚫ Маловероятно": 4, "⬜ Нет данных": 5,
            "⬜ Нет оценки — не указан балл за ДВИ": 6,
        }
        result["_chance_p"] = result["Шансы"].map(CHANCE_PRIORITY)
        result["_pb_num"] = pd.to_numeric(result["Проходной балл"], errors="coerce").fillna(0)
        result = result.sort_values(["Город", "Вуз", "_chance_p", "_pb_num"],
                                     ascending=[True, True, True, False])
        result = result.drop(columns=["_pb_num"])

        good_zones = {"🟡 Вероятно", "🔵 Реалистично", "🟢 Уверенно"}
        real_competition = result[
            ~((pd.to_numeric(result["Проходной балл"], errors="coerce") <= 1) &
              (pd.to_numeric(result["Средний балл"], errors="coerce") <= 1))
        ]
        vuz_good_count = real_competition[real_competition["Шансы"].isin(good_zones)].groupby("Вуз").size()
        main_vuz = vuz_good_count[vuz_good_count >= 3].sort_values(ascending=False)
        few_vuz = vuz_good_count[(vuz_good_count >= 1) & (vuz_good_count < 3)].sort_values(ascending=False)
        top_vuz = list(main_vuz.head(7).index)
        few_vuz_list = list(few_vuz.index)

        result_main_rows = []
        for vuz in top_vuz:
            vuz_df = result[result["Вуз"] == vuz].copy()
            seen_codes = set()
            for _, row in vuz_df.iterrows():
                code_prefix = row["Код и специальность"].split(" ")[0][:5]
                if code_prefix not in seen_codes:
                    if len(seen_codes) >= 5: continue
                    seen_codes.add(code_prefix)
                result_main_rows.append(row)
        result_main = pd.DataFrame(result_main_rows).reset_index(drop=True)
        result_main = result_main.drop(columns=["_chance_p"])

        result_few_rows = []
        for vuz in few_vuz_list:
            vuz_df = result[result["Вуз"] == vuz].copy()
            seen_codes = set()
            for _, row in vuz_df.iterrows():
                code_prefix = row["Код и специальность"].split(" ")[0][:5]
                if code_prefix not in seen_codes:
                    if len(seen_codes) >= 5: continue
                    seen_codes.add(code_prefix)
                result_few_rows.append(row)
        result_few = pd.DataFrame(result_few_rows).reset_index(drop=True) if result_few_rows else pd.DataFrame()
        if len(result_few) > 0:
            result_few = result_few.drop(columns=["_chance_p"])

        result = result_main
        st.info(f"Показаны топ-{len(top_vuz)} вузов с наибольшим количеством подходящих специальностей")
        if len(result_few) > 0:
            st.caption(f"Ещё {len(few_vuz_list)} вузов с 1-2 подходящими вариантами показаны ниже")

    if not paid:
        preview_cols = ["Город", "Вуз", "Факультет", "Код и специальность", "Профиль"]
        preview = result[[c for c in preview_cols if c in result.columns]].copy()
        st.dataframe(preview, use_container_width=True, hide_index=True)
        st.info("""
🔒 **Полная таблица доступна после оплаты**

В полной версии вы увидите:
- Проходной и средний балл
- Ваш конкурсный балл
- Оценку шансов и рекомендуемый приоритет
- Баллы за индивидуальные достижения
- Возможность скачать таблицу в Excel

**Стоимость: 1 790 руб.**
        """)
        if st.button("💳 Оплатить и получить полную таблицу", type="primary", key="pay_btn"):
            if not st.session_state.get("user_email"):
                st.error("Введите email для получения таблицы")
            else:
                try:
                    # Генерируем order_id заранее и передаём в return_url
                    order_id = str(uuid.uuid4())
                    return_url = f"https://vuzline-2026.streamlit.app/?order_id={order_id}"
                    payment_url, payment_id = create_payment(
                        amount=1790,
                        description="Подбор вузов по ЕГЭ — полная таблица",
                        return_url=return_url,
                        order_id=order_id
                    )
                    # Сохраняем данные в файл
                    search_params = {
                        "Дата и время": datetime.now().strftime("%d.%m.%Y %H:%M"),
                        "Email": st.session_state.get("user_email", ""),
                        "Предметы и баллы": str(st.session_state.get("last_subjects", {})),
                        "Города": str(st.session_state.get("last_cities", [])),
                        "ГТО": str(st.session_state.get("last_gto", "Нет")),
                        "Аттестат": str(st.session_state.get("last_attestat", False)),
                        "Балл за ДВИ": str(st.session_state.get("last_dvi", "")),
                        "Payment ID": payment_id,
                    }
                    result_df = pd.DataFrame.from_dict(st.session_state["last_result"])
                    save_payment_data(
                        order_id, result_df, search_params,
                        st.session_state["user_email"],
                        st.session_state.get("last_flow", 2),
                        payment_id=payment_id
                    )
                    st.session_state["payment_id"] = payment_id
                    st.session_state["payment_url"] = payment_url
                    st.rerun()
                except Exception as e:
                    st.error(f"Ошибка при создании платежа: {e}")
    else:
        st.dataframe(result, use_container_width=True, hide_index=True)
        show_disclaimers()
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            result.to_excel(writer, index=False, sheet_name="Результаты")
            workbook = writer.book
            worksheet = writer.sheets["Результаты"]
            num_fmt = workbook.add_format({"num_format": "0"})
            for col_name in ["Мест", "Проходной балл", "Средний балл",
                              "Ваш балл (ЕГЭ)", "Конкурсный балл",
                              "ГТО золото", "ГТО серебро", "ГТО бронза", "Аттестат"]:
                if col_name in result.columns:
                    col_idx = result.columns.get_loc(col_name)
                    worksheet.set_column(col_idx, col_idx, 12, num_fmt)
        st.download_button(
            "📥 Скачать таблицу Excel",
            data=buf.getvalue(),
            file_name="результаты_подбора.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("""
💡 **Хотите разобрать результаты вместе?**
Запишитесь на персональную консультацию со скидкой 500 руб. по промокоду **VUZLINE500**
        """)

    if flow == 2 and 'result_few' in dir() and len(result_few) > 0:
        with st.expander("📋 Вузы с 1-2 подходящими вариантами"):
            st.caption("В этих вузах мало подходящих специальностей под ваши предметы и баллы")
            if not paid:
                preview_cols = ["Город", "Вуз", "Факультет", "Код и специальность", "Профиль"]
                preview_few = result_few[[c for c in preview_cols if c in result_few.columns]].copy()
                st.dataframe(preview_few, use_container_width=True, hide_index=True)
            else:
                st.dataframe(result_few, use_container_width=True, hide_index=True)


# ─── ИНТЕРФЕЙС ────────────────────────────────────────────────────────────
df = load_data()
Configuration.account_id = st.secrets["YUKASSA_SHOP_ID"]
Configuration.secret_key = st.secrets["YUKASSA_SECRET_KEY"]
city_options = get_city_options(df)
all_codes = get_all_codes(df)

# Проверяем payment_id из URL и отправляем письмо
query_params = st.query_params
order_id_from_url = query_params.get("order_id", "")

if order_id_from_url and not st.session_state.get(f"sent_{order_id_from_url}"):
    data = load_payment_data(order_id_from_url)
    if data:
        payment_id = data.get("payment_id", "")
        if check_payment_status(payment_id):
            result_df = pd.DataFrame.from_dict(data["result"])
            if send_email(data["user_email"], result_df, data["search_params"]):
                st.session_state[f"sent_{order_id_from_url}"] = True
                st.success(f"✅ Таблица отправлена на {data['user_email']}!")
            else:
                st.error("Ошибка отправки письма. Напишите нам и мы пришлём вручную.")
    else:
        st.info("Платёж обрабатывается... Обновите страницу через минуту.")

st.title("🎓 Подбор вузов по ЕГЭ")

st.info("""
⚠️ **Важно перед использованием**

Этот сервис подходит только для поступающих **на бюджет на общих основаниях на основном этапе** приёмной кампании 2026.

Если вы поступаете по квоте, как победитель олимпиады или с другими особыми условиями — данные прогнозы вам не подойдут. С такими кейсами приходите на персональную консультацию.
""")

# Если есть payment_url — показываем кнопку перехода к оплате
if "payment_url" in st.session_state:
    st.link_button("💳 Перейти к оплате (1 790 руб.)", st.session_state["payment_url"], type="primary")
    st.caption("После оплаты вернитесь на эту страницу — таблица откроется автоматически")
    if st.button("❌ Отменить и начать заново"):
        del st.session_state["payment_url"]
        if "payment_id" in st.session_state:
            del st.session_state["payment_id"]
        st.rerun()
    st.stop()

flow = st.radio(
    "Как вы хотите искать?",
    ["🔍 Подобрать варианты по моим ЕГЭ",
     "🎯 Я знаю вузы и специальности которые хочу"],
    horizontal=True
)

st.divider()

st.subheader("Введите баллы ЕГЭ")
subjects = {}

rus = st.number_input("Русский язык *", min_value=0, max_value=100,
                       value=None, placeholder="Введите балл")
if rus and rus > 0:
    subjects["Русский язык"] = rus

subjects_list = ["Математика","Обществознание","История","Иностранный язык",
                 "Биология","Химия","Физика","Информатика","География","Литература"]
selected_subj = st.multiselect("Выберите остальные предметы *", subjects_list)

if selected_subj:
    cols = st.columns(min(len(selected_subj), 3))
    for i, subj in enumerate(selected_subj):
        with cols[i % 3]:
            score = st.number_input(subj, min_value=0, max_value=100,
                                     value=None, placeholder="Введите балл", key=f"score_{subj}")
            if score and score > 0:
                subjects[subj] = score

st.subheader("Индивидуальные достижения")
st.caption("Суммарно не более 10 баллов")
col1, col2 = st.columns(2)
with col1:
    gto = st.selectbox("ГТО", ["Нет", "Золото", "Серебро", "Бронза"])
with col2:
    attestat = st.checkbox("Аттестат с отличием")

dvi_score = st.number_input(
    "Балл за ДВИ (если уже известен)",
    min_value=0, max_value=1000,
    value=None, placeholder="Необязательно",
    help="ДВИ — дополнительное вступительное испытание в вузе. Если несколько ДВИ — введите суммарный балл."
)

gto_val = gto if gto != "Нет" else None
st.divider()
st.subheader("Введите email для получения таблицы")
user_email = st.text_input(
    "Email *",
    placeholder="example@mail.ru",
    help="На этот адрес мы отправим полную таблицу после оплаты"
)
if user_email:
    st.session_state["user_email"] = user_email

st.divider()

# ─── ФЛОУ 2 ───────────────────────────────────────────────────────────────
if flow == "🔍 Подобрать варианты по моим ЕГЭ":
    st.subheader("Выберите города (до 3)")
    selected_cities = st.multiselect(
        "Города поиска *", city_options, max_selections=3,
        help="Москва и МО / Питер и ЛО идут как один выбор"
    )
    show_dvi = st.toggle(
        "Показывать специальности с ДВИ", value=False,
        help="ДВИ — дополнительное вступительное испытание в вузе"
    )
    if st.button("🔍 Найти специальности", type="primary"):
        errors = []
        if not subjects.get("Русский язык"): errors.append("Введите балл за Русский язык")
        if len(subjects) < 2: errors.append("Введите минимум 2 предмета")
        if not selected_cities: errors.append("Выберите хотя бы один город")
        if errors:
            for e in errors: st.error(e)
        else:
            with st.spinner("Подбираем варианты..."):
                result = filter_rows_flow2(df, subjects, show_dvi, selected_cities, gto_val, attestat, dvi_score)
            if len(result) == 0:
                st.warning("По вашему запросу ничего не найдено.")
            else:
                st.session_state["last_result"] = result.to_dict()
                st.session_state["last_flow"] = 2
                st.session_state["last_subjects"] = subjects
                st.session_state["last_cities"] = selected_cities
                st.session_state["last_gto"] = gto_val
                st.session_state["last_attestat"] = attestat
                st.session_state["last_dvi"] = dvi_score
                st.rerun()

# ─── ФЛОУ 1 ───────────────────────────────────────────────────────────────
else:
    st.subheader("Выберите города (до 3)")
    selected_cities_flow1 = st.multiselect(
        "Города *", city_options, max_selections=3,
        help="Москва и МО / Питер и ЛО идут как один выбор",
        key="cities_flow1"
    )
    st.subheader("Выберите коды специальностей (до 5)")
    if len(subjects) >= 2:
        available_codes = []
        seen = set()
        for _, row in df.iterrows():
            if selected_cities_flow1:
                city_raw = str(row.iloc[22]).strip()
                if get_city_group(city_raw) not in selected_cities_flow1: continue
            status = check_row(row, subjects)
            if status is not None:
                code = clean_str(row.iloc[25])
                if code and code not in seen:
                    seen.add(code)
                    available_codes.append(code)
        available_codes = sorted(available_codes)
        st.caption(f"Показаны специальности подходящие под ваши предметы ({len(available_codes)} из {len(all_codes)})")
    else:
        available_codes = all_codes
        st.caption("Введите предметы ЕГЭ выше чтобы отфильтровать подходящие специальности")

    selected_codes = st.multiselect(
        "Коды специальностей *", available_codes,
        help="Начните вводить название или код для поиска"
    )
    if selected_codes:
        slots = count_slots(selected_codes)
        if slots < len(selected_codes):
            st.info(f"Выбрано {len(selected_codes)} кодов — засчитывается как {slots} слота из 5")
        else:
            st.info(f"Использовано {slots} из 5 слотов")

    st.subheader("Выберите вузы (до 5)")
    if selected_cities_flow1:
        vuz_options = []
        for city_group in selected_cities_flow1:
            vuz_options.extend(get_vuz_by_city(df, city_group))
        vuz_options = sorted(set(vuz_options))
        if selected_codes:
            filtered_vuz = set()
            for _, row in df.iterrows():
                city_raw = str(row.iloc[22]).strip()
                if get_city_group(city_raw) not in selected_cities_flow1: continue
                code = clean_str(row.iloc[25])
                if code in selected_codes:
                    vuz = clean_str(row.iloc[23])
                    if vuz: filtered_vuz.add(vuz)
            vuz_options = sorted(filtered_vuz)
            st.caption(f"Показаны вузы где есть выбранные специальности ({len(vuz_options)})")
    else:
        vuz_options = []

    selected_vuz = st.multiselect("Вузы *", vuz_options, max_selections=5)

    if st.button("🎯 Найти по моему списку", type="primary"):
        errors = []
        if not subjects.get("Русский язык"): errors.append("Введите балл за Русский язык")
        if len(subjects) < 2: errors.append("Введите минимум 2 предмета")
        if not selected_vuz: errors.append("Выберите хотя бы один вуз")
        if not selected_codes: errors.append("Выберите хотя бы один код специальности")
        if errors:
            for e in errors: st.error(e)
        else:
            with st.spinner("Ищем по вашему списку..."):
                result = filter_rows_flow1(df, subjects, selected_vuz, selected_codes,
                                           gto_val, attestat, selected_cities_flow1, dvi_score)
            if len(result) == 0:
                st.warning("По выбранным вузам и специальностям ничего не найдено.")
            else:
                st.session_state["last_result"] = result.to_dict()
                st.session_state["last_flow"] = 1
                st.session_state["last_subjects"] = subjects
                st.session_state["last_cities"] = selected_cities_flow1
                st.session_state["last_gto"] = gto_val
                st.session_state["last_attestat"] = attestat
                st.session_state["last_dvi"] = dvi_score
                st.rerun()
# После rerun показываем результаты и кнопку оплаты
if "last_result" in st.session_state and "last_flow" in st.session_state:
    if "payment_url" not in st.session_state:
        result = pd.DataFrame.from_dict(st.session_state["last_result"])
        paid = st.session_state.get("paid", False)
        show_results(result, flow=st.session_state["last_flow"], paid=paid)