"""
Тестовое задание AK Furs — детерминированный решатель.
Использование: python solve.py data/ > answers.json
Все логи пишутся в stderr. В stdout выходит только JSON.
"""

import csv
import json
import os
import re
import sys
import zipfile
from collections import defaultdict
from datetime import datetime, timedelta


# ---------------------------------------------------------------------------
# Загрузка данных
# ---------------------------------------------------------------------------

def load_crm(path):
    """Загружает CRM-выгрузку. Возвращает список словарей."""
    rows = []
    with open(path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            rows.append({k: v for k, v in row.items()})
    print(f"[CRM] загружено {len(rows)} сырых строк", file=sys.stderr)
    return rows


def load_bank(path):
    """Загружает банковскую выписку (Поступления + Списания). Пропускает строку-заголовок счёта."""
    rows = []
    with open(path, encoding="cp1251") as f:
        reader = csv.reader(f, delimiter=";")
        next(reader)  # пропустить «Выписка по счёту...»
        header = next(reader)
        for row in reader:
            if len(row) >= len(header):
                rows.append(dict(zip(header, row)))
    print(f"[Банк] загружено {len(rows)} сырых строк", file=sys.stderr)
    return rows


def load_xlsx_sheets(path):
    """Открывает xlsx со всеми листами, включая скрытые."""
    import openpyxl
    return openpyxl.load_workbook(path)


def load_managers(wb):
    """
    Строит словарь: вариант_написания_lower -> каноническое_имя.
    Варианты разделены « | » в листе «Менеджеры».
    """
    ws = wb["Менеджеры"]
    mgr_map = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        canonical, variants_raw = row[0], row[1]
        if not canonical:
            continue
        for variant in str(variants_raw).split("|"):
            v = variant.strip().lower()
            if v:
                mgr_map[v] = canonical
    print(f"[Менеджеры] загружено {len(mgr_map)} вариантов написания", file=sys.stderr)
    return mgr_map


def load_renames(wb):
    """Строит словарь: старый_артикул -> новый_артикул из листа «Переименования»."""
    ws = wb["Переименования"]
    renames = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        old, new = row[0], row[1]
        if old and new:
            renames[old] = new
    print(f"[Переименования] загружено {len(renames)} записей", file=sys.stderr)
    return renames


def load_prices(wb):
    """
    Строит словарь: новый_артикул -> список (дата_с, дата_по, мин_цена).
    дата_по = None означает «действует по сей день» (открытый правый край).
    В Прайс_история артикулы хранятся только под НОВЫМ именем.
    """
    ws = wb["Прайс_история"]
    prices = defaultdict(list)
    for row in ws.iter_rows(min_row=2, values_only=True):
        art, d_from, d_to, _, min_p = row[0], row[1], row[2], row[3], row[4]
        if art and d_from is not None and min_p is not None:
            prices[art].append((d_from, d_to, float(min_p)))
    print(f"[Прайс] загружено {len(prices)} артикулов", file=sys.stderr)
    return dict(prices)


# ---------------------------------------------------------------------------
# Парсеры и нормализаторы
# ---------------------------------------------------------------------------

_RU_MONTHS = {
    "янв": 1, "фев": 2, "мар": 3, "апр": 4, "май": 5, "мая": 5,
    "июн": 6, "июл": 7, "авг": 8, "сен": 9, "окт": 10, "ноя": 11, "дек": 12,
}


def parse_date(s):
    """
    Парсит строку с датой в datetime. При неудаче возвращает None.
    Форматы: ISO, DD.MM.YYYY, DD/MM/YYYY, Excel-serial (5 цифр), «23 сен 2026».
    """
    if not s:
        return None
    s = s.strip()
    # ISO: 2026-03-03
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:
        return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    # DD.MM.YYYY: 09.01.2026
    m = re.match(r"^(\d{2})\.(\d{2})\.(\d{4})$", s)
    if m:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    # DD/MM/YYYY: 31/10/2025 (подтверждено данными: 138 строк с днём >12)
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", s)
    if m:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    # Excel serial: 45752 → база 1899-12-30 + дней
    m = re.match(r"^\d{5}$", s)
    if m:
        return datetime(1899, 12, 30) + timedelta(days=int(s))
    # Русский текст: «23 сен 2026» или «23 сентября 2026»
    m = re.match(r"^(\d{1,2})\s+([а-яё]+)\.?\s+(\d{4})$", s, re.IGNORECASE)
    if m:
        day, mon_str, year = int(m.group(1)), m.group(2).lower()[:3], int(m.group(3))
        month = _RU_MONTHS.get(mon_str)
        if month:
            return datetime(year, month, day)
    print(f"[parse_date] не удалось разобрать: {s!r}", file=sys.stderr)
    return None


def normalize_order_num(s):
    """
    Извлекает целочисленный номер заказа из любого формата.
    A-00075, А-75 (кириллица), a-75, A75 → 75. При неудаче возвращает None.
    """
    m = re.search(r"\d+", s.strip())
    if m:
        return int(m.group())
    return None


def normalize_status(s):
    """Приводит сырой статус к одному из 4 канонических значений."""
    s = s.strip().upper()
    if s in ("ОПЛАЧЕН", "PAID", "ОПЛАЧЕНО"):
        return "Оплачен"
    if s in ("ОТМЕНЁН", "ОТМЕНЕН", "CANCELLED"):
        return "Отменён"
    if s in ("ВОЗВРАТ", "REFUND"):
        return "Возврат"
    return "В работе"


def normalize_manager(s, mgr_map):
    """Возвращает каноническое имя менеджера или исходную строку при неизвестном варианте."""
    key = s.strip().lower()
    result = mgr_map.get(key)
    if result is None:
        print(f"[Менеджер] неизвестный вариант: {s!r}", file=sys.stderr)
    return result or s.strip()


def parse_price_crm(s):
    """
    Парсит строку с ценой из CRM в рубли (float).
    Поддерживает: '50 000', '17,000.00', '245,000.00', '2 200 000',
                  '140000,00', '2,48 млн', '500000,00', '32,000.00'.
    """
    if not s:
        return None
    s = s.strip()
    # Миллионы: «2,48 млн»
    m = re.match(r"^([\d][0-9\s,\.]+)\s*млн$", s, re.IGNORECASE)
    if m:
        num_str = m.group(1).strip().replace(" ", "").replace(",", ".")
        return float(num_str) * 1_000_000

    s = s.replace(" ", "")

    # Если есть и запятая и точка — запятая = разделитель тысяч, точка = десятичный
    if "," in s and "." in s:
        s = s.replace(",", "")

    # Точки как разделители тысяч: 2.094.000 или 385.000
    # Правило: если все части после первой точки ровно по 3 цифры — это тысячи
    parts = s.split(".")
    if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
        s = s.replace(".", "")
    else:
        s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        print(f"[parse_price_crm] не удалось разобрать: {s!r}", file=sys.stderr)
        return None


def parse_price_chat(s):
    """
    Извлекает первую цену (в рублях) из сообщения менеджера в чате.
    Поддерживает: '53 000 ₽', '138 тыс.', '2,404 млн', '385.000 руб',
                  'сто пятьдесят восемь тысяч рублей', '2.094.000 руб'.
    Возвращает None, если цены нет или текст содержит только время/проценты.
    """
    # Убираем упоминания времени вида «15:00», чтобы не принять за число
    text = re.sub(r"\b\d{1,2}:\d{2}\b", "", s)

    # Сначала пробуем прописные числа («сто двадцать тысяч рублей»)
    written = _parse_written_rubles(text)
    if written is not None:
        return written

    # Числовые паттерны с единицей измерения
    pattern = re.compile(
        r"([\d][\d\s]*(?:[,\.][\d]+)*(?:\.\d{3})*)"
        r"\s*(₽|руб\.?|тыс\.?|к\b|млн\.?)",
        re.IGNORECASE,
    )
    m = pattern.search(text)
    if not m:
        return None

    num_str = m.group(1).strip()
    unit = m.group(2).lower().rstrip(".")

    num_str = num_str.replace(" ", "")
    dot_parts = num_str.split(".")
    if len(dot_parts) > 1 and all(len(p) == 3 for p in dot_parts[1:]):
        # Точки как разделители тысяч
        num_str = num_str.replace(".", "")
    else:
        num_str = num_str.replace(",", ".").replace(".", ".", 1)
        parts2 = num_str.split(".")
        num_str = parts2[0].replace(".", "") + ("." + parts2[1] if len(parts2) > 1 else "")

    try:
        value = float(num_str.replace(",", ""))
    except ValueError:
        return None

    multipliers = {"тыс": 1_000, "к": 1_000, "млн": 1_000_000, "₽": 1, "руб": 1}
    for key, mult in multipliers.items():
        if unit.startswith(key):
            return value * mult

    return value


_RU_ONES = {
    "один": 1, "одна": 1, "два": 2, "две": 2, "три": 3, "четыре": 4,
    "пять": 5, "шесть": 6, "семь": 7, "восемь": 8, "девять": 9,
    "десять": 10, "одиннадцать": 11, "двенадцать": 12, "тринадцать": 13,
    "четырнадцать": 14, "пятнадцать": 15, "шестнадцать": 16,
    "семнадцать": 17, "восемнадцать": 18, "девятнадцать": 19,
    "двадцать": 20, "тридцать": 30, "сорок": 40, "пятьдесят": 50,
    "шестьдесят": 60, "семьдесят": 70, "восемьдесят": 80, "девяносто": 90,
    "сто": 100, "двести": 200, "триста": 300, "четыреста": 400,
    "пятьсот": 500, "шестьсот": 600, "семьсот": 700, "восемьсот": 800,
    "девятьсот": 900,
}


def _parse_written_rubles(text):
    """
    Парсит сумму прописью: «сто пятьдесят восемь тысяч рублей» → 158000.0.
    Возвращает float или None.
    """
    text_lower = text.lower()
    if "тысяч" not in text_lower and "миллион" not in text_lower:
        return None

    tokens = re.findall(r"[а-яё]+", text_lower)
    total = 0
    current = 0
    for tok in tokens:
        if tok in _RU_ONES:
            current += _RU_ONES[tok]
        elif tok in ("тысяч", "тысячи", "тысяча"):
            total += (current or 1) * 1_000
            current = 0
        elif tok in ("миллион", "миллиона", "миллионов"):
            total += (current or 1) * 1_000_000
            current = 0
        elif tok in ("рублей", "рубля", "рубль"):
            total += current
            current = 0

    total += current
    return float(total) if total > 0 else None


def calc_discounted_unit_price(price_per_unit, qty, discount_str):
    """
    Применяет скидку к цене за единицу.
    «N%» — процент от суммы строки (цена × кол-во), результат на единицу.
    Число без «%» — фиксированная сумма в рублях на всю строку.
    Возвращает цену за единицу после скидки.
    """
    if not discount_str or not discount_str.strip():
        return price_per_unit
    d = discount_str.strip()
    if d.endswith("%"):
        pct = float(d[:-1].replace(",", ".").replace(" ", ""))
        total_discount = price_per_unit * qty * pct / 100
        return price_per_unit - total_discount / qty
    else:
        # Фиксированная скидка в рублях на строку — используем тот же парсер
        amt = parse_price_crm(d)
        if amt is None:
            print(f"[скидка] не удалось разобрать: {d!r}", file=sys.stderr)
            return price_per_unit
        return price_per_unit - amt / qty


def extract_order_from_payment(text, known_orders):
    """
    Извлекает номер заказа из назначения платежа.
    Возвращает int, если номер найден среди известных заказов, иначе None.
    Поддерживает: 'A877', 'заказ 775', 'зак. 528', 'A-00243',
                  'по счету №828', 'Счёт № 00991', 'А-00391' (кириллическая А).
    """
    # Паттерн с префиксом A/А (латиница или кириллица) и необязательным дефисом
    candidates = re.findall(
        r"[AАaа]-?0*(\d+)",
        text,
        re.IGNORECASE,
    )
    # Числа после ключевых слов: «заказ», «зак.», «счёт №»
    candidates += re.findall(
        r"(?:заказ[а-я\s]*|зак\.?\s*|счет[а-я\s]*[№#\s]*|счёт[а-я\s]*[№#\s]*)0*(\d+)",
        text,
        re.IGNORECASE,
    )
    for c in candidates:
        num = int(c)
        if num in known_orders:
            return num
    return None


def is_invite(text):
    """
    Возвращает True, если сообщение менеджера — приглашение на встречу или примерку.
    Исключает информационные фразы: «шьём в нашем ателье», «шоурум находится у метро».
    """
    t = text.lower()
    # Не приглашения — общая информация об ателье/адресе
    if "шьём в нашем ателье" in t:
        return False
    if "находится у метро" in t or "находится по адрес" in t:
        return False
    # Паттерны явного приглашения
    invite_patterns = [
        r"прим[её]рк",          # примерку, примерки
        r"подъехать",
        r"приезжайт",
        r"приглаш",
        r"запиш.{0,10}встреч",  # «запишу вас на встречу»
        r"встреч.{0,10}шоурум", # «встречу в шоурум»
    ]
    return any(re.search(p, t) for p in invite_patterns)


# ---------------------------------------------------------------------------
# Дедупликация CRM
# ---------------------------------------------------------------------------

def deduplicate_crm(rows):
    """
    Оставляет одну строку на пару (номер_заказа, № строки) — с самым поздним «Выгружено».
    Строки с одинаковым «Выгружено» — полные дубли, берём первую.
    Добавляет служебные поля: _order_num (int), _status (канонический статус).
    """
    groups = defaultdict(list)
    skipped = 0
    for row in rows:
        num = normalize_order_num(row["Номер заказа"])
        if num is None:
            skipped += 1
            continue
        line = row["№ строки"].strip()
        try:
            ts = datetime.strptime(row["Выгружено"].strip(), "%Y-%m-%d %H:%M")
        except ValueError:
            ts = datetime.min
        groups[(num, line)].append((ts, row))

    if skipped:
        print(f"[CRM] пропущено {skipped} строк с нечитаемым номером заказа", file=sys.stderr)

    clean = []
    for (num, line), entries in groups.items():
        # Сортируем по убыванию времени, берём первую (самую позднюю)
        entries.sort(key=lambda x: x[0], reverse=True)
        _, best_row = entries[0]
        best_row = dict(best_row)
        best_row["_order_num"] = num
        best_row["_status"] = normalize_status(best_row["Статус"])
        clean.append(best_row)

    print(f"[CRM] после дедупликации: {len(clean)} уникальных пар (заказ, строка)", file=sys.stderr)
    return clean


# ---------------------------------------------------------------------------
# Q1 — уникальные заказы и строки
# ---------------------------------------------------------------------------

def solve_q1(clean_rows):
    unique_orders = len(set(r["_order_num"] for r in clean_rows))
    unique_lines = len(clean_rows)
    return {
        "q1_unique_orders": unique_orders,
        "q1_unique_order_lines": unique_lines,
    }


# ---------------------------------------------------------------------------
# Q2 — сверка с банком
# ---------------------------------------------------------------------------

def solve_q2(clean_rows, bank_rows):
    known_orders = set(r["_order_num"] for r in clean_rows)
    paid_orders = set(r["_order_num"] for r in clean_rows if r["_status"] == "Оплачен")

    # Только поступления (возвраты — Списания, не считаются)
    incoming = [r for r in bank_rows if r["Тип"].strip() == "Поступление"]
    print(f"[Q2] поступлений: {len(incoming)} из {len(bank_rows)} строк выписки", file=sys.stderr)

    payment_orders = set()
    unidentified = []
    for row in incoming:
        ref = row["Назначение платежа"].strip()
        order_num = extract_order_from_payment(ref, known_orders)
        if order_num is not None:
            payment_orders.add(order_num)
        else:
            unidentified.append(row)

    # (a) Оплаченные заказы без поступлений в выписке
    paid_without_payment = sorted(paid_orders - payment_orders)

    # (b) Неопознанные поступления
    unidentified_sum = 0.0
    for row in unidentified:
        amt_str = row["Сумма"].strip().replace(" ", "").replace(",", ".")
        try:
            unidentified_sum += float(amt_str)
        except ValueError:
            print(f"[Q2] не удалось разобрать сумму: {row['Сумма']!r}", file=sys.stderr)

    print(
        f"[Q2] оплачен без поступления: {len(paid_without_payment)}, "
        f"неопознанных: {len(unidentified)}, сумма: {unidentified_sum:.2f}",
        file=sys.stderr,
    )
    return {
        "q2a_paid_orders_without_payment": paid_without_payment,
        "q2b_unidentified_payments_count": len(unidentified),
        "q2b_unidentified_payments_sum_rub": round(unidentified_sum, 2),
    }


# ---------------------------------------------------------------------------
# Q3 — ценовая политика
# ---------------------------------------------------------------------------

def solve_q3(clean_rows, prices, renames, mgr_map):
    violations_total = 0
    violations_by_manager = defaultdict(int)
    skipped_no_price = []

    for row in clean_rows:
        if row["_status"] == "Отменён":
            continue

        art_raw = row["Артикул"].strip()
        # Резолвим старый артикул → новый через таблицу переименований
        art = renames.get(art_raw, art_raw)

        order_date = parse_date(row["Дата создания"])
        if order_date is None:
            continue

        price_raw = parse_price_crm(row["Цена за ед."])
        if price_raw is None:
            continue

        try:
            qty = int(row["Кол-во"].strip())
        except ValueError:
            qty = 1

        discounted = calc_discounted_unit_price(price_raw, qty, row["Скидка"])

        # Находим минимально допустимую цену на дату создания заказа
        periods = prices.get(art)
        if not periods:
            skipped_no_price.append(art_raw)
            continue

        min_price = None
        for d_from, d_to, mp in periods:
            if d_from <= order_date and (d_to is None or order_date <= d_to):
                min_price = mp
                break

        if min_price is None:
            skipped_no_price.append(art_raw)
            continue

        if discounted < min_price:
            violations_total += 1
            manager = normalize_manager(row["Менеджер"], mgr_map)
            violations_by_manager[manager] += 1

    if skipped_no_price:
        unique_skipped = sorted(set(skipped_no_price))
        print(
            f"[Q3] пропущено {len(skipped_no_price)} строк (нет записи в прайсе): {unique_skipped}",
            file=sys.stderr,
        )

    return {
        "q3_price_violation_lines_total": violations_total,
        "q3_price_violation_lines_by_manager": dict(violations_by_manager),
    }


# ---------------------------------------------------------------------------
# Q4 — переписки: приглашение до цены
# ---------------------------------------------------------------------------

def solve_q4(zip_path):
    chats_invite_before_price = []
    ambiguous = []
    debug_lines = []

    with zipfile.ZipFile(zip_path) as z:
        for name in sorted(z.namelist()):
            if not name.endswith(".txt"):
                continue
            with z.open(name) as f:
                content = f.read().decode("utf-8")

            lines = content.strip().split("\n")
            m_id = re.search(r"Чат (C\d+)", lines[0])
            chat_id = m_id.group(1) if m_id else name

            first_price_ts = first_price_quote = None
            first_invite_ts = first_invite_quote = None

            for line in lines[1:]:
                m = re.match(r"\[(.+?)\] (Менеджер|Клиент): (.+)", line)
                if not m:
                    continue
                ts_str, role, text = m.group(1), m.group(2), m.group(3)
                try:
                    ts = datetime.strptime(ts_str, "%d.%m.%Y %H:%M")
                except ValueError:
                    continue

                if role == "Менеджер":
                    # Фиксируем первую цену менеджера
                    if first_price_ts is None:
                        price = parse_price_chat(text)
                        if price is not None:
                            first_price_ts = ts
                            first_price_quote = text
                    # Фиксируем первое приглашение менеджера
                    if first_invite_ts is None and is_invite(text):
                        first_invite_ts = ts
                        first_invite_quote = text

            # Классификация чата
            if first_invite_ts is None and first_price_ts is None:
                ambiguous.append((chat_id, "нет ни цены, ни приглашения"))
            elif first_invite_ts is None:
                ambiguous.append((chat_id, "нет приглашения"))
            elif first_price_ts is None:
                ambiguous.append((chat_id, "нет цены от менеджера"))
            elif first_invite_ts < first_price_ts:
                chats_invite_before_price.append(chat_id)
                debug_lines.append(
                    f"\n--- {chat_id} ---\n"
                    f"  Приглашение [{first_invite_ts:%d.%m.%Y %H:%M}]: {first_invite_quote}\n"
                    f"  Цена        [{first_price_ts:%d.%m.%Y %H:%M}]: {first_price_quote}"
                )

    # Записываем отладочный файл с цитатами
    debug_path = os.path.normpath(os.path.join(os.path.dirname(zip_path), "..", "debug_q4.txt"))
    with open(debug_path, "w", encoding="utf-8") as f:
        f.write(f"Чатов с приглашением ДО цены: {len(chats_invite_before_price)}\n")
        f.write("\n".join(debug_lines))

    # Неоднозначные чаты — в stderr для ручной проверки
    if ambiguous:
        print(f"\n[Q4] Неоднозначные чаты ({len(ambiguous)}):", file=sys.stderr)
        for cid, reason in ambiguous:
            print(f"  {cid}: {reason}", file=sys.stderr)

    print(
        f"[Q4] приглашение до цены: {len(chats_invite_before_price)}, "
        f"неоднозначных: {len(ambiguous)}",
        file=sys.stderr,
    )
    return {"q4_chats_meeting_before_price": sorted(chats_invite_before_price)}


# ---------------------------------------------------------------------------
# Поле notes — принятые решения
# ---------------------------------------------------------------------------

def build_notes():
    return (
        "Принятые решения: "
        "(1) Даты DD/MM/YYYY (подтверждено данными: 138 строк с днём >12). "
        "(2) Excel serial: база 1899-12-30 + дней. "
        "(3) Дубли (заказ, строка): берём строку с самым поздним Выгружено; "
        "26 пар с равным Выгружено — полные дубли, берём первую. "
        "(4) Статусы: PAID/оплачено - Оплачен, отменен/CANCELLED - Отменён, "
        "REFUND/возврат - Возврат, остальное - В работе. "
        "(5) Скидка N%: от суммы строки (цена*кол-во); число без % - рублей на строку. "
        "(6) Счёт No = номер заказа (173/173 совпали с CRM). "
        "(7) Возвраты в выписке - Списания, не считаются Поступлениями. "
        "(8) Артикулы в Прайс_история - только под НОВЫМ именем; "
        "старый артикул резолвится через Переименования. "
        "(9) Строка Q3 без записи в прайсе - пропускается (в данных таких нет). "
        "(10) Q4: 'шьём в нашем ателье' и 'шоурум находится у метро' - "
        "не приглашения. Дубли в выписке: 0. "
        "(11) Q2b: 4 из 16 неопознанных платежей содержат номера заказов (A-01169, "
        "A-01180, A-01191, A-01202), которых нет в CRM - по правилу считаются неопознанными."
    )


# ---------------------------------------------------------------------------
# Точка входа
# ---------------------------------------------------------------------------

def main(data_dir):
    crm_path   = os.path.join(data_dir, "01_crm_orders_export.csv")
    bank_path  = os.path.join(data_dir, "02_bank_statement_rub.csv")
    xlsx_path  = os.path.join(data_dir, "03_catalog_prices_reference.xlsx")
    chats_path = os.path.join(data_dir, "04_client_chats.zip")

    wb        = load_xlsx_sheets(xlsx_path)
    mgr_map   = load_managers(wb)
    renames   = load_renames(wb)
    prices    = load_prices(wb)
    crm_raw   = load_crm(crm_path)
    bank_rows = load_bank(bank_path)

    clean = deduplicate_crm(crm_raw)

    result = {}
    result.update(solve_q1(clean))
    result.update(solve_q2(clean, bank_rows))
    result.update(solve_q3(clean, prices, renames, mgr_map))
    result.update(solve_q4(chats_path))
    result["run_command"] = "python solve.py data/ > answers.json"
    result["notes"] = build_notes()

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    if len(sys.argv) != 2:
        print("Использование: python solve.py data/", file=sys.stderr)
        sys.exit(1)
    main(sys.argv[1])
