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

# Счётчик предупреждений — инкрементируется по всему коду, выводится в сводке в конце
_warn = defaultdict(int)


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
    """Загружает банковскую выписку (Поступления + Списания).
    Если первая строка — «Выписка по счёту…», пропускает её; иначе трактует как заголовок."""
    rows = []
    with open(path, encoding="cp1251") as f:
        reader = csv.reader(f, delimiter=";")
        first = next(reader)
        # Первая ячейка — «Выписка по счёту ...» (прочие поля пустые) ⇒ титульная строка
        if first and first[0].strip().lower().startswith("выписка"):
            header = next(reader)
        else:
            header = first
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
    Строит словарь: новый_артикул -> список (дата_с, дата_по, мин_цена),
    отсортированный по дата_с (от старых к новым). Это защищает Q3 от
    случайного порядка строк в xlsx при перекрытии периодов.
    дата_по = None означает «действует по сей день».
    В Прайс_история артикулы хранятся только под НОВЫМ именем.
    """
    ws = wb["Прайс_история"]
    prices = defaultdict(list)
    for row in ws.iter_rows(min_row=2, values_only=True):
        art, d_from, d_to, _, min_p = row[0], row[1], row[2], row[3], row[4]
        if art and d_from is not None and min_p is not None:
            prices[art].append((d_from, d_to, float(min_p)))
    # Сортировка периодов по дате начала — критично при перекрывающихся диапазонах
    result = {art: sorted(periods, key=lambda t: t[0]) for art, periods in prices.items()}
    print(f"[Прайс] загружено {len(result)} артикулов", file=sys.stderr)
    return result


# ---------------------------------------------------------------------------
# Парсеры и нормализаторы
# ---------------------------------------------------------------------------

_RU_MONTHS = {
    "янв": 1, "фев": 2, "мар": 3, "апр": 4, "май": 5, "мая": 5,
    "июн": 6, "июл": 7, "авг": 8, "сен": 9, "окт": 10, "ноя": 11, "дек": 12,
}


def _safe_date(year, month, day):
    """Создаёт datetime; при недопустимых значениях возвращает None (не падает)."""
    try:
        return datetime(year, month, day)
    except (ValueError, OverflowError):
        return None


def parse_date(s):
    """
    Парсит строку с датой в datetime. При неудаче возвращает None.
    Форматы: ISO, D[D].M[M].YYYY, D[D]/M[M]/YYYY, Excel-serial (4-6 цифр), «23 сен 2026».
    Если строка содержит время («2026-03-03 12:34»), оно отбрасывается — берётся только дата.
    Невалидные даты («30.02.2026») логируются как warn, но не роняют скрипт.
    """
    if not s:
        return None
    s = s.strip()
    # Отбрасываем всё после первого пробела (на случай «2026-03-03 12:34» или «09.01.2026 15:00»)
    # Исключение — русские «23 сен 2026», где пробелы часть формата: пробуем сначала с ними.
    def _try_all(txt):
        # ISO: 2026-03-03
        m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", txt)
        if m:
            return _safe_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        # DD.MM.YYYY или D.M.YYYY: 09.01.2026 / 9.1.2026
        m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$", txt)
        if m:
            return _safe_date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        # DD/MM/YYYY или D/M/YYYY (подтверждено данными: 138 строк с днём >12)
        m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", txt)
        if m:
            return _safe_date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        # Excel serial: 4-6 цифр (1927-2170 годы) → база 1899-12-30 + дней
        m = re.match(r"^\d{4,6}$", txt)
        if m:
            try:
                return datetime(1899, 12, 30) + timedelta(days=int(txt))
            except (ValueError, OverflowError):
                return None
        return None

    # Пробуем сначала «русский текст» (с пробелами в формате)
    m = re.match(r"^(\d{1,2})\s+([а-яё]+)\.?\s+(\d{4})$", s, re.IGNORECASE)
    if m:
        day, mon_str, year = int(m.group(1)), m.group(2).lower()[:3], int(m.group(3))
        month = _RU_MONTHS.get(mon_str)
        if month:
            d = _safe_date(year, month, day)
            if d:
                return d

    # Остальные форматы — пробуем сперва всю строку, затем «только дату» (без времени)
    result = _try_all(s)
    if result:
        return result
    date_part = s.split()[0] if s.split() else s
    if date_part != s:
        result = _try_all(date_part)
        if result:
            return result

    _warn["даты_не_разобраны"] += 1
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
        _warn["менеджеры_не_опознаны"] += 1
        print(f"[Менеджер] неизвестный вариант: {s!r}", file=sys.stderr)
    return result or s.strip()


def parse_price_crm(s):
    """
    Парсит строку с ценой из CRM в рубли (float).
    Поддерживает: '50 000', '17,000.00', '245,000.00', '2 200 000',
                  '140000,00', '2,48 млн', '500000,00', '32,000.00',
                  '25,000' (US-формат, запятая = тысячи).
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
    elif s.count(",") >= 1 and "." not in s:
        # Только запятые: если ВСЕ фрагменты после запятых по 3 цифры — это тысячи,
        # иначе одна запятая = десятичный разделитель (русский формат).
        comma_parts = s.split(",")
        if len(comma_parts) > 2 and all(len(p) == 3 for p in comma_parts[1:]):
            # '1,234,567' — тысячи
            s = s.replace(",", "")
        elif len(comma_parts) == 2 and len(comma_parts[1]) == 3:
            # '25,000' — тысячи (эвристика: 3 цифры после единственной запятой)
            s = s.replace(",", "")
        else:
            # '140000,00' — десятичный разделитель
            s = s.replace(",", ".")

    # Точки как разделители тысяч: 2.094.000 или 385.000
    parts = s.split(".")
    if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
        s = s.replace(".", "")

    try:
        return float(s)
    except ValueError:
        _warn["цены_crm_не_разобраны"] += 1
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
        r"\s*(₽|руб\.?|рублей|рубля|рубль|тыс\.?|т\.р\.?|тр\b|к\b|млн\.?|rub|rur|р\.|p\.)",
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

    multipliers = {
        "тыс": 1_000, "т.р": 1_000, "тр": 1_000, "к": 1_000,
        "млн": 1_000_000,
        "₽": 1, "руб": 1, "рубл": 1, "rub": 1, "rur": 1, "р.": 1, "p.": 1,
    }
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
    При qty <= 0 возвращает исходную цену (защита от деления на ноль).
    """
    if not discount_str or not discount_str.strip():
        return price_per_unit
    if qty <= 0:
        _warn["qty_ноль_или_отриц"] += 1
        print(f"[скидка] qty={qty} — пропускаем расчёт скидки", file=sys.stderr)
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
            _warn["скидки_не_разобраны"] += 1
            print(f"[скидка] не удалось разобрать: {d!r}", file=sys.stderr)
            return price_per_unit
        return price_per_unit - amt / qty


def extract_order_from_payment(text, known_orders):
    """
    Извлекает номер заказа из назначения платежа.
    Возвращает int, если номер найден среди известных заказов, иначе None.
    Поддерживает: 'A877', 'заказ 775', 'зак. 528', 'A-00243',
                  'по счету №828', 'Счёт № 00991', 'А-00391' (кириллическая А).
    Префикс A/А должен быть на границе слова, чтобы не ловить случайные
    совпадения внутри слов (например, «Автор», «iPhone»).
    """
    # Паттерн: A/А на границе слова + необязательный дефис + необязательные нули + цифры
    candidates = re.findall(
        r"(?:^|[\s,.;:()№#])[AАaа]-?0*(\d+)",
        text,
        re.IGNORECASE,
    )
    # Числа после ключевых слов: «заказ», «зак.», «счёт №»
    candidates += re.findall(
        r"(?:заказ[а-я\s]*|зак\.?\s*|счет[а-я\s]*[№#\s]*|счёт[а-я\s]*[№#\s]*)0*(\d+)",
        text,
        re.IGNORECASE,
    )
    # Голое «№<число>» или «No<число>» (без слова «счёт»/«заказ» перед ним)
    candidates += re.findall(r"(?:№|#|\bno\.?\s*)0*(\d+)", text, re.IGNORECASE)
    for c in candidates:
        num = int(c)
        if num in known_orders:
            return num
    return None


def is_invite(text):
    """
    Возвращает True, если сообщение менеджера — приглашение на встречу или примерку.
    Исключает информационные фразы (о работе ателье, адресе шоурума), которые
    упоминают ателье/шоурум без приглашения.
    """
    t = text.lower()
    # Не приглашения — общая информация об ателье/адресе/услугах
    if "шьём в нашем ателье" in t or "шьем в нашем ателье" in t:
        return False
    if "находится у метро" in t or "находится по адрес" in t:
        return False
    # Паттерны явного приглашения
    invite_patterns = [
        r"прим[её]рк",                   # примерку, примерки
        r"подъехать",
        r"приезжайт",
        r"приглаш",
        r"приход(?:ите|и)\b",            # «приходите/приходи» — только императив, не «приходит»
        r"прийти",
        r"жд[её]м\s+вас",                # «ждём вас»
        r"бронир\w*\s+(?:визит|примерк)",# «забронируйте визит/примерку»
        r"запиш\w*.{0,10}на\s+(?:встреч|примерк|визит)",
        r"встреч\w*.{0,10}шоурум",
        r"посмотр\w*\s+вживую",
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

def _is_rub(currency_str):
    """Проверяет, что строка-валюта обозначает рубли (₽, RUB, руб)."""
    if currency_str is None:
        return True  # отсутствует колонка валюты — считаем ₽
    c = currency_str.strip().upper().rstrip(".")
    return c in ("", "₽", "RUB", "RUR", "РУБ", "РУБЛЬ", "РУБЛИ", "РУБЛЕЙ")


def solve_q2(clean_rows, bank_rows):
    known_orders = set(r["_order_num"] for r in clean_rows)
    paid_orders = set(r["_order_num"] for r in clean_rows if r["_status"] == "Оплачен")

    # Только поступления в рублях (возвраты — Списания; валюты ≠ ₽ не учитываем)
    incoming = []
    for r in bank_rows:
        if r["Тип"].strip() != "Поступление":
            continue
        if not _is_rub(r.get("Валюта")):
            _warn["банк_не_рубли"] += 1
            print(f"[Q2] не-рублёвое поступление: {r.get('Валюта')!r} — пропускаем", file=sys.stderr)
            continue
        incoming.append(r)
    print(f"[Q2] поступлений в ₽: {len(incoming)} из {len(bank_rows)} строк выписки", file=sys.stderr)

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

    # (b) Неопознанные поступления — используем общий парсер цен (а не упрощённую replace-логику),
    # чтобы корректно обработать форматы "1,000,000.00", "145 000,50" и прочие.
    unidentified_sum = 0.0
    for row in unidentified:
        amt = parse_price_crm(row["Сумма"])
        if amt is None:
            _warn["суммы_банк_не_разобраны"] += 1
            print(f"[Q2] не удалось разобрать сумму: {row['Сумма']!r}", file=sys.stderr)
            continue
        if amt < 0:
            # Поступление с отрицательной суммой — скорее всего ошибка данных; логируем
            _warn["суммы_банк_отрицательные"] += 1
            print(f"[Q2] отрицательное поступление: {row['Сумма']!r} — пропускаем", file=sys.stderr)
            continue
        unidentified_sum += amt

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

        # Валюта: сравнение с минимальной ценой в прайсе осмысленно только для ₽
        if not _is_rub(row.get("Валюта")):
            _warn["crm_не_рубли"] += 1
            print(f"[Q3] не-рублёвая строка: {row.get('Валюта')!r} — пропускаем", file=sys.stderr)
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

        qty_raw = row["Кол-во"].strip()
        try:
            qty = int(qty_raw)
        except ValueError:
            try:
                qty = int(float(qty_raw.replace(",", ".")))
            except ValueError:
                _warn["qty_не_разобрано"] += 1
                print(f"[Q3] не удалось разобрать qty: {qty_raw!r} — считаем 1", file=sys.stderr)
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
        _warn["Q3_строк_без_прайса"] += len(skipped_no_price)
        unique_skipped = sorted(set(skipped_no_price))
        print(
            f"[Q3] пропущено {len(skipped_no_price)} строк (нет записи в прайсе): {unique_skipped}",
            file=sys.stderr,
        )

    # Детерминированный порядок: сортировка по убыванию числа нарушений, затем по имени
    sorted_by_manager = dict(
        sorted(violations_by_manager.items(), key=lambda kv: (-kv[1], kv[0]))
    )
    return {
        "q3_price_violation_lines_total": violations_total,
        "q3_price_violation_lines_by_manager": sorted_by_manager,
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

            # Склеиваем многострочные сообщения: строки без префикса [DD.MM.YYYY HH:MM]
            # присоединяются к предыдущему сообщению через пробел.
            messages = []  # список (ts_str, role, text)
            msg_header_re = re.compile(r"^\[(.+?)\] (Менеджер|Клиент): (.*)$")
            for raw in lines[1:]:
                m = msg_header_re.match(raw)
                if m:
                    messages.append([m.group(1), m.group(2), m.group(3)])
                elif messages and raw.strip():
                    messages[-1][2] += " " + raw.strip()

            first_price_ts = first_price_quote = None
            first_invite_ts = first_invite_quote = None

            for ts_str, role, text in messages:
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
# Сводка предупреждений
# ---------------------------------------------------------------------------

def print_warn_summary(result):
    """Выводит в stderr итоговую сводку красных флагов после расчёта всех вопросов."""
    labels = {
        "даты_не_разобраны":       "Дат не разобрано         ",
        "цены_crm_не_разобраны":   "Цен CRM не разобрано     ",
        "суммы_банк_не_разобраны": "Сумм банка не разобрано  ",
        "суммы_банк_отрицательные":"Отриц. поступлений       ",
        "скидки_не_разобраны":     "Скидок не разобрано      ",
        "менеджеры_не_опознаны":   "Менеджеров не опознано   ",
        "Q3_строк_без_прайса":     "Строк Q3 без прайса      ",
        "qty_ноль_или_отриц":      "Строк с qty≤0            ",
        "qty_не_разобрано":        "qty не разобрано         ",
        "crm_не_рубли":            "CRM строк не в ₽         ",
        "банк_не_рубли":           "Банк строк не в ₽        ",
    }
    total = sum(_warn.values())
    sep = "─" * 44
    print(f"\n{'═' * 44}", file=sys.stderr)
    print("СВОДКА  —  потенциальные проблемы с данными", file=sys.stderr)
    print(sep, file=sys.stderr)
    for key, label in labels.items():
        val = _warn.get(key, 0)
        flag = " ⚠" if val > 0 else ""
        print(f"  {label}: {val}{flag}", file=sys.stderr)
    print(sep, file=sys.stderr)
    print(f"  Итого строк Q1:  {result['q1_unique_order_lines']} уникальных строк заказов", file=sys.stderr)
    print(f"  Итого Q4 чатов:  {len(result['q4_chats_meeting_before_price'])} с приглашением до цены", file=sys.stderr)
    print(sep, file=sys.stderr)
    if total == 0:
        print("  ✓  Красных флагов нет — данные обработаны полностью", file=sys.stderr)
    else:
        print(f"  ⚠  Красных флагов: {total} — проверьте форматы входных данных!", file=sys.stderr)
    print(f"{'═' * 44}\n", file=sys.stderr)


# ---------------------------------------------------------------------------
# Поле notes — принятые решения
# ---------------------------------------------------------------------------

def build_notes():
    """Текст принятых правил расчёта — без привязки к конкретным числам из учебного набора,
    чтобы на новой версии данных notes оставались корректными."""
    return (
        "Принятые правила расчёта: "
        "(1) Даты в формате DD/MM/YYYY (не MM/DD) — определено по дням >12 в поле. "
        "(2) Excel-serial даты: база 1899-12-30 + дней. "
        "(3) Дубли (номер заказа, № строки) схлопываются по полю «Выгружено»: "
        "оставляем строку с самым поздним timestamp; при равном Выгружено берём первую "
        "(в тестовых данных такие строки — полные дубли). "
        "(4) Нормализация статусов: PAID/Оплачен/оплачено → Оплачен; "
        "Отменён/отменен/CANCELLED → Отменён; REFUND/Возврат → Возврат; остальное → В работе. "
        "(5) Скидка «N%» — процент от (цена × кол-во); число без % — фиксированная сумма "
        "в рублях на строку, делится на qty для получения скидки на единицу. "
        "(6) Привязка платежа к заказу: по номеру заказа в назначении (A-NNN / заказ NNN / "
        "зак. NNN / счёт №NNN / №NNN); «Счёт №» трактуется как номер заказа. "
        "(7) Поступления = Тип «Поступление» в ₽; возвраты (Тип «Списание») и не-рублёвые "
        "строки не учитываются. "
        "(8) Артикулы в Прайс_история хранятся под НОВЫМ именем — старый резолвится через "
        "лист «Переименования» перед поиском цены. "
        "(9) Q3: строка без подходящей записи в прайсе на дату создания — пропускается "
        "(логируется в stderr). "
        "(10) Q4: детектор приглашений не считает приглашением фразы о работе ателье "
        "(«шьём в нашем ателье») и адресе шоурума («находится у метро/по адресу»). "
        "(11) Q2b: платёж с номером, которого нет в CRM, считается неопознанным "
        "(по правилу: без привязки к существующему заказу). "
        "(12) Валюта ≠ ₽ в CRM/банке — строка пропускается с warn в stderr."
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

    print_warn_summary(result)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    if len(sys.argv) != 2:
        print("Использование: python solve.py data/", file=sys.stderr)
        sys.exit(1)
    main(sys.argv[1])
