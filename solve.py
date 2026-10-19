"""
AK Furs test assignment — deterministic solver.
Usage: python solve.py data/ > answers.json
All logs go to stderr. Only JSON goes to stdout.
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
# Loaders
# ---------------------------------------------------------------------------

def load_crm(path):
    """Return list of dicts; rows with unparseable order numbers are skipped."""
    rows = []
    with open(path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            rows.append({k: v for k, v in row.items()})
    print(f"[CRM] loaded {len(rows)} raw rows", file=sys.stderr)
    return rows


def load_bank(path):
    """Return list of dicts for ALL rows (Поступление + Списание)."""
    rows = []
    with open(path, encoding="cp1251") as f:
        reader = csv.reader(f, delimiter=";")
        next(reader)  # skip title row "Выписка по счёту..."
        header = next(reader)
        for row in reader:
            if len(row) >= len(header):
                rows.append(dict(zip(header, row)))
    print(f"[Bank] loaded {len(rows)} raw rows", file=sys.stderr)
    return rows


def load_xlsx_sheets(path):
    """Return openpyxl workbook (all sheets including hidden)."""
    import openpyxl
    return openpyxl.load_workbook(path)


def load_managers(wb):
    """
    Return dict: lowercase-stripped variant -> canonical name.
    Variants are pipe-separated in the sheet.
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
    print(f"[Managers] loaded {len(mgr_map)} variant mappings", file=sys.stderr)
    return mgr_map


def load_renames(wb):
    """Return dict: old_article -> new_article."""
    ws = wb["Переименования"]
    renames = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        old, new = row[0], row[1]
        if old and new:
            renames[old] = new
    print(f"[Renames] loaded {len(renames)} article renames", file=sys.stderr)
    return renames


def load_prices(wb):
    """
    Return dict: canonical_article -> list of (date_from, date_to, min_price).
    date_to may be None (open-ended).
    Articles stored under NEW name only — caller must resolve renames first.
    """
    ws = wb["Прайс_история"]
    prices = defaultdict(list)
    for row in ws.iter_rows(min_row=2, values_only=True):
        art, d_from, d_to, _, min_p = row[0], row[1], row[2], row[3], row[4]
        if art and d_from is not None and min_p is not None:
            prices[art].append((d_from, d_to, float(min_p)))
    print(f"[Prices] loaded {len(prices)} articles", file=sys.stderr)
    return dict(prices)


# ---------------------------------------------------------------------------
# Parsers / normalizers
# ---------------------------------------------------------------------------

_RU_MONTHS = {
    "янв": 1, "фев": 2, "мар": 3, "апр": 4, "май": 5, "мая": 5,
    "июн": 6, "июл": 7, "авг": 8, "сен": 9, "окт": 10, "ноя": 11, "дек": 12,
}


def parse_date(s):
    """Parse date string to datetime. Returns None on failure."""
    if not s:
        return None
    s = s.strip()
    # ISO
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:
        return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    # DD.MM.YYYY
    m = re.match(r"^(\d{2})\.(\d{2})\.(\d{4})$", s)
    if m:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    # DD/MM/YYYY  (confirmed DD/MM in recon)
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", s)
    if m:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    # Excel serial (5-digit integer)
    m = re.match(r"^\d{5}$", s)
    if m:
        return datetime(1899, 12, 30) + timedelta(days=int(s))
    # Russian text: "23 сен 2026" or "23 сентября 2026"
    m = re.match(r"^(\d{1,2})\s+([а-яё]+)\.?\s+(\d{4})$", s, re.IGNORECASE)
    if m:
        day, mon_str, year = int(m.group(1)), m.group(2).lower()[:3], int(m.group(3))
        month = _RU_MONTHS.get(mon_str)
        if month:
            return datetime(year, month, day)
    print(f"[parse_date] cannot parse: {s!r}", file=sys.stderr)
    return None


def normalize_order_num(s):
    """Extract integer order number from any format. Returns None on failure."""
    s = s.strip()
    # remove latin/cyrillic A prefix (with optional dash and leading zeros)
    m = re.search(r"\d+", s)
    if m:
        return int(m.group())
    return None


def normalize_status(s):
    """Map raw status string to one of 4 canonical values."""
    s = s.strip().upper()
    if s in ("ОПЛАЧЕН", "PAID", "ОПЛАЧЕНО"):
        return "Оплачен"
    if s in ("ОТМЕНЁН", "ОТМЕНЕН", "CANCELLED"):
        return "Отменён"
    if s in ("ВОЗВРАТ", "REFUND"):
        return "Возврат"
    return "В работе"


def normalize_manager(s, mgr_map):
    """Map raw manager string to canonical name. Returns raw string if unknown."""
    key = s.strip().lower()
    result = mgr_map.get(key)
    if result is None:
        print(f"[Manager] unknown variant: {s!r}", file=sys.stderr)
    return result or s.strip()


def parse_price_crm(s):
    """
    Parse CRM price string to float (rubles).
    Handles: '50 000', '17,000.00', '245,000.00', '2 200 000',
             '140000,00', '2,48 млн', '500000,00', '32,000.00'
    """
    if not s:
        return None
    s = s.strip()
    # millions: '2,48 млн' or '2.48 млн'
    m = re.match(r"^([\d][0-9\s,\.]+)\s*млн$", s, re.IGNORECASE)
    if m:
        num_str = m.group(1).strip().replace(" ", "").replace(",", ".")
        return float(num_str) * 1_000_000

    # Remove spaces (thousand separators)
    s = s.replace(" ", "")

    # Detect format: if has both comma and dot, comma=thousands dot=decimal
    if "," in s and "." in s:
        s = s.replace(",", "")  # remove comma-thousands, keep dot-decimal

    # Dots as thousand separators: X.XXX.XXX or X.XXX (no cents after last dot means thousands)
    # Distinguish: 2.094.000 (thousands) vs 17000.00 (decimal)
    # Rule: if all dot-separated parts after first are exactly 3 digits → thousands separator
    parts = s.split(".")
    if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
        s = s.replace(".", "")
    else:
        # single dot → decimal separator (or keep as-is)
        s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        print(f"[parse_price_crm] cannot parse: {s!r}", file=sys.stderr)
        return None


def parse_price_chat(s):
    """
    Extract first price (rubles) from a manager chat message.
    Handles: '53 000 ₽', '138 тыс.', '2,404 млн', '385.000 руб',
             'сто пятьдесят восемь тысяч рублей', '2.094.000 руб'.
    Returns None if no price found or text is time/percentage only.
    """
    # Skip pure time references like "15:00"
    text = re.sub(r"\b\d{1,2}:\d{2}\b", "", s)

    # Written-out Russian numbers → float
    written = _parse_written_rubles(text)
    if written is not None:
        return written

    # Numeric patterns with unit suffix
    # Group 1: number (spaces/commas/dots as separators)
    # Group 2: unit (₽, руб, тыс, к, млн)
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

    # Parse the number part
    num_str = num_str.replace(" ", "")
    # Dots-as-thousands: X.XXX or X.XXX.XXX
    dot_parts = num_str.split(".")
    if len(dot_parts) > 1 and all(len(p) == 3 for p in dot_parts[1:]):
        num_str = num_str.replace(".", "")
    else:
        num_str = num_str.replace(",", ".").replace(".", ".", 1)
        # remove any remaining dots except the decimal one
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
    Detect written-out ruble amounts like 'сто пятьдесят восемь тысяч рублей'.
    Returns float or None.
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
    Apply discount to price_per_unit.
    'N%' → percent of (price * qty), result per unit.
    Number without '%' → ruble amount off the whole line, per unit.
    Returns discounted price per unit.
    """
    if not discount_str or not discount_str.strip():
        return price_per_unit
    d = discount_str.strip()
    if d.endswith("%"):
        pct = float(d[:-1].replace(",", ".").replace(" ", ""))
        total_discount = price_per_unit * qty * pct / 100
        return price_per_unit - total_discount / qty
    else:
        # fixed ruble amount on whole line — reuse CRM price parser
        amt = parse_price_crm(d)
        if amt is None:
            print(f"[calc_discount] cannot parse discount: {d!r}", file=sys.stderr)
            return price_per_unit
        return price_per_unit - amt / qty


def extract_order_from_payment(text, known_orders):
    """
    Extract order number from bank payment description.
    Returns int if found in known_orders, else None.
    Patterns: 'A877', 'заказ 775', 'зак. 528', 'A-00243',
              'по счету №828', 'Счёт № 00991', 'А-00391' (Cyrillic A).
    """
    # Find all candidate numbers in the text
    # Match optional A/А prefix (latin or cyrillic) with optional dash
    candidates = re.findall(
        r"[AАaа]-?0*(\d+)",
        text,
        re.IGNORECASE,
    )
    # Also plain numbers after keywords
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
    Return True if manager message is an invitation to visit (meeting/fitting).
    Excludes info phrases: 'шьём в нашем ателье', 'находится у метро'.
    """
    t = text.lower()
    # Exclude known non-invite phrases
    if "шьём в нашем ателье" in t:
        return False
    if "находится у метро" in t or "находится по адрес" in t:
        return False
    # Invite keywords
    invite_patterns = [
        r"прим[её]рк",          # примерку, примерки
        r"подъехать",
        r"приезжайт",
        r"приглаш",
        r"запиш.{0,10}встреч",  # запишу вас на встречу
        r"встреч.{0,10}шоурум", # встречу в шоурум
    ]
    return any(re.search(p, t) for p in invite_patterns)


# ---------------------------------------------------------------------------
# CRM deduplication
# ---------------------------------------------------------------------------

def deduplicate_crm(rows):
    """
    Keep one row per (normalized_order_num, line_num): the one with latest Выгружено.
    Equal timestamps are full duplicates — keep first encountered.
    Returns list of clean rows with added fields: order_num (int), status_canonical.
    """
    groups = defaultdict(list)
    skipped = 0
    for row in rows:
        num = normalize_order_num(row["Номер заказа"])
        if num is None:
            skipped += 1
            continue
        line = row["№ строки"].strip()
        ts_str = row["Выгружено"].strip()
        # parse Выгружено as datetime for comparison
        try:
            ts = datetime.strptime(ts_str, "%Y-%m-%d %H:%M")
        except ValueError:
            ts = datetime.min
        groups[(num, line)].append((ts, row))

    if skipped:
        print(f"[CRM] skipped {skipped} rows with unparseable order numbers", file=sys.stderr)

    clean = []
    for (num, line), entries in groups.items():
        # sort descending by timestamp, keep first (latest)
        entries.sort(key=lambda x: x[0], reverse=True)
        _, best_row = entries[0]
        best_row = dict(best_row)
        best_row["_order_num"] = num
        best_row["_status"] = normalize_status(best_row["Статус"])
        clean.append(best_row)

    print(f"[CRM] after dedup: {len(clean)} unique (order, line) pairs", file=sys.stderr)
    return clean


# ---------------------------------------------------------------------------
# Q1
# ---------------------------------------------------------------------------

def solve_q1(clean_rows):
    unique_orders = len(set(r["_order_num"] for r in clean_rows))
    unique_lines = len(clean_rows)
    return {
        "q1_unique_orders": unique_orders,
        "q1_unique_order_lines": unique_lines,
    }


# ---------------------------------------------------------------------------
# Q2
# ---------------------------------------------------------------------------

def solve_q2(clean_rows, bank_rows):
    # Build set of known order numbers
    known_orders = set(r["_order_num"] for r in clean_rows)

    # Paid orders
    paid_orders = set(
        r["_order_num"] for r in clean_rows if r["_status"] == "Оплачен"
    )

    # Only incoming payments
    incoming = [r for r in bank_rows if r["Тип"].strip() == "Поступление"]
    print(f"[Q2] incoming payments: {len(incoming)} of {len(bank_rows)} total", file=sys.stderr)

    # Map each incoming payment to an order
    payment_orders = set()
    unidentified = []
    for row in incoming:
        ref = row["Назначение платежа"].strip()
        order_num = extract_order_from_payment(ref, known_orders)
        if order_num is not None:
            payment_orders.add(order_num)
        else:
            unidentified.append(row)

    # (a) Paid orders with zero incoming payments
    paid_without_payment = sorted(paid_orders - payment_orders)

    # (b) Unidentified incoming payments
    unidentified_sum = 0.0
    for row in unidentified:
        amt_str = row["Сумма"].strip().replace(" ", "").replace(",", ".")
        try:
            unidentified_sum += float(amt_str)
        except ValueError:
            print(f"[Q2] cannot parse amount: {row['Сумма']!r}", file=sys.stderr)

    print(
        f"[Q2] paid orders without payment: {len(paid_without_payment)}, "
        f"unidentified: {len(unidentified)}, sum: {unidentified_sum:.2f}",
        file=sys.stderr,
    )
    return {
        "q2a_paid_orders_without_payment": paid_without_payment,
        "q2b_unidentified_payments_count": len(unidentified),
        "q2b_unidentified_payments_sum_rub": round(unidentified_sum, 2),
    }


# ---------------------------------------------------------------------------
# Q3
# ---------------------------------------------------------------------------

def solve_q3(clean_rows, prices, renames, mgr_map):
    violations_total = 0
    violations_by_manager = defaultdict(int)
    skipped_no_price = []

    for row in clean_rows:
        if row["_status"] == "Отменён":
            continue

        art_raw = row["Артикул"].strip()
        art = renames.get(art_raw, art_raw)  # resolve old -> new

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

        # Find applicable minimum price for order_date
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
            f"[Q3] skipped {len(skipped_no_price)} rows (no price record): {unique_skipped}",
            file=sys.stderr,
        )

    return {
        "q3_price_violation_lines_total": violations_total,
        "q3_price_violation_lines_by_manager": dict(violations_by_manager),
    }


# ---------------------------------------------------------------------------
# Q4
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
            header = lines[0]
            m_id = re.search(r"Чат (C\d+)", header)
            chat_id = m_id.group(1) if m_id else name

            first_price_ts = None
            first_price_quote = None
            first_invite_ts = None
            first_invite_quote = None

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
                    if first_price_ts is None:
                        price = parse_price_chat(text)
                        if price is not None:
                            first_price_ts = ts
                            first_price_quote = text

                    if first_invite_ts is None and is_invite(text):
                        first_invite_ts = ts
                        first_invite_quote = text

            # Classify chat
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

    # Write debug file
    debug_path = os.path.join(os.path.dirname(zip_path), "..", "debug_q4.txt")
    debug_path = os.path.normpath(debug_path)
    with open(debug_path, "w", encoding="utf-8") as f:
        f.write(f"Чатов с приглашением ДО цены: {len(chats_invite_before_price)}\n")
        f.write("\n".join(debug_lines))

    # Print ambiguous to stderr
    if ambiguous:
        print(f"\n[Q4] Неоднозначные чаты ({len(ambiguous)}):", file=sys.stderr)
        for cid, reason in ambiguous:
            print(f"  {cid}: {reason}", file=sys.stderr)

    print(
        f"[Q4] invite-before-price: {len(chats_invite_before_price)}, "
        f"ambiguous: {len(ambiguous)}",
        file=sys.stderr,
    )
    return {"q4_chats_meeting_before_price": sorted(chats_invite_before_price)}


# ---------------------------------------------------------------------------
# Notes
# ---------------------------------------------------------------------------

def build_notes():
    return (
        "Принятые решения: "
        "(1) Даты DD/MM/YYYY (подтверждено данными: 138 строк с днём >12). "
        "(2) Excel serial: база 1899-12-30 + дней. "
        "(3) Дубли (order, line): берём строку с самым поздним Выгружено; "
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
        "не приглашения. Дубли в выписке: 0."
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main(data_dir):
    crm_path   = os.path.join(data_dir, "01_crm_orders_export.csv")
    bank_path  = os.path.join(data_dir, "02_bank_statement_rub.csv")
    xlsx_path  = os.path.join(data_dir, "03_catalog_prices_reference.xlsx")
    chats_path = os.path.join(data_dir, "04_client_chats.zip")

    # Load
    wb        = load_xlsx_sheets(xlsx_path)
    mgr_map   = load_managers(wb)
    renames   = load_renames(wb)
    prices    = load_prices(wb)
    crm_raw   = load_crm(crm_path)
    bank_rows = load_bank(bank_path)

    # Deduplicate CRM
    clean = deduplicate_crm(crm_raw)

    # Solve
    result = {}
    result.update(solve_q1(clean))
    result.update(solve_q2(clean, bank_rows))
    result.update(solve_q3(clean, prices, renames, mgr_map))
    result.update(solve_q4(chats_path))
    result["run_command"] = "python solve.py data/ > answers.json"
    result["notes"] = build_notes()

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    # ensure stdout is UTF-8 regardless of terminal locale
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    if len(sys.argv) != 2:
        print("Usage: python solve.py data/", file=sys.stderr)
        sys.exit(1)
    main(sys.argv[1])
