"""
Юнит-тесты для solve.py — проверяют парсеры и нормализаторы изолированно от данных.
Запуск: python -m pytest tests/ -v
"""

import sys
import os
import subprocess
import json
from datetime import datetime

# Добавляем корень проекта в path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from solve import (
    parse_date,
    normalize_order_num,
    normalize_status,
    parse_price_crm,
    parse_price_chat,
    _parse_written_rubles,
    calc_discounted_unit_price,
    extract_order_from_payment,
    is_invite,
)


# ---------------------------------------------------------------------------
# parse_date — все 5 форматов
# ---------------------------------------------------------------------------

class TestParseDate:
    def test_iso(self):
        assert parse_date("2026-03-03") == datetime(2026, 3, 3)

    def test_dot(self):
        assert parse_date("09.01.2026") == datetime(2026, 1, 9)

    def test_slash_dd_mm(self):
        # подтверждён данными: первая часть — день
        assert parse_date("31/10/2025") == datetime(2025, 10, 31)

    def test_excel_serial(self):
        # 45752 = 2025-03-31 (проверяем формулой 1899-12-30 + 45752 дней)
        result = parse_date("45752")
        assert result is not None
        assert result == datetime(1899, 12, 30) + __import__("datetime").timedelta(days=45752)

    def test_russian_text(self):
        assert parse_date("23 сен 2026") == datetime(2026, 9, 23)

    def test_russian_text_may(self):
        # «мая» — особый падеж
        assert parse_date("15 мая 2025") == datetime(2025, 5, 15)

    def test_invalid_returns_none(self):
        assert parse_date("неизвестно") is None

    def test_empty_returns_none(self):
        assert parse_date("") is None


# ---------------------------------------------------------------------------
# normalize_order_num
# ---------------------------------------------------------------------------

class TestNormalizeOrderNum:
    def test_padded(self):
        assert normalize_order_num("A-00075") == 75

    def test_cyrillic_a(self):
        assert normalize_order_num("А-00075") == 75

    def test_lower_short(self):
        assert normalize_order_num("a-75") == 75

    def test_no_dash(self):
        assert normalize_order_num("A75") == 75

    def test_space_prefix(self):
        assert normalize_order_num(" a-75") == 75


# ---------------------------------------------------------------------------
# normalize_status
# ---------------------------------------------------------------------------

class TestNormalizeStatus:
    def test_paid_variants(self):
        assert normalize_status("Оплачен") == "Оплачен"
        assert normalize_status("Оплачен ") == "Оплачен"
        assert normalize_status("оплачено") == "Оплачен"
        assert normalize_status("PAID") == "Оплачен"

    def test_cancelled_variants(self):
        assert normalize_status("Отменён") == "Отменён"
        assert normalize_status("отменен") == "Отменён"
        assert normalize_status("CANCELLED") == "Отменён"

    def test_refund_variants(self):
        assert normalize_status("Возврат") == "Возврат"
        assert normalize_status("возврат") == "Возврат"
        assert normalize_status("REFUND") == "Возврат"

    def test_in_progress_fallback(self):
        assert normalize_status("В работе") == "В работе"
        assert normalize_status("NEW") == "В работе"
        assert normalize_status("Ждём оплату") == "В работе"
        assert normalize_status("неизвестный статус") == "В работе"


# ---------------------------------------------------------------------------
# parse_price_crm
# ---------------------------------------------------------------------------

class TestParsePriceCrm:
    def test_space_thousands(self):
        assert parse_price_crm("50 000") == 50_000.0

    def test_space_millions(self):
        assert parse_price_crm("2 200 000") == 2_200_000.0

    def test_comma_dot(self):
        assert parse_price_crm("17,000.00") == 17_000.0

    def test_comma_decimal(self):
        assert parse_price_crm("140000,00") == 140_000.0

    def test_mln_suffix(self):
        assert parse_price_crm("2,48 млн") == 2_480_000.0

    def test_dot_thousands(self):
        assert parse_price_crm("385.000") == 385_000.0

    def test_plain_int(self):
        assert parse_price_crm("252000") == 252_000.0


# ---------------------------------------------------------------------------
# parse_price_chat / _parse_written_rubles
# ---------------------------------------------------------------------------

class TestParsePriceChat:
    def test_ruble_symbol(self):
        assert parse_price_chat("53 000 ₽") == 53_000.0

    def test_tys(self):
        assert parse_price_chat("138 тыс.") == 138_000.0

    def test_mln(self):
        assert parse_price_chat("2,404 млн") == 2_404_000.0

    def test_written_hundreds(self):
        assert _parse_written_rubles("сто пятьдесят восемь тысяч рублей") == 158_000.0

    def test_written_millions(self):
        assert _parse_written_rubles("два миллиона двести тысяч рублей") == 2_200_000.0

    def test_time_not_parsed_as_price(self):
        # «15:00» не должно считаться ценой
        assert parse_price_chat("давайте в 15:00?") is None

    def test_dot_thousands_in_chat(self):
        assert parse_price_chat("2.094.000 руб") == 2_094_000.0


# ---------------------------------------------------------------------------
# calc_discounted_unit_price
# ---------------------------------------------------------------------------

class TestCalcDiscountedUnitPrice:
    def test_percent_discount(self):
        # 10% от (252000 × 1) → итого 226800 за единицу
        result = calc_discounted_unit_price(252_000, 1, "10%")
        assert abs(result - 226_800.0) < 0.01

    def test_fixed_rub_discount(self):
        # скидка 1000 руб на строку из 2 единиц → -500 на единицу
        result = calc_discounted_unit_price(5_000, 2, "1000")
        assert abs(result - 4_500.0) < 0.01

    def test_no_discount(self):
        assert calc_discounted_unit_price(100_000, 1, "") == 100_000

    def test_formatted_fixed_discount(self):
        # «24,700.00» — рублей на строку
        result = calc_discounted_unit_price(50_000, 1, "24,700.00")
        assert abs(result - (50_000 - 24_700)) < 0.01


# ---------------------------------------------------------------------------
# extract_order_from_payment
# ---------------------------------------------------------------------------

class TestExtractOrderFromPayment:
    ORDERS = {877, 775, 528, 243, 391, 828, 991}

    def test_a_prefix(self):
        assert extract_order_from_payment("A877 доплата", self.ORDERS) == 877

    def test_zakaz_word(self):
        assert extract_order_from_payment("Предоплата заказ 775, тел. ...", self.ORDERS) == 775

    def test_zak_abbr(self):
        assert extract_order_from_payment("зак. 528", self.ORDERS) == 528

    def test_a_padded(self):
        assert extract_order_from_payment("Оплата заказа A-00243", self.ORDERS) == 243

    def test_cyrillic_a(self):
        assert extract_order_from_payment("Оплата по заказу А-00391 (шуба)", self.ORDERS) == 391

    def test_schet(self):
        assert extract_order_from_payment("по счету №828 от ...", self.ORDERS) == 828

    def test_schet_padded(self):
        assert extract_order_from_payment("Счёт № 00991 от ...", self.ORDERS) == 991

    def test_no_order(self):
        assert extract_order_from_payment("Оплата по договору", self.ORDERS) is None

    def test_unknown_order_returns_none(self):
        assert extract_order_from_payment("A-01169", self.ORDERS) is None


# ---------------------------------------------------------------------------
# is_invite
# ---------------------------------------------------------------------------

class TestIsInvite:
    def test_primerka(self):
        assert is_invite("Лучше приезжайте к нам в ателье на примерку!") is True

    def test_priezzhajte(self):
        assert is_invite("Приезжайте, будем рады видеть вас!") is True

    def test_shyom_v_atelie_excluded(self):
        assert is_invite("Все изделия шьём в нашем ателье индивидуально.") is False

    def test_metro_excluded(self):
        assert is_invite("Наш шоурум находится у метро Арбатская.") is False

    def test_price_message_not_invite(self):
        assert is_invite("Цена изделия 234 000 ₽.") is False


# ---------------------------------------------------------------------------
# Интеграционный тест: итоговые значения Q1–Q4
# ---------------------------------------------------------------------------

class TestIntegration:
    """Запускает solve.py и проверяет эталонные значения Q1–Q4."""

    _result = None

    @classmethod
    def _get_result(cls):
        if cls._result is None:
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
            if not os.path.isdir(data_dir):
                return None
            proc = subprocess.run(
                [sys.executable, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "solve.py"), data_dir],
                capture_output=True, text=True, encoding="utf-8",
            )
            cls._result = json.loads(proc.stdout)
        return cls._result

    def test_q1_unique_orders(self):
        r = self._get_result()
        if r is None:
            return  # data/ недоступна — пропустить
        assert r["q1_unique_orders"] == 1000

    def test_q1_unique_lines(self):
        r = self._get_result()
        if r is None:
            return
        assert r["q1_unique_order_lines"] == 1368

    def test_q2a_paid_without_payment(self):
        r = self._get_result()
        if r is None:
            return
        assert sorted(r["q2a_paid_orders_without_payment"]) == [12, 59, 84, 221, 600, 805, 827, 857, 993]

    def test_q2b_count(self):
        r = self._get_result()
        if r is None:
            return
        assert r["q2b_unidentified_payments_count"] == 16

    def test_q2b_sum(self):
        r = self._get_result()
        if r is None:
            return
        assert abs(r["q2b_unidentified_payments_sum_rub"] - 3_437_100.0) < 1.0

    def test_q3_total(self):
        r = self._get_result()
        if r is None:
            return
        assert r["q3_price_violation_lines_total"] == 68

    def test_q4_count(self):
        r = self._get_result()
        if r is None:
            return
        assert len(r["q4_chats_meeting_before_price"]) == 21
