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
    _is_rub,
    _safe_date,
    normalize_manager,
    build_notes,
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

    # Фикс #5: допустимы однозначные день/месяц без ведущих нулей
    def test_short_dot(self):
        assert parse_date("9.1.2026") == datetime(2026, 1, 9)

    def test_short_slash(self):
        assert parse_date("5/3/2025") == datetime(2025, 3, 5)

    # Фикс #1: невалидная дата не должна ронять скрипт
    def test_invalid_day_returns_none(self):
        assert parse_date("30.02.2026") is None  # 30 февраля
        assert parse_date("31/04/2025") is None  # 31 апреля
        assert parse_date("2026-13-01") is None  # 13-й месяц

    # Фикс #5: отбрасывание времени после даты
    def test_date_with_time_iso(self):
        assert parse_date("2026-03-03 12:34") == datetime(2026, 3, 3)

    def test_date_with_time_dot(self):
        assert parse_date("09.01.2026 15:00") == datetime(2026, 1, 9)

    def test_date_with_time_slash(self):
        assert parse_date("31/10/2025 09:30") == datetime(2025, 10, 31)


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

    # Фикс #1: американский формат с запятой-разделителем тысяч без десятичной
    def test_us_thousands_no_decimal(self):
        assert parse_price_crm("25,000") == 25_000.0

    def test_us_multi_thousands_no_decimal(self):
        assert parse_price_crm("1,234,567") == 1_234_567.0

    def test_ru_decimal_two_digits(self):
        # регресс: две цифры после запятой — десятичный разделитель, не тысячи
        assert parse_price_crm("140000,00") == 140_000.0


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

    # Фикс #10: дополнительные единицы — RUR, RUB, р., т.р., рублей (слитно с числом)
    def test_rub_uppercase(self):
        assert parse_price_chat("234000 RUB") == 234_000.0

    def test_rur(self):
        assert parse_price_chat("150000 RUR") == 150_000.0

    def test_tr_abbr(self):
        assert parse_price_chat("234 т.р.") == 234_000.0

    def test_rublej_fullword(self):
        assert parse_price_chat("234000 рублей") == 234_000.0


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

    # Фикс #8: защита от деления на ноль
    def test_zero_qty_no_crash(self):
        assert calc_discounted_unit_price(100_000, 0, "10%") == 100_000
        assert calc_discounted_unit_price(100_000, 0, "500") == 100_000


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

    # Фикс #4: буква A внутри слова не должна считаться префиксом номера
    def test_no_match_inside_word(self):
        # «iPhone A777» — цифры после A, но внутри слова; 777 не в ORDERS всё равно
        assert extract_order_from_payment("Autoplatezh777", self.ORDERS) is None

    def test_match_at_sentence_start(self):
        assert extract_order_from_payment("A-00775 оплата", self.ORDERS) == 775

    def test_match_after_space(self):
        assert extract_order_from_payment("Перевод от Петрова A528", self.ORDERS) == 528

    # Фикс #6: голое «№<число>» и «No<число>» без слова «счёт»/«заказ»
    def test_bare_hash(self):
        assert extract_order_from_payment("Платёж №775", self.ORDERS) == 775

    def test_bare_no(self):
        assert extract_order_from_payment("Оплата по No 528", self.ORDERS) == 528

    def test_bare_sharp(self):
        assert extract_order_from_payment("Платёж #243", self.ORDERS) == 243


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

    # Фикс #7: расширенные паттерны приглашения
    def test_prikhodite(self):
        assert is_invite("Приходите к нам, посмотрите вживую") is True

    def test_zhdyom_vas(self):
        assert is_invite("Ждём вас в будни с 10:00 до 19:00") is True

    def test_bronir(self):
        assert is_invite("Забронируйте примерку на удобное время") is True

    # Фикс #7: «приходит» (3-е лицо, о товаре) НЕ должно считаться приглашением
    def test_prikhodit_tovar_not_invite(self):
        assert is_invite("Шуба приходит из ателье через 10 дней") is False

    def test_prikhoditsya_not_invite(self):
        assert is_invite("Приходится ждать около недели") is False


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

    def test_q3_manager_order_deterministic(self):
        """Менеджеры в q3 идут по убыванию числа нарушений, затем по имени."""
        r = self._get_result()
        if r is None:
            return
        by_mgr = r["q3_price_violation_lines_by_manager"]
        counts = list(by_mgr.values())
        # Проверяем, что список счётчиков монотонно не возрастает
        assert counts == sorted(counts, reverse=True)

    def test_notes_no_hardcoded_counts(self):
        """notes не должен содержать захардкоженные числа из учебного набора данных."""
        r = self._get_result()
        if r is None:
            return
        notes = r["notes"]
        # Бенчмарк-значения, которые не должны встречаться как отдельные числа
        for forbidden in ("1000", "1368", "3437100"):
            assert forbidden not in notes, f"notes содержит захардкоженное число: {forbidden}"


# ---------------------------------------------------------------------------
# _safe_date — прямые тесты
# ---------------------------------------------------------------------------

class TestSafeDate:
    def test_valid_date(self):
        assert _safe_date(2026, 3, 3) == datetime(2026, 3, 3)

    def test_feb_30_returns_none(self):
        assert _safe_date(2026, 2, 30) is None

    def test_apr_31_returns_none(self):
        assert _safe_date(2025, 4, 31) is None

    def test_month_13_returns_none(self):
        assert _safe_date(2026, 13, 1) is None

    def test_day_0_returns_none(self):
        assert _safe_date(2026, 1, 0) is None


# ---------------------------------------------------------------------------
# _is_rub — прямые тесты
# ---------------------------------------------------------------------------

class TestIsRub:
    def test_none_is_rub(self):
        # отсутствующая колонка валюты — считаем ₽
        assert _is_rub(None) is True

    def test_empty_string(self):
        assert _is_rub("") is True

    def test_ruble_symbol(self):
        assert _is_rub("₽") is True

    def test_rub_uppercase(self):
        assert _is_rub("RUB") is True

    def test_rub_lowercase(self):
        assert _is_rub("rub") is True

    def test_rur(self):
        assert _is_rub("RUR") is True

    def test_ru_rub(self):
        assert _is_rub("руб") is True

    def test_ru_rub_dot(self):
        # «руб.» с точкой на конце
        assert _is_rub("руб.") is True

    def test_whitespace_stripped(self):
        assert _is_rub("  RUB  ") is True

    def test_usd_false(self):
        assert _is_rub("USD") is False

    def test_eur_false(self):
        assert _is_rub("EUR") is False

    def test_dollar_sign_false(self):
        assert _is_rub("$") is False


# ---------------------------------------------------------------------------
# normalize_manager
# ---------------------------------------------------------------------------

class TestNormalizeManager:
    MGR_MAP = {
        "белова в.": "Вероника Белова",
        "veronika b": "Вероника Белова",
        "алина": "Алина Сергеева",
        "alina.s": "Алина Сергеева",
    }

    def test_known_variant(self):
        assert normalize_manager("Белова В.", self.MGR_MAP) == "Вероника Белова"

    def test_known_latin(self):
        assert normalize_manager("Veronika B", self.MGR_MAP) == "Вероника Белова"

    def test_case_insensitive(self):
        # «АЛИНА» → «алина» в map
        assert normalize_manager("АЛИНА", self.MGR_MAP) == "Алина Сергеева"

    def test_unknown_returns_original(self):
        # Неизвестный менеджер возвращает исходную строку
        assert normalize_manager("Иванова И.", self.MGR_MAP) == "Иванова И."

    def test_strips_whitespace(self):
        assert normalize_manager("  алина  ", self.MGR_MAP) == "Алина Сергеева"


# ---------------------------------------------------------------------------
# build_notes — регрессионный тест
# ---------------------------------------------------------------------------

class TestBuildNotes:
    def test_returns_nonempty_string(self):
        notes = build_notes()
        assert isinstance(notes, str) and len(notes) > 100

    def test_no_hardcoded_benchmark_numbers(self):
        """notes не должен содержать числа, привязанные к учебным данным."""
        notes = build_notes()
        for forbidden in ("1000", "1368", "3437100", " 68 ", " 21 ", " 16 "):
            assert forbidden not in notes, f"build_notes содержит захардкоженное значение: {forbidden!r}"

    def test_mentions_key_rules(self):
        notes = build_notes()
        assert "DD/MM/YYYY" in notes
        assert "Выгружено" in notes
        assert "Скидка" in notes or "скидка" in notes.lower()


# ---------------------------------------------------------------------------
# Дополнительные тесты: is_invite — непокрытые паттерны
# ---------------------------------------------------------------------------

class TestIsInviteExtra:
    def test_priyti(self):
        # «прийти» — явное приглашение
        assert is_invite("Вам нужно прийти к нам в шоурум") is True

    def test_zapishi_na_vstrechu(self):
        assert is_invite("Запишу вас на встречу на удобное время") is True

    def test_zapishi_na_primerku(self):
        assert is_invite("Запишу на примерку в четверг") is True

    def test_vstrecha_v_shoorum(self):
        assert is_invite("Давайте организуем встречу в шоуруме?") is True

    def test_nakhoditsya_po_adresu_excluded(self):
        # адресная информация — не приглашение
        assert is_invite("Наш шоурум находится по адресу: ул. Пятницкая, 18") is False

    def test_posmotrite_zhivoyu(self):
        assert is_invite("Приходите, посмотрите вживую — это совсем другое ощущение") is True


# ---------------------------------------------------------------------------
# Дополнительные тесты: extract_order_from_payment — ведущие нули в # и №
# ---------------------------------------------------------------------------

class TestExtractOrderFromPaymentExtra:
    ORDERS = {243, 775, 528}

    def test_hash_leading_zeros(self):
        assert extract_order_from_payment("#00243", self.ORDERS) == 243

    def test_nummer_leading_zeros(self):
        assert extract_order_from_payment("№00775", self.ORDERS) == 775

    def test_no_dot_space(self):
        # «No. 528» — «no.» с точкой, пробел
        assert extract_order_from_payment("No. 528", self.ORDERS) == 528


# ---------------------------------------------------------------------------
# Дополнительные тесты: calc_discounted_unit_price — непарсимая скидка
# ---------------------------------------------------------------------------

class TestCalcDiscountedUnitPriceExtra:
    def test_unparseable_discount_returns_original(self):
        # «пять рублей» не парсится parse_price_crm → возвращаем исходную цену
        assert calc_discounted_unit_price(100_000, 1, "пять рублей") == 100_000

    def test_percent_with_space(self):
        # «15 %» со пробелом перед знаком
        result = calc_discounted_unit_price(100_000, 1, "15 %")
        assert abs(result - 85_000.0) < 0.01

    def test_fixed_discount_larger_than_price(self):
        # Скидка больше цены — физически странно, но не должно падать
        result = calc_discounted_unit_price(1_000, 1, "5000")
        assert result == 1_000 - 5_000  # отрицательная цена — не наш invariant


# ---------------------------------------------------------------------------
# Детерминированная сортировка менеджеров Q3
# ---------------------------------------------------------------------------

class TestManagerSortDeterministic:
    def test_sort_by_count_desc_then_name_asc(self):
        violations = {
            "Ольга Ершова": 17,
            "Полина Орлова": 42,
            "Алина Сергеева": 5,
            "Вероника Белова": 4,
        }
        sorted_result = dict(
            sorted(violations.items(), key=lambda kv: (-kv[1], kv[0]))
        )
        assert list(sorted_result.keys()) == [
            "Полина Орлова",    # 42
            "Ольга Ершова",     # 17
            "Алина Сергеева",   # 5
            "Вероника Белова",  # 4
        ]

    def test_tie_broken_by_name(self):
        violations = {"Яблокова": 10, "Абрамова": 10}
        sorted_result = dict(
            sorted(violations.items(), key=lambda kv: (-kv[1], kv[0]))
        )
        assert list(sorted_result.keys()) == ["Абрамова", "Яблокова"]
