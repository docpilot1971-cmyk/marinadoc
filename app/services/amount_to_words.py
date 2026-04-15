from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP

ONES_MALE = [
    "",
    "один",
    "два",
    "три",
    "четыре",
    "пять",
    "шесть",
    "семь",
    "восемь",
    "девять",
]
ONES_FEMALE = [
    "",
    "одна",
    "две",
    "три",
    "четыре",
    "пять",
    "шесть",
    "семь",
    "восемь",
    "девять",
]
TEENS = [
    "десять",
    "одиннадцать",
    "двенадцать",
    "тринадцать",
    "четырнадцать",
    "пятнадцать",
    "шестнадцать",
    "семнадцать",
    "восемнадцать",
    "девятнадцать",
]
TENS = [
    "",
    "",
    "двадцать",
    "тридцать",
    "сорок",
    "пятьдесят",
    "шестьдесят",
    "семьдесят",
    "восемьдесят",
    "девяносто",
]
HUNDREDS = [
    "",
    "сто",
    "двести",
    "триста",
    "четыреста",
    "пятьсот",
    "шестьсот",
    "семьсот",
    "восемьсот",
    "девятьсот",
]


def _plural_form(value: int, one: str, few: str, many: str) -> str:
    mod10 = value % 10
    mod100 = value % 100
    if mod10 == 1 and mod100 != 11:
        return one
    if 2 <= mod10 <= 4 and not 12 <= mod100 <= 14:
        return few
    return many


def _triplet_to_words(num: int, feminine: bool = False) -> str:
    words: list[str] = []
    words.append(HUNDREDS[num // 100])
    rem = num % 100
    if 10 <= rem <= 19:
        words.append(TEENS[rem - 10])
    else:
        words.append(TENS[rem // 10])
        ones_list = ONES_FEMALE if feminine else ONES_MALE
        words.append(ones_list[rem % 10])
    return " ".join(w for w in words if w)


def amount_to_words_ru(value: Decimal) -> str:
    quantized = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    rubles = int(quantized)
    kopeks = int((quantized - rubles) * 100)

    if rubles == 0:
        rubles_words = "ноль"
    else:
        rubles_words_parts: list[str] = []
        billions = rubles // 1_000_000_000
        millions = (rubles // 1_000_000) % 1000
        thousands = (rubles // 1000) % 1000
        units = rubles % 1000

        if billions:
            rubles_words_parts.append(_triplet_to_words(billions))
            rubles_words_parts.append(_plural_form(billions, "миллиард", "миллиарда", "миллиардов"))
        if millions:
            rubles_words_parts.append(_triplet_to_words(millions))
            rubles_words_parts.append(_plural_form(millions, "миллион", "миллиона", "миллионов"))
        if thousands:
            rubles_words_parts.append(_triplet_to_words(thousands, feminine=True))
            rubles_words_parts.append(_plural_form(thousands, "тысяча", "тысячи", "тысяч"))
        if units:
            rubles_words_parts.append(_triplet_to_words(units))

        rubles_words = " ".join(rubles_words_parts)

    rubles_unit = _plural_form(rubles, "рубль", "рубля", "рублей")
    kopeks_unit = _plural_form(kopeks, "копейка", "копейки", "копеек")
    return f"{rubles_words} {rubles_unit} {kopeks:02d} {kopeks_unit}".strip().capitalize()
