from __future__ import annotations

import re
from datetime import date
from decimal import Decimal, InvalidOperation

_MONTHS_RU = {
    "января": 1,
    "февраля": 2,
    "марта": 3,
    "апреля": 4,
    "мая": 5,
    "июня": 6,
    "июля": 7,
    "августа": 8,
    "сентября": 9,
    "октября": 10,
    "ноября": 11,
    "декабря": 12,
}


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def try_parse_date(value: str) -> date | None:
    value = normalize_text(value.replace("«", "").replace("»", ""))
    if not value:
        return None

    match_numeric = re.search(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})", value)
    if match_numeric:
        day, month, year = match_numeric.groups()
        try:
            return date(int(year), int(month), int(day))
        except ValueError:
            return None

    match_ru = re.search(
        r"(\d{1,2})\s+([А-Яа-яЁё]+)\s+(\d{4})",
        value,
        flags=re.IGNORECASE,
    )
    if match_ru:
        day_text, month_text, year_text = match_ru.groups()
        month = _MONTHS_RU.get(month_text.lower())
        if month is None:
            return None
        try:
            return date(int(year_text), month, int(day_text))
        except ValueError:
            return None

    return None


def parse_decimal(value: str | None) -> Decimal | None:
    if value is None:
        return None
    
    # Extract the first number from the text (before any parentheses or text)
    # This handles cases like "319 000 (Триста девятнадцать тысяч...)"
    match = re.match(r"\s*([0-9][0-9\s.,]*)", value)
    if match:
        number_str = match.group(1)
    else:
        # Fallback: try to find any number in the text
        match = re.search(r"([0-9][0-9\s.,]*)", value)
        if not match:
            return None
        number_str = match.group(1)
    
    cleaned = number_str.replace("\xa0", "").replace(" ", "").replace(",", ".")
    cleaned = re.sub(r"[^\d.\-]", "", cleaned)
    if not cleaned or cleaned == "-":
        return None
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def find_first(pattern: str, text: str, flags: int = re.IGNORECASE) -> str | None:
    match = re.search(pattern, text, flags=flags)
    if not match:
        return None
    return normalize_text(match.group(1))


def extract_between(text: str, start_keywords: tuple[str, ...], end_keywords: tuple[str, ...]) -> str:
    lower = text.lower()
    start_positions = [lower.find(k.lower()) for k in start_keywords if lower.find(k.lower()) >= 0]
    if not start_positions:
        return ""

    start = min(start_positions)
    end = len(text)
    for keyword in end_keywords:
        idx = lower.find(keyword.lower(), start + 1)
        if idx >= 0:
            end = min(end, idx)
    return text[start:end]


def is_numeric_like(value: str) -> bool:
    return parse_decimal(value) is not None
