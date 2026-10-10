"""Árfolyamhiányok keresése, korábbi értékkel pótlása és kézi megadása."""

from datetime import date
from decimal import Decimal

from .import_state import PreparedImport
from .models import DailyExchangeRate, Movement


def store_exchange_rate(
    exchange_rates: dict[date, DailyExchangeRate],
    record: DailyExchangeRate,
) -> None:
    """Ellenőrzött rekord hozzáadása a már alkalmazott értékek megőrzésével."""
    day = record.application_date
    existing = exchange_rates.get(day)
    if existing is not None and existing != record:
        raise ValueError(f"Eltérő tárolt árfolyamrekord: {day:%Y-%m-%d}")
    for known in exchange_rates.values():
        if known.source_date == record.source_date and known.rate != record.rate:
            raise ValueError(f"Eltérő forrásárfolyam: {record.source_date:%Y-%m-%d}")
    exchange_rates[day] = record


def find_missing_exchange_rate_dates(
    new_movements: list[Movement],
    exchange_rates: dict[date, DailyExchangeRate],
) -> list[date]:
    """Az új eurós tételekhez hiányzó alkalmazási dátumok."""
    required_dates = {
        movement.movement_date
        for movement in new_movements
        if movement.currency == "EUR"
    }

    return sorted(required_dates - set(exchange_rates))


def fill_missing_exchange_rates(
    missing_dates: list[date],
    exchange_rates: dict[date, DailyExchangeRate],
) -> dict[date, DailyExchangeRate]:
    """Hiányzó napok kiegészítése korábbi ismert forrásárfolyammal."""
    completed_rates = exchange_rates.copy()
    source_rates: dict[date, Decimal] = {}

    for record in exchange_rates.values():
        source_date = record.source_date

        if source_date in source_rates and source_rates[source_date] != record.rate:
            raise ValueError(
                f"Eltérő árfolyamok ugyanahhoz a forrásdátumhoz: "
                f"{source_date:%Y-%m-%d}"
            )

        source_rates[source_date] = record.rate

    for day in missing_dates:
        previous_dates = [
            source_date for source_date in source_rates if source_date <= day
        ]

        if not previous_dates:
            continue

        latest_source_date = max(previous_dates)

        completed_rates[day] = DailyExchangeRate(
            application_date=day,
            source_date=latest_source_date,
            rate=source_rates[latest_source_date],
        )

    return completed_rates


def get_processor_exchange_rates(
    exchange_rates: dict[date, DailyExchangeRate],
) -> dict[date, Decimal]:
    """Az alkalmazási dátumokhoz tartozó árfolyamértékek kigyűjtése."""
    return {day: record.rate for day, record in exchange_rates.items()}


def add_source_exchange_rate(
    prepared: PreparedImport,
    source_date: date,
    rate: Decimal,
) -> None:
    """Lekért vagy kézzel megadott tényleges forrásárfolyam hozzáadása."""
    store_exchange_rate(
        prepared.exchange_rates,
        DailyExchangeRate(source_date, source_date, rate),
    )


def add_manual_exchange_rate(
    prepared: PreparedImport,
    application_date: date,
    source_date: date,
    rate: Decimal,
) -> None:
    """Hiányzó alkalmazási nap árfolyamának kézi megadása."""
    store_exchange_rate(
        prepared.exchange_rates,
        DailyExchangeRate(application_date, source_date, rate),
    )
