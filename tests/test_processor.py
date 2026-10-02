from datetime import date
from decimal import Decimal
from pathlib import Path

from axelpro_riport.csv_reader import read_movements
from axelpro_riport.models import DailyTotals, MonthlyTotals, Movement
from axelpro_riport.processor import process_movements, round_huf


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SAMPLES_DIR = PROJECT_ROOT / "samples"

SAMPLE_FILES = (
    "napokon_at.csv",
    "honapokon_at.csv",
    "eveken_at.csv",
)


def test_sample_totals() -> None:
    for filename in SAMPLE_FILES:
        movements = read_movements(SAMPLES_DIR / filename)
        dates = sorted({movement.movement_date for movement in movements})
        exchange_rates = {
            movement_date: Decimal("400")
            for movement_date in dates
        }
        processed_documents = set()

        result = process_movements(
            movements=movements,
            processed_documents=processed_documents,
            exchange_rates=exchange_rates,
            additional_new_part_groups=set(),
        )

        assert len(dates) == 3, filename
        assert len(result.daily_totals) == 3, filename
        assert len(result.new_documents) == 30, filename
        assert processed_documents == set(), filename

        expected_daily = []
        monthly_factors = {}

        for k, movement_date in enumerate(dates, start=1):
            expected_daily.append(
                DailyTotals(
                    movement_date=movement_date,
                    used_parts=14000 * k,
                    new_parts=37000 * k,
                    labor=22500 * k,
                    workshop_used_parts=10000 * k,
                    workshop_new_parts=33000 * k,
                    purchases=13125 * k,
                    shipping=1250 * k,
                    trailer=7500 * k,
                )
            )

            month_key = (movement_date.year, movement_date.month)
            monthly_factors[month_key] = (
                monthly_factors.get(month_key, 0) + k
            )

        expected_monthly = [
            MonthlyTotals(
                year=year,
                month=month,
                sany=500 * factor,
                diag=1000 * factor,
                kj=1500 * factor,
            )
            for (year, month), factor in sorted(monthly_factors.items())
        ]

        assert result.daily_totals == expected_daily, (
            filename,
            result.daily_totals,
        )
        assert result.monthly_totals == expected_monthly, (
            filename,
            result.monthly_totals,
        )

        repeated = process_movements(
            movements=movements,
            processed_documents=result.new_documents,
            exchange_rates={},
            additional_new_part_groups=set(),
        )

        assert repeated.daily_totals == [], filename
        assert repeated.monthly_totals == [], filename
        assert repeated.new_documents == set(), filename

        print(f"Sikeres összesítési és ismételtimport-teszt: {filename}")


def test_additional_new_part_groups() -> None:
    movements = read_movements(SAMPLES_DIR / "napokon_at.csv")
    dates = sorted({movement.movement_date for movement in movements})
    exchange_rates = {
        movement_date: Decimal("400")
        for movement_date in dates
    }

    result = process_movements(
        movements=movements,
        processed_documents=set(),
        exchange_rates=exchange_rates,
        additional_new_part_groups={"  sZűRőK  "},
    )

    assert len(result.daily_totals) == 3

    for k, totals in enumerate(result.daily_totals, start=1):
        assert totals.new_parts == 40500 * k, totals
        assert totals.workshop_new_parts == 36500 * k, totals

    print("Sikeres kiegészítőcsoport-teszt.")


def test_ignored_eur_movement() -> None:
    movement = Movement(
        movement_date=date(2026, 9, 21),
        transaction_type="Számla",
        document_number="TESZT-001",
        group="Egyéb",
        item_code="EGYEB",
        net_value=Decimal("100"),
        currency="EUR",
    )

    result = process_movements(
        movements=[movement],
        processed_documents=set(),
        exchange_rates={},
        additional_new_part_groups=set(),
    )

    expected_daily = DailyTotals(
        movement_date=movement.movement_date,
        used_parts=0,
        new_parts=0,
        labor=0,
        workshop_used_parts=0,
        workshop_new_parts=0,
        purchases=0,
        shipping=0,
        trailer=0,
    )
    expected_monthly = MonthlyTotals(
        year=2026,
        month=9,
        sany=0,
        diag=0,
        kj=0,
    )

    assert result.daily_totals == [expected_daily], result.daily_totals
    assert result.monthly_totals == [expected_monthly], result.monthly_totals
    assert result.new_documents == {(2026, "TESZT-001")}

    print("Sikeres teszt: a nem összesített EUR-tételhez nem kell árfolyam.")


def test_rounding() -> None:
    assert round_huf(Decimal("12.49")) == 12
    assert round_huf(Decimal("12.50")) == 13
    assert round_huf(Decimal("-12.49")) == -12
    assert round_huf(Decimal("-12.50")) == -13

    print("Sikeres kerekítési teszt.")


if __name__ == "__main__":
    test_sample_totals()
    test_additional_new_part_groups()
    test_ignored_eur_movement()
    test_rounding()
    print("Minden feldolgozási teszt sikeres.")