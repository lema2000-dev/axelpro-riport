from datetime import date
from decimal import Decimal, ROUND_HALF_UP

from .models import Movement, DailyTotals, MonthlyTotals, ImportResult


def filter_new_movements(
    movements: list[Movement], processed_documents: set[tuple[int, str]]
) -> list[Movement]:
    return [
        movement
        for movement in movements
        if movement.document_key not in processed_documents
    ]


def find_workshop_documents(movements: list[Movement]) -> set[tuple[int, str]]:
    return {
        movement.document_key
        for movement in movements
        if movement.transaction_type.casefold() == "számla"
        and movement.group.casefold() == "munkadíj"
    }


def net_value_in_huf(
    movement: Movement, exchange_rates: dict[date, Decimal]
) -> Decimal:
    if movement.currency == "HUF":
        return movement.net_value

    rate = exchange_rates.get(movement.movement_date)

    if rate is None:
        raise ValueError(f"Hiányzó EUR/HUF árfolyam: {movement.movement_date}")

    if not rate.is_finite() or rate <= 0:
        raise ValueError(f"Hibás EUR/HUF árfolyam: {movement.movement_date} -> {rate}")

    return movement.net_value * rate


def round_huf(value: Decimal) -> int:
    return int(value.quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def is_new_part(movement: Movement, new_part_groups: set[str]) -> bool:
    return movement.group.casefold() in new_part_groups


def prepare_new_part_groups(additional_groups: set[str]) -> set[str]:
    return {"új alkatrész"} | {
        group.strip().casefold() for group in additional_groups if group.strip()
    }


def calculate_daily_totals(
    movement_date: date,
    movements: list[Movement],
    exchange_rates: dict[date, Decimal],
    new_part_groups: set[str],
    workshop_documents: set[tuple[int, str]],
) -> DailyTotals:
    totals = {
        "used_parts": Decimal(0),
        "new_parts": Decimal(0),
        "labor": Decimal(0),
        "workshop_used_parts": Decimal(0),
        "workshop_new_parts": Decimal(0),
        "purchases": Decimal(0),
        "shipping": Decimal(0),
        "trailer": Decimal(0),
    }

    for movement in movements:
        if movement.movement_date != movement_date:
            continue

        transaction_type = movement.transaction_type.casefold()
        group = movement.group.casefold()

        if transaction_type == "számla" and (
            group == "bontott alkatrész"
            or is_new_part(movement, new_part_groups)
            or group == "munkadíj"
        ):
            value = net_value_in_huf(movement, exchange_rates)

            if group == "bontott alkatrész":
                totals["used_parts"] += value

                if movement.document_key in workshop_documents:
                    totals["workshop_used_parts"] += value

            elif is_new_part(movement, new_part_groups):
                totals["new_parts"] += value

                if movement.document_key in workshop_documents:
                    totals["workshop_new_parts"] += value

            elif group == "munkadíj":
                totals["labor"] += value

        if transaction_type == "beszerzés":
            value = net_value_in_huf(movement, exchange_rates)
            totals["purchases"] += value

        if group == "szállítási költség":
            value = net_value_in_huf(movement, exchange_rates)
            totals["shipping"] += value

        if group == "trailer":
            value = net_value_in_huf(movement, exchange_rates)
            totals["trailer"] += value

    return DailyTotals(
        movement_date=movement_date,
        used_parts=round_huf(totals["used_parts"]),
        new_parts=round_huf(totals["new_parts"]),
        labor=round_huf(totals["labor"]),
        workshop_used_parts=round_huf(totals["workshop_used_parts"]),
        workshop_new_parts=round_huf(totals["workshop_new_parts"]),
        purchases=round_huf(totals["purchases"]),
        shipping=round_huf(totals["shipping"]),
        trailer=round_huf(totals["trailer"]),
    )


def calculate_monthly_totals(
    year: int,
    month: int,
    movements: list[Movement],
    exchange_rates: dict[date, Decimal],
) -> MonthlyTotals:
    totals = {
        "sany": Decimal(0),
        "diag": Decimal(0),
        "kj": Decimal(0),
    }

    for movement in movements:
        if movement.movement_date.year != year or movement.movement_date.month != month:
            continue

        if movement.transaction_type.casefold() != "számla":
            continue

        item_code = movement.item_code.casefold()

        if item_code not in totals:
            continue

        value = net_value_in_huf(movement, exchange_rates)
        totals[item_code] += value

    return MonthlyTotals(
        year=year,
        month=month,
        sany=round_huf(totals["sany"]),
        diag=round_huf(totals["diag"]),
        kj=round_huf(totals["kj"]),
    )


def aggregate_movements(
    movements: list[Movement],
    exchange_rates: dict[date, Decimal],
    additional_new_part_groups: set[str],
) -> ImportResult:
    """Már előszűrt tételek napi és havi összesítése."""
    new_movements = movements

    new_part_groups = prepare_new_part_groups(additional_new_part_groups)

    workshop_documents = find_workshop_documents(new_movements)

    dates = sorted({movement.movement_date for movement in new_movements})

    months = sorted({(day.year, day.month) for day in dates})

    daily_totals = [
        calculate_daily_totals(
            movement_date=day,
            movements=new_movements,
            exchange_rates=exchange_rates,
            new_part_groups=new_part_groups,
            workshop_documents=workshop_documents,
        )
        for day in dates
    ]

    monthly_totals = [
        calculate_monthly_totals(
            year=year,
            month=month,
            movements=new_movements,
            exchange_rates=exchange_rates,
        )
        for year, month in months
    ]

    new_documents = {movement.document_key for movement in new_movements}

    return ImportResult(
        daily_totals=daily_totals,
        monthly_totals=monthly_totals,
        new_documents=new_documents,
    )


def process_movements(
    movements: list[Movement],
    processed_documents: set[tuple[int, str]],
    exchange_rates: dict[date, Decimal],
    additional_new_part_groups: set[str],
) -> ImportResult:
    """Szűrés, majd napi és havi összesítés."""
    return aggregate_movements(
        filter_new_movements(movements, processed_documents),
        exchange_rates,
        additional_new_part_groups,
    )
