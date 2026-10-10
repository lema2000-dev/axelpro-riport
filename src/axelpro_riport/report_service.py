"""Több éves riport megnyitása, frissítése és mentése."""

from datetime import date
from pathlib import Path

from openpyxl.workbook.workbook import Workbook as WorkbookType
from openpyxl.worksheet.worksheet import Worksheet

from .excel_documents import read_processed_documents
from .excel_files import (
    get_year_workbook_path,
    load_or_create_year_workbook,
    save_workbook_safely,
)
from .excel_rates import (
    add_daily_exchange_rates,
    get_monthly_exchange_rate_values,
    read_daily_exchange_rates,
    update_monthly_exchange_rates,
)
from .excel_writer import write_import_result
from .models import DailyExchangeRate, ImportResult, Movement


def get_movement_years(movements: list[Movement]) -> list[int]:
    """Az árumozgásokban szereplő évek növekvő sorrendben."""
    return sorted({movement.movement_date.year for movement in movements})


def load_year_workbooks(
    output_directory: str | Path, template_sheet: Worksheet, movements: list[Movement]
) -> dict[int, WorkbookType]:
    """Az árumozgásokhoz szükséges éves munkafüzetek megnyitása vagy létrehozása."""
    workbooks: dict[int, WorkbookType] = {}

    try:
        for year in get_movement_years(movements):
            last_month = max(
                movement.movement_date.month
                for movement in movements
                if movement.movement_date.year == year
            )

            workbooks[year] = load_or_create_year_workbook(
                output_directory=output_directory,
                template_sheet=template_sheet,
                year=year,
                last_month=last_month,
            )

        return workbooks

    except Exception:
        for workbook in workbooks.values():
            workbook.close()

        raise


def collect_processed_documents(
    workbooks: dict[int, WorkbookType],
) -> set[tuple[int, str]]:
    """A megnyitott éves riportok bizonylatkulcsainak összegyűjtése."""
    processed_documents: set[tuple[int, str]] = set()

    for workbook in workbooks.values():
        processed_documents.update(read_processed_documents(workbook))

    return processed_documents


def apply_import_result(
    workbooks: dict[int, WorkbookType], template_sheet: Worksheet, result: ImportResult
) -> None:
    """Az új összegek és bizonylatkulcsok beírása az éves riportokba."""
    affected_years = {year for year, _ in result.new_documents}

    missing_years = affected_years - set(workbooks)

    if missing_years:
        raise ValueError(f"Hiányzó éves munkafüzetek: {sorted(missing_years)}")

    for year in sorted(affected_years):
        write_import_result(
            workbook=workbooks[year],
            template_sheet=template_sheet,
            year=year,
            result=result,
        )


def save_year_workbooks(
    output_directory: str | Path,
    workbooks: dict[int, WorkbookType],
    result: ImportResult,
    *,
    additional_years: set[int] | None = None,
) -> None:
    """Az új bizonylatokkal frissített éves riportok mentése."""
    affected_years = {year for year, _ in result.new_documents}

    if additional_years is not None:
        affected_years.update(additional_years)

    for year in sorted(affected_years):
        path = get_year_workbook_path(
            output_directory=output_directory,
            year=year,
        )

        save_workbook_safely(
            workbook=workbooks[year],
            workbook_path=path,
        )


def collect_daily_exchange_rates(
    workbooks: dict[int, WorkbookType],
) -> dict[date, DailyExchangeRate]:
    """A megnyitott éves riportok napi árfolyamainak összegyűjtése."""
    exchange_rates: dict[date, DailyExchangeRate] = {}

    for year, workbook in workbooks.items():
        stored_rates = read_daily_exchange_rates(workbook)

        for day, record in stored_rates.items():
            if day.year != year:
                raise ValueError(
                    f"A(z) {year}. évi riport más évhez tartozó "
                    f"alkalmazási dátumot tartalmaz: {day:%Y-%m-%d}"
                )

            exchange_rates[day] = record

    return exchange_rates


def apply_exchange_rates(
    workbooks: dict[int, WorkbookType],
    exchange_rates: dict[date, DailyExchangeRate],
) -> set[int]:
    """Árfolyamok beírása; visszaadja a módosult riportok éveit."""
    changed_years: set[int] = set()

    for year, workbook in workbooks.items():
        year_rates = {
            day: record for day, record in exchange_rates.items() if day.year == year
        }

        existing_rates = read_daily_exchange_rates(workbook)
        previous_monthly_values = get_monthly_exchange_rate_values(workbook)

        add_daily_exchange_rates(
            workbook=workbook,
            exchange_rates=year_rates,
        )

        update_monthly_exchange_rates(
            workbook=workbook,
            year=year,
        )

        has_new_rates = bool(set(year_rates) - set(existing_rates))
        monthly_values_changed = (
            get_monthly_exchange_rate_values(workbook) != previous_monthly_values
        )

        if has_new_rates or monthly_values_changed:
            changed_years.add(year)

    return changed_years
