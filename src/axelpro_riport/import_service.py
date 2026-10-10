"""A CSV-import és az éves riportok frissítéseinek összehangolása."""
from pathlib import Path
from datetime import date, datetime
from decimal import Decimal
from dataclasses import dataclass
from collections.abc import Callable

from openpyxl.workbook.workbook import Workbook as WorkbookType
from openpyxl.worksheet.worksheet import Worksheet

from .csv_reader import read_movements
from .processor import process_movements, filter_new_movements
from .excel_writer import (
    load_or_create_year_workbook,
    read_processed_documents,
    write_import_result,
    get_year_workbook_path,
    save_workbook_safely,
    read_daily_exchange_rates,
    add_daily_exchange_rates,
    update_monthly_exchange_rates,
)

from .models import Movement, ImportResult, DailyExchangeRate

@dataclass
class PreparedImport:
    workbooks: dict[int, WorkbookType]
    new_movements: list[Movement]
    exchange_rates: dict[date, DailyExchangeRate]

def get_movement_years(
    movements: list[Movement]
) -> set[int]:
    """Az árumozgásokban szereplő évek növekvő sorrendben."""
    return sorted({
        movement.movement_date.year for movement in movements
    })

def load_year_workbooks(
    output_directory: str | Path,
    template_sheet: Worksheet,
    movements: list[Movement]
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
    workbooks: dict[int, WorkbookType]
) -> set[tuple[int, str]]:
    """A megnyitott éves riportok bizonylatkulcsinak összegyűjtése."""
    processed_documents: set[tuple[int, str]] = set()

    for workbook in workbooks.values():
        processed_documents.update(read_processed_documents(workbook))

    return processed_documents

def apply_import_result(
    workbooks: dict[int, WorkbookType],
    template_sheet: Worksheet,
    result: ImportResult
) -> None:
    """Az új össszegek és bizonylatkulcsok beírása at éves riportokba."""
    affected_years = {year for year, _ in result.new_documents}

    missing_years = affected_years - set(workbooks)

    if missing_years:
        raise ValueError(
            f"Hiányzó éves munkafüzetek: {sorted(missing_years)}"
        )

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
    additional_years: set[int] | None = None
) -> dict[int, Path | None]:
    """Az új bizonylatokkal frissített éves riportok mentése."""
    affected_years = {year for year, _ in result.new_documents}

    if additional_years is not None:
        affected_years.update(additional_years)

    missing_years = affected_years - set(workbooks)

    if missing_years:
        raise ValueError(
            f"Hiányzó éves munkafüzetek: {sorted(missing_years)}"
        )

    backups: dict[int, Path | None] = {}

    for year in sorted(affected_years):
        path = get_year_workbook_path(
            output_directory=output_directory,
            year=year,
        )

        backups[year] = save_workbook_safely(
            workbook=workbooks[year],
            workbook_path=path,
        )

    return backups

def import_csv(
    csv_path: str | Path,
    output_directory: str | Path,
    template_sheet: Worksheet,
    current_source_date: date,
    current_rate: Decimal,
    additional_new_part_groups: set[str],
    *,
    request_manual_rates: Callable[
        [list[date]],
        dict[date, DailyExchangeRate],
    ] | None = None,
) -> ImportResult:
    """A teljes importfolyamat végrehajtása."""
    prepared = prepare_import(
        csv_path=csv_path,
        output_directory=output_directory,
        template_sheet=template_sheet,
    )

    handed_to_finish = False

    try:
        add_source_exchange_rate(
            prepared=prepared,
            source_date=current_source_date,
            rate=current_rate,
        )

        missing_dates = find_missing_exchange_rate_dates(
            prepared.new_movements,
            prepared.exchange_rates,
        )

        prepared.exchange_rates = fill_missing_exchange_rates(
            missing_dates,
            prepared.exchange_rates,
        )

        missing_dates = find_missing_exchange_rate_dates(
            prepared.new_movements,
            prepared.exchange_rates,
        )

        if missing_dates:
            if request_manual_rates is None:
                raise ValueError(
                    f"Kézi árfolyammegadás szükséges: {missing_dates}"
                )

            manual_rates = request_manual_rates(missing_dates)

            for day, record in manual_rates.items():
                if not isinstance(record, DailyExchangeRate):
                    raise ValueError(
                        f"Hibás kézi árfolyamrekord: {record!r}"
                    )

                if (
                    day != record.application_date
                    or day not in missing_dates
                ):
                    raise ValueError(
                        f"Nem kért alkalmazási dátum: {day!r}"
                    )

                add_manual_exchange_rate(
                    prepared=prepared,
                    application_date=day,
                    source_date=record.source_date,
                    rate=record.rate,
                )

        handed_to_finish = True

        return finish_import(
            prepared=prepared,
            output_directory=output_directory,
            template_sheet=template_sheet,
            additional_new_part_groups=additional_new_part_groups,
        )

    finally:
        if not handed_to_finish:
            close_prepared_import(prepared)

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

        if (
            source_date in source_rates
            and source_rates[source_date] != record.rate
        ):
            raise ValueError(
                f"Eltérő árfolyamok ugyanahhoz a forrásdátumhoz: "
                f"{source_date:%Y-%m-%d}"
            )

        source_rates[source_date] = record.rate

    for day in missing_dates:
        previous_dates = [
            source_date
            for source_date in source_rates
            if source_date <= day
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
    exchange_rates: dict[date, DailyExchangeRate]
) -> dict[date, Decimal]:
    """Az alkamazási datumokhoz tartozó árfolyamértékek kigyűjtése."""
    return {
        day : record.rate
        for day, record in exchange_rates.items()
    }

def get_monthly_exchange_rate_values(
    workbook: WorkbookType,
) -> tuple:
    """A havi árfolyamtábla értékei a változások összehasonlításához."""
    if "Euro" not in workbook.sheetnames:
        return ()

    return tuple(
        workbook["Euro"].iter_rows(
            min_row=1,
            max_row=13,
            max_col=3,
            values_only=True,
        )
    )

def apply_exchange_rates(
    workbooks: dict[int, WorkbookType],
    exchange_rates: dict[date, DailyExchangeRate],
) -> set[int]:
    """Árfolyamok beírása; visszaadja a módosult riportok éveit."""
    changed_years: set[int] = set()

    for year, workbook in workbooks.items():
        year_rates = {
            day: record
            for day, record in exchange_rates.items()
            if day.year == year
        }

        existing_rates = read_daily_exchange_rates(workbook)
        previous_monthly_values = get_monthly_exchange_rate_values(
            workbook
        )

        add_daily_exchange_rates(
            workbook=workbook,
            exchange_rates=year_rates,
        )

        update_monthly_exchange_rates(
            workbook=workbook,
            year=year,
        )

        has_new_rates = bool(
            set(year_rates) - set(existing_rates)
        )
        monthly_values_changed = (
            get_monthly_exchange_rate_values(workbook)
            != previous_monthly_values
        )

        if has_new_rates or monthly_values_changed:
            changed_years.add(year)

    return changed_years

def prepare_import(
    csv_path: str | Path,
    output_directory: str | Path,
    template_sheet: Worksheet,
) -> PreparedImport:
    """Riportok megnyitása, új tételek és tárolt árfolyamok előkészítése."""
    movements = read_movements(csv_path)

    workbooks = load_year_workbooks(
        output_directory=output_directory,
        template_sheet=template_sheet,
        movements=movements,
    )

    try:
        processed_documents = collect_processed_documents(workbooks)

        new_movements = filter_new_movements(
            movements,
            processed_documents,
        )

        exchange_rates = collect_daily_exchange_rates(workbooks)

        return PreparedImport(
            workbooks=workbooks,
            new_movements=new_movements,
            exchange_rates=exchange_rates,
        )

    except Exception:
        for workbook in workbooks.values():
            workbook.close()
        raise

def add_source_exchange_rate(
    prepared: PreparedImport,
    source_date: date,
    rate: Decimal,
) -> None:
    """Egy tényleges napi forrásárfolyam hozzáadása az importhoz."""
    if (
        not isinstance(source_date, date)
        or isinstance(source_date, datetime)
        or not isinstance(rate, Decimal)
        or not rate.is_finite()
        or rate <= 0
    ):
        raise ValueError(
            f"Hibás forrásárfolyam: {source_date!r} -> {rate!r}"
        )

    record = DailyExchangeRate(
        application_date=source_date,
        source_date=source_date,
        rate=rate,
    )

    existing = prepared.exchange_rates.get(source_date)

    if existing is not None and existing != record:
        raise ValueError(
            f"A(z) {source_date:%Y-%m-%d} naphoz "
            "már eltérő árfolyamrekord tartozik."
        )

    prepared.exchange_rates[source_date] = record

def add_manual_exchange_rate(
    prepared: PreparedImport,
    application_date: date,
    source_date: date,
    rate: Decimal,
) -> None:
    """Hiányzó alkalmazási nap árfolyamának kézi megadása."""
    if (
        not isinstance(application_date, date)
        or isinstance(application_date, datetime)
        or not isinstance(source_date, date)
        or isinstance(source_date, datetime)
        or source_date > application_date
        or not isinstance(rate, Decimal)
        or not rate.is_finite()
        or rate <= 0
    ):
        raise ValueError("Hibás kézzel megadott árfolyam.")

    record = DailyExchangeRate(
        application_date=application_date,
        source_date=source_date,
        rate=rate,
    )

    existing = prepared.exchange_rates.get(application_date)

    if existing is not None and existing != record:
        raise ValueError(
            f"A(z) {application_date:%Y-%m-%d} naphoz "
            "már eltérő árfolyamrekord tartozik."
        )

    for known_record in prepared.exchange_rates.values():
        if (
            known_record.source_date == source_date
            and known_record.rate != rate
        ):
            raise ValueError(
                f"A(z) {source_date:%Y-%m-%d} forrásdátumhoz "
                "már eltérő árfolyam tartozik."
            )

    prepared.exchange_rates[application_date] = record

def finish_import(
    prepared: PreparedImport,
    output_directory: str | Path,
    template_sheet: Worksheet,
    additional_new_part_groups: set[str],
) -> ImportResult:
    """Az előkészített import feldolgozása, mentése és lezárása."""
    try:
        missing_dates = find_missing_exchange_rate_dates(
            prepared.new_movements,
            prepared.exchange_rates,
        )

        if missing_dates:
            raise ValueError(
                f"Hiányzó napi árfolyamok: {missing_dates}"
            )

        processor_rates = get_processor_exchange_rates(
            prepared.exchange_rates
        )

        result = process_movements(
            movements=prepared.new_movements,
            processed_documents=set(),
            exchange_rates=processor_rates,
            additional_new_part_groups=additional_new_part_groups,
        )

        apply_import_result(
            workbooks=prepared.workbooks,
            template_sheet=template_sheet,
            result=result,
        )

        changed_years = apply_exchange_rates(
            workbooks=prepared.workbooks,
            exchange_rates=prepared.exchange_rates,
        )

        save_year_workbooks(
            output_directory=output_directory,
            workbooks=prepared.workbooks,
            result=result,
            additional_years=changed_years,
        )

        return result

    finally:
        for workbook in prepared.workbooks.values():
            workbook.close()

def close_prepared_import(prepared: PreparedImport) -> None:
    """Az előkészítés során megnyitott munkafüzetek bezárása."""
    for workbook in prepared.workbooks.values():
        workbook.close()

