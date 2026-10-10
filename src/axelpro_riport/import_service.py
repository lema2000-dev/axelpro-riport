"""A CSV-import előkészítése és teljes végrehajtási sorrendje."""

from collections.abc import Callable
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl.worksheet.worksheet import Worksheet

from .csv_reader import read_movements
from .import_state import PreparedImport, close_prepared_import
from .models import DailyExchangeRate, ImportResult
from .processor import aggregate_movements, filter_new_movements
from .rate_logic import (
    add_source_exchange_rate,
    fill_missing_exchange_rates,
    find_missing_exchange_rate_dates,
    get_processor_exchange_rates,
    store_exchange_rate,
)
from .report_service import (
    apply_exchange_rates,
    apply_import_result,
    collect_daily_exchange_rates,
    collect_processed_documents,
    load_year_workbooks,
    save_year_workbooks,
)


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


def _finish_import(
    prepared: PreparedImport,
    output_directory: str | Path,
    template_sheet: Worksheet,
    additional_new_part_groups: set[str],
) -> ImportResult:
    """Az előkészített import feldolgozása, mentése és lezárása."""
    missing_dates = find_missing_exchange_rate_dates(
        prepared.new_movements,
        prepared.exchange_rates,
    )

    if missing_dates:
        raise ValueError(f"Hiányzó napi árfolyamok: {missing_dates}")

    processor_rates = get_processor_exchange_rates(prepared.exchange_rates)

    result = aggregate_movements(
        movements=prepared.new_movements,
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


def finish_import(
    prepared: PreparedImport,
    output_directory: str | Path,
    template_sheet: Worksheet,
    additional_new_part_groups: set[str],
) -> ImportResult:
    """Önálló befejezés: feldolgozás és mentés, majd lezárás hiba esetén is."""
    try:
        return _finish_import(
            prepared, output_directory, template_sheet, additional_new_part_groups
        )
    finally:
        close_prepared_import(prepared)


def import_csv(
    csv_path: str | Path,
    output_directory: str | Path,
    template_sheet: Worksheet,
    current_source_date: date,
    current_rate: Decimal,
    additional_new_part_groups: set[str],
    *,
    request_manual_rates: (
        Callable[
            [list[date]],
            dict[date, DailyExchangeRate],
        ]
        | None
    ) = None,
) -> ImportResult:
    """A teljes importfolyamat végrehajtása."""
    prepared = prepare_import(
        csv_path=csv_path,
        output_directory=output_directory,
        template_sheet=template_sheet,
    )

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
                raise ValueError(f"Kézi árfolyammegadás szükséges: {missing_dates}")

            manual_rates = request_manual_rates(missing_dates)

            for day, record in manual_rates.items():
                if not isinstance(record, DailyExchangeRate):
                    raise ValueError(f"Hibás kézi árfolyamrekord: {record!r}")

                if day != record.application_date or day not in missing_dates:
                    raise ValueError(f"Nem kért alkalmazási dátum: {day!r}")

                store_exchange_rate(prepared.exchange_rates, record)

        return _finish_import(
            prepared=prepared,
            output_directory=output_directory,
            template_sheet=template_sheet,
            additional_new_part_groups=additional_new_part_groups,
        )

    finally:
        close_prepared_import(prepared)
