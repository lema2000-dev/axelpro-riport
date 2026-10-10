"""Feldolgozott napi és havi összegek beírása a célcellákba."""

from datetime import date, datetime

from openpyxl.workbook.workbook import Workbook as WorkbookType
from openpyxl.worksheet.worksheet import Worksheet

from .excel_constants import DAILY_TARGET_ROWS, MONTHLY_TARGET_ROWS, MONTH_NAMES
from .excel_documents import add_processed_documents
from .excel_layout import ensure_months_through
from .models import DailyTotals, ImportResult, MonthlyTotals


def find_day_column(sheet: Worksheet, target_date: date) -> int:
    """Az adott nap dátumcellájának oszlopszámát keresi meg."""
    for cell in sheet[2]:
        value = cell.value

        # A fájlból visszaolvasott Excel-dátum datetime is lehet.
        if isinstance(value, datetime):
            value = value.date()

        if value == target_date:
            return cell.column

    raise ValueError(
        f"A(z) {target_date:%Y-%m-%d} dátum nem található a(z) {sheet.title} munkalapban."
    )


def add_cell_amount(sheet: Worksheet, row: int, column: int, amount: int) -> None:
    """Az egész forintra kerekített növekmény hozzáadása a cellához."""
    cell = sheet.cell(row=row, column=column)
    current_value = cell.value

    if current_value is None:
        current_value = 0

    if isinstance(current_value, bool) or not isinstance(current_value, (int, float)):
        raise ValueError(
            f"A(z) {sheet.title}!{cell.coordinate} cella nem számot tartalmaz."
        )

    cell.value = current_value + amount


def write_daily_totals(sheet: Worksheet, totals: DailyTotals) -> None:
    """Az új napi összegek hozzáadása a megfelelő célcellákhoz."""
    column = find_day_column(sheet, totals.movement_date)

    for field_name, row in DAILY_TARGET_ROWS.items():
        amount = getattr(totals, field_name)

        add_cell_amount(sheet=sheet, row=row, column=column, amount=amount)


def write_monthly_totals(sheet: Worksheet, totals: MonthlyTotals) -> None:
    """Az új havi összegek hozzáadása a havi záró celláihoz."""
    for field_name, row in MONTHLY_TARGET_ROWS.items():
        amount = getattr(totals, field_name)

        add_cell_amount(sheet=sheet, row=row, column=2, amount=amount)


def write_import_result(
    workbook: WorkbookType, template_sheet: Worksheet, year: int, result: ImportResult
) -> None:
    """Az import adatok adott évhez tartozó napi és havi összegeinek beírása."""
    months = {
        totals.movement_date.month
        for totals in result.daily_totals
        if totals.movement_date.year == year
    } | {totals.month for totals in result.monthly_totals if totals.year == year}
    if months:
        ensure_months_through(workbook, template_sheet, year, max(months))

    for totals in result.daily_totals:
        if totals.movement_date.year != year:
            continue

        month = totals.movement_date.month

        sheet = workbook[MONTH_NAMES[month - 1]]
        write_daily_totals(sheet=sheet, totals=totals)

    for totals in result.monthly_totals:
        if totals.year != year:
            continue

        sheet = workbook[MONTH_NAMES[totals.month - 1]]
        write_monthly_totals(sheet=sheet, totals=totals)

    year_documents = {
        document_key for document_key in result.new_documents if document_key[0] == year
    }

    add_processed_documents(workbook=workbook, documents=year_documents)
