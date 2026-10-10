"""Az előkészített import állapota és munkafüzeteinek lezárása."""

from dataclasses import dataclass
from datetime import date

from openpyxl.workbook.workbook import Workbook as WorkbookType

from .models import DailyExchangeRate, Movement


@dataclass
class PreparedImport:
    workbooks: dict[int, WorkbookType]
    new_movements: list[Movement]
    exchange_rates: dict[date, DailyExchangeRate]


def close_prepared_import(prepared: PreparedImport) -> None:
    """Az előkészítés során megnyitott munkafüzetek bezárása."""
    for workbook in prepared.workbooks.values():
        workbook.close()
