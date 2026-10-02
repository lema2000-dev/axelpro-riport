from dataclasses import dataclass
from datetime import date
from decimal import Decimal


@dataclass(frozen=True)
class Movement:
    movement_date: date
    transaction_type: str
    document_number: str
    group: str
    item_code: str
    net_value: Decimal
    currency: str

    @property
    def document_key(self) -> tuple[int, str]:
        return self.movement_date.year, self.document_number

@dataclass(frozen=True)
class DailyTotals:
    movement_date: date
    used_parts: int
    new_parts: int
    labor: int
    workshop_used_parts: int
    workshop_new_parts: int
    purchases: int
    shipping: int
    trailer: int

@dataclass(frozen=True)
class MonthlyTotals:
    year: int
    month: int
    sany: int
    diag: int
    kj: int

@dataclass(frozen=True)
class ImportResult:
    daily_totals: list[DailyTotals]
    monthly_totals: list[MonthlyTotals]
    new_documents: set[tuple[int, str]]

