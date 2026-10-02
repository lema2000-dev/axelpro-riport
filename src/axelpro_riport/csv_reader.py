import csv
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from .models import Movement

REQUIRED_COLUMNS = {
    "Árumozgás időpontja",
    "Ügylet típus",
    "Bizonylatszám",
    "Csoport",
    "Cikkszám",
    "Ne. érték",
    "Pénznem",
}

def parse_net_value(value: str) -> Decimal:
    normalized = value.strip().replace(',', '.')

    try:
        result = Decimal(normalized)
    except InvalidOperation as exc:
        raise ValueError(f"Hibás nettó érték: {value!r}") from exc

    if not result.is_finite():
        raise ValueError(f"Nem véges nettó érték: {value!r}")

    return result

def parse_movement_date(value: str) -> date:
    normalized = value.strip()

    try:
        timestamp = datetime.strptime(normalized, "%Y. %m. %d. %H:%M:%S")
    except ValueError as exc:
        raise ValueError(f"Hibás árumozgási időpont: {value!r}") from exc

    return timestamp.date()
    
def parse_movement(row: dict[str, str]) -> Movement:
    currency = row["Pénznem"].strip().upper()

    if currency not in ("HUF", "EUR"):
        raise ValueError(f"Nem támogatott pénznem: {currency!r}")

    document_number = row["Bizonylatszám"].strip()

    if not document_number:
        raise ValueError("Hiányzó bizonylatszám")

    return Movement(
        movement_date=parse_movement_date(row["Árumozgás időpontja"]),
        transaction_type=row["Ügylet típus"].strip(),
        document_number=document_number,
        group=row["Csoport"].strip(),
        item_code=row["Cikkszám"].strip(),
        net_value=parse_net_value(row["Ne. érték"]),
        currency=currency,
    )

def read_movements(path: str | Path) -> list[Movement]:
    movements: list[Movement] = []

    with Path(path).open(encoding="utf-16", newline="") as file:
        reader = csv.DictReader(file, delimiter="\t")

        columns = set(reader.fieldnames or [])
        missing_columns = REQUIRED_COLUMNS - columns

        if missing_columns:
            raise ValueError("Hiányzó CSV-oszlopok: " + ", ".join(sorted(missing_columns)))

        for row in reader:
            if None in row or any(value is None for value in row.values()):
                raise ValueError(f"Hibás CSV-szerkezet a(z) {reader.line_num}. sorban.")

            if row["Ügylet típus"].strip().casefold() == "leltár":
                continue

            try:
                movement = parse_movement(row)
            except ValueError as exc:
                raise ValueError(f"Hiba a(z) {reader.line_num}. sorban: {exc}") from exc

            movements.append(movement)

    return movements

