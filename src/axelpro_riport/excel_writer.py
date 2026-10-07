"""A riport munkalapjainak létrehozása a végleges Excel-sablon alapján."""

from copy import copy
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook as WorkbookType
from openpyxl.worksheet.worksheet import Worksheet

from .calendar_utils import get_month_weeks


MONTH_NAMES = (
    "Január",
    "Február",
    "Március",
    "Április",
    "Május",
    "Június",
    "Július",
    "Augusztus",
    "Szeptember",
    "Október",
    "November",
    "December"
)

WEEK_START_ROW = 1
WEEK_END_ROW = 40
WEEK_START_COLUMN = 1
WEEK_END_COLUMN = 15

MONTHLY_CLOSE_START_ROW = 43
MONTHLY_CLOSE_END_ROW = 85
MONTHLY_CLOSE_START_COLUMN = 1
MONTHLY_CLOSE_END_COLUMN = 3


# Sablon betöltése


def load_template(template_path: str | Path) -> WorkbookType:
    """A végleges szeptemberi sablont tartalmazó munkafüzet betöltése."""
    path = Path(template_path)

    workbook = load_workbook(path, data_only=False)

    if "Szeptember" not in workbook.sheetnames:
        workbook.close()
        raise ValueError("A sablon nem tartalmaz Szeptember nevű munkalapot.")

    return workbook


# Cellák és tartományok másolása


def copy_cell(
    source_sheet: Worksheet,
    target_sheet: Worksheet,
    source_row: int,
    source_column: int,
    target_row: int,
    target_column: int,
    *,
    translate_formulas: bool = True
) -> None:
    """Egy cella értékének, képletének és formázásának másolása."""
    source = source_sheet.cell(row=source_row, column=source_column)
    target = target_sheet.cell(row=target_row, column=target_column)

    if source.data_type == "f" and translate_formulas:
        target.value = Translator(
            source.value,
            origin=source.coordinate
        ).translate_formula(target.coordinate)
    else:
        target.value = source.value

    # A stílusobjektumokat másoljuk, így nem osztunk meg StyleProxy-t.
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.alignment = copy(source.alignment)
    target.number_format = source.number_format
    target.protection = copy(source.protection)


def copy_range(
    source_sheet: Worksheet,
    target_sheet: Worksheet,
    min_row: int,
    max_row: int,
    min_column: int,
    max_column: int,
    target_start_row: int,
    target_start_column: int
) -> None:
    """Egy téglalap alakú cellatartomány másolása."""
    for source_row in range(min_row, max_row + 1):
        for source_column in range(min_column, max_column + 1):
            target_row = target_start_row + (source_row - min_row)
            target_column = target_start_column + (source_column - min_column)

            copy_cell(
                source_sheet=source_sheet,
                target_sheet=target_sheet,
                source_row=source_row,
                source_column=source_column,
                target_row=target_row,
                target_column=target_column
            )


def copy_row_heights(
    source_sheet: Worksheet,
    target_sheet: Worksheet,
    min_row: int,
    max_row: int,
    target_start_row: int
) -> None:
    """A másolt sorok magasságának átvitele."""
    for source_row in range(min_row, max_row + 1):
        target_row = target_start_row + (source_row - min_row)
        target_sheet.row_dimensions[target_row].height = source_sheet.row_dimensions[source_row].height


def copy_column_widths(
    source_sheet: Worksheet,
    target_sheet: Worksheet,
    min_column: int,
    max_column: int,
    target_start_column: int
) -> None:
    """A másolt oszlopok szélességének átvitele."""
    for source_column in range(min_column, max_column + 1):
        target_column = target_start_column + (source_column - min_column)

        source_letter = get_column_letter(source_column)
        target_letter = get_column_letter(target_column)

        target_sheet.column_dimensions[target_letter].width = source_sheet.column_dimensions[source_letter].width


def copy_merged_ranges(
    source_sheet: Worksheet,
    target_sheet: Worksheet,
    min_row: int,
    max_row: int,
    min_column: int,
    max_column: int,
    target_start_row: int,
    target_start_column: int
) -> None:
    """A másolt tartományon belüli cellaösszevonások átvitele."""
    row_offset = target_start_row - min_row
    column_offset = target_start_column - min_column

    for merged_range in list(source_sheet.merged_cells.ranges):
        if (
            merged_range.min_row >= min_row and
            merged_range.max_row <= max_row and
            merged_range.min_col >= min_column and
            merged_range.max_col <= max_column
        ):
            target_sheet.merge_cells(
                start_row=merged_range.min_row + row_offset,
                end_row=merged_range.max_row + row_offset,
                start_column=merged_range.min_col + column_offset,
                end_column=merged_range.max_col + column_offset
            )


def copy_block(
    source_sheet: Worksheet,
    target_sheet: Worksheet,
    min_row: int,
    max_row: int,
    min_column: int,
    max_column: int,
    target_start_row: int,
    target_start_column: int
) -> None:
    """Cellák, méretek és összevonások együttes másolása."""
    copy_range(
        source_sheet=source_sheet,
        target_sheet=target_sheet,
        min_row=min_row,
        max_row=max_row,
        min_column=min_column,
        max_column=max_column,
        target_start_row=target_start_row,
        target_start_column=target_start_column
    )

    copy_row_heights(
        source_sheet=source_sheet,
        target_sheet=target_sheet,
        min_row=min_row,
        max_row=max_row,
        target_start_row=target_start_row
    )

    copy_column_widths(
        source_sheet=source_sheet,
        target_sheet=target_sheet,
        min_column=min_column,
        max_column=max_column,
        target_start_column=target_start_column
    )

    # Az összevonások csak a cellák másolása után kerüljenek át.
    copy_merged_ranges(
        source_sheet=source_sheet,
        target_sheet=target_sheet,
        min_row=min_row,
        max_row=max_row,
        min_column=min_column,
        max_column=max_column,
        target_start_row=target_start_row,
        target_start_column=target_start_column
    )


def clear_input_values(
    sheet: Worksheet,
    min_row: int,
    max_row: int,
    min_column: int,
    max_column: int
) -> None:
    """A kézzel kitölthető cellák ürítése, a képletek megtartásával."""
    for row in sheet.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_column,
        max_col=max_column
    ):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.data_type != "f":
                cell.value = None


# Heti blokkok felépítése


def get_day_start_column(
    week_start_column: int,
    *,
    include_labels: bool,
) -> int:
    """A napi oszlopok kezdete, az esetleges sorfeliratok után."""
    return week_start_column + int(include_labels)


def get_week_trim_counts(
    week: list[date],
) -> tuple[int, int]:
    """A hiányzó hétfő–szombati napok száma a hét két szélén."""
    weekdays = [day.weekday() for day in week if day.weekday() < 6]

    if not weekdays:
        return 6, 0

    leading_days = min(weekdays)
    trailing_days = 5 - max(weekdays)

    return leading_days, trailing_days


def copy_week_days(
    template_sheet: Worksheet,
    target_sheet: Worksheet,
    week: list[date],
    target_start_column: int
) -> int:
    """A hónapba eső napi oszloppárok másolása és ürítése."""
    leading_days, trailing_days = get_week_trim_counts(week)
    day_count = 6 - leading_days - trailing_days

    if day_count == 0:
        return 0

    source_start_column = 2 + leading_days * 2
    source_end_column = 13 - trailing_days * 2

    copy_block(
        source_sheet=template_sheet,
        target_sheet=target_sheet,
        min_row=WEEK_START_ROW,
        max_row=WEEK_END_ROW,
        min_column=source_start_column,
        max_column=source_end_column,
        target_start_row=WEEK_START_ROW,
        target_start_column=target_start_column
    )

    clear_input_values(
        sheet=target_sheet,
        min_row=4,
        max_row=WEEK_END_ROW,
        min_column=target_start_column,
        max_column=target_start_column + day_count * 2 - 1
    )

    return day_count


def copy_week_closing(
    template_sheet: Worksheet,
    target_sheet: Worksheet,
    target_start_column: int
) -> None:
    """A heti záró két oszlopának másolása és ürítése."""
    # A range felső határa nem része a bejárásnak: N és O is kell.
    for row in range(WEEK_START_ROW, WEEK_END_ROW + 1):
        for source_column in range(14, 16):
            copy_cell(
                source_sheet=template_sheet,
                target_sheet=target_sheet,
                source_row=row,
                source_column=source_column,
                target_row=row,
                target_column=target_start_column + source_column - 14,
                translate_formulas=False
            )

    copy_column_widths(
        source_sheet=template_sheet,
        target_sheet=target_sheet,
        min_column=14,
        max_column=15,
        target_start_column=target_start_column
    )

    copy_merged_ranges(
        source_sheet=template_sheet,
        target_sheet=target_sheet,
        min_row=WEEK_START_ROW,
        max_row=WEEK_END_ROW,
        min_column=14,
        max_column=15,
        target_start_row=WEEK_START_ROW,
        target_start_column=target_start_column
    )

    clear_input_values(
        sheet=target_sheet,
        min_row=4,
        max_row=WEEK_END_ROW,
        min_column=target_start_column,
        max_column=target_start_column + 1
    )


def set_week_dates(
    sheet: Worksheet,
    week: list[date],
    week_start_column: int,
    *,
    include_labels: bool
) -> None:
    """A hónaphoz tartozó dátumok és az ISO-hétszám beírása."""
    days = [day for day in week if day.weekday() < 6]

    if not days:
        return

    day_start_column = get_day_start_column(
        week_start_column=week_start_column,
        include_labels=include_labels
    )

    for day_index, day in enumerate(days):
        column = day_start_column + day_index * 2
        sheet.cell(row=2, column=column).value = day

    closing_column = day_start_column + len(days) * 2
    sheet.cell(row=2, column=closing_column).value = days[0].isocalendar().week


def set_week_closing_formulas(
    template_sheet: Worksheet,
    target_sheet: Worksheet,
    day_start_column: int,
    day_count: int
) -> None:
    """A heti összegek újraépítése a ténylegesen jelen lévő napokból."""
    closing_column = day_start_column + day_count * 2

    for row in range(4, WEEK_END_ROW + 1):
        for currency_offset in (0, 1):
            source = template_sheet.cell(row=row, column=14 + currency_offset)

            if source.data_type != "f":
                continue

            references = [
                (
                    f"{get_column_letter(day_start_column + day_index * 2 + currency_offset)}{row}"
                )
                for day_index in range(day_count)
            ]

            target_sheet.cell(
                row=row,
                column=closing_column + currency_offset
            ).value = "=" + "+".join(references)


def create_week_block(
    template_sheet: Worksheet,
    target_sheet: Worksheet,
    week: list[date],
    week_start_column: int,
    *,
    include_labels: bool
) -> int:
    """Egy heti blokk létrehozása; a felhasznált oszlopok számát adja vissza."""
    leading_days, trailing_days = get_week_trim_counts(week)
    day_count = 6 - leading_days - trailing_days

    if day_count == 0:
        return 0

    if include_labels:
        copy_block(
            source_sheet=template_sheet,
            target_sheet=target_sheet,
            min_row=WEEK_START_ROW,
            max_row=WEEK_END_ROW,
            min_column=1,
            max_column=1,
            target_start_row=WEEK_START_ROW,
            target_start_column=week_start_column
        )

    day_start_column = get_day_start_column(
        week_start_column=week_start_column,
        include_labels=include_labels
    )

    copy_week_days(
        template_sheet=template_sheet,
        target_sheet=target_sheet,
        week=week,
        target_start_column=day_start_column
    )

    closing_column = day_start_column + day_count * 2

    copy_week_closing(
        template_sheet=template_sheet,
        target_sheet=target_sheet,
        target_start_column=closing_column
    )

    set_week_closing_formulas(
        template_sheet=template_sheet,
        target_sheet=target_sheet,
        day_start_column=day_start_column,
        day_count=day_count
    )

    set_week_dates(
        sheet=target_sheet,
        week=week,
        week_start_column=week_start_column,
        include_labels=include_labels
    )

    return int(include_labels) + day_count * 2 + 2


def create_month_weeks(
    template_sheet: Worksheet,
    target_sheet: Worksheet,
    year: int,
    month: int
) -> list[int]:
    """A hónap heti blokkjai; visszaadja a heti zárók kezdőoszlopait."""
    next_column = 1
    include_labels = True
    closing_columns = []

    for week in get_month_weeks(year, month):
        block_width = create_week_block(
            template_sheet=template_sheet,
            target_sheet=target_sheet,
            week=week,
            week_start_column=next_column,
            include_labels=include_labels
        )

        if block_width == 0:
            continue

        closing_columns.append(next_column + block_width - 2)
        next_column += block_width
        include_labels = False

    return closing_columns


# Havi záró és képletei


def build_monthly_sum(
    closing_columns: list[int],
    source_row: int,
    *,
    currency_offset: int = 0
) -> str:
    """A heti zárókat összeadó képletrész összeállítása."""
    references = [
        f"{get_column_letter(column + currency_offset)}{source_row}"
        for column in closing_columns
    ]

    return "(" + "+".join(references) + ")"


def copy_monthly_close(
    template_sheet: Worksheet,
    target_sheet: Worksheet
) -> None:
    """A havi záró sablonjának másolása, kézi értékek nélkül."""
    copy_block(
        source_sheet=template_sheet,
        target_sheet=target_sheet,
        min_row=MONTHLY_CLOSE_START_ROW,
        max_row=MONTHLY_CLOSE_END_ROW,
        min_column=MONTHLY_CLOSE_START_COLUMN,
        max_column=MONTHLY_CLOSE_END_COLUMN,
        target_start_row=MONTHLY_CLOSE_START_ROW,
        target_start_column=MONTHLY_CLOSE_START_COLUMN
    )

    clear_input_values(
        sheet=target_sheet,
        min_row=MONTHLY_CLOSE_START_ROW,
        max_row=MONTHLY_CLOSE_END_ROW,
        min_column=MONTHLY_CLOSE_START_COLUMN + 1,
        max_column=MONTHLY_CLOSE_END_COLUMN
    )


def set_monthly_basic_formulas(
    sheet: Worksheet,
    closing_columns: list[int]
) -> None:
    """A forintos havi összegek képleteinek beállítása."""
    row_mapping = {
        46: 12,
        47: 13,
        48: 14,
        50: 16,
        51: 17,
        52: 18,
        54: 20,
        55: 21,
        72: 40,
        73: 32,
        75: 22,
        76: 23,
        77: 24
    }

    for target_row, source_row in row_mapping.items():
        sheet.cell(row=target_row, column=2).value = (
            "=" + build_monthly_sum(closing_columns, source_row)
        )


def set_monthly_currency_formulas(
    sheet: Worksheet,
    closing_columns: list[int],
    exchange_rate_reference: str
) -> None:
    """A havi forint- és euróösszegek összevonása az árfolyammal."""
    row_mapping = {
        58: 26,
        60: 28,
        61: 29,
        62: 30,
        63: 31,
        65: 34,
        66: 35,
        67: 36,
        68: 37,
        70: 38,
        71: 39
    }

    for target_row, source_row in row_mapping.items():
        huf_sum = build_monthly_sum(closing_columns, source_row)

        eur_sum = build_monthly_sum(closing_columns, source_row, currency_offset=1)

        sheet.cell(row=target_row, column=2).value = (
            f"={huf_sum}+({eur_sum}*{exchange_rate_reference})"
        )


def set_monthly_purchase_formula(
    sheet: Worksheet,
    closing_columns: list[int]
) -> None:
    """A havi beszerzés képletének beállítása a levonásokkal."""
    purchases = build_monthly_sum(closing_columns, 27)

    sheet["B59"] = (f"={purchases}-(B61+B62+B63+B66+B67+B68)")


def create_monthly_close(
    template_sheet: Worksheet,
    target_sheet: Worksheet,
    closing_columns: list[int],
    month: int,
    exchange_rate_reference: str
) -> None:
    """A havi záró fejléceinek és számított sorainak beállítása."""
    copy_monthly_close(
        template_sheet=template_sheet,
        target_sheet=target_sheet
    )

    target_sheet["B43"] = "Havi záró"
    target_sheet["B44"] = MONTH_NAMES[month - 1]
    target_sheet["B45"] = "Forint"

    set_monthly_basic_formulas(
        sheet=target_sheet,
        closing_columns=closing_columns
    )

    set_monthly_currency_formulas(
        sheet=target_sheet,
        closing_columns=closing_columns,
        exchange_rate_reference=exchange_rate_reference
    )

    set_monthly_purchase_formula(
        sheet=target_sheet,
        closing_columns=closing_columns
    )


# Árfolyam- és havi munkalapok


def ensure_exchange_rate_sheet(
    workbook: WorkbookType
) -> Worksheet:
    """Az Euro munkalap lekérése vagy létrehozása."""
    if "Euro" in workbook.sheetnames:
        return workbook["Euro"]

    sheet = workbook.create_sheet(title="Euro")

    sheet["A1"] = "Hónap"
    sheet["B1"] = "EUR/HUF árfolyam"

    sheet.column_dimensions["A"].width = 18
    sheet.column_dimensions["B"].width = 22

    return sheet


def ensure_month_exchange_rate(
    workbook: WorkbookType,
    month: int
) -> str:
    """A havi árfolyam helyének előkészítése; cellahivatkozást ad vissza."""
    sheet = ensure_exchange_rate_sheet(workbook)
    row = month + 1

    sheet.cell(row=row, column=1).value = MONTH_NAMES[month - 1]
    # A korábban megadott árfolyam értékét nem írjuk felül.
    sheet.cell(row=row, column=2).number_format = "0.00"

    return f"'Euro'!$B${row}"


def create_month_sheet(
    workbook: WorkbookType,
    template_sheet: Worksheet,
    year: int,
    month: int,
    exchange_rate_reference: str
) -> Worksheet:
    """Egy új havi munkalap létrehozása a heti és havi blokkokkal."""
    month_name = MONTH_NAMES[month - 1]

    if month_name in workbook.sheetnames:
        raise ValueError(
            f"A(z) {month_name} munkalap már létezik."
        )

    sheet = workbook.create_sheet(title=month_name)

    closing_columns = create_month_weeks(
        template_sheet=template_sheet,
        target_sheet=sheet,
        year=year,
        month=month
    )

    create_monthly_close(
        template_sheet=template_sheet,
        target_sheet=sheet,
        closing_columns=closing_columns,
        month=month,
        exchange_rate_reference=exchange_rate_reference
    )

    return sheet


def ensure_month_sheet(
    workbook: WorkbookType,
    template_sheet: Worksheet,
    year: int,
    month: int
) -> Worksheet:
    """A meglévő havi munkalap lekérése, szükség esetén létrehozása."""
    month_name = MONTH_NAMES[month - 1]

    if month_name in workbook.sheetnames:
        return workbook[month_name]

    exchange_rate_reference = ensure_month_exchange_rate(
        workbook=workbook,
        month=month
    )

    return create_month_sheet(
        workbook=workbook,
        template_sheet=template_sheet,
        year=year,
        month=month,
        exchange_rate_reference=exchange_rate_reference
    )


# Összesítő munkalap


def ensure_summary_sheet(
    workbook: WorkbookType,
    template_sheet: Worksheet
) -> Worksheet:
    """Az Összesítő lekérése vagy a sorfeliratainak létrehozása."""
    sheet_name = "Összesítő"

    if sheet_name in workbook.sheetnames:
        return workbook[sheet_name]

    sheet = workbook.create_sheet(title=sheet_name, index=0)

    copy_block(
        source_sheet=template_sheet,
        target_sheet=sheet,
        min_row=MONTHLY_CLOSE_START_ROW,
        max_row=MONTHLY_CLOSE_END_ROW,
        min_column=1,
        max_column=1,
        target_start_row=1,
        target_start_column=1
    )

    return sheet


def ensure_summary_month(
    workbook: WorkbookType,
    template_sheet: Worksheet,
    month: int
) -> None:
    """Egy hónap összesítő oszlopának kitöltése a havi záró hivatkozásaival."""
    month_name = MONTH_NAMES[month - 1]

    if month_name not in workbook.sheetnames:
        raise ValueError(
            f"A(z) {month_name} munkalap még nem létezik."
        )

    month_sheet = workbook[month_name]
    summary = ensure_summary_sheet(workbook, template_sheet)
    target_column = month + 1

    column_letter = get_column_letter(target_column)
    summary.column_dimensions[column_letter].width = 22

    for source_row in range(
        MONTHLY_CLOSE_START_ROW,
        MONTHLY_CLOSE_END_ROW + 1
    ):
        target_row = source_row - MONTHLY_CLOSE_START_ROW + 1

        copy_cell(
            source_sheet=month_sheet,
            target_sheet=summary,
            source_row=source_row,
            source_column=2,
            target_row=target_row,
            target_column=target_column,
            translate_formulas=False
        )

        target = summary.cell(
            row=target_row,
            column=target_column
        )

        if source_row <= 45:
            continue

        label = month_sheet.cell(row=source_row, column=1).value
        source_value = month_sheet.cell(row=source_row, column=2).value

        if label is None and source_value is None:
            target.value = None
            continue

        # Az üres kézi cellák az összesítőben is üresek maradnak.
        reference = f"'{month_name}'!B{source_row}"
        target.value = f'=IF({reference}="","",{reference})'


# Éves munkafüzet létrehozása és bővítése


def ensure_months_through(
    workbook: WorkbookType,
    template_sheet: Worksheet,
    year: int,
    last_month: int
) -> None:
    """Munkalapok és összesítő oszlopok biztosítása januártól a megadott hónapig."""
    if not 1 <= last_month <= 12:
        raise ValueError("Az utolsó hónap sorszáma 1 és 12 közötti lehet.")

    for month in range(1, last_month + 1):
        ensure_month_sheet(
            workbook=workbook,
            template_sheet=template_sheet,
            year=year,
            month=month
        )

        ensure_summary_month(
            workbook=workbook,
            template_sheet=template_sheet,
            month=month
        )


def create_year_workbook(
    template_sheet: Worksheet,
    year: int,
    last_month: int
) -> WorkbookType:
    """Új éves munkafüzet létrehozása a megadott hónapig."""
    if not 1 <= year <= 9999:
        raise ValueError("Az évszám 1 és 9999 közötti lehet.")

    workbook = Workbook()
    workbook.remove(workbook.active)

    ensure_months_through(
        workbook=workbook,
        template_sheet=template_sheet,
        year=year,
        last_month=last_month
    )

    return workbook
