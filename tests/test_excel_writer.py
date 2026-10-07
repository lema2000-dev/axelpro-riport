from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from copy import copy

from datetime import date

from axelpro_riport.excel_writer import (
    copy_block,
    clear_input_values,
    create_week_block,
    create_year_workbook,
    set_week_dates,
    create_month_sheet,
    ensure_month_exchange_rate,
    ensure_month_sheet,
    ensure_summary_month,
    ensure_months_through,
)
from axelpro_riport.calendar_utils import get_month_weeks


def test_copy_block() -> None:
    workbook = Workbook()
    source = workbook.active
    source.title = "Forrás"
    target = workbook.create_sheet(title="Cél")

    source["A1"] = "Teszt"
    source["B2"] = 100
    source["B3"] = "=B2+$B$2"

    source["A1"].font = Font(bold=True)
    source["A1"].fill = PatternFill(fill_type="solid", fgColor="FFFF00")

    source["B2"].number_format = "#,##0"

    source.row_dimensions[1].height = 25
    source.column_dimensions["A"].width = 20
    source.merge_cells("A1:C1")

    copy_block(
        source_sheet=source,
        target_sheet=target,
        min_row=1,
        max_row=3,
        min_column=1,
        max_column=3,
        target_start_row=5,
        target_start_column=5
    )

    assert target["E5"].value == "Teszt"
    assert target["F6"].value == 100
    assert target["F7"].value == "=F6+$B$2"

    assert "E5:G5" in {str(merged_range) for merged_range in target.merged_cells.ranges}

    assert copy(target["E5"].font) == copy(source["A1"].font)
    assert copy(target["E5"].fill) == copy(source["A1"].fill)
    assert target["F6"].number_format == "#,##0"

    assert target.row_dimensions[5].height == 25
    assert target.column_dimensions["E"].width == 20

    workbook.close()
    print("Sikeres blokkmásolási teszt.")

def test_clear_input_values() -> None:
    workbook = Workbook()
    sheet = workbook.active

    sheet["A1"] = "Sorfelirat"
    sheet["B1"] = 150
    sheet.merge_cells("B1:C1")

    sheet["B2"] = "=B1*2"
    sheet["C2"] = 200
    sheet["B1"].font = Font(bold=True)

    clear_input_values(
        sheet=sheet,
        min_row=1,
        max_row=2,
        min_column=2,
        max_column=3
    )

    assert sheet["A1"].value == "Sorfelirat"
    assert sheet["B1"].value is None
    assert sheet["C2"].value is None
    assert sheet["B1"].font.bold is True
    assert "B1:C1" in {str(merged_range) for merged_range in sheet.merged_cells.ranges}

    workbook.close()
    print("Sikeres adatürítési teszt.")

def test_create_week_block() -> None:
    workbook = Workbook()
    template = workbook.active
    target = workbook.create_sheet(title="Cél")

    template["A8"] = "Bontott alkatrész sz."

    day_names = (
        "Hétfő",
        "Kedd",
        "Szerda",
        "Csütörtök",
        "Péntek",
        "Szombat"
    )

    for weekday, day_name in enumerate(day_names):
        column = 2 + weekday * 2

        template.cell(row=1, column=column).value = day_name
        template.cell(row=2, column=column).value = date(2026, 1, 1)
        template.cell(row=4, column=column).value = 100
        template.cell(row=8, column=column).value = 200

        for row in (1, 2, 4, 8):
            template.merge_cells(
                start_row=row,
                end_row=row,
                start_column=column,
                end_column=column + 1
            )

    template["N1"] = "Heti záró"
    template["N8"] = "=B8+D8+F8+H8+J8+L8"
    template["N26"] = "=B26+D26+F26+H26+J26+L26"
    template["O26"] = "=C26+E26+G26+I26+K26+M26"

    for row in (1, 2, 4, 8):
        template.merge_cells(
            start_row=row,
            end_row=row,
            start_column=14,
            end_column=15
        )

    first_width = create_week_block(
        template_sheet=template,
        target_sheet=target,
        week=get_month_weeks(2026, 10)[0],
        week_start_column=1,
        include_labels=True
    )

    assert first_width == 9
    assert target["A8"].value == "Bontott alkatrész sz."
    assert target["B1"].value == "Csütörtök"
    assert target["D1"].value == "Péntek"
    assert target["F1"].value == "Szombat"
    assert target["H1"].value == "Heti záró"

    assert target["B2"].value == date(2026, 10, 1)
    assert target["D2"].value == date(2026, 10, 2)
    assert target["F2"].value == date(2026, 10, 3)
    assert target["H2"].value == 40
    assert target["B4"].value is None
    assert target["B8"].value is None

    assert target["H8"].value == "=B8+D8+F8"
    assert target["H26"].value == "=B26+D26+F26"
    assert target["I26"].value == "=C26+E26+G26"

    next_column = 1 + first_width

    last_width = create_week_block(
        template_sheet=template,
        target_sheet=target,
        week=get_month_weeks(2026, 9)[-1],
        week_start_column=next_column,
        include_labels=False
    )

    assert last_width == 8
    assert target["J1"].value == "Hétfő"
    assert target["L1"].value == "Kedd"
    assert target["N1"].value == "Szerda"
    assert target["P1"].value == "Heti záró"

    assert target["J2"].value == date(2026, 9, 28)
    assert target["L2"].value == date(2026, 9, 29)
    assert target["N2"].value == date(2026, 9, 30)
    assert target["P2"].value == 40

    assert target["P8"].value == "=J8+L8+N8"
    assert target["P26"].value == "=J26+L26+N26"
    assert target["Q26"].value == "=K26+M26+O26"

    merged_ranges = {str(merged_range) for merged_range in target.merged_cells.ranges}

    assert "B8:C8" in merged_ranges
    assert "H8:I8" in merged_ranges
    assert "J8:K8" in merged_ranges
    assert "P8:Q8" in merged_ranges

    next_column = next_column + last_width

    sunday_width = create_week_block(
        template_sheet=template,
        target_sheet=target,
        week=[date(2026, 11, 1)],
        week_start_column=next_column,
        include_labels=False
    )

    assert sunday_width == 0
    assert target.max_column == next_column - 1

    workbook.close()
    print("Sikeres rövidített heti blokk teszt.")

def test_set_week_dates() -> None:
    workbook = Workbook()
    sheet = workbook.active
    week = get_month_weeks(2026, 10)[0]

    set_week_dates(
        sheet=sheet,
        week=week,
        week_start_column=1,
        include_labels=True
    )

    assert sheet["B2"].value == date(2026, 10, 1)
    assert sheet["D2"].value == date(2026, 10, 2)
    assert sheet["F2"].value == date(2026, 10, 3)
    assert sheet["H2"].value == 40

    set_week_dates(
        sheet=sheet,
        week=week,
        week_start_column=10,
        include_labels=False
    )

    assert sheet["J2"].value == date(2026, 10, 1)
    assert sheet["L2"].value == date(2026, 10, 2)
    assert sheet["N2"].value == date(2026, 10, 3)
    assert sheet["P2"].value == 40

    workbook.close()
    print("Sikeres heti dátumbeállítási teszt.")

def test_create_month_sheet() -> None:
    template_workbook = Workbook()
    template = template_workbook.active

    day_names = (
        "Hétfő",
        "Kedd",
        "Szerda",
        "Csütörtök",
        "Péntek",
        "Szombat"
    )

    template["A8"] = "Bontott alkatrész sz."

    for weekday, day_name in enumerate(day_names):
        column = 2 + weekday * 2
        template.cell(row=1, column=column).value = day_name

    template["N1"] = "Heti záró"
    template["N12"] = "=B12+D12+F12+H12+J12+L12"
    template["N26"] = "=B26+D26+F26+H26+J26+L26"
    template["O26"] = "=C26+E26+G26+I26+K26+M26"

    template["A43"] = "Dátum:"
    template["B43"] = "Havi záró"
    template["B44"] = "Szeptember"
    template["B45"] = "Forint"
    template["B53"] = "=B52+B48"
    template["B74"] = 123
    template["B78"] = "=B74*0.08"

    for row in range(43, 86):
        template.merge_cells(
            start_row=row,
            end_row=row,
            start_column=2,
            end_column=3
        )

    workbook = Workbook()
    workbook.remove(workbook.active)

    rate_reference = "'Euro'!$B$11"

    sheet = create_month_sheet(
        workbook=workbook,
        template_sheet=template,
        year=2026,
        month=10,
        exchange_rate_reference=rate_reference
    )

    assert sheet.title == "Október"
    assert sheet.max_column == 65

    assert sheet["B1"].value == "Csütörtök"
    assert sheet["B2"].value == date(2026, 10, 1)
    assert sheet["H2"].value == 40

    assert sheet["J1"].value == "Hétfő"
    assert sheet["J2"].value == date(2026, 10, 5)
    assert sheet["BJ2"].value == date(2026, 10, 31)
    assert sheet["BL2"].value == 44

    assert sheet["B43"].value == "Havi záró"
    assert sheet["B44"].value == "Október"
    assert sheet["B45"].value == "Forint"

    assert sheet["B46"].value == "=(H12+V12+AJ12+AX12+BL12)"
    assert sheet["B58"].value == "=(H26+V26+AJ26+AX26+BL26)+((I26+W26+AK26+AY26+BM26)*'Euro'!$B$11)"

    assert sheet["B59"].value == "=(H27+V27+AJ27+AX27+BL27)-(B61+B62+B63+B66+B67+B68)"

    assert sheet["B53"].value == "=B52+B48"
    assert sheet["B74"].value is None
    assert sheet["B78"].value == "=B74*0.08"

    assert "B46:C46" in {str(merged_range) for merged_range in sheet.merged_cells.ranges}

    try:
        create_month_sheet(
            workbook=workbook,
            template_sheet=template,
            year=2026,
            month=10,
            exchange_rate_reference=rate_reference
        )
    except ValueError:
        pass
    else:
        raise AssertionError(
            "A már létező havi munkalap létrehozása nem jelzett hibát."
        )

    workbook.close()
    template_workbook.close()
    print("Sikeres teljes havi munkalap létrehozási teszt.")

def test_ensure_month_sheet() -> None:
    template_workbook = Workbook()
    template = template_workbook.active

    template["N26"] = "=B26+D26+F26+H26+J26+L26"
    template["O26"] = "=C26+E26+G26+I26+K26+M26"

    workbook = Workbook()
    workbook.remove(workbook.active)

    sheet = ensure_month_sheet(
        workbook=workbook,
        template_sheet=template,
        year=2026,
        month=10,
    )

    assert set(workbook.sheetnames) == {"Euro", "Október"}
    assert workbook["Euro"]["A11"].value == "Október"
    assert workbook["Euro"]["B11"].value is None
    assert workbook["Euro"]["B11"].number_format == "0.00"

    assert sheet["B58"].value == (
        "=(H26+V26+AJ26+AX26+BL26)"
        "+((I26+W26+AK26+AY26+BM26)*'Euro'!$B$11)"
    )

    sheet["B8"] = 12345
    workbook["Euro"]["B11"] = 400.25

    repeated_sheet = ensure_month_sheet(
        workbook=workbook,
        template_sheet=template,
        year=2026,
        month=10,
    )

    assert repeated_sheet is sheet
    assert repeated_sheet["B8"].value == 12345
    assert workbook["Euro"]["B11"].value == 400.25
    assert len(workbook.sheetnames) == 2

    reference = ensure_month_exchange_rate(
        workbook=workbook,
        month=10,
    )

    assert reference == "'Euro'!$B$11"
    assert workbook["Euro"]["B11"].value == 400.25

    workbook.close()
    template_workbook.close()
    print("Sikeres havi munkalap- és árfolyammegőrzési teszt.")

def test_summary_month() -> None:
    template_workbook = Workbook()
    template = template_workbook.active
    template["A46"] = "Alkatrész értékesítés bontott sz.n.:"
    template["A74"] = "Havi óraszám:"

    workbook = Workbook()
    workbook.remove(workbook.active)

    month_sheet = ensure_month_sheet(
        workbook=workbook,
        template_sheet=template,
        year=2026,
        month=1,
    )

    ensure_summary_month(
        workbook=workbook,
        template_sheet=template,
        month=1,
    )

    summary = workbook["Összesítő"]

    assert workbook.sheetnames[0] == "Összesítő"
    assert summary["B2"].value == "Január"
    assert summary["A4"].value == template["A46"].value

    assert summary["B4"].value == (
        '=IF(\'Január\'!B46="","",\'Január\'!B46)'
    )
    assert summary["B32"].value == (
        '=IF(\'Január\'!B74="","",\'Január\'!B74)'
    )

    assert summary["B7"].value is None
    assert summary.column_dimensions["B"].width == 22

    month_sheet["B74"] = 150

    ensure_summary_month(
        workbook=workbook,
        template_sheet=template,
        month=1,
    )

    assert workbook["Összesítő"] is summary
    assert month_sheet["B74"].value == 150
    assert workbook.sheetnames.count("Összesítő") == 1

    try:
        ensure_summary_month(
            workbook=workbook,
            template_sheet=template,
            month=2,
        )
    except ValueError:
        pass
    else:
        raise AssertionError(
            "Hiányzó havi munkalap esetén nem jelzett hibát."
        )

    assert summary.max_column == 2

    workbook.close()
    template_workbook.close()
    print("Sikeres összesítőmunkalap-teszt.")

def test_year_workbook() -> None:
    template_workbook = Workbook()
    template = template_workbook.active

    workbook = create_year_workbook(
        template_sheet=template,
        year=2026,
        last_month=2,
    )

    assert workbook.sheetnames == [
        "Összesítő",
        "Euro",
        "Január",
        "Február",
    ]

    assert workbook["Összesítő"]["B2"].value == "Január"
    assert workbook["Összesítő"]["C2"].value == "Február"
    assert workbook["Összesítő"].max_column == 3

    assert workbook["Euro"]["A2"].value == "Január"
    assert workbook["Euro"]["A3"].value == "Február"

    january = workbook["Január"]
    january["B8"] = 12345
    workbook["Euro"]["B2"] = 400.25

    ensure_months_through(
        workbook=workbook,
        template_sheet=template,
        year=2026,
        last_month=3,
    )

    assert workbook.sheetnames == [
        "Összesítő",
        "Euro",
        "Január",
        "Február",
        "Március",
    ]

    assert workbook["Január"] is january
    assert january["B8"].value == 12345
    assert workbook["Euro"]["B2"].value == 400.25

    assert workbook["Összesítő"]["D2"].value == "Március"
    assert workbook["Euro"]["A4"].value == "Március"

    new_year = create_year_workbook(
        template_sheet=template,
        year=2027,
        last_month=1,
    )

    assert new_year.sheetnames == [
        "Összesítő",
        "Euro",
        "Január",
    ]

    assert new_year["Összesítő"].max_column == 2
    assert new_year["Január"]["B8"].value is None
    assert new_year["Euro"]["B2"].value is None

    workbook.close()
    new_year.close()
    template_workbook.close()
    print("Sikeres éves munkafüzet- és hónapbővítési teszt.")


if __name__ == "__main__":
    test_copy_block()
    test_clear_input_values()
    test_create_week_block()
    test_set_week_dates()
    test_create_month_sheet()
    test_ensure_month_sheet()
    test_summary_month()
    test_year_workbook()