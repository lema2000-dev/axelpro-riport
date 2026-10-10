from datetime import date
from decimal import Decimal
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

from axelpro_riport.models import (
    Movement,
    DailyTotals,
    MonthlyTotals,
    ImportResult,
    DailyExchangeRate,
)
from axelpro_riport.excel_documents import add_processed_documents
from axelpro_riport.excel_files import get_year_workbook_path
from axelpro_riport.excel_documents import read_processed_documents
from axelpro_riport.report_service import get_movement_years
from axelpro_riport.report_service import collect_processed_documents
from axelpro_riport.report_service import load_year_workbooks
from axelpro_riport.report_service import apply_import_result
from axelpro_riport.report_service import save_year_workbooks
from axelpro_riport.rate_logic import find_missing_exchange_rate_dates
from axelpro_riport.rate_logic import fill_missing_exchange_rates
from axelpro_riport.report_service import apply_exchange_rates


def test_get_movement_years() -> None:
    movements = [
        Movement(
            movement_date=day,
            transaction_type="Számla",
            document_number=document_number,
            group="új alkatrész",
            item_code="TESZT",
            net_value=Decimal("1000"),
            currency="HUF",
        )
        for day, document_number in [
            (date(2027, 1, 4), "SZ-2027/001"),
            (date(2026, 12, 30), "SZ-2026/001"),
            (date(2026, 12, 31), "SZ-2026/002"),
        ]
    ]

    assert get_movement_years(movements) == [2026, 2027]
    assert get_movement_years([]) == []
    print("Az árumozgásokhoz tartozó évek helyesen lettek meghatározva.")


def test_collect_processed_documents() -> None:
    workbook_2026 = Workbook()
    workbook_2027 = Workbook()

    try:
        # Azonos bizonylatszám, különböző év: két külün kulcs.
        add_processed_documents(workbook_2026, {(2026, "SZ-001")})
        add_processed_documents(workbook_2027, {(2027, "SZ-001")})

        workbooks = {
            2026: workbook_2026,
            2027: workbook_2027,
        }

        assert collect_processed_documents(workbooks) == {
            (2026, "SZ-001"),
            (2027, "SZ-001"),
        }
        assert collect_processed_documents({}) == set()
    finally:
        workbook_2026.close()
        workbook_2027.close()

    print("A feldolgozott dokumentumok helyesen lettek összegyűjtve.")


def test_load_year_workbooks() -> None:
    template_workbook = Workbook()
    template = template_workbook.active

    movements = [
        Movement(
            movement_date=day,
            transaction_type="Számla",
            document_number=document_number,
            group="Új alkatrész",
            item_code="TESZT",
            net_value=Decimal("1000"),
            currency="HUF",
        )
        for day, document_number in [
            (date(2026, 12, 30), "SZ-2026/001"),
            (date(2027, 2, 1), "SZ-2027/001"),
        ]
    ]

    try:
        with TemporaryDirectory() as directory:
            # Egy már létező riport megőrzendő adattal.
            existing = Workbook()
            try:
                existing.active.title = "Január"
                existing["Január"]["B8"] = 1234
                existing.save(get_year_workbook_path(directory, 2026))
            finally:
                existing.close()

            workbooks = load_year_workbooks(
                output_directory=directory,
                template_sheet=template,
                movements=movements,
            )

            try:
                assert set(workbooks) == {2026, 2027}

                # A meglévő riport adata megmaradt.
                assert workbooks[2026]["Január"]["B8"].value == 1234

                # Az új riport az érintett legutolsó hónapig készül.
                assert "Január" in workbooks[2027].sheetnames
                assert "Február" in workbooks[2027].sheetnames
                assert "Március" not in workbooks[2027].sheetnames

                # Az új munkafüzet még csak a memóriában létezik.
                assert not get_year_workbook_path(directory, 2027).exists()

            finally:
                for workbook in workbooks.values():
                    workbook.close()

            assert (
                load_year_workbooks(
                    output_directory=directory,
                    template_sheet=template,
                    movements=[],
                )
                == {}
            )

    finally:
        template_workbook.close()

    print(
        "Az árumozgásokhoz tartozó éves munkafüzetek helyesen lettek betöltve vagy létrehozva."
    )


def test_apply_import_result() -> None:
    template_workbook = Workbook()
    template = template_workbook.active

    workbooks = {
        2026: Workbook(),
        2027: Workbook(),
    }

    result = ImportResult(
        daily_totals=[
            DailyTotals(
                movement_date=date(2026, 1, 5),
                used_parts=1000,
                new_parts=2000,
                labor=3000,
                workshop_used_parts=0,
                workshop_new_parts=0,
                purchases=0,
                shipping=0,
                trailer=0,
            ),
        ],
        monthly_totals=[
            MonthlyTotals(
                year=2027,
                month=1,
                sany=4000,
                diag=5000,
                kj=6000,
            ),
        ],
        new_documents={
            (2026, "SZ-2026/001"),
            (2027, "SZ-2027/001"),
        },
    )

    try:
        # Hiányzó év esetén még egyik riportot sem módosíthatja.
        try:
            apply_import_result(
                workbooks={2026: workbooks[2026]},
                template_sheet=template,
                result=result,
            )
        except ValueError:
            pass
        else:
            raise AssertionError("Nem jelezte a hiányzó éves riportot.")

        assert workbooks[2026].sheetnames == ["Sheet"]
        assert read_processed_documents(workbooks[2026]) == set()

        apply_import_result(
            workbooks=workbooks,
            template_sheet=template,
            result=result,
        )

        # Január 5. hétfő: az első napi oszloppár B–C.
        sheet_2026 = workbooks[2026]["Január"]

        assert sheet_2026["J8"].value == 1000
        assert sheet_2026["J9"].value == 2000
        assert sheet_2026["J10"].value == 3000

        sheet_2027 = workbooks[2027]["Január"]
        assert sheet_2027["B82"].value == 4000
        assert sheet_2027["B83"].value == 5000
        assert sheet_2027["B84"].value == 6000

        assert read_processed_documents(workbooks[2026]) == {
            (2026, "SZ-2026/001"),
        }
        assert read_processed_documents(workbooks[2027]) == {
            (2027, "SZ-2027/001"),
        }

    finally:
        for workbook in workbooks.values():
            workbook.close()
        template_workbook.close()

    print(
        "Az új összegek és bizonylatkulcsok helyesen lettek beírva az éves riportokba."
    )


def test_save_year_workbooks() -> None:
    workbooks = {
        2026: Workbook(),
        2027: Workbook(),
    }

    result = ImportResult(
        daily_totals=[],
        monthly_totals=[],
        new_documents={(2026, "SZ-2026/001")},
    )

    try:
        workbooks[2026].active["B8"] = 1000
        add_processed_documents(
            workbooks[2026],
            result.new_documents,
        )

        with TemporaryDirectory() as directory:
            path = get_year_workbook_path(directory, 2026)
            untouched_path = get_year_workbook_path(directory, 2027)

            saved = save_year_workbooks(
                output_directory=directory,
                workbooks=workbooks,
                result=result,
            )

            assert saved is None
            assert path.exists()
            assert not untouched_path.exists()

            loaded = load_workbook(path, data_only=False)
            try:
                assert loaded.active["B8"].value == 1000
                assert read_processed_documents(loaded) == {
                    (2026, "SZ-2026/001"),
                }
            finally:
                loaded.close()

            # Második mentés: a riport lecserélődik, másolat nem készül.
            workbooks[2026].active["B8"] = 2500

            saved = save_year_workbooks(
                output_directory=directory,
                workbooks=workbooks,
                result=result,
            )

            assert saved is None
            assert not (path.parent / "backups").exists()

            loaded = load_workbook(path)
            try:
                assert loaded.active["B8"].value == 2500
            finally:
                loaded.close()

            # Új bizonylat nélkül nincs mentés.
            previous_content = path.read_bytes()
            workbooks[2026].active["B8"] = 9000

            empty_result = ImportResult(
                daily_totals=[],
                monthly_totals=[],
                new_documents=set(),
            )

            assert (
                save_year_workbooks(
                    output_directory=directory,
                    workbooks=workbooks,
                    result=empty_result,
                )
                is None
            )

            assert path.read_bytes() == previous_content

    finally:
        for workbook in workbooks.values():
            workbook.close()

    print(
        "Az új bizonylatokkal frissített éves riportok helyesen lettek mentve, a riportok ideiglenes fájlon keresztül frissültek."
    )


def test_missing_exchange_rates() -> None:
    friday = date(2026, 10, 9)
    sunday = date(2026, 10, 11)
    earlier_day = date(2026, 10, 8)

    movements = [
        Movement(
            movement_date=day,
            transaction_type="Számla",
            document_number=number,
            group="Új alkatrész",
            item_code="TESZT",
            net_value=Decimal("100"),
            currency=currency,
        )
        for day, number, currency in [
            (earlier_day, "SZ-001", "EUR"),
            (friday, "SZ-002", "EUR"),
            (sunday, "SZ-003", "EUR"),
            (sunday, "SZ-004", "EUR"),
            (date(2026, 10, 12), "SZ-005", "HUF"),
        ]
    ]

    rates = {
        friday: DailyExchangeRate(
            application_date=friday,
            source_date=friday,
            rate=Decimal("395.12"),
        ),
    }

    missing = find_missing_exchange_rate_dates(movements, rates)

    # Az ismétlődő nap egyszer szerepel, a HUF-tétel nem igényel árfolyamot.
    assert missing == [earlier_day, sunday]

    completed = fill_missing_exchange_rates(missing, rates)

    assert completed[sunday] == DailyExchangeRate(
        application_date=sunday,
        source_date=friday,
        rate=Decimal("395.12"),
    )

    # Október 8-hoz nem használhatjuk a későbbi, október 9-i árfolyamot.
    assert earlier_day not in completed
    assert find_missing_exchange_rate_dates(movements, completed) == [earlier_day]

    # Az eredeti szótár változatlan.
    assert set(rates) == {friday}

    # Egy helyettesített rekord továbbra is az eredeti forrásdátumot őrzi.
    monday = date(2026, 10, 12)

    extended = fill_missing_exchange_rates(
        [monday],
        {sunday: completed[sunday]},
    )

    assert extended[monday].source_date == friday
    assert extended[monday].rate == Decimal("395.12")

    assert fill_missing_exchange_rates([sunday], {}) == {}

    print(
        "A hiányzó árfolyamok kiegészítése a legutóbbi ismert forrásdátummal helyesen történt."
    )


def test_apply_exchange_rates() -> None:
    workbooks = {
        2026: Workbook(),
        2027: Workbook(),
    }

    day = date(2026, 10, 9)
    rates = {
        day: DailyExchangeRate(
            application_date=day,
            source_date=day,
            rate=Decimal("395.12"),
        ),
    }

    empty_result = ImportResult(
        daily_totals=[],
        monthly_totals=[],
        new_documents=set(),
    )

    try:
        changed_years = apply_exchange_rates(workbooks, rates)

        assert changed_years == {2026}
        assert "Euro" not in workbooks[2027].sheetnames

        euro = workbooks[2026]["Euro"]
        assert euro["B11"].value == 395.12
        assert euro["C11"].value == day

        # Változatlan adatok esetén nincs újabb módosult év.
        assert apply_exchange_rates(workbooks, rates) == set()

        with TemporaryDirectory() as directory:
            saved = save_year_workbooks(
                output_directory=directory,
                workbooks=workbooks,
                result=empty_result,
                additional_years=changed_years,
            )

            assert saved is None

            path = get_year_workbook_path(directory, 2026)
            assert path.exists()
            assert not get_year_workbook_path(directory, 2027).exists()

            loaded = load_workbook(path, data_only=False)
            try:
                assert loaded["Euro"]["B11"].value == 395.12
                assert loaded["Euro"]["C11"].value.date() == day
            finally:
                loaded.close()

        # A havi cella javítása önmagában is módosítás.
        euro["B11"] = 1

        assert apply_exchange_rates(workbooks, rates) == {2026}
        assert euro["B11"].value == 395.12

    finally:
        for workbook in workbooks.values():
            workbook.close()

    print(
        "A havi árfolyamok helyesen lettek beírva az éves riportokba, a változások mentése és visszaállítása is megfelelően működött."
    )


def run_all_tests() -> None:
    test_get_movement_years()
    test_collect_processed_documents()
    test_load_year_workbooks()
    test_apply_import_result()
    test_save_year_workbooks()
    test_missing_exchange_rates()
    test_apply_exchange_rates()


if __name__ == "__main__":
    run_all_tests()
