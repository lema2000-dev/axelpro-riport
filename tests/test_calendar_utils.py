from datetime import date

from axelpro_riport.calendar_utils import get_month_weeks


def test_month_boundary() -> None:
    september = get_month_weeks(2026, 9)
    october = get_month_weeks(2026, 10)

    assert september[-1] == [
        date(2026, 9, 28),
        date(2026, 9, 29),
        date(2026, 9, 30)
    ]

    assert october[0] == [
        date(2026, 10, 1),
        date(2026, 10, 2),
        date(2026, 10, 3),
        date(2026, 10, 4)
    ]

    assert september[-1][0].isocalendar().week == 40
    assert october[0][0].isocalendar().week == 40

    print("Sikeres hónaphatár teszt.")

def test_year_boundary() -> None:
    december = get_month_weeks(2026, 12)
    january = get_month_weeks(2027, 1)

    assert december[-1] == [
        date(2026, 12, 28),
        date(2026, 12, 29),
        date(2026, 12, 30),
        date(2026, 12, 31)
    ]

    assert january[0] == [
        date(2027, 1, 1),
        date(2027, 1, 2),
        date(2027, 1, 3)
    ]

    assert december[-1][0].isocalendar()[:2] == (2026, 53)
    assert january[0][0].isocalendar()[:2] == (2026, 53)

    print("Sikeres évhatár teszt.")


if __name__ == "__main__":
    test_month_boundary()
    test_year_boundary()