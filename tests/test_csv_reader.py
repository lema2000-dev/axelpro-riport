from pathlib import Path

from axelpro_riport.csv_reader import read_movements


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SAMPLES_DIR = PROJECT_ROOT / "samples"

SAMPLE_FILES = (
    "napokon_at.csv",
    "honapokon_at.csv",
    "eveken_at.csv",
)


def test_sample_files() -> None:
    for filename in SAMPLE_FILES:
        movements = read_movements(SAMPLES_DIR / filename)

        assert len(movements) == 102, (filename, len(movements))

        currencies = {movement.currency for movement in movements}
        assert currencies == {"HUF", "EUR"}, (filename, currencies)

        assert all(
            movement.transaction_type.casefold() != "leltár"
            for movement in movements
        ), filename

        print(f"Sikeres beolvasási teszt: {filename}")


if __name__ == "__main__":
    test_sample_files()