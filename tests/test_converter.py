from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

import pandas as pd

from converter import DataValidationError, transform_excel


def write_source(path: Path, rows: list[dict]) -> None:
    pd.DataFrame(rows).to_excel(path, index=False, engine="openpyxl")


class ConverterTests(unittest.TestCase):
    def test_transforms_and_anonymises_valid_export(self) -> None:
        with TemporaryDirectory() as directory:
            workdir = Path(directory)
            source = workdir / "source.xlsx"
            output_dir = workdir / "out"
            rows = [
                {
                    "Status": 11,
                    "Rsv Number": "001",
                    "Cancel Number": "",
                    "Rate": 150,
                    "Rsv Date1": "2026-06-11 09:30:00",
                    "Check In": "2026-07-01",
                    "Check Out": "2026-07-03",
                    "Name": "Smith",
                    "Firstname": "Anna",
                    "Address1": "Main Street 1",
                    "Email": "anna@example.com",
                    "Hotel City": "Zurich",
                    "City": "Bern, Switzerland",
                },
                {
                    "Status": 13,
                    "Rsv Number": "001",
                    "Cancel Number": "C-1",
                    "Rate": 150,
                    "Rsv Date1": "2026-06-12 10:15:00",
                    "Check In": "2026-07-01",
                    "Check Out": "2026-07-03",
                    "Name": "Smith",
                    "Firstname": "Anna",
                    "Address1": "Main Street 1",
                    "Email": "anna@example.com",
                    "Hotel City": "Zurich",
                    "City": "Bern, Switzerland",
                },
            ]
            write_source(source, rows)

            output_path = transform_excel(source, output_dir, anonymization_salt="test-salt")
            result = pd.read_excel(output_path, engine="openpyxl")

            self.assertEqual(output_path.name, "promodata.xlsx")
            self.assertEqual(result.loc[0, "Cancel Number"], "cancelled")
            self.assertEqual(result.loc[0, "NotTravelled"], 150)
            self.assertTrue(pd.isna(result.loc[0, "travelled"]))
            self.assertEqual(result.loc[0, "Booking Hour"], 9)
            self.assertEqual(result.loc[0, "Hotel City"], "Zurich, Switzerland")
            self.assertEqual(result.loc[0, "City"], "Bern, Switzerland")
            self.assertNotIn("Email", result.columns)
            self.assertNotIn("Address1", result.columns)
            self.assertIn("AnonID", result.columns)
            self.assertNotEqual(result.loc[0, "Name"], "Smith")
            self.assertEqual(result.loc[0, "Name"], result.loc[1, "Name"])
            self.assertEqual(result.loc[0, "AnonID"], result.loc[1, "AnonID"])

    def test_uses_unique_output_name_when_file_exists(self) -> None:
        with TemporaryDirectory() as directory:
            workdir = Path(directory)
            source = workdir / "source.xlsx"
            output_dir = workdir / "out"
            output_dir.mkdir()
            (output_dir / "promodata.xlsx").write_text("existing", encoding="utf-8")
            write_source(
                source,
                [
                    {
                        "Status": 11,
                        "Rsv Number": "002",
                        "Cancel Number": "",
                        "Rate": 100,
                        "Rsv Date1": "2026-06-11",
                        "Check In": "2026-07-01",
                        "Check Out": "2026-07-03",
                    }
                ],
            )

            output_path = transform_excel(source, output_dir)

            self.assertNotEqual(output_path.name, "promodata.xlsx")
            self.assertTrue(output_path.name.startswith("promodata_"))

    def test_rejects_missing_required_columns(self) -> None:
        with TemporaryDirectory() as directory:
            source = Path(directory) / "source.xlsx"
            write_source(source, [{"Status": 11}])

            with self.assertRaises(DataValidationError):
                transform_excel(source, Path(directory) / "out")

    def test_rejects_invalid_dates(self) -> None:
        with TemporaryDirectory() as directory:
            source = Path(directory) / "source.xlsx"
            write_source(
                source,
                [
                    {
                        "Status": 11,
                        "Rsv Number": "003",
                        "Cancel Number": "",
                        "Rate": 100,
                        "Rsv Date1": "not a date",
                        "Check In": "2026-07-01",
                        "Check Out": "2026-07-03",
                    }
                ],
            )

            with self.assertRaises(DataValidationError):
                transform_excel(source, Path(directory) / "out")


if __name__ == "__main__":
    unittest.main()
