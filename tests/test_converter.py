from pathlib import Path
from tempfile import TemporaryDirectory
import os
import unittest
from unittest.mock import patch
import urllib.request

import pandas as pd

from converter import AzureMapsGeocoder, DataValidationError, transform_excel


def write_source(path: Path, rows: list[dict]) -> None:
    pd.DataFrame(rows).to_excel(path, index=False, engine="openpyxl")


class FakeGeocoder:
    def __init__(self) -> None:
        self.calls: list[tuple[str, str]] = []

    def geocode(self, city: str, country: str) -> tuple[float, float] | None:
        self.calls.append((city, country))
        return {
            ("Zurich", "Switzerland"): (47.3769, 8.5417),
            ("Bern", "Switzerland"): (46.9480, 7.4474),
        }.get((city, country))


class FakeResponse:
    def __init__(self, payload: str) -> None:
        self.payload = payload.encode("utf-8")

    def __enter__(self) -> "FakeResponse":
        return self

    def __exit__(self, exc_type: object, exc_value: object, traceback: object) -> None:
        return None

    def read(self) -> bytes:
        return self.payload


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

            geocoder = FakeGeocoder()
            output_path = transform_excel(
                source,
                output_dir,
                anonymization_salt="test-salt",
                geocoder=geocoder,
            )
            result = pd.read_excel(output_path, engine="openpyxl")

            self.assertEqual(output_path.name, "promodata.xlsx")
            self.assertEqual(result.loc[0, "Cancel Number"], "cancelled")
            self.assertNotEqual(result.loc[0, "Rsv Number"], "001")
            self.assertEqual(result.loc[0, "Rsv Number"], result.loc[1, "Rsv Number"])
            self.assertNotEqual(result.loc[1, "Cancel Number"], "C-1")
            self.assertEqual(result.loc[0, "NotTravelled"], 150)
            self.assertTrue(pd.isna(result.loc[0, "travelled"]))
            self.assertEqual(result.loc[0, "Booking Hour"], 9)
            self.assertEqual(result.loc[0, "Hotel City"], "Zurich")
            self.assertEqual(result.loc[0, "Hotel Land"], "Switzerland")
            self.assertAlmostEqual(result.loc[0, "Hotel Latitude"], 47.3769)
            self.assertAlmostEqual(result.loc[0, "Hotel Longitude"], 8.5417)
            self.assertEqual(result.loc[0, "City"], "Bern")
            self.assertEqual(result.loc[0, "Land"], "Switzerland")
            self.assertAlmostEqual(result.loc[0, "Latitude"], 46.9480)
            self.assertAlmostEqual(result.loc[0, "Longitude"], 7.4474)
            columns = result.columns.tolist()
            hotel_city_index = columns.index("Hotel City")
            city_index = columns.index("City")
            self.assertEqual(
                columns[hotel_city_index : hotel_city_index + 4],
                ["Hotel City", "Hotel Land", "Hotel Latitude", "Hotel Longitude"],
            )
            self.assertEqual(
                columns[city_index : city_index + 4],
                ["City", "Land", "Latitude", "Longitude"],
            )
            self.assertNotIn("Email", result.columns)
            self.assertNotIn("Address1", result.columns)
            self.assertIn("AnonID", result.columns)
            self.assertNotEqual(result.loc[0, "Name"], "Smith")
            self.assertEqual(result.loc[0, "Name"], result.loc[1, "Name"])
            self.assertEqual(result.loc[0, "AnonID"], result.loc[1, "AnonID"])
            self.assertEqual(geocoder.calls, [("Zurich", "Switzerland"), ("Bern", "Switzerland")])

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

    def test_requires_azure_maps_key_when_locations_need_geocoding(self) -> None:
        with TemporaryDirectory() as directory:
            source = Path(directory) / "source.xlsx"
            write_source(
                source,
                [
                    {
                        "Status": 11,
                        "Rsv Number": "004",
                        "Cancel Number": "",
                        "Rate": 100,
                        "Rsv Date1": "2026-06-11",
                        "Check In": "2026-07-01",
                        "Check Out": "2026-07-03",
                        "City": "Bern",
                    }
                ],
            )

            with patch.dict(
                os.environ,
                {"AZURE_MAPS_SUBSCRIPTION_KEY": "", "AZURE_MAPS_KEY": ""},
                clear=False,
            ):
                with self.assertRaisesRegex(DataValidationError, "AZURE_MAPS_SUBSCRIPTION_KEY"):
                    transform_excel(source, Path(directory) / "out")

    def test_uses_passed_azure_maps_key_for_geocoding(self) -> None:
        with TemporaryDirectory() as directory:
            workdir = Path(directory)
            source = workdir / "source.xlsx"
            output_dir = workdir / "out"
            write_source(
                source,
                [
                    {
                        "Status": 11,
                        "Rsv Number": "005",
                        "Cancel Number": "",
                        "Rate": 100,
                        "Rsv Date1": "2026-06-11",
                        "Check In": "2026-07-01",
                        "Check Out": "2026-07-03",
                        "City": "Bern",
                    }
                ],
            )
            requests: list[tuple[str, str, str]] = []

            def fake_fetch(self: AzureMapsGeocoder, city: str, country: str) -> tuple[float, float]:
                requests.append((self.subscription_key, city, country))
                return 46.9480, 7.4474

            with patch.dict(
                os.environ,
                {"AZURE_MAPS_SUBSCRIPTION_KEY": "", "AZURE_MAPS_KEY": ""},
                clear=False,
            ):
                with patch.object(AzureMapsGeocoder, "_fetch_coordinates", fake_fetch):
                    output_path = transform_excel(
                        source,
                        output_dir,
                        azure_maps_subscription_key="ui-key",
                    )

            result = pd.read_excel(output_path, engine="openpyxl")

            self.assertEqual(requests, [("ui-key", "Bern", "Switzerland")])
            self.assertAlmostEqual(result.loc[0, "Latitude"], 46.9480)
            self.assertAlmostEqual(result.loc[0, "Longitude"], 7.4474)
            self.assertTrue((output_dir / ".azure_maps_geocode_cache.json").exists())

    def test_geocoder_retries_connection_reset(self) -> None:
        payload = (
            '{"features":[{"geometry":{"coordinates":[8.7721,47.2523]}}]}'
        )

        with TemporaryDirectory() as directory:
            geocoder = AzureMapsGeocoder(
                "test-key",
                Path(directory) / "cache.json",
                retry_delay_seconds=0,
            )

            with patch.object(
                urllib.request,
                "urlopen",
                side_effect=[ConnectionResetError(10054, "connection reset"), FakeResponse(payload)],
            ) as urlopen_mock:
                coordinates = geocoder.geocode("Hombrechtikon", "Switzerland")

            request = urlopen_mock.call_args_list[1].args[0]

            self.assertEqual(urlopen_mock.call_count, 2)
            self.assertEqual(coordinates, (47.2523, 8.7721))
            self.assertEqual(request.headers["User-agent"], "STC-Hotel-Data-PowerBI-Converter/1.0")
            self.assertEqual(request.headers["Connection"], "close")

    def test_rejects_locations_without_coordinates(self) -> None:
        with TemporaryDirectory() as directory:
            source = Path(directory) / "source.xlsx"
            write_source(
                source,
                [
                    {
                        "Status": 11,
                        "Rsv Number": "005",
                        "Cancel Number": "",
                        "Rate": 100,
                        "Rsv Date1": "2026-06-11",
                        "Check In": "2026-07-01",
                        "Check Out": "2026-07-03",
                        "City": "Atlantis",
                    }
                ],
            )

            with self.assertRaisesRegex(DataValidationError, "Atlantis, Switzerland"):
                transform_excel(source, Path(directory) / "out", geocoder=FakeGeocoder())

    def test_reports_progress_messages(self) -> None:
        with TemporaryDirectory() as directory:
            workdir = Path(directory)
            source = workdir / "source.xlsx"
            output_dir = workdir / "out"
            messages: list[str] = []
            write_source(
                source,
                [
                    {
                        "Status": 11,
                        "Rsv Number": "006",
                        "Cancel Number": "",
                        "Rate": 100,
                        "Rsv Date1": "2026-06-11",
                        "Check In": "2026-07-01",
                        "Check Out": "2026-07-03",
                        "City": "Bern",
                    }
                ],
            )

            transform_excel(
                source,
                output_dir,
                geocoder=FakeGeocoder(),
                progress_callback=messages.append,
            )

            self.assertIn("Reading source workbook", messages)
            self.assertIn("Normalising city/country and coordinates", messages)
            self.assertIn("Geocoding Bern, Switzerland", messages)
            self.assertIn("Conversion complete", messages)


if __name__ == "__main__":
    unittest.main()
