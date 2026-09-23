from __future__ import annotations

import hashlib
import json
import os
import time
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable
import urllib.error
import urllib.parse
import urllib.request

import pandas as pd


OUTPUT_FILENAME = "promodata.xlsx"
ANONYMIZATION_SALT_ENV = "HOTEL_PROMO_ANON_SALT"
DEFAULT_ANONYMIZATION_SALT = "hotel-promo-converter-v1"
AZURE_MAPS_SUBSCRIPTION_KEY_ENV = "AZURE_MAPS_SUBSCRIPTION_KEY"
AZURE_MAPS_KEY_ENV = "AZURE_MAPS_KEY"
AZURE_MAPS_CLIENT_ID_ENV = "AZURE_MAPS_CLIENT_ID"
AZURE_MAPS_ENDPOINT = "https://atlas.microsoft.com/geocode"
AZURE_MAPS_API_VERSION = "2025-01-01"
GEOCODE_CACHE_FILENAME = ".azure_maps_geocode_cache.json"
GEOCODE_RETRY_COUNT = 3
GEOCODE_RETRY_DELAY_SECONDS = 1.5
RETRIABLE_HTTP_STATUS_CODES = {408, 429, 500, 502, 503, 504}

REQUIRED_COLUMNS = (
    "Status",
    "Rsv Number",
    "Cancel Number",
    "Rate",
    "Rsv Date1",
    "Check In",
    "Check Out",
)
TEXT_COLUMNS = ("Rsv Number", "Cancel Number", "Address1", "Name", "Firstname")
PERSONAL_COLUMNS = ("Name", "Firstname")
IDENTIFIER_COLUMNS = ("Rsv Number",)
DROP_COLUMNS = ("Rsv Date1", "Address2", "Email", "Remark")
DEFAULT_COUNTRY = "Switzerland"
LOCATION_COLUMN_GROUPS = {
    "Hotel City": ("Hotel Land", "Hotel Latitude", "Hotel Longitude"),
    "City": ("Land", "Latitude", "Longitude"),
}
ProgressCallback = Callable[[str], None]


class DataValidationError(RuntimeError):
    """Raised when the source workbook does not match the expected export format."""


class AzureMapsGeocoder:
    def __init__(
        self,
        subscription_key: str,
        cache_path: Path,
        *,
        endpoint: str = AZURE_MAPS_ENDPOINT,
        api_version: str = AZURE_MAPS_API_VERSION,
        client_id: str | None = None,
        timeout_seconds: int = 15,
        retry_count: int = GEOCODE_RETRY_COUNT,
        retry_delay_seconds: float = GEOCODE_RETRY_DELAY_SECONDS,
        progress_callback: ProgressCallback | None = None,
    ) -> None:
        self.subscription_key = subscription_key
        self.cache_path = cache_path
        self.endpoint = endpoint
        self.api_version = api_version
        self.client_id = client_id
        self.timeout_seconds = timeout_seconds
        self.retry_count = retry_count
        self.retry_delay_seconds = retry_delay_seconds
        self.progress_callback = progress_callback
        self._cache = self._load_cache()

    def geocode(self, city: str, country: str) -> tuple[float, float] | None:
        cache_key = _location_cache_key(city, country)

        if cache_key in self._cache:
            self._progress(f"Using cached coordinates for {city}, {country}")
            cached = self._cache[cache_key]
            return _coordinates_from_cache(cached) if cached else None

        coordinates = self._fetch_coordinates(city, country)
        self._cache[cache_key] = list(coordinates) if coordinates else None
        return coordinates

    def save(self) -> None:
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        self.cache_path.write_text(json.dumps(self._cache, indent=2, sort_keys=True), encoding="utf-8")

    def _load_cache(self) -> dict[str, list[float] | None]:
        if not self.cache_path.exists():
            return {}

        try:
            data = json.loads(self.cache_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}

        return data if isinstance(data, dict) else {}

    def _progress(self, message: str) -> None:
        if self.progress_callback:
            self.progress_callback(message)

    def _fetch_coordinates(self, city: str, country: str) -> tuple[float, float] | None:
        query = f"{city}, {country}"
        query_params = {
            "api-version": self.api_version,
            "subscription-key": self.subscription_key,
            "query": query,
            "top": "1",
        }
        url = f"{self.endpoint}?{urllib.parse.urlencode(query_params)}"
        headers = {
            "Accept": "application/json",
            "Connection": "close",
            "User-Agent": "STC-Hotel-Data-PowerBI-Converter/1.0",
        }

        if self.client_id:
            headers["x-ms-client-id"] = self.client_id

        try:
            data = self._request_json_with_retries(url, headers, query)
        except urllib.error.HTTPError as error:
            details = error.read().decode("utf-8", errors="replace")
            raise DataValidationError(
                f"Azure Maps geocoding failed for '{query}' with HTTP {error.code}: {details[:300]}"
            ) from error
        except json.JSONDecodeError as error:
            raise DataValidationError(
                f"Azure Maps geocoding returned an invalid response for '{query}': {error}"
            ) from error

        return _coordinates_from_geocoding_response(data)

    def _request_json_with_retries(self, url: str, headers: dict[str, str], query: str) -> object:
        last_error: BaseException | None = None

        for attempt in range(1, self.retry_count + 1):
            request = urllib.request.Request(url, headers=headers)

            try:
                with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                    return json.loads(response.read().decode("utf-8"))
            except urllib.error.HTTPError as error:
                if error.code not in RETRIABLE_HTTP_STATUS_CODES or attempt == self.retry_count:
                    raise

                last_error = error
            except (urllib.error.URLError, TimeoutError, ConnectionResetError, OSError) as error:
                if attempt == self.retry_count:
                    raise DataValidationError(
                        "Azure Maps geocoding could not connect after "
                        f"{self.retry_count} attempts for '{query}'. Last error: {error}. "
                        "Check VPN/proxy/firewall access to https://atlas.microsoft.com and try again."
                    ) from error

                last_error = error

            time.sleep(self.retry_delay_seconds * attempt)

        raise DataValidationError(
            f"Azure Maps geocoding failed for '{query}' after {self.retry_count} attempts: {last_error}"
        )


def _blank_mask(series: pd.Series) -> pd.Series:
    return series.isna() | (series.astype("string").str.strip() == "")


def _clean_datetime_series(series: pd.Series, column_name: str) -> pd.Series:
    cleaned = series.replace(r"^\s*$", pd.NA, regex=True)
    parsed = pd.to_datetime(cleaned, errors="coerce")
    invalid = parsed.isna() & ~_blank_mask(series)

    if invalid.any():
        sample_values = series[invalid].astype(str).head(3).tolist()
        raise DataValidationError(
            f"Column '{column_name}' contains invalid date values: {', '.join(sample_values)}"
        )

    return parsed


def _validate_required_columns(columns: Iterable[str]) -> None:
    available = set(columns)
    missing = [column for column in REQUIRED_COLUMNS if column not in available]

    if missing:
        raise DataValidationError(f"Missing required columns: {', '.join(missing)}")


def _anonymize_value(value: object, purpose: str, salt: str) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return ""

    raw = f"{salt}:{purpose}:{str(value).strip()}".encode("utf-8")
    return hashlib.sha256(raw).hexdigest()[:12].upper()


def _anonymize_series(series: pd.Series, purpose: str, salt: str) -> pd.Series:
    return series.map(lambda value: _anonymize_value(value, purpose, salt))


def _anonymize_cancel_number(value: object, salt: str) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return ""

    text = str(value).strip()

    if text.lower() == "cancelled":
        return "cancelled"

    return _anonymize_value(text, "cancel_number", salt)


def _clean_location_text(value: object) -> str:
    if pd.isna(value):
        return ""

    return str(value).strip()


def _location_cache_key(city: str, country: str) -> str:
    return f"{city.strip().casefold()}|{country.strip().casefold()}"


def _coordinates_from_cache(value: object) -> tuple[float, float] | None:
    if not isinstance(value, list | tuple) or len(value) != 2:
        return None

    try:
        return float(value[0]), float(value[1])
    except (TypeError, ValueError):
        return None


def _coordinates_from_api_pair(value: object) -> tuple[float, float] | None:
    if not isinstance(value, list | tuple) or len(value) < 2:
        return None

    try:
        longitude = float(value[0])
        latitude = float(value[1])
    except (TypeError, ValueError):
        return None

    return latitude, longitude


def _coordinates_from_geocoding_response(data: object) -> tuple[float, float] | None:
    if not isinstance(data, dict):
        return None

    features = data.get("features")

    if not isinstance(features, list):
        return None

    for feature in features:
        if not isinstance(feature, dict):
            continue

        geometry = feature.get("geometry")

        if isinstance(geometry, dict):
            coordinates = _coordinates_from_api_pair(geometry.get("coordinates"))

            if coordinates:
                return coordinates

        properties = feature.get("properties")

        if not isinstance(properties, dict):
            continue

        geocode_points = properties.get("geocodePoints")

        if not isinstance(geocode_points, list):
            continue

        for geocode_point in geocode_points:
            if not isinstance(geocode_point, dict):
                continue

            point_geometry = geocode_point.get("geometry")

            if not isinstance(point_geometry, dict):
                continue

            coordinates = _coordinates_from_api_pair(point_geometry.get("coordinates"))

            if coordinates:
                return coordinates

    return None


def _split_city_country(value: object) -> tuple[object, str]:
    if pd.isna(value):
        return value, ""

    text = str(value).strip()

    if not text:
        return "", ""

    if "," in text:
        city, country = [part.strip() for part in text.rsplit(",", maxsplit=1)]
        return city, country or DEFAULT_COUNTRY

    return text, DEFAULT_COUNTRY


def _azure_maps_subscription_key(subscription_key: str | None = None) -> str:
    if subscription_key and subscription_key.strip():
        return subscription_key.strip()

    for env_name in (AZURE_MAPS_SUBSCRIPTION_KEY_ENV, AZURE_MAPS_KEY_ENV):
        value = os.getenv(env_name)

        if value and value.strip():
            return value.strip()

    return ""


def _geocode_cache_path(output_dir: Path, geocode_cache_path: str | Path | None) -> Path:
    if geocode_cache_path:
        return Path(geocode_cache_path)

    return output_dir / GEOCODE_CACHE_FILENAME


def _create_default_geocoder(
    output_dir: Path,
    geocode_cache_path: str | Path | None,
    *,
    subscription_key: str | None = None,
    client_id: str | None = None,
    progress_callback: ProgressCallback | None = None,
) -> AzureMapsGeocoder:
    subscription_key = _azure_maps_subscription_key(subscription_key)

    if not subscription_key:
        raise DataValidationError(
            "Azure Maps subscription key is missing. Enter the key in the UI or set "
            "AZURE_MAPS_SUBSCRIPTION_KEY before converting files with location columns."
        )

    client_id = client_id or os.getenv(AZURE_MAPS_CLIENT_ID_ENV)
    client_id = client_id.strip() if client_id and client_id.strip() else None
    return AzureMapsGeocoder(
        subscription_key,
        _geocode_cache_path(output_dir, geocode_cache_path),
        client_id=client_id,
        progress_callback=progress_callback,
    )


def _progress(progress_callback: ProgressCallback | None, message: str) -> None:
    if progress_callback:
        progress_callback(message)


def _move_columns_after(df: pd.DataFrame, anchor_column: str, columns_to_move: Iterable[str]) -> pd.DataFrame:
    move_columns = [column for column in columns_to_move if column in df.columns]

    if not move_columns or anchor_column not in df.columns:
        return df

    remaining_columns = [column for column in df.columns if column not in move_columns]
    anchor_index = remaining_columns.index(anchor_column)
    reordered_columns = (
        remaining_columns[: anchor_index + 1]
        + move_columns
        + remaining_columns[anchor_index + 1 :]
    )
    return df.loc[:, reordered_columns]


def _has_cell_value(value: object) -> bool:
    return not pd.isna(value) and str(value).strip() != ""


def _resolve_coordinates(
    geocoder: object,
    city: str,
    country: str,
    resolved_locations: dict[str, tuple[float, float] | None],
) -> tuple[float, float] | None:
    cache_key = _location_cache_key(city, country)

    if cache_key not in resolved_locations:
        resolved_locations[cache_key] = geocoder.geocode(city, country)

    return resolved_locations[cache_key]


def _fill_location_coordinates(
    df: pd.DataFrame,
    city_column: str,
    country_column: str,
    latitude_column: str,
    longitude_column: str,
    geocoder: object,
    resolved_locations: dict[str, tuple[float, float] | None],
    progress_callback: ProgressCallback | None,
) -> list[str]:
    unresolved_locations = []

    for index in df.index:
        if _has_cell_value(df.at[index, latitude_column]) and _has_cell_value(df.at[index, longitude_column]):
            continue

        city = _clean_location_text(df.at[index, city_column])
        country = _clean_location_text(df.at[index, country_column])

        if not city or not country:
            continue

        _progress(progress_callback, f"Geocoding {city}, {country}")
        coordinates = _resolve_coordinates(geocoder, city, country, resolved_locations)

        if not coordinates:
            unresolved_locations.append(f"{city}, {country}")
            continue

        latitude, longitude = coordinates
        df.at[index, latitude_column] = latitude
        df.at[index, longitude_column] = longitude

    return unresolved_locations


def _normalise_location_columns(
    df: pd.DataFrame,
    geocoder: object | None = None,
    progress_callback: ProgressCallback | None = None,
) -> pd.DataFrame:
    resolved_locations: dict[str, tuple[float, float] | None] = {}
    unresolved_locations: set[str] = set()

    for city_column, (country_column, latitude_column, longitude_column) in LOCATION_COLUMN_GROUPS.items():
        if city_column not in df.columns:
            continue

        city_country = df[city_column].map(_split_city_country)
        df[city_column] = city_country.map(lambda item: item[0])
        df[country_column] = city_country.map(lambda item: item[1])

        if latitude_column not in df.columns:
            df[latitude_column] = pd.NA
        if longitude_column not in df.columns:
            df[longitude_column] = pd.NA

        if geocoder is not None:
            unresolved_locations.update(
                _fill_location_coordinates(
                    df,
                    city_column,
                    country_column,
                    latitude_column,
                    longitude_column,
                    geocoder,
                    resolved_locations,
                    progress_callback,
                )
            )

        df = _move_columns_after(
            df,
            city_column,
            (country_column, latitude_column, longitude_column),
        )

    if unresolved_locations:
        unresolved_sample = ", ".join(sorted(unresolved_locations)[:10])
        raise DataValidationError(
            f"Azure Maps could not generate coordinates for: {unresolved_sample}"
        )

    return df


def _unique_output_path(output_dir: Path, filename: str) -> Path:
    output_path = output_dir / filename

    if not output_path.exists():
        return output_path

    stem = output_path.stem
    suffix = output_path.suffix
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return output_dir / f"{stem}_{timestamp}{suffix}"


def _read_source_workbook(input_path: Path) -> pd.DataFrame:
    if input_path.suffix.lower() == ".xls":
        raise DataValidationError(
            "This tool only supports .xlsx files. Open the file in Excel and save it as .xlsx."
        )

    if input_path.suffix.lower() != ".xlsx":
        raise DataValidationError("Please select an .xlsx Excel file.")

    if not input_path.exists():
        raise DataValidationError(f"Input file does not exist: {input_path}")

    df = pd.read_excel(input_path, engine="openpyxl", dtype={column: "string" for column in TEXT_COLUMNS})
    df.columns = [str(column).strip() for column in df.columns]
    return df


def transform_excel(
    input_path: str | Path,
    output_dir: str | Path,
    *,
    output_filename: str = OUTPUT_FILENAME,
    anonymization_salt: str | None = None,
    geocode_locations: bool = True,
    geocoder: object | None = None,
    geocode_cache_path: str | Path | None = None,
    azure_maps_subscription_key: str | None = None,
    azure_maps_client_id: str | None = None,
    progress_callback: ProgressCallback | None = None,
) -> Path:
    source_path = Path(input_path)
    target_dir = Path(output_dir)
    salt = anonymization_salt or os.getenv(ANONYMIZATION_SALT_ENV, DEFAULT_ANONYMIZATION_SALT)

    try:
        _progress(progress_callback, "Reading source workbook")
        df = _read_source_workbook(source_path)
        _progress(progress_callback, "Validating required columns")
        _validate_required_columns(df.columns)

        _progress(progress_callback, "Calculating booking and cancellation fields")
        df["QtyCorrect"] = pd.to_numeric(df["Status"], errors="coerce").map({11: 1, 13: 2})

        cancel_clean = df["Cancel Number"].astype("string").str.strip()
        rsv_clean = df["Rsv Number"].astype("string").str.strip()
        has_cancel_number = cancel_clean.notna() & (cancel_clean != "")
        cancelled_reservations = rsv_clean[has_cancel_number & rsv_clean.notna() & (rsv_clean != "")]
        cancelled_booking_rows = rsv_clean.isin(cancelled_reservations) & ~has_cancel_number
        df.loc[cancelled_booking_rows, "Cancel Number"] = "cancelled"

        cancelled = df["Cancel Number"].astype("string").str.strip().str.lower() == "cancelled"
        active = df["Cancel Number"].isna() | (df["Cancel Number"].astype("string").str.strip() == "")
        df["NotTravelled"] = df["Rate"].where(cancelled, "")
        df["travelled"] = df["Rate"].where(active, "")

        _progress(progress_callback, "Cleaning date fields")
        df["Rsv Date1"] = _clean_datetime_series(df["Rsv Date1"], "Rsv Date1")
        df["Check In"] = _clean_datetime_series(df["Check In"], "Check In")
        df["Check Out"] = _clean_datetime_series(df["Check Out"], "Check Out")

        df["Rsv Date"] = df["Rsv Date1"].dt.date
        df["Rsv Time"] = df["Rsv Date1"].dt.time
        df["Check In"] = df["Check In"].dt.date
        df["Check Out"] = df["Check Out"].dt.date
        df["Booking Hour"] = df["Rsv Date1"].dt.hour.astype("Int64")

        _progress(progress_callback, "Anonymising personal and booking data")
        for column in PERSONAL_COLUMNS:
            if column in df.columns:
                df[column] = _anonymize_series(df[column], column.lower(), salt)

        for column in IDENTIFIER_COLUMNS:
            if column in df.columns:
                df[column] = _anonymize_series(df[column], column.lower().replace(" ", "_"), salt)

        if "Cancel Number" in df.columns:
            df["Cancel Number"] = df["Cancel Number"].map(lambda value: _anonymize_cancel_number(value, salt))

        if "Address1" in df.columns:
            df["Address1"] = _anonymize_series(df["Address1"], "address", salt)
            df.rename(columns={"Address1": "AnonID"}, inplace=True)

        has_location_columns = any(column in df.columns for column in LOCATION_COLUMN_GROUPS)

        if geocode_locations and geocoder is None and has_location_columns:
            _progress(progress_callback, "Preparing Azure Maps geocoding")
            geocoder = _create_default_geocoder(
                target_dir,
                geocode_cache_path,
                subscription_key=azure_maps_subscription_key,
                client_id=azure_maps_client_id,
                progress_callback=progress_callback,
            )
        elif not geocode_locations:
            geocoder = None

        _progress(progress_callback, "Normalising city/country and coordinates")
        df = _normalise_location_columns(df, geocoder, progress_callback)

        _progress(progress_callback, "Dropping sensitive source columns")
        df.drop(columns=list(DROP_COLUMNS), inplace=True, errors="ignore")

        target_dir.mkdir(parents=True, exist_ok=True)
        output_path = _unique_output_path(target_dir, output_filename)
        _progress(progress_callback, f"Writing output workbook: {output_path.name}")
        df.to_excel(output_path, index=False, sheet_name="Sheet1", engine="openpyxl")

        if isinstance(geocoder, AzureMapsGeocoder):
            _progress(progress_callback, "Saving geocoding cache")
            geocoder.save()

        _progress(progress_callback, "Conversion complete")
        return output_path
    except DataValidationError:
        raise
    except Exception as error:
        raise RuntimeError(f"Error during transformation: {error}") from error


if __name__ == "__main__":
    from gui import run_gui

    run_gui()
