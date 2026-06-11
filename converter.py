from __future__ import annotations

import hashlib
import os
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd


OUTPUT_FILENAME = "promodata.xlsx"
ANONYMIZATION_SALT_ENV = "HOTEL_PROMO_ANON_SALT"
DEFAULT_ANONYMIZATION_SALT = "hotel-promo-converter-v1"

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
DROP_COLUMNS = ("Rsv Date1", "Address2", "Email", "Remark")
CITY_COLUMNS = ("Hotel City", "City")


class DataValidationError(RuntimeError):
    """Raised when the source workbook does not match the expected export format."""


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


def _add_country_suffix(value: object) -> object:
    if pd.isna(value):
        return value

    text = str(value).strip()

    if not text or text.lower().endswith(", switzerland"):
        return text

    return f"{text}, Switzerland"


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
) -> Path:
    source_path = Path(input_path)
    target_dir = Path(output_dir)
    salt = anonymization_salt or os.getenv(ANONYMIZATION_SALT_ENV, DEFAULT_ANONYMIZATION_SALT)

    try:
        df = _read_source_workbook(source_path)
        _validate_required_columns(df.columns)

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

        df["Rsv Date1"] = _clean_datetime_series(df["Rsv Date1"], "Rsv Date1")
        df["Check In"] = _clean_datetime_series(df["Check In"], "Check In")
        df["Check Out"] = _clean_datetime_series(df["Check Out"], "Check Out")

        df["Rsv Date"] = df["Rsv Date1"].dt.date
        df["Rsv Time"] = df["Rsv Date1"].dt.time
        df["Check In"] = df["Check In"].dt.date
        df["Check Out"] = df["Check Out"].dt.date
        df["Booking Hour"] = df["Rsv Date1"].dt.hour.astype("Int64")

        for column in PERSONAL_COLUMNS:
            if column in df.columns:
                df[column] = _anonymize_series(df[column], column.lower(), salt)

        if "Address1" in df.columns:
            df["Address1"] = _anonymize_series(df["Address1"], "address", salt)
            df.rename(columns={"Address1": "AnonID"}, inplace=True)

        for column in CITY_COLUMNS:
            if column in df.columns:
                df[column] = df[column].map(_add_country_suffix)

        df.drop(columns=list(DROP_COLUMNS), inplace=True, errors="ignore")

        target_dir.mkdir(parents=True, exist_ok=True)
        output_path = _unique_output_path(target_dir, output_filename)
        df.to_excel(output_path, index=False, sheet_name="Sheet1", engine="openpyxl")

        return output_path
    except DataValidationError:
        raise
    except Exception as error:
        raise RuntimeError(f"Error during transformation: {error}") from error
