# Hotel Promo Converter

Small desktop tool for converting a Coop Promo hotel-data Excel export into a cleaned and anonymised `promodata.xlsx` file for analysis.

## Requirements

- Python 3.10+
- Dependencies from `requirements.txt`

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Usage

Run the GUI:

```powershell
python PromoConvertv4.py
```

Then select:

1. The source `.xlsx` export file.
2. The output folder.
3. Convert and save.

The converter writes `promodata.xlsx`. If that file already exists, it writes a timestamped filename such as `promodata_20260611_143012.xlsx` instead of overwriting the existing file.

## Expected Input Columns

Required columns:

- `Status`
- `Rsv Number`
- `Cancel Number`
- `Rate`
- `Rsv Date1`
- `Check In`
- `Check Out`

Optional columns used when present:

- `Name`
- `Firstname`
- `Address1`
- `Address2`
- `Email`
- `Remark`
- `Hotel City`
- `City`

## Transformations

- Creates `QtyCorrect` from `Status`:
  - `11` -> `1`
  - `13` -> `2`
- Marks matching booking rows as `cancelled` when a cancellation row exists for the same reservation number.
- Creates `NotTravelled` and `travelled` from `Rate`.
- Splits `Rsv Date1` into `Rsv Date`, `Rsv Time`, and `Booking Hour`.
- Converts `Check In` and `Check Out` to dates.
- Drops `Address2`, `Email`, `Remark`, and the original `Rsv Date1`.
- Adds `, Switzerland` to `Hotel City` and `City` for Power BI map matching.
- Anonymises `Name`, `Firstname`, and `Address1` deterministically.

## Anonymisation

The converter uses SHA-256 based deterministic IDs. This means the same source value produces the same anonymised value across runs, which is useful for reporting, but the original value is not retained in the output.

For stronger privacy, set a private salt before running:

```powershell
$env:HOTEL_PROMO_ANON_SALT = "your-private-company-salt"
python PromoConvertv4.py
```

Do not commit real Excel exports to GitHub. The `.gitignore` intentionally excludes `.xls` and `.xlsx` files.
