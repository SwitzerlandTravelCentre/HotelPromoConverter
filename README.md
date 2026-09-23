# Hotel Data Power BI Converter

Small desktop tool for converting hotel-data exports from [swisshotels.com](https://www.swisshotels.com/) / Switzerland Travel Centre into a cleaned and fully anonymised `promodata.xlsx` file for Power BI.

The converter is not tied to a single campaign or partner. It can process every export that matches the expected Swisshotels/STC export structure.

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

`python converter.py` also starts the same GUI for convenience.

Then select:

1. The source `.xlsx` export file.
2. The output folder.
3. Enter the Azure Maps subscription key.
4. Convert and save.

The converter writes `promodata.xlsx`. If that file already exists, it writes a timestamped filename such as `promodata_20260611_143012.xlsx` instead of overwriting the existing file.

During conversion, the GUI shows a progress bar and a live log. Longer steps such as Azure Maps geocoding run in the background so the window stays responsive.

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
- Splits `Hotel City` and `City` values such as `Zurich, Switzerland` into separate location fields for Power BI/Azure.
- Adds latitude and longitude output fields for map reporting.
- Fully anonymises configured personal-data and booking-identifier fields before writing the output file.

Location output columns:

- `Hotel City`, `Hotel Land`, `Hotel Latitude`, `Hotel Longitude`
- `City`, `Land`, `Latitude`, `Longitude`

## Anonymisation

The converter uses SHA-256 based deterministic IDs for personal-data and booking-identifier fields such as `Name`, `Firstname`, `Address1`, `Rsv Number`, and real `Cancel Number` values. This means the same source value produces the same anonymised value across runs, which is useful for reporting, but the original value is not retained in the output.

The converter also removes direct contact and free-text fields such as `Email`, `Address2`, and `Remark` from the output.

## Azure Maps Geocoding

Power BI/Azure expects location data as separate columns:

- `City`
- `Land`
- `Latitude`
- `Longitude`

Latitude and longitude are generated during conversion through Azure Maps Geocoding:

```text
https://atlas.microsoft.com/geocode?api-version=2025-01-01&subscription-key=...&query=...
```

Azure Maps returns coordinates as `longitude, latitude`; the converter writes them to Excel as `Latitude`, `Longitude`.

Enter the Azure Maps subscription key in the GUI. The key is never written to the output file or geocoding cache.

To avoid entering the key every time, enable `Save key securely for next time` in the GUI. The converter stores the key encrypted with Windows DPAPI under the current Windows user profile:

```text
%LOCALAPPDATA%\STC\HotelPromoConverter\azure_maps_subscription_key.bin
```

Only the same Windows user can decrypt that saved key. Use `Forget saved key` in the GUI to remove it.

As an alternative, store the Azure Maps subscription key outside the code as an environment variable:

```powershell
$env:AZURE_MAPS_SUBSCRIPTION_KEY = "your-azure-maps-subscription-key"
python PromoConvertv4.py
```

`AZURE_MAPS_KEY` is also supported as a backwards-compatible fallback. If IT requires the Azure Maps account client ID header, set it separately:

```powershell
$env:AZURE_MAPS_CLIENT_ID = "your-azure-maps-client-id"
```

The Azure subscription ID is not the same value as the Azure Maps `subscription-key`. Use Key 1 or Key 2 from the Azure Maps account.

Geocoding results are cached in `.azure_maps_geocode_cache.json` in the selected output folder. The cache contains only city/country coordinate lookups and no personal data.

If Azure Maps cannot resolve a non-empty location, the conversion stops with the affected city/country value so the source data can be corrected before importing the file into Power BI.

For temporary network resets such as `WinError 10054`, the converter retries the Azure Maps request automatically. If the error still appears after the retries, check VPN/proxy/firewall access to `https://atlas.microsoft.com`.

For stronger privacy, set a private salt before running:

```powershell
$env:HOTEL_PROMO_ANON_SALT = "your-private-company-salt"
python PromoConvertv4.py
```

Do not commit real Excel exports to GitHub. The `.gitignore` intentionally excludes `.xls` and `.xlsx` files.
