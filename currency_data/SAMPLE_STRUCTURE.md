# TGJU SANA Currency Data - Expected Output Structure

This directory will contain the fetched currency data from the TGJU SANA API.

## Files Generated

After running `fetch_currency_data.py`, the following files will be created:

### 1. `currency_data_raw.json`
Contains the raw API response exactly as received.

### 2. `currency_data_organized.json`
Contains organized data with the following structure:

```json
{
  "metadata": {
    "fetch_timestamp": "2026-01-01T12:00:00.000000",
    "source_api": "https://api.tgju.org/v1/data/sana/json",
    "total_records": 100,
    "date_range": {
      "earliest": "1402/01/01",
      "latest": "1402/12/29"
    }
  },
  "by_date": {
    "1402/10/15": [
      {
        "currency": "USD",
        "price": "520000",
        "category": "main",
        "raw": { ... }
      }
    ]
  },
  "by_currency": {
    "USD": [...],
    "EUR": [...],
    "GBP": [...]
  },
  "all_records": [...]
}
```

### 3. `currency_data.csv`
A flat CSV file with all records for easy analysis in Excel or other tools.

## Data Fields

The SANA (سنا - سامانه نظارت ارز) system typically provides:

| Field | Description |
|-------|-------------|
| currency | Currency code (USD, EUR, etc.) |
| price/rate | Exchange rate in IRR |
| date/jdate | Jalali date of the rate |
| time | Time of update |
| change | Change from previous value |

## Running the Script

```bash
# Install dependencies
pip install -r requirements.txt

# Run the fetcher
python fetch_currency_data.py
```

## Notes

- The API returns currency exchange rates from Iran's official SANA monitoring system
- Dates are typically in Jalali (Persian) calendar format
- The script handles various API response structures automatically
- Historical data availability depends on API support
