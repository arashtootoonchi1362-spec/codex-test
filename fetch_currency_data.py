#!/usr/bin/env python3
"""
TGJU SANA Currency Data Fetcher
================================
Fetches all currency exchange data from the TGJU SANA API.
The SANA system (سنا - سامانه نظارت ارز) is Iran's currency monitoring system
that provides official exchange rates.

This script fetches all available currency data organized by date.

API Endpoint: https://api.tgju.org/v1/data/sana/json
"""

import requests
import json
import time
from datetime import datetime, timedelta
import os
import sys
from typing import Optional, Dict, Any, List

# Configuration
BASE_API_URL = "https://api.tgju.org/v1/data/sana/json"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9,fa;q=0.8",
    "Referer": "https://www.tgju.org/",
    "Origin": "https://www.tgju.org"
}

# Output files
OUTPUT_DIR = "currency_data"
RAW_DATA_FILE = "currency_data_raw.json"
ORGANIZED_DATA_FILE = "currency_data_organized.json"
CSV_OUTPUT_FILE = "currency_data.csv"


def create_output_directory():
    """Create output directory if it doesn't exist"""
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"Created output directory: {OUTPUT_DIR}")


def fetch_api_data(url: str, params: Optional[Dict] = None, retries: int = 3) -> Optional[Dict]:
    """
    Fetch data from the API with retry logic

    Args:
        url: API endpoint URL
        params: Optional query parameters
        retries: Number of retry attempts

    Returns:
        JSON response as dictionary or None if failed
    """
    for attempt in range(retries):
        try:
            print(f"  Fetching: {url}")
            if params:
                print(f"  Params: {params}")

            response = requests.get(url, headers=HEADERS, params=params, timeout=60)
            response.raise_for_status()

            data = response.json()
            print(f"  ✓ Success (Status: {response.status_code})")
            return data

        except requests.exceptions.HTTPError as e:
            print(f"  ✗ HTTP Error: {e}")
            if response.status_code == 429:  # Rate limited
                wait_time = 2 ** (attempt + 1)
                print(f"  Rate limited. Waiting {wait_time} seconds...")
                time.sleep(wait_time)
            elif attempt < retries - 1:
                time.sleep(1)

        except requests.exceptions.ConnectionError as e:
            print(f"  ✗ Connection Error: {e}")
            if attempt < retries - 1:
                time.sleep(2)

        except requests.exceptions.Timeout as e:
            print(f"  ✗ Timeout Error: {e}")
            if attempt < retries - 1:
                time.sleep(2)

        except json.JSONDecodeError as e:
            print(f"  ✗ JSON Parse Error: {e}")
            return None

        except Exception as e:
            print(f"  ✗ Unexpected Error: {e}")
            if attempt < retries - 1:
                time.sleep(1)

    return None


def fetch_main_data() -> Optional[Dict]:
    """Fetch the main SANA data from the API"""
    print("\n" + "=" * 60)
    print("FETCHING MAIN SANA DATA")
    print("=" * 60)

    return fetch_api_data(BASE_API_URL)


def fetch_historical_data(start_date: Optional[str] = None, end_date: Optional[str] = None) -> Optional[Dict]:
    """
    Attempt to fetch historical data if the API supports date parameters

    Args:
        start_date: Start date in YYYY-MM-DD format
        end_date: End date in YYYY-MM-DD format
    """
    params = {}
    if start_date:
        params['start'] = start_date
        params['from'] = start_date
    if end_date:
        params['end'] = end_date
        params['to'] = end_date

    if params:
        print(f"\nAttempting historical fetch: {start_date} to {end_date}")
        return fetch_api_data(BASE_API_URL, params)

    return None


def explore_api_structure(data: Any, indent: int = 0, max_depth: int = 4) -> str:
    """
    Explore and describe the API data structure

    Returns a string description of the structure
    """
    lines = []
    prefix = "  " * indent

    if indent > max_depth:
        lines.append(f"{prefix}...")
        return "\n".join(lines)

    if isinstance(data, dict):
        lines.append(f"{prefix}Dict ({len(data)} keys):")
        for i, (key, value) in enumerate(data.items()):
            if i >= 10:  # Limit to first 10 keys
                lines.append(f"{prefix}  ... and {len(data) - 10} more keys")
                break

            if isinstance(value, dict):
                lines.append(f"{prefix}  '{key}': Dict({len(value)} keys)")
                if indent < 2:
                    lines.append(explore_api_structure(value, indent + 2, max_depth))
            elif isinstance(value, list):
                lines.append(f"{prefix}  '{key}': List({len(value)} items)")
                if value and indent < 2:
                    lines.append(f"{prefix}    First item:")
                    lines.append(explore_api_structure(value[0], indent + 3, max_depth))
            else:
                val_str = str(value)[:50]
                if len(str(value)) > 50:
                    val_str += "..."
                lines.append(f"{prefix}  '{key}': {type(value).__name__} = {val_str}")

    elif isinstance(data, list):
        lines.append(f"{prefix}List ({len(data)} items):")
        if data:
            lines.append(f"{prefix}  First item:")
            lines.append(explore_api_structure(data[0], indent + 2, max_depth))
    else:
        val_str = str(data)[:100]
        lines.append(f"{prefix}{type(data).__name__}: {val_str}")

    return "\n".join(lines)


def organize_data_by_date(data: Any) -> Dict[str, Any]:
    """
    Parse and organize the API response by date

    Returns organized data structure with metadata
    """
    organized = {
        "metadata": {
            "fetch_timestamp": datetime.now().isoformat(),
            "source_api": BASE_API_URL,
            "total_records": 0,
            "date_range": {
                "earliest": None,
                "latest": None
            }
        },
        "by_date": {},
        "by_currency": {},
        "all_records": []
    }

    all_dates = []

    def extract_date_from_item(item: Dict) -> Optional[str]:
        """Extract date from various possible field names"""
        date_fields = ['date', 'd', 'time', 'jdate', 'jalali_date', 'created_at',
                      'updated_at', 'timestamp', 'dt', 'تاریخ']

        for field in date_fields:
            if field in item and item[field]:
                return str(item[field])
        return None

    def extract_currency_info(item: Dict) -> Dict:
        """Extract currency-related information from an item"""
        currency_fields = ['currency', 'name', 'symbol', 'code', 'title', 'ارز']
        price_fields = ['price', 'value', 'rate', 'amount', 'قیمت', 'نرخ']

        info = {"raw": item}

        for field in currency_fields:
            if field in item:
                info['currency'] = item[field]
                break

        for field in price_fields:
            if field in item:
                info['price'] = item[field]
                break

        return info

    def process_items(items: List, category: str = "main"):
        """Process a list of items"""
        for item in items:
            if not isinstance(item, dict):
                continue

            date = extract_date_from_item(item)
            currency_info = extract_currency_info(item)
            currency_info['category'] = category

            organized["all_records"].append(currency_info)
            organized["metadata"]["total_records"] += 1

            if date:
                all_dates.append(date)
                if date not in organized["by_date"]:
                    organized["by_date"][date] = []
                organized["by_date"][date].append(currency_info)

            if 'currency' in currency_info:
                curr = currency_info['currency']
                if curr not in organized["by_currency"]:
                    organized["by_currency"][curr] = []
                organized["by_currency"][curr].append(currency_info)

    # Process based on data structure
    if isinstance(data, dict):
        # Check if this is a response wrapper
        if 'data' in data:
            inner_data = data['data']
            if isinstance(inner_data, list):
                process_items(inner_data)
            elif isinstance(inner_data, dict):
                for key, value in inner_data.items():
                    if isinstance(value, list):
                        process_items(value, key)
                    elif isinstance(value, dict):
                        process_items([value], key)
        else:
            # Process each top-level key
            for key, value in data.items():
                if isinstance(value, list):
                    process_items(value, key)
                elif isinstance(value, dict):
                    # Could be a single record or nested structure
                    date = extract_date_from_item(value)
                    if date or any(f in value for f in ['price', 'rate', 'value']):
                        process_items([value], key)
                    else:
                        # Nested structure - go deeper
                        for subkey, subvalue in value.items():
                            if isinstance(subvalue, list):
                                process_items(subvalue, f"{key}.{subkey}")
                            elif isinstance(subvalue, dict):
                                process_items([subvalue], f"{key}.{subkey}")

    elif isinstance(data, list):
        process_items(data)

    # Set date range
    if all_dates:
        sorted_dates = sorted(set(all_dates))
        organized["metadata"]["date_range"]["earliest"] = sorted_dates[0]
        organized["metadata"]["date_range"]["latest"] = sorted_dates[-1]

    return organized


def save_to_json(data: Any, filename: str):
    """Save data to a JSON file"""
    filepath = os.path.join(OUTPUT_DIR, filename)
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    print(f"  ✓ Saved: {filepath}")
    return filepath


def save_to_csv(organized_data: Dict, filename: str):
    """Save organized data to CSV format"""
    import csv

    filepath = os.path.join(OUTPUT_DIR, filename)

    # Flatten records for CSV
    rows = []
    for record in organized_data.get("all_records", []):
        row = {
            "category": record.get("category", ""),
            "currency": record.get("currency", ""),
            "price": record.get("price", ""),
        }

        # Add date from raw data
        raw = record.get("raw", {})
        date_fields = ['date', 'd', 'time', 'jdate', 'jalali_date']
        for field in date_fields:
            if field in raw:
                row["date"] = raw[field]
                break

        # Add other relevant fields from raw
        for key, value in raw.items():
            if key not in row and not isinstance(value, (dict, list)):
                row[key] = value

        rows.append(row)

    if rows:
        # Get all unique keys
        all_keys = set()
        for row in rows:
            all_keys.update(row.keys())

        # Write CSV
        fieldnames = sorted(all_keys)
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)

        print(f"  ✓ Saved: {filepath} ({len(rows)} rows)")
    else:
        print(f"  ⚠ No data to save to CSV")

    return filepath


def print_summary(organized_data: Dict):
    """Print a summary of the fetched data"""
    print("\n" + "=" * 60)
    print("DATA SUMMARY")
    print("=" * 60)

    meta = organized_data.get("metadata", {})

    print(f"\n📊 Total Records: {meta.get('total_records', 0)}")
    print(f"📅 Date Range: {meta.get('date_range', {}).get('earliest', 'N/A')} to {meta.get('date_range', {}).get('latest', 'N/A')}")
    print(f"📆 Unique Dates: {len(organized_data.get('by_date', {}))}")
    print(f"💱 Unique Currencies: {len(organized_data.get('by_currency', {}))}")

    # Show sample dates
    by_date = organized_data.get("by_date", {})
    if by_date:
        sorted_dates = sorted(by_date.keys())
        print(f"\n📅 Sample Dates (first 10):")
        for date in sorted_dates[:10]:
            print(f"   • {date}: {len(by_date[date])} records")

        if len(sorted_dates) > 10:
            print(f"   ... and {len(sorted_dates) - 10} more dates")

    # Show currencies
    by_currency = organized_data.get("by_currency", {})
    if by_currency:
        print(f"\n💱 Currencies found:")
        for currency in list(by_currency.keys())[:15]:
            print(f"   • {currency}: {len(by_currency[currency])} records")

        if len(by_currency) > 15:
            print(f"   ... and {len(by_currency) - 15} more currencies")


def main():
    """Main execution function"""
    print("=" * 60)
    print("       TGJU SANA CURRENCY DATA FETCHER")
    print("=" * 60)
    print(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"API Endpoint: {BASE_API_URL}")
    print("-" * 60)

    # Create output directory
    create_output_directory()

    # Fetch main data
    raw_data = fetch_main_data()

    if not raw_data:
        print("\n❌ Failed to fetch data from the API")
        print("\nPossible reasons:")
        print("  • Network connectivity issues")
        print("  • API temporarily unavailable")
        print("  • IP-based restrictions")
        print("\nSuggestions:")
        print("  • Try running the script again later")
        print("  • Check your internet connection")
        print("  • Try using a VPN if in a restricted region")
        return 1

    # Explore API structure
    print("\n" + "=" * 60)
    print("API RESPONSE STRUCTURE")
    print("=" * 60)
    print(explore_api_structure(raw_data))

    # Organize data by date
    print("\n" + "=" * 60)
    print("ORGANIZING DATA BY DATE")
    print("=" * 60)
    organized_data = organize_data_by_date(raw_data)

    # Save raw data
    print("\n" + "=" * 60)
    print("SAVING DATA")
    print("=" * 60)
    save_to_json(raw_data, RAW_DATA_FILE)
    save_to_json(organized_data, ORGANIZED_DATA_FILE)
    save_to_csv(organized_data, CSV_OUTPUT_FILE)

    # Print summary
    print_summary(organized_data)

    # Try to fetch additional historical data
    print("\n" + "=" * 60)
    print("ATTEMPTING HISTORICAL DATA FETCH")
    print("=" * 60)

    # Try different date parameters (API might support these)
    today = datetime.now()
    one_year_ago = today - timedelta(days=365)

    historical_endpoints = [
        f"{BASE_API_URL}?page=1",
        f"{BASE_API_URL}?limit=1000",
        f"{BASE_API_URL}?start={one_year_ago.strftime('%Y-%m-%d')}",
        "https://api.tgju.org/v1/data/sana/history/json",
        "https://api.tgju.org/v1/market/sana/json",
    ]

    for endpoint in historical_endpoints:
        print(f"\nTrying: {endpoint}")
        additional_data = fetch_api_data(endpoint)
        if additional_data and additional_data != raw_data:
            print("  ✓ Found additional data!")
            save_to_json(additional_data, f"historical_data_{endpoint.split('/')[-1].replace('?', '_')}.json")

    print("\n" + "=" * 60)
    print("✅ FETCH COMPLETED SUCCESSFULLY")
    print("=" * 60)
    print(f"Finished at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"\nOutput files saved in: {os.path.abspath(OUTPUT_DIR)}/")
    print(f"  • {RAW_DATA_FILE}")
    print(f"  • {ORGANIZED_DATA_FILE}")
    print(f"  • {CSV_OUTPUT_FILE}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
