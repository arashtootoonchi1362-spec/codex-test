#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Smart CSV/Excel Scraper for Contractor vs Employer Comparison
============================================================
This script dynamically detects headers and extracts amounts from Excel files
with variable structures. It compares contractor claims against employer approved amounts.

Author: Data Architect
Version: 1.0
"""

import pandas as pd
import numpy as np
import re
import warnings
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass, field

warnings.filterwarnings('ignore')

# =============================================================================
# CONFIGURATION
# =============================================================================

# Keywords for header detection - higher weight = more important
HEADER_KEYWORDS = {
    'W.B.S': 10,
    'ردیف': 8,
    'شرح': 7,
    'عنوان فعالیت': 9,
    'عنوان فعاليت': 9,
    'مشاور زیربنایی': 8,
    'مشاور زیر بنایی': 8,
    'گروه مشارکت': 8,
    'پیمانکار': 7,
    'کارفرما': 7,
    'هزینه کامل': 6,
    'وزن کامل': 5,
    'فصل مرتبط': 4,
}

# Sub-header keywords for amount columns
AMOUNT_KEYWORDS = ['مبلغ', 'Amount', 'ریال']

# Persian to English digit mapping
PERSIAN_DIGITS = {
    '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
    '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9',
    '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4',
    '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9',
}


# =============================================================================
# DATA CLASSES
# =============================================================================

@dataclass
class ColumnMapping:
    """Stores column indices for a specific sheet"""
    item_id_col: Optional[int] = None
    wbs_col: Optional[int] = None
    description_col: Optional[int] = None
    contractor_amount_col: Optional[int] = None
    employer_amount_col: Optional[int] = None
    contractor_comments_col: Optional[int] = None
    employer_comments_col: Optional[int] = None
    header_row: int = 0
    subheader_row: int = 1
    data_start_row: int = 2


@dataclass
class SheetData:
    """Stores extracted data from a sheet"""
    sheet_name: str
    items: List[Dict[str, Any]] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    contractor_total: float = 0.0
    employer_total: float = 0.0


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def persian_to_english(text: str) -> str:
    """Convert Persian/Arabic digits to English"""
    if not isinstance(text, str):
        return str(text) if text is not None else ''
    for persian, english in PERSIAN_DIGITS.items():
        text = text.replace(persian, english)
    return text


def clean_numeric_value(value: Any) -> float:
    """
    Clean and convert a value to float.
    Handles: commas, Persian digits, dashes, slashes, text, blanks
    """
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return 0.0

    if isinstance(value, (int, float)):
        if np.isnan(value) if isinstance(value, float) else False:
            return 0.0
        return float(value)

    # Convert to string and clean
    text = str(value).strip()

    # Handle empty strings, dashes, or non-numeric text
    if not text or text == '-' or text == '—':
        return 0.0

    # Convert Persian digits
    text = persian_to_english(text)

    # Remove commas (thousand separators)
    text = text.replace(',', '')

    # Remove spaces
    text = text.replace(' ', '')

    # Handle slash-separated numbers (take first part or treat as error)
    if '/' in text and not text.replace('/', '').replace('.', '').replace('-', '').isdigit():
        parts = text.split('/')
        try:
            return float(parts[0])
        except:
            return 0.0

    # Try to extract number using regex
    match = re.search(r'[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?', text)
    if match:
        try:
            return float(match.group())
        except:
            return 0.0

    return 0.0


def normalize_text(text: Any) -> str:
    """Normalize text for comparison (strip whitespace, normalize unicode)"""
    if text is None or (isinstance(text, float) and np.isnan(text)):
        return ''
    return str(text).strip().replace('\n', ' ').replace('\r', '')


def calculate_header_confidence(row: pd.Series) -> float:
    """
    Calculate confidence score for a row being the main header.
    Returns score based on presence of header keywords.
    """
    score = 0.0
    for cell in row:
        cell_text = normalize_text(cell).lower()
        for keyword, weight in HEADER_KEYWORDS.items():
            if keyword.lower() in cell_text:
                score += weight
    return score


# =============================================================================
# STEP 1: DYNAMIC HEADER DETECTION
# =============================================================================

def detect_header_row(df: pd.DataFrame, max_scan_rows: int = 20) -> Tuple[int, int]:
    """
    Scan first N rows to find the main header row based on keyword confidence.

    Returns:
        Tuple of (header_row_index, subheader_row_index)
    """
    max_rows = min(max_scan_rows, len(df))
    scores = []

    for idx in range(max_rows):
        row = df.iloc[idx]
        score = calculate_header_confidence(row)
        scores.append((idx, score))

    # Sort by score descending
    scores.sort(key=lambda x: x[1], reverse=True)

    if scores and scores[0][1] > 0:
        header_row = scores[0][0]
        subheader_row = header_row + 1
        return header_row, subheader_row

    # Fallback: assume row 5 is header (common pattern in these files)
    return 5, 6


# =============================================================================
# STEP 2: HIERARCHICAL COLUMN MAPPING
# =============================================================================

def find_column_by_keywords(row: pd.Series, keywords: List[str],
                            start_col: int = 0, end_col: Optional[int] = None) -> Optional[int]:
    """
    Find column index containing any of the keywords within specified range.
    """
    end = end_col if end_col is not None else len(row)

    for col_idx in range(start_col, min(end, len(row))):
        cell_text = normalize_text(row.iloc[col_idx]).lower()
        for keyword in keywords:
            if keyword.lower() in cell_text:
                return col_idx
    return None


def find_group_column_span(header_row: pd.Series, group_name: str) -> Tuple[Optional[int], Optional[int]]:
    """
    Find the start and likely end column for a group header (handles merged cells).

    In pandas, merged cells have value in first cell and NaN in subsequent cells.
    We estimate the span by finding the next non-empty cell.
    """
    start_col = None

    for col_idx, cell in enumerate(header_row):
        cell_text = normalize_text(cell).lower()
        if group_name.lower() in cell_text:
            start_col = col_idx
            break

    if start_col is None:
        return None, None

    # Find end of span (next non-empty cell or end of row)
    end_col = start_col + 1
    for col_idx in range(start_col + 1, len(header_row)):
        cell_val = header_row.iloc[col_idx]
        if pd.notna(cell_val) and str(cell_val).strip():
            end_col = col_idx
            break
    else:
        end_col = len(header_row)

    # Typical span is 3 columns (percentage, weighted percentage, amount)
    # But ensure at least 3 columns
    if end_col - start_col < 3:
        end_col = min(start_col + 4, len(header_row))

    return start_col, end_col


def find_amount_column_in_span(subheader_row: pd.Series, start_col: int, end_col: int) -> Optional[int]:
    """
    Find the 'مبلغ' (Amount) column within a specified span.
    """
    for col_idx in range(start_col, end_col):
        if col_idx < len(subheader_row):
            cell_text = normalize_text(subheader_row.iloc[col_idx]).lower()
            for keyword in AMOUNT_KEYWORDS:
                if keyword.lower() in cell_text:
                    return col_idx

    # Fallback: amount is typically the 3rd column in the group (index +2)
    fallback_col = start_col + 2
    if fallback_col < end_col:
        return fallback_col

    return None


def map_columns_employer(df: pd.DataFrame, header_row: int, subheader_row: int) -> ColumnMapping:
    """
    Map columns for employer.xlsx file.
    Target: Find 'مشاور زیربنایی' -> 'مبلغ' column for approved amounts.
    """
    mapping = ColumnMapping()
    mapping.header_row = header_row
    mapping.subheader_row = subheader_row
    mapping.data_start_row = subheader_row + 1

    header = df.iloc[header_row]
    subheader = df.iloc[subheader_row] if subheader_row < len(df) else pd.Series()

    # Find anchor columns (Item ID, WBS, Description)
    mapping.item_id_col = find_column_by_keywords(header, ['ردیف'])
    mapping.wbs_col = find_column_by_keywords(header, ['W.B.S', 'WBS', 'w.b.s'])
    mapping.description_col = find_column_by_keywords(header, ['شرح', 'عنوان فعالیت', 'عنوان فعاليت', 'Description'])

    # Find Employer Approved Amount (مشاور زیربنایی -> مبلغ)
    # Primary search: 'مشاور زیربنایی' or 'مشاور زیر بنایی'
    group_names = ['مشاور زیربنایی', 'مشاور زیر بنایی', 'Infrastructure']

    for group_name in group_names:
        start_col, end_col = find_group_column_span(header, group_name)
        if start_col is not None:
            amount_col = find_amount_column_in_span(subheader, start_col, end_col)
            if amount_col is not None:
                mapping.employer_amount_col = amount_col
                break

    # Fallback: try 'تایید کارفرما' or 'Approved' or 'کارفرما'
    if mapping.employer_amount_col is None:
        fallback_groups = ['تایید کارفرما', 'کارفرما', 'Approved']
        for group_name in fallback_groups:
            start_col, end_col = find_group_column_span(header, group_name)
            if start_col is not None:
                amount_col = find_amount_column_in_span(subheader, start_col, end_col)
                if amount_col is not None:
                    mapping.employer_amount_col = amount_col
                    break

    # Find comments column (توضیحات)
    for col_idx, cell in enumerate(header):
        cell_text = normalize_text(cell).lower()
        if 'توضیحات' in cell_text:
            mapping.employer_comments_col = col_idx
            break

    return mapping


def map_columns_contractor(df: pd.DataFrame, header_row: int, subheader_row: int) -> ColumnMapping:
    """
    Map columns for contractor.xlsx file.
    Target: Find 'گروه مشارکت' or 'پیمانکار' -> 'مبلغ' column for claimed amounts.
    """
    mapping = ColumnMapping()
    mapping.header_row = header_row
    mapping.subheader_row = subheader_row
    mapping.data_start_row = subheader_row + 1

    header = df.iloc[header_row]
    subheader = df.iloc[subheader_row] if subheader_row < len(df) else pd.Series()

    # Find anchor columns
    mapping.item_id_col = find_column_by_keywords(header, ['ردیف'])
    mapping.wbs_col = find_column_by_keywords(header, ['W.B.S', 'WBS', 'w.b.s'])
    mapping.description_col = find_column_by_keywords(header, ['شرح', 'عنوان فعالیت', 'عنوان فعاليت', 'Description'])

    # Find Contractor Claim Amount
    # Primary: 'پیمانکار-قطعی' or 'پیمانکار'
    group_names = ['پیمانکار-قطعی', 'پیمانکار', 'گروه مشارکت', 'Contractor']

    for group_name in group_names:
        start_col, end_col = find_group_column_span(header, group_name)
        if start_col is not None:
            amount_col = find_amount_column_in_span(subheader, start_col, end_col)
            if amount_col is not None:
                mapping.contractor_amount_col = amount_col
                break

    # Find contractor comments column
    for col_idx, cell in enumerate(header):
        cell_text = normalize_text(cell).lower()
        if 'توضیحات پیمانکار' in cell_text or 'توضیحات' in cell_text:
            mapping.contractor_comments_col = col_idx
            break

    return mapping


# =============================================================================
# STEP 3: DATA EXTRACTION & CLEANING
# =============================================================================

def extract_sheet_data(df: pd.DataFrame, mapping: ColumnMapping, sheet_name: str,
                       is_employer: bool = True) -> SheetData:
    """
    Extract data from a sheet using the mapped columns.
    """
    sheet_data = SheetData(sheet_name=sheet_name)

    # Determine which amount column to use
    amount_col = mapping.employer_amount_col if is_employer else mapping.contractor_amount_col
    comments_col = mapping.employer_comments_col if is_employer else mapping.contractor_comments_col

    if amount_col is None:
        sheet_data.warnings.append(f"Could not find amount column for sheet '{sheet_name}'")
        return sheet_data

    # Extract data rows
    for row_idx in range(mapping.data_start_row, len(df)):
        row = df.iloc[row_idx]

        # Get WBS (primary identifier)
        wbs = ''
        if mapping.wbs_col is not None and mapping.wbs_col < len(row):
            wbs = normalize_text(row.iloc[mapping.wbs_col])
            wbs = persian_to_english(wbs)

        # Skip rows without WBS
        if not wbs:
            continue

        # Get Item ID
        item_id = ''
        if mapping.item_id_col is not None and mapping.item_id_col < len(row):
            item_id = normalize_text(row.iloc[mapping.item_id_col])
            item_id = persian_to_english(item_id)

        # Get Description
        description = ''
        if mapping.description_col is not None and mapping.description_col < len(row):
            description = normalize_text(row.iloc[mapping.description_col])

        # Get Amount
        amount = 0.0
        if amount_col < len(row):
            amount = clean_numeric_value(row.iloc[amount_col])

        # Get Comments
        comments = ''
        if comments_col is not None and comments_col < len(row):
            comments = normalize_text(row.iloc[comments_col])

        item = {
            'wbs': wbs,
            'item_id': item_id,
            'description': description,
            'amount': amount,
            'comments': comments,
            'row_num': row_idx + 1  # 1-indexed for human readability
        }

        sheet_data.items.append(item)

        if is_employer:
            sheet_data.employer_total += amount
        else:
            sheet_data.contractor_total += amount

    return sheet_data


# =============================================================================
# STEP 4: COMPARISON & REPORT GENERATION
# =============================================================================

def compare_and_merge(contractor_data: SheetData, employer_data: SheetData) -> List[Dict[str, Any]]:
    """
    Compare contractor and employer data, match by WBS, and identify disputes.
    """
    results = []

    # Create lookup dictionaries
    contractor_lookup = {item['wbs']: item for item in contractor_data.items}
    employer_lookup = {item['wbs']: item for item in employer_data.items}

    # Get all unique WBS codes
    all_wbs = set(contractor_lookup.keys()) | set(employer_lookup.keys())

    for wbs in sorted(all_wbs):
        contractor_item = contractor_lookup.get(wbs, {})
        employer_item = employer_lookup.get(wbs, {})

        contractor_amount = contractor_item.get('amount', 0.0)
        employer_amount = employer_item.get('amount', 0.0)
        difference = contractor_amount - employer_amount

        result = {
            'wbs': wbs,
            'item_id': contractor_item.get('item_id', '') or employer_item.get('item_id', ''),
            'description': contractor_item.get('description', '') or employer_item.get('description', ''),
            'contractor_amount': contractor_amount,
            'employer_amount': employer_amount,
            'difference': difference,
            'contractor_comments': contractor_item.get('comments', ''),
            'employer_comments': employer_item.get('comments', ''),
            'is_dispute': abs(difference) > 0.01,  # Consider small rounding differences
            'only_in_contractor': wbs not in employer_lookup,
            'only_in_employer': wbs not in contractor_lookup,
        }

        results.append(result)

    return results


def process_files(contractor_path: str, employer_path: str) -> Dict[str, List[Dict[str, Any]]]:
    """
    Main processing function. Reads both files, processes all sheets, and returns comparison results.
    """
    print("=" * 80)
    print("SMART CSV/EXCEL SCRAPER - Contractor vs Employer Comparison")
    print("=" * 80)

    # Load Excel files
    print(f"\nLoading files...")
    contractor_xls = pd.ExcelFile(contractor_path)
    employer_xls = pd.ExcelFile(employer_path)

    contractor_sheets = set(contractor_xls.sheet_names)
    employer_sheets = set(employer_xls.sheet_names)

    # Find common sheets (where comparison makes sense)
    common_sheets = contractor_sheets & employer_sheets

    print(f"  Contractor sheets: {len(contractor_sheets)}")
    print(f"  Employer sheets: {len(employer_sheets)}")
    print(f"  Common sheets: {len(common_sheets)}")

    # Skip sheets that are clearly summary/non-data sheets
    skip_sheets = {'summary', 'summary-2', 'summary-Adjustment', 'CBS', 'فروش كالا و خدمات'}

    all_results = {}
    warnings_log = []

    for sheet_name in sorted(common_sheets):
        if sheet_name.lower() in [s.lower() for s in skip_sheets]:
            print(f"\n  Skipping summary sheet: '{sheet_name}'")
            continue

        print(f"\n{'=' * 60}")
        print(f"Processing sheet: '{sheet_name}'")
        print("=" * 60)

        try:
            # Read sheets without header (we'll detect it dynamically)
            contractor_df = pd.read_excel(contractor_xls, sheet_name=sheet_name, header=None)
            employer_df = pd.read_excel(employer_xls, sheet_name=sheet_name, header=None)

            # STEP 1: Detect header rows
            c_header_row, c_subheader_row = detect_header_row(contractor_df)
            e_header_row, e_subheader_row = detect_header_row(employer_df)

            print(f"  Contractor - Header row: {c_header_row}, Sub-header row: {c_subheader_row}")
            print(f"  Employer - Header row: {e_header_row}, Sub-header row: {e_subheader_row}")

            # STEP 2: Map columns
            contractor_mapping = map_columns_contractor(contractor_df, c_header_row, c_subheader_row)
            employer_mapping = map_columns_employer(employer_df, e_header_row, e_subheader_row)

            print(f"  Contractor columns - WBS: {contractor_mapping.wbs_col}, Amount: {contractor_mapping.contractor_amount_col}")
            print(f"  Employer columns - WBS: {employer_mapping.wbs_col}, Amount: {employer_mapping.employer_amount_col}")

            # STEP 3: Extract data
            contractor_data = extract_sheet_data(contractor_df, contractor_mapping, sheet_name, is_employer=False)
            employer_data = extract_sheet_data(employer_df, employer_mapping, sheet_name, is_employer=True)

            print(f"  Contractor items extracted: {len(contractor_data.items)}")
            print(f"  Employer items extracted: {len(employer_data.items)}")

            # STEP 4: Sanity checks
            if employer_data.employer_total == 0 and len(employer_data.items) > 0:
                warning_msg = f"⚠️  WARNING: Sheet '{sheet_name}' has {len(employer_data.items)} rows but EMPLOYER TOTAL = 0!"
                print(warning_msg)
                warnings_log.append(warning_msg)

            if contractor_data.contractor_total == 0 and len(contractor_data.items) > 0:
                warning_msg = f"⚠️  WARNING: Sheet '{sheet_name}' has {len(contractor_data.items)} rows but CONTRACTOR TOTAL = 0!"
                print(warning_msg)
                warnings_log.append(warning_msg)

            # Compare and merge
            comparison_results = compare_and_merge(contractor_data, employer_data)

            # Count disputes
            disputes = [r for r in comparison_results if r['is_dispute']]
            print(f"  Disputes found: {len(disputes)} out of {len(comparison_results)} items")

            if comparison_results:
                all_results[sheet_name] = comparison_results

        except Exception as e:
            error_msg = f"❌ ERROR processing sheet '{sheet_name}': {str(e)}"
            print(error_msg)
            warnings_log.append(error_msg)

    # Print summary warnings
    if warnings_log:
        print("\n" + "=" * 80)
        print("⚠️  WARNINGS SUMMARY")
        print("=" * 80)
        for warning in warnings_log:
            print(warning)

    return all_results


def generate_excel_report(results: Dict[str, List[Dict[str, Any]]], output_path: str):
    """
    Generate the final Excel report with formatting using xlsxwriter.
    """
    print(f"\n{'=' * 80}")
    print(f"Generating report: {output_path}")
    print("=" * 80)

    import xlsxwriter

    workbook = xlsxwriter.Workbook(output_path)

    # Define formats
    # Title format
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'font_name': 'B Nazanin',
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#4472C4',
        'font_color': 'white',
        'border': 1,
        'reading_order': 2,  # RTL
    })

    # Header format
    header_format = workbook.add_format({
        'bold': True,
        'font_size': 11,
        'font_name': 'B Nazanin',
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#D9E2F3',
        'border': 1,
        'text_wrap': True,
        'reading_order': 2,  # RTL
    })

    # Normal cell format (RTL)
    cell_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'B Nazanin',
        'align': 'right',
        'valign': 'vcenter',
        'border': 1,
        'reading_order': 2,  # RTL
    })

    # Number format
    number_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'B Nazanin',
        'align': 'right',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '#,##0',
        'reading_order': 2,  # RTL
    })

    # Dispute (RED) format for differences
    dispute_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'B Nazanin',
        'align': 'right',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '#,##0',
        'bg_color': '#FFCDD2',
        'font_color': '#B71C1C',
        'bold': True,
        'reading_order': 2,  # RTL
    })

    # Link format
    link_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'B Nazanin',
        'font_color': 'blue',
        'underline': True,
        'align': 'right',
        'reading_order': 2,  # RTL
    })

    # =========================================================================
    # Create Dashboard sheet
    # =========================================================================
    dashboard = workbook.add_worksheet('Dashboard')
    dashboard.right_to_left()

    # Dashboard headers
    dashboard.merge_range('A1:G1', 'داشبورد مقایسه پیمانکار و کارفرما', title_format)

    dashboard_headers = ['ردیف', 'نام برگه', 'تعداد آیتم', 'اختلاف‌ها', 'جمع پیمانکار', 'جمع کارفرما', 'جمع اختلاف']
    for col, header in enumerate(dashboard_headers):
        dashboard.write(2, col, header, header_format)

    dashboard.set_column('A:A', 8)
    dashboard.set_column('B:B', 35)
    dashboard.set_column('C:C', 12)
    dashboard.set_column('D:D', 12)
    dashboard.set_column('E:G', 18)

    row = 3
    for idx, (sheet_name, items) in enumerate(results.items(), 1):
        disputes = [i for i in items if i['is_dispute']]
        total_contractor = sum(i['contractor_amount'] for i in items)
        total_employer = sum(i['employer_amount'] for i in items)
        total_diff = sum(i['difference'] for i in items if i['is_dispute'])

        dashboard.write(row, 0, idx, cell_format)

        # Create internal link to sheet
        safe_sheet_name = sheet_name[:31] if len(sheet_name) > 31 else sheet_name
        dashboard.write_url(row, 1, f"internal:'{safe_sheet_name}'!A1", link_format, sheet_name)

        dashboard.write(row, 2, len(items), number_format)
        dashboard.write(row, 3, len(disputes), dispute_format if disputes else number_format)
        dashboard.write(row, 4, total_contractor, number_format)
        dashboard.write(row, 5, total_employer, number_format)
        dashboard.write(row, 6, total_diff, dispute_format if total_diff != 0 else number_format)

        row += 1

    # =========================================================================
    # Create individual sheets for each category
    # =========================================================================
    column_headers = [
        'ردیف',
        'W.B.S',
        'شرح',
        'مبلغ پیمانکار (ریال)',
        'مبلغ کارفرما (ریال)',
        'اختلاف (ریال)',
        'توضیحات پیمانکار',
        'توضیحات کارفرما'
    ]

    for sheet_name, items in results.items():
        # Truncate sheet name if too long (Excel limit is 31 chars)
        safe_name = sheet_name[:31] if len(sheet_name) > 31 else sheet_name

        # Handle duplicate names (shouldn't happen but just in case)
        try:
            worksheet = workbook.add_worksheet(safe_name)
        except:
            safe_name = safe_name[:28] + '...'
            worksheet = workbook.add_worksheet(safe_name)

        worksheet.right_to_left()

        # Title
        worksheet.merge_range('A1:H1', f'برگه: {sheet_name}', title_format)

        # Headers
        for col, header in enumerate(column_headers):
            worksheet.write(2, col, header, header_format)

        # Set column widths
        worksheet.set_column('A:A', 8)   # ردیف
        worksheet.set_column('B:B', 15)  # W.B.S
        worksheet.set_column('C:C', 45)  # شرح
        worksheet.set_column('D:F', 18)  # Amounts
        worksheet.set_column('G:H', 40)  # Comments

        # Data rows
        row = 3
        for idx, item in enumerate(items, 1):
            worksheet.write(row, 0, idx, cell_format)
            worksheet.write(row, 1, item['wbs'], cell_format)
            worksheet.write(row, 2, item['description'], cell_format)
            worksheet.write(row, 3, item['contractor_amount'], number_format)
            worksheet.write(row, 4, item['employer_amount'], number_format)

            # Highlight differences
            diff_fmt = dispute_format if item['is_dispute'] else number_format
            worksheet.write(row, 5, item['difference'], diff_fmt)

            worksheet.write(row, 6, item['contractor_comments'], cell_format)
            worksheet.write(row, 7, item['employer_comments'], cell_format)

            row += 1

        # Add totals row
        if items:
            total_contractor = sum(i['contractor_amount'] for i in items)
            total_employer = sum(i['employer_amount'] for i in items)
            total_diff = total_contractor - total_employer

            worksheet.write(row, 2, 'جمع کل:', header_format)
            worksheet.write(row, 3, total_contractor, header_format)
            worksheet.write(row, 4, total_employer, header_format)
            worksheet.write(row, 5, total_diff, dispute_format if total_diff != 0 else header_format)

    # =========================================================================
    # Create a "Disputes Only" summary sheet
    # =========================================================================
    disputes_sheet = workbook.add_worksheet('اختلافات')
    disputes_sheet.right_to_left()

    disputes_sheet.merge_range('A1:I1', 'خلاصه اختلافات', title_format)

    disputes_headers = ['ردیف', 'برگه', 'W.B.S', 'شرح', 'پیمانکار', 'کارفرما', 'اختلاف', 'توضیحات پیمانکار', 'توضیحات کارفرما']
    for col, header in enumerate(disputes_headers):
        disputes_sheet.write(2, col, header, header_format)

    disputes_sheet.set_column('A:A', 8)
    disputes_sheet.set_column('B:B', 25)
    disputes_sheet.set_column('C:C', 15)
    disputes_sheet.set_column('D:D', 40)
    disputes_sheet.set_column('E:G', 16)
    disputes_sheet.set_column('H:I', 35)

    row = 3
    idx = 1
    for sheet_name, items in results.items():
        for item in items:
            if item['is_dispute']:
                disputes_sheet.write(row, 0, idx, cell_format)
                disputes_sheet.write(row, 1, sheet_name, cell_format)
                disputes_sheet.write(row, 2, item['wbs'], cell_format)
                disputes_sheet.write(row, 3, item['description'], cell_format)
                disputes_sheet.write(row, 4, item['contractor_amount'], number_format)
                disputes_sheet.write(row, 5, item['employer_amount'], number_format)
                disputes_sheet.write(row, 6, item['difference'], dispute_format)
                disputes_sheet.write(row, 7, item['contractor_comments'], cell_format)
                disputes_sheet.write(row, 8, item['employer_comments'], cell_format)
                row += 1
                idx += 1

    workbook.close()

    print(f"\n✅ Report generated successfully: {output_path}")
    print(f"   - Total sheets: {len(results) + 2} (Dashboard + Disputes + {len(results)} category sheets)")

    total_disputes = sum(1 for items in results.values() for i in items if i['is_dispute'])
    print(f"   - Total disputes found: {total_disputes}")


# =============================================================================
# MAIN EXECUTION
# =============================================================================

def main():
    """Main entry point"""
    # File paths
    contractor_file = 'contractor.xlsx'
    employer_file = 'employer.xlsx'
    output_file = 'disputes_verified_master.xlsx'

    # Check files exist
    if not Path(contractor_file).exists():
        print(f"❌ ERROR: Contractor file not found: {contractor_file}")
        return 1

    if not Path(employer_file).exists():
        print(f"❌ ERROR: Employer file not found: {employer_file}")
        return 1

    try:
        # Process files
        results = process_files(contractor_file, employer_file)

        if not results:
            print("\n⚠️  No data extracted. Please check the input files.")
            return 1

        # Generate report
        generate_excel_report(results, output_file)

        # Print final summary
        print("\n" + "=" * 80)
        print("PROCESSING COMPLETE")
        print("=" * 80)

        total_items = sum(len(items) for items in results.values())
        total_disputes = sum(1 for items in results.values() for i in items if i['is_dispute'])
        total_contractor = sum(i['contractor_amount'] for items in results.values() for i in items)
        total_employer = sum(i['employer_amount'] for items in results.values() for i in items)

        print(f"  Sheets processed: {len(results)}")
        print(f"  Total line items: {total_items}")
        print(f"  Total disputes: {total_disputes}")
        print(f"  Contractor Total: {total_contractor:,.0f} ریال")
        print(f"  Employer Total: {total_employer:,.0f} ریال")
        print(f"  Net Difference: {total_contractor - total_employer:,.0f} ریال")
        print(f"\n  Output file: {output_file}")

        return 0

    except Exception as e:
        print(f"\n❌ FATAL ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    exit(main())
