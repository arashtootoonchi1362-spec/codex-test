#!/usr/bin/env python3
"""
Dispute Generator Script
Compares contractor.xlsx and employer.xlsx to identify disputes
and outputs a well-formatted RTL Excel report using xlsxwriter.
"""

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name
import warnings
import sys

warnings.filterwarnings('ignore')

# Configuration
CONTRACTOR_FILE = 'contractor.xlsx'
EMPLOYER_FILE = 'employer.xlsx'
OUTPUT_FILE = 'disputes_final_formatted.xlsx'

# Sheets to analyze (common between both files)
SHEETS_TO_ANALYZE = [
    'تجهیز کارگاه اولیه',
    'تجهیز کارگاه مستمر',
    'خدمات طراحي  مهندسي پیش از اجرا',
    'خدمات طراحي و مهندسي حین اجرا',
    'خدمات مهندسي حين اجرا',
    'خدمات تامین',
    'خدمات اجرایی ',
    'اختتامیه',
    'بیمه نامه ها',
    'سازمان دهی و برنامه ریزی پروژه',
]


def clean_numeric(val):
    """Convert value to numeric, handling NaN and strings."""
    if pd.isna(val):
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def find_disputes_in_sheet(contractor_df, employer_df, sheet_name):
    """
    Compare contractor and employer data to find disputes.
    A dispute is when contractor claims a value but employer shows 0 or different value.

    Data structure (based on analysis):
    - Row 0-4: Header info (کارفرما، مشاور، etc.)
    - Row 5: Main column headers (ردیف, W.B.S, عنوان فعالیت, etc.)
    - Row 6: Sub-headers (درصد پیشرفت, مبلغ, etc.)
    - Row 7+: Actual data

    Column mapping:
    - Col 0: Row number (ردیف)
    - Col 1: WBS code
    - Col 2: Activity name (عنوان فعالیت)
    - Col 3: Total cost (هزینه کامل)
    - Col 4: Weight % (وزن کامل)
    - Col 6-8: Previous approved (تایید شده کارکرد قبلی)
    - Col 9-11: Contractor claimed (پیمانکار-قطعی) - Col 11 is amount
    - Col 12: Contractor notes (توضیحات پیمانکار)
    - Col 14-16: Employer approved (مشاور کارفرما) - Col 16 is amount
    - Col 17: Employer notes (توضیحات مشاور)
    """
    disputes = []

    # Data starts at row 7 (0-indexed)
    data_start_row = 7

    for idx in range(data_start_row, len(contractor_df)):
        try:
            row_c = contractor_df.iloc[idx]

            # Get row number (column 0)
            row_num = row_c.iloc[0] if pd.notna(row_c.iloc[0]) else ''

            # Skip non-data rows (headers, empty rows, subtotals)
            if pd.isna(row_num) or str(row_num).strip() in ['', 'ردیف', 'NaN', 'جمع', 'جمع کل']:
                continue

            # Try to get numeric row number
            try:
                row_num_int = int(float(row_num))
            except:
                continue

            # Get WBS code (column 1)
            wbs = str(row_c.iloc[1]) if pd.notna(row_c.iloc[1]) else ''

            # Get activity name (column 2)
            activity = str(row_c.iloc[2]) if pd.notna(row_c.iloc[2]) else ''

            # Get total cost/budget (column 3)
            total_cost = clean_numeric(row_c.iloc[3]) if len(row_c) > 3 else 0

            # Get contractor's progress % (column 9 - درصد پیشرفت از سرفصل)
            contractor_progress = clean_numeric(row_c.iloc[9]) if len(row_c) > 9 else 0

            # Get contractor's claimed amount (column 11 - مبلغ پیمانکار)
            contractor_amount = clean_numeric(row_c.iloc[11]) if len(row_c) > 11 else 0

            # Get employer's progress % (column 14 - درصد پیشرفت کارفرما)
            employer_progress = clean_numeric(row_c.iloc[14]) if len(row_c) > 14 else 0

            # Get employer's approved amount (column 16 - مبلغ کارفرما)
            employer_amount = clean_numeric(row_c.iloc[16]) if len(row_c) > 16 else 0

            # Get contractor notes (column 12)
            contractor_notes = str(row_c.iloc[12]) if len(row_c) > 12 and pd.notna(row_c.iloc[12]) else ''

            # Get employer notes (column 17)
            employer_notes = str(row_c.iloc[17]) if len(row_c) > 17 and pd.notna(row_c.iloc[17]) else ''

            # Calculate dispute amount
            dispute_amount = contractor_amount - employer_amount

            # A dispute exists if there's a meaningful difference
            # (using threshold to avoid floating point issues)
            if abs(dispute_amount) > 100:  # 100 Rial threshold
                disputes.append({
                    'شماره ردیف': row_num_int,
                    'کد WBS': wbs,
                    'شرح فعالیت': activity,
                    'برگه': sheet_name,
                    'هزینه کامل (ریال)': total_cost,
                    'درصد پیشرفت پیمانکار': contractor_progress,
                    'مبلغ ادعای پیمانکار (ریال)': contractor_amount,
                    'درصد پیشرفت کارفرما': employer_progress,
                    'مبلغ تایید کارفرما (ریال)': employer_amount,
                    'مبلغ اختلاف (ریال)': dispute_amount,
                    'توضیحات پیمانکار': contractor_notes if contractor_notes != 'nan' else '',
                    'توضیحات مشاور': employer_notes if employer_notes != 'nan' else '',
                })
        except Exception as e:
            continue

    return disputes


def load_and_process_data():
    """Load both Excel files and find all disputes."""
    print("Loading Excel files...")

    try:
        xls_contractor = pd.ExcelFile(CONTRACTOR_FILE)
        xls_employer = pd.ExcelFile(EMPLOYER_FILE)
    except Exception as e:
        print(f"Error loading files: {e}")
        sys.exit(1)

    all_disputes = []
    sheet_disputes = {}

    # Get common sheets
    contractor_sheets = set(xls_contractor.sheet_names)
    employer_sheets = set(xls_employer.sheet_names)

    for sheet_name in SHEETS_TO_ANALYZE:
        if sheet_name not in contractor_sheets:
            print(f"  Skipping {sheet_name} (not in contractor file)")
            continue

        print(f"  Processing: {sheet_name}")

        try:
            # Read with header at row 3 (0-indexed)
            df_contractor = pd.read_excel(xls_contractor, sheet_name=sheet_name, header=None)

            # For employer, check if sheet exists
            if sheet_name in employer_sheets:
                df_employer = pd.read_excel(xls_employer, sheet_name=sheet_name, header=None)
            else:
                df_employer = pd.DataFrame()

            disputes = find_disputes_in_sheet(df_contractor, df_employer, sheet_name)

            if disputes:
                sheet_disputes[sheet_name] = disputes
                all_disputes.extend(disputes)
                print(f"    Found {len(disputes)} disputes")
        except Exception as e:
            print(f"    Error processing {sheet_name}: {e}")
            continue

    return all_disputes, sheet_disputes


def write_formatted_excel(all_disputes, sheet_disputes):
    """Write disputes to a well-formatted Excel file using xlsxwriter."""

    print(f"\nWriting {len(all_disputes)} disputes to {OUTPUT_FILE}...")

    # Create workbook with xlsxwriter
    workbook = xlsxwriter.Workbook(OUTPUT_FILE, {'strings_to_urls': False})

    # Define formats
    # Title format
    title_format = workbook.add_format({
        'font_name': 'B Nazanin',
        'font_size': 16,
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_color': '#1F4E79',
        'reading_order': 2,  # RTL
    })

    # Subtitle format
    subtitle_format = workbook.add_format({
        'font_name': 'B Nazanin',
        'font_size': 12,
        'align': 'center',
        'valign': 'vcenter',
        'font_color': '#5B9BD5',
        'reading_order': 2,
    })

    # Header format (Row 4 - bold and distinct)
    header_format = workbook.add_format({
        'font_name': 'B Nazanin',
        'font_size': 11,
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'bg_color': '#1F4E79',
        'font_color': '#FFFFFF',
        'border': 1,
        'border_color': '#000000',
        'reading_order': 2,
    })

    # Data format - text
    text_format = workbook.add_format({
        'font_name': 'B Nazanin',
        'font_size': 10,
        'align': 'right',
        'valign': 'vcenter',
        'text_wrap': True,
        'border': 1,
        'border_color': '#D9D9D9',
        'reading_order': 2,
    })

    # Data format - numbers
    number_format = workbook.add_format({
        'font_name': 'Tahoma',
        'font_size': 10,
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '#,##0',
        'border': 1,
        'border_color': '#D9D9D9',
        'reading_order': 2,
    })

    # Data format - percentage
    percent_format = workbook.add_format({
        'font_name': 'Tahoma',
        'font_size': 10,
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '0.00%',
        'border': 1,
        'border_color': '#D9D9D9',
        'reading_order': 2,
    })

    # Dispute amount format (positive - contractor claims more)
    dispute_positive_format = workbook.add_format({
        'font_name': 'Tahoma',
        'font_size': 10,
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '#,##0',
        'font_color': '#006600',
        'bg_color': '#C6EFCE',
        'border': 1,
        'border_color': '#D9D9D9',
        'reading_order': 2,
    })

    # Dispute amount format (negative - employer claims more)
    dispute_negative_format = workbook.add_format({
        'font_name': 'Tahoma',
        'font_size': 10,
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '#,##0',
        'font_color': '#9C0006',
        'bg_color': '#FFC7CE',
        'border': 1,
        'border_color': '#D9D9D9',
        'reading_order': 2,
    })

    # Notes format
    notes_format = workbook.add_format({
        'font_name': 'B Nazanin',
        'font_size': 9,
        'align': 'right',
        'valign': 'vcenter',
        'text_wrap': True,
        'border': 1,
        'border_color': '#D9D9D9',
        'reading_order': 2,
    })

    # Alternating row format
    alt_text_format = workbook.add_format({
        'font_name': 'B Nazanin',
        'font_size': 10,
        'align': 'right',
        'valign': 'vcenter',
        'text_wrap': True,
        'border': 1,
        'border_color': '#D9D9D9',
        'bg_color': '#F2F2F2',
        'reading_order': 2,
    })

    alt_number_format = workbook.add_format({
        'font_name': 'Tahoma',
        'font_size': 10,
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '#,##0',
        'border': 1,
        'border_color': '#D9D9D9',
        'bg_color': '#F2F2F2',
        'reading_order': 2,
    })

    alt_percent_format = workbook.add_format({
        'font_name': 'Tahoma',
        'font_size': 10,
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '0.00%',
        'border': 1,
        'border_color': '#D9D9D9',
        'bg_color': '#F2F2F2',
        'reading_order': 2,
    })

    alt_notes_format = workbook.add_format({
        'font_name': 'B Nazanin',
        'font_size': 9,
        'align': 'right',
        'valign': 'vcenter',
        'text_wrap': True,
        'border': 1,
        'border_color': '#D9D9D9',
        'bg_color': '#F2F2F2',
        'reading_order': 2,
    })

    # Summary format
    summary_format = workbook.add_format({
        'font_name': 'B Nazanin',
        'font_size': 12,
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '#,##0',
        'bg_color': '#D9E1F2',
        'border': 2,
        'reading_order': 2,
    })

    # Column headers
    headers = [
        'ردیف',
        'کد WBS',
        'شرح فعالیت',
        'برگه کاری',
        'هزینه کامل (ریال)',
        'درصد پیشرفت پیمانکار',
        'مبلغ ادعای پیمانکار (ریال)',
        'درصد پیشرفت کارفرما',
        'مبلغ تایید کارفرما (ریال)',
        'مبلغ اختلاف (ریال)',
        'توضیحات پیمانکار',
        'توضیحات مشاور/کارفرما',
    ]

    # Column widths (approximate for Persian text)
    col_widths = [8, 15, 40, 25, 22, 18, 25, 18, 25, 25, 50, 50]

    def write_sheet(worksheet, disputes_data, sheet_title):
        """Write a formatted sheet with disputes data."""

        # Set RTL
        worksheet.right_to_left()

        # Set column widths
        for col_idx, width in enumerate(col_widths):
            worksheet.set_column(col_idx, col_idx, width)

        # Row 0: Empty (for spacing)
        worksheet.set_row(0, 20)

        # Row 1: Title
        worksheet.merge_range(1, 0, 1, len(headers) - 1, sheet_title, title_format)
        worksheet.set_row(1, 30)

        # Row 2: Subtitle with count
        subtitle = f"تعداد کل اختلافات: {len(disputes_data)}"
        worksheet.merge_range(2, 0, 2, len(headers) - 1, subtitle, subtitle_format)
        worksheet.set_row(2, 25)

        # Row 3: Empty (spacing before headers)
        worksheet.set_row(3, 10)

        # Row 4: Headers (bold and distinct)
        worksheet.set_row(4, 35)
        for col_idx, header in enumerate(headers):
            worksheet.write(4, col_idx, header, header_format)

        # Freeze panes: freeze first row (header) and first column
        worksheet.freeze_panes(5, 1)

        # Data rows starting at row 5
        for row_idx, dispute in enumerate(disputes_data):
            excel_row = row_idx + 5
            is_alt_row = row_idx % 2 == 1

            # Set row height
            worksheet.set_row(excel_row, 30)

            # Select formats based on alternating rows
            txt_fmt = alt_text_format if is_alt_row else text_format
            num_fmt = alt_number_format if is_alt_row else number_format
            pct_fmt = alt_percent_format if is_alt_row else percent_format
            note_fmt = alt_notes_format if is_alt_row else notes_format

            # Column 0: Row number
            worksheet.write_number(excel_row, 0, row_idx + 1, num_fmt)

            # Column 1: WBS
            worksheet.write_string(excel_row, 1, str(dispute['کد WBS']), txt_fmt)

            # Column 2: Activity
            worksheet.write_string(excel_row, 2, str(dispute['شرح فعالیت']), txt_fmt)

            # Column 3: Sheet name
            worksheet.write_string(excel_row, 3, str(dispute['برگه']), txt_fmt)

            # Column 4: Total cost
            worksheet.write_number(excel_row, 4, dispute['هزینه کامل (ریال)'], num_fmt)

            # Column 5: Contractor progress %
            contractor_prog = dispute['درصد پیشرفت پیمانکار']
            if contractor_prog > 0:
                worksheet.write_number(excel_row, 5, contractor_prog, pct_fmt)
            else:
                worksheet.write_number(excel_row, 5, 0, pct_fmt)

            # Column 6: Contractor amount
            worksheet.write_number(excel_row, 6, dispute['مبلغ ادعای پیمانکار (ریال)'], num_fmt)

            # Column 7: Employer progress %
            employer_prog = dispute['درصد پیشرفت کارفرما']
            if employer_prog > 0:
                worksheet.write_number(excel_row, 7, employer_prog, pct_fmt)
            else:
                worksheet.write_number(excel_row, 7, 0, pct_fmt)

            # Column 8: Employer amount
            worksheet.write_number(excel_row, 8, dispute['مبلغ تایید کارفرما (ریال)'], num_fmt)

            # Column 9: Dispute amount (with conditional formatting)
            dispute_amt = dispute['مبلغ اختلاف (ریال)']
            if dispute_amt > 0:
                worksheet.write_number(excel_row, 9, dispute_amt, dispute_positive_format)
            else:
                worksheet.write_number(excel_row, 9, dispute_amt, dispute_negative_format)

            # Column 10: Contractor notes
            worksheet.write_string(excel_row, 10, str(dispute['توضیحات پیمانکار']), note_fmt)

            # Column 11: Employer notes
            worksheet.write_string(excel_row, 11, str(dispute['توضیحات مشاور']), note_fmt)

        # Summary row
        if disputes_data:
            summary_row = len(disputes_data) + 6
            worksheet.set_row(summary_row, 35)

            worksheet.write_string(summary_row, 0, 'جمع کل', summary_format)
            worksheet.merge_range(summary_row, 1, summary_row, 5, '', summary_format)

            total_contractor = sum(d['مبلغ ادعای پیمانکار (ریال)'] for d in disputes_data)
            worksheet.write_number(summary_row, 6, total_contractor, summary_format)

            worksheet.write_string(summary_row, 7, '', summary_format)

            total_employer = sum(d['مبلغ تایید کارفرما (ریال)'] for d in disputes_data)
            worksheet.write_number(summary_row, 8, total_employer, summary_format)

            total_dispute = sum(d['مبلغ اختلاف (ریال)'] for d in disputes_data)
            worksheet.write_number(summary_row, 9, total_dispute, summary_format)

            worksheet.merge_range(summary_row, 10, summary_row, 11, '', summary_format)

        # Use autofit for better column sizing
        worksheet.autofit()

    # Sheet 1: All disputes (Summary)
    ws_all = workbook.add_worksheet('خلاصه کل اختلافات')
    write_sheet(ws_all, all_disputes, 'گزارش جامع اختلافات پیمانکار و کارفرما')

    # Individual sheets per category
    for sheet_name, disputes in sheet_disputes.items():
        # Clean sheet name for Excel (max 31 chars)
        safe_name = sheet_name[:31].replace('[', '(').replace(']', ')').replace(':', '-')
        ws = workbook.add_worksheet(safe_name)
        write_sheet(ws, disputes, f'اختلافات - {sheet_name}')

    # Close workbook
    workbook.close()
    print(f"Report saved to: {OUTPUT_FILE}")


def main():
    """Main entry point."""
    print("=" * 60)
    print("  Dispute Report Generator")
    print("  Comparing Contractor vs Employer Claims")
    print("=" * 60)
    print()

    # Load and process data
    all_disputes, sheet_disputes = load_and_process_data()

    if not all_disputes:
        print("\nNo disputes found!")
        return

    print(f"\nTotal disputes found: {len(all_disputes)}")

    # Calculate totals
    total_contractor = sum(d['مبلغ ادعای پیمانکار (ریال)'] for d in all_disputes)
    total_employer = sum(d['مبلغ تایید کارفرما (ریال)'] for d in all_disputes)
    total_dispute = sum(d['مبلغ اختلاف (ریال)'] for d in all_disputes)

    print(f"Total contractor claims: {total_contractor:,.0f} Rial")
    print(f"Total employer approved: {total_employer:,.0f} Rial")
    print(f"Total dispute amount:    {total_dispute:,.0f} Rial")

    # Write formatted report
    write_formatted_excel(all_disputes, sheet_disputes)

    print()
    print("=" * 60)
    print("  Report generation complete!")
    print("=" * 60)


if __name__ == '__main__':
    main()
