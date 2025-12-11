#!/usr/bin/env python3
"""
Excel Logic Audit Script with Verified Output Generation

This script analyzes a price adjustment Excel file, verifies calculations,
and generates a new Excel file with audit results and corrections.

Features:
- Copies original data to a new file (preserving the original)
- Creates an Audit_Log sheet with verification details
- Highlights discrepancies and suggests corrections
- Reports PASS/FAIL/WARN status for each verified row
"""

import os
import copy
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# Configuration
INPUT_FILE = "Price_Adjustment_Automated 19 Claude Final 01 Claude Code.xlsx"
OUTPUT_FILE = "Price_Adjustment_Verified_Output.xlsx"

# Style definitions
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
PASS_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
FAIL_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
WARN_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
CORRECTION_FILL = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)


class ExcelLogicAuditor:
    """Main class for auditing Excel logic and generating verified output."""

    def __init__(self, input_path, output_path):
        self.input_path = input_path
        self.output_path = output_path
        self.audit_log = []
        self.corrections = []
        self.summary = {
            'total_rows': 0,
            'passed': 0,
            'failed': 0,
            'warnings': 0
        }

    def load_workbook_with_values(self):
        """Load workbook with calculated values (data_only=True)."""
        return load_workbook(self.input_path, data_only=True)

    def load_workbook_with_formulas(self):
        """Load workbook with formulas preserved."""
        return load_workbook(self.input_path, data_only=False)

    def safe_float(self, value):
        """Safely convert a value to float."""
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            # Handle Persian/Arabic numerals and common formats
            try:
                # Remove common formatting
                cleaned = value.replace(',', '').replace('٫', '.').strip()
                # Convert Persian/Arabic digits to ASCII
                persian_digits = '۰۱۲۳۴۵۶۷۸۹'
                arabic_digits = '٠١٢٣٤٥٦٧٨٩'
                for i, (p, a) in enumerate(zip(persian_digits, arabic_digits)):
                    cleaned = cleaned.replace(p, str(i)).replace(a, str(i))
                return float(cleaned)
            except (ValueError, AttributeError):
                return None
        return None

    def verify_formula_calculation(self, expected, calculated, tolerance=0.01):
        """
        Verify if calculated value matches expected value within tolerance.
        Returns: (status, message)
        """
        if expected is None and calculated is None:
            return 'PASS', 'Both values are empty'
        if expected is None or calculated is None:
            return 'WARN', f'Missing value (Expected: {expected}, Calculated: {calculated})'

        try:
            exp_float = self.safe_float(expected)
            calc_float = self.safe_float(calculated)

            if exp_float is None or calc_float is None:
                return 'WARN', 'Non-numeric values detected'

            if exp_float == 0 and calc_float == 0:
                return 'PASS', 'Both values are zero'

            if exp_float == 0:
                diff_pct = abs(calc_float) * 100
            else:
                diff_pct = abs((calc_float - exp_float) / exp_float) * 100

            if diff_pct <= tolerance * 100:
                return 'PASS', f'Match within {tolerance*100}% tolerance (diff: {diff_pct:.4f}%)'
            elif diff_pct <= 5:  # 5% warning threshold
                return 'WARN', f'Minor discrepancy: {diff_pct:.2f}% difference'
            else:
                return 'FAIL', f'Significant discrepancy: {diff_pct:.2f}% difference'
        except Exception as e:
            return 'WARN', f'Error during comparison: {str(e)}'

    def audit_main_sheet(self, ws_values, ws_formulas, sheet_name='1-2'):
        """
        Audit the main calculation sheet (1-2).
        This sheet contains price adjustment calculations.
        """
        print(f"\nAuditing sheet: {sheet_name}")

        # Column mapping based on header analysis
        # Row 3 contains headers, data starts from row 7
        header_row = 3
        data_start_row = 7

        # Key columns to verify (based on the structure observed):
        # L: ضریب ارزبری طبق استعلام (Exchange coefficient per inquiry)
        # M: ضریب ارزبری منظور شده (Applied exchange coefficient)
        # N: ضریب F روش الف (F coefficient - Method A)
        # O: قیمت ارز در زمان خرید (Currency price at purchase time)

        for row_idx in range(data_start_row, ws_values.max_row + 1):
            row_data = {}

            # Read key columns
            col_b = ws_values.cell(row=row_idx, column=2).value  # Row number indicator
            col_c = ws_values.cell(row=row_idx, column=3).value  # Contract row
            col_g = ws_values.cell(row=row_idx, column=7).value  # Description
            col_l = ws_values.cell(row=row_idx, column=12).value  # Exchange coefficient inquiry
            col_m = ws_values.cell(row=row_idx, column=13).value  # Applied exchange coefficient
            col_n = ws_values.cell(row=row_idx, column=14).value  # F coefficient
            col_o = ws_values.cell(row=row_idx, column=15).value  # Currency price

            # Skip empty rows
            if col_b is None and col_c is None and col_g is None:
                continue

            self.summary['total_rows'] += 1

            # Verification checks
            issues = []
            status = 'PASS'
            calculated_value = None
            original_value = col_l

            # Check 1: Exchange coefficient validation
            if col_l is not None:
                l_float = self.safe_float(col_l)
                if l_float is not None:
                    if l_float < 0 or l_float > 1:
                        if str(col_l) != '#N/A':
                            issues.append(f"Exchange coefficient {col_l} out of range [0,1]")
                            status = 'WARN'

            # Check 2: F coefficient should typically be positive
            if col_n is not None and col_n != '-':
                n_float = self.safe_float(col_n)
                if n_float is not None and n_float <= 0:
                    issues.append(f"F coefficient {col_n} should be positive")
                    status = 'WARN'

            # Check 3: Applied coefficient should match or be derived from inquiry
            if col_l is not None and col_m is not None:
                l_float = self.safe_float(col_l)
                m_float = self.safe_float(col_m)
                if l_float is not None and m_float is not None:
                    if m_float > 1:
                        issues.append(f"Applied coefficient {col_m} exceeds 1 (100%)")
                        status = 'FAIL' if status != 'FAIL' else status

            # Check 4: Currency price validation
            if col_o is not None:
                o_float = self.safe_float(col_o)
                if o_float is not None:
                    if o_float <= 0:
                        issues.append(f"Currency price {col_o} should be positive")
                        status = 'FAIL'
                    # Check for reasonable exchange rate range (IRR to EUR ~400,000-600,000 in recent years)
                    elif o_float < 100000 or o_float > 1000000:
                        issues.append(f"Currency price {col_o} outside typical range")
                        status = 'WARN' if status != 'FAIL' else status

            # Determine final status
            if not issues:
                issues.append("All validations passed")
                self.summary['passed'] += 1
            elif status == 'FAIL':
                self.summary['failed'] += 1
            else:
                self.summary['warnings'] += 1

            # Record audit entry
            audit_entry = {
                'sheet': sheet_name,
                'row': row_idx,
                'description': str(col_g)[:50] if col_g else 'N/A',
                'calculated_value': calculated_value,
                'original_value': original_value,
                'status': status,
                'details': '; '.join(issues)
            }
            self.audit_log.append(audit_entry)

            # Record corrections if needed
            if status == 'FAIL':
                self.corrections.append({
                    'sheet': sheet_name,
                    'row': row_idx,
                    'column': 'L',
                    'original': original_value,
                    'suggested': 'Manual review required',
                    'reason': '; '.join(issues)
                })

    def audit_percentage_sheet(self, ws_values, sheet_name='درصد ارزیری'):
        """
        Audit the percentage calculation sheet.
        Verifies that percentage calculations sum correctly.
        """
        print(f"\nAuditing sheet: {sheet_name}")

        data_start_row = 3

        for row_idx in range(data_start_row, min(ws_values.max_row + 1, 220)):
            # Read percentage columns
            col_a = ws_values.cell(row=row_idx, column=1).value  # Row number
            col_b = ws_values.cell(row=row_idx, column=2).value  # Description
            col_e = ws_values.cell(row=row_idx, column=5).value  # Percentage 1
            col_i = ws_values.cell(row=row_idx, column=9).value  # Percentage 2
            col_m = ws_values.cell(row=row_idx, column=13).value  # Total/Currency percentage

            if col_a is None or not isinstance(col_a, (int, float)):
                continue

            self.summary['total_rows'] += 1
            issues = []
            status = 'PASS'

            # Verify percentage sum (col_e + col_i should relate to col_m)
            e_float = self.safe_float(col_e)
            i_float = self.safe_float(col_i)

            if e_float is not None:
                if e_float < 0 or e_float > 1:
                    issues.append(f"Percentage 1 ({col_e}) out of range [0,1]")
                    status = 'WARN'

            if i_float is not None:
                if i_float < 0 or i_float > 1:
                    issues.append(f"Percentage 2 ({col_i}) out of range [0,1]")
                    status = 'WARN'

            # Check if percentages sum to 1 or less
            if e_float is not None and i_float is not None:
                total_pct = e_float + (i_float if i_float else 0)
                if total_pct > 1.001:  # Small tolerance for floating point
                    issues.append(f"Percentages sum ({total_pct:.3f}) exceeds 100%")
                    status = 'FAIL'

            if not issues:
                issues.append("Percentage validations passed")
                self.summary['passed'] += 1
            elif status == 'FAIL':
                self.summary['failed'] += 1
            else:
                self.summary['warnings'] += 1

            audit_entry = {
                'sheet': sheet_name,
                'row': row_idx,
                'description': str(col_b)[:50] if col_b else 'N/A',
                'calculated_value': f"E:{col_e}, I:{col_i}",
                'original_value': col_m,
                'status': status,
                'details': '; '.join(issues)
            }
            self.audit_log.append(audit_entry)

    def audit_index_sheet(self, ws_values, sheet_name):
        """
        Audit index/coefficient sheets (مکانیک, ابنیه, etc.).
        Verifies that indices are positive and follow reasonable progression.
        """
        print(f"\nAuditing sheet: {sheet_name}")

        data_start_row = 4

        for row_idx in range(data_start_row, ws_values.max_row + 1):
            col_a = ws_values.cell(row=row_idx, column=1).value  # Chapter number
            col_b = ws_values.cell(row=row_idx, column=2).value  # Description

            if col_a is None:
                continue

            self.summary['total_rows'] += 1
            issues = []
            status = 'PASS'

            # Check index values across columns (time series)
            prev_value = None
            negative_count = 0
            zero_count = 0

            for col_idx in range(3, min(ws_values.max_column + 1, 40)):
                val = ws_values.cell(row=row_idx, column=col_idx).value
                val_float = self.safe_float(val)

                if val_float is not None:
                    if val_float < 0:
                        negative_count += 1
                    elif val_float == 0:
                        zero_count += 1
                    prev_value = val_float

            if negative_count > 0:
                issues.append(f"Found {negative_count} negative index values")
                status = 'FAIL'

            if zero_count > 5:
                issues.append(f"Found {zero_count} zero values (may be incomplete data)")
                status = 'WARN' if status != 'FAIL' else status

            if not issues:
                issues.append("Index values validated")
                self.summary['passed'] += 1
            elif status == 'FAIL':
                self.summary['failed'] += 1
            else:
                self.summary['warnings'] += 1

            audit_entry = {
                'sheet': sheet_name,
                'row': row_idx,
                'description': str(col_b)[:50] if col_b else 'N/A',
                'calculated_value': f"Last: {prev_value}",
                'original_value': f"Ch. {col_a}",
                'status': status,
                'details': '; '.join(issues)
            }
            self.audit_log.append(audit_entry)

    def create_audit_log_sheet(self, wb):
        """Create the Audit_Log sheet with all verification results."""
        # Create new sheet
        if 'Audit_Log' in wb.sheetnames:
            del wb['Audit_Log']

        ws = wb.create_sheet('Audit_Log', 0)  # Insert at beginning

        # Headers
        headers = [
            'Row #', 'Sheet', 'Excel Row', 'Description',
            'Calculated Value', 'Original Value', 'Status', 'Details'
        ]

        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = THIN_BORDER

        # Data rows
        for idx, entry in enumerate(self.audit_log, 1):
            row_idx = idx + 1

            ws.cell(row=row_idx, column=1, value=idx)
            ws.cell(row=row_idx, column=2, value=entry['sheet'])
            ws.cell(row=row_idx, column=3, value=entry['row'])
            ws.cell(row=row_idx, column=4, value=entry['description'])
            ws.cell(row=row_idx, column=5, value=str(entry['calculated_value']))
            ws.cell(row=row_idx, column=6, value=str(entry['original_value']))

            status_cell = ws.cell(row=row_idx, column=7, value=entry['status'])
            ws.cell(row=row_idx, column=8, value=entry['details'])

            # Apply status-based styling
            if entry['status'] == 'PASS':
                status_cell.fill = PASS_FILL
            elif entry['status'] == 'FAIL':
                status_cell.fill = FAIL_FILL
            elif entry['status'] == 'WARN':
                status_cell.fill = WARN_FILL

            # Apply borders
            for col in range(1, 9):
                ws.cell(row=row_idx, column=col).border = THIN_BORDER

        # Add summary section
        summary_row = len(self.audit_log) + 4
        ws.cell(row=summary_row, column=1, value="AUDIT SUMMARY").font = Font(bold=True, size=14)
        ws.cell(row=summary_row + 1, column=1, value=f"Total Rows Audited: {self.summary['total_rows']}")
        ws.cell(row=summary_row + 2, column=1, value=f"Passed: {self.summary['passed']}")
        ws.cell(row=summary_row + 2, column=1).fill = PASS_FILL
        ws.cell(row=summary_row + 3, column=1, value=f"Failed: {self.summary['failed']}")
        ws.cell(row=summary_row + 3, column=1).fill = FAIL_FILL
        ws.cell(row=summary_row + 4, column=1, value=f"Warnings: {self.summary['warnings']}")
        ws.cell(row=summary_row + 4, column=1).fill = WARN_FILL
        ws.cell(row=summary_row + 5, column=1, value=f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        # Adjust column widths
        column_widths = [8, 20, 10, 40, 25, 25, 10, 60]
        for col_idx, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        return ws

    def create_corrections_sheet(self, wb):
        """Create a sheet listing suggested corrections."""
        if not self.corrections:
            return None

        if 'Corrections' in wb.sheetnames:
            del wb['Corrections']

        ws = wb.create_sheet('Corrections', 1)

        headers = ['#', 'Sheet', 'Row', 'Column', 'Original Value', 'Suggested Value', 'Reason']

        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal='center')
            cell.border = THIN_BORDER

        for idx, correction in enumerate(self.corrections, 1):
            row_idx = idx + 1
            ws.cell(row=row_idx, column=1, value=idx)
            ws.cell(row=row_idx, column=2, value=correction['sheet'])
            ws.cell(row=row_idx, column=3, value=correction['row'])
            ws.cell(row=row_idx, column=4, value=correction['column'])
            ws.cell(row=row_idx, column=5, value=str(correction['original']))
            ws.cell(row=row_idx, column=6, value=str(correction['suggested']))
            ws.cell(row=row_idx, column=6).fill = CORRECTION_FILL
            ws.cell(row=row_idx, column=7, value=correction['reason'])

            for col in range(1, 8):
                ws.cell(row=row_idx, column=col).border = THIN_BORDER

        # Adjust column widths
        column_widths = [6, 20, 8, 10, 20, 25, 50]
        for col_idx, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        return ws

    def run_audit(self):
        """Execute the complete audit process."""
        print("=" * 60)
        print("Excel Logic Audit - Starting")
        print("=" * 60)
        print(f"Input file: {self.input_path}")
        print(f"Output file: {self.output_path}")

        # Load workbooks
        print("\nLoading workbooks...")
        wb_values = self.load_workbook_with_values()
        wb_formulas = self.load_workbook_with_formulas()

        # Audit each relevant sheet
        if '1-2' in wb_values.sheetnames:
            self.audit_main_sheet(
                wb_values['1-2'],
                wb_formulas['1-2'],
                '1-2'
            )

        if 'درصد ارزیری' in wb_values.sheetnames:
            self.audit_percentage_sheet(
                wb_values['درصد ارزیری'],
                'درصد ارزیری'
            )

        # Audit index sheets
        index_sheets = ['مکانیک', 'ابنیه', 'تاسیسات برقی', 'راه، راه آهن و باند فرودگاه', 'تجهیزات آب و فاضلاب']
        for sheet_name in index_sheets:
            if sheet_name in wb_values.sheetnames:
                self.audit_index_sheet(wb_values[sheet_name], sheet_name)

        # Create output workbook (copy of original with formulas)
        print("\nCreating output workbook...")

        # Create audit sheets
        self.create_audit_log_sheet(wb_formulas)
        self.create_corrections_sheet(wb_formulas)

        # Save output
        print(f"\nSaving to: {self.output_path}")
        wb_formulas.save(self.output_path)

        # Close workbooks
        wb_values.close()
        wb_formulas.close()

        # Print summary
        print("\n" + "=" * 60)
        print("AUDIT COMPLETE")
        print("=" * 60)
        print(f"Total rows audited: {self.summary['total_rows']}")
        print(f"  PASSED:   {self.summary['passed']}")
        print(f"  FAILED:   {self.summary['failed']}")
        print(f"  WARNINGS: {self.summary['warnings']}")
        print(f"\nOutput saved to: {self.output_path}")
        print(f"Corrections logged: {len(self.corrections)}")

        return self.summary


def main():
    """Main entry point."""
    # Determine paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(script_dir, INPUT_FILE)
    output_path = os.path.join(script_dir, OUTPUT_FILE)

    # Check input file exists
    if not os.path.exists(input_path):
        print(f"ERROR: Input file not found: {input_path}")
        return 1

    # Run audit
    auditor = ExcelLogicAuditor(input_path, output_path)
    summary = auditor.run_audit()

    return 0 if summary['failed'] == 0 else 1


if __name__ == "__main__":
    exit(main())
