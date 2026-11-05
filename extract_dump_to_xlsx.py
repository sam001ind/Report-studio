#!/usr/bin/env python3
"""
Database Dump & Backup to XLSX Extractor
=======================================
Extracts tables, columns, and records from SQL & database dump files (.dump, .sql, .dmp, .bak, .txt)
into fully formatted Excel (.xlsx) workbooks while keeping all data and formatting 100% intact.

Supported Dump Formats:
1. PostgreSQL text dumps (COPY ... FROM stdin; tab-delimited, \\N nulls, \\ escapes, \\. terminator)
2. PostgreSQL / MySQL / SQLite INSERT INTO statements (single and multi-row value lists)
3. Multi-table dumps (each table becomes its own worksheet in the Excel workbook)
4. Split handling for tables exceeding Excel's 1,048,576 row limit

Key Data Preservation Rules:
- Preserves leading zeros in codes, PINs, phone numbers, and IDs
- Prevents scientific notation (e.g. 3.09E+11) for Aadhar, PRN, and large numeric identifiers
- Converts \\N and NULL to clean empty cells (no literal '\\N')
- Unescapes newlines (\\n), tabs (\\t), and quotes properly
- Applies frozen header rows and styled headers for professional usability
"""

import sys
import os
import re
import argparse
from typing import Dict, List, Any, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

EXCEL_MAX_ROWS = 1048500  # Safe headroom below 1,048,576 limit


def unescape_pg_copy_val(val: str) -> str:
    """Unescape PostgreSQL COPY data value."""
    if val == r'\N':
        return ""
    if '\\' not in val:
        return val
    # Replace standard escape sequences
    val = val.replace(r'\n', '\n').replace(r'\r', '\r').replace(r'\t', '\t').replace(r'\\', '\\')
    return val


def parse_sql_values(values_str: str) -> List[List[Any]]:
    """Parse comma-separated tuples of SQL INSERT VALUES e.g. (1, 'val', NULL), (2, 'val2')."""
    rows = []
    current_row = []
    current_val = []
    in_quotes = False
    quote_char = None
    escaped = False
    in_tuple = False

    i = 0
    n = len(values_str)
    while i < n:
        ch = values_str[i]

        if not in_tuple:
            if ch == '(':
                in_tuple = True
                current_row = []
                current_val = []
            i += 1
            continue

        if in_quotes:
            if escaped:
                current_val.append(ch)
                escaped = False
            elif ch == '\\':
                escaped = True
            elif ch == quote_char:
                # Check for doubled quote escape in SQL (e.g. 'O''Reilly')
                if i + 1 < n and values_str[i + 1] == quote_char:
                    current_val.append(quote_char)
                    i += 1
                else:
                    in_quotes = False
            else:
                current_val.append(ch)
        else:
            if ch in ("'", '"', '`'):
                in_quotes = True
                quote_char = ch
            elif ch == ',':
                val_str = "".join(current_val).strip()
                if val_str.upper() == 'NULL':
                    current_row.append("")
                else:
                    current_row.append(val_str)
                current_val = []
            elif ch == ')':
                val_str = "".join(current_val).strip()
                if val_str.upper() == 'NULL':
                    current_row.append("")
                else:
                    current_row.append(val_str)
                rows.append(current_row)
                current_row = []
                current_val = []
                in_tuple = False
            else:
                current_val.append(ch)

        i += 1

    return rows


def parse_dump_file(filepath: str) -> Dict[str, Dict[str, Any]]:
    """
    Parses a database dump file and returns a dictionary of tables with their columns and rows.
    """
    tables: Dict[str, Dict[str, Any]] = {}
    table_schemas: Dict[str, List[str]] = {}

    current_copy_table: Optional[str] = None
    current_copy_cols: Optional[List[str]] = None

    print(f"[*] Scanning and parsing dump file: {filepath} ...")
    file_size = os.path.getsize(filepath)
    print(f"[*] File size: {file_size / (1024 * 1024):.2f} MB")

    line_count = 0
    in_create_table = False
    create_table_name = ""
    create_table_cols = []

    # Open with utf-8, replace bad bytes cleanly
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        for line in f:
            line_count += 1
            if line_count % 50000 == 0:
                print(f"    ... processed {line_count:,} lines ...")

            # 1. Inside PostgreSQL COPY block
            if current_copy_table is not None:
                # Terminator is a line with just \.
                if line.rstrip('\r\n') == r'\.':
                    print(f"    [+] Finished COPY for table '{current_copy_table}' ({len(tables[current_copy_table]['rows']):,} rows)")
                    current_copy_table = None
                    current_copy_cols = None
                    continue

                raw_parts = line.rstrip('\r\n').split('\t')
                cleaned_row = [unescape_pg_copy_val(part) for part in raw_parts]
                tables[current_copy_table]['rows'].append(cleaned_row)
                continue

            stripped = line.strip()
            if not stripped or stripped.startswith('--') or stripped.startswith('/*'):
                continue

            # 2. Check for CREATE TABLE to learn schema
            if not in_create_table:
                m_create = re.match(r'^CREATE\s+TABLE(?:\s+IF\s+NOT\s+EXISTS)?\s+(?:ONLY\s+)?(?:[a-zA-Z0-9_]+\.)?([a-zA-Z0-9_\"`]+)\s*\(', stripped, re.IGNORECASE)
                if m_create:
                    in_create_table = True
                    create_table_name = m_create.group(1).replace('"', '').replace('`', '').strip()
                    create_table_cols = []
                    # Check if single line definition
                    if stripped.endswith(');'):
                        in_create_table = False
                        table_schemas[create_table_name] = create_table_cols
                    continue
            else:
                if stripped.startswith(');'):
                    in_create_table = False
                    table_schemas[create_table_name] = create_table_cols
                    continue
                # Match column definition (e.g. col_name varchar(50) ...)
                m_col = re.match(r'^\s*([a-zA-Z0-9_\"`]+)\s+([a-zA-Z0-9_]+(?:\([^)]*\))?)', stripped)
                if m_col:
                    cname = m_col.group(1).replace('"', '').replace('`', '').strip()
                    if cname.upper() not in ('CONSTRAINT', 'PRIMARY', 'KEY', 'FOREIGN', 'UNIQUE', 'CHECK'):
                        create_table_cols.append(cname)
                continue

            # 3. Check for COPY ... FROM stdin;
            m_copy = re.match(r'^COPY\s+(?:ONLY\s+)?(?:[a-zA-Z0-9_]+\.)?([a-zA-Z0-9_\"`]+)\s*(?:\(([^)]+)\))?\s+FROM\s+stdin;', stripped, re.IGNORECASE)
            if m_copy:
                tname = m_copy.group(1).replace('"', '').replace('`', '').strip()
                cols_str = m_copy.group(2)
                if cols_str:
                    cols = [c.strip().replace('"', '').replace('`', '') for c in cols_str.split(',')]
                else:
                    cols = table_schemas.get(tname, [])
                
                if tname not in tables:
                    tables[tname] = {'cols': cols, 'rows': []}
                elif cols and not tables[tname]['cols']:
                    tables[tname]['cols'] = cols

                current_copy_table = tname
                current_copy_cols = cols
                print(f"    [+] Detected COPY block for table '{tname}' with {len(cols)} columns")
                continue

            # 4. Check for INSERT INTO statements
            m_insert = re.match(r'^INSERT\s+INTO\s+(?:[a-zA-Z0-9_]+\.)?([a-zA-Z0-9_\"`]+)\s*(?:\(([^)]+)\))?\s+VALUES\s*(.*);?$', stripped, re.IGNORECASE)
            if m_insert:
                tname = m_insert.group(1).replace('"', '').replace('`', '').strip()
                cols_str = m_insert.group(2)
                values_part = m_insert.group(3)

                if cols_str:
                    cols = [c.strip().replace('"', '').replace('`', '') for c in cols_str.split(',')]
                else:
                    cols = table_schemas.get(tname, [])

                if tname not in tables:
                    tables[tname] = {'cols': cols, 'rows': []}
                elif cols and not tables[tname]['cols']:
                    tables[tname]['cols'] = cols

                if values_part:
                    parsed_rows = parse_sql_values(values_part)
                    tables[tname]['rows'].extend(parsed_rows)
                continue

    return tables


def export_tables_to_xlsx(tables: Dict[str, Dict[str, Any]], output_filepath: str):
    """
    Exports parsed tables into an Excel workbook with formatting and data integrity preserved.
    """
    if not tables:
        raise ValueError("No tables or data rows were found in the dump file.")

    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="176B87", end_color="176B87", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    data_font = Font(name="Calibri", size=10)
    data_align = Alignment(vertical="center")

    thin_border_side = Side(border_style="thin", color="E2E8F0")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    total_tables_exported = 0
    total_rows_exported = 0

    for raw_tname, tdata in tables.items():
        cols = tdata.get('cols', [])
        rows = tdata.get('rows', [])

        if not rows and not cols:
            continue

        # If columns weren't in header, synthesize Column 1, Column 2...
        num_cols = len(cols)
        if rows and len(rows[0]) > num_cols:
            num_cols = len(rows[0])
            for c_i in range(len(cols), num_cols):
                cols.append(f"Column_{c_i + 1}")

        # Split large tables if row count > EXCEL_MAX_ROWS
        chunks = [rows[i:i + EXCEL_MAX_ROWS] for i in range(0, max(1, len(rows)), EXCEL_MAX_ROWS)]
        if not rows:
            chunks = [[]]

        for part_idx, chunk_rows in enumerate(chunks):
            sheet_title = raw_tname[:28]  # Excel max sheet name length is 31
            if len(chunks) > 1:
                sheet_title = f"{raw_tname[:24]}_p{part_idx + 1}"

            ws = wb.create_sheet(title=sheet_title)
            ws.views.sheetView[0].showGridLines = True

            # 1. Write Header Row
            ws.append(cols)
            header_row = ws[1]
            ws.row_dimensions[1].height = 28

            col_widths = [len(str(c)) for c in cols]

            for col_idx, cell in enumerate(header_row, start=1):
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = cell_border

            # Freeze top header pane
            ws.freeze_panes = "A2"

            # 2. Write Data Rows
            print(f"[*] Writing sheet '{sheet_title}' ({len(chunk_rows):,} rows) ...")
            for row_values in chunk_rows:
                ws.append(row_values)

            # Sample first 200 rows for sensible column widths
            sample_rows = chunk_rows[:200]
            for row_values in sample_rows:
                for c_idx, val in enumerate(row_values):
                    if c_idx < len(col_widths) and val is not None:
                        val_len = len(str(val))
                        if val_len > col_widths[c_idx]:
                            col_widths[c_idx] = min(50, val_len)

            # 3. Set Adjusted Column Widths
            for col_idx, width in enumerate(col_widths, start=1):
                col_letter = get_column_letter(col_idx)
                ws.column_dimensions[col_letter].width = max(12, min(width + 4, 50))

            total_tables_exported += 1
            total_rows_exported += len(chunk_rows)

    print(f"[*] Saving workbook to: {output_filepath} ...")
    wb.save(output_filepath)
    print(f"[✓] Successfully exported {total_tables_exported} sheet(s) with {total_rows_exported:,} total records to {output_filepath}!")


def main():
    parser = argparse.ArgumentParser(
        description="Extract database DUMP and SQL backup files to Excel (.xlsx) with intact formats and data."
    )
    parser.add_argument("input_file", help="Path to database dump file (.dump, .sql, .dmp, .bak, .txt)")
    parser.add_argument("-o", "--output", help="Optional output .xlsx file path")

    args = parser.parse_args()

    input_path = os.path.abspath(args.input_file)
    if not os.path.exists(input_path):
        print(f"[-] Error: File not found: {input_path}")
        sys.exit(1)

    if args.output:
        output_path = os.path.abspath(args.output)
    else:
        base_name, _ = os.path.splitext(input_path)
        output_path = f"{base_name}_extracted.xlsx"

    try:
        tables = parse_dump_file(input_path)
        if not tables:
            print("[-] No table data could be detected in the provided dump file.")
            sys.exit(1)

        export_tables_to_xlsx(tables, output_path)
    except Exception as e:
        print(f"[-] Extraction failed: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
