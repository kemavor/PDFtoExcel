#!/usr/bin/env python3
"""
PDF to Excel Converter
Converts tables from PDF documents to Excel files
"""

import sys
import os
import argparse
import tabula
import pandas as pd


def convert_pdf_to_excel(pdf_path, excel_path=None, pages='all'):
    """
    Convert PDF tables to Excel format
    
    Args:
        pdf_path (str): Path to input PDF file
        excel_path (str): Path to output Excel file (optional)
        pages (str): Pages to extract tables from (default: 'all')
    
    Returns:
        str: Path to the generated Excel file
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")
    
    # Generate output filename if not provided
    if excel_path is None:
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        excel_path = f"{base_name}.xlsx"
    
    print(f"Converting {pdf_path} to {excel_path}...")
    
    try:
        # Extract tables from PDF
        tables = tabula.read_pdf(pdf_path, pages=pages, multiple_tables=True)
        
        if not tables:
            print("Warning: No tables found in the PDF")
            # Create an empty Excel file
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                pd.DataFrame().to_excel(writer, sheet_name='Sheet1', index=False)
            return excel_path
        
        # Write tables to Excel file
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for i, table in enumerate(tables):
                sheet_name = f'Table_{i+1}' if len(tables) > 1 else 'Sheet1'
                table.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"Successfully converted {len(tables)} table(s) to {excel_path}")
        return excel_path
        
    except Exception as e:
        raise Exception(f"Error converting PDF to Excel: {str(e)}") from e


def main():
    """Main entry point for the PDF to Excel converter"""
    parser = argparse.ArgumentParser(
        description='Convert PDF documents with tables to Excel format'
    )
    parser.add_argument(
        'pdf_file',
        help='Path to the PDF file to convert'
    )
    parser.add_argument(
        '-o', '--output',
        help='Path to the output Excel file (default: same name as PDF with .xlsx extension)',
        default=None
    )
    parser.add_argument(
        '-p', '--pages',
        help='Pages to extract (default: all). Examples: "1", "1-3", "1,3,5"',
        default='all'
    )
    
    args = parser.parse_args()
    
    try:
        output_file = convert_pdf_to_excel(args.pdf_file, args.output, args.pages)
        print(f"\n✓ Conversion complete! Excel file saved as: {output_file}")
        return 0
    except Exception as e:
        print(f"\n✗ Error: {str(e)}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
