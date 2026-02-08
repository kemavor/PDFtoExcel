#!/usr/bin/env python3
"""
Example usage of PDF to Excel converter
"""

from pdf_to_excel import convert_pdf_to_excel


def example_basic_conversion():
    """Example: Basic PDF to Excel conversion"""
    print("Example 1: Basic conversion")
    print("-" * 50)
    
    # This would convert a PDF file to Excel
    # convert_pdf_to_excel('sample.pdf')
    print("Usage: convert_pdf_to_excel('input.pdf')")
    print("Output: Creates 'input.xlsx' with extracted tables\n")


def example_custom_output():
    """Example: Custom output filename"""
    print("Example 2: Custom output filename")
    print("-" * 50)
    
    # convert_pdf_to_excel('report.pdf', 'my_data.xlsx')
    print("Usage: convert_pdf_to_excel('report.pdf', 'my_data.xlsx')")
    print("Output: Creates 'my_data.xlsx' with extracted tables\n")


def example_specific_pages():
    """Example: Extract from specific pages"""
    print("Example 3: Extract from specific pages")
    print("-" * 50)
    
    # convert_pdf_to_excel('document.pdf', pages='1-3')
    print("Usage: convert_pdf_to_excel('document.pdf', pages='1-3')")
    print("Output: Extracts tables only from pages 1-3\n")


if __name__ == '__main__':
    print("=" * 50)
    print("PDF to Excel Converter - Examples")
    print("=" * 50)
    print()
    
    example_basic_conversion()
    example_custom_output()
    example_specific_pages()
    
    print("=" * 50)
    print("For actual conversion, provide a real PDF file")
    print("=" * 50)
