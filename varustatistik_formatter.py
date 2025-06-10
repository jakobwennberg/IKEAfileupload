#!/usr/bin/env python3
"""
Varustatistik Excel Formatter - External Forecast Format
Formats Swedish restaurant statistics Excel files into external forecast data format.

Usage:
    python varustatistik_formatter.py input_file.xlsx [output_file.txt]
"""

import pandas as pd
import re
import sys
from datetime import datetime
from typing import List, Tuple, Dict
import warnings

# Suppress openpyxl warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def get_variable_id(category: str) -> str:
    """
    Map category to external forecast variable ID.
    Café, Sallad -> kallmat
    Food (kitchen printer) -> varmmat
    """
    category_lower = category.lower()
    
    # Handle different variations of category names
    if 'café' in category_lower or 'sallad' in category_lower:
        return 'kallmat'
    elif 'food' in category_lower and ('kitchen' in category_lower or 'printer' in category_lower):
        return 'varmmat'
    elif 'food' in category_lower:  # Handle cases where it's just "Food"
        return 'varmmat'
    
    return ''  # Unknown category


def process_excel_file(filename: str) -> List[Dict]:
    """
    Process the Excel file and extract hourly totals.
    """
    results = []
    
    # Read Excel file
    xl_file = pd.ExcelFile(filename)
    
    for sheet_name in xl_file.sheet_names:
        # Skip sheets that don't look like date sheets (YYYY-MM format)
        if not re.match(r'^\d{4}-\d{2}$', sheet_name) and sheet_name != 'Blad1':
            continue
        
        if sheet_name == 'Blad1':
            continue
            
        print(f"Processing sheet: {sheet_name}")
        
        # Read sheet without headers to get raw data
        df = pd.read_excel(xl_file, sheet_name=sheet_name, header=None)
        
        # Process each row
        for idx, row in df.iterrows():
            # Check if first cell contains "Totalt" pattern
            if pd.notna(row[0]) and isinstance(row[0], str) and row[0].startswith('Totalt'):
                # Extract date, category, and hour using regex
                match = re.match(r'Totalt (\d{4}-\d{2}-\d{2}) (.+?) Kl: (\d{2})', row[0])
                
                if match:
                    date_str, category, hour_str = match.groups()
                    
                    # Get antal (quantity) from column 4 (0-indexed)
                    antal = row[4] if pd.notna(row[4]) else 0
                    
                    # Skip if antal is empty or 0
                    if antal == '' or pd.isna(antal):
                        continue
                    
                    # Get variable ID based on category
                    variable_id = get_variable_id(category)
                    
                    if variable_id:
                        # Format hour as HH:00:00
                        hour_formatted = f"{hour_str}:00:00"
                        
                        # Add to results
                        results.append({
                            'external_forecast_variable_id': variable_id,
                            'external_forecast_configuration_id': '',  # Not provided
                            'external_unit_id': 'produktion',
                            'external_section_id': 'köket',
                            'date': date_str,
                            'hour': hour_formatted,
                            'value': float(antal)
                        })
    
    # Sort results by date and hour
    results.sort(key=lambda x: (x['date'], x['hour']))
    
    return results


def format_output(results: List[Dict]) -> str:
    """
    Format results into the required external forecast format.
    """
    # Header with asterisks indicating required fields
    lines = [
        'External_forecast_variable_ID\tExternal_forecast_configuration_ID\tExternal_unit_ID \tExternal_section_ID\tDate_YYYY-MM-DD\tHour_00:00:00\tValue'
    ]
    
    for item in results:
        line = f"{item['external_forecast_variable_id']}\t{item['external_forecast_configuration_id']}\t{item['external_unit_id']}\t{item['external_section_id']}\t{item['date']}\t{item['hour']}\t{item['value']}"
        lines.append(line)
    
    return '\n'.join(lines)


def main():
    """
    Main function to process the file.
    """
    if len(sys.argv) < 2:
        print("Usage: python varustatistik_formatter.py input_file.xlsx [output_file.txt]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'external_forecast_output.txt'
    
    try:
        print(f"Reading file: {input_file}")
        results = process_excel_file(input_file)
        
        print(f"Found {len(results)} hourly totals")
        
        # Format output
        output = format_output(results)
        
        # Write to file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(output)
        
        print(f"Output written to: {output_file}")
        
        # Show summary
        if results:
            print(f"\nSummary:")
            print(f"- Total records: {len(results)}")
            print(f"- Date range: {results[0]['date']} to {results[-1]['date']}")
            
            # Count by variable ID
            kallmat_count = sum(1 for r in results if r['external_forecast_variable_id'] == 'kallmat')
            varmmat_count = sum(1 for r in results if r['external_forecast_variable_id'] == 'varmmat')
            print(f"- Kallmat records: {kallmat_count}")
            print(f"- Varmmat records: {varmmat_count}")
        
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()


# Alternative: Function version for use in other scripts or notebooks
def format_varustatistik_external(input_file: str, output_file: str = None) -> str:
    """
    Format a varustatistik Excel file into external forecast format.
    
    Args:
        input_file: Path to the Excel file
        output_file: Optional path to save the output (if not provided, returns string only)
    
    Returns:
        Formatted data as string
    """
    results = process_excel_file(input_file)
    output = format_output(results)
    
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(output)
    
    return output