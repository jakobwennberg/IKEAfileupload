#!/usr/bin/env python3
"""
Varustatistik Excel Formatter - Updated External Forecast Format
Formats Swedish restaurant statistics Excel files into external forecast data format.
Aggregates multiple entries with the same timestamp and variable ID.

Usage:
    python varustatistik_formatter.py input_file.xlsx [output_file.txt]
"""

import pandas as pd
import re
import sys
from datetime import datetime
from typing import List, Tuple, Dict
import warnings
from collections import defaultdict

# Suppress openpyxl warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def get_swedish_timezone_offset(date_str: str) -> str:
    """
    Determine Swedish timezone offset (+01:00 or +02:00) based on date.
    Sweden uses CEST (+02:00) from last Sunday in March to last Sunday in October.
    Returns in format +HH:00
    """
    year, month, day = map(int, date_str.split('-'))
    date = datetime(year, month, day)
    
    # Find last Sunday in March
    march_last = datetime(year, 3, 31)
    while march_last.weekday() != 6:  # 6 = Sunday
        march_last = march_last.replace(day=march_last.day - 1)
    
    # Find last Sunday in October
    october_last = datetime(year, 10, 31)
    while october_last.weekday() != 6:
        october_last = october_last.replace(day=october_last.day - 1)
    
    # Check if date is in DST period
    if march_last <= date < october_last:
        return '+02:00'
    return '+01:00'


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
    Aggregates values for the same timestamp and variable ID.
    """
    # Use a dictionary to aggregate values by key
    aggregated_data = defaultdict(lambda: {
        'value': 0.0,
        'externalForecastConfigurationId': '',
        'Unit integration key': 'produktion',
        'Section integration key': 'köket'
    })
    
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
                        # Format time as HH:00:00
                        time_formatted = f"{hour_str}:00:00"
                        
                        # Get timezone offset
                        timezone = get_swedish_timezone_offset(date_str)
                        
                        # Create unique key for aggregation
                        key = (date_str, time_formatted, timezone, variable_id)
                        
                        # Aggregate the value
                        aggregated_data[key]['value'] += float(antal)
    
    # Convert aggregated data to list of dictionaries
    results = []
    for (date_str, time_formatted, timezone, variable_id), data in aggregated_data.items():
        results.append({
            'date': date_str,
            'time': time_formatted,
            'timezone': timezone,
            'value': data['value'],
            'externalForecastVariableId': variable_id,
            'externalForecastConfigurationId': data['externalForecastConfigurationId'],
            'Unit integration key': data['Unit integration key'],
            'Section integration key': data['Section integration key']
        })
    
    # Sort results by date and time
    results.sort(key=lambda x: (x['date'], x['time'], x['externalForecastVariableId']))
    
    return results


def format_output(results: List[Dict]) -> str:
    """
    Format results into the required external forecast format.
    """
    # Header with asterisks indicating required fields
    lines = [
        'date *\ttime *\ttimezone *\tvalue *\texternalForecastVariableId *\texternalForecastConfigurationId\tUnit integration key\tSection integration key'
    ]
    
    for item in results:
        line = f"{item['date']}\t{item['time']}\t{item['timezone']}\t{item['value']}\t{item['externalForecastVariableId']}\t{item['externalForecastConfigurationId']}\t{item['Unit integration key']}\t{item['Section integration key']}"
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
        
        print(f"Found {len(results)} unique timestamp/variable combinations")
        
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
            kallmat_count = sum(1 for r in results if r['externalForecastVariableId'] == 'kallmat')
            varmmat_count = sum(1 for r in results if r['externalForecastVariableId'] == 'varmmat')
            print(f"- Kallmat records: {kallmat_count}")
            print(f"- Varmmat records: {varmmat_count}")
            
            # Show first few records as example
            print(f"\nExample output (first 5 records):")
            for i, record in enumerate(results[:5]):
                print(f"  {record['date']}\t{record['time']}\t{record['timezone']}\t{record['value']}\t{record['externalForecastVariableId']}")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
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