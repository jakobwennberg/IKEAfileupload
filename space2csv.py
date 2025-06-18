#!/usr/bin/env python3
import sys
import csv

def space_to_csv(input_path: str, output_path: str) -> None:
    """
    Reads a space-separated file (any number of spaces between columns)
    and writes it out as a comma-separated CSV.
    """
    with open(input_path, 'r', encoding='utf-8') as fin, \
         open(output_path, 'w', encoding='utf-8', newline='') as fout:

        writer = csv.writer(fout)
        for line in fin:
            # split on any whitespace sequence
            row = line.strip().split()

            if row:
                writer.writerow(row)

def main():
    if len(sys.argv) < 2:
        print("Usage: python space2csv.py input.txt [output.csv]")
        sys.exit(1)

    input_file  = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'output.csv'

    try:
        space_to_csv(input_file, output_file)
        print(f"Converted '{input_file}' → '{output_file}'")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
