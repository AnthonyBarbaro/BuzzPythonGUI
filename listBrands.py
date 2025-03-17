#!/usr/bin/env python3
"""
listBrands.py

Script to iterate all CSV files in a folder (default: "files") and
list the unique Brand names found in each file.
"""

import os
import sys
import pandas as pd

FILES_DIR = "files"

def main():
    # Make sure the folder exists
    if not os.path.isdir(FILES_DIR):
        print(f"[ERROR] The folder '{FILES_DIR}' does not exist. Exiting.")
        sys.exit(1)

    # Iterate over all CSV files in FILES_DIR
    for file_name in os.listdir(FILES_DIR):
        if file_name.lower().endswith(".csv"):
            csv_path = os.path.join(FILES_DIR, file_name)
            try:
                # Read the CSV
                df = pd.read_csv(csv_path)
                # Check if there's a "Brand" column
                if "Brand" in df.columns:
                    brands = df["Brand"].dropna().unique().tolist()
                    # Print results
                    print(f"File: {file_name}")
                    print("Brands found:")
                    for b in sorted(brands):
                        print(f"  - {b}")
                    print("-" * 40)  # separator line
                else:
                    print(f"[INFO] '{file_name}' has no 'Brand' column.")
            except Exception as e:
                print(f"[ERROR] Could not read '{file_name}': {e}")

if __name__ == "__main__":
    main()
