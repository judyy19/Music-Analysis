"""
Python script to merge duplicate ISRC rows in song_final Excel reports.
This script groups the report by ISRC, sums all numeric metrics (clicks, revenue),
and concatenates differing artists.
"""
import os
import pandas as pd
import sys

def merge_isrc_excel(input_path: str, output_path: str):
    if not os.path.exists(input_path):
        print(f"Error: Input file '{input_path}' does not exist.")
        return

    print(f"Reading input file: {input_path}")
    df = pd.read_excel(input_path)

    # 1. Standardize column types
    if 'ISRC' in df.columns:
        df['ISRC'] = df['ISRC'].astype(str).str.strip().str.upper()
    if 'song' in df.columns:
        df['song'] = df['song'].astype(str).str.strip()
    if 'artist' in df.columns:
        df['artist'] = df['artist'].astype(str).str.strip()

    # 2. Identify metadata columns vs numeric columns
    metadata_cols = ['ISRC', 'song', 'artist', 'UPC', 'album']
    present_metadata = [col for col in metadata_cols if col in df.columns]
    numeric_cols = [col for col in df.columns if col not in present_metadata]

    # 3. Define custom aggregation rules
    agg_dict = {}
    if 'song' in df.columns:
        agg_dict['song'] = 'first'
    if 'artist' in df.columns:
        # Join unique artists alphabetically
        agg_dict['artist'] = lambda x: ', '.join(sorted(list(set(x.dropna().astype(str).str.strip()))))
        # agg_dict['artist'] = 'first'
    if 'album' in df.columns:
        agg_dict['album'] = 'first'
    if 'UPC' in df.columns:
        agg_dict['UPC'] = 'first'

    # Sum all numeric columns (clicks, revenue, totals)
    for col in numeric_cols:
        agg_dict[col] = 'sum'

    print("Grouping by ISRC and merging records...")
    # Group by ISRC and apply aggregation
    merged_df = df.groupby('ISRC', as_index=False).agg(agg_dict)

    # 4. Reorder columns to match original order
    original_col_order = [col for col in df.columns if col in merged_df.columns]
    merged_df = merged_df[original_col_order]

    # 5. Sort by total revenue descending (or total click if revenue not present)
    if 'total_revenue' in merged_df.columns:
        merged_df = merged_df.sort_values(by='total_revenue', ascending=False)
    elif 'total_click' in merged_df.columns:
        merged_df = merged_df.sort_values(by='total_click', ascending=False)

    merged_df = merged_df.reset_index(drop=True)

    # 6. Save to Excel
    print(f"Saving merged report to: {output_path}")
    merged_df.to_excel(output_path, index=False)
    print("Merge completed successfully!")

if __name__ == "__main__":
    # If paths are passed via command line, use them. Otherwise, default to Sony reports.
    if len(sys.argv) > 2:
        in_path = sys.argv[1]
        out_path = sys.argv[2]
    else:
        # Default paths relative to script location
        script_dir = os.path.dirname(os.path.abspath(__file__))
        in_path = os.path.abspath(os.path.join(script_dir, "../output/TME_Song_Report_for_Sony_大曲庫_new.xlsx"))
        out_path = os.path.abspath(os.path.join(script_dir, "../output/TME_Song_Report_for_Sony_大曲庫_merged.xlsx"))
        
        print("Usage: python merge_isrc_reports.py <input_excel_path> <output_excel_path>")
        print(f"Using default paths:\nInput:  {in_path}\nOutput: {out_path}\n")

    merge_isrc_excel(in_path, out_path)
