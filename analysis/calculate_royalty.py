import pandas as pd
import os
import re
import logging
from reporting import style_excel_file

# Set up basic logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger("RoyaltyCalculator")

# Path definitions
ROYALTY_RATE_PATH = "/Users/chu-chun/Mirror/Eva/input/Royalty_Fee_Rate.xlsx"
SONG_REPORT_PATH = "/Users/chu-chun/Mirror/Eva/output/FORWARD/TME_Song_Report_Monthly_Forward.xlsx"
OUTPUT_DIR = "/Users/chu-chun/Mirror/Eva/output/Royalty/"
RIGHT_HOLDER = "Bob"
START_MONTH = "2026-01"
END_MONTH = "2026-06"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f"{RIGHT_HOLDER}_Royalty_Report.xlsx")

def main():
    logger.info("Starting Royalty Calculation Process...")
    
    # 1. Load Royalty Rates
    logger.info(f"Loading royalty rates from: {ROYALTY_RATE_PATH}")
    if not os.path.exists(ROYALTY_RATE_PATH):
        logger.error(f"Royalty rate file not found at: {ROYALTY_RATE_PATH}")
        return
    df_rates = pd.read_excel(ROYALTY_RATE_PATH)
    
    # 2. Load Monthly Song Report
    logger.info(f"Loading monthly song report from: {SONG_REPORT_PATH}")
    if not os.path.exists(SONG_REPORT_PATH):
        logger.error(f"Monthly song report not found at: {SONG_REPORT_PATH}")
        return
    df_report = pd.read_excel(SONG_REPORT_PATH)

    # Normalize ISRC strings for reliable matching
    df_rates['ISRC'] = df_rates['ISRC'].astype(str).str.strip().str.upper()
    df_report['ISRC'] = df_report['ISRC'].astype(str).str.strip().str.upper()

    # 3. Filter for Target Right Holder & Group by ISRC
    logger.info(f"Filtering rates for right holder: '{RIGHT_HOLDER}'")
    df_holder_rates = df_rates[df_rates['right_holder'].str.strip() == RIGHT_HOLDER]
    if df_holder_rates.empty:
        logger.warning(f"No rates found for right holder '{RIGHT_HOLDER}' in royalty rate sheet.")
        return
    # Group by ISRC to get unique percentages per ISRC
    df_holder_rates_unique = df_holder_rates.groupby('ISRC')['percentage'].first().reset_index()

    logger.info("Aggregating monthly song report by ISRC...")
    # Sum up revenue and clicks by ISRC (handling songs with multiple rows)
    numeric_cols = [col for col in df_report.columns if col.endswith('_revenue') or col.endswith('_clicks') or col in ['total_revenue', 'total_click']]
    agg_dict = {col: 'sum' for col in numeric_cols}
    # For text metadata, take the first occurrence
    agg_dict['song'] = 'first'
    agg_dict['artist'] = 'first'
    
    df_report_grouped = df_report.groupby('ISRC').agg(agg_dict).reset_index()

    logger.info("Merging rates with aggregated song report...")
    df_merged = pd.merge(
        df_holder_rates_unique,
        df_report_grouped,
        on='ISRC',
        how='inner'
    )
    
    if df_merged.empty:
        logger.warning("No matching songs found between the royalty rate sheet and the monthly song report.")
        return
        
    logger.info(f"Matched {len(df_merged)} unique songs for '{RIGHT_HOLDER}'.")

    # 4. Calculate Royalty Fees (revenue * percentage)
    logger.info("Calculating monthly royalty fees...")
    revenue_cols = [col for col in df_merged.columns if col.endswith('_revenue') or col == 'total_revenue']
    for col in revenue_cols:
        royalty_col = col.replace('_revenue', '_royalty_fee') if col.endswith('_revenue') else 'total_royalty_fee'
        df_merged[royalty_col] = df_merged[col] * df_merged['percentage']

    # Extract all month strings (YYYY-MM) from columns to reconstruct time series columns logically
    months = sorted(list(set(re.match(r'^\d{4}-\d{2}', col).group(0) for col in df_report.columns if re.match(r'^\d{4}-\d{2}', col))))
    
    # Filter months by target analysis period
    months = [m for m in months if START_MONTH <= m <= END_MONTH]
    logger.info(f"Filtered details for period: {START_MONTH} to {END_MONTH} ({len(months)} months)")

    # Re-calculate total_royalty_fee based on the filtered months to ensure accuracy
    logger.info("Re-calculating total royalty fee based on filtered months...")
    df_merged['total_royalty_fee'] = df_merged[[f"{m}_royalty_fee" for m in months]].sum(axis=1)

    # Define final column order: Metadata, total royalty fee, and monthly royalty details (no clicks, no original revenue)
    base_cols = ['ISRC', 'song', 'artist', 'total_royalty_fee']
    time_series_cols = [f"{m}_royalty_fee" for m in months]
            
    final_cols = base_cols + time_series_cols
    # Ensure all final columns exist in the dataframe before subsetting
    final_cols = [col for col in final_cols if col in df_merged.columns]
    df_final = df_merged[final_cols].copy()
    
    # Rename monthly columns from YYYY-MM_royalty_fee to YYYY-MM
    rename_dict = {f"{m}_royalty_fee": m for m in months}
    df_final = df_final.rename(columns=rename_dict)
    
    # Sort by total royalty fee descending
    df_final = df_final.sort_values('total_royalty_fee', ascending=False).reset_index(drop=True)

    # 5. Output and style the report
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    logger.info(f"Exporting royalty report to: {OUTPUT_FILE}")
    df_final.to_excel(OUTPUT_FILE, index=False)
    
    # Call the professional styling function
    style_excel_file(OUTPUT_FILE)
    logger.info(f"Royalty report successfully generated and styled at: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
