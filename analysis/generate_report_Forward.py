import pandas as pd
import logging

# Opt-in to future downcasting behavior to silence FutureWarning
pd.set_option('future.no_silent_downcasting', True)

from config import (
    BASE_PATH, 
    OUTPUT_SONG_REPORT_PATH_MONTHLY
)
from pipeline import run_etl_pipeline, clean_and_unify_data
from reporting import generate_song_report

# Configure professional logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger("ReportGenerator")

START_MONTH = "2023-01"
END_MONTH = "2026-06"

# 1. Load Data
logger.info(f"Starting pipeline. Ingesting raw Excel files from: {BASE_PATH}")
df_raw = run_etl_pipeline(BASE_PATH)
if df_raw.empty:
    logger.error("Raw dataset is empty. Please verify BASE_PATH contains TME excel files.")
    exit(1)
logger.info(f"Raw data successfully loaded. Total rows: {len(df_raw)}")

# 2. Clean Data
logger.info("Cleaning dates, standardizing metadata, and unifying platform names...")
df_cleaned = clean_and_unify_data(df_raw)
# Exclude Taiwan songs (ISRC starting with 'TW') and drop empty UPC
df_cleaned = df_cleaned[~df_cleaned['ISRC'].str.startswith('TW', na=False)]
df_cleaned = df_cleaned[df_cleaned['UPC'].notna()]
logger.info(f"Data cleaning completed. Exclude 'TW' songs and empty UPC. Total cleaned rows: {len(df_cleaned)}")

# 3. Filter Data
logger.info(f"Filtering dataset for target analysis period: {START_MONTH} to {END_MONTH}...")
df_filtered = df_cleaned[
    (df_cleaned['YearMonth'] >= START_MONTH) & 
    (df_cleaned['YearMonth'] <= END_MONTH)
]

if df_filtered.empty:
    logger.warning(f"No records found within the filtered period: {START_MONTH} to {END_MONTH}")
else:
    logger.info(f"Filtered dataset contains {len(df_filtered)} rows.")
    # Debug metrics
    platforms = df_filtered['Platform'].unique().tolist() if 'Platform' in df_filtered.columns else []
    unique_songs = df_filtered['Song'].nunique() if 'Song' in df_filtered.columns else 0
    total_rev = df_filtered['Revenue'].sum() if 'Revenue' in df_filtered.columns else 0.0
    
    logger.info(f"Platforms found: {platforms}")
    logger.info(f"Unique songs count: {unique_songs}")
    logger.info(f"Total revenue for the period: ${total_rev:,.2f}")

# 4. Generate Report
logger.info(f"Generating monthly pivoted song report...")
song_final = generate_song_report(df_filtered, OUTPUT_SONG_REPORT_PATH_MONTHLY)
logger.info(f"Report generated in: {OUTPUT_SONG_REPORT_PATH_MONTHLY}")