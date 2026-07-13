"""
ETL Logic for scanning raw music revenue reports, cleaning fields,
calculating clicks, performing validation checks, and loading them
into the sqlite database.
"""
import os
import glob
import re
import json
import sqlite3
import pandas as pd
import numpy as np
from typing import List

from config import (
    STANDARD_SCHEMA, PLATFORM_MAP, EXCLUDE_FILE_NAME, 
    ARTIST_ALIAS_MAP, SONG_ALIAS_MAP
)
from catalog.catalog_logic import add_song_metadata

def init_etl_tables(conn: sqlite3.Connection):
    """
    Initialize SQLite tables for storing processed files metadata,
    cleaned ingestion records, and parsing/validation errors.
    
    Args:
        conn (sqlite3.Connection): SQLite connection object.
    """
    cursor = conn.cursor()
    
    # processed_files table for tracking already-ingested files and their metadata
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS processed_files (
            filepath TEXT PRIMARY KEY,
            last_modified REAL,
            row_count INTEGER
        )
    """)
    
    # etl_records stores validated data rows ready for royalty allocation
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS etl_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            year_month TEXT,
            isrc TEXT,
            song TEXT,
            artist TEXT,
            album TEXT,
            upc TEXT,
            platform TEXT,
            revenue REAL,
            clicks REAL,
            source_file TEXT
        )
    """)
    
    # etl_errors stores validation errors or structural anomalies for the Error Reporter
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS etl_errors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_file TEXT,
            error_reason TEXT,
            original_data TEXT
        )
    """)
    
    # Create indexes for etl_records to significantly speed up filtering and aggregation
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_etl_records_ym_song ON etl_records(year_month, song)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_etl_records_isrc ON etl_records(isrc)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_etl_records_song ON etl_records(song)")
    
    conn.commit()

def get_file_list(base_path: str) -> List[str]:
    """
    Search recursively for all Excel spreadsheets inside the base path,
    filtering out files matching any patterns in EXCLUDE_FILE_NAME.
    
    Args:
        base_path (str): Root folder path to search.
        
    Returns:
        List[str]: List of absolute paths of matching raw Excel files.
    """
    all_files = glob.glob(os.path.join(base_path, "**/*.xlsx"), recursive=True)
    valid_files = []
    for f in all_files:
        basename = os.path.basename(f)
        if basename.startswith("~$") or basename.startswith("."):
            continue
        # Check against exclusions
        if not basename.startswith(tuple(EXCLUDE_FILE_NAME)):
            valid_files.append(os.path.abspath(f))
    return valid_files

def normalize_col(col: str) -> str:
    """Normalize raw header name for better alias mapping."""
    col = str(col).lower()
    col = re.sub(r'[_\-]', ' ', col)
    col = re.sub(r'\s+', ' ', col)
    return col.strip()

def match_column(col: str, schema: dict) -> str:
    """Matches a single header column name against aliases in STANDARD_SCHEMA."""
    col_norm = str(col).lower().strip()
    for standard, aliases in schema.items():
        choices = [str(a).lower().strip() for a in aliases]
        if col_norm in choices:
            return standard
    return None

def auto_map_columns(columns: list, schema: dict) -> dict:
    """Create a dictionary mapping raw columns to standardized schema column names."""
    rename_map = {}
    for col in columns:
        matched = match_column(col, schema)
        if matched:
            rename_map[col] = matched
    return rename_map

def clean_upc(val) -> str:
    """Standardize UPC values to uniform 'UPC-xxxxxxxx' strings."""
    if pd.isna(val):
        return np.nan
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    if val_str.startswith('UPC-'):
        val_str = val_str.replace('UPC-', '')
    # If it is empty after replacing, return NaN
    if not val_str or val_str.lower() == 'nan':
        return np.nan
    return f"UPC-{val_str}"

def standardize_artist_string(artist_val) -> str:
    """
    Cleans and standardizes artist names by splitting combinations, 
    matching against ARTIST_ALIAS_MAP, and ordering them alphabetically.
    """
    if isinstance(artist_val, (int, float)) and not pd.isna(artist_val):
        if float(artist_val).is_integer():
            artist_str = str(int(artist_val))
        else:
            artist_str = str(artist_val)
    elif pd.isna(artist_val):
        return 'UNKNOWNARTIST'
    else:
        artist_str = str(artist_val).strip()

    if not artist_str or artist_str.lower() == 'nan':
        return 'UNKNOWNARTIST'

    # Check for direct whole name match
    whole_key = artist_str.replace(' ', '').upper()
    for std_name, alias_list in ARTIST_ALIAS_MAP.items():
        std_clean = std_name.replace(' ', '').upper()
        if whole_key == std_clean:
            return std_name
        if isinstance(alias_list, str):
            alias_list = [alias_list]
        cleaned_aliases = [a.replace(' ', '').upper() for a in alias_list]
        if whole_key in cleaned_aliases:
            return std_name

    # Split by common comma delimiters
    parts = re.split(r',|，', artist_str)
    standardized_parts = []
    for part in parts:
        part_clean = part.strip()
        if not part_clean:
            continue
        
        lookup_key = part_clean.replace(' ', '').upper()
        mapped_name = part_clean
        
        for std_name, alias_list in ARTIST_ALIAS_MAP.items():
            std_clean = std_name.replace(' ', '').upper()
            if lookup_key == std_clean:
                mapped_name = std_name
                break
            if isinstance(alias_list, str):
                alias_list = [alias_list]
            cleaned_aliases = [a.replace(' ', '').upper() for a in alias_list]
            if lookup_key in cleaned_aliases:
                mapped_name = std_name
                break
        standardized_parts.append(mapped_name)
        
    unique_parts = sorted(list(set(standardized_parts)))
    if not unique_parts:
        return 'UNKNOWNARTIST'
    return ', '.join(unique_parts)

def standardize_song_string(song_val) -> str:
    """Cleans and maps song titles against SONG_ALIAS_MAP."""
    if isinstance(song_val, (int, float)) and not pd.isna(song_val):
        if float(song_val).is_integer():
            song_str = str(int(song_val))
        else:
            song_str = str(song_val)
    elif pd.isna(song_val):
        return 'UNKNOWNSONG'
    else:
        song_str = str(song_val).strip()

    if not song_str or song_str.lower() == 'nan':
        return 'UNKNOWNSONG'

    lookup_key = song_str.replace(' ', '').upper()
    for std_name, alias_list in SONG_ALIAS_MAP.items():
        std_clean = std_name.replace(' ', '').upper()
        if lookup_key == std_clean:
            return std_name
        if isinstance(alias_list, str):
            alias_list = [alias_list]
        cleaned_aliases = [a.replace(' ', '').upper() for a in alias_list]
        if lookup_key in cleaned_aliases:
            return std_name
            
    return song_str

def calculate_clicks(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    """
    Aggregates metrics columns into a unified 'Total_Clicks' column
    based on the file name prefix.
    """
    filename_lower = filename.lower()
    
    if filename_lower.startswith('aiting') or filename_lower.startswith('爱听'):
        click_cols = ['Aiting_Free', 'Aiting_Sub']
    elif filename_lower.startswith('k_') or filename_lower.startswith('k歌'):
        click_cols = ['K_Lyrics', 'K_Comp', 'K_Rec_Orig', 'K_Rec_Kara_Lic', 'K_Rec_Kara_TME']
    elif filename_lower.startswith('single') or filename_lower.startswith('单曲'):
        click_cols = ['Single_IOS', 'Single_Others']
    elif filename_lower.startswith('song') or filename_lower.startswith('歌曲'):
        click_cols = ['Song_Free_Normal', 'Song_Free_NonNormal', 'Song_Sub_Basic', 'Song_Sub_Senior', 'Song_MuCoin']
    elif filename_lower.startswith('mv'):
        click_cols = ['MV_Comp']
    else:
        click_cols = []
        
    available_cols = [c for c in click_cols if c in df.columns]
    if available_cols:
        df['Total_Clicks'] = df[available_cols].sum(axis=1)
        df.drop(columns=available_cols, inplace=True, errors='ignore')
    else:
        df['Total_Clicks'] = 0.0
        
    return df

def clean_and_unify_dataframe(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    """
    Applies standardization rules to loaded dataframes including
    date formatting, platform translation, name cleaning, and ISRC default filling.
    """
    if df.empty:
        return df
        
    df_cleaned = df.copy()
    
    # 1. Date clean to YearMonth YYYY-MM
    if 'Date' in df_cleaned.columns:
        df_cleaned['Date'] = df_cleaned['Date'].astype(str).str.slice(0, 8)
        # Handle string inputs like '20240901' or '2024-09'
        df_cleaned['Date'] = pd.to_datetime(df_cleaned['Date'], format='%Y%m%d', errors='coerce')
        df_cleaned['YearMonth'] = df_cleaned['Date'].dt.to_period('M').astype(str)
        # Drop temporary 'Date' if renamed
        df_cleaned.drop(columns=['Date'], inplace=True, errors='ignore')
    else:
        df_cleaned['YearMonth'] = 'UNKNOWN'
        
    # 2. Platform unification
    if 'Platform' in df_cleaned.columns:
        df_cleaned['Platform'] = df_cleaned['Platform'].replace(PLATFORM_MAP)
    else:
        df_cleaned['Platform'] = 'Unknown'
        
    # 3. Song title formatting
    if 'Song' in df_cleaned.columns:
        df_cleaned['Song'] = df_cleaned['Song'].apply(
            lambda x: str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) and float(x).is_integer()
                      else (str(x).strip() if pd.notna(x) else '')
        )
        df_cleaned['standard_song'] = df_cleaned['Song'].apply(standardize_song_string)
    else:
        df_cleaned['standard_song'] = 'UNKNOWNSONG'
        
    # 4. Artist name formatting
    if 'Artist' in df_cleaned.columns:
        df_cleaned['standard_artist'] = df_cleaned['Artist'].apply(standardize_artist_string)
    else:
        df_cleaned['standard_artist'] = 'UNKNOWNARTIST'
        
    # 5. Missing ISRC Fallback mapping: 'Song - Artist'
    if 'ISRC' in df_cleaned.columns:
        # Standardize ISRC input formatting
        df_cleaned['ISRC'] = df_cleaned['ISRC'].apply(
            lambda x: str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) and float(x).is_integer()
                      else (str(x).strip() if pd.notna(x) else x)
        )
        
        song_cleaned = df_cleaned['standard_song'].astype(str).str.replace(r'\s+', '', regex=True).str.upper()
        artist_cleaned = df_cleaned['standard_artist'].astype(str).str.replace(r'\s+', '', regex=True).str.upper()
        isrc_fill = song_cleaned + " - " + artist_cleaned
        
        df_cleaned['ISRC'] = df_cleaned['ISRC'].fillna(isrc_fill)
        df_cleaned.loc[df_cleaned['ISRC'].astype(str).str.strip() == '', 'ISRC'] = isrc_fill
        df_cleaned['ISRC'] = df_cleaned['ISRC'].astype(str).str.upper().str.strip()
    else:
        # ISRC column completely missing
        song_cleaned = df_cleaned['standard_song'].astype(str).str.replace(r'\s+', '', regex=True).str.upper()
        artist_cleaned = df_cleaned['standard_artist'].astype(str).str.replace(r'\s+', '', regex=True).str.upper()
        isrc_fill = song_cleaned + " - " + artist_cleaned
        df_cleaned['ISRC'] = isrc_fill

    # 6. clean UPC format
    if 'UPC' in df_cleaned.columns:
        df_cleaned['UPC'] = df_cleaned['UPC'].apply(clean_upc)
    else:
        df_cleaned['UPC'] = np.nan
        
    # 7. Album empty fill
    if 'Album' not in df_cleaned.columns:
        df_cleaned['Album'] = 'Unknown Album'
    else:
        df_cleaned['Album'] = df_cleaned['Album'].fillna('Unknown Album')
        
    # 8. Revenue ensure float
    if 'Revenue' in df_cleaned.columns:
        df_cleaned['Revenue'] = pd.to_numeric(df_cleaned['Revenue'], errors='coerce').fillna(0.0)
    else:
        df_cleaned['Revenue'] = 0.0

    return df_cleaned

def run_etl_pipeline(conn: sqlite3.Connection, input_dir: str) -> dict:
    """
    Scans files, matches schema, parses contents, validates rules, 
    inserts valid rows to etl_records and errors to etl_errors.
    
    Args:
        conn (sqlite3.Connection): DB connection.
        input_dir (str): Base folder containing physical files.
        
    Returns:
        dict: Ingestion summary metrics (files processed, rows ingested, errors count).
    """
    cursor = conn.cursor()
    
    # Retrieve modified time mapping of previously processed files
    cursor.execute("SELECT filepath, last_modified FROM processed_files")
    processed_map = {row[0]: row[1] for row in cursor.fetchall()}
    
    # Get all physical raw files
    file_list = get_file_list(input_dir)
    
    new_or_modified_files = []
    for f in file_list:
        abs_path = os.path.abspath(f)
        mtime = os.path.getmtime(abs_path)
        if abs_path not in processed_map or processed_map[abs_path] != mtime:
            new_or_modified_files.append((abs_path, mtime))
            
    summary = {
        'files_processed': 0,
        'records_added': 0,
        'records_skipped_duplicate': len(file_list) - len(new_or_modified_files),
        'errors_logged': 0
    }
    
    if not new_or_modified_files:
        return summary
        
    for abs_path, mtime in new_or_modified_files:
        filename = os.path.basename(abs_path)
        
        try:
            # Read excel using auto schema logic
            cols = pd.read_excel(abs_path, nrows=0).columns
            rename_map = auto_map_columns(cols, STANDARD_SCHEMA)
            
            if not rename_map:
                # No columns matches the minimum mapping requirement
                cursor.execute("""
                    INSERT INTO etl_errors (source_file, error_reason, original_data)
                    VALUES (?, ?, ?)
                """, (filename, "Header mapping failed. No recognized schema columns.", json.dumps(list(cols))))
                summary['errors_logged'] += 1
                continue
                
            # Load only mapped columns
            df_raw = pd.read_excel(abs_path, usecols=list(rename_map.keys()))
            df_raw = df_raw.rename(columns=rename_map)
            
            # Resolve specific conflicts
            filename_lower = filename.lower()
            if (filename_lower.startswith('song') or filename_lower.startswith('歌曲')) and 'Aiting_Free' in df_raw.columns:
                df_raw = df_raw.rename(columns={'Aiting_Free': 'Song_Free_Normal'})
                
            # Calculate total clicks column
            df_raw = calculate_clicks(df_raw, filename)
            
            # Run standardization
            df_cleaned = clean_and_unify_dataframe(df_raw, filename)
            
            valid_rows = []
            error_rows = []
            
            # Delete any previous records of this file before writing new ones
            cursor.execute("DELETE FROM etl_records WHERE source_file = ?", (filename,))
            cursor.execute("DELETE FROM etl_errors WHERE source_file = ?", (filename,))
            
            for index, row in df_cleaned.iterrows():
                # Get clean values
                ym = row['YearMonth']
                isrc = row['ISRC']
                song = row['standard_song']
                artist = row['standard_artist']
                album = row['Album']
                upc = row['UPC']
                platform = row['Platform']
                rev = float(row['Revenue'])
                clicks = float(row['Total_Clicks'])
                
                # Check validation rules:
                # 1. ISRC must not be empty
                if pd.isna(isrc) or not isrc:
                    error_rows.append((
                        filename,
                        "Missing ISRC.",
                        json.dumps(row.to_dict())
                    ))
                    continue
                
                # 2. UPC must not be empty
                if pd.isna(upc) or not upc:
                    error_rows.append((
                        filename, 
                        "Missing UPC.", 
                        json.dumps(row.to_dict())
                    ))
                    continue
                
                # 3. YearMonth check
                if ym == 'UNKNOWN' or not re.match(r'^\d{4}-\d{2}$', ym):
                    error_rows.append((
                        filename,
                        f"Invalid Date format: '{ym}'.",
                        json.dumps(row.to_dict())
                    ))
                    continue
                    
                # 4. Taiwan song exclusion
                if isrc.startswith("TW"):
                    error_rows.append((
                        filename,
                        "Excluded Taiwan ISRC (starting with TW).",
                        json.dumps(row.to_dict())
                    ))
                    continue
                
                # Add to valid list
                valid_rows.append((
                    ym, isrc, song, artist, album, upc, platform, rev, clicks, filename
                ))
                
                # Upsert standard song metadata into catalog
                add_song_metadata(conn, isrc, song, artist, album, upc)
                
            # Perform bulk inserts for performance
            if valid_rows:
                cursor.executemany("""
                    INSERT INTO etl_records (
                        year_month, isrc, song, artist, album, upc, platform, revenue, clicks, source_file
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, valid_rows)
                summary['records_added'] += len(valid_rows)
                
            if error_rows:
                cursor.executemany("""
                    INSERT INTO etl_errors (source_file, error_reason, original_data)
                    VALUES (?, ?, ?)
                """, error_rows)
                summary['errors_logged'] += len(error_rows)
                
            # Log successful file process
            cursor.execute("""
                INSERT OR REPLACE INTO processed_files (filepath, last_modified, row_count)
                VALUES (?, ?, ?)
            """, (abs_path, mtime, len(valid_rows)))
            
            conn.commit()
            summary['files_processed'] += 1
            
        except Exception as e:
            conn.rollback()
            cursor.execute("""
                INSERT INTO etl_errors (source_file, error_reason, original_data)
                VALUES (?, ?, ?)
            """, (filename, f"Exception during file parsing: {str(e)}", ""))
            conn.commit()
            summary['errors_logged'] += 1
            
    return summary
