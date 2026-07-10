"""
Data ETL Pipeline for reading, cleaning, and preprocessing TME raw files.
"""
import os
import glob
import re
import pandas as pd
import sqlite3

from typing import List
from config import BASE_PATH, STANDARD_SCHEMA, PLATFORM_MAP, EXCLUDE_FILE_NAME, ARTIST_ALIAS_MAP, SONG_ALIAS_MAP
from utils import read_excel_auto_schema, clean_upc

def get_file_list(base_path: str = BASE_PATH) -> List[str]:
    """Retrieve all Excel files recursively from the base path, excluding 'bill' and temp lock files."""
    all_files = glob.glob(os.path.join(base_path, "**/*.xlsx"), recursive=True)
    return [
        f for f in all_files 
        # if not os.path.basename(f).lower().startswith('bill') 
        # if not any(ex in os.path.basename(f) for ex in EXCLUDE_FILE_NAME)
        if not os.path.basename(f).startswith(tuple(EXCLUDE_FILE_NAME))
    ]


def calculate_clicks(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    """Calculate the total clicks depending on file prefix flow."""
    filename_lower = filename.lower()
    
    if filename_lower.startswith('aiting') or filename_lower.startswith('爱听'):
        # Aiting clicks
        click_cols = ['Aiting_Free', 'Aiting_Sub']
        available_cols = [c for c in click_cols if c in df.columns]
        if available_cols:
            df['Total_Clicks'] = df[available_cols].sum(axis=1)
            df.drop(columns=available_cols, inplace=True, errors='ignore')
        else:
            df['Total_Clicks'] = 0.0
            
    elif filename_lower.startswith('k_') or filename_lower.startswith('k歌'):
        # K clicks
        click_cols = ['K_Lyrics', 'K_Comp', 'K_Rec_Orig', 'K_Rec_Kara_Lic', 'K_Rec_Kara_TME']
        available_cols = [c for c in click_cols if c in df.columns]
        if available_cols:
            df['Total_Clicks'] = df[available_cols].sum(axis=1)
            df.drop(columns=available_cols, inplace=True, errors='ignore')
        else:
            df['Total_Clicks'] = 0.0
            
    elif filename_lower.startswith('single') or filename_lower.startswith('单曲'):
        # Single clicks
        click_cols = ['Single_IOS', 'Single_Others']
        available_cols = [c for c in click_cols if c in df.columns]
        if available_cols:
            df['Total_Clicks'] = df[available_cols].sum(axis=1)
            df.drop(columns=available_cols, inplace=True, errors='ignore')
        else:
            df['Total_Clicks'] = 0.0
    
    elif filename_lower.startswith('song') or filename_lower.startswith('歌曲'):
        # Song clicks
        click_cols = ['Song_Free_Normal', 'Song_Free_NonNormal', 'Song_Sub_Basic', 'Song_Sub_Senior', 'Song_MuCoin']
        available_cols = [c for c in click_cols if c in df.columns]
        if available_cols:
            df['Total_Clicks'] = df[available_cols].sum(axis=1)
            df.drop(columns=available_cols, inplace=True, errors='ignore')
        else:
            df['Total_Clicks'] = 0.0
            
    elif filename_lower.startswith('mv'):
        # MV clicks
        click_cols = ['MV_Comp']
        available_cols = [c for c in click_cols if c in df.columns]
        if available_cols:
            df['Total_Clicks'] = df[available_cols].sum(axis=1)
            df.drop(columns=available_cols, inplace=True, errors='ignore')
        else:
            df['Total_Clicks'] = 0.0
            
    else:
        # Default clicks to 0
        df['Total_Clicks'] = 0.0
        
    return df

def run_etl_pipeline(base_path: str = BASE_PATH) -> pd.DataFrame:
    """Execute the complete ETL pipeline to extract, transform and load TME Excel files."""
    file_list = get_file_list(base_path)
    if not file_list:
        print(f"Warning: No Excel files found in {base_path}")
        return pd.DataFrame()
        
    all_data = []
    for file in file_list:
        filename = os.path.basename(file)
        # 1. Read Excel using auto schema mapping
        temp_df = read_excel_auto_schema(file, STANDARD_SCHEMA)
        
        if temp_df is not None:
            # 2. Clean UPC
            if 'UPC' in temp_df.columns:
                temp_df['UPC'] = temp_df['UPC'].apply(clean_upc)
            
            # Resolve column mapping conflict: Song files using "广告收入分成-使用量" will map to Aiting_Free initially; rename to Song_Free_Normal
            filename_lower = filename.lower()
            if (filename_lower.startswith('song') or filename_lower.startswith('歌曲')) and 'Aiting_Free' in temp_df.columns:
                temp_df = temp_df.rename(columns={'Aiting_Free': 'Song_Free_Normal'})

            # 3. Calculate Clicks
            temp_df = calculate_clicks(temp_df, filename)
            
            # Record source file name
            temp_df['source_file'] = filename
            all_data.append(temp_df)
        else:
            print(f"Skipping file {file} due to read error.")
            
    # Filter out empty DataFrames to prevent FutureWarning during concatenation
    all_data = [df for df in all_data if not df.empty]
    if not all_data:
        return pd.DataFrame()
        
    df_raw = pd.concat(all_data, ignore_index=True)
    return df_raw

def standardize_artist_string(artist_val) -> str:
    # Convert numeric values like 5566 or 5566.0 to '5566' string first
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

    # Check the whole string first (for names containing multiple artist seperated by commas like '吉克隽逸,장혁,朴宰范(JAYPARK)')
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

    # 1. Split by English comma ',' or Chinese comma '，'
    parts = re.split(r',|，', artist_str)
    
    standardized_parts = []
    for part in parts:
        part_clean = part.strip()
        if not part_clean:
            continue
        # Convert key to uppercase and remove spaces for dictionary lookup
        lookup_key = part_clean.replace(' ', '').upper()
        # Find if lookup_key matches standard name or any of its aliases
        mapped_name = part_clean
        for std_name, alias_list in ARTIST_ALIAS_MAP.items():
            std_clean = std_name.replace(' ', '').upper()
            if lookup_key == std_clean:
                mapped_name = std_name
                break
            # Handle if alias_list is a string or list/iterable
            if isinstance(alias_list, str):
                alias_list = [alias_list]
            cleaned_aliases = [a.replace(' ', '').upper() for a in alias_list]
            if lookup_key in cleaned_aliases:
                mapped_name = std_name
                break
        standardized_parts.append(mapped_name)
        
    # 2. Deduplicate and sort alphabetically (to ensure order independent matching)
    unique_parts = sorted(list(set(standardized_parts)))
    
    # 3. Join back
    if not unique_parts:
        return 'UNKNOWNARTIST'
    return ', '.join(unique_parts)


def standardize_song_string(song_val) -> str:
    # Handle numeric/null values
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

    # Case-insensitive & space-insensitive matching against SONG_ALIAS_MAP
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

def clean_and_unify_data(df: pd.DataFrame) -> pd.DataFrame:
    """Apply cleaning rules: clean dates, map platform, standardize song/artist metadata."""
    if df.empty:
        return df
        
    df_cleaned = df.copy()
    
    # 1. Clean dates to YearMonth format (e.g. 20240901 -> '2024-09')
    if 'Date' in df_cleaned.columns:
        df_cleaned['Date'] = df_cleaned['Date'].astype(str).str.slice(0, 8)
        df_cleaned['Date'] = pd.to_datetime(df_cleaned['Date'], format='%Y%m%d', errors='coerce')
        df_cleaned['Date'] = df_cleaned['Date'].dt.to_period('M').astype(str)
        df_cleaned.rename(columns={'Date': 'YearMonth'}, inplace=True)
        
    # 2. Unify Platform Names
    if 'Platform' in df_cleaned.columns:
        df_cleaned['Platform'] = df_cleaned['Platform'].replace(PLATFORM_MAP)
        
    # Ensure Song names are string types and clean trailing '.0' (caused by Excel numeric cell parsing)
    if 'Song' in df_cleaned.columns:
        df_cleaned['Song'] = df_cleaned['Song'].apply(
            lambda x: str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) and float(x).is_integer()
                      else (str(x).strip() if pd.notna(x) else '')
        )
        df_cleaned['standard_song'] = df_cleaned['Song'].apply(standardize_song_string)
        
        
    # 3. Standardize Song and Artist (Resolve multiple names per ISRC / UPC)
    # Standardize Song Names (first occurrence per ISRC)
    
    # if 'ISRC' in df_cleaned.columns and 'Song' in df_cleaned.columns:
    #     standard_songs = df_cleaned.groupby('ISRC')['Song'].agg('first')
    #     df_cleaned['standard_song'] = df_cleaned['ISRC'].map(standard_songs).fillna(df_cleaned['Song'])
    # else:
    #     df_cleaned['standard_song'] = df_cleaned['Song']

    # Standardize Artist Names (first occurrence per UPC)
    # if 'UPC' in df_cleaned.columns and 'Artist' in df_cleaned.columns:
    #     standard_artists = df_cleaned.groupby('UPC')['Artist'].agg('first')
    #     df_cleaned['standard_artist'] = df_cleaned['UPC'].map(standard_artists)
    # df_cleaned['standard_artist'] = df_cleaned['Artist']

    # Clean and standardize artist using map dictionary, sorting, and deduplication
    if 'Artist' in df_cleaned.columns:
        df_cleaned['standard_artist'] = df_cleaned['Artist'].apply(standardize_artist_string)
    else:
        df_cleaned['standard_artist'] = 'UNKNOWNARTIST'
    # keep 'Artist' column original without any modification


    # Fill missing ISRC with "Song-Artist" (no spaces, uppercase)
    if 'ISRC' in df_cleaned.columns and 'Song' in df_cleaned.columns and 'Artist' in df_cleaned.columns:
        song_cleaned = df_cleaned['Song'].fillna('UNKNOWNSONG').astype(str).str.replace(r'\s+', '', regex=True).str.upper()
        artist_cleaned = df_cleaned['standard_artist'].fillna('UNKNOWNARTIST').astype(str).str.replace(r'\s+', '', regex=True).str.upper()
        isrc_fill = song_cleaned + " - " + artist_cleaned
        # Strip leading/trailing whitespaces from existing ISRCs
        df_cleaned['ISRC'] = df_cleaned['ISRC'].apply(
            lambda x: str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) and float(x).is_integer()
                      else (str(x).strip() if pd.notna(x) else x)
        )
        df_cleaned['ISRC'] = df_cleaned['ISRC'].fillna(isrc_fill)
        df_cleaned.loc[df_cleaned['ISRC'].astype(str).str.strip() == '', 'ISRC'] = isrc_fill


    # 4. Exclude Taiwan songs (ISRC starting with 'TW') and drop empty ISRC
    # if 'ISRC' in df_cleaned.columns:
    #     df_cleaned = df_cleaned[~df_cleaned['ISRC'].str.startswith('TW', na=False)]
    #     df_cleaned = df_cleaned.dropna(subset=['ISRC'])

    # Capitalize all the letters of ISRC
    if 'ISRC' in df_cleaned.columns:
        df_cleaned['ISRC'] = df_cleaned['ISRC'].apply(lambda x: str(x).upper() if pd.notna(x) else x)
        
    if 'YearMonth' in df_cleaned.columns:
        df_cleaned.sort_values('YearMonth', inplace=True)
        
    return df_cleaned

def sync_to_sqlite(base_path: str = BASE_PATH, db_path: str = None) -> pd.DataFrame:
    """
    Synchronize raw Excel records to an SQLite database cache.
    Reads only new or modified files, then returns the full raw DataFrame from SQLite.
    """
    if db_path is None:
        from config import DB_PATH
        db_path = DB_PATH

    os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Create processed_files table if not exists
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS processed_files (
            filepath TEXT PRIMARY KEY,
            last_modified REAL,
            row_count INTEGER
        )
    """)
    conn.commit()

    # Get processed files and metadata
    cursor.execute("SELECT filepath, last_modified FROM processed_files")
    processed_map = {row[0]: row[1] for row in cursor.fetchall()}

    # Get list of all physical raw files
    file_list = get_file_list(base_path)
    
    # Clean up records for deleted physical files
    physical_abs_paths = {os.path.abspath(f) for f in file_list}
    deleted_files = [path for path in processed_map if path not in physical_abs_paths]
    if deleted_files:
        print(f"🗑️ SQLite Cleanup: Found {len(deleted_files)} deleted Excel files. Removing from database...")
        for path in deleted_files:
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='raw_records'")
            if cursor.fetchone():
                cursor.execute("DELETE FROM raw_records WHERE db_file_path = ?", (path,))
            cursor.execute("DELETE FROM processed_files WHERE filepath = ?", (path,))
            conn.commit()
            print(f"   [Deleted Cache] {os.path.basename(path)}")

    new_or_modified_files = []
    for f in file_list:
        abs_path = os.path.abspath(f)
        mtime = os.path.getmtime(abs_path)
        if abs_path not in processed_map or processed_map[abs_path] != mtime:
            new_or_modified_files.append((abs_path, mtime))

    if new_or_modified_files:
        print(f"🔄 SQLite Ingestion: Found {len(new_or_modified_files)} new/modified Excel files to process.")
        for abs_path, mtime in new_or_modified_files:
            filename = os.path.basename(abs_path)
            # Read Excel
            temp_df = read_excel_auto_schema(abs_path, STANDARD_SCHEMA)
            if temp_df is not None:
                # Clean UPC
                if 'UPC' in temp_df.columns:
                    temp_df['UPC'] = temp_df['UPC'].apply(clean_upc)
                
                # Resolve column mapping conflict
                filename_lower = filename.lower()
                if (filename_lower.startswith('song') or filename_lower.startswith('歌曲')) and 'Aiting_Free' in temp_df.columns:
                    temp_df = temp_df.rename(columns={'Aiting_Free': 'Song_Free_Normal'})

                # Calculate clicks
                temp_df = calculate_clicks(temp_df, filename)
                temp_df['source_file'] = filename
                temp_df['db_file_path'] = abs_path
                
                # Delete existing records from raw_records if file already existed
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='raw_records'")
                if cursor.fetchone():
                    cursor.execute("DELETE FROM raw_records WHERE db_file_path = ?", (abs_path,))
                    conn.commit()

                # Save raw records to SQLite
                temp_df.to_sql('raw_records', conn, if_exists='append', index=False)
                
                # Update processed_files
                cursor.execute(
                    "INSERT OR REPLACE INTO processed_files (filepath, last_modified, row_count) VALUES (?, ?, ?)",
                    (abs_path, mtime, len(temp_df))
                )
                conn.commit()
                print(f"   [Imported] {filename} ({len(temp_df)} rows)")
            else:
                print(f"Skipping file {abs_path} due to read error.")

    # Check if raw_records table has been created/populated
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='raw_records'")
    if not cursor.fetchone():
        print("❌ SQLite Cache is empty and no raw records are cached yet.")
        conn.close()
        return pd.DataFrame()

    # Load all raw records from SQLite
    df_raw = pd.read_sql_query("SELECT * FROM raw_records", conn)
    conn.close()
    return df_raw

