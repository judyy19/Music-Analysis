"""
Data ETL Pipeline for reading, cleaning, and preprocessing TME raw files.
"""
import os
import glob
import pandas as pd
from typing import List
from config import BASE_PATH, STANDARD_SCHEMA, PLATFORM_MAP, EXCLUDE_FILE_NAME
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
            
    if not all_data:
        return pd.DataFrame()
        
    df_raw = pd.concat(all_data, ignore_index=True)
    return df_raw

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
        
        
    # 3. Standardize Song and Artist (Resolve multiple names per ISRC / UPC)
    # Standardize Song Names (first occurrence per ISRC)
    df_cleaned['standard_song'] = df_cleaned['Song'].astype(str)
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

    # TODO: implement artist mapping table
    # Clean and standardize artist: remove spaces, convert to uppercase, handle NaN/blank/'NA'
    temp_artist = df_cleaned['Artist'].fillna('UNKNOWNARTIST').astype(str).str.replace(r'\s+', '', regex=True).str.upper()
    df_cleaned['standard_artist'] = temp_artist.replace(['nan'], 'UNKNOWNARTIST')
    # keep 'Artist' column original without any modification


    # Fill missing ISRC with "Song-Artist" (no spaces, uppercase)
    if 'ISRC' in df_cleaned.columns and 'Song' in df_cleaned.columns and 'Artist' in df_cleaned.columns:
        song_cleaned = df_cleaned['Song'].fillna('UNKNOWNSONG').astype(str).str.replace(r'\s+', '', regex=True).str.upper()
        artist_cleaned = df_cleaned['Artist'].fillna('UNKNOWNARTIST').astype(str).str.replace(r'\s+', '', regex=True).str.upper()
        isrc_fill = song_cleaned + " - " + artist_cleaned
        df_cleaned['ISRC'] = df_cleaned['ISRC'].fillna(isrc_fill)
        df_cleaned.loc[df_cleaned['ISRC'].astype(str).str.strip() == '', 'ISRC'] = isrc_fill


    # 4. Exclude Taiwan songs (ISRC starting with 'TW') and drop empty ISRC
    # if 'ISRC' in df_cleaned.columns:
    #     df_cleaned = df_cleaned[~df_cleaned['ISRC'].str.startswith('TW', na=False)]
    #     df_cleaned = df_cleaned.dropna(subset=['ISRC'])

    # Capitalize all the letters of ISRC
    if 'ISRC' in df_cleaned.columns:
        df_cleaned['ISRC'] = df_cleaned['ISRC'].str.upper()
        
    if 'YearMonth' in df_cleaned.columns:
        df_cleaned.sort_values('YearMonth', inplace=True)
        
    return df_cleaned
