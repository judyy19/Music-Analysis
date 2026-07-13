"""
Catalog logic for managing the song baseline metadata and right holder split ratios.
Provides functions to initialize tables, seed from rate spreadsheets,
and fetch catalog details.
"""
import os
import pandas as pd
import sqlite3

def init_catalog_tables(conn: sqlite3.Connection):
    """
    Initialize SQLite tables for song metadata and royalty splits.
    
    Args:
        conn (sqlite3.Connection): Database connection object.
    """
    cursor = conn.cursor()
    
    # catalog_songs stores physical metadata about each music asset
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS catalog_songs (
            isrc TEXT PRIMARY KEY,
            song_name TEXT,
            artist_name TEXT,
            album_name TEXT,
            upc TEXT
        )
    """)
    
    # catalog_splits stores royalty split shares for right holders per ISRC
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS catalog_splits (
            isrc TEXT,
            right_holder TEXT,
            percentage REAL,
            PRIMARY KEY (isrc, right_holder)
        )
    """)
    
    # Index on right_holder to speed up dropdown populating and statements querying
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_catalog_splits_holder ON catalog_splits(right_holder)")
    
    conn.commit()

def import_catalog_from_excel(conn: sqlite3.Connection, excel_path: str):
    """
    Load royalty split percentages from Royalty_Fee_Rate.xlsx into catalog_splits.
    
    Args:
        conn (sqlite3.Connection): Database connection object.
        excel_path (str): File path to Royalty_Fee_Rate.xlsx.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Royalty fee rate excel not found at: {excel_path}")
        
    df = pd.read_excel(excel_path)
    
    # Standardize column names
    df.columns = [col.strip().lower() for col in df.columns]
    
    # Map raw columns to required keys
    isrc_col = next((c for c in df.columns if 'isrc' in c), None)
    holder_col = next((c for c in df.columns if 'holder' in c or 'right' in c), None)
    pct_col = next((c for c in df.columns if 'percent' in c or 'rate' in c or 'share' in c), None)
    
    if not isrc_col or not holder_col or not pct_col:
        raise ValueError(f"Missing required columns in rate sheet. Columns found: {df.columns.tolist()}")
        
    # Extract clean series
    df_clean = pd.DataFrame({
        'isrc': df[isrc_col].astype(str).str.strip().str.upper(),
        'right_holder': df[holder_col].astype(str).str.strip(),
        'percentage': pd.to_numeric(df[pct_col], errors='coerce').fillna(0.0)
    })
    
    # Remove null or empty entries
    df_clean = df_clean[df_clean['isrc'].notna() & (df_clean['isrc'] != '') & (df_clean['isrc'] != 'NAN')]
    
    # Upsert into database
    cursor = conn.cursor()
    
    # To keep imports clean, we clear previous records in catalog_splits
    cursor.execute("DELETE FROM catalog_splits")
    
    for _, row in df_clean.iterrows():
        cursor.execute("""
            INSERT OR REPLACE INTO catalog_splits (isrc, right_holder, percentage)
            VALUES (?, ?, ?)
        """, (row['isrc'], row['right_holder'], row['percentage']))
        
    conn.commit()

def add_song_metadata(conn: sqlite3.Connection, isrc: str, song_name: str, artist_name: str, album_name: str, upc: str):
    """
    Add or update song physical metadata in the catalog database.
    
    Args:
        conn (sqlite3.Connection): Database connection object.
        isrc (str): ISRC of the song.
        song_name (str): Title of the song.
        artist_name (str): Name of the artist.
        album_name (str): Name of the album.
        upc (str): UPC code of the album.
    """
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR REPLACE INTO catalog_songs (isrc, song_name, artist_name, album_name, upc)
        VALUES (?, ?, ?, ?, ?)
    """, (isrc.strip().upper(), song_name, artist_name, album_name, upc))
    conn.commit()

def add_split_ratio(conn: sqlite3.Connection, isrc: str, right_holder: str, percentage: float):
    """
    Manually add or update a right holder split ratio for a specific song.
    
    Args:
        conn (sqlite3.Connection): Database connection object.
        isrc (str): Standardized ISRC.
        right_holder (str): Name of the artist/payee.
        percentage (float): Split percentage (decimal, e.g., 0.50).
    """
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR REPLACE INTO catalog_splits (isrc, right_holder, percentage)
        VALUES (?, ?, ?)
    """, (isrc.strip().upper(), right_holder.strip(), float(percentage)))
    conn.commit()

def get_split_ratios(conn: sqlite3.Connection, isrc: str) -> dict:
    """
    Get all split ratios registered for an ISRC.
    
    Args:
        conn (sqlite3.Connection): Database connection object.
        isrc (str): Standardized ISRC.
        
    Returns:
        dict: Mapping of right_holder -> percentage.
    """
    cursor = conn.cursor()
    cursor.execute("""
        SELECT right_holder, percentage 
        FROM catalog_splits 
        WHERE isrc = ?
    """, (isrc.strip().upper(),))
    
    return {row[0]: row[1] for row in cursor.fetchall()}

def get_all_right_holders(conn: sqlite3.Connection) -> list:
    """
    Retrieve list of all unique right holders registered in the split table.
    
    Args:
        conn (sqlite3.Connection): Database connection object.
        
    Returns:
        list: Alphabetically sorted list of right holders.
    """
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT right_holder FROM catalog_splits")
    holders = [row[0] for row in cursor.fetchall() if row[0]]
    return sorted(holders)

def get_all_songs(conn: sqlite3.Connection) -> list:
    """
    Retrieve all songs with their metadata and aggregated splits info.
    
    Args:
        conn (sqlite3.Connection): Database connection.
        
    Returns:
        list: List of dicts describing songs.
    """
    cursor = conn.cursor()
    cursor.execute("""
        SELECT s.isrc, s.song_name, s.artist_name, s.album_name, s.upc,
               (SELECT GROUP_CONCAT(right_holder || ': ' || (percentage * 100) || '%', ', ')
                FROM catalog_splits WHERE isrc = s.isrc) as splits
        FROM catalog_songs s
    """)
    songs = []
    for row in cursor.fetchall():
        songs.append({
            'isrc': row[0],
            'song_name': row[1],
            'artist_name': row[2],
            'album_name': row[3],
            'upc': row[4],
            'splits_summary': row[5] or "No splits registered"
        })
    return songs
