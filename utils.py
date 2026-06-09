"""
Utility functions for text normalization, schema mapping, and cleaning.
"""
import re
import numpy as np
import pandas as pd

def normalize_col(col: str) -> str:
    """Normalize column name by lowercasing and replacing separators with space."""
    col = str(col).lower()
    col = re.sub(r'[_\-]', ' ', col)
    col = re.sub(r'\s+', ' ', col)
    return col.strip()

def match_column(col: str, schema: dict, threshold: int = 95) -> str:
    """Match original column names to standard column name using schema aliases."""
    col_norm = str(col).lower().strip()
    for standard, aliases in schema.items():
        choices = [str(a).lower().strip() for a in aliases]
        if col_norm in choices:
            return standard
    return None

def auto_map_columns(columns: list, schema: dict) -> dict:
    """Create a dictionary mapping of original column name -> standard column name."""
    rename_map = {}
    for col in columns:
        matched = match_column(col, schema)
        if matched:
            rename_map[col] = matched
    return rename_map

def read_excel_auto_schema(file_path: str, schema: dict) -> pd.DataFrame:
    """Read excel file header and only read columns that match the standard schema."""
    try:
        # Read header first
        cols = pd.read_excel(file_path, nrows=0).columns
        rename_map = auto_map_columns(cols, schema)
        
        # Only read matched columns
        df = pd.read_excel(file_path, usecols=list(rename_map.keys()))
        df = df.rename(columns=rename_map)
        return df
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return None

def clean_upc(val) -> str:
    """Standardize UPC strings, removing decimals and prefixes and prefixing 'UPC-'."""
    if pd.isna(val):
        return np.nan
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    if val_str.startswith('UPC-'):
        val_str = val_str.replace('UPC-', '')
    return f"UPC-{val_str}"
