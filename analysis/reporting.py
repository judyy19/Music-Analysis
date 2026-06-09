"""
Reporting module for TME Music data aggregation and report export.
"""
import pandas as pd

def generate_song_report(df_filtered: pd.DataFrame, output_path: str) -> pd.DataFrame:
    """Generate and save Song Revenue and Clicks report pivoted by YearMonth."""
    if df_filtered.empty:
        print("Warning: Input data is empty. Skipping song report.")
        return pd.DataFrame()

    # 1. Group by song attributes and month
    revenue_by_song_per_month = (
        df_filtered
        .groupby(['UPC', 'ISRC', 'YearMonth'])
        .agg(
            song=('standard_song', 'first'),
            artist=('standard_artist', 'first'),
            album=('Album', 'first'),
            revenue=('Revenue', 'sum'),
            clicks=('Total_Clicks', 'sum')
        )
        .reset_index()
    )
    
    # 2. Pivot table (MultiIndex columns: [clicks/revenue] x [YearMonth])
    song_pivot = revenue_by_song_per_month.pivot_table(
        index=['ISRC', 'song', 'artist', 'UPC', 'album'],
        columns='YearMonth',
        values=['clicks', 'revenue'],
        aggfunc='sum',
        fill_value=0
    )
    
    # 3. Calculate horizontal sums using MultiIndex slicing
    click_cols = [col for col in song_pivot.columns if col[0] == 'clicks']
    rev_cols = [col for col in song_pivot.columns if col[0] == 'revenue']
    
    song_pivot[('total_click', '')] = song_pivot[click_cols].sum(axis=1)
    song_pivot[('total_revenue', '')] = song_pivot[rev_cols].sum(axis=1)
    
    # Extract month list
    months = sorted(list(set(ym for metric, ym in song_pivot.columns if ym != '' and metric in ['clicks', 'revenue'])))
    
    # Reset index to move row index into columns
    song_pivot_df = song_pivot.reset_index()
    
    # 4. Flatten columns and reorder
    flat_cols = []
    for col in song_pivot_df.columns:
        level0, level1 = col
        if level1 == '':
            flat_cols.append(level0)
        else:
            flat_cols.append(f"{level1}_{level0}")
            
    song_pivot_df.columns = flat_cols
    
    # Build desired column order
    base_cols = ['ISRC', 'song', 'artist', 'UPC', 'album']
    summary_cols = ['total_click', 'total_revenue']
    time_cols = []
    for m in months:
        time_cols.extend([f"{m}_clicks", f"{m}_revenue"])
        
    final_cols = base_cols + summary_cols + time_cols
    song_final = song_pivot_df[final_cols].copy()
    
    # Sort by total revenue descending
    song_final = song_final.sort_values('total_revenue', ascending=False).reset_index(drop=True)
    
    # Export to Excel
    song_final.to_excel(output_path, index=False)
    print(f"Song report exported successfully to: {output_path}")
    return song_final

def generate_album_report(df_filtered: pd.DataFrame, output_path: str) -> pd.DataFrame:
    """Generate and save Album Revenue report pivoted by YearMonth."""
    if df_filtered.empty:
        print("Warning: Input data is empty. Skipping album report.")
        return pd.DataFrame()
        
    # 1. Group by album attributes and month
    revenue_by_album_per_month = (
        df_filtered
        .groupby(['UPC', 'YearMonth'])
        .agg(
            album=('Album', 'first'),
            artist=('standard_artist', 'first'),
            revenue=('Revenue', 'sum')
        )
        .reset_index()
    )
    
    # 2. Pivot table (SingleIndex columns of YearMonth strings)
    album_pivot = revenue_by_album_per_month.pivot_table(
        index=['UPC', 'album', 'artist'],
        columns='YearMonth',
        values='revenue',
        aggfunc='sum',
        fill_value=0
    )
    
    # 3. Calculate row totals
    album_pivot['total_revenue'] = album_pivot.sum(axis=1)
    
    # Reset index to move row index into columns
    album_pivot_df = album_pivot.reset_index()
    
    # Sort months chronologically
    months = sorted([col for col in album_pivot.columns if col != 'total_revenue'])
    
    # Reorder columns: UPC, album, artist, total_revenue, then months
    final_cols = ['UPC', 'album', 'artist', 'total_revenue'] + months
    album_final = album_pivot_df[final_cols].copy()
    
    # Sort by total revenue descending
    album_final = album_final.sort_values('total_revenue', ascending=False).reset_index(drop=True)
    
    # Export to Excel
    album_final.to_excel(output_path, index=False)
    print(f"Album report exported successfully to: {output_path}")
    return album_final

def generate_song_report_by_year(df_filtered: pd.DataFrame, output_path: str) -> pd.DataFrame:
    """Generate and save Song Revenue and Clicks report pivoted by Year."""
    if df_filtered.empty:
        print("Warning: Input data is empty. Skipping yearly song report.")
        return pd.DataFrame()

    # 1. Create a copy and extract Year (YYYY) from YearMonth (YYYY-MM)
    df = df_filtered.copy()
    if 'YearMonth' in df.columns:
        df['Year'] = df['YearMonth'].str.slice(0, 4)
    else:
        df['Year'] = 'Unknown'

    # 2. Group by song attributes (excluding UPC/Album as requested) and Year
    revenue_by_song_per_year = (
        df
        .groupby(['ISRC', 'Year'])
        .agg(
            song=('standard_song', 'first'),
            artist=('standard_artist', 'first'),
            revenue=('Revenue', 'sum'),
            clicks=('Total_Clicks', 'sum')
        )
        .reset_index()
    )
    
    # 3. Pivot table (MultiIndex columns: [clicks/revenue] x [Year])
    song_pivot = revenue_by_song_per_year.pivot_table(
        index=['ISRC', 'song', 'artist'],
        columns='Year',
        values=['clicks', 'revenue'],
        aggfunc='sum',
        fill_value=0
    )
    
    # 4. Calculate horizontal sums using MultiIndex slicing
    click_cols = [col for col in song_pivot.columns if col[0] == 'clicks']
    rev_cols = [col for col in song_pivot.columns if col[0] == 'revenue']
    
    song_pivot[('total_click', '')] = song_pivot[click_cols].sum(axis=1)
    song_pivot[('total_revenue', '')] = song_pivot[rev_cols].sum(axis=1)
    
    # Extract year list
    years = sorted(list(set(y for metric, y in song_pivot.columns if y != '' and metric in ['clicks', 'revenue'])))
    
    # Reset index to move row index into columns
    song_pivot_df = song_pivot.reset_index()
    
    # 5. Flatten columns and reorder
    flat_cols = []
    for col in song_pivot_df.columns:
        level0, level1 = col
        if level1 == '':
            flat_cols.append(level0)
        else:
            flat_cols.append(f"{level1}_{level0}")
            
    song_pivot_df.columns = flat_cols
    
    # Build desired column order (only including ISRC, song, artist, total metrics, and year metrics)
    base_cols = ['ISRC', 'song', 'artist']
    summary_cols = ['total_click', 'total_revenue']
    
    # Group all years' clicks first, then all years' revenue
    time_cols = [f"{y}_clicks" for y in years] + [f"{y}_revenue" for y in years]
        
    final_cols = base_cols + summary_cols + time_cols
    song_final = song_pivot_df[final_cols].copy()
    
    # Sort by total revenue descending
    song_final = song_final.sort_values('total_revenue', ascending=False).reset_index(drop=True)
    
    # Export to Excel
    song_final.to_excel(output_path, index=False)
    print(f"Yearly song report exported successfully to: {output_path}")
    return song_final

