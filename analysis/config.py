"""
Configuration settings and constants for the TME Music Analysis pipeline.
"""
import os

# Base paths
BASE_PATH = "/Users/chu-chun/Mirror/Eva/input/TME/"
OUTPUT_SONG_REPORT_PATH = "../output/TME_Song_Report_202401_202603.xlsx"
OUTPUT_ALBUM_REPORT_PATH = "../output/TME_Album_Revenue_Report_202306_202603.xlsx"


# Analysis parameters
START_MONTH = "2024-01"
END_MONTH = "2026-03"
TOP_N_SONGS = 10
TOP_N_ALBUMS = 10

# Schema definition for columns mapping (ETL mapping)
STANDARD_SCHEMA = {
    "Date": ["date", "period", "upload date", "日期", "结算期间"],
    "ISRC": ["isrc", "歌曲ISRC"],
    "Song": ["song", "mv", "music video", "track", "title", "歌曲名"],
    "Artist": ["artist", "歌手名"],
    "Revenue": ["license fees - total", "CP分成收入"],
    "Platform": ["platform", "平台"],
    "Album": ["album", "专辑名"],
    "UPC": ["upc", "专辑UPC"],
    "Clicks": ["clicks", "点击量"],
    
    # 1. Aiting Clicks
    "Aiting_Free": ["free music service"],
    "Aiting_Sub": ["consumption - subscription music service"],
    
    # 2. K Clicks
    "K_Lyrics": ["license fees-per play-consumption of lyrics"],
    "K_Comp": ["license fees-per play-consumption of composition"],
    "K_Rec_Orig": ["license fees-per play-consumption of recording-original version"],
    "K_Rec_Kara_Lic": ["license fees-per play-consumption of recording-karaoke version provided by licensor"],
    "K_Rec_Kara_TME": ["license fees-per play-consumption of recording-karaoke version processed by tme"],
    
    # 3. Single Clicks
    "Single_IOS": ["sales_ios"],
    "Single_Others": ["sales_others"],
    
    # 4. Song Clicks
    "Song_Free_Normal": ["license fees - free music service-free mode-number of content used"],
    "Song_Free_NonNormal": ["license fees - free music service-non-free mode-number of content used"],
    "Song_Sub_Basic": ["subscription music service(basic)"],
    "Song_Sub_Senior": ["subscription music service(senior)"],
    "Song_MuCoin": ["mucoin & gift"]
}

# Platform translation/unification map
PLATFORM_MAP = {
    'QQ音乐': 'QQMusic',
    '酷狗': 'Kugou',
    '酷我': 'Kuwo',
    '酷狗K歌': 'KugouKaraoke',
    '酷狗直播': 'KugouLive',
    '全民K歌': 'Wesing',
    '爱听': 'UltimateMusic'
}

# Plot styling configuration
PLOT_STYLE = {
    "font_sans": ['Arial Unicode MS'], # Support Chinese on Mac
    "figure_dpi": 100,
}
