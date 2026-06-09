"""
Configuration settings and constants for the TME Music Analysis pipeline.
"""
import os

# Base paths
BASE_PATH = "/Users/chu-chun/Mirror/Eva/input/sony_TME/大曲庫/2025/"
OUTPUT_SONG_REPORT_PATH = "../output/TME_Song_Report_for_Sony_大曲庫-2025.xlsx"
OUTPUT_SONG_REPORT_PATH_MONTHLY = "../output/TME_Song_Report_Monthly_for_Sony_大曲庫-2025.xlsx"
OUTPUT_ALBUM_REPORT_PATH = "../output/TME_Album_Revenue_Report_for_Sony.xlsx"

# Analysis parameters
# START_MONTH = "2024-01"
# END_MONTH = "2026-03"
TOP_N_SONGS = 10
TOP_N_ALBUMS = 10

# Schema definition for columns mapping (ETL mapping)
STANDARD_SCHEMA = {
    "Date": ["date", "period", "upload date", "日期", "结算期间"],
    "ISRC": ["isrc", "歌曲ISRC"],
    "Song": ["song", "mv", "music video", "track", "title", "歌曲名", "mv名"],
    "Artist": ["artist", "歌手名"],
    "Revenue": ["license fees - total", "CP分成收入"],
    "Platform": ["platform", "平台"],
    "Album": ["album", "专辑名"],
    "UPC": ["upc", "专辑UPC"],
    "Clicks": ["clicks", "点击量"],
    
    # 1. Aiting Clicks
    "Aiting_Free": ["free music service", "广告收入分成-使用量"],
    "Aiting_Sub": ["consumption - subscription music service", "包月收入分成-使用量"],
    
    # 2. K Clicks
    "K_Lyrics": ["license fees-per play-consumption of lyrics", "按次分成-词使用量"],
    "K_Comp": ["license fees-per play-consumption of composition", "按次分成-曲使用量"],
    "K_Rec_Orig": ["license fees-per play-consumption of recording-original version", "按次分成-邻接权使用量-原版音源"],
    "K_Rec_Kara_Lic": ["license fees-per play-consumption of recording-karaoke version provided by licensor", "按次分成-邻接权使用量-版权方提供伴奏"],
    "K_Rec_Kara_TME": ["license fees-per play-consumption of recording-karaoke version processed by tme", "按次分成-邻接权使用量-依据版权方提供音源制作伴奏"],
    
    # 3. Single Clicks
    "Single_IOS": ["sales_ios", "IOS销量"],
    "Single_Others": ["sales_others", "非IOS销量"],
    
    # 4. Song Clicks
    "Song_Free_Normal": ["license fees - free music service-free mode-number of content used","广告收入分成-使用量"],
    "Song_Free_NonNormal": ["license fees - free music service-non-free mode-number of content used"],
    "Song_Sub_Basic": ["subscription music service(basic)", "基本包月收入分成-使用量"],
    "Song_Sub_Senior": ["subscription music service(senior)", "高级包月收入分成-使用量"],
    "Song_MuCoin": ["mucoin & gift", "打榜收入"],

    # 5. MV Clicks
    "MV_Comp": ["License Fees - Free Music Service-Non-free Mode-Consumption of MV", "广告收入分成-非免模-MV使用量"]
}

# Platform translation/unification map
PLATFORM_MAP = {
    'QQ音乐': 'QQMusic',
    '酷狗': 'Kugou',
    '酷我': 'Kuwo',
    '酷狗K歌': 'KugouKaraoke',
    '酷狗直播': 'KugouLive',
    '全民K歌': 'Wesing',
    '爱听': 'UltimateMusic',
    '全民K歌国际版': 'Wesing_International'
}

# Plot styling configuration
PLOT_STYLE = {
    "font_sans": ['Arial Unicode MS'], # Support Chinese on Mac
    "figure_dpi": 100,
}
