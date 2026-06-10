"""
Configuration settings and constants for the TME Music Analysis pipeline.
"""
import os

# Base paths
# BASE_PATH = "/Users/chu-chun/Mirror/Eva/test/"
BASE_PATH = "/Users/chu-chun/Mirror/Eva/input/sony_TME/大曲庫/"
OUTPUT_SONG_REPORT_PATH = "../output/TME_Song_Report_for_Sony_大曲庫.xlsx"
OUTPUT_SONG_REPORT_PATH_MONTHLY = "../output/TME_Song_Report_Monthly_for_Sony_test.xlsx"
OUTPUT_ALBUM_REPORT_PATH = "../output/TME_Album_Revenue_Report_for_Sony.xlsx"


# Exclude file name
EXCLUDE_FILE_NAME = ['bill', '~$', '保底金明细', '数专', '音乐汇总单', '合同收入', 'digital_album']

# Analysis parameters
# START_MONTH = "2024-01"
# END_MONTH = "2026-03"
TOP_N_SONGS = 10
TOP_N_ALBUMS = 10

# Schema definition for columns mapping (ETL mapping)
STANDARD_SCHEMA = {
    "Date": ["date", "period", "upload date", "日期", "结算期间", "结算开始时间"],
    "ISRC": ["isrc", "歌曲ISRC"],
    "Song": ["song", "mv", "music video", "track", "title", "歌曲名", "mv名"],
    "Artist": ["artist", "歌手名"],
    "Revenue": ["license fees - total", "CP分成收入"],
    "Platform": ["platform", "平台", "结算平台", "结算平台ID（32=aiting）"],
    "Album": ["album", "专辑名"],
    "UPC": ["upc", "专辑UPC"],
    "Clicks": ["clicks", "点击量"],
    
    # 1. Aiting Click/ 愛聽
    "Aiting_Free": ["free music service", "广告收入分成-使用量"],
    "Aiting_Sub": ["consumption - subscription music service", "包月收入分成-使用量"],
    
    # 2. K Clicks/ K歌
    "K_Lyrics": ["license fees-per play-consumption of lyrics", "按次分成-词使用量", "K歌按次分成使用量"],
    "K_Comp": ["license fees-per play-consumption of composition", "按次分成-曲使用量"],
    "K_Rec_Orig": ["license fees-per play-consumption of recording-original version", "按次分成-邻接权使用量-原版音源"],
    "K_Rec_Kara_Lic": ["license fees-per play-consumption of recording-karaoke version provided by licensor", "按次分成-邻接权使用量-版权方提供伴奏"],
    "K_Rec_Kara_TME": ["license fees-per play-consumption of recording-karaoke version processed by tme", "按次分成-邻接权使用量-依据版权方提供音源制作伴奏"],
    
    # 3. Single Clicks/ 單曲
    "Single_IOS": ["sales_ios", "IOS销量"],
    "Single_Others": ["sales_others", "非IOS销量"],
    
    # 4. Song Clicks/ 歌曲
    "Song_Free_Normal": ["license fees - free music service-free mode-number of content used","广告收入分成-使用量"],
    "Song_Free_NonNormal": ["license fees - free music service-non-free mode-number of content used"],
    "Song_Sub_Basic": ["subscription music service(basic)", "基本包月收入分成-使用量"],
    "Song_Sub_Senior": ["subscription music service(senior)", "高级包月收入分成-使用量"],
    "Song_MuCoin": ["mucoin & gift", "打榜收入"],

    # 5. MV Clicks/ 
    "MV_Comp": ["License Fees - Free Music Service-Non-free Mode-Consumption of MV", "广告收入分成-非免模-MV使用量"]
}

# Platform translation/unification map
PLATFORM_MAP = {
    'QQ音乐': 'QQMusic',
    '酷狗': 'Kugou',
    '酷我': 'Kuwo',
    '酷狗音乐': 'Kugou',
    '酷我音乐': 'Kuwo',
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

# Artist alias mapping table (Keys should be uppercase and without whitespace for matching)
ARTIST_ALIAS_MAP = {
    '黄安祖': ['黄安祖', '安祖ANTZU'],
    '蔡黄汝': ['蔡黄汝', '蔡黃汝'],
    '王彩桦': ['王彩桦', '王彩樺'],
    '不才': ['不才', '蔡明希-不才'],
    '刘凡': ['刘凡', '刘凡FANCY'],
    '張萱妍': ['張萱妍', '张萱妍'],
    'AKB48TEAMTP': ['AKB48TEAMTP', 'AKB48'],
    '金大为': ['DK金大为', '金大为'],
    'THEPUZZLE5,KOHA': ['THEPUZZLE5,KOHA', 'THEPUZZLE5'],
    '云の泣': ['云の泣', '云之泣'],
    '吉克隽逸,장혁,朴宰范(JAYPARK)': ['吉克隽逸,장혁,朴宰范(JAYPARK)', '吉克隽逸,장혁,박재범'],
    '吉克隽逸,杭盖乐队': ['吉克隽逸,杭盖乐队', '吉克隽逸'],
    '郭京飞,안칠현,金圣洙': ['郭京飞,안칠현,金圣洙', '郭京飞,安七炫,金圣洙'],
    '顾晓宇,代轩齐': ['顾晓宇,代轩齐', '代轩齐,顾晓宇'],
    '黄昺翔SEANHUANG': ['黄昺翔SEANHUANG', '黄昺翔SEANH.', '黃昺翔'],
    '龙飞龙泽1983组合,刘彦英': ['龙飞龙泽1983组合,刘彦英', '龙飞龙泽,刘彦英'],
    'SHILAAMZAH,五洲唱响乐团': ['SHILAAMZAH,五洲唱响乐团', '茜拉(SHILAAMZAH),五洲唱响乐团', 'SHILAAMZAH,平安', 'SHILAAMZAH', '茜拉(SHILAAMZAH)', '茜拉(SHILAAMZAH),平安', '五洲唱响乐团, 茜拉 (Shila Amzah)'],
    '方泂鑌': ['方烱彬', '方泂鑌', '方炯鑌']
}
