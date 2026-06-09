"""
Data Visualization module for plotting TME Music analysis charts.
"""
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
from matplotlib.ticker import StrMethodFormatter
from config import PLOT_STYLE

def set_style():
    """Apply global plotting configurations."""
    plt.rcParams['font.sans-serif'] = PLOT_STYLE['font_sans']
    plt.rcParams['axes.unicode_minus'] = False
    sns.set_theme(style="whitegrid", rc={"font.sans-serif": PLOT_STYLE['font_sans']})

def plot_revenue_trend(revenue_by_month: pd.DataFrame):
    """Plot monthly total revenue trend line graph."""
    set_style()
    fig, ax1 = plt.subplots(figsize=(12, 6))
    
    ax1.plot(
        revenue_by_month['YearMonth'], 
        revenue_by_month['revenue'], 
        color='steelblue', 
        marker='o', 
        linewidth=2, 
        label='Revenue'
    )
    ax1.set_xlabel('Year-Month', fontsize=12, labelpad=10)
    ax1.set_ylabel('Total Revenue', color='steelblue', fontsize=12)
    ax1.tick_params(axis='y', labelcolor='steelblue')
    ax1.yaxis.set_major_formatter(StrMethodFormatter('{x:,.0f}'))
    
    plt.xticks(rotation=90)
    plt.title('Revenue Trend of TME Music', fontsize=15, pad=15)
    fig.tight_layout()
    plt.show()

def plot_clicks_vs_revenue(revenue_by_song: pd.DataFrame, title: str = "Clicks vs Revenue Scatter Plot"):
    """Plot scatter plot comparing Clicks vs Revenue."""
    set_style()
    plt.figure(figsize=(10, 6))
    plt.scatter(
        revenue_by_song['clicks'], 
        revenue_by_song['revenue'], 
        color='royalblue', 
        alpha=0.6, 
        edgecolors='w', 
        s=60
    )
    plt.xlabel('Total Clicks', fontsize=12)
    plt.ylabel('Total Revenue', fontsize=12)
    plt.title(title, fontsize=15, pad=15)
    plt.grid(True, linestyle=':', alpha=0.5)
    plt.tight_layout()
    plt.show()

def plot_pie_charts(df_pie: pd.DataFrame):
    """Plot side-by-side Revenue and Clicks pie charts."""
    set_style()
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))
    
    # Colors
    colors = plt.cm.tab20c.colors
    explode = [0.1] + [0] * (len(df_pie) - 1)
    
    # 1. Revenue Pie Chart
    ax1.pie(
        df_pie['revenue'],
        labels=df_pie['song'],
        autopct='%1.1f%%',
        startangle=140,
        colors=colors,
        explode=explode,
        shadow=True
    )
    ax1.set_title('Revenue Share by Song', fontsize=14, weight='bold', pad=10)
    
    # 2. Clicks Pie Chart
    ax2.pie(
        df_pie['clicks'],
        labels=df_pie['song'],
        autopct='%1.1f%%',
        startangle=140,
        colors=colors,
        explode=explode,
        shadow=True
    )
    ax2.set_title('Clicks Share by Song', fontsize=14, weight='bold', pad=10)
    
    plt.suptitle('Top Songs Contribution Analysis', fontsize=18, weight='bold', y=0.98)
    plt.tight_layout()
    plt.show()

def plot_monthly_revenue_by_top_songs(top_song_df: pd.DataFrame, ntop: int):
    """Plot monthly trend lines for the top N songs."""
    set_style()
    plt.figure(figsize=(14, 7))
    sns.lineplot(data=top_song_df, x='YearMonth', y='revenue', hue='song', marker='o', linewidth=2)
    plt.title(f'Monthly Revenue Trend by Top {ntop} Songs', fontsize=15, pad=15)
    plt.xlabel('Year-Month', fontsize=12)
    plt.ylabel('Revenue', fontsize=12)
    plt.xticks(rotation=90)
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', title='Song')
    plt.tight_layout()
    plt.show()

def plot_revenue_by_platform(revenue_by_platform: pd.DataFrame):
    """Plot platform revenue trend over time."""
    set_style()
    plt.figure(figsize=(14, 7))
    sns.lineplot(data=revenue_by_platform, x='YearMonth', y='revenue', hue='Platform', marker='o', linewidth=2)
    plt.title('Revenue by Platform', fontsize=15, pad=15)
    plt.xlabel('Year-Month', fontsize=12)
    plt.ylabel('Revenue', fontsize=12)
    plt.xticks(rotation=90)
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', title='Platform')
    plt.tight_layout()
    plt.show()

def plot_revenue_by_album(top_album_df: pd.DataFrame):
    """Plot monthly trend lines for the top albums."""
    set_style()
    plt.figure(figsize=(14, 7))
    sns.lineplot(data=top_album_df, x='YearMonth', y='revenue', hue='album', marker='o', linewidth=2)
    plt.title('Revenue by Album', fontsize=15, pad=15)
    plt.xlabel('Year-Month', fontsize=12)
    plt.ylabel('Revenue', fontsize=12)
    plt.xticks(rotation=90)
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', title='Album')
    plt.tight_layout()
    plt.show()
