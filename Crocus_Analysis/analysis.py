"""
analysis.py
Crocus Fitness Dushanbe — Commercial Analytics Dashboard 2025
Author: Bogdan Khudoidodov
"""

import os
import logging
import argparse

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec

# ─────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────

CONFIG = {
    'attendance_csv': 'data/attendance_monthly.csv',
    'programs_xlsx':  'data/crocus_fitness_commercial_EN_2026.xlsx',
    'sheet_name':     '📊 DASHBOARD',
    'output_img':     'outputs/crocus_dashboard.png',
}

THEME = {
    'bg':      '#0f1117',
    'card':    '#1a1d27',
    'text':    '#e8eaf0',
    'subtext': '#8b8fa8',
    'grid':    '#2a2d3a',
    'blue':    '#00d4ff',
    'red':     '#ff6b35',
    'green':   '#39d353',
    'gold':    '#ffd700',
    'summer':  '#5a7fb5',
}

PRIORITY_COLORS = {
    'PILOT':    THEME['green'],
    'TOP':      THEME['gold'],
    'SCALE':    THEME['blue'],
    'VALIDATE': THEME['subtext'],
}

SEASONS = {
    'ramadan':      [3, 4],
    'post_ramadan': [5],
    'summer':       [6, 7, 8],
}

MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# ─────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────

def load_attendance(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        raise FileNotFoundError(f'Attendance file not found: {path}')
    logging.info('Loading attendance data...')
    return pd.read_csv(path)


def load_programs(path: str, sheet: str) -> pd.DataFrame:
    if not os.path.exists(path):
        raise FileNotFoundError(f'Programs file not found: {path}')
    logging.info('Loading programs data...')
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    df  = raw.iloc[16:26, [1, 3, 4, 5, 6, 8, 10]].copy()
    df.columns = ['program', 'type', 'spots', 'sessions', 'price', 'cycles', 'priority']
    return df.reset_index(drop=True)

# ─────────────────────────────────────────────────────────────
# PROCESSING
# ─────────────────────────────────────────────────────────────

def clean_programs(df: pd.DataFrame) -> pd.DataFrame:
    numeric = ['spots', 'sessions', 'price', 'cycles']
    df[numeric] = df[numeric].apply(pd.to_numeric, errors='coerce')
    df = df.dropna(subset=numeric).reset_index(drop=True)
    df['priority'] = df['priority'].str.extract(r'(PILOT|TOP|SCALE|VALIDATE)')
    df['type']     = df['type'].str.extract(r'(MINI-GROUP|STANDARD|KIDS)')
    return df


def validate_programs(df: pd.DataFrame) -> None:
    for col in ['spots', 'price', 'priority']:
        if df[col].isnull().any():
            raise ValueError(f'Missing values in column: {col}')


def calculate_metrics(df: pd.DataFrame) -> pd.DataFrame:
    df['rev_cycle'] = df['spots'] * df['price']
    df['rev_100']   = df['rev_cycle'] * df['cycles']
    df['rev_70']    = (df['rev_100'] * 0.70).round(0)
    df['breakeven'] = np.ceil(df['rev_cycle'] * 0.30 / df['price']).astype(int)
    return df

# ─────────────────────────────────────────────────────────────
# AGGREGATIONS
# ─────────────────────────────────────────────────────────────

def aggregate(att: pd.DataFrame, prog: pd.DataFrame) -> dict:
    return {
        'monthly':     att.groupby('month')['monthly_visits'].sum().values,
        'zone_annual': att.groupby('zone')['monthly_visits'].sum().sort_values(),
        'util':        att.groupby(['month', 'zone_type'])['utilization_pct'].mean().unstack(),
        'pilots':      prog[prog['priority'] == 'PILOT'],
    }

# ─────────────────────────────────────────────────────────────
# PLOT HELPERS
# ─────────────────────────────────────────────────────────────

def style_ax(ax, title: str, grid_axis: str = 'y') -> None:
    ax.set_facecolor(THEME['card'])
    for s in ax.spines.values():
        s.set_color(THEME['grid'])
    ax.tick_params(colors=THEME['subtext'], labelsize=8)
    ax.set_title(title, color=THEME['text'], fontsize=10, fontweight='bold', pad=10)
    ax.grid(axis=grid_axis, color=THEME['grid'], linewidth=0.5, alpha=0.7)


def label_bars(ax, bars, values, fmt='{:.1f}k', div=1000, offset=200) -> None:
    for b, v in zip(bars, values):
        ax.text(b.get_x() + b.get_width()/2, b.get_height() + offset,
                fmt.format(v / div), ha='center', va='bottom',
                color=THEME['text'], fontsize=7.5, fontweight='bold')


def label_hbars(ax, bars, values, fmt='{:.1f}k', div=1000, offset=300) -> None:
    for b, v in zip(bars, values):
        ax.text(v + offset, b.get_y() + b.get_height()/2,
                fmt.format(v / div), va='center',
                color=THEME['text'], fontsize=7.5, fontweight='bold')

# ─────────────────────────────────────────────────────────────
# CHARTS
# ─────────────────────────────────────────────────────────────

def plot_monthly(ax, monthly: np.ndarray) -> None:
    style_ax(ax, 'Monthly Zone Visits — All Zones (2025)')
    colors = [
        THEME['red']    if m in SEASONS['ramadan']      else
        THEME['green']  if m in SEASONS['post_ramadan'] else
        THEME['summer'] if m in SEASONS['summer']       else
        THEME['blue']
        for m in range(1, 13)
    ]
    bars = ax.bar(MONTHS, monthly, color=colors, width=0.65, zorder=3)
    label_bars(ax, bars, monthly)
    ax.annotate('Ramadan', xy=(1.5, monthly.min() * 0.95),
                color=THEME['red'], fontsize=8, ha='center')
    ax.annotate('Post-Ramadan boom', xy=(4, monthly.max() * 0.97),
                color=THEME['green'], fontsize=8, ha='center')
    ax.set_ylabel('Zone Visits per Month', fontsize=8, color=THEME['subtext'])
    ax.legend(handles=[mpatches.Patch(color=c, label=l) for l, c in [
        ('Ramadan', THEME['red']), ('Post-Ramadan', THEME['green']),
        ('Summer', THEME['summer']), ('Regular', THEME['blue']),
    ]], fontsize=7, labelcolor=THEME['text'],
       facecolor=THEME['card'], edgecolor=THEME['grid'], loc='lower right')


def plot_kpis(ax, att: pd.DataFrame, pilots: pd.DataFrame) -> None:
    ax.set_facecolor(THEME['card'])
    for s in ax.spines.values():
        s.set_color(THEME['grid'])
    ax.set_xticks([]); ax.set_yticks([])
    ax.set_title('Key Metrics', color=THEME['text'], fontsize=10, fontweight='bold', pad=10)
    kpis = [
        ('Active Members',      '2,500',                                    THEME['blue']),
        ('Regular Visitors',    '600',                                       THEME['green']),
        ('Annual Zone Visits',  f'{att["monthly_visits"].sum()/1000:.0f}k', THEME['gold']),
        ('Peak Utilization',    '75%',                                       THEME['red']),
        ('3 Pilots Rev (70%)',  f'{pilots["rev_70"].sum()/1000:.0f}k TJS',  THEME['green']),
        ('3 Pilots Rev (100%)', f'{pilots["rev_100"].sum()/1000:.0f}k TJS', THEME['blue']),
    ]
    for i, (label, val, col) in enumerate(kpis):
        y = 0.88 - i * 0.16
        ax.text(0.08, y, label, transform=ax.transAxes, color=THEME['subtext'], fontsize=8)
        ax.text(0.92, y, val,   transform=ax.transAxes, color=col,
                fontsize=10, fontweight='bold', ha='right')
        ax.add_line(plt.Line2D([0.05, 0.95], [y - 0.04, y - 0.04],
                    transform=ax.transAxes, color=THEME['grid'], linewidth=0.5))


def plot_zones(ax, zone_annual: pd.Series) -> None:
    style_ax(ax, 'Annual Visits by Zone', grid_axis='x')
    ax.grid(axis='y', visible=False)
    labels = [z.replace('_', ' ') for z in zone_annual.index]
    colors = [
        THEME['green'] if any(x in z for x in ['Kids', 'Martial', 'Cross']) else
        THEME['red']   if 'Gym' in z else
        THEME['blue']
        for z in zone_annual.index
    ]
    bars = ax.barh(labels, zone_annual.values, color=colors, height=0.6, zorder=3)
    label_hbars(ax, bars, zone_annual.values)
    ax.set_xlabel('Annual Zone Visits', fontsize=8, color=THEME['subtext'])


def plot_revenue(ax, prog: pd.DataFrame) -> None:
    style_ax(ax, 'Revenue Potential (TJS/year @ 70%)', grid_axis='x')
    ax.grid(axis='y', visible=False)
    short = [p[:16] + '..' if len(p) > 16 else p for p in prog['program']]
    bars  = ax.barh(short, prog['rev_70'],
                    color=[PRIORITY_COLORS[p] for p in prog['priority']],
                    height=0.6, zorder=3)
    label_hbars(ax, bars, prog['rev_70'].values, fmt='{:.0f}k')
    ax.legend(handles=[mpatches.Patch(color=c, label=l) for l, c in PRIORITY_COLORS.items()],
              fontsize=6.5, labelcolor=THEME['text'],
              facecolor=THEME['card'], edgecolor=THEME['grid'], loc='lower right')


def plot_utilization(ax, util: pd.DataFrame) -> None:
    style_ax(ax, 'Avg Utilization: Open Zones vs Sections (% of capacity)')
    ax.plot(MONTHS, util['open_zone'].values, color=THEME['red'],
            marker='o', lw=2, ms=5, label='Open Zones (Gym / Pool)')
    ax.plot(MONTHS, util['section'].values, color=THEME['green'],
            marker='s', lw=2, ms=5, label='Sections (Classes)')
    ax.fill_between(range(12), util['open_zone'].values, alpha=0.08, color=THEME['red'])
    ax.fill_between(range(12), util['section'].values,   alpha=0.08, color=THEME['green'])
    ax.set_xticks(range(12)); ax.set_xticklabels(MONTHS, fontsize=8)
    ax.set_ylabel('Utilization %', fontsize=8, color=THEME['subtext'])
    ax.axvspan(2, 3.5, alpha=0.10, color=THEME['red'])
    ax.axvspan(5, 7.9, alpha=0.08, color=THEME['summer'])
    ax.text(2.6, util.values.min() + 3, 'Ramadan', color=THEME['red'],   fontsize=7.5)
    ax.text(5.8, util.values.min() + 3, 'Summer',  color='#8ab4f8',      fontsize=7.5)
    ax.legend(fontsize=8, labelcolor=THEME['text'],
              facecolor=THEME['card'], edgecolor=THEME['grid'])


def plot_breakeven(ax, pilots: pd.DataFrame) -> None:
    style_ax(ax, 'Pilot Break-even Analysis')
    x = np.arange(len(pilots))
    ax.bar(x - 0.2, pilots['spots'].values,     0.35,
           label='Max Spots',  color=THEME['blue'],  alpha=0.7, zorder=3)
    ax.bar(x + 0.2, pilots['breakeven'].values, 0.35,
           label='Break-even', color=THEME['green'], zorder=3)
    ax.set_xticks(x)
    ax.set_xticklabels(["Women's\nStrength", "Kids\nAthletic", "Healthy\nBack"],
                       fontsize=8, color=THEME['text'])
    ax.set_ylabel('Spots', fontsize=8, color=THEME['subtext'])
    ax.legend(fontsize=7.5, labelcolor=THEME['text'],
              facecolor=THEME['card'], edgecolor=THEME['grid'])
    for i, (cap, be) in enumerate(zip(pilots['spots'].values, pilots['breakeven'].values)):
        ax.text(i, cap + 0.2, f'{int(be/cap*100)}%\nneeded',
                ha='center', color=THEME['subtext'], fontsize=7)

# ─────────────────────────────────────────────────────────────
# DASHBOARD
# ─────────────────────────────────────────────────────────────

def build_dashboard(att: pd.DataFrame, prog: pd.DataFrame, agg: dict, output: str) -> None:
    fig = plt.figure(figsize=(18, 14), facecolor=THEME['bg'])
    gs  = GridSpec(3, 3, figure=fig, hspace=0.45, wspace=0.35)

    plot_monthly    (fig.add_subplot(gs[0, :2]), agg['monthly'])
    plot_kpis       (fig.add_subplot(gs[0,  2]), att, agg['pilots'])
    plot_zones      (fig.add_subplot(gs[1, :2]), agg['zone_annual'])
    plot_revenue    (fig.add_subplot(gs[1,  2]), prog)
    plot_utilization(fig.add_subplot(gs[2, :2]), agg['util'])
    plot_breakeven  (fig.add_subplot(gs[2,  2]), agg['pilots'])

    fig.text(0.5, 0.97,
             'CROCUS FITNESS DUSHANBE — Commercial Analytics Report 2025',
             ha='center', color=THEME['text'], fontsize=13, fontweight='bold')
    fig.text(0.5, 0.945,
             'Attendance Analysis · Revenue Modelling · Pilot Strategy  |  Bogdan Khudoidodov',
             ha='center', color=THEME['subtext'], fontsize=8.5)

    os.makedirs(os.path.dirname(output), exist_ok=True)
    plt.savefig(output, dpi=150, bbox_inches='tight', facecolor=THEME['bg'])
    logging.info(f'Dashboard saved → {output}')

# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(description='Crocus Fitness Analytics Dashboard')
    parser.add_argument('--output', default=CONFIG['output_img'], help='Output image path')
    args = parser.parse_args()

    att  = load_attendance(CONFIG['attendance_csv'])
    prog = load_programs(CONFIG['programs_xlsx'], CONFIG['sheet_name'])
    prog = clean_programs(prog)
    validate_programs(prog)
    prog = calculate_metrics(prog)

    agg = aggregate(att, prog)
    build_dashboard(att, prog, agg, args.output)


if __name__ == '__main__':
    main()