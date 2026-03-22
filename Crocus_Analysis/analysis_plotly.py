"""
analysis_plotly.py
Crocus Fitness Dushanbe — Interactive Commercial Analytics Dashboard 2025
Author: Bogdan Khudoidodov (Plotly version)
"""

import os
import logging
import argparse

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ─────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────

CONFIG = {
    'attendance_csv': 'data/attendance_monthly.csv',
    'programs_xlsx':  'data/crocus_fitness_commercial_EN_2026.xlsx',
    'sheet_name':     '📊 DASHBOARD',
    'output_html':    'outputs/crocus_dashboard.html',
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
    initial_len = len(df)
    numeric = ['spots', 'sessions', 'price', 'cycles']
    df[numeric] = df[numeric].apply(pd.to_numeric, errors='coerce')
    df = df.dropna(subset=numeric).reset_index(drop=True)
    dropped = initial_len - len(df)
    if dropped > 0:
        logging.warning(f'Dropped {dropped} rows due to invalid numeric data.')
    else:
        logging.info('No rows dropped during cleaning.')
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
# DASHBOARD
# ─────────────────────────────────────────────────────────────

def build_dashboard(att: pd.DataFrame, prog: pd.DataFrame, agg: dict, output: str) -> None:
    fig = make_subplots(
        rows=3, cols=3,
        specs=[
            [{'colspan': 2}, None, {'type': 'table'}],
            [{'colspan': 2}, None, {'type': 'xy'}],
            [{'colspan': 2}, None, {'type': 'xy'}],
        ],
        subplot_titles=(
            'Monthly Zone Visits — All Zones (2025)', 'Key Metrics',
            'Annual Visits by Zone',                  'Revenue Potential (TJS/year @ 70%)',
            'Avg Utilization: Open Zones vs Sections','Pilot Break-even Analysis',
        ),
        vertical_spacing=0.10,
        horizontal_spacing=0.08,
    )

    # 1 — Monthly bar
    monthly_colors = [
        THEME['red']    if m in SEASONS['ramadan']      else
        THEME['green']  if m in SEASONS['post_ramadan'] else
        THEME['summer'] if m in SEASONS['summer']       else
        THEME['blue']
        for m in range(1, 13)
    ]
    fig.add_trace(go.Bar(
        x=MONTHS, y=agg['monthly'],
        marker_color=monthly_colors,
        text=[f'{v/1000:.1f}k' for v in agg['monthly']],
        textposition='outside',
        name='Visits',
    ), row=1, col=1)

    # 2 — KPI table
    kpi_labels = ['Active Members', 'Regular Visitors', 'Annual Zone Visits',
                  'Peak Utilization', '3 Pilots Rev (70%)', '3 Pilots Rev (100%)']
    kpi_values = [
        '<b>2,500</b>', '<b>600</b>',
        f'<b>{att["monthly_visits"].sum()/1000:.0f}k</b>',
        '<b>75%</b>',
        f'<b>{agg["pilots"]["rev_70"].sum()/1000:.0f}k TJS</b>',
        f'<b>{agg["pilots"]["rev_100"].sum()/1000:.0f}k TJS</b>',
    ]
    kpi_colors = [THEME['blue'], THEME['green'], THEME['gold'],
                  THEME['red'], THEME['green'], THEME['blue']]
    fig.add_trace(go.Table(
        header=dict(
            values=['Metric', 'Value'],
            fill_color=THEME['bg'],
            font=dict(color=THEME['bg'], size=1),  # скрываем заголовок цветом фона
            height=4,
        ),
        cells=dict(
            values=[kpi_labels, kpi_values],
            align=['left', 'right'],
            font=dict(color=[[[THEME['subtext']] * 6, kpi_colors]], size=14),
            fill_color=THEME['card'],
            line_color=THEME['grid'],
            height=35,
        ),
    ), row=1, col=3)

    # 3 — Annual visits by zone (horizontal bar)
    zones_labels = [z.replace('_', ' ') for z in agg['zone_annual'].index]
    zones_colors = [
        THEME['green'] if any(x in z for x in ['Kids', 'Martial', 'Cross']) else
        THEME['red']   if 'Gym' in z else
        THEME['blue']
        for z in agg['zone_annual'].index
    ]
    fig.add_trace(go.Bar(
        y=zones_labels, x=agg['zone_annual'].values,
        orientation='h',
        marker_color=zones_colors,
        text=[f'{v/1000:.1f}k' for v in agg['zone_annual'].values],
        textposition='outside',
        name='Annual Visits',
    ), row=2, col=1)

    # 4 — Revenue potential (horizontal bar)
    short_progs = [p[:16] + '..' if len(p) > 16 else p for p in prog['program']]
    fig.add_trace(go.Bar(
        y=short_progs, x=prog['rev_70'],
        orientation='h',
        marker_color=[PRIORITY_COLORS[p] for p in prog['priority']],
        text=[f'{v/1000:.0f}k' for v in prog['rev_70'].values],
        textposition='outside',
        name='Revenue 70%',
    ), row=2, col=3)

    # 5 — Utilization lines
    fig.add_trace(go.Scatter(
        x=MONTHS, y=agg['util']['open_zone'].values,
        mode='lines+markers', name='Open Zones (Gym/Pool)',
        line=dict(color=THEME['red'], width=3),
        fill='tozeroy', fillcolor='rgba(255,107,53,0.1)',
    ), row=3, col=1)
    fig.add_trace(go.Scatter(
        x=MONTHS, y=agg['util']['section'].values,
        mode='lines+markers', name='Sections (Classes)',
        line=dict(color=THEME['green'], width=3),
        fill='tozeroy', fillcolor='rgba(57,211,83,0.1)',
    ), row=3, col=1)

    # FIX: add_vrect с row/col не работает когда в figure есть Table.
    # Используем add_shape с явным xref='x4' (row=3,col=1 → xaxis4)
    # x-ось категориальная — индексы: Mar=2, Apr=3, Jun=5, Aug=7
    for x0, x1, color, label in [
        (1.5, 3.5, THEME['red'],    'Ramadan'),
        (4.5, 7.5, THEME['summer'], 'Summer'),
    ]:
        fig.add_shape(
            type='rect', x0=x0, x1=x1, y0=0, y1=1,
            xref='x4', yref='y4 domain',
            fillcolor=color, opacity=0.12, line_width=0,
        )
        fig.add_annotation(
            x=(x0 + x1) / 2, y=1.05,
            xref='x4', yref='y4 domain',
            text=label, showarrow=False,
            font=dict(color=color, size=11),
        )

    # 6 — Break-even grouped bars
    # FIX: barmode='group' глобальный — isolate через offset вручную не нужен,
    # т.к. только этот subplot имеет 2 bar trace → grouping применится только здесь
    pilot_names = [p.replace(' ', '<br>') for p in agg['pilots']['program'].values]
    fig.add_trace(go.Bar(
        x=pilot_names, y=agg['pilots']['spots'].values,
        name='Max Spots', marker_color=THEME['blue'], opacity=0.7,
        offsetgroup=0,
    ), row=3, col=3)
    fig.add_trace(go.Bar(
        x=pilot_names, y=agg['pilots']['breakeven'].values,
        name='Break-even', marker_color=THEME['green'],
        text=[f'{int(be/cap*100)}% needed'
              for cap, be in zip(agg['pilots']['spots'].values, agg['pilots']['breakeven'].values)],
        textposition='outside',
        offsetgroup=1,
    ), row=3, col=3)

    # Layout
    fig.update_layout(
        title=dict(
            text=(
                '<b>CROCUS FITNESS DUSHANBE — Commercial Analytics Report 2025</b><br>'
                '<sup>Attendance Analysis · Revenue Modelling · Pilot Strategy  |  Bogdan Khudoidodov</sup>'
            ),
            font=dict(size=22, color=THEME['text']),
            x=0.5, xanchor='center',
        ),
        template='plotly_dark',
        paper_bgcolor=THEME['bg'],
        plot_bgcolor=THEME['card'],
        font=dict(color=THEME['text']),
        showlegend=True,
        legend=dict(bgcolor=THEME['card'], bordercolor=THEME['grid'], borderwidth=1),
        height=1100,
        barmode='group',
    )

    fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor=THEME['grid'])
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor=THEME['grid'])
    # Убираем y-grid для горизонтальных баров
    fig.update_yaxes(showgrid=False, row=2, col=1)
    fig.update_yaxes(showgrid=False, row=2, col=3)

    os.makedirs(os.path.dirname(output), exist_ok=True)
    fig.write_html(output)
    logging.info(f'Interactive dashboard saved → {output}')

# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(description='Crocus Fitness Interactive Dashboard')
    parser.add_argument('--output', default=CONFIG['output_html'], help='Output HTML path')
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