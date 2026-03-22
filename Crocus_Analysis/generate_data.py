"""
generate_data.py
Generates attendance_monthly.csv based on real Crocus Fitness Dushanbe schedules.

Sources:
- БИ__2___РАСПИСАПНИЕ_2026_ВЕСНА_.xlsx
- расписание_гп_РАМАДАН.xlsx
- ДЕТСКИЕ_СЕКЦИИ_ВП_ДЕКАБРЬ.xlsx
- Расписание_ДК__1___2_.xlsx
- Operational data: ~2500 members, 600 regular, 260 evening / 80-100 morning
"""

import pandas as pd
import numpy as np

np.random.seed(42)

OUTPUT = 'data/attendance_monthly.csv'

# ── Zone definitions from real schedules ──────────────────────────────────────
# sessions_per_day: counted from actual schedule files
# avg_per_session:  estimated from capacity constraints + member data
# days_per_week:    from schedule (some zones closed Sunday)
# capacity:         from price list notes (10-18 limited, gyms open)

zones = {
    # БИ зона — 10-12 занятий/день, пн-сб (дети + взрослые разные виды)
    'Martial_Arts_Kids':   {'sessions': 6,  'avg': 11, 'capacity': 15, 'days': 6, 'type': 'section'},
    'Martial_Arts_Adults': {'sessions': 5,  'avg': 10, 'capacity': 12, 'days': 6, 'type': 'section'},

    # Групповые программы — 10-12 занятий/день, пн-вс
    'Group_Morning':       {'sessions': 4,  'avg': 14, 'capacity': 20, 'days': 7, 'type': 'section'},
    'Group_Evening':       {'sessions': 7,  'avg': 16, 'capacity': 22, 'days': 7, 'type': 'section'},

    # Cycle (ограничено 18 мест, пн/ср/пт вечер)
    'Cycle_Studio':        {'sessions': 3,  'avg': 15, 'capacity': 18, 'days': 5, 'type': 'section'},

    # Antigravity (ограничено 10 мест)
    'Antigravity':         {'sessions': 2,  'avg': 8,  'capacity': 10, 'days': 5, 'type': 'section'},

    # Детский бассейн — 5-6 слотов, пн-сб
    'Pool_Kids':           {'sessions': 5,  'avg': 11, 'capacity': 15, 'days': 6, 'type': 'section'},

    # Детский клуб секции — 8-10 занятий, пн-сб
    'Kids_Club':           {'sessions': 8,  'avg': 10, 'capacity': 12, 'days': 6, 'type': 'section'},

    # Тренажёрный зал — open zone, пн-вс
    # 260 вечером + 90 утром = ~350/день всего, делим муж/жен ~65/35
    'Gym_Male':            {'sessions': 1,  'avg': 95, 'capacity': 130, 'days': 7, 'type': 'open_zone'},
    'Gym_Female':          {'sessions': 1,  'avg': 55, 'capacity': 80,  'days': 7, 'type': 'open_zone'},

    # Взрослый бассейн — open zone
    'Pool_Main':           {'sessions': 1,  'avg': 40, 'capacity': 60,  'days': 7, 'type': 'open_zone'},
    'Pool_Female':         {'sessions': 1,  'avg': 20, 'capacity': 35,  'days': 6, 'type': 'open_zone'},
}

# ── Seasonality per month ─────────────────────────────────────────────────────
# Base coefficient applied to avg attendance
# Ramadan 2025: ~Mar-Apr | Summer slump: Jun-Aug | Post-Ramadan boom: May
seasonality = {
    1: 0.90, 2: 0.92,
    3: 0.68, 4: 0.63,   # Ramadan — gym hits hard
    5: 1.10,             # Post-Ramadan boom
    6: 0.82, 7: 0.75, 8: 0.78,  # Summer
    9: 1.05, 10: 1.08, 11: 1.05, 12: 0.88,
}

days_in_month = {1:31,2:28,3:31,4:30,5:31,6:30,7:31,8:31,9:30,10:31,11:30,12:31}

# ── Generate records ──────────────────────────────────────────────────────────
records = []

for month, base_s in seasonality.items():
    working_days = days_in_month[month]

    for zone, cfg in zones.items():
        # Sections: boom in summer (sports, kids active), drop in Ramadan
        # Open zones: harder drop in Ramadan and summer
        if cfg['type'] == 'section':
            if month in [6, 7, 8]:
                s = min(base_s * 1.25, 1.0)   # sections busier in summer
            elif month in [3, 4]:
                s = base_s * 1.05              # sections hold better in Ramadan
            else:
                s = base_s
        else:  # open_zone
            if month in [3, 4]:
                s = base_s * 0.80              # gym drops hardest in Ramadan
            elif month in [6, 7, 8]:
                s = base_s * 0.88
            else:
                s = base_s

        # Effective working days (days_per_week → monthly)
        effective_days = working_days * (cfg['days'] / 7)

        # Daily sessions × avg attendance per session
        daily_visits = cfg['sessions'] * cfg['avg'] * s
        daily_visits *= np.random.normal(1.0, 0.05)  # ±5% noise
        daily_visits = max(round(daily_visits), 0)

        monthly_visits = round(daily_visits * effective_days)
        capacity_day   = cfg['sessions'] * cfg['capacity']
        utilization    = round(min(daily_visits / capacity_day, 1.0) * 100, 1)

        records.append({
            'month':               month,
            'month_name':          pd.Timestamp(f'2025-{month:02d}-01').strftime('%B'),
            'zone':                zone,
            'zone_type':           cfg['type'],
            'sessions_per_day':    cfg['sessions'],
            'days_per_week':       cfg['days'],
            'avg_per_session':     cfg['avg'],
            'avg_daily_visits':    daily_visits,
            'monthly_visits':      monthly_visits,
            'capacity_per_day':    capacity_day,
            'utilization_pct':     utilization,
            'seasonality_coef':    round(s, 2),
        })

df = pd.DataFrame(records)
df.to_csv(OUTPUT, index=False)

# ── Sanity check ──────────────────────────────────────────────────────────────
print('Monthly totals (all zones):')
print(df.groupby('month')['monthly_visits'].sum().to_string())
print(f'\nTotal annual zone-visits : {df["monthly_visits"].sum():,}')
print(f'Avg daily (all zones)    : {df["avg_daily_visits"].sum()/12:.0f}')
print(f'\nAvg daily gym only       : {df[df["zone"].str.startswith("Gym")]["avg_daily_visits"].mean():.0f} per gym zone')
print(f'Zones                    : {df["zone"].nunique()}')
print(f'\nSaved → {OUTPUT}')