#!/usr/bin/env python3
"""
Wildwind Room Availability Updater
Downloads Excel from Dropbox and generates availability.html
"""

import json
import requests
import pandas as pd
from datetime import datetime
from pathlib import Path

# ─── CONFIG ───────────────────────────────────────────────────────────
DROPBOX_URL = "https://www.dropbox.com/scl/fi/i0ji499omsoc0yl4fuzo7/STEPH-VERSION-ALL-ROOMS-2026.xlsx?rlkey=lq3j8vz7hdlnoqvairj0yp2ih&dl=1"
OUTPUT_FILE = "availability.html"

# Excel-rad → visningsnamn (A+B kolumnerna, förkortat)
ROW_TO_NAME = {
    44:  "Melas 1",
    48:  "Melas 2AB",   # kombineras med rad 52
    56:  "Melas 3",
    60:  "Melas 4",
    72:  "Melas 7AB",   # kombineras med rad 76
    80:  "Melas 8",
    140: "Kav 2",
    144: "Kav 3",
    148: "Kav 4",
    184: "Kav 13",
    188: "Kav 14",
    192: "Kav 15",
    204: "Kav 18",
    208: "Kav 19",
    216: "Akti 7",
    220: "Akti 8",
    224: "Akti 9",
    244: "Akti 4",
    248: "Akti 5",
    272: "Xristina 6",
    284: "NM208",
}

# Kombinerade familjerum – bokad om NÅGON rad är bokad
COMBINED_ROOMS = {
    "Melas 2AB": [48, 52],
    "Melas 7AB": [72, 76],
}

# Sailing-priser från wildwind.se (SEK/person/v, 2 delar rum)
SAILING_PRICES = {
    "25 Apr": 8990,  "2 May":  9430,  "9 May":  10300, "16 May": 10840,
    "23 May": 10840, "30 May": 11930, "6 Jun":  12690, "13 Jun": 13020,
    "20 Jun": 13890, "27 Jun": 14110, "4 Jul":  14440, "11 Jul": 14980,
    "18 Jul": 15200, "25 Jul": 15200, "1 Aug":  15200, "8 Aug":  15200,
    "15 Aug": 14110, "22 Aug": 12800, "29 Aug": 11930, "5 Sep":  11710,
    "12 Sep": 11170, "19 Sep": 10300, "26 Sep": 9540,  "3 Oct":  8990,
}

# Alla priser från webben (SEK/person/v, 2 delar rum)
ALL_PRICES = {
    # datum: (1v Sailing, 2v Sailing, 1v FatW, 2v FatW, 1v HO, 2v HO, enkelrum)
    "25 Apr": (8990,  14110, None,  None,  7570,  10840, 1910),
    "2 May":  (9430,  14440, None,  None,  7900,  13020, 2130),
    "9 May":  (10300, 16290, None,  None,  8120,  13350, 2180),
    "16 May": (10840, 17380, 14000, 23690, 8230,  13350, 2400),
    "23 May": (10840, 18470, 14000, 24780, 7570,  14110, 2400),
    "30 May": (11930, 19340, 15090, 25660, 8660,  14440, 2670),
    "6 Jun":  (12690, 19560, 15850, 25870, 9540,  14650, 2780),
    "13 Jun": (13020, 20650, 16180, 26960, 9750,  15200, 2950),
    "20 Jun": (13890, 22610, 17050, 28920, 11380, 16290, 3050),
    "27 Jun": (14110, 22820, 17270, 29140, 11930, 17380, 3160),
    "4 Jul":  (14440, 22820, 17600, 29140, 12260, 17380, 3160),
    "11 Jul": (14980, 23690, 18140, 30010, 12260, 17380, 3160),
    "18 Jul": (15200, 23910, 18360, 30230, 12260, 17380, 3160),
    "25 Jul": (15200, 23910, 18360, 30230, 12260, 17380, 3160),
    "1 Aug":  (15200, 23910, 18360, 30230, 12260, 17380, 3160),
    "8 Aug":  (15200, 23910, 18360, 30230, 12260, 17380, 3160),
    "15 Aug": (14110, 21730, 17270, 28050, 12260, 17380, 3160),
    "22 Aug": (12800, 19880, 15960, 26200, 11930, 16290, 2730),
    "29 Aug": (11930, 18470, 15090, 24780, 9750,  15200, 2460),
    "5 Sep":  (11710, 17380, 14870, 23690, 9210,  14110, 2400),
    "12 Sep": (11170, 16290, 14330, 22610, 8990,  13890, 2130),
    "19 Sep": (10300, 15200, 13460, None,  8660,  12260, 2070),
    "26 Sep": (9540,  13570, None,  None,  7570,  11380, 1910),
    "3 Oct":  (8990,  13020, None,  None,  7360,  None,  1910),
}

# Ruminfo för mobilvy
ROOM_INFO = {
    "Melas 1":    {"tillagg": "0 kr",     "tillagg_num": 0,    "info": "Twin · Sidoutsikt hav + Kav Bar"},
    "Melas 2AB":  {"tillagg": "0 kr",     "tillagg_num": 0,    "info": "Twin+Double · Familjerum, havsutsikt"},
    "Melas 3":    {"tillagg": "1 758 kr", "tillagg_num": 1758, "info": "Double · Full havsutsikt"},
    "Melas 4":    {"tillagg": "1 758 kr", "tillagg_num": 1758, "info": "Twin · Full havsutsikt"},
    "Melas 7AB":  {"tillagg": "0 kr",     "tillagg_num": 0,    "info": "Double+Twin · Familjerum, mot Ponti"},
    "Melas 8":    {"tillagg": "0 kr",     "tillagg_num": 0,    "info": "Twin · Mot Ponti + berg västerut"},
    "Kav 2":      {"tillagg": "0 kr",     "tillagg_num": 0,    "info": "Small Double · Ingen utsikt, baksida"},
    "Kav 3":      {"tillagg": "0 kr",     "tillagg_num": 0,    "info": "Twin · Ingen utsikt, baksida (litet)"},
    "Kav 4":      {"tillagg": "0 kr",     "tillagg_num": 0,    "info": "Double · Sidoutsikt mot trädgård"},
    "Kav 13":     {"tillagg": "1 320 kr", "tillagg_num": 1320, "info": "Double · Pool + berg västerut, övervåning"},
    "Kav 14":     {"tillagg": "1 320 kr", "tillagg_num": 1320, "info": "Double · Pool + berg, delar balkong Kav 13"},
    "Kav 15":     {"tillagg": "1 320 kr", "tillagg_num": 1320, "info": "Double · Pool + berg västerut, övervåning"},
    "Kav 18":     {"tillagg": "1 320 kr", "tillagg_num": 1320, "info": "Twin · Havsutsikt över Kav Bar, övervåning"},
    "Kav 19":     {"tillagg": "1 320 kr", "tillagg_num": 1320, "info": "Twin · Sidoutsikt mot hav + byn"},
    "Akti 4":     {"tillagg": "1 758 kr", "tillagg_num": 1758, "info": "Studio · Twin, pentry, övervåning"},
    "Akti 5":     {"tillagg": "1 758 kr", "tillagg_num": 1758, "info": "Studio · Twin, pentry, övervåning"},
    "Akti 7":     {"tillagg": "3 516 kr", "tillagg_num": 3516, "info": "2-sovrum · Double+Twin+extra, fullt kök (min 4 pers)"},
    "Akti 8":     {"tillagg": "3 516 kr", "tillagg_num": 3516, "info": "2-sovrum · Double+Twin+extra, fullt kök (min 4 pers)"},
    "Akti 9":     {"tillagg": "1 758 kr", "tillagg_num": 1758, "info": "1-sovrum · Twin + extra i hall, pentry"},
    "Xristina 6": {"tillagg": "3 516 kr", "tillagg_num": 3516, "info": "1-sovrum · Double+2 enklar, kitchenette (13 jul–15 sep)"},
    "NM208":      {"tillagg": "2 638 kr", "tillagg_num": 2638, "info": "Double · Bergutsikt bakåt, övervåning"},
}

# ─── DOWNLOAD ─────────────────────────────────────────────────────────
def download_excel():
    print("📥 Laddar ner Excel från Dropbox...")
    r = requests.get(DROPBOX_URL, timeout=30)
    r.raise_for_status()
    tmp = Path("temp_availability.xlsx")
    tmp.write_bytes(r.content)
    print(f"✅ {len(r.content):,} bytes")
    return tmp

# ─── PARSE ────────────────────────────────────────────────────────────
def parse_availability(excel_path):
    print("📊 Läser tillgänglighet...")
    df = pd.read_excel(excel_path, header=None)

    # Hitta lördagskolumner
    saturdays = []
    for col in range(3, len(df.columns), 2):
        date_val = df.iloc[0, col]
        day_name = df.iloc[1, col]
        if pd.notna(day_name) and 'SATURDAY' in str(day_name).upper():
            if pd.notna(date_val):
                dt = pd.to_datetime(date_val)
                label = dt.strftime('%-d %b')
                prices = ALL_PRICES.get(label, (None,)*7)
                saturdays.append({
                    'col':       col,
                    'date':      dt.strftime('%Y-%m-%d'),
                    'display':   label,
                    'price':     SAILING_PRICES.get(label),  # 1v sailing
                    'sail2':     prices[1],
                    'fatw1':     prices[2],
                    'fatw2':     prices[3],
                    'ho1':       prices[4],
                    'ho2':       prices[5],
                    'single':    prices[6],
                })

    print(f"   {len(saturdays)} lördagar hittade")

    def get_week_statuses(r_idx):
        statuses = []
        for sat in saturdays:
            sat_col = sat['col']
            booked  = False
            on_hold = False
            for d in range(7):
                col = sat_col + d * 2
                if col >= len(df.columns):
                    break
                # Bokad: namn i rumraden (rad 0)
                v = df.iloc[r_idx, col]
                if pd.notna(v) and str(v).strip() not in ('', 'nan'):
                    booked = True
                # On hold: sök i raderna +1, +2, +3 under rumraden
                for offset in [1, 2, 3]:
                    if r_idx + offset < len(df):
                        h = df.iloc[r_idx + offset, col]
                        if pd.notna(h):
                            h_str = str(h).strip().lower()
                            if h_str not in ('', 'nan') and any(
                                k in h_str for k in ('hold', 'option', 'prel')
                            ):
                                on_hold = True
            statuses.append(
                "booked" if booked else "on_hold" if on_hold else "available"
            )
        return statuses

    rooms = []
    added = set()

    for row, name in ROW_TO_NAME.items():
        if name in added:
            continue
        r_idx = row - 1
        if r_idx >= len(df):
            print(f"   ⚠️ Rad {row} ({name}) finns inte")
            continue

        if name in COMBINED_ROOMS:
            # Slå ihop rader – bokad om NÅGON är bokad
            all_st = [get_week_statuses(r - 1) for r in COMBINED_ROOMS[name]]
            week_status = []
            for i in range(len(saturdays)):
                col_st = [s[i] for s in all_st]
                week_status.append(
                    "booked"    if "booked"   in col_st else
                    "on_hold"   if "on_hold"  in col_st else
                    "available"
                )
        else:
            week_status = get_week_statuses(r_idx)

        rooms.append({"name": name, "weeks": week_status})
        added.add(name)

    print(f"   {len(rooms)} rum klara")
    return {"weeks": saturdays, "rooms": rooms}

# ─── GENERATE HTML ────────────────────────────────────────────────────
def generate_html(data):
    weeks = data["weeks"]
    rooms = data["rooms"]
    now   = datetime.now().strftime("%d %b %Y, %H:%M")

    data_js  = f"const WEEKS = {json.dumps(weeks, ensure_ascii=False)};\n"
    data_js += f"const ROOMS = {json.dumps(rooms, ensure_ascii=False)};\n"
    data_js += f"const ROOM_INFO = {json.dumps(ROOM_INFO, ensure_ascii=False)};\n"
    data_js += f"const UPDATED = '{now}';\n"

    return f'''<!DOCTYPE html>
<html lang="sv">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Wildwind – Rumstillgänglighet 2026</title>
<link href="https://fonts.googleapis.com/css2?family=Nunito+Sans:wght@300;400;600;700&family=Cormorant+Garamond:ital,wght@0,400;0,600;1,400&display=swap" rel="stylesheet"/>
<style>
:root{{
  --azure:#0B3D72;--azure2:#1A5C9A;--sky:#5BA4CF;--sand:#F5EDD8;--sand2:#E6D5BA;
  --green:#2E8B57;--orange:#E08C2A;--red:#C0392B;--gold:#C9963A;
  --green-bg:#E8F5EE;--orange-bg:#FEF3E2;--red-bg:#FDECEC;
  --white:#FDFAF5;--text:#2A2A2A;--mid:#666;
}}
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'Nunito Sans',sans-serif;background:var(--white);color:var(--text);font-size:14px;}}

/* ── HEADER ── */
header{{background:var(--azure);color:#fff;padding:14px 20px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;}}
header h1{{font-family:'Cormorant Garamond',serif;font-size:22px;font-weight:400;}}
.header-right{{display:flex;align-items:center;gap:12px;}}
.updated{{font-size:11px;opacity:0.6;}}
.reload-btn{{background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.3);color:#fff;padding:6px 14px;border-radius:20px;cursor:pointer;font-size:12px;font-family:inherit;font-weight:600;}}
.reload-btn:hover{{background:rgba(255,255,255,0.28);}}

/* ── CONTROLS ── */
.controls{{background:var(--sand);padding:14px 20px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;border-bottom:2px solid var(--sand2);}}
.controls label{{font-size:12px;font-weight:700;color:var(--azure);text-transform:uppercase;letter-spacing:0.06em;}}
.controls select{{padding:7px 12px;border:1.5px solid var(--sand2);border-radius:20px;font-family:inherit;font-size:13px;background:#fff;color:var(--text);cursor:pointer;}}
.controls select:focus{{outline:none;border-color:var(--azure);}}
.prog-btns{{display:flex;gap:6px;flex-wrap:wrap;}}
.prog-btn{{padding:7px 14px;border-radius:20px;border:1.5px solid var(--sand2);background:#fff;font-family:inherit;font-size:12px;font-weight:600;cursor:pointer;transition:all .2s;color:var(--mid);}}
.prog-btn.active{{background:var(--azure);color:#fff;border-color:var(--azure);}}

/* ── VECKONAVIGERING ── */
.week-nav{{background:var(--azure2);color:#fff;padding:0 20px;display:flex;align-items:stretch;}}
.week-nav-inner{{display:flex;align-items:center;gap:0;width:100%;overflow-x:auto;scrollbar-width:none;}}
.week-nav-inner::-webkit-scrollbar{{display:none;}}
.week-tab{{
  padding:14px 16px;cursor:pointer;white-space:nowrap;
  font-size:12px;font-weight:600;color:rgba(255,255,255,0.65);
  border-bottom:3px solid transparent;transition:all .2s;flex-shrink:0;
  display:flex;flex-direction:column;align-items:center;gap:3px;
}}
.week-tab:hover{{color:#fff;background:rgba(255,255,255,0.08);}}
.week-tab.active{{color:#fff;border-bottom-color:var(--gold);}}
.week-tab .wt-date{{font-size:13px;}}
.week-tab .wt-avail{{font-size:10px;opacity:0.7;}}
.week-tab.active .wt-avail{{opacity:1;color:var(--gold);}}
.nav-arrow-btn{{
  background:transparent;border:none;color:rgba(255,255,255,0.7);
  font-size:22px;cursor:pointer;padding:0 10px;flex-shrink:0;
  align-self:center;transition:color .2s;
}}
.nav-arrow-btn:hover{{color:#fff;}}
.nav-arrow-btn:disabled{{opacity:0.2;cursor:default;}}

/* ── PRIS-BANNER ── */
.price-banner{{
  background:#fff;border-bottom:1px solid #e8e0d0;
  padding:12px 20px;display:flex;align-items:center;gap:24px;flex-wrap:wrap;
}}
.price-banner-date{{
  font-family:'Cormorant Garamond',serif;font-size:20px;color:var(--azure);font-weight:600;
}}
.price-chips{{display:flex;gap:8px;flex-wrap:wrap;}}
.price-chip{{
  padding:5px 14px;border-radius:20px;font-size:12px;font-weight:600;
  background:var(--sand);color:var(--azure);border:1px solid var(--sand2);
}}
.price-chip.highlight{{background:var(--azure);color:#fff;border-color:var(--azure);}}
.price-chip.dim{{opacity:0.45;}}

/* ── RUMLISTA ── */
.room-grid{{
  display:grid;
  grid-template-columns:repeat(auto-fill, minmax(280px, 1fr));
  gap:14px;padding:20px;
}}
.hotel-section{{grid-column:1/-1;}}
.hotel-label{{
  font-size:11px;font-weight:700;letter-spacing:0.14em;
  text-transform:uppercase;color:var(--mid);
  padding:4px 0 8px;border-bottom:2px solid var(--sand2);
  margin-bottom:4px;
}}
.room-card{{
  background:#fff;border-radius:10px;padding:16px;
  border:1.5px solid #e8e0d0;transition:all .2s;
  display:flex;flex-direction:column;gap:8px;
}}
.room-card:hover{{border-color:var(--sky);box-shadow:0 4px 16px rgba(11,61,114,0.08);}}
.room-card.booked{{opacity:0.42;border-color:#eee;}}
.room-card.on_hold{{border-color:var(--orange);background:var(--orange-bg);}}
.room-card-top{{display:flex;align-items:center;gap:10px;}}
.room-dot{{width:10px;height:10px;border-radius:50%;flex-shrink:0;}}
.room-dot.available{{background:var(--green);}}
.room-dot.on_hold{{background:var(--orange);}}
.room-dot.booked{{background:var(--red);}}
.room-card-name{{font-weight:700;font-size:15px;color:var(--azure);}}
.room-card-status{{font-size:11px;font-weight:700;margin-left:auto;}}
.status-available{{color:var(--green);}}
.status-on_hold{{color:var(--orange);}}
.status-booked{{color:var(--red);}}
.room-card-info{{font-size:12px;color:var(--mid);line-height:1.5;}}
.room-card-prices{{
  background:var(--sand);border-radius:6px;padding:8px 10px;
  display:flex;flex-direction:column;gap:3px;margin-top:2px;
}}
.room-price-row{{display:flex;justify-content:space-between;font-size:12px;}}
.room-price-label{{color:var(--mid);}}
.room-price-val{{font-weight:700;color:var(--azure);}}
.room-price-val.selected{{color:var(--green);font-size:13px;}}
.room-card-cta{{
  display:flex;align-items:center;justify-content:center;gap:6px;
  background:var(--azure);color:#fff;border-radius:6px;
  padding:9px;font-size:12px;font-weight:600;
  text-decoration:none;margin-top:2px;transition:background .2s;
}}
.room-card-cta:hover{{background:var(--azure2);}}
.empty-state{{
  grid-column:1/-1;text-align:center;
  padding:60px 20px;color:var(--mid);
}}
.empty-state h3{{font-family:'Cormorant Garamond',serif;font-size:24px;color:var(--azure);margin-bottom:8px;}}

/* ── LEGEND ── */
.legend{{display:flex;gap:16px;padding:8px 20px;background:#fff;border-bottom:1px solid #e8e0d0;flex-wrap:wrap;}}
.legend-item{{display:flex;align-items:center;gap:6px;font-size:11px;font-weight:600;}}
.dot{{width:10px;height:10px;border-radius:50%;}}
.dot-green{{background:var(--green);}} .dot-orange{{background:var(--orange);}} .dot-red{{background:var(--red);}}

.info-box{{margin:0 20px 20px;padding:12px 16px;background:#EDF3FA;border-left:4px solid var(--sky);border-radius:6px;font-size:12px;color:var(--mid);line-height:1.6;}}

@media(max-width:600px){{
  .room-grid{{grid-template-columns:1fr;padding:12px;gap:10px;}}
  .controls{{padding:12px;gap:8px;}}
  .price-banner{{padding:10px 12px;gap:12px;}}
  .price-banner-date{{font-size:17px;}}
}}
</style>
</head>
<body>

<header>
  <h1>🌊 Wildwind 2026</h1>
  <div class="header-right">
    <span class="updated">Uppdaterad: <span id="upd"></span></span>
    <button class="reload-btn" onclick="location.reload(true)">↺ Uppdatera</button>
  </div>
</header>

<div class="controls">
  <label>Program:</label>
  <div class="prog-btns">
    <button class="prog-btn active" onclick="setProgram('sailing',this)">Sailing</button>
    <button class="prog-btn" onclick="setProgram('faw',this)">FAW</button>
    <button class="prog-btn" onclick="setProgram('ho',this)">Healthy Options</button>
  </div>
  <label style="margin-left:8px;">Varaktighet:</label>
  <div class="prog-btns">
    <button class="prog-btn active" onclick="setDuration(1,this)">1 vecka</button>
    <button class="prog-btn" onclick="setDuration(2,this)">2 veckor</button>
  </div>
  <label style="margin-left:8px;">Visa:</label>
  <select id="showFilter" onchange="render()">
    <option value="available">Bara lediga</option>
    <option value="all">Alla rum</option>
  </select>
</div>

<div class="week-nav">
  <button class="nav-arrow-btn" id="navPrev" onclick="shiftWeeks(-1)">&#8249;</button>
  <div class="week-nav-inner" id="weekTabs"></div>
  <button class="nav-arrow-btn" id="navNext" onclick="shiftWeeks(1)">&#8250;</button>
</div>

<div class="price-banner" id="priceBanner"></div>

<div class="legend">
  <div class="legend-item"><div class="dot dot-green"></div>Ledig</div>
  <div class="legend-item"><div class="dot dot-orange"></div>Preliminärt bokad</div>
  <div class="legend-item"><div class="dot dot-red"></div>Bokad</div>
</div>

<div class="room-grid" id="roomGrid"></div>

<div class="info-box">ℹ️ Rum kan vara preliminärt bokade i upp till 48h utan att synas här. Kontakta oss för aktuell tillgänglighet.</div>

<script>
{data_js}

document.getElementById('upd').textContent = UPDATED;

const SECTIONS = [
  {{ label:"Melas Hotel",         rooms:["Melas 1","Melas 2AB","Melas 3","Melas 4","Melas 7AB","Melas 8"] }},
  {{ label:"Kavadias Hotel",      rooms:["Kav 2","Kav 3","Kav 4","Kav 13","Kav 14","Kav 15","Kav 18","Kav 19"] }},
  {{ label:"AKTI Apartments",     rooms:["Akti 4","Akti 5","Akti 7","Akti 8","Akti 9"] }},
  {{ label:"Xristina Apartments", rooms:["Xristina 6"] }},
  {{ label:"New Melas",           rooms:["NM208"] }},
];

let currentWeek = 0;
let program = 'sailing';
let duration = 1;
let tabOffset = 0;
const TABS_VISIBLE = window.innerWidth < 600 ? 4 : 8;

// Starta på närmaste kommande lördag
const today = new Date();
for (let i = 0; i < WEEKS.length; i++) {{
  if (new Date(WEEKS[i].date) >= today) {{ currentWeek = i; break; }}
}}
tabOffset = Math.max(0, currentWeek - 2);

function getPrice(w) {{
  if (!w) return null;
  if (program==='sailing') return duration===1 ? w.price  : w.sail2;
  if (program==='faw')     return duration===1 ? w.fatw1  : w.fatw2;
  if (program==='ho')      return duration===1 ? w.ho1    : w.ho2;
  return null;
}}

function progLabel() {{
  if (program==='sailing') return 'Sailing';
  if (program==='faw')     return 'FAW';
  return 'Healthy Options';
}}

function setProgram(p, btn) {{
  program = p;
  document.querySelectorAll('.prog-btns .prog-btn').forEach(b => b.classList.remove('active'));
  // Mark only program buttons
  btn.classList.add('active');
  render();
}}

function setDuration(d, btn) {{
  duration = d;
  // Mark only duration buttons
  const allBtns = document.querySelectorAll('.prog-btns .prog-btn');
  allBtns.forEach(b => {{ if (b.textContent.includes('vecka')) b.classList.remove('active'); }});
  btn.classList.add('active');
  render();
}}

function shiftWeeks(dir) {{
  tabOffset = Math.max(0, Math.min(WEEKS.length - TABS_VISIBLE, tabOffset + dir));
  renderTabs();
}}

function countAvailable(wi) {{
  return ROOMS.filter(r => r.weeks[wi] === 'available').length;
}}

function renderTabs() {{
  const c = document.getElementById('weekTabs');
  c.innerHTML = '';
  const end = Math.min(WEEKS.length, tabOffset + TABS_VISIBLE);
  for (let i = tabOffset; i < end; i++) {{
    const w = WEEKS[i];
    const avail = countAvailable(i);
    const d = document.createElement('div');
    d.className = 'week-tab' + (i===currentWeek?' active':'');
    d.innerHTML = `<span class="wt-date">${{w.display}}</span><span class="wt-avail">${{avail}} lediga</span>`;
    d.onclick = () => {{ currentWeek=i; renderTabs(); renderBanner(); renderRooms(); }};
    c.appendChild(d);
  }}
  document.getElementById('navPrev').disabled = tabOffset === 0;
  document.getElementById('navNext').disabled = tabOffset + TABS_VISIBLE >= WEEKS.length;
}}

function renderBanner() {{
  const w = WEEKS[currentWeek];
  const endDate = new Date(w.date); endDate.setDate(endDate.getDate() + (duration===2?14:7));
  const endLabel = endDate.toLocaleDateString('sv-SE', {{day:'numeric', month:'short'}});
  const basePrice = getPrice(w);
  const b = document.getElementById('priceBanner');

  const chips = [
    {{label:'Sailing', val: duration===1?w.price:w.sail2, prog:'sailing'}},
    {{label:'FAW',     val: duration===1?w.fatw1:w.fatw2, prog:'faw'}},
    {{label:'HO',      val: duration===1?w.ho1:w.ho2,     prog:'ho'}},
  ].filter(c => c.val);

  let html = `<div class="price-banner-date">${{w.display}} – ${{endLabel}}</div>`;
  html += `<div class="price-chips">`;
  chips.forEach(c => {{
    const cls = c.prog===program ? 'price-chip highlight' : 'price-chip';
    html += `<div class="${{cls}}" onclick="setProgram('${{c.prog}}', this)" style="cursor:pointer">${{c.label}} ${{c.val.toLocaleString('sv-SE')}} kr</div>`;
  }});
  if (w.single) html += `<div class="price-chip dim">Enkelrum +${{w.single.toLocaleString('sv-SE')}} kr</div>`;
  html += `</div>`;
  b.innerHTML = html;
}}

function renderRooms() {{
  const showAll = document.getElementById('showFilter').value === 'all';
  const w = WEEKS[currentWeek];
  const basePrice = getPrice(w);
  const grid = document.getElementById('roomGrid');
  let html = '';
  let totalShown = 0;

  SECTIONS.forEach(sec => {{
    const secRooms = sec.rooms.filter(name => {{
      const room = ROOMS.find(r => r.name===name);
      if (!room) return false;
      if (!showAll && room.weeks[currentWeek] !== 'available') return false;
      return true;
    }});
    if (!secRooms.length) return;
    totalShown += secRooms.length;
    html += `<div class="hotel-section"><div class="hotel-label">${{sec.label}}</div></div>`;
    secRooms.forEach(name => {{
      const room = ROOMS.find(r => r.name===name);
      const st = room.weeks[currentWeek] || 'available';
      const info = ROOM_INFO[name] || {{}};
      const tillagg = info.tillagg_num || 0;
      const stLabel = {{available:'Ledig', on_hold:'Preliminär', booked:'Bokad'}};

      // Bygg prisrader
      let priceRows = '';
      const progs = [
        {{key:'sailing', label:'Sailing',         val1: w.price, val2: w.sail2}},
        {{key:'faw',     label:'FAW',             val1: w.fatw1, val2: w.fatw2}},
        {{key:'ho',      label:'Healthy Options', val1: w.ho1,   val2: w.ho2}},
      ];
      progs.forEach(p => {{
        const base = duration===1 ? p.val1 : p.val2;
        if (!base) return;
        const total = base + tillagg;
        const isSel = p.key === program;
        priceRows += `<div class="room-price-row">
          <span class="room-price-label">${{p.label}} ${{duration}}v + tillägg</span>
          <span class="room-price-val${{isSel?' selected':''}}">${{total.toLocaleString('sv-SE')}} kr</span>
        </div>`;
      }});
      if (info.tillagg && info.tillagg !== '0 kr') {{
        priceRows += `<div class="room-price-row" style="border-top:1px solid #ddd;margin-top:4px;padding-top:4px;">
          <span class="room-price-label">varav rumstillägg</span>
          <span class="room-price-val">${{info.tillagg}}</span>
        </div>`;
      }}

      const subject = encodeURIComponent(`Bokningsförfrågan: ${{name}} vecka ${{w.display}}`);
      html += `<div class="room-card ${{st !== 'available' ? st : ''}}">
        <div class="room-card-top">
          <div class="room-dot ${{st}}"></div>
          <div class="room-card-name">${{name}}</div>
          <div class="room-card-status status-${{st}}">${{stLabel[st]}}</div>
        </div>
        ${{info.info ? `<div class="room-card-info">${{info.info}}</div>` : ''}}
        <div class="room-card-prices">${{priceRows}}</div>
        ${{st==='available' ? `<a href="mailto:pia@wildwind.se?subject=${{subject}}" class="room-card-cta">✉ Boka ${{name}}</a>` : ''}}
      </div>`;
    }});
  }});

  if (totalShown === 0) {{
    html = `<div class="empty-state"><h3>Inga lediga rum</h3><p>Alla rum är bokade denna vecka. Välj en annan vecka!</p></div>`;
  }}
  grid.innerHTML = html;
}}

function render() {{
  renderTabs();
  renderBanner();
  renderRooms();
}}

render();
</script>
</body>
</html>'''

# ─── MAIN ─────────────────────────────────────────────────────────────
def main():
    print("🚀 Wildwind Availability Updater")
    print("="*40)
    try:
        excel = download_excel()
        data  = parse_availability(excel)
        html  = generate_html(data)
        Path(OUTPUT_FILE).write_text(html, encoding='utf-8')
        print(f"✅ Genererade {OUTPUT_FILE}")
        Path("temp_availability.xlsx").unlink(missing_ok=True)
        print("🎉 Klar!")
    except Exception as e:
        print(f"❌ Fel: {e}")
        raise

if __name__ == "__main__":
    main()
