#!/usr/bin/env python3
"""
Wildwind Room Availability Updater
Downloads Excel from Dropbox and generates availability.html (NOT index.html)
"""

import re
import pandas as pd
import json
import requests
from datetime import datetime, date
from pathlib import Path

# ─── CONFIG ───────────────────────────────────────────────────────────
DROPBOX_URL = "https://www.dropbox.com/scl/fi/i0ji499omsoc0yl4fuzo7/STEPH-VERSION-ALL-ROOMS-2026.xlsx?rlkey=lq3j8vz7hdlnoqvairj0yp2ih&dl=1"
OUTPUT_FILE = "availability.html"   # ← OBS: inte index.html

# Excel-rader som Seafari AB kan boka (0-indexerade)
ALLOWED_ROWS = [44, 48, 52, 56, 60, 140, 144, 148, 168, 184, 188, 192,
                204, 208, 220, 224, 244, 248, 256, 264, 272, 284]

# Sailing-priser 2026 (SEK/person/vecka, 2 delar rum)
SAILING_PRICES = {
    "Apr 25": 8990,  "May 2": 9430,   "May 9": 10300,  "May 16": 10840,
    "May 23": 10840, "May 30": 11940, "Jun 06": 12700, "Jun 13": 13030,
    "Jun 20": 13900, "Jun 27": 14120, "Jul 4": 14440,  "Jul 11": 14990,
    "Jul 18": 15210, "Jul 25": 15210, "Aug 01": 15210, "Aug 8": 15210,
    "Aug 15": 14120, "Aug 22": 12810, "Aug 29": 11940, "Sep 5": 11720,
    "Sep 12": 11170, "Sep 19": 10300, "Sep 26": 9540,  "Oct 3": 8990,
}

# Mappning från Excel-datum (lördag) → prisliste-nyckel
def date_to_price_key(dt):
    """Matcha ett datum mot SAILING_PRICES-nycklar."""
    if not isinstance(dt, (date, datetime)):
        return None
    month_abbr = dt.strftime("%b")   # "Apr", "May" etc
    day = dt.day
    key = f"{month_abbr} {day}"
    if key in SAILING_PRICES:
        return key
    # Prova nollpaddat (Jun 06 etc)
    key2 = f"{month_abbr} {day:02d}"
    if key2 in SAILING_PRICES:
        return key2
    return None

# ─── DOWNLOAD ─────────────────────────────────────────────────────────
def download_excel():
    print("📥 Laddar ner Excel från Dropbox...")
    response = requests.get(DROPBOX_URL, timeout=30)
    response.raise_for_status()
    tmp = Path("temp_availability.xlsx")
    tmp.write_bytes(response.content)
    print(f"✅ {len(response.content):,} bytes")
    return tmp

# ─── PARSE ────────────────────────────────────────────────────────────
import re

ROOM_PATTERN = re.compile(
    r'\b(NM\d+|MN\d+|[MK]\d+[AB/]*|[AX]\d+[A-Z]*|COSMOS[- ]?II?)\b',
    re.IGNORECASE
)

def parse_availability(excel_path):
    df = pd.read_excel(excel_path, header=None, engine="openpyxl")

    # ── Hitta alla kolumner med lördagsdatum ──
    weeks = []
    for row_idx in range(min(5, len(df))):          # kolla de 5 första raderna
        for col_idx, val in enumerate(df.iloc[row_idx]):
            if isinstance(val, (date, datetime)):
                dt = val.date() if isinstance(val, datetime) else val
                if dt.weekday() == 5:               # 5 = lördag
                    # undvik dubletter
                    if not any(w["col"] == col_idx for w in weeks):
                        weeks.append({"col": col_idx, "date": dt})
        if weeks:
            break   # hittade lördagar i denna rad, gå vidare

    # Sortera kronologiskt
    weeks.sort(key=lambda w: w["date"])

    # ── Hämta rum ──
    rooms = []
    for row_idx in ALLOWED_ROWS:
        if row_idx >= len(df):
            continue
        row = df.iloc[row_idx]

        # Leta efter rumsnummer (t.ex. M1, K14, A7, NM208) i de första 5 kolumnerna
        room_name = None
        for col_idx in range(min(5, len(row))):
            val = row.iloc[col_idx]
            if pd.isna(val):
                continue
            match = ROOM_PATTERN.search(str(val))
            if match:
                room_name = match.group(0).upper()
                break

        if not room_name:
            # Fallback: använd första icke-tomma cell
            for col_idx in range(min(5, len(row))):
                val = row.iloc[col_idx]
                if pd.notna(val) and str(val).strip() not in ("", "nan"):
                    room_name = str(val).strip()
                    break
            if not room_name:
                room_name = f"Rad {row_idx}"

        # ── Status per vecka ──
        week_status = []
        for wk in weeks:
            cell_val = row.iloc[wk["col"]] if wk["col"] < len(row) else None
            if cell_val is None or (isinstance(cell_val, float) and pd.isna(cell_val)):
                status = "available"
            else:
                val_str = str(cell_val).strip().lower()
                if val_str in ("", "nan"):
                    status = "available"
                elif any(k in val_str for k in ("hold", "option", "prel", "~")):
                    status = "on_hold"
                else:
                    status = "booked"
            week_status.append(status)

        rooms.append({"name": room_name, "weeks": week_status})

    return {"weeks": weeks, "rooms": rooms}

# ─── GENERATE HTML ────────────────────────────────────────────────────
def generate_html(data):
    weeks  = data["weeks"]
    rooms  = data["rooms"]
    now    = datetime.now().strftime("%d %b %Y, %H:%M")

    # Bygg vecko-JSON med pris
    weeks_json = []
    for wk in weeks:
        dt = wk["date"]
        label = dt.strftime("%-d %b") if hasattr(dt, 'strftime') else str(dt)
        price_key = date_to_price_key(dt)
        price = SAILING_PRICES.get(price_key) if price_key else None
        weeks_json.append({
            "label": label,
            "iso": dt.isoformat() if hasattr(dt, 'isoformat') else str(dt),
            "price": price
        })

    rooms_json = rooms  # already serializable

    data_js = f"const WEEKS = {json.dumps(weeks_json, ensure_ascii=False)};\n"
    data_js += f"const ROOMS = {json.dumps(rooms_json, ensure_ascii=False)};\n"
    data_js += f"const UPDATED = '{now}';\n"

    html = f'''<!DOCTYPE html>
<html lang="sv">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Wildwind – Rumstillgänglighet 2026</title>
<link href="https://fonts.googleapis.com/css2?family=Nunito+Sans:wght@300;400;600&family=Cormorant+Garamond:wght@400;600&display=swap" rel="stylesheet"/>
<style>
:root {{
  --azure: #0B3D72; --sky: #5BA4CF; --sand: #F5EDD8;
  --green: #2E8B57; --orange: #E08C2A; --red: #C0392B;
  --green-bg: #E8F5EE; --orange-bg: #FEF3E2; --red-bg: #FDECEC;
  --white: #FDFAF5; --text: #2A2A2A; --text-mid: #666;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: 'Nunito Sans', sans-serif; background: var(--white); color: var(--text); font-size: 14px; }}

header {{
  background: var(--azure); color: #fff;
  padding: 20px 32px; display: flex; align-items: center; justify-content: space-between;
}}
header h1 {{ font-family: 'Cormorant Garamond', serif; font-size: 24px; font-weight: 400; }}
header .updated {{ font-size: 12px; opacity: 0.6; }}

.legend {{
  display: flex; gap: 20px; padding: 14px 32px;
  background: #fff; border-bottom: 1px solid #e8e0d0;
  flex-wrap: wrap; align-items: center;
}}
.legend-item {{ display: flex; align-items: center; gap: 7px; font-size: 12px; font-weight: 600; }}
.dot {{ width: 14px; height: 14px; border-radius: 3px; flex-shrink: 0; }}
.dot-green  {{ background: var(--green); }}
.dot-orange {{ background: var(--orange); }}
.dot-red    {{ background: var(--red); }}
.note {{ font-size: 11px; color: var(--text-mid); margin-left: auto; }}

.filter-bar {{
  padding: 12px 32px; background: var(--sand);
  display: flex; gap: 12px; align-items: center; flex-wrap: wrap;
}}
.filter-bar label {{ font-size: 12px; font-weight: 600; color: var(--azure); }}
.filter-bar select, .filter-bar input {{
  padding: 6px 10px; border: 1px solid #ccc; border-radius: 20px;
  font-family: inherit; font-size: 12px; background: #fff;
}}

.table-wrap {{ overflow-x: auto; padding: 24px 32px; }}
table {{ border-collapse: collapse; min-width: 100%; }}

th {{
  background: var(--azure); color: #fff;
  padding: 0; text-align: center; font-size: 11px;
  font-weight: 600; white-space: nowrap; position: sticky; top: 0; z-index: 2;
}}
th .week-date {{ padding: 8px 10px 2px; display: block; }}
th .week-price {{
  display: block; padding: 2px 10px 8px;
  font-size: 10px; font-weight: 400; opacity: 0.75;
  color: #FFE099;
}}
th:first-child {{ text-align: left; min-width: 130px; padding: 8px 14px; position: sticky; left: 0; z-index: 3; }}

td {{ padding: 0; border: 1px solid #e8e0d0; }}
td:first-child {{
  padding: 8px 14px; font-weight: 600; font-size: 12px;
  background: #fff; position: sticky; left: 0; z-index: 1;
  white-space: nowrap; border-right: 2px solid #d0c8b8;
}}
tr:nth-child(even) td:first-child {{ background: #faf6ee; }}

.cell {{
  width: 52px; height: 36px; display: flex;
  align-items: center; justify-content: center;
  font-size: 10px; font-weight: 700; cursor: default;
  transition: opacity .15s;
}}
.cell:hover {{ opacity: 0.8; }}
.cell-available {{ background: var(--green-bg); color: var(--green); }}
.cell-on_hold   {{ background: var(--orange-bg); color: var(--orange); }}
.cell-booked    {{ background: var(--red-bg); color: var(--red); }}

.info-box {{
  margin: 0 32px 32px; padding: 14px 20px;
  background: #EDF3FA; border-left: 4px solid var(--sky);
  border-radius: 6px; font-size: 12px; color: var(--text-mid); line-height: 1.6;
}}
.contact-btn {{
  display: inline-flex; align-items: center; gap: 8px;
  margin: 0 32px 32px; padding: 14px 28px;
  background: var(--azure); color: #fff; border-radius: 40px;
  text-decoration: none; font-size: 13px; font-weight: 600;
  transition: background .2s;
}}
.contact-btn:hover {{ background: #1A5C9A; }}
</style>
</head>
<body>
<header>
  <h1>Wildwind – Rumstillgänglighet 2026</h1>
  <div style="display:flex;align-items:center;gap:16px;">
    <span class="updated">Uppdaterad: <span id="upd"></span></span>
    <button onclick="location.reload(true)" style="background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.3);color:#fff;padding:7px 16px;border-radius:20px;cursor:pointer;font-size:12px;font-family:inherit;font-weight:600;transition:background .2s;" onmouseover="this.style.background='rgba(255,255,255,0.25)'" onmouseout="this.style.background='rgba(255,255,255,0.15)'">↺ Uppdatera</button>
  </div>
</header>

<div class="legend">
  <div class="legend-item"><div class="dot dot-green"></div> Ledig</div>
  <div class="legend-item"><div class="dot dot-orange"></div> Preliminärt bokad</div>
  <div class="legend-item"><div class="dot dot-red"></div> Bokad</div>
  <span class="note">Alla veckor lördag–lördag &nbsp;·&nbsp; Sailing-pris per person/v (2 delar rum)</span>
</div>

<div class="filter-bar">
  <label>Filtrera:</label>
  <input type="text" id="roomFilter" placeholder="Rum (t.ex. M3)" oninput="renderTable()"/>
  <select id="statusFilter" onchange="renderTable()">
    <option value="all">Alla statusar</option>
    <option value="available">Visa bara lediga</option>
  </select>
  <select id="weekFilter" onchange="renderTable()">
    <option value="all">Alla veckor</option>
  </select>
</div>

<div class="table-wrap">
  <table id="avail-table"></table>
</div>

<div class="info-box">
  ℹ️ Vissa rum kan vara preliminärt bokade i upp till 48 timmar utan att det syns här.
  Kontakta oss för att säkerställa aktuell tillgänglighet innan du bekräftar resan till dina kunder.
</div>

<a href="mailto:pia@wildwind.se?subject=Wildwind%20bokningsf%C3%B6rfr%C3%A5gan" class="contact-btn">
  ✉️ Skicka bokningsförfrågan
</a>

<script>
{data_js}

document.getElementById('upd').textContent = UPDATED;

// Populate week filter
const wf = document.getElementById('weekFilter');
WEEKS.forEach((w,i) => {{
  const opt = document.createElement('option');
  opt.value = i; opt.textContent = w.label + (w.price ? ' – ' + w.price.toLocaleString('sv-SE') + ' kr' : '');
  wf.appendChild(opt);
}});

function renderTable() {{
  const roomQ  = document.getElementById('roomFilter').value.toLowerCase();
  const statQ  = document.getElementById('statusFilter').value;
  const weekQ  = document.getElementById('weekFilter').value;

  const filtWeeks = weekQ === 'all'
    ? WEEKS.map((_,i) => i)
    : [parseInt(weekQ)];

  const table = document.getElementById('avail-table');
  let html = '<thead><tr><th>Rum</th>';
  filtWeeks.forEach(i => {{
    const w = WEEKS[i];
    html += `<th><span class="week-date">${{w.label}}</span>`;
    html += w.price ? `<span class="week-price">${{w.price.toLocaleString('sv-SE')}} kr</span>` : '<span class="week-price">&nbsp;</span>';
    html += '</th>';
  }});
  html += '</tr></thead><tbody>';

  ROOMS.forEach(room => {{
    if (roomQ && !room.name.toLowerCase().includes(roomQ)) return;
    if (statQ === 'available') {{
      const anyFree = filtWeeks.some(i => room.weeks[i] === 'available');
      if (!anyFree) return;
    }}
    html += `<tr><td>${{room.name}}</td>`;
    filtWeeks.forEach(i => {{
      const st = room.weeks[i] || 'available';
      const label = st === 'available' ? '✓' : st === 'on_hold' ? '~' : '✕';
      html += `<td><div class="cell cell-${{st}}" title="${{WEEKS[i].label}} – ${{st === 'available' ? 'Ledig' : st === 'on_hold' ? 'Prel. bokad' : 'Bokad'}}">${{label}}</div></td>`;
    }});
    html += '</tr>';
  }});
  html += '</tbody>';
  table.innerHTML = html;
}}

renderTable();
</script>
</body>
</html>'''
    return html

# ─── MAIN ─────────────────────────────────────────────────────────────
def main():
    print("🚀 Wildwind Availability Updater")
    print("=" * 40)
    try:
        excel_file = download_excel()
        data = parse_availability(excel_file)
        html = generate_html(data)
        Path(OUTPUT_FILE).write_text(html, encoding='utf-8')
        print(f"✅ Genererade {OUTPUT_FILE}")
        Path("temp_availability.xlsx").unlink(missing_ok=True)
        print("🎉 Klar!")
    except Exception as e:
        print(f"❌ Fel: {e}")
        raise

if __name__ == "__main__":
    main()
