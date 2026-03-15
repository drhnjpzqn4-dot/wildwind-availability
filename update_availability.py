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

# Exakt mappning: Excel-rad (1-indexerad) → visningsnamn
# Ordningen styr visningsordningen på sidan
ROW_TO_NAME = {
    44:  "M1",
    48:  "M2A/B",   # M2A och M2B kombineras – kolla båda raderna (48+52)
    56:  "M3",
    60:  "M4",
    72:  "M7A/B",   # M7A och M7B kombineras (72+76)
    80:  "M8",
    140: "K2",
    144: "K3",
    148: "K4",
    188: "K14",
    192: "K15",
    204: "K18",
    208: "K19",
    216: "A7",
    220: "A8",
    224: "A9",
    244: "A4",
    248: "A5",
    272: "X6",
    284: "NM208",
}

# Kombinerade rum: dessa rader slås ihop till ett rum
# Om NÅGON dag i NÅGOT av raderna är bokad → bokad
COMBINED_ROOMS = {
    "M2A/B": [48, 52],   # Båda raderna måste vara lediga
    "M7A/B": [72, 76],
}

# Sailing-priser 2026
SAILING_PRICES = {
    "25 Apr": 8990,  "2 May": 9430,   "9 May": 10300,  "16 May": 10840,
    "23 May": 10840, "30 May": 11940, "6 Jun": 12700,  "13 Jun": 13030,
    "20 Jun": 13900, "27 Jun": 14120, "4 Jul": 14440,  "11 Jul": 14990,
    "18 Jul": 15210, "25 Jul": 15210, "1 Aug": 15210,  "8 Aug": 15210,
    "15 Aug": 14120, "22 Aug": 12810, "29 Aug": 11940, "5 Sep": 11720,
    "12 Sep": 11170, "19 Sep": 10300, "26 Sep": 9540,  "3 Oct": 8990,
}

# Ruminfo för mobilvy
ROOM_INFO = {
    "M1":    {"tillagg": "0 kr",     "info": "Twin · Sidoutsikt hav + Kav Bar"},
    "M2A/B": {"tillagg": "0 kr",     "info": "Twin+Double · Familjerum, sid+havsutsikt"},
    "M3":    {"tillagg": "1 758 kr", "info": "Double · Full havsutsikt"},
    "M4":    {"tillagg": "1 758 kr", "info": "Twin · Full havsutsikt"},
    "M7A/B": {"tillagg": "0 kr",     "info": "Double+Twin · Familjerum, mot Ponti"},
    "M8":    {"tillagg": "0 kr",     "info": "Twin · Mot Ponti + berg västerut"},
    "K2":    {"tillagg": "0 kr",     "info": "Small Double · Ingen utsikt, baksida"},
    "K3":    {"tillagg": "0 kr",     "info": "Twin · Ingen utsikt, baksida (litet)"},
    "K4":    {"tillagg": "0 kr",     "info": "Double · Sidoutsikt mot trädgård"},
    "K14":   {"tillagg": "1 320 kr", "info": "Double · Pool + berg västerut, övervåning"},
    "K15":   {"tillagg": "1 320 kr", "info": "Double · Pool + berg, delar balkong K14"},
    "K18":   {"tillagg": "1 320 kr", "info": "Twin · Havsutsikt över Kav Bar, övervåning"},
    "K19":   {"tillagg": "1 320 kr", "info": "Twin · Sidoutsikt mot hav + byn"},
    "A4":    {"tillagg": "1 758 kr", "info": "Studio · Twin, pentry, övervåning"},
    "A5":    {"tillagg": "1 758 kr", "info": "Studio · Twin, pentry, övervåning"},
    "A7":    {"tillagg": "3 516 kr", "info": "2-sovrum · Double+Twin+extra, fullt kök (min 4 pers)"},
    "A8":    {"tillagg": "3 516 kr", "info": "2-sovrum · Double+Twin+extra, fullt kök (min 4 pers)"},
    "A9":    {"tillagg": "1 758 kr", "info": "1-sovrum · Twin + extra i hall, pentry"},
    "X6":    {"tillagg": "3 516 kr", "info": "1-sovrum · Double+2 enklar, kitchenette (13 jul–15 sep)"},
    "NM208": {"tillagg": "2 638 kr", "info": "Double · Bergutsikt bakåt, övervåning"},
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

    # Hitta lördagskolumner (rad 2 = index 1 innehåller "SATURDAY")
    saturdays = []
    for col in range(3, len(df.columns), 2):
        date_val = df.iloc[0, col]
        day_name = df.iloc[1, col]
        if pd.notna(day_name) and 'SATURDAY' in str(day_name).upper():
            if pd.notna(date_val):
                dt = pd.to_datetime(date_val)
                label = dt.strftime('%-d %b')
                price = SAILING_PRICES.get(label)
                saturdays.append({
                    'col':     col,
                    'date':    dt.strftime('%Y-%m-%d'),
                    'display': label,
                    'price':   price,
                })

    print(f"   {len(saturdays)} lördagar hittade")

    def week_status_for_row(r_idx, saturdays, df):
        """Räkna ut status per vecka för en rad."""
        result = []
        for sat in saturdays:
            sat_col = sat['col']
            booked  = False
            on_hold = False
            for d in range(7):
                col = sat_col + d * 2
                if col >= len(df.columns):
                    break
                # Bokad: namn i rumraden
                v = df.iloc[r_idx, col]
                if pd.notna(v) and str(v).strip() not in ('', 'nan'):
                    booked = True
                # On hold: "on hold" i raden +2 under
                if r_idx + 2 < len(df):
                    h = df.iloc[r_idx + 2, col]
                    if pd.notna(h) and 'hold' in str(h).lower():
                        on_hold = True
            result.append("booked" if booked else "on_hold" if on_hold else "available")
        return result

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
            # Slå ihop: bokad om NÅGON av raderna är bokad
            all_statuses = []
            for sub_row in COMBINED_ROOMS[name]:
                all_statuses.append(week_status_for_row(sub_row - 1, saturdays, df))
            week_status = []
            for i in range(len(saturdays)):
                combined = [s[i] for s in all_statuses]
                if "booked" in combined:
                    week_status.append("booked")
                elif "on_hold" in combined:
                    week_status.append("on_hold")
                else:
                    week_status.append("available")
        else:
            week_status = week_status_for_row(r_idx, saturdays, df)

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
<link href="https://fonts.googleapis.com/css2?family=Nunito+Sans:wght@300;400;600;700&family=Cormorant+Garamond:wght@400;600&display=swap" rel="stylesheet"/>
<style>
:root{{
  --azure:#0B3D72;--azure2:#1A5C9A;--sky:#5BA4CF;--sand:#F5EDD8;
  --green:#2E8B57;--orange:#E08C2A;--red:#C0392B;
  --green-bg:#E8F5EE;--orange-bg:#FEF3E2;--red-bg:#FDECEC;
  --white:#FDFAF5;--text:#2A2A2A;--mid:#666;
}}
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'Nunito Sans',sans-serif;background:var(--white);color:var(--text);font-size:14px;}}

header{{background:var(--azure);color:#fff;padding:14px 20px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;}}
header h1{{font-family:'Cormorant Garamond',serif;font-size:21px;font-weight:400;}}
.header-right{{display:flex;align-items:center;gap:12px;}}
.updated{{font-size:11px;opacity:0.6;}}
.reload-btn{{background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.3);color:#fff;padding:6px 14px;border-radius:20px;cursor:pointer;font-size:12px;font-family:inherit;font-weight:600;}}
.reload-btn:hover{{background:rgba(255,255,255,0.28);}}

.legend{{display:flex;gap:16px;padding:10px 20px;background:#fff;border-bottom:1px solid #e8e0d0;flex-wrap:wrap;align-items:center;}}
.legend-item{{display:flex;align-items:center;gap:6px;font-size:12px;font-weight:600;}}
.dot{{width:12px;height:12px;border-radius:3px;}}
.dot-green{{background:var(--green);}} .dot-orange{{background:var(--orange);}} .dot-red{{background:var(--red);}}
.note{{font-size:11px;color:var(--mid);margin-left:auto;}}

/* DESKTOP */
.filter-bar{{padding:10px 20px;background:var(--sand);display:flex;gap:12px;align-items:center;flex-wrap:wrap;}}
.filter-bar label{{font-size:12px;font-weight:600;color:var(--azure);}}
.filter-bar select,.filter-bar input{{padding:5px 10px;border:1px solid #ccc;border-radius:20px;font-family:inherit;font-size:12px;background:#fff;}}
.table-wrap{{overflow-x:auto;padding:20px;}}
table{{border-collapse:collapse;min-width:100%;}}
th{{background:var(--azure);color:#fff;padding:0;text-align:center;font-size:11px;font-weight:600;white-space:nowrap;position:sticky;top:0;z-index:2;}}
th .wk-date{{padding:7px 8px 2px;display:block;}}
th .wk-price{{display:block;padding:2px 8px 7px;font-size:10px;font-weight:400;opacity:0.72;color:#FFE099;}}
th:first-child{{text-align:left;min-width:90px;padding:8px 14px;position:sticky;left:0;z-index:3;}}
td{{padding:0;border:1px solid #e8e0d0;}}
td:first-child{{padding:8px 14px;font-weight:700;font-size:12px;background:#fff;position:sticky;left:0;z-index:1;white-space:nowrap;border-right:2px solid #d0c8b8;cursor:help;}}
tr:nth-child(even) td:first-child{{background:#faf6ee;}}
.cell{{width:46px;height:32px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;}}
.cell:hover{{opacity:0.75;}}
.cell-available{{background:var(--green-bg);color:var(--green);}}
.cell-on_hold{{background:var(--orange-bg);color:var(--orange);}}
.cell-booked{{background:var(--red-bg);color:var(--red);}}

/* SECTION HEADERS i tabellen */
.section-row td{{background:var(--azure)!important;color:#fff;font-size:11px;font-weight:600;letter-spacing:0.1em;text-transform:uppercase;padding:5px 14px!important;border:none!important;}}

/* MOBIL */
.mobile-view{{display:none;}}
.week-nav{{background:var(--azure);color:#fff;padding:12px 16px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10;}}
.week-nav-center{{text-align:center;flex:1;}}
.week-nav-date{{font-family:'Cormorant Garamond',serif;font-size:20px;font-weight:600;}}
.week-nav-price{{font-size:12px;color:#FFE099;margin-top:2px;}}
.nav-arrow{{background:rgba(255,255,255,0.15);border:none;color:#fff;width:40px;height:40px;border-radius:50%;font-size:22px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:background .2s;flex-shrink:0;}}
.nav-arrow:hover{{background:rgba(255,255,255,0.3);}}
.nav-arrow:disabled{{opacity:0.2;cursor:default;}}
.week-dots{{display:flex;gap:4px;justify-content:center;padding:8px 16px;background:var(--azure2);flex-wrap:wrap;}}
.wdot{{width:6px;height:6px;border-radius:50%;background:rgba(255,255,255,0.3);cursor:pointer;transition:all .2s;}}
.wdot.active{{background:#fff;width:8px;height:8px;}}
.mobile-filter{{padding:10px 16px;background:var(--sand);display:flex;gap:10px;align-items:center;flex-wrap:wrap;}}
.mobile-filter select,.mobile-filter input{{padding:7px 12px;border:1px solid #ccc;border-radius:20px;font-family:inherit;font-size:13px;background:#fff;flex:1;min-width:120px;}}
.section-header{{padding:8px 16px;background:var(--azure);color:#fff;font-size:11px;font-weight:700;letter-spacing:0.12em;text-transform:uppercase;}}
.room-row{{display:flex;align-items:center;padding:12px 16px;border-bottom:1px solid #ede8df;gap:14px;}}
.room-badge{{width:56px;height:56px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:12px;flex-shrink:0;text-align:center;line-height:1.2;}}
.badge-available{{background:var(--green-bg);color:var(--green);border:1.5px solid var(--green);}}
.badge-on_hold{{background:var(--orange-bg);color:var(--orange);border:1.5px solid var(--orange);}}
.badge-booked{{background:var(--red-bg);color:var(--red);border:1.5px solid var(--red);}}
.room-details{{flex:1;min-width:0;}}
.room-name{{font-weight:700;font-size:15px;color:var(--azure);}}
.room-desc{{font-size:12px;color:var(--mid);margin-top:2px;line-height:1.4;}}
.room-tillagg{{font-size:12px;font-weight:600;color:var(--text);margin-top:2px;}}
.room-status{{font-size:12px;font-weight:700;flex-shrink:0;text-align:right;}}
.status-available{{color:var(--green);}} .status-on_hold{{color:var(--orange);}} .status-booked{{color:var(--red);}}
.info-box{{margin:16px;padding:12px 16px;background:#EDF3FA;border-left:4px solid var(--sky);border-radius:6px;font-size:12px;color:var(--mid);line-height:1.6;}}
.contact-btn{{display:flex;align-items:center;justify-content:center;gap:8px;margin:0 16px 32px;padding:15px;background:var(--azure);color:#fff;border-radius:40px;text-decoration:none;font-size:14px;font-weight:600;}}
.contact-btn:hover{{background:var(--azure2);}}

@media (max-width:700px){{
  .desktop-view{{display:none;}}
  .mobile-view{{display:block;}}
  .legend .note{{display:none;}}
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

<div class="legend">
  <div class="legend-item"><div class="dot dot-green"></div>Ledig</div>
  <div class="legend-item"><div class="dot dot-orange"></div>Preliminärt</div>
  <div class="legend-item"><div class="dot dot-red"></div>Bokad</div>
  <span class="note">Lördag–lördag · Sailing fr. pris/person/v (2 delar rum)</span>
</div>

<!-- DESKTOP -->
<div class="desktop-view">
  <div class="filter-bar">
    <label>Filtrera:</label>
    <input type="text" id="roomFilter" placeholder="Rum (t.ex. M3)" oninput="renderTable()"/>
    <select id="statusFilter" onchange="renderTable()">
      <option value="all">Alla statusar</option>
      <option value="available">Bara lediga</option>
    </select>
    <select id="weekFilter" onchange="renderTable()">
      <option value="all">Alla veckor</option>
    </select>
  </div>
  <div class="table-wrap"><table id="tbl"></table></div>
</div>

<!-- MOBIL -->
<div class="mobile-view">
  <div class="week-nav">
    <button class="nav-arrow" id="prevBtn" onclick="changeWeek(-1)">&#8249;</button>
    <div class="week-nav-center">
      <div class="week-nav-date" id="mWeekDate">–</div>
      <div class="week-nav-price" id="mWeekPrice"></div>
    </div>
    <button class="nav-arrow" id="nextBtn" onclick="changeWeek(1)">&#8250;</button>
  </div>
  <div class="week-dots" id="weekDots"></div>
  <div class="mobile-filter">
    <input type="text" id="mRoomFilter" placeholder="🔍 Sök rum (t.ex. K14)" oninput="renderMobile()"/>
    <select id="mStatusFilter" onchange="renderMobile()">
      <option value="all">Alla rum</option>
      <option value="available">Bara lediga</option>
    </select>
  </div>
  <div id="roomList"></div>
</div>

<div class="info-box">ℹ️ Rum kan vara prel. bokade i upp till 48h utan att synas här. Kontakta oss för aktuell status.</div>
<a href="mailto:pia@wildwind.se?subject=Wildwind%20bokningsf%C3%B6rfr%C3%A5gan" class="contact-btn">✉️ Skicka bokningsförfrågan</a>

<script>
{data_js}

document.getElementById('upd').textContent = UPDATED;

const SECTIONS = [
  {{ label: "Melas Hotel",        rooms: ["M1","M2A/B","M3","M4","M7A/B","M8"] }},
  {{ label: "Kavadias Hotel",     rooms: ["K2","K3","K4","K14","K15","K18","K19"] }},
  {{ label: "AKTI Apartments",    rooms: ["A4","A5","A7","A8","A9"] }},
  {{ label: "Xristina Apartments",rooms: ["X6"] }},
  {{ label: "New Melas",          rooms: ["NM208"] }},
];

// Närmaste kommande lördag
const today = new Date();
let currentWeek = 0;
for (let i = 0; i < WEEKS.length; i++) {{
  if (new Date(WEEKS[i].date) >= today) {{ currentWeek = i; break; }}
}}

// Desktop vecko-dropdown
const wf = document.getElementById('weekFilter');
WEEKS.forEach((w,i) => {{
  const o = document.createElement('option');
  o.value = i;
  o.textContent = w.display + (w.price ? ' – ' + w.price.toLocaleString('sv-SE') + ' kr' : '');
  wf.appendChild(o);
}});

function roomsByFilter(roomQ, statQ, idxs) {{
  return ROOMS.filter(r => {{
    if (roomQ && !r.name.toLowerCase().includes(roomQ)) return false;
    if (statQ === 'available' && !idxs.some(i => r.weeks[i] === 'available')) return false;
    return true;
  }});
}}

function renderTable() {{
  const roomQ = document.getElementById('roomFilter').value.toLowerCase();
  const statQ = document.getElementById('statusFilter').value;
  const weekQ = document.getElementById('weekFilter').value;
  const idxs  = weekQ === 'all' ? WEEKS.map((_,i)=>i) : [parseInt(weekQ)];
  const filtered = roomsByFilter(roomQ, statQ, idxs);
  const filtNames = new Set(filtered.map(r => r.name));

  let h = '<thead><tr><th>Rum</th>';
  idxs.forEach(i => {{
    const w = WEEKS[i];
    h += `<th><span class="wk-date">${{w.display}}</span>`;
    h += w.price ? `<span class="wk-price">${{w.price.toLocaleString('sv-SE')}} kr</span>` : `<span class="wk-price">&nbsp;</span>`;
    h += '</th>';
  }});
  h += '</tr></thead><tbody>';

  SECTIONS.forEach(sec => {{
    const secRooms = sec.rooms.filter(n => filtNames.has(n));
    if (!secRooms.length) return;
    h += `<tr class="section-row"><td colspan="${{idxs.length+1}}">${{sec.label}}</td></tr>`;
    secRooms.forEach(name => {{
      const room = ROOMS.find(r => r.name === name);
      if (!room) return;
      const info = ROOM_INFO[name] || {{}};
      const tip = info.info ? `${{name}}: ${{info.info}} | Tillägg: ${{info.tillagg||'–'}}` : name;
      h += `<tr><td title="${{tip}}">${{name}}</td>`;
      idxs.forEach(i => {{
        const st  = room.weeks[i] || 'available';
        const lbl = st==='available'?'✓':st==='on_hold'?'~':'✕';
        h += `<td><div class="cell cell-${{st}}" title="${{WEEKS[i].display}}: ${{st==='available'?'Ledig':st==='on_hold'?'Prel.':'Bokad'}}">${{lbl}}</div></td>`;
      }});
      h += '</tr>';
    }});
  }});
  h += '</tbody>';
  document.getElementById('tbl').innerHTML = h;
}}

function renderMobile() {{
  const roomQ = document.getElementById('mRoomFilter').value.toLowerCase();
  const statQ = document.getElementById('mStatusFilter').value;
  const w = WEEKS[currentWeek];

  const endDate = new Date(w.date);
  endDate.setDate(endDate.getDate() + 7);
  const endLabel = endDate.toLocaleDateString('sv-SE', {{day:'numeric', month:'short'}});

  document.getElementById('mWeekDate').textContent  = w.display + ' – ' + endLabel;
  document.getElementById('mWeekPrice').textContent = w.price ? 'Sailing fr. ' + w.price.toLocaleString('sv-SE') + ' kr/person/v' : '';
  document.getElementById('prevBtn').disabled = currentWeek === 0;
  document.getElementById('nextBtn').disabled = currentWeek === WEEKS.length - 1;

  // Dots
  const dc = document.getElementById('weekDots');
  dc.innerHTML = '';
  WEEKS.forEach((_,i) => {{
    const d = document.createElement('div');
    d.className = 'wdot' + (i===currentWeek?' active':'');
    d.onclick = () => {{ currentWeek=i; renderMobile(); }};
    dc.appendChild(d);
  }});

  const stLabel = {{available:'Ledig', on_hold:'Preliminär', booked:'Bokad'}};
  let html = '';

  SECTIONS.forEach(sec => {{
    const secRooms = sec.rooms.filter(name => {{
      if (roomQ && !name.toLowerCase().includes(roomQ)) return false;
      const room = ROOMS.find(r => r.name===name);
      if (!room) return false;
      if (statQ==='available' && room.weeks[currentWeek] !== 'available') return false;
      return true;
    }});
    if (!secRooms.length) return;
    html += `<div class="section-header">${{sec.label}}</div>`;
    secRooms.forEach(name => {{
      const room = ROOMS.find(r => r.name===name);
      if (!room) return;
      const st = room.weeks[currentWeek] || 'available';
      const info = ROOM_INFO[name] || {{}};
      html += `<div class="room-row">
        <div class="room-badge badge-${{st}}">${{name}}</div>
        <div class="room-details">
          <div class="room-name">${{name}}</div>
          ${{info.info ? `<div class="room-desc">${{info.info}}</div>` : ''}}
          ${{info.tillagg ? `<div class="room-tillagg">Tillägg: ${{info.tillagg}}</div>` : ''}}
        </div>
        <div class="room-status status-${{st}}">${{stLabel[st]}}</div>
      </div>`;
    }});
  }});

  document.getElementById('roomList').innerHTML = html || '<p style="padding:20px;color:var(--mid);text-align:center;">Inga rum matchar.</p>';
}}

function changeWeek(dir) {{
  currentWeek = Math.max(0, Math.min(WEEKS.length-1, currentWeek+dir));
  renderMobile();
}}

renderTable();
renderMobile();
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
