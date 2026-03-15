#!/usr/bin/env python3
"""
Wildwind Room Availability Updater
Downloads Excel from Dropbox and generates availability.html
"""

import re
import json
import requests
import pandas as pd
from datetime import datetime
from pathlib import Path

# ─── CONFIG ───────────────────────────────────────────────────────────
DROPBOX_URL = "https://www.dropbox.com/scl/fi/i0ji499omsoc0yl4fuzo7/STEPH-VERSION-ALL-ROOMS-2026.xlsx?rlkey=lq3j8vz7hdlnoqvairj0yp2ih&dl=1"
OUTPUT_FILE = "availability.html"

# Excel-rader som Seafari AB kan boka (1-indexerade som i Excel)
ALLOWED_ROWS = [44, 48, 52, 56, 60, 140, 144, 148, 168, 184, 188, 192,
                204, 208, 220, 224, 244, 248, 256, 264, 272, 284]

# Sailing-priser 2026 (SEK/person/vecka, 2 delar rum)
SAILING_PRICES = {
    "Apr 25": 8990,  "May 2":  9430,  "May 9":  10300, "May 16": 10840,
    "May 23": 10840, "May 30": 11940, "Jun 06": 12700, "Jun 13": 13030,
    "Jun 20": 13900, "Jun 27": 14120, "Jul 4":  14440, "Jul 11": 14990,
    "Jul 18": 15210, "Jul 25": 15210, "Aug 01": 15210, "Aug 8":  15210,
    "Aug 15": 14120, "Aug 22": 12810, "Aug 29": 11940, "Sep 5":  11720,
    "Sep 12": 11170, "Sep 19": 10300, "Sep 26": 9540,  "Oct 3":  8990,
}

def price_for_date(dt):
    for fmt in ["%-d %b", "%b %-d", "%b %d"]:
        try:
            key = dt.strftime(fmt)
            if key in SAILING_PRICES:
                return SAILING_PRICES[key]
        except:
            pass
    return None

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

    # ── Hitta lördagskolumner: rad 1 innehåller "SATURDAY" ──
    saturdays = []
    for col_idx in range(3, len(df.columns), 2):
        date_val = df.iloc[0, col_idx]
        day_name = df.iloc[1, col_idx]
        if pd.notna(day_name) and 'SATURDAY' in str(day_name).upper():
            if pd.notna(date_val):
                dt = pd.to_datetime(date_val)
                saturdays.append({
                    'col':     col_idx,
                    'date':    dt.strftime('%Y-%m-%d'),
                    'display': dt.strftime('%-d %b'),
                    'price':   price_for_date(dt),
                })

    print(f"   Hittade {len(saturdays)} lördagar")

    # ── Hämta rum ──
    rooms = []
    for row in ALLOWED_ROWS:
        r_idx = row - 1   # 0-indexerat
        if r_idx >= len(df):
            continue

        # Kolumn 0 = hotellnamn, kolumn 1 = rumsnummer
        building = str(df.iloc[r_idx, 0]).strip() if pd.notna(df.iloc[r_idx, 0]) else ""
        room_num = str(df.iloc[r_idx, 1]).strip() if pd.notna(df.iloc[r_idx, 1]) else ""

        building = re.sub(r'\s+', ' ', building).replace("nan", "").strip()
        room_num = re.sub(r'\s+', ' ', room_num).replace("nan", "").strip()

        # Använd rumsnummer om det finns, annars hotellnamn
        room_name = room_num if room_num and room_num.lower() != "nan" else building
        if not room_name:
            room_name = f"Rad {row}"

        # ── Status per vecka (kolla alla 7 dagar) ──
        week_status = []
        for sat in saturdays:
            sat_col     = sat['col']
            booked      = False
            on_hold     = False

            for day_offset in range(7):
                col = sat_col + (day_offset * 2)
                if col >= len(df.columns):
                    break
                cell_val = df.iloc[r_idx, col]
                if pd.notna(cell_val) and str(cell_val).strip() not in ("", "nan"):
                    val_str = str(cell_val).strip().lower()
                    if any(k in val_str for k in ("hold", "option", "prel")):
                        on_hold = True
                    else:
                        booked = True

            week_status.append("booked" if booked else "on_hold" if on_hold else "available")

        rooms.append({"name": room_name, "weeks": week_status})

    print(f"   Hittade {len(rooms)} rum")
    return {"weeks": saturdays, "rooms": rooms}

# ─── GENERATE HTML ────────────────────────────────────────────────────
ROOM_INFO = {
    "M1":    {"tillagg": "0 kr",      "info": "Twin · Sidoutsikt hav + Kav Bar"},
    "M2A/B": {"tillagg": "0 kr",      "info": "Twin+Double · Familjerum, sid- + full havsutsikt"},
    "M7A/B": {"tillagg": "0 kr",      "info": "Double+Twin · Familjerum, mot Ponti västerut"},
    "M8":    {"tillagg": "0 kr",      "info": "Twin · Mot Ponti + berg västerut"},
    "M3":    {"tillagg": "1 758 kr",  "info": "Double · Full havsutsikt"},
    "M4":    {"tillagg": "1 758 kr",  "info": "Twin · Full havsutsikt"},
    "NM208": {"tillagg": "2 638 kr",  "info": "Double · Bergutsikt bakåt, övervåning"},
    "K2":    {"tillagg": "0 kr",      "info": "Small Double · Ingen utsikt, baksida"},
    "K3":    {"tillagg": "0 kr",      "info": "Twin · Ingen utsikt, baksida (litet rum)"},
    "K4":    {"tillagg": "0 kr",      "info": "Double · Sidoutsikt mot trädgård"},
    "K14":   {"tillagg": "1 320 kr",  "info": "Double · Pool + berg västerut, övervåning"},
    "K15":   {"tillagg": "1 320 kr",  "info": "Double · Pool + berg, delar balkong K14"},
    "K18":   {"tillagg": "1 320 kr",  "info": "Twin · Havsutsikt över Kav Bar, övervåning"},
    "K19":   {"tillagg": "1 320 kr",  "info": "Twin · Sidoutsikt mot hav + byn"},
    "A4":    {"tillagg": "1 758 kr",  "info": "Studio · Twin, pentry, övervåning"},
    "A5":    {"tillagg": "1 758 kr",  "info": "Studio · Twin, pentry, övervåning"},
    "A9":    {"tillagg": "1 758 kr",  "info": "1-sovrum · Twin + extra i hall, pentry"},
    "A7":    {"tillagg": "3 516 kr",  "info": "2-sovrum · Double+Twin+extra, fullt kök (min 4 pers)"},
    "A8":    {"tillagg": "3 516 kr",  "info": "2-sovrum · Double+Twin+extra, fullt kök (min 4 pers)"},
    "X1":    {"tillagg": "1 320 kr",  "info": "Studio · Twin, pentry, havsutsikt på avstånd (13 jul–15 sep)"},
    "X6":    {"tillagg": "3 516 kr",  "info": "1-sovrum · Double+2 enklar i lounge, kitchenette (13 jul–15 sep)"},
}

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
  --azure:#0B3D72;--azure2:#1A5C9A;--sky:#5BA4CF;--sand:#F5EDD8;--sand2:#E6D5BA;
  --green:#2E8B57;--orange:#E08C2A;--red:#C0392B;
  --green-bg:#E8F5EE;--orange-bg:#FEF3E2;--red-bg:#FDECEC;
  --white:#FDFAF5;--text:#2A2A2A;--mid:#666;--light:#999;
}}
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'Nunito Sans',sans-serif;background:var(--white);color:var(--text);font-size:14px;}}

/* ── HEADER ── */
header{{background:var(--azure);color:#fff;padding:14px 20px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;}}
header h1{{font-family:'Cormorant Garamond',serif;font-size:21px;font-weight:400;}}
.header-right{{display:flex;align-items:center;gap:12px;}}
.updated{{font-size:11px;opacity:0.6;}}
.reload-btn{{background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.3);color:#fff;padding:6px 14px;border-radius:20px;cursor:pointer;font-size:12px;font-family:inherit;font-weight:600;transition:background .2s;}}
.reload-btn:hover{{background:rgba(255,255,255,0.28);}}

/* ── LEGEND ── */
.legend{{display:flex;gap:16px;padding:10px 20px;background:#fff;border-bottom:1px solid #e8e0d0;flex-wrap:wrap;align-items:center;}}
.legend-item{{display:flex;align-items:center;gap:6px;font-size:12px;font-weight:600;}}
.dot{{width:12px;height:12px;border-radius:3px;flex-shrink:0;}}
.dot-green{{background:var(--green);}} .dot-orange{{background:var(--orange);}} .dot-red{{background:var(--red);}}
.note{{font-size:11px;color:var(--mid);margin-left:auto;}}

/* ══════════════════════════════════
   DESKTOP TABLE VIEW
══════════════════════════════════ */
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
.cell{{width:46px;height:32px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;cursor:default;}}
.cell:hover{{opacity:0.75;}}
.cell-available{{background:var(--green-bg);color:var(--green);}}
.cell-on_hold{{background:var(--orange-bg);color:var(--orange);}}
.cell-booked{{background:var(--red-bg);color:var(--red);}}

/* ══════════════════════════════════
   MOBILE VIEW
══════════════════════════════════ */
.mobile-view{{display:none;}}

/* Veckonavigering */
.week-nav{{
  background:var(--azure);color:#fff;
  padding:12px 16px;display:flex;align-items:center;justify-content:space-between;
  position:sticky;top:0;z-index:10;
}}
.week-nav-center{{text-align:center;flex:1;}}
.week-nav-date{{font-family:'Cormorant Garamond',serif;font-size:20px;font-weight:600;}}
.week-nav-price{{font-size:12px;color:#FFE099;margin-top:2px;}}
.nav-arrow{{
  background:rgba(255,255,255,0.15);border:none;color:#fff;
  width:40px;height:40px;border-radius:50%;font-size:20px;
  cursor:pointer;display:flex;align-items:center;justify-content:center;
  transition:background .2s;flex-shrink:0;
}}
.nav-arrow:hover{{background:rgba(255,255,255,0.3);}}
.nav-arrow:disabled{{opacity:0.25;cursor:default;}}

/* Veckoindikator (dots) */
.week-dots{{
  display:flex;gap:4px;justify-content:center;align-items:center;
  padding:8px 16px;background:var(--azure2);
}}
.wdot{{width:6px;height:6px;border-radius:50%;background:rgba(255,255,255,0.3);transition:all .2s;cursor:pointer;}}
.wdot.active{{background:#fff;width:8px;height:8px;}}

/* Filter mobil */
.mobile-filter{{
  padding:10px 16px;background:var(--sand);
  display:flex;gap:10px;align-items:center;flex-wrap:wrap;
}}
.mobile-filter select,.mobile-filter input{{
  padding:7px 12px;border:1px solid #ccc;border-radius:20px;
  font-family:inherit;font-size:13px;background:#fff;flex:1;min-width:120px;
}}

/* Rumlista */
.room-list{{padding:8px 0;}}
.room-row{{
  display:flex;align-items:center;
  padding:13px 16px;border-bottom:1px solid #ede8df;
  gap:14px;transition:background .15s;
}}
.room-row:active{{background:var(--sand);}}
.room-badge{{
  width:52px;height:52px;border-radius:10px;
  display:flex;align-items:center;justify-content:center;
  font-weight:700;font-size:13px;flex-shrink:0;letter-spacing:0.01em;
}}
.badge-available{{background:var(--green-bg);color:var(--green);border:1.5px solid var(--green);}}
.badge-on_hold  {{background:var(--orange-bg);color:var(--orange);border:1.5px solid var(--orange);}}
.badge-booked   {{background:var(--red-bg);color:var(--red);border:1.5px solid var(--red);}}
.room-details{{flex:1;min-width:0;}}
.room-name{{font-weight:700;font-size:15px;color:var(--azure);}}
.room-info{{font-size:12px;color:var(--mid);margin-top:3px;line-height:1.4;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.room-tillagg{{font-size:12px;font-weight:600;color:var(--text);margin-top:2px;}}
.room-status-text{{font-size:12px;font-weight:700;flex-shrink:0;}}
.status-available{{color:var(--green);}}
.status-on_hold  {{color:var(--orange);}}
.status-booked   {{color:var(--red);}}

/* Info + knapp */
.info-box{{margin:16px;padding:12px 16px;background:#EDF3FA;border-left:4px solid var(--sky);border-radius:6px;font-size:12px;color:var(--mid);line-height:1.6;}}
.contact-btn{{display:flex;align-items:center;justify-content:center;gap:8px;margin:0 16px 32px;padding:15px;background:var(--azure);color:#fff;border-radius:40px;text-decoration:none;font-size:14px;font-weight:600;transition:background .2s;}}
.contact-btn:hover{{background:var(--azure2);}}

/* ── SWITCH ── */
@media (max-width: 700px) {{
  .desktop-view {{ display: none; }}
  .mobile-view  {{ display: block; }}
  .legend .note {{ display: none; }}
}}
</style>
</head>
<body>

<header>
  <h1>🌊 Wildwind 2026</h1>
  <div class="header-right">
    <span class="updated">Uppdaterad: <span id="upd"></span></span>
    <button class="reload-btn" onclick="location.reload(true)">↺</button>
  </div>
</header>

<div class="legend">
  <div class="legend-item"><div class="dot dot-green"></div>Ledig</div>
  <div class="legend-item"><div class="dot dot-orange"></div>Preliminärt</div>
  <div class="legend-item"><div class="dot dot-red"></div>Bokad</div>
  <span class="note">Lördag–lördag · Sailing fr. pris/person/v</span>
</div>

<!-- ══ DESKTOP ══ -->
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

<!-- ══ MOBIL ══ -->
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
  <div class="room-list" id="roomList"></div>
</div>

<div class="info-box">ℹ️ Rum kan vara prel. bokade i upp till 48h utan att synas här. Kontakta oss för aktuell status.</div>
<a href="mailto:pia@wildwind.se?subject=Wildwind%20bokningsf%C3%B6rfr%C3%A5gan" class="contact-btn">✉️ Skicka bokningsförfrågan</a>

<script>
{data_js}

document.getElementById('upd').textContent = UPDATED;

/* ── Hitta närmaste kommande vecka ── */
const today = new Date();
let startWeek = 0;
for (let i = 0; i < WEEKS.length; i++) {{
  if (new Date(WEEKS[i].date) >= today) {{ startWeek = i; break; }}
}}
let currentWeek = startWeek;

/* ── Desktop: fyll vecko-dropdown ── */
const wf = document.getElementById('weekFilter');
WEEKS.forEach((w,i) => {{
  const o = document.createElement('option');
  o.value = i;
  o.textContent = w.display + (w.price ? ' – ' + w.price.toLocaleString('sv-SE') + ' kr' : '');
  wf.appendChild(o);
}});

function renderTable() {{
  const roomQ = document.getElementById('roomFilter').value.toLowerCase();
  const statQ = document.getElementById('statusFilter').value;
  const weekQ = document.getElementById('weekFilter').value;
  const idxs  = weekQ === 'all' ? WEEKS.map((_,i)=>i) : [parseInt(weekQ)];
  let h = '<thead><tr><th>Rum</th>';
  idxs.forEach(i => {{
    const w = WEEKS[i];
    h += `<th><span class="wk-date">${{w.display}}</span>`;
    h += w.price ? `<span class="wk-price">${{w.price.toLocaleString('sv-SE')}} kr</span>` : `<span class="wk-price">&nbsp;</span>`;
    h += '</th>';
  }});
  h += '</tr></thead><tbody>';
  ROOMS.forEach(room => {{
    if (roomQ && !room.name.toLowerCase().includes(roomQ)) return;
    if (statQ === 'available' && !idxs.some(i => room.weeks[i] === 'available')) return;
    const info = ROOM_INFO[room.name] || {{}};
    const tooltip = info.info ? `${{room.name}}: ${{info.info}} | Tillägg: ${{info.tillagg||'–'}}` : room.name;
    h += `<tr><td title="${{tooltip}}">${{room.name}}</td>`;
    idxs.forEach(i => {{
      const st  = room.weeks[i] || 'available';
      const lbl = st === 'available' ? '✓' : st === 'on_hold' ? '~' : '✕';
      const tip = WEEKS[i].display + ' – ' + (st==='available'?'Ledig':st==='on_hold'?'Prel. bokad':'Bokad');
      h += `<td><div class="cell cell-${{st}}" title="${{tip}}">${{lbl}}</div></td>`;
    }});
    h += '</tr>';
  }});
  h += '</tbody>';
  document.getElementById('tbl').innerHTML = h;
}}

/* ── Mobil ── */
function buildDots() {{
  const c = document.getElementById('weekDots');
  c.innerHTML = '';
  WEEKS.forEach((w,i) => {{
    const d = document.createElement('div');
    d.className = 'wdot' + (i === currentWeek ? ' active' : '');
    d.onclick = () => {{ currentWeek = i; renderMobile(); }};
    c.appendChild(d);
  }});
}}

function renderMobile() {{
  const roomQ = document.getElementById('mRoomFilter').value.toLowerCase();
  const statQ = document.getElementById('mStatusFilter').value;
  const w     = WEEKS[currentWeek];

  document.getElementById('mWeekDate').textContent  = w.display + ' – ' + formatEndDate(w.date);
  document.getElementById('mWeekPrice').textContent = w.price ? 'Sailing fr. ' + w.price.toLocaleString('sv-SE') + ' kr/person/v' : '';
  document.getElementById('prevBtn').disabled = currentWeek === 0;
  document.getElementById('nextBtn').disabled = currentWeek === WEEKS.length - 1;

  buildDots();

  const statusLabel = {{ available:'Ledig', on_hold:'Preliminär', booked:'Bokad' }};
  let html = '';
  ROOMS.forEach(room => {{
    if (roomQ && !room.name.toLowerCase().includes(roomQ)) return;
    const st = room.weeks[currentWeek] || 'available';
    if (statQ === 'available' && st !== 'available') return;
    const info = ROOM_INFO[room.name] || {{}};
    html += `
      <div class="room-row">
        <div class="room-badge badge-${{st}}">${{room.name}}</div>
        <div class="room-details">
          <div class="room-name">${{room.name}}</div>
          ${{info.info ? `<div class="room-info">${{info.info}}</div>` : ''}}
          ${{info.tillagg ? `<div class="room-tillagg">Tillägg: ${{info.tillagg}}</div>` : ''}}
        </div>
        <div class="room-status-text status-${{st}}">${{statusLabel[st]}}</div>
      </div>`;
  }});
  document.getElementById('roomList').innerHTML = html || '<p style="padding:20px;color:var(--mid);text-align:center;">Inga rum matchar filtret.</p>';
}}

function changeWeek(dir) {{
  currentWeek = Math.max(0, Math.min(WEEKS.length - 1, currentWeek + dir));
  renderMobile();
}}

function formatEndDate(isoDate) {{
  const d = new Date(isoDate);
  d.setDate(d.getDate() + 7);
  return d.toLocaleDateString('sv-SE', {{day:'numeric', month:'short'}});
}}

renderTable();
renderMobile();
</script>
</body>
</html>'''

# ─── MAIN ─────────────────────────────────────────────────────────────
def main():
    print("🚀 Wildwind Availability Updater")
    print("=" * 40)
    try:
        excel_file = download_excel()
        data       = parse_availability(excel_file)
        html       = generate_html(data)
        Path(OUTPUT_FILE).write_text(html, encoding='utf-8')
        print(f"✅ Genererade {OUTPUT_FILE}")
        Path("temp_availability.xlsx").unlink(missing_ok=True)
        print("🎉 Klar!")
    except Exception as e:
        print(f"❌ Fel: {e}")
        raise

if __name__ == "__main__":
    main()
