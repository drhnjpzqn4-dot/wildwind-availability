#!/usr/bin/env python3
"""Wildwind Room Availability Updater"""

import json, requests, pandas as pd
from datetime import datetime
from pathlib import Path

DROPBOX_URL = "https://www.dropbox.com/scl/fi/00tqgjw04ph26l50cltod/STEPH-VERSION-ALL-ROOMS-2027.xlsx?rlkey=uaw1ngacjal6m6ffs4izq25l2&dl=1"
OUTPUT_FILE = "availability.html"

ROW_TO_NAME = {
    44:"Melas 1", 48:"Melas 2AB", 56:"Melas 3", 60:"Melas 4", 80:"Melas 8",
    140:"Kav 2", 144:"Kav 3", 148:"Kav 4", 168:"Kav 9",
    184:"Kav 13", 188:"Kav 14", 192:"Kav 15", 204:"Kav 18", 208:"Kav 19",
    216:"Akti 7", 220:"Akti 8", 224:"Akti 9", 244:"Akti 4", 248:"Akti 5",
    256:"Xristina 2", 264:"Xristina 4", 272:"Xristina 6", 276:"X Bungalow",
    284:"NM208",
}
COMBINED_ROOMS = {"Melas 2AB":[48,52]}

ALL_PRICES = {  # 2027, SEK (EUR x 11.1530, Riksbanken okt 2026)
    "24 Apr": (9980, 15340, None, None, 8530, 11990, 1950),
    "1 May": (10540, 15560, None, None, 8870, 14110, 2170),
    "8 May": (11430, 17230, None, None, 9200, 14220, 2230),
    "15 May": (11990, 18900, 15220, 25370, 9310, 14780, 2450),
    "22 May": (12100, 19800, 15340, 26270, 8760, 15340, 2450),
    "29 May": (12210, 20350, 15450, 26820, 8870, 15340, 2730),
    "5 Jun": (13890, 21470, 17120, 27940, 10650, 16450, 2840),
    "12 Jun": (14440, 22250, 17680, 28720, 11100, 16670, 3010),
    "19 Jun": (15340, 24260, 18570, 30730, 12770, 17790, 3120),
    "26 Jun": (15560, 24480, 18790, 30950, 13330, 18900, 3230),
    "3 Jul": (15890, 24820, 19130, 31280, 13660, 19240, 3230),
    "10 Jul": (16450, 25370, 19690, 31840, 13660, 18900, 3230),
    "17 Jul": (16670, 25600, 19910, 32060, 13660, 18900, 3230),
    "24 Jul": (16670, 25600, 19910, 32060, 13660, 18900, 3230),
    "31 Jul": (16670, 25600, 19910, 32060, 13660, 18900, 3230),
    "7 Aug": (16670, 25600, 19910, 32060, 13660, 18900, 3230),
    "14 Aug": (15560, 23920, 18790, 30390, 13660, 19460, 3230),
    "21 Aug": (14220, 22250, 17450, 28720, 13330, 18570, 2790),
    "28 Aug": (13330, 20350, 16560, 26820, 11100, 17010, 2510),
    "4 Sep": (13100, 19800, 16340, 26270, 10540, 16450, 2450),
    "11 Sep": (12550, 18680, 15780, 25150, 10320, 16230, 2170),
    "18 Sep": (11430, 17230, 14670, 23700, 9760, 14220, 2120),
    "25 Sep": (10540, 15560, 13770, None, 8530, 13330, 1950),
    "2 Oct": (9420, 14220, None, None, 7750, None, 1950),
}

ROOM_INFO = {
    "Melas 1":   {"t":0,   "desc":"Twin · Sidoutsikt Kav Bar"},
    "Melas 2AB": {"t":0,   "desc":"Double+Twin · Familjerum · Havsutsikt"},
    "Melas 3":   {"t":2790,"desc":"Double · Full havsutsikt"},
    "Melas 4":   {"t":2790,"desc":"Twin · Full havsutsikt"},
    "Melas 8":   {"t":0,   "desc":"Twin · Utsikt mot Ponti"},
    "Kav 2":     {"t":0,   "desc":"Twin · Mot berg norrut"},
    "Kav 3":     {"t":0,   "desc":"Twin · Mot berg norrut"},
    "Kav 4":     {"t":0,   "desc":"Double · Mot gräsmattan"},
    "Kav 9":     {"t":0,   "desc":"Double · Mot poolen"},
    "Kav 13":    {"t":1560,"desc":"Twin · Övervåning"},
    "Kav 14":    {"t":1560,"desc":"Double · Mot poolen · Övervåning"},
    "Kav 15":    {"t":1560,"desc":"Double · Mot poolen · Övervåning"},
    "Kav 18":    {"t":1560,"desc":"Twin · Havsutsikt · Övervåning"},
    "Kav 19":    {"t":1560,"desc":"Twin · Mot byn · Övervåning"},
    "Akti 7":    {"t":1900,"desc":"2-sovrum · Fullt kök · Min 4 pers"},
    "Akti 8":    {"t":1900,"desc":"2-sovrum · Fullt kök · Min 4 pers"},
    "Akti 9":    {"t":1900,"desc":"1-sovrum · Pentry"},
    "Akti 4":    {"t":1560,"desc":"Studio · Pentry · Övervåning"},
    "Akti 5":    {"t":1560,"desc":"Studio · Pentry · Övervåning"},
    "Xristina 2":{"t":1560,"desc":"Studio · Havsutsikt på avstånd"},
    "Xristina 4":{"t":1560,"desc":"Studio · Trädgårdsvy norrut"},
    "Xristina 6":{"t":1900,"desc":"1-sovrum · 13 jul–15 sep"},
    "X Bungalow":{"t":1560,"desc":"Eget hus · Mot bergen"},
    "NM208":     {"t":3570,"desc":"Double · Bergutsikt · New Melas"},
}

SECTIONS = [
    {"label":"Melas Hotel",         "rooms":["Melas 1","Melas 2AB","Melas 3","Melas 4","Melas 8"]},
    {"label":"Kavadias Hotel",      "rooms":["Kav 2","Kav 3","Kav 4","Kav 9","Kav 13","Kav 14","Kav 15","Kav 18","Kav 19"]},
    {"label":"AKTI Apartments",     "rooms":["Akti 4","Akti 5","Akti 7","Akti 8","Akti 9"]},
    {"label":"Xristina Apartments", "rooms":["Xristina 2","Xristina 4","Xristina 6","X Bungalow"]},
    {"label":"New Melas",           "rooms":["NM208"]},
]

def download_excel():
    print("Laddar ner Excel...")
    r = requests.get(DROPBOX_URL, timeout=30)
    r.raise_for_status()
    p = Path("temp_availability.xlsx")
    p.write_bytes(r.content)
    print(f"OK {len(r.content):,} bytes")
    return p

def parse_availability(path):
    print("Läser Excel...")
    df = pd.read_excel(path, header=None)
    saturdays = []
    for col in range(3, len(df.columns), 2):
        dv, dn = df.iloc[0,col], df.iloc[1,col]
        if pd.notna(dn) and "SATURDAY" in str(dn).upper() and pd.notna(dv):
            dt = pd.to_datetime(dv)
            lb = dt.strftime("%-d %b")
            if lb not in ALL_PRICES:
                continue  # bara veckor i Wildwinds prislista (t.ex. inte 9 okt, reserverad för engelsk grupp)
            p  = ALL_PRICES[lb]
            saturdays.append({"col":col,"date":dt.strftime("%Y-%m-%d"),
                "display":lb,"sail1":p[0],"sail2":p[1],
                "faw1":p[2],"faw2":p[3],"ho1":p[4],"ho2":p[5],"single":p[6]})
    print(f"  {len(saturdays)} lördagar")

    def row_st(ri):
        result = []
        for sat in saturdays:
            booked = on_hold = False
            for d in range(7):
                c = sat["col"] + d*2
                if c >= len(df.columns): break
                v = df.iloc[ri,c]
                if pd.notna(v) and str(v).strip() not in ("","nan"):
                    booked = True
                for off in [1,2,3]:
                    if ri+off < len(df):
                        h = df.iloc[ri+off,c]
                        if pd.notna(h) and any(k in str(h).lower() for k in ("hold","option","prel")):
                            on_hold = True
            result.append("booked" if booked else "on_hold" if on_hold else "available")
        return result

    rooms, added = [], set()
    for row,name in ROW_TO_NAME.items():
        if name in added: continue
        ri = row-1
        if ri >= len(df): continue
        if name in COMBINED_ROOMS:
            all_st = [row_st(r-1) for r in COMBINED_ROOMS[name]]
            ws = []
            for i in range(len(saturdays)):
                col = [s[i] for s in all_st]
                ws.append("booked" if "booked" in col else "on_hold" if "on_hold" in col else "available")
        else:
            ws = row_st(ri)
        rooms.append({"name":name,"weeks":ws})
        added.add(name)
    print(f"  {len(rooms)} rum klara")
    return {"weeks":saturdays,"rooms":rooms}

def generate_html(data):
    weeks = data["weeks"]
    rooms = data["rooms"]
    now   = datetime.now().strftime("%d %b %Y, %H:%M")
    djs   = (f"const WEEKS={json.dumps(weeks,ensure_ascii=False)};\n"
             f"const ROOMS={json.dumps(rooms,ensure_ascii=False)};\n"
             f"const ROOM_INFO={json.dumps(ROOM_INFO,ensure_ascii=False)};\n"
             f"const SECTIONS={json.dumps(SECTIONS,ensure_ascii=False)};\n"
             f"const UPDATED='{now}';\n")
    return f"""<!DOCTYPE html>
<html lang="sv">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Wildwind – Rumstillgänglighet 2027</title>
<link href="https://fonts.googleapis.com/css2?family=Nunito+Sans:wght@300;400;600;700&family=Cormorant+Garamond:wght@400;600&display=swap" rel="stylesheet"/>
<style>
:root{{
  --azure:#0B3D72;--azure2:#1A5C9A;--sand:#F5EDD8;--sand2:#E6D5BA;
  --green:#2E8B57;--orange:#E08C2A;--red:#C0392B;--gold:#C9963A;
  --green-bg:#E8F5EE;--orange-bg:#FEF3E2;--red-bg:#FDECEC;
  --white:#FDFAF5;--text:#2A2A2A;--mid:#666;
}}
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'Nunito Sans',sans-serif;background:var(--white);color:var(--text);font-size:13px;}}

/* HEADER */
header{{background:var(--azure);color:#fff;padding:12px 20px;display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap;}}
header h1{{font-family:'Cormorant Garamond',serif;font-size:22px;font-weight:400;}}
.upd{{font-size:11px;opacity:.6;}}
.reload{{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);color:#fff;padding:5px 12px;border-radius:20px;cursor:pointer;font-size:11px;font-family:inherit;font-weight:600;}}

/* TOPBAR */
.topbar{{display:flex;gap:16px;padding:9px 20px;background:#fff;border-bottom:1px solid #e8e0d0;flex-wrap:wrap;align-items:center;}}
.li{{display:flex;align-items:center;gap:5px;font-size:11px;font-weight:600;}}
.dot{{width:9px;height:9px;border-radius:50%;}}
.dg{{background:var(--green);}}.do{{background:var(--orange);}}.dr{{background:var(--red);}}
.filter-wrap{{margin-left:auto;display:flex;gap:8px;align-items:center;flex-wrap:wrap;}}
.filter-wrap label{{font-size:11px;font-weight:700;color:var(--azure);}}
.filter-wrap input{{padding:5px 10px;border:1px solid #ccc;border-radius:20px;font-family:inherit;font-size:12px;width:140px;}}
.filter-wrap select{{padding:5px 10px;border:1px solid #ccc;border-radius:20px;font-family:inherit;font-size:12px;background:#fff;}}

/* TABLE */
.table-scroll{{overflow:auto;max-height:calc(100vh - 110px);min-height:320px;margin:12px 20px 20px;border:1px solid #ece7dc;border-radius:8px;}}
table{{border-collapse:collapse;font-size:12px;}}

/* Rum-kolumn (header) */
th.room-th{{
  position:sticky;left:0;top:0;z-index:6;
  background:var(--azure);color:#fff;
  padding:8px 12px;text-align:left;
  min-width:130px;font-size:11px;font-weight:700;
  border-right:2px solid rgba(255,255,255,.2);
}}

/* Vecko-headers */
th.week-th{{
  background:var(--azure);color:#fff;
  padding:0;text-align:center;
  min-width:52px;width:52px;
  position:sticky;top:0;z-index:4;
}}
th.week-th .wh-inner{{
  display:flex;flex-direction:column;align-items:center;
  padding:6px 4px 5px;
}}
th.week-th .wh-date{{font-size:11px;font-weight:700;white-space:nowrap;}}
th.week-th .wh-price{{font-size:9px;color:#FFE099;margin-top:2px;white-space:nowrap;}}

/* Sektionsrad */
tr.sec-row td{{
  background:var(--azure) !important;
  color:#fff !important;
  font-size:10px;font-weight:700;
  letter-spacing:.12em;text-transform:uppercase;
  padding:5px 12px !important;
  border:none !important;
}}

/* Rumsnamn-cell */
td.room-name-cell{{
  position:sticky;left:0;z-index:1;
  background:#fff;
  border-right:2px solid var(--sand2);
  padding:6px 10px;
  min-width:130px;
}}
tr:nth-child(even) td.room-name-cell{{background:#faf6ee;}}
.rn{{font-weight:700;font-size:12px;color:var(--azure);}}
.rt{{font-size:10px;color:var(--mid);margin-top:1px;}}
.rt-sup{{font-size:10px;color:var(--mid);}}

/* Status-celler */
td.cell{{padding:0;border:1px solid #ece7dc;}}
.ci{{
  width:52px;height:40px;
  display:flex;align-items:center;justify-content:center;
  font-size:13px;font-weight:700;
}}
.ci-available{{background:var(--green-bg);color:var(--green);}}
.ci-on_hold{{background:var(--orange-bg);color:var(--orange);}}
.ci-booked{{background:var(--red-bg);color:var(--red);opacity:.55;}}

/* INFO */
.infobox{{margin:0 20px 24px;padding:10px 14px;background:#EDF3FA;border-left:4px solid #5BA4CF;border-radius:6px;font-size:11px;color:var(--mid);}}
.cta{{display:inline-flex;align-items:center;gap:8px;margin:0 20px 28px;padding:12px 24px;background:var(--azure);color:#fff;border-radius:40px;text-decoration:none;font-size:13px;font-weight:600;}}
.cta:hover{{background:var(--azure2);}}

/* POPUP */
.popup-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:100;align-items:center;justify-content:center;padding:20px;}}
.popup-overlay.open{{display:flex;}}
.popup{{background:#fff;border-radius:14px;padding:28px;max-width:380px;width:100%;box-shadow:0 20px 60px rgba(0,0,0,.25);position:relative;animation:popIn .2s ease;}}
@keyframes popIn{{from{{opacity:0;transform:scale(.95)}}to{{opacity:1;transform:scale(1)}}}}
.popup-close{{position:absolute;top:14px;right:16px;background:none;border:none;font-size:20px;cursor:pointer;color:var(--mid);line-height:1;}}
.popup-close:hover{{color:var(--text);}}
.popup-room{{font-family:'Cormorant Garamond',serif;font-size:26px;font-weight:600;color:var(--azure);margin-bottom:2px;}}
.popup-week{{font-size:13px;color:var(--mid);margin-bottom:16px;}}
.popup-desc{{font-size:12px;color:var(--mid);margin-bottom:16px;padding:10px 12px;background:var(--sand);border-radius:8px;line-height:1.6;}}
.popup-prices{{display:flex;flex-direction:column;gap:6px;margin-bottom:20px;}}
.pp-row{{display:flex;justify-content:space-between;align-items:center;padding:8px 12px;border-radius:8px;background:var(--sand);}}
.pp-row.highlight{{background:var(--azure);color:#fff;}}
.pp-label{{font-size:12px;font-weight:600;}}
.pp-val{{font-size:14px;font-weight:700;}}
.pp-note{{font-size:10px;opacity:.7;margin-top:1px;}}
.popup-ppp{{font-size:11px;color:var(--mid);text-align:center;margin-bottom:16px;}}
.popup-book{{display:flex;align-items:center;justify-content:center;gap:8px;background:var(--azure);color:#fff;padding:14px;border-radius:40px;text-decoration:none;font-size:14px;font-weight:700;transition:background .2s;}}
.popup-book:hover{{background:var(--azure2);}}
</style>
</head>
<body>

<header>
  <h1>🌊 Wildwind 2027</h1>
  <div style="display:flex;align-items:center;gap:12px;">
    <span class="upd">Uppdaterad: <span id="upd"></span></span>
    <button class="reload" onclick="location.reload(true)">↺ Uppdatera</button>
  </div>
</header>

<div class="topbar">
  <div class="li"><div class="dot dg"></div>Ledig</div>
  <div class="li"><div class="dot do"></div>Preliminärt bokad</div>
  <div class="li"><div class="dot dr"></div>Bokad</div>
  <div class="filter-wrap">
    <label>Sök rum:</label>
    <input type="text" id="roomQ" placeholder="t.ex. Melas, Kav" oninput="render()"/>
    <label>Vecka:</label>
    <select id="weekQ" onchange="render()">
      <option value="all">Alla veckor</option>
    </select>
    <select id="statQ" onchange="render()">
      <option value="all">Alla statusar</option>
      <option value="available">Bara lediga</option>
    </select>
  </div>
</div>

<div class="table-scroll">
  <table id="tbl"></table>
</div>

<div class="infobox">ℹ️ Priserna är preliminära. Ett rum kan ha blivit bokat efter vår senaste uppdatering, och vi är små hotell med en ganska manuell bokning. Jag stämmer alltid av innan jag bekräftar.<br><strong>Priser per person baserat på att två delar rum</strong>, exkl. eventuellt rumstillägg. Klicka på en grön ruta för prisdetaljer.</div>
<a href="boka.html#steg1" target="_top" class="cta">Gå till bokningen</a>

<!-- POPUP -->
<div class="popup-overlay" id="overlay" onclick="closePopup(event)">
  <div class="popup">
    <button class="popup-close" onclick="closePopup()">✕</button>
    <div class="popup-room" id="p-room"></div>
    <div class="popup-week" id="p-week"></div>
    <div class="popup-desc" id="p-desc"></div>
    <div class="popup-prices" id="p-prices"></div>
    <div class="popup-ppp">Pris per person baserat på att två delar rum</div>
    <a id="p-book" href="boka.html#steg1" target="_top" class="popup-book">Boka eller anmäl intresse</a>
  </div>
</div>

<script>
{djs}
document.getElementById('upd').textContent = UPDATED;

// Fyll veckodropdown
const wsel = document.getElementById('weekQ');
WEEKS.forEach((w,i) => {{
  const o = document.createElement('option');
  o.value = i;
  o.textContent = w.display + (w.sail1 ? ' – ' + w.sail1.toLocaleString('sv-SE') + ' kr' : '');
  wsel.appendChild(o);
}});

function fmt(n) {{ return n ? n.toLocaleString('sv-SE') : '–'; }}

function render() {{
  const roomQ = document.getElementById('roomQ').value.toLowerCase();
  const weekQ = document.getElementById('weekQ').value;
  const statQ = document.getElementById('statQ').value;
  const idxs  = weekQ === 'all' ? WEEKS.map((_,i)=>i) : [parseInt(weekQ)];

  // Bygg thead
  let h = '<thead><tr><th class="room-th">Rum</th>';
  idxs.forEach(i => {{
    const w = WEEKS[i];
    h += '<th class="week-th"><div class="wh-inner">';
    h += '<span class="wh-date">' + w.display + '</span>';
    h += w.sail1 ? '<span class="wh-price">' + fmt(w.sail1) + ' kr</span>' : '<span class="wh-price">&nbsp;</span>';
    h += '</div></th>';
  }});
  h += '</tr></thead><tbody>';

  // Bygg tbody sektionsvis
  SECTIONS.forEach(sec => {{
    // Filtrera rum
    const sr = sec.rooms.filter(name => {{
      if (roomQ && !name.toLowerCase().includes(roomQ)) return false;
      const r = ROOMS.find(x => x.name===name);
      if (!r) return false;
      if (statQ === 'available' && !idxs.some(i => r.weeks[i] === 'available')) return false;
      return true;
    }});
    if (!sr.length) return;

    // Sektionsrad
    h += '<tr class="sec-row"><td colspan="' + (idxs.length+1) + '">' + sec.label + '</td></tr>';

    sr.forEach(name => {{
      const r = ROOMS.find(x => x.name===name);
      if (!r) return;
      const info = ROOM_INFO[name] || {{}};
      const til = info.t || 0;
      const tilStr = til > 0 ? '+' + til.toLocaleString('sv-SE') + ' kr' : '';

      // Tooltip med alla priser
      const w0 = WEEKS[idxs[0]];
      const tipLines = [];
      if (w0) {{
        if (w0.sail1) tipLines.push('Sailing 1v: ' + fmt(w0.sail1 + til) + ' kr');
        if (w0.faw1)  tipLines.push('FAW 1v: ' + fmt(w0.faw1 + til) + ' kr');
        if (w0.ho1)   tipLines.push('HO 1v: ' + fmt(w0.ho1 + til) + ' kr');
      }}
      const tip = tipLines.join(' | ');

      h += '<tr>';
      h += '<td class="room-name-cell" title="' + tip + '">';
      h += '<div class="rn">' + name + '</div>';
      if (info.desc || tilStr) {{
        h += '<div class="rt">' + (info.desc||'');
        if (tilStr) h += ' <span class="rt-sup">tillägg ' + tilStr + '</span>';
        h += '</div>';
      }}
      h += '</td>';

      idxs.forEach(i => {{
        const st = r.weeks[i] || 'available';
        const lbl = st==='available' ? '✓' : st==='on_hold' ? '~' : '✕';
        const tip2 = WEEKS[i].display + ': ' + (st==='available'?'Ledig':st==='on_hold'?'Preliminär':'Bokad');
        const clickAttr = st==='available' ? ` onclick="openPopup('${{name}}',${{i}})" style="cursor:pointer;"` : '';
        h += '<td class="cell"><div class="ci ci-' + st + '"' + clickAttr + ' title="' + tip2 + '">' + lbl + '</div></td>';
      }});
      h += '</tr>';
    }});
  }});

  h += '</tbody>';
  document.getElementById('tbl').innerHTML = h;
}}

render();

function openPopup(name, wi) {{
  const w    = WEEKS[wi];
  const info = ROOM_INFO[name] || {{}};
  const til  = info.t || 0;

  document.getElementById('p-room').textContent = name;
  document.getElementById('p-week').textContent = 'Vecka ' + w.display + (w.sail1 ? ' · Sailing fr. ' + w.sail1.toLocaleString('sv-SE') + ' kr' : '');
  document.getElementById('p-desc').textContent = info.desc || '';

  const progs = [
    ['Sailing 1 vecka', w.sail1], ['Sailing 2 veckor', w.sail2],
    ['FAW 1 vecka', w.faw1],      ['FAW 2 veckor', w.faw2],
    ['Healthy Options 1v', w.ho1],['Healthy Options 2v', w.ho2],
  ];
  let ph = '';
  progs.forEach(([lbl, base], idx) => {{
    if (!base) return;
    const total = base;
    const hi = idx === 0 ? ' highlight' : '';
    ph += '<div class="pp-row' + hi + '">';
    ph += '<div><div class="pp-label">' + lbl + '</div>';

    ph += '</div>';
    ph += '<div class="pp-val">' + total.toLocaleString('sv-SE') + ' kr</div>';
    ph += '</div>';
  }});
  if (til) {{
    ph += '<div class="pp-row"><div class="pp-label">Rumstillägg (för hela rummet)</div><div class="pp-val">+' + til.toLocaleString('sv-SE') + ' kr/v</div></div>';
  }}
  if (w.single) {{
    ph += '<div class="pp-row"><div class="pp-label">Enkelrumstillägg</div><div class="pp-val">+' + w.single.toLocaleString('sv-SE') + ' kr/v</div></div>';
  }}
  document.getElementById('p-prices').innerHTML = ph;

  document.getElementById('p-book').href = 'boka.html?vecka=' + encodeURIComponent(w.display) + '&rum=' + encodeURIComponent(name) + '#steg1';
  document.getElementById('overlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}}

function closePopup(e) {{
  if (e && e.target !== document.getElementById('overlay') && !e.target.classList.contains('popup-close')) return;
  document.getElementById('overlay').classList.remove('open');
  document.body.style.overflow = '';
}}

document.addEventListener('keydown', e => {{ if (e.key === 'Escape') closePopup({{target:document.getElementById('overlay')}}); }});
</script>
</body>
</html>"""

def main():
    print("Wildwind Availability Updater")
    try:
        data = parse_availability(download_excel())
        Path(OUTPUT_FILE).write_text(generate_html(data), encoding='utf-8')
        print(f"Genererade {OUTPUT_FILE}")
        # Samma data som JSON, sa bokningsraknaren kan filtrera rum per vecka
        Path("availability.json").write_text(json.dumps({
            "updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "weeks": data["weeks"],
            "rooms": data["rooms"],
            "room_info": ROOM_INFO,
            "sections": SECTIONS,
        }, ensure_ascii=False), encoding='utf-8')
        print("Genererade availability.json")
        Path("temp_availability.xlsx").unlink(missing_ok=True)
        print("Klar!")
    except Exception as e:
        print(f"Fel: {e}"); raise

if __name__ == "__main__":
    main()
