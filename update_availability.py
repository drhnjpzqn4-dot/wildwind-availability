#!/usr/bin/env python3
"""Wildwind Room Availability Updater – genererar availability.html"""

import json, requests, pandas as pd
from datetime import datetime
from pathlib import Path

DROPBOX_URL = "https://www.dropbox.com/scl/fi/i0ji499omsoc0yl4fuzo7/STEPH-VERSION-ALL-ROOMS-2026.xlsx?rlkey=lq3j8vz7hdlnoqvairj0yp2ih&dl=1"
OUTPUT_FILE = "availability.html"

# ── Excel-rad → visningsnamn ──────────────────────────────────────────
# M7AB borttaget (nu UK-rum). K9, Xristina 2, Xristina 4, X Bungalow tillagda.
ROW_TO_NAME = {
    44:  "Melas 1",
    48:  "Melas 2AB",   # kombineras med rad 52
    56:  "Melas 3",
    60:  "Melas 4",
    80:  "Melas 8",
    140: "Kav 2",
    144: "Kav 3",
    148: "Kav 4",
    168: "Kav 9",
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
    256: "Xristina 2",
    264: "Xristina 4",
    272: "Xristina 6",
    276: "X Bungalow",
    284: "NM208",
}

# Melas 2AB: kombinera rad 48+52 (bokad om NÅGON är bokad)
COMBINED_ROOMS = {
    "Melas 2AB": [48, 52],
}

# ── Priser från wildwind.se ───────────────────────────────────────────
# (sail1, sail2, faw1, faw2, ho1, ho2, single)
ALL_PRICES = {
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

# ── Ruminfo ───────────────────────────────────────────────────────────
# tillagg_num = per rum (2 pers), dvs kr/pers × 2
ROOM_INFO = {
    "Melas 1":    {"t": 0,    "sängar": "Twin",           "info": "Sidoutsikt mot Kav Bar · Bottenvåning"},
    "Melas 2AB":  {"t": 0,    "sängar": "Double + Twin",  "info": "Familjerum · Sid- och havsutsikt · Bottenvåning"},
    "Melas 3":    {"t": 1758, "sängar": "Double",         "info": "Havsutsikt mot stranden · Bottenvåning"},
    "Melas 4":    {"t": 1758, "sängar": "Twin",           "info": "Havsutsikt mot stranden · Bottenvåning"},
    "Melas 8":    {"t": 0,    "sängar": "Twin",           "info": "Utsikt mot Ponti och berg · Bottenvåning"},
    "Kav 2":      {"t": 0,    "sängar": "Twin",           "info": "Utsikt mot berg norrut · Bottenvåning"},
    "Kav 3":      {"t": 0,    "sängar": "Twin",           "info": "Utsikt mot berg norrut · Bottenvåning"},
    "Kav 4":      {"t": 0,    "sängar": "Double",         "info": "Utsikt mot gräsmattan österut · Bottenvåning"},
    "Kav 9":      {"t": 0,    "sängar": "Double",         "info": "Mot poolen, utan uteplats · Bottenvåning"},
    "Kav 13":     {"t": 1320, "sängar": "Twin",           "info": "Utsikt mot berg norrut · Övervåning"},
    "Kav 14":     {"t": 1320, "sängar": "Double",         "info": "Mot poolen · Övervåning"},
    "Kav 15":     {"t": 1320, "sängar": "Double",         "info": "Mot poolen · Övervåning"},
    "Kav 18":     {"t": 1320, "sängar": "Twin",           "info": "Havsutsikt · Övervåning"},
    "Kav 19":     {"t": 1320, "sängar": "Twin",           "info": "Mot byn österut · Övervåning"},
    "Akti 7":     {"t": 1758, "sängar": "Double + Twin",  "info": "2-sovrum, fullt kök · Min 4 pers · Bottenvåning"},
    "Akti 8":     {"t": 1758, "sängar": "Double + Twin",  "info": "2-sovrum, fullt kök · Min 4 pers · Bottenvåning"},
    "Akti 9":     {"t": 1758, "sängar": "Twin + extra",   "info": "1-sovrum, pentry · Bottenvåning"},
    "Akti 4":     {"t": 1320, "sängar": "Twin",           "info": "Studio, pentry · Övervåning"},
    "Akti 5":     {"t": 1320, "sängar": "Twin",           "info": "Studio, pentry · Övervåning"},
    "Xristina 2": {"t": 1320, "sängar": "Twin",           "info": "Studio, pentry · Havsutsikt på avstånd"},
    "Xristina 4": {"t": 1320, "sängar": "Twin",           "info": "Studio, pentry · Trädgårdsvy norrut"},
    "Xristina 6": {"t": 1758, "sängar": "Double + 2 enk","info": "1-sovrum, kitchenette · Trädgårdsvy · 13 jul–15 sep"},
    "X Bungalow": {"t": 1320, "sängar": "Double + 2 enk","info": "Eget hus i trädgården · Mot bergen norrut"},
    "NM208":      {"t": 2638, "sängar": "Double",         "info": "Bergutsikt bakåt · Övervåning · New Melas"},
}

SECTIONS = [
    {"label": "Melas Hotel",         "rooms": ["Melas 1","Melas 2AB","Melas 3","Melas 4","Melas 8"]},
    {"label": "Kavadias Hotel",      "rooms": ["Kav 2","Kav 3","Kav 4","Kav 9","Kav 13","Kav 14","Kav 15","Kav 18","Kav 19"]},
    {"label": "AKTI Apartments",     "rooms": ["Akti 4","Akti 5","Akti 7","Akti 8","Akti 9"]},
    {"label": "Xristina Apartments", "rooms": ["Xristina 2","Xristina 4","Xristina 6","X Bungalow"]},
    {"label": "New Melas",           "rooms": ["NM208"]},
]

# ── Download ──────────────────────────────────────────────────────────
def download_excel():
    print("📥 Laddar ner Excel...")
    r = requests.get(DROPBOX_URL, timeout=30)
    r.raise_for_status()
    p = Path("temp_availability.xlsx")
    p.write_bytes(r.content)
    print(f"✅ {len(r.content):,} bytes")
    return p

# ── Parse ─────────────────────────────────────────────────────────────
def parse_availability(path):
    print("📊 Läser Excel...")
    df = pd.read_excel(path, header=None)

    # Hitta lördagskolumner (rad 2 = index 1 har "SATURDAY")
    saturdays = []
    for col in range(3, len(df.columns), 2):
        dv = df.iloc[0, col]
        dn = df.iloc[1, col]
        if pd.notna(dn) and "SATURDAY" in str(dn).upper() and pd.notna(dv):
            dt    = pd.to_datetime(dv)
            label = dt.strftime("%-d %b")
            p     = ALL_PRICES.get(label, (None,)*7)
            saturdays.append({
                "col": col, "date": dt.strftime("%Y-%m-%d"),
                "display": label,
                "sail1": p[0], "sail2": p[1],
                "faw1":  p[2], "faw2":  p[3],
                "ho1":   p[4], "ho2":   p[5],
                "single":p[6],
            })
    print(f"   {len(saturdays)} lördagar")

    def row_statuses(r_idx):
        result = []
        for sat in saturdays:
            sc = sat["col"]
            booked = on_hold = False
            for d in range(7):
                c = sc + d * 2
                if c >= len(df.columns): break
                # Bokad: värde i rumraden
                v = df.iloc[r_idx, c]
                if pd.notna(v) and str(v).strip() not in ("", "nan"):
                    booked = True
                # On hold: sök i +1, +2, +3 raderna under
                for off in [1, 2, 3]:
                    if r_idx + off < len(df):
                        h = df.iloc[r_idx + off, c]
                        if pd.notna(h) and any(
                            k in str(h).lower() for k in ("hold","option","prel")
                        ):
                            on_hold = True
            result.append(
                "booked" if booked else "on_hold" if on_hold else "available"
            )
        return result

    rooms, added = [], set()
    for row, name in ROW_TO_NAME.items():
        if name in added: continue
        r_idx = row - 1
        if r_idx >= len(df):
            print(f"   ⚠️ Rad {row} ({name}) saknas")
            continue
        if name in COMBINED_ROOMS:
            all_st = [row_statuses(r - 1) for r in COMBINED_ROOMS[name]]
            ws = []
            for i in range(len(saturdays)):
                col = [s[i] for s in all_st]
                ws.append("booked" if "booked" in col else "on_hold" if "on_hold" in col else "available")
        else:
            ws = row_statuses(r_idx)
        rooms.append({"name": name, "weeks": ws})
        added.add(name)

    print(f"   {len(rooms)} rum klara")
    return {"weeks": saturdays, "rooms": rooms}

# ── Generate HTML ─────────────────────────────────────────────────────
def generate_html(data):
    weeks = data["weeks"]
    rooms = data["rooms"]
    now   = datetime.now().strftime("%d %b %Y, %H:%M")

    djs  = f"const WEEKS={json.dumps(weeks, ensure_ascii=False)};\n"
    djs += f"const ROOMS={json.dumps(rooms, ensure_ascii=False)};\n"
    djs += f"const ROOM_INFO={json.dumps(ROOM_INFO, ensure_ascii=False)};\n"
    djs += f"const SECTIONS={json.dumps(SECTIONS, ensure_ascii=False)};\n"
    djs += f"const UPDATED='{now}';\n"

    html = '''<!DOCTYPE html>
<html lang="sv">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Wildwind – Rumstillgänglighet 2026</title>
<link href="https://fonts.googleapis.com/css2?family=Nunito+Sans:wght@300;400;600;700&family=Cormorant+Garamond:ital,wght@0,400;0,600;1,400&display=swap" rel="stylesheet"/>
<style>
:root{
  --azure:#0B3D72;--azure2:#1A5C9A;--sky:#5BA4CF;--sand:#F5EDD8;--sand2:#E6D5BA;
  --green:#2E8B57;--orange:#E08C2A;--red:#C0392B;--gold:#C9963A;
  --green-bg:#E8F5EE;--orange-bg:#FEF3E2;--red-bg:#FDECEC;
  --white:#FDFAF5;--text:#2A2A2A;--mid:#666;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Nunito Sans',sans-serif;background:var(--white);color:var(--text);font-size:14px;}

header{background:var(--azure);color:#fff;padding:14px 20px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;}
header h1{font-family:'Cormorant Garamond',serif;font-size:22px;font-weight:400;}
.hright{display:flex;align-items:center;gap:12px;}
.upd{font-size:11px;opacity:.6;}
.reload{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);color:#fff;padding:6px 14px;border-radius:20px;cursor:pointer;font-size:12px;font-family:inherit;font-weight:600;}
.reload:hover{background:rgba(255,255,255,.28);}

.controls{background:var(--sand);padding:12px 20px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;border-bottom:2px solid var(--sand2);}
.controls label{font-size:11px;font-weight:700;color:var(--azure);text-transform:uppercase;letter-spacing:.06em;}
.pbtn{padding:7px 14px;border-radius:20px;border:1.5px solid var(--sand2);background:#fff;font-family:inherit;font-size:12px;font-weight:600;cursor:pointer;color:var(--mid);transition:all .2s;}
.pbtn.active{background:var(--azure);color:#fff;border-color:var(--azure);}
.pbtns{display:flex;gap:6px;flex-wrap:wrap;}
.show-sel{padding:7px 12px;border:1.5px solid var(--sand2);border-radius:20px;font-family:inherit;font-size:12px;background:#fff;cursor:pointer;}

.week-nav{background:var(--azure2);color:#fff;padding:0 20px;display:flex;align-items:stretch;}
.wtabs{display:flex;align-items:center;flex:1;overflow-x:auto;scrollbar-width:none;}
.wtabs::-webkit-scrollbar{display:none;}
.wtab{padding:12px 14px;cursor:pointer;white-space:nowrap;font-size:12px;font-weight:600;color:rgba(255,255,255,.6);border-bottom:3px solid transparent;transition:all .2s;flex-shrink:0;display:flex;flex-direction:column;align-items:center;gap:2px;}
.wtab:hover{color:#fff;background:rgba(255,255,255,.08);}
.wtab.active{color:#fff;border-bottom-color:var(--gold);}
.wtab .wdate{font-size:13px;}
.wtab .wavail{font-size:10px;opacity:.7;}
.wtab.active .wavail{opacity:1;color:var(--gold);}
.narr{background:transparent;border:none;color:rgba(255,255,255,.7);font-size:22px;cursor:pointer;padding:0 8px;transition:color .2s;}
.narr:hover{color:#fff;}
.narr:disabled{opacity:.2;cursor:default;}

.pbanner{background:#fff;border-bottom:1px solid #e8e0d0;padding:12px 20px;display:flex;align-items:center;gap:16px;flex-wrap:wrap;}
.pbanner-date{font-family:'Cormorant Garamond',serif;font-size:20px;color:var(--azure);font-weight:600;min-width:120px;}
.pchips{display:flex;gap:8px;flex-wrap:wrap;}
.pchip{padding:5px 14px;border-radius:20px;font-size:12px;font-weight:600;background:var(--sand);color:var(--azure);border:1px solid var(--sand2);cursor:pointer;transition:all .2s;}
.pchip.active{background:var(--azure);color:#fff;border-color:var(--azure);}
.pchip.na{opacity:.35;cursor:default;}

.legend{display:flex;gap:16px;padding:8px 20px;background:#fff;border-bottom:1px solid #e8e0d0;flex-wrap:wrap;}
.li{display:flex;align-items:center;gap:6px;font-size:11px;font-weight:600;}
.dot{width:10px;height:10px;border-radius:50%;}
.dg{background:var(--green);}.do{background:var(--orange);}.dr{background:var(--red);}

.rgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));gap:12px;padding:16px 20px;}
.hsec{grid-column:1/-1;font-size:11px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--mid);padding:4px 0 6px;border-bottom:2px solid var(--sand2);}
.rcard{background:#fff;border-radius:10px;padding:14px;border:1.5px solid #e8e0d0;display:flex;flex-direction:column;gap:8px;transition:all .2s;}
.rcard:hover{border-color:var(--sky);box-shadow:0 4px 16px rgba(11,61,114,.08);}
.rcard.on_hold{border-color:var(--orange);background:var(--orange-bg);}
.rcard.booked{opacity:.38;pointer-events:none;}
.rtop{display:flex;align-items:center;gap:8px;}
.rdot{width:9px;height:9px;border-radius:50%;flex-shrink:0;}
.rdot.available{background:var(--green);}.rdot.on_hold{background:var(--orange);}.rdot.booked{background:var(--red);}
.rname{font-weight:700;font-size:15px;color:var(--azure);}
.rstatus{font-size:11px;font-weight:700;margin-left:auto;}
.sa{color:var(--green);}.so{color:var(--orange);}.sb{color:var(--red);}
.rinfo{font-size:12px;color:var(--mid);line-height:1.5;}
.rprices{background:var(--sand);border-radius:6px;padding:8px 10px;display:flex;flex-direction:column;gap:3px;}
.rprow{display:flex;justify-content:space-between;font-size:12px;}
.rplabel{color:var(--mid);}
.rpval{font-weight:700;color:var(--azure);}
.rpval.sel{color:var(--green);font-size:13px;}
.rcta{display:flex;align-items:center;justify-content:center;gap:6px;background:var(--azure);color:#fff;border-radius:6px;padding:9px;font-size:12px;font-weight:600;text-decoration:none;transition:background .2s;margin-top:2px;}
.rcta:hover{background:var(--azure2);}
.empty{grid-column:1/-1;text-align:center;padding:50px 20px;color:var(--mid);}
.empty h3{font-family:'Cormorant Garamond',serif;font-size:24px;color:var(--azure);margin-bottom:8px;}

.infobox{margin:0 20px 20px;padding:12px 16px;background:#EDF3FA;border-left:4px solid var(--sky);border-radius:6px;font-size:12px;color:var(--mid);line-height:1.6;}

@media(max-width:600px){
  .rgrid{grid-template-columns:1fr;padding:10px 12px;}
  .controls{padding:10px 12px;}
  .pbanner{padding:10px 12px;}
}
</style>
</head>
<body>

<header>
  <h1>🌊 Wildwind 2026</h1>
  <div class="hright">
    <span class="upd">Uppdaterad: <span id="upd"></span></span>
    <button class="reload" onclick="location.reload(true)">↺ Uppdatera</button>
  </div>
</header>

<div class="controls">
  <label>Program:</label>
  <div class="pbtns" id="progBtns">
    <button class="pbtn active" data-prog="sail1">Sailing 1v</button>
    <button class="pbtn" data-prog="sail2">Sailing 2v</button>
    <button class="pbtn" data-prog="faw1">FAW 1v</button>
    <button class="pbtn" data-prog="faw2">FAW 2v</button>
    <button class="pbtn" data-prog="ho1">HO 1v</button>
    <button class="pbtn" data-prog="ho2">HO 2v</button>
  </div>
  <label style="margin-left:8px;">Visa:</label>
  <select class="show-sel" id="showSel" onchange="render()">
    <option value="available">Bara lediga</option>
    <option value="all">Alla rum</option>
  </select>
</div>

<div class="week-nav">
  <button class="narr" id="navP" onclick="shift(-1)">&#8249;</button>
  <div class="wtabs" id="wtabs"></div>
  <button class="narr" id="navN" onclick="shift(1)">&#8250;</button>
</div>

<div class="pbanner" id="pbanner"></div>

<div class="legend">
  <div class="li"><div class="dot dg"></div>Ledig</div>
  <div class="li"><div class="dot do"></div>Preliminärt bokad</div>
  <div class="li"><div class="dot dr"></div>Bokad</div>
</div>

<div class="rgrid" id="rgrid"></div>
<div class="infobox">ℹ️ Rum kan vara preliminärt bokade i upp till 48h utan att synas här. Kontakta oss för aktuell status.</div>

<script>
PLACEHOLDER_DJS
document.getElementById('upd').textContent = UPDATED;

let prog = 'sail1', cw = 0, tabOff = 0;
const TABS = () => window.innerWidth < 600 ? 4 : 8;

// Starta på närmaste lördag
const today = new Date();
for (let i = 0; i < WEEKS.length; i++) {
  if (new Date(WEEKS[i].date) >= today) { cw = i; break; }
}
tabOff = Math.max(0, cw - 2);

// Program-knappar
document.getElementById('progBtns').addEventListener('click', e => {
  const b = e.target.closest('.pbtn');
  if (!b) return;
  const w = WEEKS[cw];
  const key = b.dataset.prog;
  const val = {sail1:w.sail1,sail2:w.sail2,faw1:w.faw1,faw2:w.faw2,ho1:w.ho1,ho2:w.ho2}[key];
  if (!val) return; // ej tillgänglig denna vecka
  prog = key;
  document.querySelectorAll('.pbtn').forEach(x => x.classList.remove('active'));
  b.classList.add('active');
  render();
});

function getVal(w, p) {
  return {sail1:w.sail1,sail2:w.sail2,faw1:w.faw1,faw2:w.faw2,ho1:w.ho1,ho2:w.ho2}[p] || null;
}

function countFree(wi) {
  return ROOMS.filter(r => r.weeks[wi] === 'available').length;
}

function shift(d) {
  tabOff = Math.max(0, Math.min(WEEKS.length - TABS(), tabOff + d));
  renderTabs();
}

function renderTabs() {
  const c = document.getElementById('wtabs'); c.innerHTML = '';
  const end = Math.min(WEEKS.length, tabOff + TABS());
  for (let i = tabOff; i < end; i++) {
    const w = WEEKS[i];
    const d = document.createElement('div');
    d.className = 'wtab' + (i===cw?' active':'');
    d.innerHTML = `<span class="wdate">${w.display}</span><span class="wavail">${countFree(i)} lediga</span>`;
    d.onclick = () => { cw=i; renderTabs(); renderBanner(); renderRooms(); };
    c.appendChild(d);
  }
  document.getElementById('navP').disabled = tabOff === 0;
  document.getElementById('navN').disabled = tabOff + TABS() >= WEEKS.length;
}

function renderBanner() {
  const w = WEEKS[cw];
  const ed = new Date(w.date); ed.setDate(ed.getDate()+7);
  const edl = ed.toLocaleDateString('sv-SE',{day:'numeric',month:'short'});
  const chips = [
    {k:'sail1',l:'Sailing 1v', v:w.sail1},
    {k:'sail2',l:'Sailing 2v', v:w.sail2},
    {k:'faw1', l:'FAW 1v',     v:w.faw1},
    {k:'faw2', l:'FAW 2v',     v:w.faw2},
    {k:'ho1',  l:'HO 1v',      v:w.ho1},
    {k:'ho2',  l:'HO 2v',      v:w.ho2},
  ];
  let h = `<div class="pbanner-date">${w.display} – ${edl}</div><div class="pchips">`;
  chips.forEach(c => {
    if (!c.v) { h += `<div class="pchip na">${c.l} –</div>`; return; }
    const a = c.k===prog ? ' active' : '';
    h += `<div class="pchip${a}" onclick="setProg('${c.k}')">${c.l}: ${c.v.toLocaleString('sv-SE')} kr</div>`;
  });
  if (w.single) h += `<div class="pchip" style="opacity:.6">Enkelrum +${w.single.toLocaleString('sv-SE')} kr</div>`;
  h += '</div>';
  document.getElementById('pbanner').innerHTML = h;
}

function setProg(p) {
  prog = p;
  document.querySelectorAll('.pbtn').forEach(b => {
    b.classList.toggle('active', b.dataset.prog===p);
  });
  render();
}

function renderRooms() {
  const showAll = document.getElementById('showSel').value === 'all';
  const w = WEEKS[cw];
  const base = getVal(w, prog);
  const stL = {available:'Ledig', on_hold:'Preliminär', booked:'Bokad'};
  let html = '', shown = 0;

  SECTIONS.forEach(sec => {
    const sr = sec.rooms.filter(name => {
      const r = ROOMS.find(x => x.name===name); if (!r) return false;
      return showAll || r.weeks[cw]==='available' || r.weeks[cw]==='on_hold';
    });
    if (!sr.length) return;
    shown += sr.length;
    html += `<div class="hsec">${sec.label}</div>`;
    sr.forEach(name => {
      const r = ROOMS.find(x => x.name===name); if (!r) return;
      const st = r.weeks[cw] || 'available';
      const info = ROOM_INFO[name] || {};
      const til = info.t || 0;
      const total = base ? base + til : null;
      const subj = encodeURIComponent(`Bokningsförfrågan: ${name} vecka ${w.display}`);

      let prows = '';
      [['sail1','Sailing 1v',w.sail1],['sail2','Sailing 2v',w.sail2],
       ['faw1','FAW 1v',w.faw1],['faw2','FAW 2v',w.faw2],
       ['ho1','HO 1v',w.ho1],['ho2','HO 2v',w.ho2]
      ].forEach(([k,l,v]) => {
        if (!v) return;
        const t = v + til;
        prows += `<div class="rprow"><span class="rplabel">${l}${til?' + tillägg':''}</span><span class="rpval${k===prog?' sel':''}">${t.toLocaleString('sv-SE')} kr</span></div>`;
      });
      if (til) prows += `<div class="rprow" style="border-top:1px solid #ddd;margin-top:3px;padding-top:3px;"><span class="rplabel">varav rumstillägg</span><span class="rpval">${til.toLocaleString('sv-SE')} kr</span></div>`;

      html += `<div class="rcard ${st!=='available'?st:''}">
        <div class="rtop">
          <div class="rdot ${st}"></div>
          <div class="rname">${name}</div>
          <div class="rstatus s${st[0]}">${stL[st]}</div>
        </div>
        ${info.sängar?`<div class="rinfo">${info.sängar} · ${info.info||''}</div>`:''}
        <div class="rprices">${prows}</div>
        ${st==='available'?`<a href="mailto:pia@wildwind.se?subject=${subj}" class="rcta">✉ Boka ${name}</a>`:''}
      </div>`;
    });
  });

  document.getElementById('rgrid').innerHTML = shown
    ? html
    : `<div class="empty"><h3>Inga lediga rum denna vecka</h3><p>Välj en annan vecka ovan.</p></div>`;
}

function render() { renderTabs(); renderBanner(); renderRooms(); }
render();
</script>
</body>
</html>'''
    html = html.replace('PLACEHOLDER_DJS', djs)
    return html

# ── Main ──────────────────────────────────────────────────────────────
def main():
    print("🚀 Wildwind Availability Updater"); print("="*40)
    try:
        data = parse_availability(download_excel())
        Path(OUTPUT_FILE).write_text(generate_html(data), encoding='utf-8')
        print(f"✅ {OUTPUT_FILE}")
        Path("temp_availability.xlsx").unlink(missing_ok=True)
        print("🎉 Klar!")
    except Exception as e:
        print(f"❌ {e}"); raise

if __name__ == "__main__":
    main()
