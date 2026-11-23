"""Pias adminsida (pia.html): Stephs bokningar med namn, krypterad med lösenord.

Namnen läggs ALDRIG ut i klartext. Data krypteras med AES-GCM; nyckeln tas fram
ur lösenordet (PBKDF2-SHA256). Lösenordet ligger som GitHub-secret PIA_PASSWORD.
Utan secret skapas ingen sida.
"""
import base64, json, os, re
from pathlib import Path

STEPH_URL = "https://www.dropbox.com/scl/fi/00tqgjw04ph26l50cltod/STEPH-VERSION-ALL-ROOMS-2027.xlsx?rlkey=uaw1ngacjal6m6ffs4izq25l2&dl=0"
FORFRAGNINGAR_ID = "14Dj0ZVfxc_Cc_klXddfnPuTGM2k0PnRxeCPpMNlwdhw"
BOKADE_ID = "1dH4XJ4YGmmDhhe1UUCN1DvMCsPQ29XxdZ-hJwtLGC-4"
ITERATIONS = 250000


def _txt(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("", "nan", "none") else s


def _is_num(s):
    return bool(re.fullmatch(r"-?\d+(\.\d+)?", s))


SKIP_LABELS = ("SAILING ONLY", "HEALTHY OPTIONS ONLY")
LAND_KODER = ("se", "de", "nl", "uk", "dir", "be", "hu", "it")
PREFIX = {"MELAS": "Melas", "KAVADIAS": "Kav", "AKTI": "Akti", "XRISTINA": "Xristina",
          "KALIAS": "Kalias", "ODYSSION": "Odyssion", "TOWER": "Tower villa"}
SEKTION = [("Melas", "Melas Hotel"), ("Kav", "Kavadias Hotel"), ("Akti", "AKTI Apartments"),
           ("Xristina", "Xristina Apartments"), ("X Bungalow", "Xristina Apartments"),
           ("NM", "New Melas")]


def _rumsnamn(a, b):
    """Rumsnamn ur kolumn A och B i Stephs lista, t.ex. MELAS + 7A -> Melas 7A."""
    a = a.replace("\u039d", "N").replace("\u039c", "M")  # grekiska NM
    a1 = a.split()[0].upper() if a else ""
    nr = b.split()[0] if b else ""
    if a1.startswith("NM"):
        return a.replace(" ", "").upper()
    if a1 == "XRISTINA" and nr.upper() == "BUNGALOW":
        return "X Bungalow"
    if a1 == "XRISTINA" and "STUDIOS" in a.upper():
        return "Xristina studios"
    if a1 == "MELAS" and len(a.split()) > 1:  # "MELAS 18"
        return "Melas " + a.split()[1]
    base = PREFIX.get(a1, a.title())
    if a1 in ("KALIAS", "ODYSSION", "TOWER"):
        return base
    return f"{base} {nr}".strip()


def _sektion(namn):
    if namn == "Xristina studios":
        return "Övriga (vid behov)"
    for pre, label in SEKTION:
        if namn.startswith(pre):
            return label
    return "Övriga (vid behov)"


def build_bookings(df, saturdays, row_to_name, combined):
    """Alla rum i Stephs lista med varje vistelse (ankomst och avresa per dag).

    Pias egna rum (row_to_name) markeras mine=True; övriga visas nertonade på sidan.
    """
    dagkol = [c for c in range(3, len(df.columns))
              if hasattr(df.iloc[0, c], "strftime") and _txt(df.iloc[0, c])]
    dagar = [df.iloc[0, c].strftime("%Y-%m-%d") for c in dagkol]
    kol_index = {c: i for i, c in enumerate(dagkol)}

    def cell(r, c):
        if r < 0 or r >= len(df) or c < 0 or c >= len(df.columns):
            return ""
        return _txt(df.iloc[r, c])

    mina_rader = {}
    for row, name in row_to_name.items():
        mina_rader[row] = name
    for name, rows in combined.items():
        for row in rows:
            mina_rader[row] = name

    # rumslistan börjar vid första MELAS-raden (ovanför ligger summeringar)
    forsta = next((ri for ri in range(len(df)) if cell(ri, 0).upper().startswith("MELAS")), 0)
    rum, ordning, sett = {}, [], set()
    for ri in range(forsta, len(df)):
        if cell(ri, 2) != "W" or not cell(ri, 0):
            continue
        if any(cell(ri, 0).upper().startswith(x) for x in SKIP_LABELS):
            continue
        row = ri + 1
        mine = row in mina_rader
        namn = mina_rader.get(row) or _rumsnamn(cell(ri, 0), cell(ri, 1))
        if namn in sett and not mine:
            n = 2
            while f"{namn} {n}" in sett:
                n += 1
            namn = f"{namn} {n}"
        if namn not in rum:
            desc = " · ".join(x for x in (cell(ri + 1, 0), cell(ri + 1, 1)) if x)
            rum[namn] = {"mine": mine, "rows": [], "desc": desc.capitalize()}
            ordning.append(namn)
            sett.add(namn)
        rum[namn]["rows"].append(ri)

    ut = {}
    for namn in ordning:
        vist = []
        for ri in rum[namn]["rows"]:
            akt = None
            for c in dagkol:
                v = cell(ri, c)
                if v and _is_num(v):
                    v = ""
                if akt and v == akt["n"] and kol_index[c] == akt["b"]:
                    akt["b"] += 1
                    continue
                if akt:
                    vist.append(akt); akt = None
                if v:
                    note = cell(ri + 1, c)
                    pax = cell(ri + 1, c + 1)
                    land = cell(ri + 3, c).lower()
                    if land and land not in LAND_KODER:
                        note = " · ".join(x for x in (note, cell(ri + 3, c)) if x and not _is_num(x))
                        land = ""
                    akt = {"n": v,
                           "x": note if note and not _is_num(note) else "",
                           "p": (pax[:-2] if pax.endswith(".0") else pax) if _is_num(pax) else "",
                           "s": cell(ri + 2, c),
                           "l": land or "uk",
                           "a": kol_index[c], "b": kol_index[c] + 1}
            if akt:
                vist.append(akt)
        # samma gäst i båda raderna av ett kombinerat rum ska bara synas en gång
        unika = {}
        for s in vist:
            unika.setdefault((s["n"], s["a"], s["b"]), s)
        ut[namn] = {"mine": rum[namn]["mine"], "desc": rum[namn]["desc"],
                    "stays": sorted(unika.values(), key=lambda s: s["a"])}

    sektioner = {}
    for namn in ordning:
        sektioner.setdefault(_sektion(namn), []).append(namn)
    sektioner = [{"label": k, "rooms": v} for k, v in sektioner.items()]
    veckor = [{"display": w["display"], "date": w["date"],
               "d": dagar.index(w["date"]) if w["date"] in dagar else -1} for w in saturdays]
    return {"days": dagar, "weeks": veckor, "rooms": ut, "sections": sektioner}


def _encrypt(plaintext, password):
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    salt, iv = os.urandom(16), os.urandom(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt,
                     iterations=ITERATIONS).derive(password.encode("utf-8"))
    ct = AESGCM(key).encrypt(iv, plaintext.encode("utf-8"), None)
    b = lambda x: base64.b64encode(x).decode()
    return {"salt": b(salt), "iv": b(iv), "ct": b(ct), "it": ITERATIONS}


def write_admin_page(data, bookings, sections, room_info, updated, path="pia.html"):
    pw = os.environ.get("PIA_PASSWORD", "")
    if not pw:
        print("PIA_PASSWORD saknas: pia.html skapas inte")
        return False
    room_desc = {k: v.get("desc", "") for k, v in room_info.items()}
    for namn, r in bookings["rooms"].items():
        if namn in room_desc:
            r["desc"] = room_desc[namn]
    payload = json.dumps(dict(bookings, updated=updated), ensure_ascii=False)
    enc = _encrypt(payload, pw)
    html = TEMPLATE.replace("__ENC__", json.dumps(enc)) \
                   .replace("__STEPH__", STEPH_URL) \
                   .replace("__FORF__", FORFRAGNINGAR_ID) \
                   .replace("__BOKADE__", BOKADE_ID)
    Path(path).write_text(html, encoding="utf-8")
    print(f"Genererade {path} (krypterad)")
    return True


TEMPLATE = r"""<!DOCTYPE html>
<html lang="sv"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>Pias översikt 2027</title>
<style>
:root{--azure:#0B3D72;--sand:#F5EDD8;--sand2:#E6D5BA;--text:#2A2A2A;--mid:#666;
--se:#9B59B6;--de:#F4D03F;--nl:#F39C12;--uk:#5DADE2;--dir:#AAB7B8;--ovr:#58D68D;}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,Helvetica,sans-serif;background:#FDFAF5;color:var(--text);font-size:13px}
header{background:var(--azure);color:#fff;padding:14px 20px;display:flex;gap:12px;align-items:center;flex-wrap:wrap}
header h1{font-size:20px;font-weight:600}
header .upd{opacity:.7;font-size:12px}
header .lankar{margin-left:auto;display:flex;gap:8px;flex-wrap:wrap}
header .lankar a{color:#fff;border:1px solid rgba(255,255,255,.4);border-radius:20px;padding:5px 12px;text-decoration:none;font-size:12px}
#las{max-width:380px;margin:80px auto;background:#fff;border:1px solid var(--sand2);border-radius:14px;padding:28px}
#las h2{color:var(--azure);margin-bottom:12px}
#las input[type=password]{width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:9px;font-size:15px;margin:8px 0}
#las button{width:100%;padding:12px;border:0;border-radius:30px;background:var(--azure);color:#fff;font-weight:700;font-size:14px;cursor:pointer;margin-top:8px}
#las .fel{color:#b04a3d;margin-top:8px;display:none}
.sektion{padding:16px 20px}
.sektion h2{color:var(--azure);font-size:17px;margin-bottom:8px}
.legend{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:10px;align-items:center}
.legend span{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:600}
.legend i{width:14px;height:14px;border-radius:3px;display:inline-block}
.legend label{margin-left:auto;font-size:12px;cursor:pointer}
.ruta{overflow:auto;max-height:75vh;border:1px solid #ece7dc;border-radius:8px;background:#fff}
table{border-collapse:collapse}
th.rum,td.rum{position:sticky;left:0;background:#fff;z-index:2;text-align:left;padding:4px 8px;min-width:120px;border-right:2px solid var(--sand2)}
th{position:sticky;top:0;background:var(--azure);color:#fff;font-size:11px;padding:6px 4px;z-index:3;white-space:nowrap}
th.rum{z-index:4;background:var(--azure);color:#fff}
tr.sek td{background:var(--azure);color:#fff;font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;padding:4px 8px}
td.c{width:70px;min-width:70px;height:34px;border:1px solid #f0ebe0;padding:0;cursor:default;position:relative}
td.c.b{cursor:pointer}
td.c .dgr{display:flex;height:100%}
td.c .dgr i{flex:1;display:block}
td.c .dgr i.h{background-image:repeating-linear-gradient(45deg,rgba(255,255,255,.75) 0 4px,transparent 4px 8px)!important}
td.c .lbl{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;overflow:hidden;white-space:nowrap;padding:0 2px;pointer-events:none}
td.c .lbl.del{justify-content:flex-start}
tr.annan td.rum .rn{color:#9a9f9f}
tr.annan td.rum .rd{color:#b5b5b5}
tr.annan td.c{background:#f7f7f5}
tr.annan td.c .dgr i.av{opacity:.28}
tr.annan td.c .lbl.av{color:#999}
td.c.dim .dgr,td.c.dim .lbl{opacity:.15}
.rn{font-weight:700;color:var(--azure);font-size:12px}.rd{font-size:10px;color:var(--mid)}
#info{position:fixed;right:16px;bottom:16px;max-width:340px;background:#fff;border:1px solid var(--sand2);border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.15);padding:16px;display:none;z-index:20}
#info h3{color:var(--azure);margin-bottom:6px}
#info .x{position:absolute;top:8px;right:12px;cursor:pointer;border:0;background:none;font-size:18px}
.ark iframe{width:100%;height:520px;border:1px solid #ece7dc;border-radius:8px;background:#fff}
.ark a{font-size:12px;color:var(--azure)}
</style></head><body>
<header><h1>Pias översikt 2027</h1><span class="upd" id="upd"></span>
<div class="lankar">
<a href="__STEPH__" target="_blank" rel="noopener" title="Öppnar Stephs aktuella fil i Dropbox">Stephs original i Dropbox</a>
<a href="#forf">Förfrågningar</a><a href="#bokade">Bokade</a>
<a href="boka.html" target="_blank">Boka-sidan</a></div></header>

<div id="las"><h2>Logga in</h2><p style="color:var(--mid)">Sidan visar gästernas namn och är skyddad med lösenord.</p>
<input type="password" id="pw" placeholder="Lösenord" autocomplete="current-password">
<label style="font-size:12px;display:flex;gap:6px;align-items:center"><input type="checkbox" id="minns"> Kom ihåg på den här enheten</label>
<button id="oppna">Öppna</button><div class="fel" id="fel">Fel lösenord.</div></div>

<div id="app" style="display:none">
<div class="sektion"><h2>Bokningar per rum och vecka</h2>
<div class="legend">
<span><i style="background:var(--se)"></i>Svenska</span><span><i style="background:var(--de)"></i>Tyska</span>
<span><i style="background:var(--nl)"></i>Holländska</span><span><i style="background:var(--uk)"></i>Brittiska</span>
<span><i style="background:var(--dir)"></i>Direkt</span><span><i style="background:var(--ovr)"></i>Övriga</span>
<span><i style="background:repeating-linear-gradient(45deg,#ccc 0 3px,#fff 3px 6px)"></i>On hold</span>
<span><i style="background:#f7f7f5;border:1px solid #ddd"></i>Grått rumsnamn = rum du inte säljer</span>
<label><input type="checkbox" id="baraSe"> Visa bara svenska</label></div>
<div class="ruta"><table id="tbl"></table></div>
<p style="font-size:11px;color:var(--mid);margin-top:6px">Varje ruta är en vecka lördag till fredag, delad i sju dagar, så att in- och utcheckning mitt i veckan syns. Håll musen över en ruta för namn och datum, klicka för detaljer. Hämtat ur Stephs lista varje morgon. Svenska bokningar visas alltid i full färg.</p></div>
<div class="sektion ark" id="forf"><h2>Förfrågningar från sajten</h2>
<iframe src="https://docs.google.com/spreadsheets/d/__FORF__/edit?rm=minimal" loading="lazy"></iframe>
<a href="https://docs.google.com/spreadsheets/d/__FORF__/edit" target="_blank" rel="noopener">Öppna i eget fönster</a></div>
<div class="sektion ark" id="bokade"><h2>Bokade (delas med Steph)</h2>
<iframe src="https://docs.google.com/spreadsheets/d/__BOKADE__/edit?rm=minimal" loading="lazy"></iframe>
<a href="https://docs.google.com/spreadsheets/d/__BOKADE__/edit" target="_blank" rel="noopener">Öppna i eget fönster</a></div>
</div>
<div id="info"><button class="x" onclick="this.parentNode.style.display='none'">×</button><div id="infoTxt"></div></div>

<script>
const ENC = __ENC__;
const FARG = {se:'var(--se)',de:'var(--de)',nl:'var(--nl)',uk:'var(--uk)',dir:'var(--dir)'};
const LAND = {se:'Sverige',de:'Tyskland',nl:'Nederländerna',uk:'Storbritannien',dir:'Direkt',be:'Belgien',hu:'Ungern',it:'Italien'};
const b64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
async function dekryptera(pw){
  const km = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
  const key = await crypto.subtle.deriveKey({name:'PBKDF2', salt:b64(ENC.salt), iterations:ENC.it, hash:'SHA-256'},
    km, {name:'AES-GCM', length:256}, false, ['decrypt']);
  const pt = await crypto.subtle.decrypt({name:'AES-GCM', iv:b64(ENC.iv)}, key, b64(ENC.ct));
  return JSON.parse(new TextDecoder().decode(pt));
}
const esc = s => String(s||'').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
let D = null;
const dag = i => new Date(D.days[i]+'T12:00:00');
function datum(i){
  if(i < D.days.length) return dag(i).toLocaleDateString('sv-SE',{weekday:'short',day:'numeric',month:'short'});
  const d = dag(D.days.length-1); d.setDate(d.getDate()+(i-D.days.length+1));
  return d.toLocaleDateString('sv-SE',{weekday:'short',day:'numeric',month:'short'});
}
const nätter = s => s.b - s.a;
const isHold = s => /hold|hld|option|prel/i.test(s.s||'');
function period(s){ const n = nätter(s); return `${datum(s.a)} till ${datum(s.b)} (${n} ${n===1?'natt':'nätter'})`; }
function beskriv(s){ return `${s.n} · ${period(s)}${s.x?' · '+s.x:''}${s.p?' · '+s.p+' pers':''}${s.s?' · '+s.s:''} · ${LAND[s.l]||s.l.toUpperCase()}`; }
function iVecka(r, w){ const d0 = w.d; return (D.rooms[r].stays||[]).filter(s=>s.a < d0+7 && s.b > d0); }
function rita(){
  const bara = document.getElementById('baraSe').checked;
  let h = '<thead><tr><th class="rum">Rum</th>' + D.weeks.map(w=>`<th>${esc(w.display)}</th>`).join('') + '</tr></thead><tbody>';
  D.sections.forEach(sec=>{
    h += `<tr class="sek"><td colspan="${D.weeks.length+1}">${esc(sec.label)}</td></tr>`;
    sec.rooms.forEach(r=>{
      const R = D.rooms[r], annan = !R.mine;
      h += `<tr class="${annan?'annan':''}"><td class="rum"><div class="rn">${esc(r)}</div><div class="rd">${esc(R.desc||'')}</div></td>`;
      D.weeks.forEach((w,i)=>{
        const lst = w.d < 0 ? [] : iVecka(r, w);
        if(!lst.length){ h += '<td class="c"></td>'; return; }
        const dim = bara && !lst.some(x=>x.l==='se');
        let seg = '', mest = null, mestN = 0;
        lst.forEach(s=>{ const n = Math.min(s.b, w.d+7) - Math.max(s.a, w.d); if(n > mestN){ mest = s; mestN = n; } });
        for(let k=0;k<7;k++){
          const s = lst.find(x=>x.a <= w.d+k && x.b > w.d+k);
          if(!s){ seg += '<i></i>'; continue; }
          const av = annan && s.l !== 'se';
          seg += `<i class="${isHold(s)?'h ':''}${av?'av':''}" style="background:${FARG[s.l]||'var(--ovr)'}"></i>`;
        }
        const kort = esc(mest.n.replace(/\s*x\s*\d+\s*$/i,'').split(/\s+/)[0].slice(0,9));
        const del = mestN < 5 ? ' del' : '';
        const av = annan && mest.l !== 'se' ? ' av' : '';
        h += `<td class="c b${dim?' dim':''}" title="${esc(lst.map(beskriv).join('\n'))}" data-r="${esc(r)}" data-i="${i}"><div class="dgr">${seg}</div><div class="lbl${del}${av}" style="${del && mest.a > w.d ? 'justify-content:flex-end' : ''}">${kort}</div></td>`;
      });
      h += '</tr>';
    });
  });
  document.getElementById('tbl').innerHTML = h + '</tbody>';
}
document.getElementById('tbl').addEventListener('click', e=>{
  const td = e.target.closest('td.b'); if(!td) return;
  const r = td.dataset.r, w = D.weeks[+td.dataset.i], lst = iVecka(r, w);
  document.getElementById('infoTxt').innerHTML = `<h3>${esc(r)} · vecka ${esc(w.display)}</h3>` +
    (D.rooms[r].mine ? '' : '<p style="font-size:12px;color:var(--mid)">Rum som du inte säljer</p>') +
    lst.map(s=>`<p style="margin:8px 0"><b>${esc(s.n)}</b><br>Ankomst ${esc(datum(s.a))}, avresa ${esc(datum(s.b))}<br>${nätter(s)} ${nätter(s)===1?'natt':'nätter'}${s.x?'<br>'+esc(s.x):''}<br>${s.p?esc(s.p)+' pers · ':''}${esc(s.s||'')} · ${esc(LAND[s.l]||s.l.toUpperCase())}</p>`).join('');
  document.getElementById('info').style.display = 'block';
});
document.getElementById('baraSe').addEventListener('change', rita);
async function oppna(pw){
  try{ D = await dekryptera(pw); }catch(e){ document.getElementById('fel').style.display='block'; return false; }
  document.getElementById('las').style.display='none';
  document.getElementById('app').style.display='block';
  document.getElementById('upd').textContent = 'Uppdaterad ' + D.updated;
  rita(); return true;
}
document.getElementById('oppna').onclick = async ()=>{
  const pw = document.getElementById('pw').value;
  if(await oppna(pw) && document.getElementById('minns').checked){ try{ localStorage.setItem('piaPw', pw); }catch(e){} }
};
document.getElementById('pw').addEventListener('keydown', e=>{ if(e.key==='Enter') document.getElementById('oppna').click(); });
try{ const s = localStorage.getItem('piaPw'); if(s) oppna(s); }catch(e){}
</script></body></html>
"""
