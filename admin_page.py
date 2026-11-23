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


def build_bookings(df, saturdays, row_to_name, combined):
    """Per rum och vecka: vem som står bokad (första namnet under veckan)."""
    rows_for = {}
    for row, name in row_to_name.items():
        rows_for.setdefault(name, [])
    for name in rows_for:
        if name in combined:
            rows_for[name] = list(combined[name])
        else:
            rows_for[name] = [r for r, n in row_to_name.items() if n == name][:1]

    def cell(r, c):
        if r < 0 or r >= len(df) or c < 0 or c >= len(df.columns):
            return ""
        return _txt(df.iloc[r, c])

    out = {}
    for name, rows in rows_for.items():
        weeks = []
        for sat in saturdays:
            found = []
            for row in rows:
                ri = row - 1
                for d in range(7):
                    c = sat["col"] + d * 2
                    v = cell(ri, c)
                    if v and not _is_num(v):
                        note = cell(ri + 1, c)
                        pax = cell(ri + 1, c + 1)
                        found.append({
                            "n": v,
                            "x": note if note and not _is_num(note) else "",
                            "p": (pax[:-2] if pax.endswith(".0") else pax) if _is_num(pax) else "",
                            "s": cell(ri + 2, c),
                            "l": (cell(ri + 3, c) or "uk").lower(),
                            "d": d,
                        })
                        break
            if not found and weeks and weeks[-1]:
                # inget namn men rummet är upptaget: samma gäst stannar kvar
                upptaget = any(cell(row - 1, sat["col"] + d * 2) for row in rows for d in range(7))
                if upptaget:
                    found = [dict(b, f=1) for b in weeks[-1]]
            weeks.append(found or None)
        out[name] = weeks
    return out


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
    payload = json.dumps({
        "updated": updated,
        "weeks": [{"display": w["display"], "date": w["date"]} for w in data["weeks"]],
        "sections": sections,
        "room_info": {k: v.get("desc", "") for k, v in room_info.items()},
        "bookings": bookings,
    }, ensure_ascii=False)
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
td.c{width:64px;min-width:64px;height:34px;border:1px solid #f0ebe0;padding:0;cursor:default}
td.c div{height:100%;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;overflow:hidden;white-space:nowrap;padding:0 2px}
td.c.b{cursor:pointer}
td.c.hold div{background-image:repeating-linear-gradient(45deg,rgba(255,255,255,.7) 0 5px,transparent 5px 10px);outline:2px dashed rgba(0,0,0,.35);outline-offset:-3px}
td.c.dim{opacity:.18}
.rn{font-weight:700;color:var(--azure);font-size:12px}.rd{font-size:10px;color:var(--mid)}
#info{position:fixed;right:16px;bottom:16px;max-width:340px;background:#fff;border:1px solid var(--sand2);border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.15);padding:16px;display:none;z-index:20}
#info h3{color:var(--azure);margin-bottom:6px}
#info .x{position:absolute;top:8px;right:12px;cursor:pointer;border:0;background:none;font-size:18px}
.ark iframe{width:100%;height:520px;border:1px solid #ece7dc;border-radius:8px;background:#fff}
.ark a{font-size:12px;color:var(--azure)}
</style></head><body>
<header><h1>Pias översikt 2027</h1><span class="upd" id="upd"></span>
<div class="lankar">
<a href="__STEPH__" target="_blank" rel="noopener">Stephs Excel</a>
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
<label><input type="checkbox" id="baraSe"> Visa bara svenska</label></div>
<div class="ruta"><table id="tbl"></table></div>
<p style="font-size:11px;color:var(--mid);margin-top:6px">Håll musen över en ruta för namn, klicka för detaljer. Hämtat ur Stephs lista varje morgon.</p></div>
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
function beskriv(b){ return `${b.n}${b.x?' · '+b.x:''}${b.p?' · '+b.p+' pers':''}${b.s?' · '+b.s:''} · ${LAND[b.l]||b.l.toUpperCase()}`; }
function rita(){
  const bara = document.getElementById('baraSe').checked;
  let h = '<thead><tr><th class="rum">Rum</th>' + D.weeks.map(w=>`<th>${esc(w.display)}</th>`).join('') + '</tr></thead><tbody>';
  D.sections.forEach(sec=>{
    h += `<tr class="sek"><td colspan="${D.weeks.length+1}">${esc(sec.label)}</td></tr>`;
    sec.rooms.forEach(r=>{
      h += `<tr><td class="rum"><div class="rn">${esc(r)}</div><div class="rd">${esc(D.room_info[r]||'')}</div></td>`;
      (D.bookings[r]||[]).forEach((lst,i)=>{
        if(!lst){ h += '<td class="c"></td>'; return; }
        const b = lst[0], hold = /hold|option|prel/i.test(b.s||'');
        const dim = bara && !lst.some(x=>x.l==='se');
        const f = FARG[b.l] || 'var(--ovr)';
        const kort = esc(b.n.replace(/\s*x\s*\d+\s*$/i,'').split(/\s+/)[0].slice(0,9));
        h += `<td class="c b${hold?' hold':''}${dim?' dim':''}" title="${esc(lst.map(beskriv).join('\n'))}" data-r="${esc(r)}" data-i="${i}"><div style="background:${f}">${kort}</div></td>`;
      });
      h += '</tr>';
    });
  });
  document.getElementById('tbl').innerHTML = h + '</tbody>';
}
document.getElementById('tbl').addEventListener('click', e=>{
  const td = e.target.closest('td.b'); if(!td) return;
  const r = td.dataset.r, i = +td.dataset.i, lst = D.bookings[r][i];
  // hur länge stannar de? räkna sammanhängande veckor med samma namn
  const namn = lst[0].n; let a=i, z=i;
  while(a>0 && (D.bookings[r][a-1]||[]).some(x=>x.n===namn)) a--;
  while(z<D.weeks.length-1 && (D.bookings[r][z+1]||[]).some(x=>x.n===namn)) z++;
  document.getElementById('infoTxt').innerHTML = `<h3>${esc(r)} · vecka ${esc(D.weeks[i].display)}</h3>` +
    lst.map(b=>`<p style="margin:6px 0"><b>${esc(b.n)}</b>${b.x?'<br>'+esc(b.x):''}<br>${b.p?esc(b.p)+' pers · ':''}${esc(b.s||'')} · ${esc(LAND[b.l]||b.l.toUpperCase())}</p>`).join('') +
    `<p style="font-size:12px;color:var(--mid)">${esc(namn)}: ankomst ${esc(D.weeks[a].display)}, ${z-a+1} ${z>a?'veckor':'vecka'}</p>`;
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
