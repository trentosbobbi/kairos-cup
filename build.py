#!/usr/bin/env python3
"""
Kairos Cup 2025 — Build Script (ODS)
Legge KairosCup_Dati.ods, calcola tutto in Python, genera docs/index.html.

Uso: python build.py
Dipendenze: pip install odfpy
"""

import json, sys
from pathlib import Path
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf import teletype

ROOT     = Path(__file__).parent
ODS      = ROOT / "KairosCup_Dati.ods"
TEMPLATE = ROOT / "index_static.html"
OUT_DIR  = ROOT / "docs"
OUT_HTML = OUT_DIR / "index.html"

ALL_TEAMS = [
    "Torino 1","Torino 2","Mestre 1","Venezia/Mestre",
    "Merano","Pordenone","San Giorgio di Nogaro",
    "Udine","Trieste 1","Trieste 2"
]

# ── ODS reader ────────────────────────────────────────────────────────────────
def expand_sheet(sheet, max_rows=300, max_cols=10):
    """Espande celle/righe ripetute e restituisce una lista di liste."""
    result = []
    for tr in sheet.getElementsByType(TableRow):
        row_rep = int(tr.getAttribute("numberrowsrepeated") or 1)
        cols = []
        for cell in tr.getElementsByType(TableCell):
            col_rep = int(cell.getAttribute("numbercolumnsrepeated") or 1)
            num = cell.getAttribute("value")
            txt = teletype.extractText(cell).strip()
            if num is not None:
                try: v = int(float(num))
                except: v = num
            else:
                v = txt if txt else None
            cols.extend([v] * col_rep)
        for _ in range(min(row_rep, max_rows - len(result))):
            result.append(cols[:max_cols])
        if len(result) >= max_rows:
            break
    return result

def safe_int(val, default=0):
    if val is None or val == "": return default
    try: return int(float(str(val)))
    except: return default

def safe_str(val):
    return str(val).strip() if val is not None else ""

# ── Calcola classifica da PARTITE ─────────────────────────────────────────────
def calc_classifica(rows_partite):
    stats = {t:{"G":0,"V":0,"N":0,"P":0,"GF":0,"GS":0} for t in ALL_TEAMS}

    # Riga 0=titolo, 1=section, 2=header → dati da riga 3
    # Leggi solo righe con giornata "G1"…"G9" (non "FIN")
    for row in rows_partite[3:]:
        if len(row) < 7: continue
        giornata = safe_str(row[1])
        if not giornata.startswith("G"): continue
        ta = safe_str(row[2]); tb = safe_str(row[6])
        try:
            ga = int(float(str(row[3]))); gb = int(float(str(row[5])))
        except: continue
        if ta not in stats or tb not in stats: continue
        stats[ta]["G"]+=1; stats[tb]["G"]+=1
        stats[ta]["GF"]+=ga; stats[ta]["GS"]+=gb
        stats[tb]["GF"]+=gb; stats[tb]["GS"]+=ga
        if   ga>gb: stats[ta]["V"]+=1; stats[tb]["P"]+=1
        elif ga<gb: stats[tb]["V"]+=1; stats[ta]["P"]+=1
        else:       stats[ta]["N"]+=1; stats[tb]["N"]+=1

    result = []
    for team, s in stats.items():
        pt = s["V"]*3 + s["N"]; dr = s["GF"]-s["GS"]
        result.append({**s, "squadra":team, "Pt":pt, "Dr":dr})
    result.sort(key=lambda t: (-t["Pt"], -t["Dr"], -t["GF"]))
    for i,r in enumerate(result): r["pos"] = i+1
    return result

# ── Calcola cannonieri/portieri da GIOCATORI ──────────────────────────────────
def calc_giocatori(rows_giocatori):
    players = []
    # Riga 0=titolo, 1=header → dati da riga 2
    for row in rows_giocatori[2:]:
        if len(row) < 9: continue
        nome = safe_str(row[2])
        if not nome: continue
        players.append({
            "nome":    nome,
            "squadra": safe_str(row[3]),
            "gol":     safe_int(row[4]),
            "assist":  safe_int(row[5]),
            "voto":    safe_int(row[6]),
            "presenze":safe_int(row[7]),
            "ruolo":   safe_str(row[8]),
        })

    cannonieri = sorted([p for p in players if p["gol"]>0],
                        key=lambda p: (-p["gol"],-p["assist"]))[:10]
    portieri   = sorted([p for p in players
                         if p["ruolo"].lower()=="portiere" and p["voto"]>0],
                        key=lambda p: -p["voto"])[:5]
    return (
        [{"nome":p["nome"],"squadra":p["squadra"],
          "gol":p["gol"],"assist":p["assist"],"voto":p["voto"]} for p in cannonieri],
        [{"nome":p["nome"],"squadra":p["squadra"],"voto":p["voto"]} for p in portieri]
    )

# ── Leggi finali da PARTITE ───────────────────────────────────────────────────
def calc_finali(rows_partite):
    labels = ["Finale 1°/2°","Finale 3°/4°","Finale 5°/6°","Finale 7°/8°","Finale 9°/10°"]
    fin_rows = [r for r in rows_partite if len(r)>1 and safe_str(r[1])=="FIN"]
    finali = []
    for i, row in enumerate(fin_rows[:5]):
        ta = safe_str(row[2]); tb = safe_str(row[6] if len(row)>6 else "")
        try: score = f"{int(float(str(row[3])))}-{int(float(str(row[5])))}"
        except: score = None
        finali.append({
            "label": labels[i] if i<len(labels) else f"Finale {i+1}",
            "teamA": ta if ta and "TBD" not in ta else "TBD",
            "teamB": tb if tb and "TBD" not in tb else "TBD",
            "score": score,
        })
    return finali

# ── Main ──────────────────────────────────────────────────────────────────────
def build():
    print(f"📂  Leggo  {ODS.name}  …")
    if not ODS.exists():      sys.exit(f"❌  File non trovato: {ODS}")
    if not TEMPLATE.exists(): sys.exit(f"❌  Template non trovato: {TEMPLATE}")

    try:
        doc = load(str(ODS))
    except Exception as e:
        sys.exit(f"❌  Impossibile aprire il file ODS: {e}")

    sheets = {s.getAttribute("name"): s
              for s in doc.spreadsheet.getElementsByType(Table)}

    for name in ("PARTITE","GIOCATORI"):
        if name not in sheets:
            sys.exit(f"❌  Foglio '{name}' non trovato nel file ODS.")

    rows_p = expand_sheet(sheets["PARTITE"])
    rows_g = expand_sheet(sheets["GIOCATORI"])

    classifica           = calc_classifica(rows_p)
    cannonieri, portieri = calc_giocatori(rows_g)
    finali               = calc_finali(rows_p)

    data = {"classifica":classifica,"cannonieri":cannonieri,
            "portieri":portieri,"finali":finali}

    print(f"✅  Classifica: {len(classifica)} squadre")
    print(f"✅  Cannonieri: {len(cannonieri)} giocatori")
    print(f"✅  Portieri:   {len(portieri)} portieri")
    print(f"✅  Finali:     {len(finali)} partite")

    template = TEMPLATE.read_text(encoding="utf-8")
    output   = template.replace("__KAIROS_DATA__", json.dumps(data, ensure_ascii=False))
    if "__KAIROS_DATA__" in output:
        sys.exit("❌  Placeholder __KAIROS_DATA__ non trovato in index_template.html")

    OUT_DIR.mkdir(exist_ok=True)
    OUT_HTML.write_text(output, encoding="utf-8")
    print(f"\n🚀  Sito generato in:  {OUT_HTML.relative_to(ROOT)}")

if __name__ == "__main__":
    build()
