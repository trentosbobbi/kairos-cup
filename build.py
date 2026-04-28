#!/usr/bin/env python3
"""
Kairos Cup 2025 — Build Script
================================
Legge KairosCup_Dati.xlsx, estrae classifica / cannonieri / portieri
e li inietta nel template HTML, producendo docs/index.html pronto
per essere pubblicato su GitHub Pages.

Uso:
    python build.py

Dipendenze:
    pip install openpyxl
"""

import json
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("❌  Installa openpyxl:  pip install openpyxl")

# ── Percorsi ─────────────────────────────────────────────────────────────────
ROOT      = Path(__file__).parent
EXCEL     = ROOT / "KairosCup_Dati.xlsx"
TEMPLATE  = ROOT / "index_template.html"
OUT_DIR   = ROOT / "docs"
OUT_HTML  = OUT_DIR / "index.html"

ALL_TEAMS = [
    "Torino 1", "Torino 2", "Mestre 1", "Venezia/Mestre",
    "Merano", "Pordenone", "San Giorgio di Nogaro",
    "Udine", "Trieste 1", "Trieste 2"
]

def empty_data():
    return {
        "classifica": [
            {"pos": i+1, "squadra": t, "G":0,"V":0,"N":0,"P":0,
             "GF":0,"GS":0,"Dr":0,"Pt":0}
            for i, t in enumerate(ALL_TEAMS)
        ],
        "cannonieri": [],
        "portieri": []
    }

def safe_int(val, default=0):
    try:
        return int(val) if val not in (None, "") else default
    except (TypeError, ValueError):
        return default

def read_calc_sheet(wb):
    """Legge il foglio _CALC (righe 2-11) → classifica grezza."""
    ws = wb["_CALC"]
    rows = []
    for row in ws.iter_rows(min_row=2, max_row=11, values_only=True):
        squadra = row[0]
        if not squadra:
            continue
        G  = safe_int(row[1])
        V  = safe_int(row[2])
        N  = safe_int(row[3])
        P  = safe_int(row[4])
        GF = safe_int(row[5])
        GS = safe_int(row[6])
        Pt = safe_int(row[7])
        Dr = GF - GS
        rows.append({"squadra": squadra, "G":G,"V":V,"N":N,"P":P,
                     "GF":GF,"GS":GS,"Dr":Dr,"Pt":Pt})

    # Ordina: Pt desc, poi Dr desc
    rows.sort(key=lambda t: (-t["Pt"], -t["Dr"]))
    for i, r in enumerate(rows):
        r["pos"] = i + 1
    return rows

def read_giocatori_sheet(wb):
    """Legge il foglio GIOCATORI (righe 3-202) → cannonieri e portieri."""
    ws = wb["GIOCATORI"]
    players = []
    for row in ws.iter_rows(min_row=3, max_row=202, values_only=True):
        nome = row[2]  # colonna C
        if not nome:
            continue
        players.append({
            "nome":    str(nome).strip(),
            "squadra": str(row[3] or "").strip(),
            "gol":     safe_int(row[4]),
            "assist":  safe_int(row[5]),
            "voto":    safe_int(row[6]),
            "presenze":safe_int(row[7]),
            "ruolo":   str(row[8] or "").strip(),
        })

    cannonieri = sorted(
        [p for p in players if p["gol"] > 0],
        key=lambda p: (-p["gol"], -p["assist"])
    )[:10]

    portieri = sorted(
        [p for p in players if p["ruolo"].lower() == "portiere" and p["voto"] > 0],
        key=lambda p: -p["voto"]
    )[:5]

    # Formato output
    cannonieri_out = [
        {"nome": p["nome"], "squadra": p["squadra"],
         "gol": p["gol"], "assist": p["assist"], "voto": p["voto"]}
        for p in cannonieri
    ]
    portieri_out = [
        {"nome": p["nome"], "squadra": p["squadra"], "voto": p["voto"]}
        for p in portieri
    ]
    return cannonieri_out, portieri_out

def build():
    print(f"📂  Leggo  {EXCEL.name}  …")

    if not EXCEL.exists():
        sys.exit(f"❌  File non trovato: {EXCEL}\n"
                 "    Assicurati che KairosCup_Dati.xlsx sia nella stessa cartella di build.py")

    if not TEMPLATE.exists():
        sys.exit(f"❌  Template non trovato: {TEMPLATE}")

    try:
        wb = openpyxl.load_workbook(EXCEL, data_only=True)
    except Exception as e:
        sys.exit(f"❌  Impossibile aprire il file Excel: {e}")

    # Verifica fogli obbligatori
    for sheet in ("_CALC", "GIOCATORI"):
        if sheet not in wb.sheetnames:
            sys.exit(f"❌  Foglio '{sheet}' non trovato nel file Excel.")

    classifica              = read_calc_sheet(wb)
    cannonieri, portieri    = read_giocatori_sheet(wb)

    data = {
        "classifica": classifica,
        "cannonieri": cannonieri,
        "portieri":   portieri,
    }

    print(f"✅  Classifica: {len(classifica)} squadre")
    print(f"✅  Cannonieri: {len(cannonieri)} giocatori")
    print(f"✅  Portieri:   {len(portieri)} portieri")

    # Leggi template e inietta dati
    template = TEMPLATE.read_text(encoding="utf-8")
    data_json = json.dumps(data, ensure_ascii=False)
    output = template.replace("__KAIROS_DATA__", data_json)

    # Verifica che il placeholder fosse presente
    if "__KAIROS_DATA__" in template and "__KAIROS_DATA__" not in output:
        pass  # ok
    elif "__KAIROS_DATA__" in output:
        sys.exit("❌  Placeholder __KAIROS_DATA__ non sostituito — controlla index_template.html")

    # Scrivi output
    OUT_DIR.mkdir(exist_ok=True)
    OUT_HTML.write_text(output, encoding="utf-8")
    print(f"\n🚀  Sito generato in:  {OUT_HTML.relative_to(ROOT)}")
    print("    Carica su GitHub e il sito si aggiornerà automaticamente.\n")

if __name__ == "__main__":
    build()
