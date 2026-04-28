# Kairos Cup 2025 — Gestione Sito

Il sito si aggiorna automaticamente ogni volta che carichi un Excel aggiornato su GitHub.
I visitatori non vedono nessun pulsante di gestione — il sito è completamente statico.

---

## Struttura della cartella

```
kairos-cup/
├── KairosCup_Dati.xlsx      ← il tuo file dati (l'unico che modifichi)
├── index_template.html      ← template del sito (non modificare)
├── build.py                 ← script di build (non modificare)
├── .github/
│   └── workflows/
│       └── deploy.yml       ← automazione GitHub (non modificare)
└── docs/
    └── index.html           ← sito generato (non modificare, si crea da solo)
```

---

## Setup iniziale (una sola volta)

### 1. Crea il repository su GitHub

1. Vai su [github.com](https://github.com) → **New repository**
2. Nome: `kairos-cup` (o come preferisci)
3. Visibilità: **Public** *(GitHub Pages gratuito richiede repo pubblico)*
4. Clicca **Create repository**

### 2. Carica i file

Trascina nella pagina del repository (o usa GitHub Desktop) tutti i file:
- `KairosCup_Dati.xlsx`
- `index_template.html`
- `build.py`
- la cartella `.github/` con tutto il contenuto

### 3. Abilita GitHub Pages

1. Nel repository → **Settings** → **Pages** (menu a sinistra)
2. Source: **GitHub Actions**
3. Salva

### 4. Prima build

Vai su **Actions** → seleziona l'ultimo workflow → clicca **Re-run jobs** se necessario.
Dopo ~1 minuto il sito sarà pubblicato all'indirizzo:

```
https://TUONOME.github.io/kairos-cup/
```

---

## Flusso di aggiornamento (ogni giornata di torneo)

1. Apri **KairosCup_Dati.xlsx** sul tuo PC
2. Inserisci i gol nel foglio **PARTITE** e i dati giocatori nel foglio **GIOCATORI**
3. Salva il file
4. Carica il file aggiornato su GitHub:
   - Vai nella pagina del repository
   - Clicca sul file `KairosCup_Dati.xlsx`
   - Clicca sull'icona matita (✏️) o trascina il nuovo file
   - Clicca **Commit changes**
5. GitHub Actions legge l'Excel, genera il sito e lo pubblica in automatico (~1 minuto)

---

## Dominio personalizzato (opzionale)

Se vuoi un indirizzo tipo `kairoscup2025.it`:

1. Acquista il dominio (es. su Namecheap, Aruba, ecc.)
2. Nel DNS del dominio aggiungi un record `CNAME`:
   - Nome: `www`
   - Valore: `TUONOME.github.io`
3. In GitHub → Settings → Pages → Custom domain → inserisci il tuo dominio
4. Spunta **Enforce HTTPS**

---

## Problemi comuni

| Problema | Soluzione |
|---|---|
| La Action non parte | Vai su Actions → abilita i workflow se richiesto |
| Il sito mostra dati vecchi | Svuota la cache del browser (Ctrl+Shift+R) |
| Errore `_CALC sheet not found` | Assicurati di usare il file Excel originale generato da `create_excel.py` |
| I cannonieri non appaiono | Controlla che nel foglio GIOCATORI ci siano valori > 0 nella colonna Gol |
