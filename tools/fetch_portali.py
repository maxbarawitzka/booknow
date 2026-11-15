#!/usr/bin/env python3
"""Legge i calendari iCal di Booking.com e Airbnb e scrive le date occupate.

Gira su GitHub Actions, che ha accesso libero alla rete (il computer e la
sessione Claude no). I link dei portali arrivano dal segreto ICAL_SOURCES e
non compaiono mai nel repository.

IMPORTANTE - privacy: dagli eventi si prendono SOLO le date (DTSTART/DTEND).
SUMMARY, DESCRIPTION, UID e qualsiasi altro campo vengono scartati: l'iCal di
Airbnb puo' contenere il link alla prenotazione e le ultime cifre del telefono
dell'ospite, e questo repository e' pubblico.

Uso: python3 tools/fetch_portali.py            (legge $ICAL_SOURCES)
     python3 tools/fetch_portali.py file.json  (legge un file locale)
"""
import json
import os
import sys
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta, timezone

OUT = "data/portali_occupati.json"
TIMEOUT = 45
UA = "Mozilla/5.0 (compatible; LeMareeBooknow/1.0)"
ORIZZONTE = 400  # giorni in avanti: oltre non serve al calendario pubblico


def unfold(testo):
    """Ricompone le righe iCal spezzate (continuano con spazio o tab)."""
    righe = []
    for r in testo.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if r[:1] in (" ", "\t") and righe:
            righe[-1] += r[1:]
        else:
            righe.append(r)
    return righe


def leggi_data(valore):
    v = valore.strip()
    if "T" in v:
        v = v.split("T")[0]
    if len(v) != 8 or not v.isdigit():
        return None
    return date(int(v[:4]), int(v[4:6]), int(v[6:8]))


def date_occupate(ics):
    """Date occupate da un iCal. DTEND e' escluso (giorno di partenza libero)."""
    giorni, dentro, inizio, fine = set(), False, None, None
    for riga in unfold(ics):
        nome = riga.split(":", 1)[0].split(";", 1)[0].upper()
        if riga.upper().startswith("BEGIN:VEVENT"):
            dentro, inizio, fine = True, None, None
        elif riga.upper().startswith("END:VEVENT"):
            if inizio and fine and fine > inizio:
                g = inizio
                while g < fine:
                    giorni.add(g)
                    g += timedelta(days=1)
            elif inizio and fine == inizio:
                giorni.add(inizio)
            dentro, inizio, fine = False, None, None
        elif dentro and ":" in riga:
            valore = riga.split(":", 1)[1]
            if nome == "DTSTART":
                inizio = leggi_data(valore)
            elif nome == "DTEND":
                fine = leggi_data(valore)
    return giorni


def scarica(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "text/calendar, */*"})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
        return r.read().decode("utf-8", "replace")


def main():
    if len(sys.argv) > 1:
        sorgenti = json.load(open(sys.argv[1], encoding="utf-8"))
    else:
        grezzo = os.environ.get("ICAL_SOURCES", "").strip()
        if not grezzo:
            print("ICAL_SOURCES non impostato: niente da fare.", file=sys.stderr)
            return 1
        sorgenti = json.loads(grezzo)

    appartamenti = sorgenti.get("appartamenti", sorgenti)
    oggi = date.today()
    limite = oggi + timedelta(days=ORIZZONTE)

    risultato = {
        "generato": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "orizzonte_giorni": ORIZZONTE,
        "nota": "Solo date. Nessun dato degli ospiti: SUMMARY e DESCRIPTION degli eventi vengono scartati.",
        "appartamenti": {},
    }
    errori_totali = 0

    for nome, portali in appartamenti.items():
        voce = {"occupate": [], "per_portale": {}, "errori": []}
        tutte = set()
        for portale, url in sorted(portali.items()):
            if not url or not str(url).startswith("http"):
                voce["per_portale"][portale] = None
                continue
            try:
                giorni = {g for g in date_occupate(scarica(url)) if oggi <= g <= limite}
                tutte |= giorni
                voce["per_portale"][portale] = len(giorni)
                print(f"{nome}/{portale}: {len(giorni)} notti occupate")
            except (urllib.error.URLError, urllib.error.HTTPError, OSError, ValueError) as e:
                errori_totali += 1
                voce["per_portale"][portale] = "errore"
                voce["errori"].append(f"{portale}: {type(e).__name__}: {e}")
                print(f"{nome}/{portale}: ERRORE {e}", file=sys.stderr)
        voce["occupate"] = sorted(g.isoformat() for g in tutte)
        risultato["appartamenti"][nome] = voce

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(risultato, f, ensure_ascii=False, indent=1, sort_keys=True)
        f.write("\n")
    print(f"\nScritto {OUT} · errori: {errori_totali}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
