#!/usr/bin/env python3
"""
export_booknow.py - Le Maree di Carloforte
Legge il booking planner (.xlsx) ed esporta per il calendario pubblico:
  data/disponibilita.json   stati giornalieri, stagioni, tariffe, minimi (NESSUN dato ospiti)
  ical/<appartamento>.ics   prenotazioni dirette (+ chiusure se ICAL_INCLUDE_CLOSURES), per Booking.com

Uso:
  python3 export_booknow.py PLANNER.xlsx OUT_DIR [--booking-json booking_busy.json]

Regole (come nel planner):
  - stagione di ogni notte: periodo speciale se presente, altrimenti stagione base (PARAMETRI)
  - tariffa notte per notte; pulizie finali aggiunte una volta
  - soggiorno minimo: quello della stagione piu' alta toccata (Bassa/Media vuote = vale Alta)
  - RICHIESTE: Confermata = occupato, In attesa (opzione) = occupato; Rifiutata/Annullata ignorate
  - CHIUSURE: 'Tutti' o nome appartamento, notti incluse
"""
import json, re, sys, datetime as dt, hashlib, argparse
from pathlib import Path
import openpyxl

APARTMENTS_PUBLIC = ["AltaMarea", "BassaMarea", "CountryLoft"]
WINDOW_DAYS = 400          # ~13 mesi dal giorno di esportazione
# Le CHIUSURE del planner vanno anche a Booking.com? Oggi no: il planner chiude tutto il 2027
# "da aprire", e mandarlo a Booking bloccherebbe anche le vendite sul portale. Da decidere con Max.
ICAL_INCLUDE_CLOSURES = False
# Tassa di soggiorno aggiunta al prezzo indicativo (per persona per notte).
# Carloforte: dal 2018 risulta sostituita dal contributo di sbarco (pagato col traghetto) -> 0 finché Max non conferma.
TOURIST_TAX = {"per_person_night": 0.0, "max_nights": None, "exempt_under_age": None}
RANK = {"Bassa": 1, "Media": 2, "Alta": 3}
CODE = {"Bassa": "B", "Media": "M", "Alta": "A"}


def d(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


def header_index(ws, row, names):
    """Trova la riga di intestazione (cerca fino a 3 righe dopo quella indicata)."""
    for r in range(max(1, row - 2), row + 4):
        hdr = [c.value for c in ws[r]]
        if all(n in hdr for n in names):
            return {n: hdr.index(n) for n in names} | {"_row": r}
    raise SystemExit(f"{ws.title}: intestazioni non trovate {names}")


def read_planner(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    # --- appartamenti
    ws = wb["APPARTAMENTI"]
    cols = ["Nome", "Posti letto", "Tariffa Bassa €/notte", "Tariffa Media €/notte",
            "Tariffa Alta €/notte", "Pulizie finali €", "Sogg. minimo Alta (notti)",
            "Sogg. minimo Bassa (notti)", "Sogg. minimo Media (notti)", "Orari check-in / check-out"]
    ix = header_index(ws, 3, cols)
    apts = {}
    for row in ws.iter_rows(min_row=ix["_row"] + 1, values_only=True):
        name = row[ix["Nome"]]
        if not name or name not in APARTMENTS_PUBLIC:
            continue
        min_a = row[ix["Sogg. minimo Alta (notti)"]] or 1
        # Orari: testo libero, es. "16:00 / 10:00" -> primo orario = check-in, secondo = check-out
        times = re.findall(r"(\d{1,2})[:.](\d{2})", str(row[ix["Orari check-in / check-out"]] or ""))
        hours = {"in": f"{int(times[0][0]):02d}:{times[0][1]}", "out": f"{int(times[1][0]):02d}:{times[1][1]}"} if len(times) >= 2 else None
        apts[name] = {
            "name": name,
            "guests": int(row[ix["Posti letto"]] or 0),
            "rates": {"B": float(row[ix["Tariffa Bassa €/notte"]] or 0),
                      "M": float(row[ix["Tariffa Media €/notte"]] or 0),
                      "A": float(row[ix["Tariffa Alta €/notte"]] or 0)},
            "cleaning": float(row[ix["Pulizie finali €"]] or 0),
            "hours": hours,
            "min": {"B": int(row[ix["Sogg. minimo Bassa (notti)"]] or min_a),
                    "M": int(row[ix["Sogg. minimo Media (notti)"]] or min_a),
                    "A": int(min_a)},
        }
    missing = [a for a in APARTMENTS_PUBLIC if a not in apts]
    if missing:
        raise SystemExit(f"Appartamenti non trovati in APPARTAMENTI: {missing}")

    # --- stagioni base (PARAMETRI B7:D..)
    ws = wb["PARAMETRI"]
    base = []
    for r in range(7, 30):
        g, m, s = ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value
        if isinstance(g, (int, float)) and isinstance(m, (int, float)) and s in RANK:
            base.append((int(m), int(g), s))
    base.sort()
    if not base or base[0][:2] != (1, 1):
        raise SystemExit("PARAMETRI: la prima stagione deve partire dal 1/1")

    # --- periodi speciali
    ws = wb["PERIODI SPECIALI"]
    specials = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        a, b, ev, s = d(row[1]), d(row[2]), row[3], row[4]
        if a and b and s in RANK:
            specials.append((a, b, s, ev))

    # --- chiusure
    ws = wb["CHIUSURE"]
    closures = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        who, a, b = row[1], d(row[2]), d(row[3])
        if who and a and b:
            closures.append((who, a, b))

    # --- richieste (solo date e stato: nessun dato personale viene letto oltre)
    ws = wb["RICHIESTE"]
    ix = header_index(ws, 4, ["ID", "Appartamento", "Arrivo", "Partenza", "Stato"])
    stays = []
    for row in ws.iter_rows(min_row=ix["_row"] + 1, values_only=True):
        apt, a, b, st = row[ix["Appartamento"]], d(row[ix["Arrivo"]]), d(row[ix["Partenza"]]), row[ix["Stato"]]
        if apt in apts and a and b and b > a and st in ("Confermata", "In attesa"):
            stays.append((apt, a, b, st, row[ix["ID"]]))

    return apts, base, specials, closures, stays


def season_of(day, base, specials):
    for a, b, s, ev in specials:
        if a <= day <= b:
            return s, ev
    cur = base[0][2]
    for m, g, s in base:
        if (day.month, day.day) >= (m, g):
            cur = s
    return cur, None


def build(apts, base, specials, closures, stays, booking_busy, today):
    days = [today + dt.timedelta(i) for i in range(WINDOW_DAYS)]
    seasons, events = [], {}
    for i, day in enumerate(days):
        s, ev = season_of(day, base, specials)
        seasons.append(CODE[s])
        if ev:
            events[str(i)] = ev

    # stato per notte: F libero, O occupato (diretto/opzione), X occupato da portale, C chiuso
    avail, ical_blocks = {}, {}
    for name in apts:
        st = ["F"] * len(days)
        direct = []
        for who, a, b in closures:
            if who in ("Tutti", name):
                for i, day in enumerate(days):
                    if a <= day <= b:
                        st[i] = "C"
                if ICAL_INCLUDE_CLOSURES:
                    direct.append((a, b + dt.timedelta(1), "Chiuso"))
        for apt, a, b, stato, rid in stays:
            if apt != name:
                continue
            for i, day in enumerate(days):
                if a <= day < b and st[i] != "C":
                    st[i] = "O"
            direct.append((a, b, "Prenotato"))
        for a, b in booking_busy.get(name, []):
            for i, day in enumerate(days):
                if a <= day < b and st[i] == "F":
                    st[i] = "X"
        avail[name] = "".join(st)
        ical_blocks[name] = direct     # solo diretto + chiusure (mai rimandare a Booking le sue date)

    data = {
        "generated": dt.datetime.now().astimezone().isoformat(timespec="minutes"),
        "start": today.isoformat(),
        "days": len(days),
        "currency": "EUR",
        "whatsapp": "393487100832",
        "tax": TOURIST_TAX,
        "seasons": "".join(seasons),
        "events": events,
        "apartments": [
            {k: apts[n][k] for k in ("name", "guests", "rates", "cleaning", "min", "hours")}
            | {"avail": avail[n]} for n in APARTMENTS_PUBLIC
        ],
    }
    return data, ical_blocks


def write_ics(path, name, blocks):
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Le Maree di Carloforte//booknow//IT",
             "CALSCALE:GREGORIAN", f"X-WR-CALNAME:{name} - Le Maree di Carloforte"]
    stamp = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    for a, b, label in sorted(blocks):
        uid = hashlib.sha1(f"{name}{a}{b}{label}".encode()).hexdigest()[:16]
        lines += ["BEGIN:VEVENT", f"UID:{uid}@booknow.lemareedicarloforte.it", f"DTSTAMP:{stamp}",
                  f"DTSTART;VALUE=DATE:{a:%Y%m%d}", f"DTEND;VALUE=DATE:{b:%Y%m%d}",
                  f"SUMMARY:{label}", "END:VEVENT"]
    lines.append("END:VCALENDAR")
    path.write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("planner")
    ap.add_argument("outdir")
    ap.add_argument("--booking-json", help="date occupate dai portali {apt: [[arrivo, partenza], ...]}")
    ap.add_argument("--today", help="AAAA-MM-GG (test)")
    args = ap.parse_args()

    today = dt.date.fromisoformat(args.today) if args.today else dt.date.today()
    booking = {}
    if args.booking_json and Path(args.booking_json).exists():
        raw = json.loads(Path(args.booking_json).read_text())
        booking = {k: [(dt.date.fromisoformat(a), dt.date.fromisoformat(b)) for a, b in v]
                   for k, v in raw.items() if isinstance(v, list)}

    data, blocks = build(*read_planner(args.planner), booking, today)
    out = Path(args.outdir)
    (out / "data").mkdir(parents=True, exist_ok=True)
    (out / "ical").mkdir(parents=True, exist_ok=True)
    (out / "data" / "disponibilita.json").write_text(
        json.dumps(data, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    for name, bl in blocks.items():
        write_ics(out / "ical" / f"{name.lower()}.ics", name, bl)
    for a in data["apartments"]:
        c = {k: a["avail"].count(k) for k in "FOXC"}
        print(f"{a['name']:12s} liberi {c['F']:3d}  occupati {c['O']:3d}  portali {c['X']:3d}  chiusi {c['C']:3d}")
    print("OK ->", out)


if __name__ == "__main__":
    main()
