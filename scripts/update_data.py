#!/usr/bin/env python3
"""Genera injuries.csv (último reporte oficial NBA + MPG) y rest.csv (B2B / 3-en-4 / descanso).
Uso (desde la raíz del repo):  python scripts/update_data.py
Requiere: pip install pdfplumber"""
import csv, io, re, sys, json, unicodedata, datetime as dt, urllib.request, urllib.error
from zoneinfo import ZoneInfo
import pdfplumber

ET = ZoneInfo("America/New_York")
UA = {"User-Agent": "Mozilla/5.0"}
PDF_URL = "https://ak-static.cms.nba.com/referee/injury/Injury-Report_{d}_{h:02d}_{m:02d}{ap}.pdf"
SCHED_URLS = ["https://cdn.nba.com/static/json/staticData/scheduleLeagueV2_1.json",
              "https://cdn.nba.com/static/json/staticData/scheduleLeagueV2.json"]
KEY_PER = 18      # PER mínimo para considerar 'clave' (promedio liga = 15)
MIN_MPG = 12      # y al menos estos minutos por partido (evita PER inflado con pocos minutos)

def fetch(url):
    try:
        return urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=40).read()
    except urllib.error.HTTPError as e:
        if e.code in (403, 404): return None
        raise
    except Exception as e:
        print("  fallo", url, e); return None

# ---------- nombres ----------
def norm(n):
    n = unicodedata.normalize("NFKD", n).encode("ascii", "ignore").decode().lower()
    if "," in n: a, b = n.split(",", 1); n = b + " " + a      # "Last, First" -> "First Last"
    n = re.sub(r"[.'’\-]", " ", n)
    n = re.sub(r"\b(jr|sr|ii|iii|iv|v)\b", " ", n)
    return re.sub(r"\s+", " ", n).strip()

def load_players(path="players.csv"):
    """Lee el CSV de Basketball-Reference (tabla Advanced: trae PER y MP total). Devuelve {nombre: (mpg, per)}."""
    d = {}
    try:
        lines = [l[1:-1].replace('""', '"') if l.startswith('"') and l.endswith('"') else l      # guardado desde Excel (todo en col. A)
                 for l in open(path, encoding="utf-8-sig").read().splitlines()]
    except FileNotFoundError:
        print("players.csv no encontrado: las lesiones se listan sin PER ni MPG"); return d
    i = next((k for k, l in enumerate(lines) if re.search(r"(^|,)Player(,|$)", l) and re.search(r"(^|,)(MP|MPG|MIN)\W*(,|$)", l, re.I)), None)
    if i is None: print("players.csv: no se encontró encabezado con Player y MP"); return d
    rows = []
    for r in csv.DictReader(lines[i:]):
        r = {re.sub(r"[^a-z]", "", (k or "").lower()): v for k, v in r.items()}
        name = r.get("player")
        if not name or name == "Player": continue                     # filas de encabezado repetidas
        try: mp = float(r.get("mpg") or r.get("mp") or r.get("min"))
        except (TypeError, ValueError): continue
        try: per = float(r.get("per"))
        except (TypeError, ValueError): per = None
        try: g = float(r.get("g") or 0)
        except ValueError: g = 0
        rows.append((norm(name), mp, g, per))
    totals = max((m for _, m, _, _ in rows), default=0) > 60          # MP = minutos totales (tabla Advanced/Totals)
    for n, m, g, per in rows:
        if totals:
            if g <= 0: continue
            m = m / g
        d.setdefault(n, (round(m, 1), per))                           # BBRef: la fila TOT va primero
    print("players.csv:", len(d), "jugadores;", "con PER" if any(v[1] is not None for v in d.values()) else "SIN columna PER (usa la tabla Advanced)")
    return d

# ---------- PDF de lesiones ----------
def latest_pdf(now):
    t = now.replace(minute=now.minute // 15 * 15, second=0, microsecond=0)
    for _ in range(200):                                   # hasta ~50 h hacia atrás
        h12 = t.hour % 12 or 12
        url = PDF_URL.format(d=t.strftime("%Y-%m-%d"), h=h12, m=t.minute, ap="AM" if t.hour < 12 else "PM")
        b = fetch(url)
        if b: return b, t, url
        t -= dt.timedelta(minutes=15)
    return None, None, None

def parse_pdf(data):
    rows, cols, cur, prev = [], None, {}, None
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=2)
            lines = {}
            for w in words: lines.setdefault(round(w["top"] / 3), []).append(w)
            for key in sorted(lines):
                ws = sorted(lines[key], key=lambda w: w["x0"]); text = " ".join(w["text"] for w in ws)
                if text.startswith("Injury Report:") or re.match(r"Page \d+ of \d+", text): continue
                if "Game Date" in text and "Matchup" in text:
                    x = lambda s, i=0: [w["x0"] for w in ws if w["text"] == s][i]
                    cols = [("date", x("Game", 0)), ("time", x("Game", 1)), ("match", x("Matchup")), ("team", x("Team")),
                            ("player", x("Player")), ("status", x("Current")), ("reason", x("Reason"))]
                    continue
                if cols is None: continue
                c = {n: [] for n, _ in cols}
                for w in ws:
                    n = [n for n, x0 in cols if w["x0"] >= x0 - 3][-1]; c[n].append(w["text"])
                c = {k: " ".join(v) for k, v in c.items()}
                if "NOT YET SUBMITTED" in text:
                    for k in ("date", "time", "match", "team"):
                        if c[k]: cur[k] = c[k]
                    rows.append({**cur, "player": "", "status": "NOT YET SUBMITTED", "reason": ""}); prev = None; continue
                for k in ("date", "time", "match", "team"):
                    if c[k]:
                        if k == "date" or (k == "time" and c["date"]): pass
                        cur[k] = c[k]
                if c["player"] or c["status"]:
                    prev = {**cur, "player": c["player"], "status": c["status"], "reason": c["reason"]}; rows.append(prev)
                elif c["reason"] and prev: prev["reason"] += " " + c["reason"]
    return rows

def build_injuries(now):
    data, t, url = latest_pdf(now)
    if not data: print("No se encontró reporte de lesiones"); return
    print("Reporte:", url)
    players, out = load_players(), []
    for r in parse_pdf(data):
        gd = r.get("date", ""); m = re.match(r"(\d\d)/(\d\d)/(\d{4})", gd)
        iso = f"{m.group(3)}-{m.group(1)}-{m.group(2)}" if m else gd
        mp, per = players.get(norm(r["player"]), (None, None)) if r["player"] else (None, None)
        gl = bool(re.search(r"g league", r["reason"], re.I))
        out.append({"report_et": t.strftime("%Y-%m-%d %H:%M"), "game_date": iso, "game_time": r.get("time", ""),
                    "matchup": r.get("match", ""), "team": r.get("team", ""), "player": r["player"], "status": r["status"],
                    "reason": r["reason"].strip(), "mpg": "" if mp is None else mp, "per": "" if per is None else per,
                    "key": int(per is not None and per >= KEY_PER and (mp or 0) >= MIN_MPG and r["status"] in ("Out", "Doubtful") and not gl)})
    with open("injuries.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(out[0].keys()) if out else ["report_et"]); w.writeheader(); w.writerows(out)
    print("injuries.csv:", len(out), "filas")

# ---------- descanso ----------
ABBR = {"Atlanta Hawks": "ATL", "Boston Celtics": "BOS", "Brooklyn Nets": "BKN", "Charlotte Hornets": "CHA", "Chicago Bulls": "CHI",
        "Cleveland Cavaliers": "CLE", "Dallas Mavericks": "DAL", "Denver Nuggets": "DEN", "Detroit Pistons": "DET", "Golden State Warriors": "GSW",
        "Houston Rockets": "HOU", "Indiana Pacers": "IND", "Los Angeles Clippers": "LAC", "LA Clippers": "LAC", "Los Angeles Lakers": "LAL",
        "Memphis Grizzlies": "MEM", "Miami Heat": "MIA", "Milwaukee Bucks": "MIL", "Minnesota Timberwolves": "MIN", "New Orleans Pelicans": "NOP",
        "New York Knicks": "NYK", "Oklahoma City Thunder": "OKC", "Orlando Magic": "ORL", "Philadelphia 76ers": "PHI", "Phoenix Suns": "PHX",
        "Portland Trail Blazers": "POR", "Sacramento Kings": "SAC", "San Antonio Spurs": "SAS", "Toronto Raptors": "TOR", "Utah Jazz": "UTA",
        "Washington Wizards": "WAS"}

def games_from_csv(path="schedule.csv"):
    """schedule.csv = calendarios mensuales de Basketball-Reference pegados uno tras otro (columnas: Date,Start,Visitor,PTS,Home,...)."""
    try: lines = open(path, encoding="utf-8-sig").read().splitlines()
    except FileNotFoundError: return []
    lines = [l[1:-1].replace('""', '"') if l.startswith('"') and l.endswith('"') else l for l in lines]
    out = []
    for row in csv.reader(lines):
        if len(row) < 5 or row[0].strip() == "Date": continue                  # encabezados repetidos / líneas vacías
        try: d = dt.datetime.strptime(" ".join(row[0].split()[1:]), "%b %d %Y").date().isoformat()   # "Sun Nov 1 2026"
        except ValueError: continue
        a, h = ABBR.get(row[2].strip()), ABBR.get(row[4].strip())
        if a and h: out.append((d, a, h))
        else: print("Equipo no reconocido en schedule.csv:", row[2], "/", row[4])
    return out

def games_from_json():
    raw = next((b for b in (fetch(u) for u in SCHED_URLS) if b), None)
    if not raw: return []
    games = []
    for gd in json.loads(raw)["leagueSchedule"]["gameDates"]:
        for g in gd["games"]:
            if str(g.get("gameId", "")).startswith("001"): continue                     # sin pretemporada
            s = g.get("gameDateEst") or g.get("gameDateTimeEst") or g.get("gameDateTimeUTC") or ""
            games.append((s[:10], g["awayTeam"]["teamTricode"], g["homeTeam"]["teamTricode"]))
    return games

def build_rest(now):
    games = games_from_csv() or games_from_json()
    if not games: print("Sin calendario: sube schedule.csv al repo"); return
    print("Calendario:", len(games), "partidos; ejemplo:", games[len(games) // 2] if games else "vacío")
    played = {}
    for d, a, h in games:
        for t in (a, h): played.setdefault(t, set()).add(d)
    D = lambda s: dt.date.fromisoformat(s)
    def ctx(t, d):
        prev = sorted(x for x in played[t] if x < d)
        rest = (D(d) - D(prev[-1])).days - 1 if prev else ""
        last4 = sum(1 for x in played[t] if 0 <= (D(d) - D(x)).days <= 3)
        return rest, int(rest == 0), int(last4 >= 3)
    today = now.date(); keep = {today.isoformat(), (today + dt.timedelta(days=1)).isoformat()}
    rows = []
    for d, a, h in sorted(games):
        if d in keep:
            ra, ha = ctx(a, d), ctx(h, d)
            rows.append(dict(date=d, road=a, home=h, road_rest=ra[0], road_b2b=ra[1], road_3in4=ra[2],
                             home_rest=ha[0], home_b2b=ha[1], home_3in4=ha[2]))
    with open("rest.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["date", "road", "home", "road_rest", "road_b2b", "road_3in4", "home_rest", "home_b2b", "home_3in4"])
        w.writeheader(); w.writerows(rows)
    print("rest.csv:", len(rows), "partidos")

if __name__ == "__main__":
    now = dt.datetime.now(ET)
    for fn in (build_injuries, build_rest):
        try: fn(now)
        except Exception as e: print("ERROR en", fn.__name__, "->", e)
