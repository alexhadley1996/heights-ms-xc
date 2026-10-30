#!/usr/bin/env python3
"""Build data/results.js (and results.json) from the "Heights MS XC Results" xlsx export.

Usage:  python3 scripts/build_data.py path/to/Heights_MS_XC_Results.xlsx

Also merges data/supplemental/ (races found online + per-race overrides for sheet races);
see README "Supplemental data and corrections".
Requires: openpyxl.  Re-run after exporting a fresh copy of the Google Sheet (File > Download > .xlsx).
New meet tabs are picked up automatically if their header row uses the usual column names
(Place, Heights Place, Grade/Year, Name or First/Last Name, Pace, Time, School/Team).
Add an entry to MEET_INFO below to give a new tab a nicer name, date or distance.
"""
import sys, re, json, datetime, os, statistics, csv, difflib

import openpyxl

# Tabs that are analysis/comparison sheets, not meet results.
SKIP_TABS = {"2026 Comparison", "Comparison 2025 - 6th Grade"}

# Per-tab metadata.  Only facts that appear in the sheet (tab name / header notes) are used.
# distance_mi: from a "Distance:" note in the tab header or the tab name; None if the sheet doesn't say.
MEET_INFO = {
    "2026 States":               dict(season=2026, name="State Championship", distance=None),
    "2026 - Bullis":             dict(season=2026, name="Bullis", distance=None),
    "2026 - Flint Hill":         dict(season=2026, name="Flint Hill", distance=2.1),
    "2025 States":               dict(season=2025, name="State Championship", distance=None),
    "Flint Hill 2025":           dict(season=2025, name="Flint Hill", distance=None),
    "2025 Bittersweet #2":       dict(season=2025, name="Bittersweet #2", distance=None),
    "2025 Bittersweet #1":       dict(season=2025, name="Bittersweet #1", distance=None),
    "2024 State Championship":   dict(season=2024, name="State Championship", distance=None),
    "2024 - Flint Hill":         dict(season=2024, name="Flint Hill", distance=None),
    "2024 - Landon 1.8 miles":   dict(season=2024, name="Landon", distance=1.8),
    "2024 - Gtown Prep":         dict(season=2024, name="Georgetown Prep", distance=2.0),
    "2024 - Bullis":             dict(season=2024, name="Bullis", distance=2.0),
    "2023 State Championship":   dict(season=2023, name="State Championship", distance=2.0),
    "2023 - Bullis":             dict(season=2023, name="Bullis", distance=1.8),
    "2023 Flint Hill":           dict(season=2023, name="Flint Hill", distance=2.1),
    "2023 - Covenant Life":      dict(season=2023, name="Covenant Life", distance=2.0),
    "2022 - States":             dict(season=2022, name="State Championship", distance=2.1),
    "2022 - CL":                 dict(season=2022, name="Covenant Life", distance=2.0),
    "2021 - States at Landon":   dict(season=2021, name="State Championship (at Landon)", distance=2.2),
    "2010 - Landon #2":          dict(season=2010, name="Landon #2", distance=None),
    "2010 - Landon #1":          dict(season=2010, name="Landon #1", distance=None),
    "2009 - Landon, Nov 3":      dict(season=2009, name="Landon", date="Nov 3", distance=None),
    "2009 - Landon, Sept 28":    dict(season=2009, name="Landon", date="Sept 28", distance=None),
    "2008 - Landon, Nov 4":      dict(season=2008, name="Landon", date="Nov 4", distance=None),
    "2008 - Landon, Oct 17":     dict(season=2008, name="Landon", date="Oct 17", distance=None),
    "2008 - Landon, Oct 6":      dict(season=2008, name="Landon", date="Oct 6", distance=None),
    "2007 - Landon, October 18": dict(season=2007, name="Landon", date="Oct 18", distance=None),
    "2007 - Landon Invit":       dict(season=2007, name="Landon Invitational", distance=None),
    "2006 - Landon, October 18": dict(season=2006, name="Landon", date="Oct 18", distance=None),
    "2006 - Landon Invit":       dict(season=2006, name="Landon Invitational", distance=None),
}
# Chronological order within a season, when known from the sheet (tab order is newest-first).
# Otherwise tab order (reversed) is used.

# Spelling variants of the same Heights runner -> canonical name.
NAME_ALIASES = {
    "Sean Carney (real)": "Sean Carney",
    "John Ortman": "Johnny Ortman",      # 2023 States; "Johnny" in 5 other tabs
    "Louie Liess": "Louie Leiss",        # 2024 Bullis
    "P Allen": "Phillip Allen",          # 2007 Landon Oct 18
    "Vince Kiernan": "Vince Keirnan",    # 2007; "Keirnan" in the 2008 tabs (true spelling unknown)
    "Santi Luli": "Santiago Luli",       # 2008 tabs
    "Andres Lulli": "Andres Luli",       # 2009 Nov 3
    "Ed Nelson": "Edward Nelson",
}

# Hand-written notes about the sheet, shown on the Meets page.
EXTRA_NOTES = [
    "2024 Bullis: no grades recorded; several non-Heights runners listed by first name only; the sheet's Meet column auto-incremented (2024 Bullis, 2025 Bullis, ...) and is ignored.",
    "2024 Landon: the sheet's 'Heights Place' column numbers every finisher, so the site recalculates Heights place from the School column (same for every race).",
    "2025 States: 'Sean Carney (real)' was added by hand with no overall place (shown without a place); Heights place 43 is skipped in the sheet.",
    "Several 2008 Landon (Oct 6) and 2023 Flint Hill times were stored as clock times (e.g. 13:36 read as 13 h 36 min); they are converted back to minutes:seconds.",
    "2010 meets have no pace column; most 2006–2010 meets, 2022 Covenant Life, 2023 Flint Hill and 2024 Flint Hill/Bullis have no grades.",
    "Race distances come from the divisor in each tab's pace formula (pace = time \u00f7 miles). The two 2010 Landon tabs have no pace formula, so their distance is unknown. 2026 Flint Hill's header says 2.1 mi but its formula divides by 2; the formula is used.",
    "Where a tab has no pace for a runner (e.g. non-Heights runners at 2023 States and 2023 Bullis), the site calculates it the same way: time \u00f7 the tab's distance.",
    "Team scores are calculated by the site from the listed finishers (standard scoring: teams with 5+ finishers, top 5 score, 6th\u20137th displace, 8th+ ignored, ties broken by the 6th runner) and compared with official meet scores where those were found online.",
    "Races from 2011\u20132019, 2021 CLS, 2022 FHS and the 2006 Landon Home Meet (Sept 15) are not in the sheet; they come from racecast.co / Athletic.net / mocorunning.com (data/supplemental). Each of those meet pages links its source.",
    "Corrected from official results: 2022 States Race 1 (A race) had 31 B-race runners mixed in, 2021 States and 2007 Landon Invitational were missing finishers, and 2025 States did not score 5th graders. See each meet page.",
    "Name spellings merged for the runner pages: John/Johnny Ortman, Louie Liess/Leiss, P/Phillip Allen, Vince Kiernan/Keirnan, Santi/Santiago Luli, Andres Lulli/Luli, Ed/Edward Nelson.",
    "2026 States lists 5th graders as exhibition ('Ex.') and two DNFs.",
]

HEIGHTS_RE = re.compile(r"\bheights\b", re.I)


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def fmt_secs(t):
    if t is None:
        return None
    m, s = divmod(t, 60)
    return f"{int(m)}:{s:04.1f}"


def parse_time(v):
    """Return seconds (float) or None. Handles 'mm:ss.t' strings and Excel time values that
    Sheets mis-read as h:mm (13:36 -> 13 min 36 s)."""
    if v is None:
        return None
    if isinstance(v, datetime.time):
        return v.hour * 60 + v.minute + v.second / 60  # h:mm stored -> mm:ss
    if isinstance(v, datetime.timedelta):
        tot = v.total_seconds()  # e.g. "1 day, 4:31:00" == 28:31 mis-read as hours
        return tot / 60
    if isinstance(v, datetime.datetime):
        return v.hour * 60 + v.minute + v.second / 60
    if isinstance(v, (int, float)):
        return None
    s = str(v).strip()
    m = re.fullmatch(r"(\d+):(\d{1,2}(?:\.\d+)?)", s)
    if m:
        return int(m.group(1)) * 60 + float(m.group(2))
    m = re.fullmatch(r"(\d+):(\d{2}):(\d{2})", s)  # e.g. 14:52:00 meaning 14:52
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return None


def parse_pace(v):
    if isinstance(v, datetime.time):
        return v.hour * 60 + v.minute
    if v is None:
        return None
    m = re.fullmatch(r"(\d+):(\d{2})", str(v).strip())
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def clean_school(s):
    s = (str(s).strip() if s is not None else "")
    s = re.sub(r"\s*-\s*#\d+$", "", s)  # "Flint Hill - #1827" -> "Flint Hill"
    return s


def find_cols(header):
    h = [str(c).strip().lower() if c is not None else "" for c in header]
    def first(*names, exclude=()):
        for i, c in enumerate(h):
            if c in names:
                return i
        return None
    cols = {
        "place": first("place", "overall place", "pl"),
        "hplace": next((i for i, c in enumerate(h) if c.startswith("place (heights") or c == "heights place"), None),
        "grade": first("grade", "year"),
        "name": first("name", "named"),
        "first": first("first name"),
        "last": first("last name"),
        "pace": first("pace"),
        "time": first("time", "finals"),
        "school": first("school", "team"),
        "points": first("points", "score"),
    }
    # 2007 tabs have an unlabeled pace column right before Time
    if cols["pace"] is None and cols["time"] is not None and h[cols["time"] - 1] == "":
        cols["pace"] = cols["time"] - 1
    return cols


PACE_DIVISOR_RE = re.compile(r"\)\)\s*/\s*([\d.]+)\s*\)\s*/\s*86400")
SCORE_NOTE_RE = re.compile(r"(\d)\w*\s+gr\w*\s+(expanded\s+)?score", re.I)
NON_TEAM_RE = re.compile(r"^(unattached|unatt|independent|individual|none|)$", re.I)


def formula_distance(ws_f):
    """Distance in miles = the divisor in the tab's pace formula (pace = time / distance)."""
    from collections import Counter
    c = Counter()
    for row in ws_f.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                m = PACE_DIVISOR_RE.search(v)
                if m:
                    c[float(m.group(1))] += 1
    return c


def sheet_score_notes(ws):
    """Team score notes typed into the side of some tabs, e.g. '8th Grade Score: 22'."""
    notes = []
    for row in ws.iter_rows(values_only=True):
        for i, v in enumerate(row):
            if isinstance(v, str) and SCORE_NOTE_RE.search(v):
                label = v.strip().rstrip(":").strip()
                val = None
                m = re.search(r":\s*(\d+)\s*$", label)
                if m:
                    val = int(m.group(1)); label = label[:m.start()].strip()
                else:
                    nxt = row[i + 1] if i + 1 < len(row) else None
                    val = int(nxt) if isinstance(nxt, (int, float)) else None
                notes.append({"label": label, "score": val})
    return notes


def team_key(school):
    return re.sub(r"[^a-z0-9]+", " ", school.lower()).strip()


def score_race(field, exclude_grades=(), max_per_team=7):
    """Standard XC scoring. Returns (standings, partial_teams). Runners in exclude_grades keep their place but don't score."""
    for e in field:
        e.pop("teamPlace", None)
    finishers = [e for e in field if isinstance(e["place"], int) and e["secs"] is not None and e.get("grade") not in exclude_grades]
    finishers.sort(key=lambda e: e["place"])
    by_team = {}
    for e in finishers:
        if NON_TEAM_RE.match(e["school"].strip()):
            continue
        by_team.setdefault(team_key(e["school"]), []).append(e)
    full = {k: v for k, v in by_team.items() if len(v) >= 5}
    scorers = set()
    for k, v in full.items():
        for e in v[:max_per_team]:
            scorers.add(id(e))
    tp = 0
    for e in finishers:
        if id(e) in scorers:
            tp += 1
            e["teamPlace"] = tp
    standings = []
    for k, v in full.items():
        top = v[:7]
        pts = [e["teamPlace"] for e in top]
        five = top[:5]
        total5 = sum(e["secs"] for e in five)
        standings.append({
            "team": v[0]["school"],
            "heights": any(e["heights"] for e in v),
            "score": sum(pts[:5]),
            "scorers": [{"name": e["name"], "place": e["teamPlace"], "overall": e["place"], "time": e["time"]} for e in five],
            "displacers": [{"name": e["name"], "place": e["teamPlace"], "overall": e["place"], "time": e["time"]} for e in top[5:7]],
            "finishers": len(v),
            "teamTime": fmt_secs(total5),
            "avgTime": fmt_secs(total5 / 5),
            "spread": fmt_secs(five[-1]["secs"] - five[0]["secs"]),
            "tiebreak": pts[5] if len(pts) > 5 else 10 ** 6,
        })
    standings.sort(key=lambda t: (t["score"], t["tiebreak"]))
    for i, t in enumerate(standings):
        t["rank"] = i + 1
        if i and t["score"] == standings[i - 1]["score"]:
            t["tieNote"] = "tie broken by 6th runner"
            standings[i - 1]["tieNote"] = "tie broken by 6th runner"
    for t in standings:
        if t["tiebreak"] == 10 ** 6:
            t["tiebreak"] = None
    partial = sorted({e["school"] for k, v in by_team.items() if len(v) < 5 for e in v[:1]})
    return standings, partial


# ---------------------------------------------------------------------------
# Supplemental data (races found online) and per-race overrides for sheet races
# ---------------------------------------------------------------------------
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
SUPP = os.path.join(ROOT, "data", "supplemental")
DIST_TOL = 0.01  # miles; official vs sheet distance closer than this counts as the same (e.g. 1.599 vs 1.6)

TEAM_STOP = {"the", "school", "schools", "middle", "ms", "sc", "sch", "of", "and", "academy", "day", "episcopal", "prep",
             "preparatory", "international", "friends", "christian", "catholic", "s", "st", "saint", "home", "homeschool", "hs"}
TEAM_ALIAS = {"geoergetown": "georgetown", "gds": "georgetown", "sssas": "stephens agnes", "sssa": "stephens agnes",
              "saes": "andrews", "cls": "covenant life", "stephen": "stephens", "alban": "albans", "andrew": "andrews"}


def _toks(n):
    n = re.sub(r"\((md|dc|va)\)", " ", n.lower()).replace("&", " and ").replace("'", "")
    w = re.findall(r"[a-z0-9]+", n)
    out = []
    for x in w:
        out += TEAM_ALIAS.get(x, x).split()
    return [x for x in out if x not in TEAM_STOP] or w


def match_team(name, cands):
    k = set(_toks(name)); best, bs = None, 0
    for c in cands:
        ck = set(_toks(c))
        if not k or not ck:
            continue
        s = len(k & ck) / min(len(k), len(ck))
        s2 = difflib.SequenceMatcher(None, " ".join(_toks(name)), " ".join(_toks(c))).ratio()
        sc = max(s, s2 * 0.95)
        if sc > bs:
            best, bs = c, sc
    return best if bs >= 0.6 else None


def load_json(name, default):
    p = os.path.join(SUPP, name)
    return json.load(open(p)) if os.path.exists(p) else default


def load_official_scores():
    p = os.path.join(SUPP, "official_team_scores.csv")
    out = {}
    if os.path.exists(p):
        for r in csv.DictReader(open(p, encoding="utf-8")):
            out.setdefault(r["Meet"], []).append(r)
    return out


def make_entry(place_raw, grade, name, school, time_raw, pace_raw, quirk_ctx=None):
    place = num(place_raw)
    name = re.sub(r"\s+", " ", str(name or "").strip())
    school = clean_school(school)
    t = parse_time(time_raw)
    status = None
    if isinstance(time_raw, str) and time_raw.strip().upper() in ("DNF", "DNS", "DQ"):
        status = time_raw.strip().upper()
    g = num(grade)
    pace = parse_pace(pace_raw)
    heights = bool(HEIGHTS_RE.search(school))
    e = {
        "place": int(place) if place is not None else (str(place_raw).strip() if place_raw not in (None, "") else None),
        "grade": int(g) if g else None,
        "name": name,
        "school": school,
        "time": fmt_secs(t) if t is not None else status,
        "secs": round(t, 2) if t is not None else None,
        "pace": f"{pace // 60}:{pace % 60:02d}" if pace else None,
        "paceSecs": pace,
        "heights": heights,
    }
    if heights and name in NAME_ALIASES:
        e["note"] = f"Listed as \u201c{name}\u201d"
        e["name"] = NAME_ALIASES[name]
    return e


def read_csv_field(rel):
    rows = list(csv.DictReader(open(os.path.join(SUPP, rel), encoding="utf-8")))
    # pace in these files is computed by the source from its own distance; recompute from ours instead
    return [make_entry(r.get("Place"), r.get("Grade"), r.get("Name"), r.get("School"), r.get("Time"), None) for r in rows]


def norm_name(n):
    return re.sub(r"[^a-z]", "", n.lower())


def fill_from_official(field, official, title, quirks):
    """Keep the sheet rows; add official finishers the sheet doesn't have. Returns number added.
    A sheet row and an official row are the same runner if the names match, or place and time match."""
    names = {norm_name(e["name"]) for e in field}
    by_place = {e["place"]: e for e in field if isinstance(e["place"], int)}
    school_map, missing = {}, []
    for o in official:
        s = by_place.get(o["place"])
        same_slot = s and s["secs"] is not None and o["secs"] is not None and abs(s["secs"] - o["secs"]) <= 1.5
        if same_slot:
            if s["school"]:
                school_map.setdefault(o["school"], s["school"])
            else:
                s["blankSchool"] = o["school"]
        if norm_name(o["name"]) in names or same_slot:
            continue
        missing.append(o)
    for e in field:
        if e.pop("blankSchool", None) is not None:
            o_school = next(o["school"] for o in official if o["place"] == e["place"])
            e["school"] = school_map.get(o_school, o_school)
            e["heights"] = bool(HEIGHTS_RE.search(e["school"]))
            e["note"] = "School was blank in the sheet; filled in from the official results"
            if e["heights"] and e["name"] in NAME_ALIASES:
                e["note"] += f" (listed as \u201c{e['name']}\u201d)"
                e["name"] = NAME_ALIASES[e["name"]]
            quirks.append(f"{title}: {e['name']} (place {e['place']}) had no school in the sheet; set to {e['school']} from the official results.")
    for o in missing:
        # use the sheet's spelling of the team name (learned from matched runners) so they score together
        o["school"] = school_map.get(o["school"], o["school"])
        o["heights"] = bool(HEIGHTS_RE.search(o["school"]))
        o["added"] = True
        field.append(o)
    field.sort(key=lambda e: (e["place"] if isinstance(e["place"], int) else 10 ** 6, e["secs"] or 10 ** 6))
    return len(missing)


def parse_md(season, s):
    if not s:
        return None
    for fmt in ("%b %d", "%B %d"):
        try:
            d = datetime.datetime.strptime(s.replace("Sept", "Sep"), fmt)
            return datetime.date(season, d.month, d.day).isoformat()
        except ValueError:
            pass
    return None


def finalize(field, meta, quirks):
    """Shared per-race processing: suspect flags, Heights places, paces, team scores, official comparison."""
    title = meta["tab"] or meta["name"]
    slowest_ahead = 0
    for e in field:
        if e["secs"] is None or not isinstance(e["place"], int):
            continue
        if e["secs"] < slowest_ahead - 1.0:
            e["suspect"] = True
            quirks.append(f"{meta['season']} {meta['name']}: {e['name']} ({e['school']}) place {e['place']} has time {e['time']}, faster than runners placed ahead \u2013 time/place likely mis-entered.")
        slowest_ahead = max(slowest_ahead, e["secs"])
    hp = 0
    for e in field:
        if e["heights"]:
            if isinstance(e["place"], int) and e["secs"] is not None:
                hp += 1; e["hplace"] = hp
            else:
                e["hplace"] = None
    dist, dsrc = meta["distance"], meta["distanceSource"]
    if dist and dsrc != "est":
        for e in field:
            if e["secs"] and not e["paceSecs"]:
                ps = int(e["secs"] / dist)
                e["paceSecs"], e["pace"] = ps, f"{ps // 60}:{ps % 60:02d}"
                if dsrc != "official":
                    e["paceCalc"] = True
    excl = tuple(meta.get("scoreExcludeGrades") or ())
    standings, partial = score_race(field, excl, 10 ** 6 if meta.get("noDisplacerLimit") else 7)
    if excl:
        for e in field:
            if e.get("grade") in excl and isinstance(e["place"], int):
                e["noScore"] = True
    caveats = []
    ints = sorted(e["place"] for e in field if isinstance(e["place"], int))
    missing = (ints[-1] - len(set(ints))) if ints else 0
    if missing > 0:
        caveats.append(f"{missing} place number(s) are missing from the list, so some finishers aren't listed; scores may be off.")
    if sum(1 for e in field if not e["heights"]) < 0.1 * len(field):
        caveats.append("Almost every runner listed is from The Heights, so the field looks incomplete.")
    if len(standings) < 2:
        caveats.append("Fewer than two teams have 5 finishers listed, so no team score is shown.")
    hv = sorted({e["school"] for e in field if e["heights"]})
    if len(hv) > 1:
        caveats.append(f"Heights runners appear under more than one team name ({', '.join(hv)}); they are scored separately.")
    if meta.get("scoreRuleNote"):
        caveats.append(meta["scoreRuleNote"])
    sb = {"standings": standings if len(standings) >= 2 else [], "partialTeams": partial, "caveats": caveats,
          "sheetNotes": meta.pop("_scoreNotes", [])}
    off = meta.pop("_official", None)
    if off:
        comp = {t["team"]: t for t in sb["standings"]}
        used, rows = set(), []
        for o in sorted(off, key=lambda o: int(o["Place"])):
            c = match_team(o["Team"], [x for x in comp if x not in used])
            t = comp.get(c)
            if c:
                used.add(c)
            ok = bool(t and t["rank"] == int(o["Place"]) and t["score"] == int(o["Score"]))
            rows.append({"team": o["Team"], "place": int(o["Place"]), "score": int(o["Score"]), "computedTeam": c,
                         "computedPlace": t["rank"] if t else None, "computedScore": t["score"] if t else None, "match": ok,
                         "heights": bool((t and t["heights"]) or HEIGHTS_RE.search(o["Team"]))})
            if t:
                t["official"] = {"place": int(o["Place"]), "score": int(o["Score"]), "match": ok}
        h = [r for r in rows if r["heights"]]
        sb["official"] = {"source": off[0]["SourceURL"], "teams": rows, "matched": sum(r["match"] for r in rows),
                          "heightsMatch": all(r["match"] for r in h) if h else None,
                          "extraComputed": [x for x in comp if x not in used]}
        bad = [r for r in rows if not r["match"]]
        if bad:
            ties = any(a["secs"] == b["secs"] and a["school"] != b["school"] for a, b in zip(field, field[1:]) if a["secs"])
            if missing > 0:
                sb["official"]["why"] = "The posted results are missing finishers, so the calculated scores can't match the official ones."
            elif ties and all(r["computedScore"] is not None and abs(r["computedScore"] - r["score"]) <= 2 for r in bad):
                sb["official"]["why"] = "The 1\u20132 point differences come from runners with identical times, whose order differs in the official scoring."
        if sb["official"]["matched"] < len(rows):
            quirks.append(f"{meta['season']} {meta['name']}: {sb['official']['matched']} of {len(rows)} official team scores reproduced.")
    meta.update({
        "scoring": sb,
        "fieldSize": len(field),
        "heightsCount": sum(e["heights"] for e in field),
        "hasPace": any(e["paceSecs"] for e in field),
        "hasGrade": any(e["grade"] for e in field),
        "results": field,
    })
    return meta


def main(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    wb_f = openpyxl.load_workbook(path, data_only=False)
    overrides = load_json("sheet_overrides.json", {}).get("races", {})
    manifest = load_json("manifest.json", {}).get("races", [])
    offscores = load_official_scores()
    meets, quirks, distance_log = [], [], []
    tabs = [ws for ws in wb if ws.title not in SKIP_TABS]
    used_overrides = set()
    for order, ws in enumerate(reversed(tabs)):
        title = ws.title
        info = dict(MEET_INFO.get(title, {}))
        if "season" not in info:
            m = re.search(r"(19|20)\d{2}", title)
            info["season"] = int(m.group(0)) if m else 0
            info["name"] = re.sub(r"(19|20)\d{2}\s*-?\s*", "", title).strip() or title
            quirks.append(f"Tab '{title}' has no MEET_INFO entry; name guessed.")
        rows = list(ws.iter_rows(values_only=True))
        divs = formula_distance(wb_f[title])
        f_dist = divs.most_common(1)[0][0] if divs else None
        if len(divs) > 1:
            quirks.append(f"Tab '{title}': pace formulas use more than one distance {dict(divs)}; using {f_dist}.")
        stated = info.get("distance")
        if f_dist and stated and abs(f_dist - stated) > 0.001:
            quirks.append(f"Tab '{title}': pace formula divides by {f_dist:g} mi but the sheet states {stated:g} mi; using the formula ({f_dist:g} mi).")
        score_notes = sheet_score_notes(ws)
        cols = find_cols(rows[0])
        if cols["time"] is None:
            quirks.append(f"Tab '{title}': no Time column found, skipped.")
            continue

        def get(r, k):
            i = cols[k]
            return r[i] if i is not None and i < len(r) else None

        races, cur, prev_place = [], [], 0
        for r in rows[1:]:
            if not any(c not in (None, "", " ") for c in r[:12]):
                continue
            place = num(get(r, "place"))
            if place is not None and place < prev_place and cur:
                races.append(cur); cur = []
            if place is not None:
                prev_place = place
            if cols["name"] is not None:
                name = get(r, "name")
            else:
                name = f"{str(get(r, 'first') or '').strip()} {str(get(r, 'last') or '').strip()}"
            cur.append(make_entry(get(r, "place"), get(r, "grade"), name, get(r, "school"), get(r, "time"), get(r, "pace")))
        if cur:
            races.append(cur)
        if len(races) > 1:
            quirks.append(f"Tab '{title}' contains {len(races)} separately-numbered races (placings restart); shown as Race 1..{len(races)}.")

        for ri, field in enumerate(races):
            ratios = [e["secs"] / e["paceSecs"] for e in field if e["secs"] and e["paceSecs"]]
            implied = round(statistics.median(ratios), 2) if ratios else None
            if f_dist:
                dist, dsrc = f_dist, "formula"
            elif stated:
                dist, dsrc = stated, "stated"
            elif implied:
                dist, dsrc = implied, "est"
            else:
                dist, dsrc = None, None
            label = info["name"] + (f" \u2013 Race {ri + 1}" if len(races) > 1 else "")
            mid = slug(f"{info['season']}-{info['name']}-{info.get('date', '')}" + (f"-race{ri + 1}" if len(races) > 1 else ""))
            ov = overrides.get(mid, {})
            if ov:
                used_overrides.add(mid)
            correction = None
            if ov.get("replace_field_with"):
                before = len(field)
                field = read_csv_field(ov["replace_field_with"])
                for e in field:
                    e["fromOfficial"] = True
                correction = ov.get("correction_note") or f"Replaced the sheet's {before} rows with the official results ({len(field)})."
            elif ov.get("fill_from_official"):
                n = fill_from_official(field, read_csv_field(ov["fill_from_official"]), title, quirks)
                correction = (ov.get("correction_note") or "") + ("" if n else " (nothing to add)")
                if n:
                    quirks.append(f"{info['season']} {label}: added {n} finishers from the official results.")
            off_d = ov.get("official_distance_mi")
            dnote = None
            if off_d and dist:
                if abs(off_d - dist) <= DIST_TOL:
                    dnote = {"official": off_d, "agrees": True}
                else:
                    dnote = {"official": off_d, "agrees": False}
                    quirks.append(f"{info['season']} {label}: sheet distance {dist:g} mi ({dsrc}) but the official results list {off_d:g} mi; the site uses the sheet's {dist:g} mi.")
            meta = {
                "id": mid, "season": info["season"], "name": label,
                "dateLabel": info.get("date"),
                "date": ov.get("official_date") or parse_md(info["season"], info.get("date")),
                "tab": title, "origin": "sheet",
                "distance": dist, "distanceSource": dsrc, "statedDistance": stated, "impliedDistance": implied,
                "officialDistance": dnote,
                "officialName": " \u2013 ".join(x for x in (ov.get("official_meet_name"), ov.get("official_race")) if x) or None,
                "location": ov.get("location"),
                "source": {"name": ov["source_name"], "url": ov["source_url"], "original": ov.get("original_url")} if ov.get("source_url") else None,
                "correction": correction,
                "scoreExcludeGrades": ov.get("score_exclude_grades"),
                "scoreRuleNote": ov.get("score_rule_note"),
                "_scoreNotes": score_notes if ri == 0 else [],
                "_official": offscores.get(ov.get("official_scores_meet")) if ov.get("official_scores_meet") else None,
                "_tabOrder": order * 10 + ri,
            }
            distance_log.append((title, ri, dist, dsrc, stated, off_d))
            meets.append(finalize(field, meta, quirks))
    for k in set(overrides) - used_overrides:
        quirks.append(f"sheet_overrides.json: no sheet race with id '{k}' (override unused).")

    # Races found online that aren't in the sheet
    for s in manifest:
        field = read_csv_field(s["csv"])
        race = s.get("race") or ""
        label = s["meet"] + (f" ({race})" if race and race not in ("Middle School Boys", "Boys") else "")
        mid = slug(f"{s['season']}-{s['meet']}-{race}-{s['date']}")
        meta = {
            "id": mid, "season": s["season"], "name": label, "dateLabel": None, "date": s["date"],
            "tab": None, "origin": "online",
            "distance": s.get("distance_mi"), "distanceSource": "official" if s.get("distance_mi") else None,
            "statedDistance": None, "impliedDistance": None, "officialDistance": None,
            "officialName": f"{s['meet']} \u2013 {race}" if race else s["meet"],
            "location": s.get("location"),
            "source": {"name": s.get("source_name"), "url": s.get("source_url"), "original": s.get("original_url")},
            "correction": None, "note": s.get("note"),
            "noDisplacerLimit": s.get("score_no_displacer_limit"), "scoreRuleNote": s.get("score_rule_note"),
            "_official": offscores.get(s.get("official_scores_meet")) if s.get("official_scores_meet") else None,
            "_tabOrder": 0,
        }
        meets.append(finalize(field, meta, quirks))

    # Chronological order: by date; sheet races without a date sit just after the previous dated race of the
    # same season in tab order.
    meets.sort(key=lambda m: (m["season"], m["_tabOrder"] if m["origin"] == "sheet" else 0))
    last = {}
    for m in [x for x in meets if x["origin"] == "sheet"]:
        if m["date"]:
            last[m["season"]] = m["date"]
        m["_sortDate"] = m["date"] or (last.get(m["season"], f"{m['season']}-09-01") + "~")
    for m in meets:
        m.setdefault("_sortDate", m["date"])
    meets.sort(key=lambda m: (m["season"], m["_sortDate"], m["_tabOrder"]))
    for i, m in enumerate(meets):
        m["order"] = i
        for k in ("_sortDate", "_tabOrder"):
            m.pop(k, None)
    seen = {}
    for m in meets:
        if m["id"] in seen:
            seen[m["id"]] += 1; m["id"] += f"-{seen[m['id']]}"
        else:
            seen[m["id"]] = 1

    runners, review = build_runners(meets)
    data = {
        "generated": datetime.date.today().isoformat(),
        "source": "Google Sheet \u201cHeights MS XC Results\u201d plus races found on racecast.co / Athletic.net / mocorunning.com",
        "meets": sorted(meets, key=lambda m: -m["order"]),
        "runners": sorted(runners, key=lambda r: (r["name"].split()[-1], r["name"], r["seasons"][0])),
        "identityReview": review,
        "quirks": EXTRA_NOTES + sorted(set(quirks)),
    }
    out = os.path.join(ROOT, "data")
    with open(os.path.join(out, "results.json"), "w") as f:
        json.dump(data, f, indent=1)
    with open(os.path.join(out, "results.js"), "w") as f:
        f.write("// Generated by scripts/build_data.py from the Heights MS XC Results sheet + data/supplemental. Do not edit by hand.\n")
        f.write("window.XC_DATA = " + json.dumps(data, separators=(",", ":")) + ";\n")
    print(f"{len(meets)} races ({sum(m['origin'] == 'online' for m in meets)} from online), {len(runners)} Heights runners, {sum(m['fieldSize'] for m in meets)} result rows")
    for q in data["quirks"]:
        print("quirk:", q)
    for t, ri, d, src, st, od in distance_log:
        print(f"distance: {t!r} race{ri + 1}: {d} ({src}){' stated ' + str(st) if st else ''}{' official ' + str(od) if od else ''}")
    for m in meets:
        o = m["scoring"].get("official")
        print("score:", m["id"], m["date"], m["distance"], m["distanceSource"], f"official {o['matched']}/{len(o['teams'])} heights={o['heightsMatch']}" if o else "no official")
    for r in review:
        print("review:", r)


def class_of(season, grade):
    return season + (8 - grade) if grade else None


def build_runners(meets):
    """Group Heights results by name, then split a name into separate runners when the seasons/grades
    can't belong to one middle-schooler (class year differs by 2+, or seasons more than 3 years apart)."""
    by_name = {}
    for m in meets:
        for e in m["results"]:
            if e["heights"]:
                by_name.setdefault(e["name"], []).append((m, e))
    runners, review = [], []
    for name, items in by_name.items():
        items.sort(key=lambda x: x[0]["order"])
        clusters = []
        for m, e in items:
            cy = class_of(m["season"], e["grade"])
            placed = None
            for c in clusters:
                cys = [y for y in c["cy"] if y]
                if cy and cys and abs(statistics.median(cys) - cy) >= 2:
                    continue
                if m["season"] - c["last"] > 3:
                    continue
                if not cy and not cys and m["season"] - c["first"] > 3:
                    continue
                placed = c; break
            if not placed:
                placed = {"first": m["season"], "last": m["season"], "cy": [], "items": []}
                clusters.append(placed)
            placed["items"].append((m, e)); placed["cy"].append(cy); placed["last"] = m["season"]
        for c in clusters:
            seasons = sorted({m["season"] for m, _ in c["items"]})
            rid = slug(name) if len(clusters) == 1 else slug(f"{name}-{seasons[0]}")
            sources = {m["origin"] for m, _ in c["items"]}
            r = {"name": name, "id": rid, "seasons": seasons, "results": []}
            if len(clusters) > 1:
                r["disambig"] = f"{seasons[0]}" + (f"\u2013{seasons[-1]}" if seasons[-1] != seasons[0] else "")
            for m, e in c["items"]:
                e["runnerId"] = rid
                r["results"].append({
                    "meet": m["id"], "place": e["place"], "hplace": e.get("hplace"), "grade": e["grade"],
                    "time": e["time"], "secs": e["secs"], "pace": e["pace"], "paceSecs": e["paceSecs"],
                    "suspect": e.get("suspect", False), "teamPlace": e.get("teamPlace"),
                    "paceCalc": e.get("paceCalc", False), "origin": m["origin"], "added": e.get("added", False),
                })
            paces = [x for x in r["results"] if x["paceSecs"] and not x["suspect"]]
            r["bestPace"] = min(paces, key=lambda x: x["paceSecs"])["pace"] if paces else None
            r["bestPaceSecs"] = min((x["paceSecs"] for x in paces), default=None)
            runners.append(r)
            if len(sources) > 1:
                cys = sorted({y for y in c["cy"] if y})
                if not cys and seasons[-1] - seasons[0] >= 2:
                    review.append(f"Merged sheet + online results for {name} ({', '.join(map(str, seasons))}) with no grades to confirm it's one runner.")
        if len(clusters) > 1:
            review.append(f"Split \u201c{name}\u201d into {len(clusters)} runners: " + "; ".join(
                f"{c['first']}\u2013{c['last']} (class of {', '.join(str(y) for y in sorted({y for y in c['cy'] if y})) or '?'})" for c in clusters))
    # possible spelling variants that were NOT merged (same last name, similar first name, overlapping seasons)
    rs = sorted(runners, key=lambda r: r["name"])
    for i, a in enumerate(rs):
        for b in rs[i + 1:]:
            if a["name"] == b["name"]:
                continue
            la, lb = a["name"].split()[-1].lower(), b["name"].split()[-1].lower()
            fa, fb = a["name"].split()[0].lower(), b["name"].split()[0].lower()
            same_last = la == lb or difflib.SequenceMatcher(None, la, lb).ratio() >= 0.85
            close_first = fa == fb or fa.startswith(fb[:3]) or fb.startswith(fa[:3]) or difflib.SequenceMatcher(None, fa, fb).ratio() >= 0.75
            overlap = min(a["seasons"][-1], b["seasons"][-1]) - max(a["seasons"][0], b["seasons"][0]) >= -1
            if same_last and close_first and overlap and (la != lb or fa != fb):
                review.append(f"Possible same runner, not merged: \u201c{a['name']}\u201d ({', '.join(map(str, a['seasons']))}) and \u201c{b['name']}\u201d ({', '.join(map(str, b['seasons']))}).")
    return runners, review


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "Heights_MS_XC_Results.xlsx")
