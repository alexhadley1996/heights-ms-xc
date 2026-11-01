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
import sys, re, json, datetime, os, statistics, csv, difflib, collections, itertools

import openpyxl

# Tabs that are analysis/comparison sheets, not meet results.
SKIP_TABS = {"2026 Comparison", "Comparison 2025 - 6th Grade"}

# Per-tab metadata.  Only facts that appear in the sheet (tab name / header notes) are used.
# distance_mi: from a "Distance:" note in the tab header or the tab name; None if the sheet doesn't say.
MEET_INFO = {
    "2026 States":               dict(season=2026, name="State Championship", distance=None),
    "2026 - Bullis":             dict(season=2026, name="Bullis", distance=None),
    "2026 - Flint Hill":         dict(season=2026, name="Flint Hill", distance=2.0),
    "2025 States":               dict(season=2025, name="State Championship", distance=None),
    "Flint Hill 2025":           dict(season=2025, name="Flint Hill", distance=None),
    "2025 Bittersweet #2":       dict(season=2025, name="Bittersweet #2", distance=None),
    "2025 Bittersweet #1":       dict(season=2025, name="Bittersweet #1", distance=None),
    "2024 State Championship":   dict(season=2024, name="State Championship", distance=None),
    "2024 - Flint Hill":         dict(season=2024, name="Flint Hill", distance=None),
    "2024 - Landon 1.6 miles":   dict(season=2024, name="Landon", distance=1.6),
    "2024 - Gtown Prep":         dict(season=2024, name="Georgetown Prep", distance=2.0),
    "2024 - Bullis":             dict(season=2024, name="Bullis", distance=2.0),
    "2023 State Championship":   dict(season=2023, name="State Championship", distance=2.0),
    "2023 - Bullis":             dict(season=2023, name="Bullis", distance=1.8),
    "2023 Flint Hill":           dict(season=2023, name="Flint Hill", distance=2.1),
    "2023 - Covenant Life":      dict(season=2023, name="Covenant Life", distance=2.0),
    # The tab's second list is a B-team re-tally of the same race; it is ignored (keep_races=1).
    "2022 - States":             dict(season=2022, name="State Championship", distance=2.1, keep_races=1),
    "2022 - CL":                 dict(season=2022, name="Covenant Life", distance=2.0),
    "2021 - States at Landon":   dict(season=2021, name="State Championship (at Landon)", distance=2.2),
    # 2010 tabs have no pace column/formula; course length per the coach (team) is 2.2 mi.
    "2010 - Landon #2":          dict(season=2010, name="Landon #2", distance=2.2, dist_from_team=True),
    "2010 - Landon #1":          dict(season=2010, name="Landon #1", distance=2.2, dist_from_team=True),
    "2009 - Landon, Nov 3":      dict(season=2009, name="Landon", date="Nov 3", distance=None),
    "2009 - Landon, Sept 28":    dict(season=2009, name="Landon", date="Sept 28", distance=None),
    "2008 - Landon, Nov 4":      dict(season=2008, name="Landon", date="Nov 4", distance=None),
    "2008 - Landon, Oct 17":     dict(season=2008, name="Landon", date="Oct 17", distance=None),
    "2008 - Landon, Oct 6":      dict(season=2008, name="Landon", date="Oct 6", distance=None),
    "2007 - Landon, October 18": dict(season=2007, name="Landon", date="Oct 18", distance=None),
    "2007 - Landon Invit":       dict(season=2007, name="Landon Invitational", distance=None),
    "2006 - Landon, October 18": dict(season=2006, name="Landon", date="Oct 18", distance=None),
    "2006 - Landon Invit":       dict(season=2006, name="Landon Invitational", distance=None),
    "2006 - Landon Home Meet Sep 15 ": dict(season=2006, name="Landon Home Meet", date="Sep 15", distance=None),
    "2005 - Landon MS Tri-Meet Sep 1": dict(season=2005, name="Landon MS Tri-Meet", date="Sep 16", distance=None),  # tab name truncated; Date column says 2005-09-16
    "2014 - Landon MS Invitational":   dict(season=2014, name="Landon MS Invitational", distance=None),
    "2014 - Landon MS Invitational (": dict(season=2014, name="Landon MS Invitational (Open race)", distance=None),
    "2015 - Landon MS Invitational -":  dict(season=2015, name="Landon MS Invitational", distance=2.19),
    "2015 - Landon MS Open - GPT":      dict(season=2015, name="Landon MS Invitational (Open race)", distance=2.19),
    "2019 - States at Landon":     dict(season=2019, name="State Championship (at Landon)", distance=2.2),
    "2019 - Norwood Invitational": dict(season=2019, name="Norwood Invitational", distance=2.1),   # team tab says "Boys (2.1 mi)"
    "2019 - Covenant Life":        dict(season=2019, name="Covenant Life", distance=2.0),
}
# Chronological order within a season, when known from the sheet (tab order is newest-first).
# Otherwise tab order (reversed) is used.

# Spelling variants of the same Heights runner -> canonical name.
# Canonical = the spelling most used in official results (racecast / Athletic.net / mocorunning), user-confirmed merges.
NAME_ALIASES = {
    "Sean Carney (real)": "Sean Carney",
    "Johnny Ortman": "John Ortman",      # official results: "John Ortman" in 8 meets, "Johnny" in 1
    "Louie Liess": "Louie Leiss",        # official: "Leiss" (2 meets), never "Liess"
    "P Allen": "Philip Allen",
    "Phillip Allen": "Philip Allen",     # official: "Philip" 10 meets (mostly HS), "Phillip" 7
    "Vince Kiernan": "Vince Keirnan",    # official: "Keirnan" 3 meets (2008), "Kiernan" 2 (2007)
    "Santi Luli": "Santiago Lulli",      # nickname; full first name kept; family name is Lulli (user-confirmed)
    "Andres Luli": "Andres Lulli",       # official: "Lulli" 8 meets (2009 + HS), "Luli" 3 (2008)
    "Justin Liess": "Justin Leiss",      # user-confirmed (same family as Louie Leiss)
    "Colin Oneil": "Cullun O'Neill",     # user-confirmed same boy; official results: "Cullun O'Neill" 2 meets (2006), "Colin Oneil" 1 (2008)
    "Ed Nelson": "Edward Nelson",
    "Henny Mangold": "Henry Mangold",    # 2005 Tri-Meet typo (user-confirmed)        # nickname; full first name kept
}

SILENT_FIXES = {"Henny Mangold"}  # plain typos: corrected without a "Listed as" note

CONFIRMED_SAME = {"Edward Nelson", "Cullun O'Neill"}  # user-confirmed single runner across sources

# Hand-written notes about the sheet, shown on the Meets page.
EXTRA_NOTES = [
    "2024 Bullis: no grades recorded; several non-Heights runners are listed by first name only.",
    "Heights place is calculated from the School column in every race (some tabs number every finisher in their Heights Place column).",
    "Some 2008 Landon (Oct 6) and 2023 Flint Hill times were stored as clock times (e.g. 13:36 read as 13 h 36 min) and are shown as minutes:seconds.",
    "Grades weren't recorded for most 2006\u20132010 meets, 2022 Covenant Life, 2023 Flint Hill and 2024 Flint Hill/Bullis. For Heights runners the site fills a missing grade from the same runner's other races (same season, or the grade in another season adjusted by the year difference) and shows it in italics.",
    "Race distances come from the divisor in each tab's pace formula (pace = time \u00f7 miles). The 2010 Landon tabs have no pace formula; per the team those courses were 2.2 mi, so their paces are calculated from that.",
    "Where a tab has no pace for a runner, the site calculates it the same way: time \u00f7 the race distance.",
    "Team scores are calculated by the site from the listed finishers (standard scoring: teams with 5+ finishers, top 5 score, 6th\u20137th displace, 8th+ ignored, ties broken by the 6th runner) and compared with official meet scores where those were found online.",
    "Races from 2011\u20132019 that aren't in the sheet, plus 2021 CLS and 2022 FHS, come from racecast.co / Athletic.net / mocorunning.com (data/supplemental). Each of those meet pages links its source.",
    "2022 States: the meet re-tallied extra runners from deep teams as B teams; the B listing is left out, those runners keep their overall places but don't count in team scores. 2021 States and 2007 Landon Invitational were missing finishers, which were added from the official results. 2025 States did not score 5th graders.",
    "Name spellings merged for the runner pages (shown with the spelling used most in official results): John Ortman (Johnny), Louie Leiss (Liess), Philip Allen (Phillip, P), Vince Keirnan (Kiernan), Santiago Lulli (Santi Luli), Andres Lulli (Luli), Daniel Lulli (Luli), Justin Leiss (Liess), Cullun O'Neill (Colin Oneil), Edward Nelson (Ed).",
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
PACE_METERS_RE = re.compile(r"/\s*\(\s*([\d.]+)\s*/\s*1609\.344\s*\)\s*\)\s*/\s*86400")
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
                    continue
                m = PACE_METERS_RE.search(v)  # ".../(3000/1609.344))/86400" -> metres to miles
                if m:
                    c[round(float(m.group(1)) / 1609.344, 3)] += 1
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


def score_race(field, exclude_grades=(), max_per_team=7, incomplete_score=False, by_overall=False, teams_only=None,
               exact_school=False, partial_ok=None):
    """Standard XC scoring. Returns (standings, partial_teams). Runners in exclude_grades keep their place but don't score.
    Meet-specific rules (HS overrides; all off by default): incomplete_score = runners of teams without 5 finishers
    keep their team places; by_overall = team score is the sum of overall places; teams_only = only these team keys
    score (everyone else runs as an individual); exact_school = teams split by the exact listed name (A/B entries);
    partial_ok = team keys the official results show as scoring teams although fewer than 5 of their runners are listed
    (their listed runners take team places, but the team itself gets no calculated score)."""
    key = (lambda s: s.strip()) if exact_school else team_key
    for e in field:
        e.pop("teamPlace", None)
    finishers = [e for e in field if isinstance(e["place"], int) and e["secs"] is not None and e.get("grade") not in exclude_grades and not e.get("noScore")]
    finishers.sort(key=lambda e: e["place"])
    if teams_only is not None:
        finishers = [e for e in finishers if team_key(e["school"]) in teams_only]
    by_team = {}
    for e in finishers:
        if NON_TEAM_RE.match(e["school"].strip()):
            continue
        by_team.setdefault(key(e["school"]), []).append(e)
    full = {k: v for k, v in by_team.items() if len(v) >= 5}
    placed = by_team if incomplete_score else dict(full, **{k: v for k, v in by_team.items()
                                                             if partial_ok and team_key(v[0]["school"]) in partial_ok})
    scorers = set()
    for k, v in placed.items():
        for e in v[:max_per_team]:
            scorers.add(id(e))
    tp = 0
    for e in finishers:
        if id(e) in scorers:
            tp += 1
            e["teamPlace"] = e["place"] if by_overall else tp
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
    if heights:
        new = NAME_ALIASES.get(name, name)
        new = re.sub(r"\bLuli$", "Lulli", new)  # the Lulli family name is always "Lulli" (user-confirmed)
        if new != name:
            if name not in SILENT_FIXES:
                e["note"] = f"Listed as \u201c{name}\u201d"
            e["name"] = new
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
    off_by_name = {norm_name(o["name"]): o for o in official}
    for e in field:
        if e["place"] is None and e["secs"] is not None:
            o = off_by_name.get(norm_name(e["name"]))
            if o and isinstance(o["place"], int):
                e["place"] = o["place"]
                if o.get("grade") and not e.get("grade"):
                    e["grade"] = o["grade"]
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


def place_unplaced(field, info, label, quirks):
    """Runners with a time but no place (added by hand, e.g. no bib) are slotted in by time and
    the overall places renumbered from there."""
    loose = [e for e in field if e["place"] is None and e["secs"] is not None]
    if not loose or not any(isinstance(e["place"], int) for e in field):
        return
    for e in loose:
        field.remove(e)
        i = next((k for k, x in enumerate(field) if isinstance(x["place"], int) and x["secs"] is not None and x["secs"] > e["secs"]), len(field))
        field.insert(i, e)
        e["note"] = "Ran without a bib; added by hand in the sheet and placed by his time"
        e["handPlaced"] = True
    n = 0
    for e in field:
        if isinstance(e["place"], int) or e.get("handPlaced"):
            n += 1
            if e["place"] != n:
                e.setdefault("sheetPlace", e["place"])
            e["place"] = n


def mark_b_team(field, a_list, b_list, quirks, label):
    """Meets sometimes re-tally the same race for 'B teams'. Runners on the official B list (and not on the A list)
    keep their overall place but don't count in the (A) team scores."""
    def same(e, o):
        return o["secs"] is not None and e["secs"] is not None and abs(o["secs"] - e["secs"]) <= 0.15 and \
            (norm_name(o["name"]) == norm_name(e["name"]) or difflib.SequenceMatcher(None, norm_name(o["name"]), norm_name(e["name"])).ratio() >= 0.6)
    n = 0
    for e in field:
        if any(same(e, o) for o in b_list) and not any(same(e, o) for o in a_list):
            e["noScore"] = True
            e["noScoreWhy"] = "Scored for a B team at this meet, so not counted in the team scores"
            n += 1
    quirks.append(f"{label}: {n} runners were scored for B teams; they keep their places but don't count in team scores.")
    # runners with identical times: follow the official finish order
    idx = {}
    for e in field:
        k = next((i for i, o in enumerate(a_list) if same(e, o)), None)
        if k is not None:
            idx[id(e)] = k
    i = 0
    while i < len(field):
        j = i
        while j + 1 < len(field) and field[j + 1]["secs"] is not None and field[j + 1]["secs"] == field[i]["secs"]:
            j += 1
        grp = field[i:j + 1]
        if len(grp) > 1 and all(id(e) in idx for e in grp):
            places = [e["place"] for e in grp]
            new = sorted(grp, key=lambda e: idx[id(e)])
            if new != grp:
                for e, pl in zip(new, places):
                    e["place"] = pl
                field[i:j + 1] = new
                quirks.append(f"{label}: tied runners at {grp[0]['time']} ({', '.join(e['name'] for e in new)}) put in the official finish order.")
        i = j + 1


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
    for e in field:
        if e.get("grade") in excl and isinstance(e["place"], int):
            e["noScore"] = True
            e["noScoreWhy"] = "This meet did not count this grade in team scores"
    standings, partial = score_race(field, excl, 10 ** 6 if meta.get("noDisplacerLimit") else 7,
                                    incomplete_score=bool(meta.get("scoreIncomplete")), by_overall=bool(meta.get("scoreByOverall")),
                                    teams_only=meta.pop("_scoreTeamsOnly", None), exact_school=bool(meta.get("scoreExactSchool")),
                                    partial_ok=meta.pop("_scorePartialOk", None))
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
    def is_team_tab(ws):
        h = [str(c or "").strip().lower() for c in next(ws.iter_rows(max_row=1, values_only=True), ())][:4]
        return h[:4] == ["race", "rank", "team", "score"]
    team_tabs = [ws.title for ws in wb if ws.title not in SKIP_TABS and is_team_tab(ws)]
    tabs = [ws for ws in wb if ws.title not in SKIP_TABS and ws.title not in team_tabs]
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
        if info.get("keep_races"):
            races = races[:info["keep_races"]]
        if len(races) > 1:
            quirks.append(f"Tab '{title}' contains {len(races)} separately-numbered races (placings restart); shown as Race 1..{len(races)}.")

        for ri, field in enumerate(races):
            ratios = [e["secs"] / e["paceSecs"] for e in field if e["secs"] and e["paceSecs"]]
            implied = round(statistics.median(ratios), 2) if ratios else None
            if f_dist:
                dist, dsrc = f_dist, "formula"
            elif stated:
                dist, dsrc = stated, "team" if info.get("dist_from_team") else "stated"
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
            place_unplaced(field, info, label, quirks)
            if ov.get("b_team_list"):
                mark_b_team(field, read_csv_field(ov["a_team_list"]), read_csv_field(ov["b_team_list"]), quirks, f"{info['season']} {label}")
                correction = ov.get("correction_note")
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
            if ov.get("replaces_supplemental"):
                # the sheet version wins; fill grades the sheet doesn't have from the online copy of the same race
                onl = {o["place"]: o for o in read_csv_field(ov["replaces_supplemental"]) if isinstance(o["place"], int)}
                ng = 0
                for e in field:
                    o = onl.get(e["place"])
                    if not e["grade"] and o and o["grade"] and o["secs"] and e["secs"] and abs(o["secs"] - e["secs"]) <= 0.6:
                        e["grade"] = o["grade"]; ng += 1
                if ng:
                    quirks.append(f"{info['season']} {label}: the sheet has no grade for {ng} runners; filled from the official results of the same race.")
            off_d = ov.get("official_distance_mi")
            dnote = None
            if off_d and dist:
                if abs(off_d - dist) <= DIST_TOL:
                    dnote = {"official": off_d, "agrees": True}
                else:
                    dnote = {"official": off_d, "agrees": False}
                    quirks.append(f"{info['season']} {label}: {'course distance (per the team)' if dsrc == 'team' else 'sheet distance'} {dist:g} mi{'' if dsrc == 'team' else f' ({dsrc})'} but the official results list {off_d:g} mi; the site uses {dist:g} mi.")
            meta = {
                "id": mid, "season": info["season"], "name": label, "division": "MS", "level": "Open", "levelLabel": "Middle school",
                "dateLabel": info.get("date"),
                "date": ov.get("official_date") or parse_md(info["season"], info.get("date")),
                "tab": title, "origin": "sheet",
                "distance": dist, "distanceSource": dsrc, "statedDistance": stated, "impliedDistance": implied,
                "officialDistance": dnote,
                "officialName": " \u2013 ".join(x for x in (ov.get("official_meet_name"), ov.get("official_race")) if x) or None,
                "location": ov.get("location"),
                "source": {"name": ov["source_name"], "url": ov["source_url"], "original": ov.get("original_url")} if ov.get("source_url") else None,
                "correction": correction,
                "note": ov.get("meet_note"),
                "scoreExcludeGrades": ov.get("score_exclude_grades"),
                "scoreRuleNote": ov.get("score_rule_note"),
                "noDisplacerLimit": ov.get("score_no_displacer_limit"),
                "_scoreNotes": score_notes if ri == 0 else [],
                "_official": offscores.get(ov.get("official_scores_meet")) if ov.get("official_scores_meet") else None,
                "_tabOrder": order * 10 + ri,
            }
            distance_log.append((title, ri, dist, dsrc, stated, off_d))
            meets.append(finalize(field, meta, quirks))
    for k in set(overrides) - used_overrides:
        quirks.append(f"sheet_overrides.json: no sheet race with id '{k}' (override unused).")

    # Races found online that aren't in the sheet
    replaced = {ov["replaces_supplemental"] for ov in overrides.values() if ov.get("replaces_supplemental")}
    for s in manifest:
        if s["csv"] in replaced:
            continue  # the sheet now has this race; the sheet version wins
        field = read_csv_field(s["csv"])
        race = s.get("race") or ""
        label = s["meet"] + (f" ({race})" if race and race not in ("Middle School Boys", "Boys") else "")
        mid = slug(f"{s['season']}-{s['meet']}-{race}-{s['date']}")
        meta = {
            "id": mid, "season": s["season"], "name": label, "division": "MS", "level": "Open", "levelLabel": "Middle school",
            "dateLabel": None, "date": s["date"],
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

    # High school races (data/hs), division "HS"
    import hs_data
    meets += hs_data.build_hs_meets(quirks)

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

    apply_row_splits(meets, quirks)
    runners, review = build_runners(meets)
    check_one_row_per_runner(meets, runners)
    write_outputs(meets, runners, review, quirks)
    ms = [m for m in meets if m["division"] == "MS"]
    gs = collections.Counter(m["division"] for m in meets if m["group"] == m["id"] or True)
    ng = collections.Counter((m["division"], m["group"]) for m in meets)
    print(f"Meet pages: {len(ng)} ({sum(1 for d, _ in ng if d == 'HS')} HS, {sum(1 for d, _ in ng if d == 'MS')} MS) for {len(meets)} races; "
          f"{len(build_runners.aliases)} old runner ids redirect")
    print(f"MS: {len(ms)} races ({sum(m['origin'] == 'online' for m in ms)} from online), "
          f"{sum(1 for r in runners if r['ms'])} Heights runners, {sum(m['heightsCount'] for m in ms)} Heights finishes")
    hs = [m for m in meets if m["division"] == "HS"]
    print(f"HS: {len(hs)} races, {sum(1 for r in runners if r['hs'])} Heights runners, {sum(m['heightsCount'] for m in hs)} Heights finishes; "
          f"both divisions: {sum(1 for r in runners if r['ms'] and r['hs'])} runners")
    for q in EXTRA_NOTES + sorted(set(quirks)):
        print("quirk:", q)
    for t, ri, d, src, st, od in distance_log:
        print(f"distance: {t!r} race{ri + 1}: {d} ({src}){' stated ' + str(st) if st else ''}{' official ' + str(od) if od else ''}")
    for m in meets:
        o = m["scoring"].get("official")
        print("score:", m["division"], m["id"], m["date"], m["distance"], m["distanceSource"], f"official {o['matched']}/{len(o['teams'])} heights={o['heightsMatch']}" if o else "no official")
    for r in review:
        print("review:", r)


def slim_scoring(m):
    """What the list pages need about team scores, without the full standings (those live in the per-race file)."""
    sc = m["scoring"]
    st, off = sc["standings"], sc.get("official")
    out = {"teamCount": len(off["teams"]) if off else len(st), "calcCount": len(st)}
    teams = []
    if off:
        out["official"] = {"matched": off["matched"], "n": len(off["teams"]), "heightsMatch": off["heightsMatch"]}
        for o in off["teams"]:
            if o["heights"]:
                teams.append({"team": o["team"], "place": o["place"], "score": o["score"], "official": True,
                              "calcPlace": o["computedPlace"], "calcScore": o["computedScore"], "match": o["match"]})
    if not teams:
        for t in st:
            if t["heights"]:
                teams.append({"team": t["team"], "place": t["rank"], "score": t["score"], "official": False})
    out["heights"] = teams
    return out


LEVEL_ORDER = {"Varsity": 0, "JV": 1, "Freshman": 2, "Open": 3, "Section": 4, "Relay": 5}
# The same meet listed under two names (one race from racecast, one from mocorunning)
MEET_GROUP_ALIASES = {(2011, "MD-DC Private School Championships"): "Maryland & DC Private School Championships"}
SEASON_NOTES = {2020: "2020 was the COVID-19 season: most invitationals and championships were cancelled, and teams ran "
                      "small dual or tri meets with few or no published results, so very few races are on record."}


def assign_groups(meets):
    """One meet page per meet: races of the same meet, date (or weekend) and division share a group.
    Returns the group list; each race gets m["group"] (a one-race meet keeps its race id as the group id)."""
    buckets = collections.defaultdict(list)
    for m in meets:
        if m["division"] == "HS":
            name = MEET_GROUP_ALIASES.get((m["season"], m["meetName"]), m["meetName"])
            key = ("HS", m["season"], name)
        else:
            name = re.sub(r"\s*\((open|b) race\)\s*$", "", m["name"], flags=re.I).strip()
            key = ("MS", m["season"], name, m.get("date") or m["id"])
        buckets[key].append(m)
    race_ids = {m["id"] for m in meets}
    groups, used = [], set()
    for key, ms in buckets.items():
        div, season, name = key[0], key[1], key[2]
        ms.sort(key=lambda m: (LEVEL_ORDER.get(m.get("level"), 9) if div == "HS" else (m["name"] != name), m["order"]))
        if len(ms) == 1:
            gid = ms[0]["id"]
        elif div == "MS":
            gid = ms[0]["id"]  # the main race's id: its old links already land on this page
        else:
            gid = slug(f"hs-{season}-{name}")
            while gid in race_ids or gid in used:
                gid += "-meet"
        used.add(gid)
        dates = sorted({m["date"] for m in ms if m.get("date")})
        g = {"id": gid, "season": season, "division": div, "name": name, "date": dates[0] if dates else None,
             "endDate": dates[-1] if len(dates) > 1 else None, "dateLabel": ms[0].get("dateLabel"),
             "location": next((m["location"] for m in ms if m.get("location")), None),
             "order": min(m["order"] for m in ms), "races": [m["id"] for m in ms],
             "levels": list(dict.fromkeys(m.get("level") for m in ms))}
        for m in ms:
            m["group"] = gid
        groups.append({k: v for k, v in g.items() if v is not None})
    return sorted(groups, key=lambda g: -g["order"])


def write_outputs(meets, runners, review, quirks):
    out = os.path.join(ROOT, "data")
    fdir = os.path.join(out, "races")
    os.makedirs(fdir, exist_ok=True)
    for f in os.listdir(fdir):
        if f.endswith(".js"):
            os.remove(os.path.join(fdir, f))
    groups = assign_groups(meets)
    core = []
    for m in sorted(meets, key=lambda m: -m["order"]):
        field = {"results": [{k: v for k, v in e.items() if v not in (None, False) or k in ("place", "time")} for e in m["results"]],
                 "scoring": m["scoring"]}
        if m.get("relay"):
            field["relay"] = m["relay"]
        with open(os.path.join(fdir, f"{m['id']}.js"), "w") as f:
            f.write(f"(window.XC_FIELD = window.XC_FIELD || {{}})[{json.dumps(m['id'])}] = " + json.dumps(field, separators=(",", ":")) + ";\n")
        c = {k: v for k, v in m.items() if k not in ("results", "scoring", "relay", "csv", "aliasIds") and v not in (None, False, [], "")}
        if c.get("notes"):
            c.pop("note", None)  # HS races: the list of notes replaces the joined text
        c["scoring"] = slim_scoring(m)
        core.append(c)
    data = {
        "generated": datetime.date.today().isoformat(),
        "source": "Google Sheet \u201cHeights MS XC Results\u201d (middle school), races found on racecast.co / Athletic.net / mocorunning.com (middle school), and high school results from racecast.co and mocorunning.com (data/hs)",
        "meets": core,
        "groups": groups,
        "seasonNotes": SEASON_NOTES,
        "runnerAliases": getattr(build_runners, "aliases", {}),
        "raceAliases": {a: m["id"] for m in meets for a in (m.get("aliasIds") or []) if a != m["id"]},
        "runners": sorted(runners, key=lambda r: (r["name"].split()[-1], r["name"], r["seasons"][0])),
        "identityReview": review,
        "quirks": EXTRA_NOTES + sorted(set(quirks)),
    }
    with open(os.path.join(out, "results.json"), "w") as f:
        json.dump(data, f, indent=1)
    with open(os.path.join(out, "results.js"), "w") as f:
        f.write("// Generated by scripts/build_data.py from the Heights MS XC Results sheet + data/supplemental + data/hs. Do not edit by hand.\n"
                "// Full results for each race are in data/races/<race id>.js, loaded by the race page.\n")
        f.write("window.XC_DATA = " + json.dumps(data, separators=(",", ":")) + ";\n")


def class_of(season, grade):
    """Graduation year (works for both middle and high school grades)."""
    return season + (12 - grade) if grade else None


def _cluster(name, items):
    """Split one name's results into separate people when seasons/grades can't belong to one runner."""
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
    return clusters


# A result row that belongs to a different Heights runner with the same name as a teammate (the team says these are
# brothers). Keyed by (race id, place, listed name) -> (tag for the runner id, label shown after his name).
ROW_SPLITS = {
    ("hs-2019-landon-invitational-varsity-b-division", 66, "Robert Spence"):
        ("2019-landon", "the second Robert Spence at the 2019 Landon Invitational"),
    ("2016-norwood-halloween-run-2016-10-27", 101, "Noah Schneider"):
        ("2016-norwood", "the second Noah Schneider at the 2016 Norwood Halloween Run"),
}


def apply_row_splits(meets, quirks):
    found = set()
    for m in meets:
        for e in m["results"]:
            k = (m["id"], e["place"], e["name"])
            if k in ROW_SPLITS:
                e["person"], e["personLabel"] = ROW_SPLITS[k]
                e["note"] = ((e.get("note") + "; ") if e.get("note") else "") + (
                    "Another Heights runner listed under the same name in this race (a brother, per the team); "
                    "kept as a separate runner.")
                found.add(k)
    for k in set(ROW_SPLITS) - found:
        quirks.append(f"ROW_SPLITS entry not found: {k}")


def check_one_row_per_runner(meets, runners):
    """Build-time check: no runner may have two results in the same race."""
    bad = []
    for r in runners:
        c = collections.Counter(x["meet"] for x in r["results"])
        bad += [f"{r['id']} \u00d7{n} in {mid}" for mid, n in c.items() if n > 1]
    for m in meets:
        c = collections.Counter(e.get("runnerId") for e in m["results"] if e.get("runnerId"))
        bad += [f"{rid} \u00d7{n} in {m['id']} (field)" for rid, n in c.items() if n > 1]
    if bad:
        sys.exit("ERROR: a runner appears more than once in a race:\n  " + "\n  ".join(sorted(set(bad))))
    print(f"Check passed: every runner appears at most once in each of {len(meets)} races.")


def build_runners(meets):
    """One runner per person. Middle and high school results are grouped by name separately (same rules as before),
    then joined into one person using the MS-HS matches in data/hs (hs_data.load_matches)."""
    import hs_data
    by_name, forms = {}, {}
    for m in meets:
        for e in m["results"]:
            if e["heights"]:
                k = (m["division"], norm_name(e["name"]))  # "Nick O'neil" == "Nick O'Neil"
                by_name.setdefault(k, []).append((m, e))
                forms.setdefault(k, collections.Counter())[e["name"]] += 1
    case_merges = []
    groups = {}
    for k, items in by_name.items():
        best = max(forms[k], key=lambda f: (forms[k][f], sum(c.isupper() for c in f), "'" in f))
        if len(forms[k]) > 1:
            case_merges.append(f"Merged capitalization/punctuation variants ({k[0]}): {', '.join(sorted(forms[k]))} \u2192 {best}.")
        for m, e in items:
            if e["name"] != best:
                e.setdefault("note", f"Listed as \u201c{e['name']}\u201d")
                e["name"] = best
        groups[(k[0], best)] = items
    review = []
    people = []  # noqa  # {name, div, items, cy, id}
    for div in ("MS", "HS"):
        names = [n for d, n in groups if d == div]
        for name in names:
            clusters = _cluster(name, groups[(div, name)])
            if div == "HS" and len(clusters) > 1 and name not in hs_data.KEEP_SPLIT:
                # A high school grade that disagrees (e.g. "12" for a freshman) is usually a source typo. Re-join
                # same-name pieces that never ran on the same day and whose seasons touch; keep real namesakes apart.
                merged = True
                while merged and len(clusters) > 1:
                    merged = False
                    for a, b in itertools.combinations(clusters, 2):
                        da = {(m["season"], m["date"]) for m, _ in a["items"]}
                        db = {(m["season"], m["date"]) for m, _ in b["items"]}
                        if not (da & db) and max(a["first"], b["first"]) - min(a["last"], b["last"]) <= 1:
                            a["items"] += b["items"]; a["cy"] += b["cy"]
                            a["first"], a["last"] = min(a["first"], b["first"]), max(a["last"], b["last"])
                            clusters.remove(b); merged = True
                            review.append(f"Kept \u201c{name}\u201d (HS) as one runner although the listed grades disagree "
                                          f"(class of {', '.join(str(y) for y in sorted(set(y for y in a['cy'] if y)))}); never ran on the same day, so likely a grade typo.")
                            break
            # rows marked as a different runner with the same name (ROW_SPLITS) become their own person
            tagged = collections.defaultdict(list)
            for c in clusters:
                keep = []
                for it in c["items"]:
                    (tagged[it[1]["person"]] if it[1].get("person") else keep).append(it)
                c["items"] = keep
            clusters = [c for c in clusters if c["items"]]
            for c in clusters:
                seasons = sorted({m["season"] for m, _ in c["items"]})
                pid = slug(name) if len(clusters) == 1 else slug(f"{name}-{seasons[0]}")
                p = {"name": name, "div": div, "items": c["items"], "cy": [y for y in c["cy"] if y], "id": pid, "seasons": seasons,
                     "split": len(clusters) > 1 or bool(tagged)}
                people.append(p)
                if div == "MS":
                    srcs = {m["origin"] for m, _ in c["items"]}
                    if len(srcs) > 1 and not p["cy"] and seasons[-1] - seasons[0] >= 2 and name not in CONFIRMED_SAME:
                        review.append(f"Merged sheet + online results for {name} ({', '.join(map(str, seasons))}) with no grades to confirm it's one runner.")
            for tag, its in tagged.items():
                seasons = sorted({m["season"] for m, _ in its})
                cy = [m["season"] + 12 - e["grade"] for m, e in its if isinstance(e.get("grade"), int)]
                people.append({"name": name, "div": div, "items": its, "cy": cy, "id": slug(f"{name}-{tag}"), "seasons": seasons,
                               "split": True, "tagged": True, "disambig": its[0][1].get("personLabel")})
                review.append(f"Split \u201c{name}\u201d ({div}): {len(its)} result(s) belong to a different runner with the same name ({tag}).")
            if len(clusters) > 1:
                review.append(f"Split \u201c{name}\u201d ({div}) into {len(clusters)} runners: " + "; ".join(
                    f"{c['first']}\u2013{c['last']} (class of {', '.join(str(y) for y in sorted(set(y for y in c['cy'] if y))) or '?'})" for c in clusters))
    # MS ids exactly as before; HS-only ids get "-hs" if an MS runner already has that id
    ms_ids = collections.Counter(p["id"] for p in people if p["div"] == "MS")
    for p in people:
        if p["div"] == "MS" and ms_ids[p["id"]] > 1:
            review.append(f"Two runners share the id for {p['name']}; using {slug(p['name'] + '-' + str(p['seasons'][0]))}.")
            p["id"] = slug(f"{p['name']}-{p['seasons'][0]}")
    ms_by_id = {p["id"]: p for p in people if p["div"] == "MS"}
    hs_people = [p for p in people if p["div"] == "HS"]
    # join MS and HS careers
    matches = hs_data.load_matches()
    merged_hs = set()
    for ms_id, hs_names in matches.items():
        msp = ms_by_id.get(ms_id)
        if not msp:
            review.append(f"MS\u2013HS match: no MS runner with id {ms_id}; not merged.")
            continue
        for hn in sorted(hs_names):
            cands = [p for p in hs_people if p["name"] == hn and id(p) not in merged_hs and not p.get("tagged")]
            if not cands:
                if hn not in hs_data.HS_NAME_ALIASES:
                    review.append(f"MS\u2013HS match: no HS runner named {hn} (for {msp['name']}); not merged.")
                continue
            ms_cy = statistics.median(msp["cy"]) if msp["cy"] else None
            hp = min(cands, key=lambda p: abs(statistics.median(p["cy"]) - ms_cy) if (p["cy"] and ms_cy) else 0)
            if ms_cy and hp["cy"] and abs(statistics.median(hp["cy"]) - ms_cy) >= 2:
                review.append(f"MS\u2013HS match {msp['name']} / {hn}: class years differ ({ms_cy:g} vs {statistics.median(hp['cy']):g}); not merged.")
                continue
            merged_hs.add(id(hp))
            msp.setdefault("hsParts", []).append(hp)
    taken = set(ms_by_id)
    runners = []
    grade_log = {"filled": 0, "conflicts": [], "unresolved": 0}
    persons = [p for p in people if p["div"] == "MS"] + [p for p in hs_people if id(p) not in merged_hs]
    by_pid = {}
    for p in persons:
        by_pid.setdefault(p["id"], p)
    for grp in hs_data.CONFIRMED_SAME_PERSON:
        prim = by_pid.get(grp[0])
        if not prim:
            review.append(f"Confirmed merge: no runner with id {grp[0]}; not merged.")
            continue
        for oid in grp[1:]:
            o = by_pid.get(oid)
            if not o or o is prim:
                review.append(f"Confirmed merge: no runner with id {oid} (for {grp[0]}); skipped.")
                continue
            prim.setdefault("hsParts", []).extend([{**o, "hsParts": []}] + o.get("hsParts", []))
            persons.remove(o)
        prim["split"] = False
    aliases = {}
    for p in persons:
        items = list(p["items"])
        names = collections.Counter({p["name"]: len(p["items"])})
        for hp in p.get("hsParts", []):
            items += hp["items"]; names[hp["name"]] += len(hp["items"])
        items.sort(key=lambda x: x[0]["order"])
        rid = p["id"]
        if p["div"] == "HS" and rid in taken:
            rid = slug(f"{p['name']}-hs")
        taken.add(rid)
        for hp in p.get("hsParts", []):  # old page ids of the joined parts redirect here
            for old_id in (hp["id"], slug(f"{hp['name']}-hs")):
                if old_id != rid:
                    aliases.setdefault(old_id, rid)
        dup = collections.Counter(m["id"] for m, _ in items)
        for mid, n in dup.items():
            if n > 1:
                review.append(f"After merging, {p['name']} has {n} results in the same race ({mid}): "
                              + "; ".join(f"{e['name']} {e['time']} (place {e['place']})" for m, e in items if m["id"] == mid) + ". All kept.")
        # the name he's listed under most (an HS full name usually wins over an MS short form)
        name = max(names, key=lambda n: (names[n], len(n)))
        infer_grades(name, items, grade_log)
        seasons = sorted({m["season"] for m, _ in items})
        r = {"name": name, "id": rid, "seasons": seasons, "results": [],
             "ms": sorted({m["season"] for m, _ in items if m["division"] == "MS"}),
             "hs": sorted({m["season"] for m, _ in items if m["division"] == "HS"})}
        listed = [e["listed"] for _, e in items if e.get("listed")]
        if len(names) > 1 or listed:
            r["aka"] = list(dict.fromkeys([n for n in names if n != name] + [n for n in listed if n != name]))
        if p.get("disambig"):
            r["disambig"] = p["disambig"]
        elif p["split"]:
            r["disambig"] = f"{seasons[0]}" + (f"\u2013{seasons[-1]}" if seasons[-1] != seasons[0] else "")
        for m, e in items:
            e["runnerId"] = rid
            x = {"meet": m["id"], "place": e["place"], "hplace": e.get("hplace"), "grade": e["grade"],
                 "gradeInferred": e.get("gradeInferred", False),
                 "time": e["time"], "secs": e["secs"], "pace": e["pace"], "paceSecs": e["paceSecs"],
                 "suspect": e.get("suspect", False), "teamPlace": e.get("teamPlace"),
                 "paceCalc": e.get("paceCalc", False), "origin": m["origin"], "added": e.get("added", False)}
            if e.get("listed") and e["listed"] != name:
                x["as"] = e["listed"]
            elif e["name"] != name:
                x["as"] = e["name"]
            if e.get("dq"):
                x["dq"] = True
                x["note"] = e.get("note")
            if e.get("relay"):
                x["relay"] = e["relay"]
            r["results"].append({k: v for k, v in x.items() if v not in (None, False) or k in ("place", "time")})
        ok = [x for x, (m, e) in zip(r["results"], items) if x.get("paceSecs") and not x.get("suspect") and m["division"] == "MS"]
        r["bestPace"] = min(ok, key=lambda x: x["paceSecs"])["pace"] if ok else None
        r["bestPaceSecs"] = min((x["paceSecs"] for x in ok), default=None)
        hs5 = [x for x, (m, e) in zip(r["results"], items) if m["division"] == "HS" and m.get("distKey") == "5K" and x.get("secs") and not x.get("relay") and not x.get("suspect")]
        if hs5:
            b = min(hs5, key=lambda x: x["secs"])
            r["hs5k"] = {"time": b["time"], "secs": b["secs"], "meet": b["meet"]}
        hsp = [x for x, (m, e) in zip(r["results"], items) if m["division"] == "HS" and x.get("paceSecs") and not x.get("relay") and not x.get("suspect")]
        if hsp:
            r["hsBestPaceSecs"] = min(x["paceSecs"] for x in hsp)
        runners.append({k: v for k, v in r.items() if v is not None})
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
    for m in meets:
        m["hasGrade"] = any(e["grade"] for e in m["results"])
    review += case_merges
    review.append(f"Grades inferred for {grade_log['filled']} Heights results; {grade_log['unresolved']} still unknown.")
    review += grade_log["conflicts"]
    ids = {r["id"] for r in runners}
    build_runners.aliases = {k: v for k, v in aliases.items() if k not in ids}
    return runners, review


def infer_grades(name, items, log):
    """Fill a missing grade from the same runner's other races: same season exactly, otherwise
    known grade +/- the year difference. Only when every known grade agrees and the result fits the
    division (5-8 middle school, 9-12 high school)."""
    known = [(m["season"], e["grade"]) for m, e in items if e["grade"] and not e.get("gradeInferred")]
    by_season = {}
    for sn, g in known:
        by_season.setdefault(sn, set()).add(g)
    for sn, gs in by_season.items():
        if len(gs) > 1:
            log["conflicts"].append(f"Grade conflict: {name} is listed as grade {'/'.join(map(str, sorted(gs)))} in {sn}.")
    classes = {sn + (12 - g) for sn, g in known}
    if len(classes) > 1:
        log["conflicts"].append(f"Grade conflict: {name}'s grades don't line up across seasons ({', '.join(f'{sn}: gr {g}' for sn, g in sorted(set(known)))}); not inferring other seasons.")
    for m, e in items:
        if e["grade"]:
            continue
        sn = m["season"]
        g = None
        if len(by_season.get(sn, ())) == 1:
            g = next(iter(by_season[sn]))
        elif known and len(classes) == 1:
            g = 12 - (next(iter(classes)) - sn)
        lo, hi = (5, 8) if m["division"] == "MS" else (9, 12)
        if g and lo <= g <= hi:
            e["grade"], e["gradeInferred"] = g, True
            log["filled"] += 1
        else:
            log["unresolved"] += 1


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "Heights_MS_XC_Results.xlsx")
