#!/usr/bin/env python3
"""Build data/results.js (and results.json) from the "Heights MS XC Results" xlsx export.

Usage:  python3 scripts/build_data.py path/to/Heights_MS_XC_Results.xlsx
Requires: openpyxl.  Re-run after exporting a fresh copy of the Google Sheet (File > Download > .xlsx).
New meet tabs are picked up automatically if their header row uses the usual column names
(Place, Heights Place, Grade/Year, Name or First/Last Name, Pace, Time, School/Team).
Add an entry to MEET_INFO below to give a new tab a nicer name, date or distance.
"""
import sys, re, json, datetime, os, statistics

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
    "Distances are only stated for some meets; others show an estimate from time ÷ pace (marked 'est.'). 2026 Flint Hill says 2.1 mi but its paces work out to about 2.0 mi.",
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
    }
    # 2007 tabs have an unlabeled pace column right before Time
    if cols["pace"] is None and cols["time"] is not None and h[cols["time"] - 1] == "":
        cols["pace"] = cols["time"] - 1
    return cols


def main(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    meets, quirks = [], []
    tabs = [ws for ws in wb if ws.title not in SKIP_TABS]
    for order, ws in enumerate(reversed(tabs)):
        title = ws.title
        info = dict(MEET_INFO.get(title, {}))
        if "season" not in info:
            m = re.search(r"(19|20)\d{2}", title)
            info["season"] = int(m.group(0)) if m else 0
            info["name"] = re.sub(r"(19|20)\d{2}\s*-?\s*", "", title).strip() or title
            quirks.append(f"Tab '{title}' has no MEET_INFO entry; name guessed.")
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        cols = find_cols(header)
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
            place_raw = get(r, "place")
            place = num(place_raw)
            if place is not None and place < prev_place and cur:
                races.append(cur); cur = []  # place numbering restarted -> second race in same tab
            if place is not None:
                prev_place = place
            if cols["name"] is not None:
                name = str(get(r, "name") or "").strip()
            else:
                name = f"{str(get(r, 'first') or '').strip()} {str(get(r, 'last') or '').strip()}".strip()
            name = re.sub(r"\s+", " ", name)
            school = clean_school(get(r, "school"))
            time_raw = get(r, "time")
            t = parse_time(time_raw)
            status = None
            if isinstance(time_raw, str) and time_raw.strip().upper() in ("DNF", "DNS", "DQ"):
                status = time_raw.strip().upper()
            grade = num(get(r, "grade"))
            pace = parse_pace(get(r, "pace"))
            heights = bool(HEIGHTS_RE.search(school))
            note = None
            if heights and name in NAME_ALIASES:
                note = f"Listed in sheet as \u201c{name}\u201d"
                name = NAME_ALIASES[name]
            entry = {
                "place": int(place) if place is not None else (str(place_raw).strip() if place_raw not in (None, "") else None),
                "grade": int(grade) if grade else None,
                "name": name,
                "school": school,
                "time": fmt_secs(t) if t is not None else status,
                "secs": round(t, 2) if t is not None else None,
                "pace": f"{pace // 60}:{pace % 60:02d}" if pace else None,
                "paceSecs": pace,
                "heights": heights,
            }
            if heights:
                entry["sheetHPlace"] = num(get(r, "hplace"))
            if note:
                entry["note"] = note
            cur.append(entry)
        if cur:
            races.append(cur)

        for ri, field in enumerate(races):
            # Flag times that are faster than someone placed ahead of them (out of order in the sheet).
            slowest_ahead = 0
            for e in field:
                if e["secs"] is None or not isinstance(e["place"], int):
                    continue
                if e["secs"] < slowest_ahead - 1.0:
                    e["suspect"] = True
                    quirks.append(f"Tab '{title}': {e['name']} ({e['school']}) place {e['place']} has time {e['time']}, faster than runners placed ahead \u2013 time/place likely mis-entered.")
                slowest_ahead = max(slowest_ahead, e["secs"])
            hp = 0
            for e in field:
                if e["heights"]:
                    if isinstance(e["place"], int) and e["secs"] is not None:
                        hp += 1
                        e["hplace"] = hp
                    else:
                        e["hplace"] = None
            heights_n = sum(e["heights"] for e in field)
            # implied distance from time/pace (median), only as a hint
            ratios = [e["secs"] / e["paceSecs"] for e in field if e["secs"] and e["paceSecs"]]
            implied = round(statistics.median(ratios), 2) if ratios else None
            label = info["name"] + (f" \u2013 Race {ri + 1}" if len(races) > 1 else "")
            mid = slug(f"{info['season']}-{info['name']}-{info.get('date', '')}" + (f"-race{ri + 1}" if len(races) > 1 else ""))
            meets.append({
                "id": mid,
                "season": info["season"],
                "name": label,
                "date": info.get("date"),
                "tab": title,
                "distance": info.get("distance"),
                "impliedDistance": implied,
                "order": order * 10 + ri,
                "fieldSize": len(field),
                "heightsCount": heights_n,
                "hasPace": any(e["paceSecs"] for e in field),
                "hasGrade": any(e["grade"] for e in field),
                "results": field,
            })
            if len(races) > 1:
                quirks.append(f"Tab '{title}' contains {len(races)} separately-numbered races (placings restart); shown as Race 1..{len(races)}.")

    # Runner index (Heights only)
    runners = {}
    for m in meets:
        for e in m["results"]:
            if not e["heights"]:
                continue
            r = runners.setdefault(e["name"], {"name": e["name"], "id": slug(e["name"]), "results": []})
            r["results"].append({
                "meet": m["id"], "place": e["place"], "hplace": e.get("hplace"),
                "grade": e["grade"], "time": e["time"], "secs": e["secs"],
                "pace": e["pace"], "paceSecs": e["paceSecs"],
                "suspect": e.get("suspect", False),
            })
    for r in runners.values():
        seasons = sorted({next(m["season"] for m in meets if m["id"] == x["meet"]) for x in r["results"]})
        r["seasons"] = seasons
        paces = [x for x in r["results"] if x["paceSecs"] and not x["suspect"]]
        r["bestPace"] = min(paces, key=lambda x: x["paceSecs"])["pace"] if paces else None
        r["bestPaceSecs"] = min((x["paceSecs"] for x in paces), default=None)

    data = {
        "generated": datetime.date.today().isoformat(),
        "source": "Google Sheet \u201cHeights MS XC Results\u201d",
        "meets": sorted(meets, key=lambda m: (-m["season"], -m["order"])),
        "runners": sorted(runners.values(), key=lambda r: r["name"].split()[-1] + r["name"]),
        "quirks": EXTRA_NOTES + sorted(set(quirks)),
    }
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")
    with open(os.path.join(out, "results.json"), "w") as f:
        json.dump(data, f, indent=1)
    with open(os.path.join(out, "results.js"), "w") as f:
        f.write("// Generated by scripts/build_data.py from the Heights MS XC Results sheet. Do not edit by hand.\n")
        f.write("window.XC_DATA = " + json.dumps(data, separators=(",", ":")) + ";\n")
    print(f"{len(meets)} races, {len(runners)} Heights runners, {sum(m['fieldSize'] for m in meets)} result rows")
    for q in data["quirks"]:
        print("quirk:", q)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "Heights_MS_XC_Results.xlsx")
