"""Races added by hand: one CSV file per race in the race-results/ folder (see race-results/README.md and
HOW-TO-UPDATE.md). Each file starts with a few "label,value" rows about the race, then a header row that starts
with "Place", then one row per runner. Files whose name starts with "_" (the template and the example) are ignored.

load_uploaded_races() returns (meets, problems): meet dicts in the same shape as the other races, plus a list of
plain-English problems. A file with any problem is left out completely (so a half-typed race never goes live).
"""
import csv, datetime, glob, os, re

import build_data as B
import hs_data as H

FOLDER = os.path.join(B.ROOT, "race-results")
REL = "race-results"

INFO_LABELS = {
    "meet": "meet", "meet name": "meet", "date": "date", "venue": "venue", "location": "venue", "course": "venue",
    "level": "level", "division": "level", "race": "race", "race name": "race", "distance": "distance",
    "source": "source", "source link": "source", "link": "source", "results link": "source", "notes": "notes", "note": "notes",
}
COLUMN_NAMES = {
    "place": "place", "pl": "place", "overall place": "place", "name": "name", "athlete": "name", "runner": "name",
    "grade": "grade", "gr": "grade", "yr": "grade", "year": "grade", "school": "school", "team": "school",
    "time": "time", "finish time": "time", "mark": "time",
    # columns that results sites often include; ignored
    "pace": None, "points": None, "pts": None, "score": None, "team score": None, "bib": None, "#": None, "photo": None,
    "avg": None, "avg mile": None, "mile": None, "split": None, "gap": None, "behind": None, "heights place": None,
    "place (heights)": None, "heights": None, "meet": None, "level": None, "sourceurl": None, "pr": None,
}
CLASS_YEARS = {"fr": 9, "so": 10, "jr": 11, "sr": 12, "freshman": 9, "sophomore": 10, "junior": 11, "senior": 12}
DATE_FORMATS = ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%b %d, %Y", "%B %d, %Y", "%b %d %Y", "%B %d %Y", "%d %b %Y",
                "%d %B %Y", "%a, %b %d, %Y", "%A, %B %d, %Y", "%d-%b-%Y", "%d-%b-%y")
GENERIC_MS_RACES = {"", "middle school", "middle school boys", "boys", "ms", "ms boys", "varsity"}


def _read_rows(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp1252"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    lines = text.splitlines()
    # text pasted straight from Google Sheets into GitHub's "Create new file" is tab-separated
    delim = "\t" if any("\t" in l for l in lines) else ","
    return list(csv.reader(lines, delimiter=delim))


def parse_date(s):
    s = re.sub(r"\s+", " ", s.strip().replace("Sept", "Sep").replace(".", ""))
    for f in DATE_FORMATS:
        try:
            return datetime.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def parse_distance(s):
    """'5K', '5 km', '3 miles', '2 mi', '3000m', '2.1' (miles) -> (metres, miles, key) or None if unreadable."""
    t = s.strip().lower().replace(",", "")
    if re.fullmatch(r"\d+(\.\d+)?", t):
        t += " mi"
    t = re.sub(r"(\d)\s*(miles?|mi)\b\.?", r"\1 mi", t)
    t = re.sub(r"(\d)\s*(kilometers?|kilometres?|km|k)\b", r"\1 km", t)
    t = re.sub(r"(\d)\s*(meters?|metres?|m)\b", r"\1 m", t)
    m = re.fullmatch(r"(\d+(?:\.\d+)?) (mi|km|m)", t)
    if not m:
        return None
    v, unit = float(m.group(1)), m.group(2)
    if unit == "mi":
        return (v * 1609.344, v, f"{v:g} mi") if 0.5 <= v <= 10 else None
    if unit == "km":
        return (v * 1000, v * 1000 / 1609.344, f"{v:g}K") if 1 <= v <= 15 else None
    if 800 <= v <= 15000:
        r = H.parse_distance(f"{int(v)} m") if v == int(v) and v >= 1000 else (None, None, None)
        return r if r[0] else (v, v / 1609.344, f"{v:g} m")
    return None


def clean_time(s):
    """Return a time the build understands ('17:42.3', 'DNF') or None. Undoes Google Sheets turning 17:42 into
    17:42:00 (a clock time) or 0:17:42."""
    t = re.sub(r"(\s*(PR|SR|SB|\*|\u2020))+$", "", s.strip(), flags=re.I).strip()
    if t.upper() in ("DNF", "DNS", "DQ"):
        return t.upper()
    m = re.fullmatch(r"(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP])\.?M\.?", t, re.I)  # "5:42:00 PM" = 17:42
    if m:
        h = int(m.group(1)) % 12 + (12 if m.group(4).upper() == "P" else 0) or 12
        if m.group(3) not in (None, "00"):
            return None
        t = f"{h}:{m.group(2)}"
    m = re.fullmatch(r"(\d{1,2}):(\d{2}):(\d{2}(?:\.\d+)?)", t)
    if m:
        h, mi, se = int(m.group(1)), int(m.group(2)), m.group(3)
        if h == 0:
            t = f"{mi}:{se}"
        elif float(se) == 0 and h >= 5:  # "17:42:00" = 17 min 42 s mis-read as a clock time
            t = f"{h}:{mi:02d}"
        else:
            return None
    m = re.fullmatch(r"(\d{1,2}):(\d{2})(\.\d+)?", t)
    if not m or int(m.group(2)) >= 60:
        return None
    total = int(m.group(1)) * 60 + int(m.group(2))
    return t if 4 * 60 <= total <= 59 * 60 else None


def parse_file(path):
    """-> (info dict, list of result row dicts with 'row' numbers, problems)"""
    name = os.path.basename(path)
    rows = _read_rows(path)
    info, problems, results, header = {}, [], [], None
    for n, r in enumerate(rows, start=1):
        cells = [c.strip() for c in r]
        if not any(cells):
            continue
        if header is None:
            first = cells[0].lower().rstrip(":")
            low = {c.lower().rstrip(":") for c in cells}
            if first in ("place", "pl", "overall place") or (low & {"place", "pl", "overall place"} and low & {"name", "athlete", "runner"}):
                header = []
                for c in cells:
                    key = c.lower().rstrip(":")
                    if key and key not in COLUMN_NAMES:
                        problems.append(f"row {n}: unknown column name '{c}'. Use: Place, Name, Grade, School, Time.")
                    header.append(COLUMN_NAMES.get(key))
                for need in ("place", "name", "time"):
                    if need not in header:
                        problems.append(f"row {n}: the results header row needs a '{need.title()}' column.")
                continue
            label = INFO_LABELS.get(first)
            if not label:
                problems.append(f"row {n}: '{cells[0]}' isn't a label this site knows. The top rows must be "
                                "Meet, Date, Venue, Level, Race, Distance, Source, Notes, then the row 'Place,Name,Grade,School,Time'.")
                continue
            info[label] = next((c for c in cells[1:] if c), "")
            continue
        row = {"row": n}
        for key, c in zip(header, cells):
            if key:
                row[key] = c
        if not any(row.get(k) for k in ("place", "name", "time", "grade")):
            continue
        results.append(row)
    if header is None:
        problems.append("there is no results header row. Add a row that starts with 'Place' (Place,Name,Grade,School,Time) "
                        "above the runners.")
    return info, results, problems


def build_meet(path, quirks):
    """-> (meet dict or None, problems)"""
    info, rows, problems = parse_file(path)
    meet = re.sub(r"\s+", " ", info.get("meet", "")).strip()
    meet = re.sub(r"^(19|20)\d{2}\s+", "", meet)  # "2027 Bullis Invitational" -> "Bullis Invitational"
    if not meet:
        problems.append("the 'Meet' row is missing or empty (the meet's name, e.g. Bullis Invitational).")
    d = parse_date(info.get("date", "")) if info.get("date") else None
    if not info.get("date"):
        problems.append("the 'Date' row is missing or empty. Write it like 2027-09-18 or Sep 18, 2027.")
    elif not d:
        problems.append(f"can't read the date '{info['date']}'. Write it like 2027-09-18 or Sep 18, 2027.")
    elif not (1950 <= d.year <= datetime.date.today().year + 1):
        problems.append(f"the date '{info['date']}' has an unlikely year ({d.year}).")
    lvl = info.get("level", "").strip().upper().replace(".", "")
    lvl = {"MIDDLE SCHOOL": "MS", "MIDDLE": "MS", "HIGH SCHOOL": "HS", "HIGH": "HS"}.get(lvl, lvl)
    if lvl not in ("MS", "HS"):
        problems.append(f"the 'Level' row must be MS (middle school) or HS (high school), not '{info.get('level', '')}'.")
    race = re.sub(r"\s+", " ", info.get("race", "")).strip()
    if lvl == "HS" and not race:
        problems.append("the 'Race' row is empty. For high school write Varsity, JV, Freshman or Open (e.g. 'Varsity' or 'JV 2A').")
    dist = None
    if info.get("distance"):
        dist = parse_distance(info["distance"])
        if not dist:
            problems.append(f"can't read the distance '{info['distance']}'. Write it like 5K, 3 miles, 2 mi or 3000m.")
    source = info.get("source", "").strip()
    if source and not re.match(r"^https?://", source):
        problems.append(f"the 'Source' row must be a web link starting with http:// or https:// (or left empty), not '{source}'.")

    lo, hi = (4, 8) if lvl == "MS" else (9, 12)
    field, seen = [], {}
    for r in rows:
        n = r["row"]
        name = re.sub(r"\s+", " ", r.get("name", "")).strip()
        if not name:
            problems.append(f"row {n}: the Name is empty.")
            continue
        if "," in name and len(name.split(",")) == 2:  # "Smith, John" -> "John Smith"
            last, first = [x.strip() for x in name.split(",")]
            name = f"{first} {last}"
        place = r.get("place", "").strip().rstrip(".")
        t = clean_time(r.get("time", ""))
        if not r.get("time", "").strip():
            problems.append(f"row {n} ({name}): the Time is empty. Use minutes:seconds like 17:42.3, or DNF.")
            continue
        if t is None:
            problems.append(f"row {n} ({name}): can't read the time '{r['time']}'. Use minutes:seconds like 17:42.3 "
                            "(or DNF / DNS / DQ). If Google Sheets changed it, format the Time column as Plain text.")
            continue
        if place and not re.fullmatch(r"\d+", place) and place.upper() not in ("DNF", "DNS", "DQ", "-", "--"):
            problems.append(f"row {n} ({name}): the Place '{place}' should be a whole number (or empty for DNF).")
            continue
        if place.upper() in ("DNF", "DNS", "DQ", "-", "--"):
            place = ""
        if t in ("DNF", "DNS", "DQ"):
            place = ""
        g = r.get("grade", "").strip().rstrip(".")
        gnum = CLASS_YEARS.get(g.lower(), None) if g else None
        if g and gnum is None:
            if not re.fullmatch(r"\d{1,2}", g):
                problems.append(f"row {n} ({name}): the Grade '{g}' should be a number like {lo}\u2013{hi} (or empty).")
                continue
            gnum = int(g)
        if gnum is not None and lvl in ("MS", "HS") and not (lo <= gnum <= hi):
            problems.append(f"row {n} ({name}): grade {gnum} isn't a {'middle' if lvl == 'MS' else 'high'} school grade "
                            f"({lo}\u2013{hi}). Check the Level row and that the columns are in the right order.")
            continue
        school = re.sub(r"\s+", " ", r.get("school", "")).strip() or "The Heights"
        row = {"Place": place, "Grade": str(gnum) if gnum else "", "Name": name, "School": school, "Time": t}
        if lvl == "HS":
            e = H.entry(row, {})
        else:
            e = B.make_entry(place or None, gnum, H.proper_case(name), school, t, None)
        e["_row"] = n
        if e["heights"]:
            k = B.norm_name(e["name"])
            if k in seen:
                problems.append(f"row {n}: {e['name']} is listed twice in this race (rows {seen[k]} and {n}). "
                                "Remove one, or if two different runners share the name, add a middle initial.")
                continue
            seen[k] = n
        field.append(e)
    if rows and not field and not problems:
        problems.append("no runners could be read.")
    if not rows:
        problems.append("there are no runners under the 'Place,Name,Grade,School,Time' row.")
    if field and not any(e["heights"] for e in field):
        problems.append("no runner is from The Heights. Put 'The Heights' in the School column for our runners "
                        "(or leave School empty for them).")
    if problems:
        return None, problems

    for e in field:
        e.pop("_row", None)
    field.sort(key=lambda e: (0, e["place"]) if isinstance(e["place"], int) else (1, e["secs"] or 1e9))
    season = d.year
    metres, miles, dkey = dist if dist else (None, None, None)
    notes = [x for x in [info.get("notes", "").strip()] if x]
    if not dist:
        notes.append("The results don't list the race distance, so no paces are shown.")
    only_heights = all(e["heights"] for e in field)
    base = os.path.basename(path)
    src = {"name": re.sub(r"^www\.", "", re.sub(r"^https?://([^/]+).*", r"\1", source)), "url": source, "original": None} if source else None
    common = {
        "season": season, "division": lvl, "date": d.isoformat(), "dateLabel": None, "tab": None,
        "distance": miles, "distanceSource": "official" if miles else None,
        "statedDistance": None, "impliedDistance": None, "officialDistance": None,
        "location": info.get("venue", "").strip() or None, "source": src, "correction": None,
        "noTeamScores": "Only Heights runners are listed for this race, so team scores aren't shown." if only_heights else None,
        "_official": None, "_tabOrder": 0, "csv": f"{REL}/{base}", "uploaded": True,
    }
    if lvl == "HS":
        level = ("JV" if re.search(r"\bjv\b|junior varsity|varsity b\b", race, re.I) else
                 "Freshman" if re.search(r"fresh|frosh|9th", race, re.I) else
                 "Open" if re.search(r"\bopen\b", race, re.I) else
                 "Varsity" if re.search(r"varsity", race, re.I) else "Open")
        meta = dict(common, **{
            "id": B.slug(f"hs-{season}-{meet}-{race}"), "level": level, "levelLabel": race, "race": race,
            "meetName": meet, "name": f"{meet} \u2013 {race}", "origin": "hs", "distanceM": metres, "distKey": dkey,
            "officialName": None, "note": " ".join(notes) or None, "notes": notes, "completeness": None,
            "teamScoreStatus": None, "doubled": False, "aliasIds": None, "scoreOnly": None, "partialScores": None,
        })
    else:
        label = meet + (f" ({race})" if race.lower() not in GENERIC_MS_RACES else "")
        meta = dict(common, **{
            "id": B.slug(f"{season}-{meet}-{race}-{d.isoformat()}"), "level": "Open", "levelLabel": "Middle school",
            "name": label, "meetName": meet, "origin": "online", "officialName": None,
            "note": " ".join(notes) or None,
        })
    return B.finalize(field, meta, quirks), []


def load_uploaded_races(quirks, existing_ids):
    meets, problems = [], []
    files = sorted(glob.glob(os.path.join(FOLDER, "*.csv")) + glob.glob(os.path.join(FOLDER, "*.CSV")))
    for path in files:
        base = os.path.basename(path)
        if base.startswith("_"):
            continue
        try:
            m, probs = build_meet(path, quirks)
        except Exception as ex:  # never let one odd file stop the whole site from building
            m, probs = None, [f"couldn't read the file ({type(ex).__name__}: {ex})."]
        if m and (m["id"] in existing_ids):
            probs = [f"this race already seems to be on the site (same meet name, year and race: '{m['id']}'). "
                     "If it is a different race, change the Race row (e.g. 'Varsity B'); otherwise delete this file."]
            m = None
        if probs:
            problems.append(f"{REL}/{base} was NOT added:\n    - " + "\n    - ".join(probs))
            continue
        existing_ids.add(m["id"])
        meets.append(m)
    return meets, problems
