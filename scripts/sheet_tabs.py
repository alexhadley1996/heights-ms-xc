"""New-style tabs in the Heights XC Results sheet: the tab title carries the date, distance, level and meet.

Old tabs ("2026 - Bullis", anything listed in MEET_INFO) are not handled here. A title only counts as new-style
when it contains a full date (year, month and day). Everything else stays on the original middle-school path.
"""
import datetime, difflib, re

import build_data as B
import hs_data as H
import new_races as N

_MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8,
           "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12}
# Longer names first so "Varsity" is not read as the single letter V, and "Freshman" before "Fr".
_LEVELS = [
    (re.compile(r"\bmiddle\s+school\b", re.I), "MS", "Open", "Middle school"),
    (re.compile(r"\bms\b", re.I), "MS", "Open", "Middle school"),
    (re.compile(r"\bfreshman\b|\bfrosh\b|\bfr\b", re.I), "HS", "Freshman", "Freshman"),
    (re.compile(r"\bjunior\s+varsity\b|\bjv\b", re.I), "HS", "JV", "JV"),
    (re.compile(r"\bvarsity\b|\bvar\b", re.I), "HS", "Varsity", "Varsity"),
    (re.compile(r"\bopen\b", re.I), "HS", "Open", "Open"),
    (re.compile(r"(?<![A-Za-z0-9])V(?![A-Za-z0-9])"), "HS", "Varsity", "Varsity"),
]
# Kept on the race label ("Varsity 2A", "JV Small Schools"), not treated as the meet name.
_EXTRA = re.compile(r"^(?:small\s+schools|[1-6]a|[ab])$", re.I)
_DIST = re.compile(r"\b(\d+(?:\.\d+)?)\s*(k|km|kilometers?|kilometres?|miles?|mi|meters?|metres?|m)\b", re.I)


def _year(y):
    y = int(y)
    if y < 100:
        y += 2000 if y < 70 else 1900
    return y


def _date(y, m, d):
    try:
        return datetime.date(_year(y), int(m), int(d))
    except ValueError:
        return None


def find_date(title):
    """-> (date, start, end) of the first full date in the title, or None."""
    patterns = [
        (re.compile(r"\b((?:19|20)\d{2})-(\d{1,2})-(\d{1,2})\b"), lambda m: _date(m.group(1), m.group(2), m.group(3))),
        (re.compile(r"\b(\d{1,2})[./](\d{1,2})[./](\d{2,4})\b"), lambda m: _date(m.group(3), m.group(1), m.group(2))),
        (re.compile(r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?\s+(\d{1,2}),?\s+((?:19|20)\d{2})\b", re.I),
         lambda m: _date(m.group(3), _MONTHS[m.group(1).lower()[:4] if m.group(1).lower().startswith("sept") else m.group(1).lower()[:3]], m.group(2))),
        (re.compile(r"\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?,?\s+((?:19|20)\d{2})\b", re.I),
         lambda m: _date(m.group(3), _MONTHS[m.group(2).lower()[:4] if m.group(2).lower().startswith("sept") else m.group(2).lower()[:3]], m.group(1))),
    ]
    found = []
    for cre, conv in patterns:
        m = cre.search(title)
        if not m:
            continue
        d = conv(m)
        if d and 1990 <= d.year <= datetime.date.today().year + 2:
            found.append((m.start(), d, m.end()))
    if not found:
        return None
    found.sort(key=lambda x: (x[0], -(x[2] - x[0])))
    _, d, end = found[0]
    start = found[0][0]
    return d, start, end


def find_distance(text):
    """-> ((metres, miles, key), start, end) or None. '5000m' becomes 5K, the same as the HS races."""
    m = _DIST.search(text)
    if not m:
        return None
    raw = m.group(0)
    parsed = N.parse_distance(raw)
    if not parsed:
        parsed = H.parse_distance(raw)
        if not parsed or not parsed[0]:
            return None
    return parsed, m.start(), m.end()


def find_level(text):
    """-> (division, level, label, leftover) or None when the title names no level."""
    for cre, div, level, label in _LEVELS:
        m = cre.search(text)
        if not m:
            continue
        leftover = (text[:m.start()] + " " + text[m.end():]).strip()
        extras = []
        def pull(src):
            out, extra = [], []
            for tok in re.split(r"\s+", src.strip()):
                if not tok:
                    continue
                if _EXTRA.match(tok) or tok.lower() == "small" or tok.lower() == "schools":
                    extra.append(tok)
                else:
                    out.append(tok)
            # "Small" "Schools" stay together; a lone "Small" or "Schools" goes back to the meet name
            if extra == ["Small"] or extra == ["Schools"] or [t.lower() for t in extra] in (["small"], ["schools"]):
                out.extend(extra)
                extra = []
            return " ".join(out), extra
        leftover, extras = pull(leftover)
        # also pull a modifier that sat immediately before the level word, already in leftover
        if extras:
            label = (label + " " + " ".join(extras)).strip()
            if label.lower().startswith("varsity ") or label.lower().startswith("jv ") or label.lower().startswith("freshman "):
                pass
        return div, level, label, leftover
    return None


def _cleanup(text):
    text = re.sub(r"[-–—|/]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip(" -–,")
    return text


def parse_title(title):
    """-> dict with date/distance/level/meet, or None if the title has no full date (an old-style tab)."""
    found = find_date(title)
    if not found:
        return None
    date, a, b = found
    rest = _cleanup(title[:a] + " " + title[b:])
    dist = find_distance(rest)
    if dist:
        parsed, a, b = dist
        rest = _cleanup(rest[:a] + " " + rest[b:])
    else:
        parsed = None
    level = find_level(rest)
    if level:
        division, lvl, label, rest = level
        rest = _cleanup(rest)
        defaulted = False
    else:
        division, lvl, label, defaulted = "MS", "Open", "Middle school", True
    meet = _cleanup(rest)
    return {"date": date, "distance": parsed, "division": division, "level": lvl, "levelLabel": label,
            "race": label if division == "HS" else "", "meet": meet, "defaultedMS": defaulted}


def _cell(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return f"{v.hour}:{v.minute:02d}:{v.second:02d}"
    if isinstance(v, datetime.time):
        # 17:42:00 is how Sheets stores a race time of 17:42. clean_time turns that back into 17:42.
        if v.microsecond:
            return f"{v.hour}:{v.minute:02d}:{v.second + v.microsecond / 1e6:.1f}"
        return f"{v.hour}:{v.minute:02d}:{v.second:02d}"
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _header_map(cells):
    """-> list of column keys (or None) using the same names as a race-results CSV, or None if this isn't a header."""
    lows = [_cell(c).lower().rstrip(":") for c in cells]
    if not any(lows):
        return None
    known = set(N.COLUMN_NAMES)
    if not ({"place", "pl", "overall place"} & set(lows) and {"name", "athlete", "runner", "first name"} & set(lows)):
        return None
    header = []
    unknown = []
    for c in lows:
        if not c:
            header.append(None)
            continue
        if c not in known and c not in ("first name", "last name"):
            unknown.append(c)
            header.append(None)
        elif c in ("first name", "last name"):
            header.append(c)
        else:
            header.append(N.COLUMN_NAMES[c])
    return header, unknown


def read_tab(ws):
    """-> (field entries or None, problems). One bad row is a problem; the caller drops the whole tab."""
    rows = list(ws.iter_rows(values_only=True))
    header = None
    start = None
    unknown = []
    for i, r in enumerate(rows):
        hit = _header_map(r)
        if hit:
            header, unknown = hit
            start = i + 1
            break
    problems = []
    if header is None:
        return None, ["no results header row. The first row should be Place, Name, Grade, School, Time "
                      "(or the usual sheet headings: Place, Grade, Name, Pace, Time, School)."]
    for u in unknown:
        if u not in ("meet",):  # a Meet column pasted from another sheet is ignored
            problems.append(f"unknown column name '{u}'. Use Place, Name, Grade, School, Time.")
    # old sheet layout (First Name / Last Name, or Place (Heights)) goes through the existing reader
    lows = [_cell(c).lower() for c in rows[start - 1]]
    return header, start, rows, problems, ("first name" in lows or "last name" in lows or any(x.startswith("place (heights") for x in lows))


def field_from_tab(ws, division):
    """-> (list of result entries, problems)."""
    parsed = read_tab(ws)
    if parsed[0] is None:
        return None, parsed[1]
    header, start, rows, problems, old_layout = parsed
    title_row = start  # 1-based header row number is start
    field, seen = [], {}
    lo, hi = (4, 8) if division == "MS" else (9, 12)
    if old_layout:
        cols = B.find_cols(rows[start - 1])
        if cols["time"] is None:
            return None, problems + ["no Time column found."]
        def get(r, k):
            i = cols[k]
            return r[i] if i is not None and i < len(r) else None
        for n, r in enumerate(rows[start:], start=start + 1):
            if not any(c not in (None, "", " ") for c in r[:12]):
                continue
            if cols["name"] is not None:
                name = get(r, "name")
            else:
                name = f"{_cell(get(r, 'first'))} {_cell(get(r, 'last'))}".strip()
            name = re.sub(r"\s+", " ", _cell(name) if not isinstance(name, str) else name).strip()
            if not name:
                problems.append(f"row {n}: the Name is empty.")
                continue
            t = get(r, "time")
            if t in (None, ""):
                problems.append(f"row {n} ({name}): the Time is empty. Use minutes:seconds like 17:42.3, or DNF.")
                continue
            school = B.clean_school(get(r, "school")) or "The Heights"
            e = B.make_entry(get(r, "place"), get(r, "grade"), H.proper_case(name) if division == "HS" else name,
                             school, t, get(r, "pace"))
            if division == "HS":
                e["heights"] = bool(H.HS_HEIGHTS_RE.match(str(e["school"]).strip()))
            if e["heights"]:
                k = B.norm_name(e["name"])
                if k in seen:
                    problems.append(f"row {n}: {e['name']} is listed twice (also row {seen[k]}).")
                    continue
                seen[k] = n
            field.append(e)
        return field, problems

    for n, r in enumerate(rows[start:], start=start + 1):
        cells = [_cell(c) for c in r]
        if not any(cells):
            continue
        row = {}
        for key, c in zip(header, cells):
            if key and key not in ("first name", "last name"):
                row[key] = c
        name = re.sub(r"\s+", " ", row.get("name", "")).strip()
        if "," in name and len(name.split(",")) == 2:
            last, first = [x.strip() for x in name.split(",")]
            name = f"{first} {last}"
        if not name and not row.get("time") and not row.get("place"):
            continue
        if not name:
            problems.append(f"row {n}: the Name is empty.")
            continue
        if not row.get("time"):
            problems.append(f"row {n} ({name}): the Time is empty. Use minutes:seconds like 17:42.3, or DNF.")
            continue
        t = N.clean_time(row["time"])
        if t is None:
            problems.append(f"row {n} ({name}): can't read the time '{row['time']}'. Use minutes:seconds like 17:42.3, or DNF.")
            continue
        place = row.get("place", "").strip().rstrip(".")
        if place.upper() in ("DNF", "DNS", "DQ", "-", "--") or t in ("DNF", "DNS", "DQ"):
            place = ""
        elif place and not re.fullmatch(r"\d+", place):
            problems.append(f"row {n} ({name}): the Place '{place}' should be a whole number.")
            continue
        g = row.get("grade", "").strip().rstrip(".")
        gnum = N.CLASS_YEARS.get(g.lower()) if g else None
        if g and gnum is None:
            if not re.fullmatch(r"\d{1,2}", g):
                problems.append(f"row {n} ({name}): the Grade '{g}' should be a number like {lo}–{hi} (or empty).")
                continue
            gnum = int(g)
        if gnum is not None and not (lo <= gnum <= hi):
            problems.append(f"row {n} ({name}): grade {gnum} isn't a {'middle' if division == 'MS' else 'high'} school grade ({lo}–{hi}).")
            continue
        school = re.sub(r"\s+", " ", row.get("school", "")).strip() or "The Heights"
        packed = {"Place": place, "Grade": str(gnum) if gnum else "", "Name": name, "School": school, "Time": t}
        e = H.entry(packed, {}) if division == "HS" else B.make_entry(place or None, gnum, name, school, t, None)
        if e["heights"]:
            k = B.norm_name(e["name"])
            if k in seen:
                problems.append(f"row {n}: {e['name']} is listed twice (also row {seen[k]}).")
                continue
            seen[k] = n
        field.append(e)
    if not field and not problems:
        problems.append("there are no runners under the header row.")
    if field and not any(e["heights"] for e in field):
        problems.append("no runner is from The Heights. Write 'The Heights' in the School column, or leave School empty.")
    return field, problems


def norm_meet(name):
    s = name.lower().replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _similar(a, b):
    if not a or not b:
        return 0
    if a == b:
        return 1
    short, long = (a, b) if len(a) <= len(b) else (b, a)
    if len(short) >= 8 and (short in long or long.startswith(short)):
        return 0.9
    return difflib.SequenceMatcher(None, a, b).ratio()


def _existing_hs(meets, parsed):
    """The HS meet this tab belongs with, and the one race it replaces (or None)."""
    year = parsed["date"].year
    by_name = {}
    for m in meets:
        # only races that were already on the site, not other new tabs from this same sheet
        if m.get("sheetNew") or m.get("division") != "HS" or m.get("season") != year or not m.get("meetName"):
            continue
        by_name.setdefault(m["meetName"], []).append(m)
    target = norm_meet(parsed["meet"])
    best, score = None, 0
    for name in by_name:
        s = _similar(target, norm_meet(name))
        if s > score:
            best, score = name, s
    if best is None or score < 0.84:
        return None, None
    pool = [m for m in by_name[best] if m.get("level") == parsed["level"]]
    extras = [w for w in parsed["levelLabel"].split()[1:]]
    if extras and pool:
        narrowed = [m for m in pool if all(e.lower() in f"{m.get('levelLabel','')} {m.get('race','')}".lower() for e in extras)]
        if len(narrowed) == 1:
            return best, narrowed[0]
        if len(narrowed) > 1:
            return best, None
    if len(pool) == 1:
        return best, pool[0]
    return best, None


def _meta(parsed, field, order):
    d = parsed["date"]
    dist = parsed["distance"]
    metres, miles, dkey = dist if dist else (None, None, None)
    only_heights = bool(field) and all(e["heights"] for e in field)
    common = {
        "season": d.year, "division": parsed["division"], "date": d.isoformat(), "dateLabel": None,
        "tab": parsed["title"], "origin": "sheet", "distance": miles,
        "distanceSource": "official" if miles else None, "statedDistance": miles, "impliedDistance": None,
        "officialDistance": None, "location": None, "source": None, "correction": None,
        "meetName": parsed["meet"], "sheetNew": True, "_tabOrder": order, "_scoreNotes": [],
        "noTeamScores": "Only Heights runners are listed for this race, so team scores aren't shown." if only_heights else None,
    }
    if parsed["division"] == "HS":
        label = parsed["levelLabel"]
        common.update({
            "level": parsed["level"], "levelLabel": label, "race": label, "origin": "hs",
            "name": f"{parsed['meet']} \u2013 {label}", "distanceM": metres, "distKey": dkey,
            "officialName": None, "note": None, "notes": [], "completeness": None, "teamScoreStatus": None,
            "doubled": False, "aliasIds": None, "scoreOnly": None, "partialScores": None,
            "id": B.slug(f"hs-{d.year}-{parsed['meet']}-{label}"),
        })
    else:
        common.update({
            "level": "Open", "levelLabel": "Middle school",
            "name": parsed["meet"], "id": B.slug(f"{d.year}-{parsed['meet']}-{d.isoformat()}"),
        })
    return common


def official_rows(meet):
    off = (meet.get("scoring") or {}).get("official")
    if not off:
        return None
    return [{"Place": t["place"], "Score": t["score"], "Team": t["team"], "SourceURL": off.get("source")} for t in off["teams"]]


def apply_titled_tabs(tabs, meets, quirks):
    """tabs: list of (order, worksheet, parsed title). Mutates meets. -> list of problem strings."""
    problems = []
    for order, ws, parsed in tabs:
        parsed = dict(parsed, title=ws.title)
        if parsed["defaultedMS"]:
            print(f"note: tab {ws.title!r} has a date but no level (Varsity, JV, Freshman, Open, or MS); treated as middle school.")
        if not parsed["meet"]:
            problems.append(f"tab {ws.title!r} was NOT added:\n    - the tab name has a date, but no meet name left after the date, distance and level.")
            continue
        try:
            field, probs = field_from_tab(ws, parsed["division"])
        except Exception as ex:
            problems.append(f"tab {ws.title!r} was NOT added:\n    - couldn't read the tab ({type(ex).__name__}: {ex}).")
            continue
        if probs or not field:
            if not probs:
                probs = ["no runners could be read."]
            problems.append(f"tab {ws.title!r} was NOT added:\n    - " + "\n    - ".join(probs))
            continue
        field.sort(key=lambda e: (0, e["place"]) if isinstance(e["place"], int) else (1, e["secs"] or 1e9))
        canon, replace = _existing_hs(meets, parsed) if parsed["division"] == "HS" else (None, None)
        if replace is not None:
            meta = _meta(parsed, field, order)
            meta["id"] = replace["id"]
            meta["meetName"] = replace.get("meetName") or parsed["meet"]
            meta["sheetNew"] = False
            meta["joinedExisting"] = True
            meta["location"] = replace.get("location")
            meta["source"] = replace.get("source")
            meta["notes"] = replace.get("notes") or []
            meta["note"] = replace.get("note")
            meta["aliasIds"] = replace.get("aliasIds")
            meta["_official"] = official_rows(replace)
            built = B.finalize(field, meta, quirks)
            meets[meets.index(replace)] = built
            print(f"tab {ws.title} replaced existing race {replace.get('name')}")
            continue
        meta = _meta(parsed, field, order)
        if canon:
            meta["meetName"] = canon
            meta["joinedExisting"] = True
            if parsed["division"] == "HS":
                meta["name"] = f"{canon} \u2013 {parsed['levelLabel']}"
                meta["id"] = B.slug(f"hs-{parsed['date'].year}-{canon}-{parsed['levelLabel']}")
        meets.append(B.finalize(field, meta, quirks))
    return problems
