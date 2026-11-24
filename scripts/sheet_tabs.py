"""New-style tabs in the Heights XC Results sheet: the tab title carries the date, distance, level and meet.

Old tabs ("2026 - Bullis", anything listed in MEET_INFO) are not handled here. A title only counts as new-style
when it contains a full date (year, month and day). Everything else stays on the original middle-school path.
"""
import datetime, difflib, os, re

import build_data as B
import hs_data as H
import new_races as N

_MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8,
           "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12}
# Longer names first so "Varsity" is not read as the single letter V, and "Freshman" before "Fr".
_LEVELS = [
    (re.compile(r"\bmiddle\s+school\b", re.I), "MS", "Open", "Middle school"),
    (re.compile(r"\bms\b", re.I), "MS", "Open", "Middle school"),
    (re.compile(r"\bfresh(?:man|men)\b|\bfrosh\b|\bfr\b", re.I), "HS", "Freshman", "Freshman"),
    (re.compile(r"\bjunior\s+varsity\b|\bjv\b", re.I), "HS", "JV", "JV"),
    (re.compile(r"\bvarsity\b|\bvar\b", re.I), "HS", "Varsity", "Varsity"),
    (re.compile(r"\brising\s+stars?\b", re.I), "HS", "Open", "Rising Stars"),
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


# Cross country meets are in the fall. A compact date that can be read two ways prefers these months.
_XC_MONTHS = {8, 9, 10, 11}


def _compact_dates(digits):
    """Every valid M/D/YY or MM/DD/YY reading of a 4–6 digit run with no separators.

    The last two digits are the year. 11426 is 11/4/26 and 1/14/26. 9426 is 9/4/26.
    """
    if not digits.isdigit() or not (4 <= len(digits) <= 6):
        return []
    body, yy = digits[:-2], digits[-2:]
    found = []
    for mlen in (1, 2):
        if len(body) <= mlen:
            continue
        month, day = int(body[:mlen]), int(body[mlen:])
        if not (1 <= month <= 12 and 1 <= day <= 31):
            continue
        d = _date(yy, month, day)
        if d and d not in found:
            found.append(d)
    return found


def _pick_compact(cands):
    """Prefer a month in Aug–Nov. Otherwise the first valid reading (1-digit month first)."""
    season = [d for d in cands if d.month in _XC_MONTHS]
    return (season or cands)[0]


def find_date(title):
    """-> (date, start, end, note) or None. note is set when a compact number was read as a date."""
    patterns = [
        (re.compile(r"\b((?:19|20)\d{2})-(\d{1,2})-(\d{1,2})\b"), lambda m: _date(m.group(1), m.group(2), m.group(3))),
        (re.compile(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\b"), lambda m: _date(m.group(3), m.group(1), m.group(2))),
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
            found.append((m.start(), -(m.end() - m.start()), d, m.end(), None))
    for m in re.finditer(r"\b(\d{4,6})\b", title):
        cands = _compact_dates(m.group(1))
        if not cands:
            continue
        chosen = _pick_compact(cands)
        alts = [d.isoformat() for d in cands if d != chosen]
        note = f"date {m.group(1)} read as {chosen.isoformat()}" + (f" (also possible: {', '.join(alts)})" if alts else "")
        found.append((m.start(), -(m.end() - m.start()), chosen, m.end(), note))
    if not found:
        return None
    found.sort()
    _, _, d, end, note = found[0]
    return d, found[0][0], end, note


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
    date, a, b, date_note = found
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
            "race": label if division == "HS" else "", "meet": meet, "defaultedMS": defaulted, "dateNote": date_note}


def meet_column(ws):
    """The most common value in the tab's Meet column ('2026 NXR Southeast Varsity'), or ''."""
    import collections
    rows = ws.iter_rows(values_only=True)
    head = [_cell(c).lower() for c in next(rows, ())]
    if "meet" not in head:
        return ""
    i = head.index("meet")
    seen = collections.Counter(_cell(r[i]) for r in rows if i < len(r) and _cell(r[i]))
    return seen.most_common(1)[0][0] if seen else ""


def hs_level_from_meet_column(ws):
    """-> (level, label, meet name) when the Meet column names a high school race (Varsity, JV, Freshmen,
    Rising Stars, Open), else None.

    Google's Excel download cuts tab names to 31 characters, so '2026 - NXR Southeast Regionals Varsity'
    arrives as '2026 - NXR Southeast Regionals '. The Meet column inside the tab still says 'Varsity'.
    """
    text = meet_column(ws)
    if not text:
        return None
    found = find_level(_cleanup(text))
    if not found or found[0] != "HS":
        return None
    _, lvl, label, rest = found
    meet = _cleanup(re.sub(r"\b(?:19|20)\d{2}\b", " ", rest))
    return lvl, label, meet


def _distance_from_miles(miles):
    if not miles:
        return None
    if abs(miles * 1609.344 - 5000) < 20:  # a pace formula that divides by 3.107 or 3.1 is a 5K
        return N.parse_distance("5k")
    return N.parse_distance(f"{miles:g} mi")


def parse_undated_hs(title, ws, formula_miles=None):
    """A tab with no full date in its name that is still a high school race.

    The level comes from the tab name ('2026 - Bullis Varsity') or, when the name doesn't say, from the
    Meet column. A name that says MS or middle school always stays middle school. -> parsed dict or None.
    """
    ym = re.search(r"\b((?:19|20)\d{2})\b", title)
    if not ym:
        return None
    season = int(ym.group(1))
    rest = _cleanup(title[:ym.start()] + " " + title[ym.end():])
    dist = find_distance(rest)
    if dist:
        parsed_dist, a, b = dist
        rest = _cleanup(rest[:a] + " " + rest[b:])
    else:
        parsed_dist = _distance_from_miles(formula_miles)
    level = find_level(rest)
    if level and level[0] != "HS":
        return None
    if level:
        _, lvl, label, meet = level
        meet = _cleanup(meet)
        source = "tab name"
    else:
        hit = hs_level_from_meet_column(ws)
        if not hit:
            return None
        lvl, label, meet = hit
        meet = meet or rest
        source = "Meet column"
    if not meet:
        return None
    return {"date": None, "season": season, "distance": parsed_dist, "division": "HS", "level": lvl,
            "levelLabel": label, "race": label, "meet": meet, "defaultedMS": False, "dateNote": None,
            "levelNote": f"read as high school {label} from the {source}"}


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


# Header words a results website might use. Extra columns (pace, bib, points) are ignored.
_HEADER = {
    "place": "place", "pl": "place", "pos": "place", "position": "place", "overall": "place", "overall place": "place",
    "name": "name", "runner": "name", "athlete": "name",
    "first": "first", "first name": "first", "last": "last", "last name": "last",
    "grade": "grade", "gr": "grade", "yr": "grade", "year": "grade", "class": "grade",
    "school": "school", "team": "school", "affiliation": "school",
    "time": "time", "finish": "time", "finish time": "time", "chip": "time", "chip time": "time",
    "net": "time", "net time": "time", "gun": "time", "gun time": "time", "mark": "time",
    "pace": None, "bib": None, "bib #": None, "#": None, "points": None, "pts": None, "score": None,
    "team score": None, "photo": None, "avg": None, "gap": None, "behind": None, "heights place": None,
    "place (heights)": "hplace", "meet": None,
}
_TIME_TOKEN = re.compile(r"\b(\d{1,2}:\d{2}(?:\.\d+)?)\b")
_JUNK_LINE = re.compile(
    r"^(results|scoring|team\s+scores?|page\s+\d+(\s+of\s+\d+)?|boys|girls|men|women|varsity|jv|freshman|freshmen|open)\s*$",
    re.I)
_GRADE_WORD = {"sr": 12, "jr": 11, "so": 10, "fr": 9, "senior": 12, "junior": 11, "sophomore": 10, "freshman": 9,
               "frosh": 9}
_schools = None


def _known_schools():
    """School names already on the site, longest first, so a pasted line can be split into name and school."""
    global _schools
    if _schools is not None:
        return _schools
    found = {"the heights", "heights"}
    import csv, glob
    for path in glob.glob(os.path.join(B.ROOT, "data", "hs", "races", "*.csv")) + glob.glob(os.path.join(B.ROOT, "data", "supplemental", "races", "*.csv")):
        try:
            rows = list(csv.reader(open(path, encoding="utf-8", errors="replace")))
        except OSError:
            continue
        if not rows:
            continue
        head = [c.strip().lower() for c in rows[0]]
        idx = next((i for i, c in enumerate(head) if c in ("school", "team", "affiliation")), None)
        if idx is None:
            continue
        for r in rows[1:]:
            if idx < len(r) and r[idx].strip():
                found.add(re.sub(r"\s+", " ", r[idx].strip().lower()))
    _schools = sorted(found, key=len, reverse=True)
    return _schools


def _header_map(cells):
    """-> (keys, unknown) or None if this row is not a results header."""
    lows = [_cell(c).lower().rstrip(":") for c in cells]
    if not any(lows):
        return None
    keys_present = set()
    header, unknown = [], []
    for c in lows:
        if not c:
            header.append(None)
            continue
        if c not in _HEADER:
            unknown.append(c)
            header.append(None)
            continue
        header.append(_HEADER[c])
        if _HEADER[c]:
            keys_present.add(_HEADER[c])
    if "place" in keys_present and ("name" in keys_present or "first" in keys_present) and "time" in keys_present:
        return header, unknown
    return None


def _nonempty(row):
    return [_cell(c) for c in row if _cell(c)]


def _looks_header_text(text):
    low = text.lower()
    has_place = bool(re.search(r"\b(place|pl|pos)\b", low))
    has_name = bool(re.search(r"\b(name|runner|athlete)\b", low))
    has_time = bool(re.search(r"\btime\b|\bfinish\b", low))
    return has_place and has_name and has_time and not _TIME_TOKEN.search(text)


def _split_aligned(header_line, line):
    """Split one pasted line into the header's columns, using the header's character positions."""
    parts = [p for p in re.split(r" {2,}", header_line.strip()) if p]
    starts, pos = [], 0
    for p in parts:
        i = header_line.find(p, pos)
        starts.append(i if i >= 0 else pos)
        pos = (i if i >= 0 else pos) + len(p)
    cols = [c.strip() for c in re.split(r" {2,}", line.strip()) if c.strip()]
    if len(cols) == len(parts):
        return cols
    sliced = []
    for i, s in enumerate(starts):
        end = starts[i + 1] if i + 1 < len(starts) else len(line)
        sliced.append(line[s:end].strip() if s < len(line) else "")
    return sliced


def _expand_rows(rows):
    """Turn one-column pastes (tabs, or spaces lined up under a header) into cells."""
    out = []
    for row in rows:
        cells = _nonempty(row)
        if len(cells) == 1 and "\t" in cells[0]:
            out.append(tuple(c.strip() for c in cells[0].split("\t")))
        else:
            out.append(tuple(row))
    header_at, header_line = None, None
    for i, row in enumerate(out):
        cells = _nonempty(row)
        if len(cells) == 1 and _looks_header_text(cells[0]) and "  " in cells[0]:
            header_at, header_line = i, cells[0]
            break
    if header_at is None:
        return out
    expanded = []
    for i, row in enumerate(out):
        cells = _nonempty(row)
        if i >= header_at and len(cells) == 1:
            expanded.append(tuple(_split_aligned(header_line, cells[0])))
        else:
            expanded.append(row)
    return expanded


def _flip_last_first(name):
    if "," in name and len(name.split(",")) == 2:
        last, first = [x.strip() for x in name.split(",", 1)]
        if last and first:
            return f"{first} {last}"
    return name


def _school_from_rest(text):
    """Pull a known school off the end of the leftover text. 'Heights' counts as The Heights."""
    low = re.sub(r"\s+", " ", text).strip()
    if re.search(r"\bheights\b", low, re.I):
        name = re.sub(r"\b(the\s+)?heights(\s+school)?\b", "", text, flags=re.I)
        return re.sub(r"\s+", " ", name).strip(" ,-"), "The Heights"
    for school in _known_schools():
        if low == school or low.endswith(" " + school):
            cut = len(text) - len(school)
            return re.sub(r"\s+", " ", text[:cut]).strip(" ,-"), text[cut:].strip()
    return re.sub(r"\s+", " ", text).strip(), ""


def _parse_free_line(text, division):
    """One pasted line with no columns. -> ('ok', entry) | ('skip', reason) | ('junk', None)."""
    raw = text.strip()
    if not raw or _JUNK_LINE.match(raw) or _looks_header_text(raw):
        return "junk", None
    tm = _TIME_TOKEN.search(raw)
    if not tm:
        # "1 The Heights 45" is a team score. "1 Alex Runner The Heights" with no time is a row we could not read.
        if re.match(r"^\d+\.?\s+.+?\s+\d{1,3}$", raw):
            return "junk", None
        if re.match(r"^\d+\.?\s+\S+", raw):
            return "skip", "no finish time"
        return "junk", None
    time = N.clean_time(tm.group(1))
    if time is None:
        return "skip", f"can't read the time '{tm.group(1)}'"
    rest = (raw[:tm.start()] + " " + raw[tm.end():])
    rest = re.sub(r"\s+", " ", rest).strip()
    place = ""
    pm = re.match(r"^(\d+)\.?\s+", rest)
    if pm:
        place, rest = pm.group(1), rest[pm.end():].strip()
    grade = None
    tokens = rest.split()
    kept = []
    for tok in tokens:
        word = tok.strip(".").lower()
        if grade is None and word in _GRADE_WORD:
            grade = _GRADE_WORD[word]
            continue
        if grade is None and re.fullmatch(r"\d{1,2}", tok) and 6 <= int(tok) <= 12 and tok != place:
            grade = int(tok)
            continue
        if re.fullmatch(r"\d{3,6}", tok):
            continue  # bib
        kept.append(tok)
    name, school = _school_from_rest(" ".join(kept))
    name = _flip_last_first(name)
    if not name:
        return "skip", "no runner name left after the time"
    lo, hi = (4, 8) if division == "MS" else (9, 12)
    # an open high school race (e.g. NXR Rising Stars) can include other schools' middle schoolers
    other_school = division == "HS" and school and not H.HS_HEIGHTS_RE.match(school.strip())
    if grade is not None and not (lo <= grade <= hi) and not other_school:
        return "skip", f"grade {grade} isn't a {'middle' if division == 'MS' else 'high'} school grade ({lo}–{hi})"
    if place and time in ("DNF", "DNS", "DQ"):
        place = ""
    school = school or "The Heights"
    return "ok", {"place": place, "grade": str(grade) if grade else "", "name": name, "school": school, "time": time}


def _row_to_entry(row, division, lineno):
    """Column dict -> ('ok', entry) | ('skip', reason) | ('junk', None)."""
    lo, hi = (4, 8) if division == "MS" else (9, 12)
    first, last = row.get("first", ""), row.get("last", "")
    name = re.sub(r"\s+", " ", str(row.get("name") or f"{first} {last}")).strip()
    name = _flip_last_first(name)
    if not name and not str(row.get("time") or "").strip() and not str(row.get("place") or "").strip():
        return "junk", None
    if not name:
        return "skip", "the Name is empty"
    raw_time = row.get("time", "")
    if not str(raw_time).strip():
        return "skip", f"({name}): the Time is empty"
    t = N.clean_time(_cell(raw_time))
    if t is None:
        return "skip", f"({name}): can't read the time '{raw_time}'"
    place = str(row.get("place", "")).strip().rstrip(".")
    if place.upper() in ("DNF", "DNS", "DQ", "-", "--") or t in ("DNF", "DNS", "DQ"):
        place = ""
    elif place and not re.fullmatch(r"\d+", place):
        return "skip", f"({name}): the Place '{place}' should be a whole number"
    g = str(row.get("grade", "")).strip().rstrip(".")
    gnum = _GRADE_WORD.get(g.lower()) if g else None
    if g and gnum is None:
        gnum = N.CLASS_YEARS.get(g.lower())
    if g and gnum is None:
        if not re.fullmatch(r"\d{1,2}", g):
            return "skip", f"({name}): the Grade '{g}' should be a number or SR/JR/SO/FR"
        gnum = int(g)
    school = re.sub(r"\s+", " ", str(row.get("school") or "")).strip() or "The Heights"
    # an open high school race (e.g. NXR Rising Stars) can include other schools' middle schoolers
    other_school = division == "HS" and not H.HS_HEIGHTS_RE.match(school)
    if gnum is not None and not (lo <= gnum <= hi) and not other_school:
        return "skip", f"({name}): grade {gnum} isn't a {'middle' if division == 'MS' else 'high'} school grade ({lo}–{hi})"
    packed = {"Place": place, "Grade": str(gnum) if gnum else "", "Name": name, "School": school, "Time": t}
    if division == "HS":
        e = H.entry(packed, {})
        e["heights"] = bool(H.HS_HEIGHTS_RE.match(str(e["school"]).strip()))
    else:
        e = B.make_entry(place or None, gnum, name, school, t, row.get("pace"))
    return "ok", e


def field_from_tab(ws, division):
    """-> (entries, fatal problems, skipped-row notes).

    A row that can't be read is skipped and named in the summary. The tab is dropped only when
    more than half of the result-like rows fail, or when no Heights runner can be read.
    """
    rows = _expand_rows(list(ws.iter_rows(values_only=True)))
    header, start = None, 0
    for i, r in enumerate(rows):
        hit = _header_map(r)
        if hit:
            header, _unknown = hit
            start = i + 1
            break
    field, seen, skipped, attempted = [], {}, [], 0

    def take(kind, payload, label):
        nonlocal attempted
        if kind == "junk":
            return
        attempted += 1
        if kind == "skip":
            skipped.append(f"row {label}: {payload}")
            return
        e = payload
        if e["heights"]:
            k = B.norm_name(e["name"])
            if k in seen:
                skipped.append(f"row {label}: {e['name']} is listed twice (also row {seen[k]})")
                return
            seen[k] = label
        field.append(e)

    if header is not None:
        for n, r in enumerate(rows[start:], start=start + 1):
            cells = [_cell(c) for c in r]
            if not any(cells):
                continue
            joined = " ".join(c for c in cells if c)
            if _JUNK_LINE.match(joined) or _looks_header_text(joined):
                continue
            row = {}
            for key, c in zip(header, cells):
                if key:
                    row[key] = c
            # no finish time anywhere: a team-score line or page furniture, not a runner we failed to read
            if not _TIME_TOKEN.search(joined) and not str(row.get("time") or "").strip():
                continue
            kind, payload = _row_to_entry(row, division, n)
            take(kind, payload, n)
    else:
        for n, r in enumerate(rows, start=1):
            text = " ".join(c for c in (_cell(x) for x in r) if c)
            if not text:
                continue
            kind, payload = _parse_free_line(text, division)
            if kind == "ok":
                kind, payload = _row_to_entry(payload, division, n)
            take(kind, payload, n)
    fatal = []
    if attempted and len(skipped) > attempted / 2:
        fatal.append(f"more than half the rows could not be read ({len(skipped)} of {attempted}).")
        fatal.extend(skipped[:12])
        skipped = []
    elif not field:
        fatal.append("no runners could be read. Paste the results with a place, a name, and a time.")
    elif not any(e["heights"] for e in field):
        fatal.append("no runner is from The Heights. The school should say 'The Heights', or be left empty.")
    return field, fatal, skipped


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
    year = parsed.get("season") or parsed["date"].year
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
    season = parsed.get("season") or d.year
    dist = parsed["distance"]
    metres, miles, dkey = dist if dist else (None, None, None)
    only_heights = bool(field) and all(e["heights"] for e in field)
    common = {
        "season": season, "division": parsed["division"], "date": d.isoformat() if d else None, "dateLabel": None,
        "tab": parsed["title"], "origin": "sheet", "distance": miles,
        "distanceSource": "official" if miles else None, "statedDistance": miles, "impliedDistance": None,
        "officialDistance": None, "location": parsed.get("venue"), "source": None, "correction": None,
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
            "id": B.slug(f"hs-{season}-{parsed['meet']}-{label}"),
        })
    else:
        common.update({
            "level": "Open", "levelLabel": "Middle school",
            "name": parsed["meet"], "id": B.slug(f"{season}-{parsed['meet']}-{d.isoformat() if d else ''}"),
        })
    return common


def official_rows(meet):
    off = (meet.get("scoring") or {}).get("official")
    if not off:
        return None
    return [{"Place": t["place"], "Score": t["score"], "Team": t["team"], "SourceURL": off.get("source")} for t in off["teams"]]


# Labels written in a cell, with the value in the next cell to the right (or the cell below):
# "Venue:" | "WakeMed Soccer Park, Cary, NC". A trailing colon is needed for Meet and Distance, because a
# results table can have a column headed "Meet" or "Distance".
_DETAIL_KEYS = {"venue": "venue", "location": "venue", "course": "venue", "site": "venue",
                "meet": "meet", "meet name": "meet", "distance": "distance"}
_ALWAYS_LABEL = {"venue", "location", "course"}


def sheet_details(ws, max_rows=40):
    """-> {'venue': ..., 'meet': ..., 'distance': ...} from 'Label:' cells near the top of a tab."""
    rows = [[_cell(c) for c in r] for r in ws.iter_rows(max_row=max_rows, values_only=True)]
    found = {}
    for i, r in enumerate(rows):
        for j, c in enumerate(r):
            low = c.lower().strip()
            key = low.rstrip(":").strip()
            if key not in _DETAIL_KEYS or not (low.endswith(":") or key in _ALWAYS_LABEL):
                continue
            field = _DETAIL_KEYS[key]
            if field in found:
                continue
            value = next((v for v in r[j + 1:] if v), "")
            if not value and i + 1 < len(rows) and j < len(rows[i + 1]):
                value = rows[i + 1][j]
            if value and value.lower().rstrip(":") not in _DETAIL_KEYS:
                found[field] = re.sub(r"\s+", " ", value).strip()
    return found


def _words(text):
    return [w for w in re.split(r"\s+", text.strip()) if w]


def with_sheet_details(tabs):
    """Fill in venue, distance and meet name from the tab's own cells.

    When the tab name has no meet name ('Nov 22 2026 - 5k - Varsity'), the meet name comes from a
    'Meet:' cell or the Meet column. Tabs of the same day whose names share a beginning share one
    meet page: 'NXR Southeast Regionals Championship' and 'NXR Southeast Regionals Rising Stars'
    become the meet 'NXR Southeast Regionals' with the races 'Championship' and 'Rising Stars'.
    """
    out, from_sheet = [], []
    for order, ws, parsed in tabs:
        parsed = dict(parsed)
        try:
            details = sheet_details(ws)
        except Exception:
            details = {}
        if details.get("venue"):
            parsed["venue"] = details["venue"]
        if not parsed.get("distance") and details.get("distance"):
            got = find_distance(details["distance"])
            if got:
                parsed["distance"] = got[0]
        if not parsed.get("meet"):
            full = details.get("meet") or meet_column(ws)
            full = _cleanup(re.sub(r"\b(?:19|20)\d{2}\b", " ", full))
            if full:
                parsed["meet"] = full
                parsed["meetFromSheet"] = True
                from_sheet.append(parsed)
        out.append((order, ws, parsed))
    # same day, same division: split the shared beginning (the meet) from the rest (the race)
    by_day = {}
    for p in from_sheet:
        by_day.setdefault((p.get("date"), p.get("season"), p["division"]), []).append(p)
    for group in by_day.values():
        names = {p["meet"] for p in group}
        if len(group) < 2 or len(names) < 2:
            for p in group:
                _strip_level_word(p)
            continue
        split = [_words(n) for n in names]
        common = []
        for ws_ in zip(*split):
            if len({w.lower() for w in ws_}) != 1:
                break
            common.append(ws_[0])
        if not common or any(len(w) == len(common) for w in split):
            for p in group:
                _strip_level_word(p)
            continue
        meet = " ".join(common).strip(" -–,")
        for p in group:
            race = " ".join(_words(p["meet"])[len(common):]).strip(" -–,")
            p["meet"] = meet
            if race and p["division"] == "HS":
                p["levelLabel"] = p["race"] = race
    return out


def _strip_level_word(p):
    """A lone tab: '2026 Bullis Invitational Varsity' in the Meet column -> meet 'Bullis Invitational'."""
    found = find_level(p["meet"])
    if found and found[0] == p["division"] and _cleanup(found[3]):
        p["meet"] = _cleanup(found[3])


def apply_titled_tabs(tabs, meets, quirks):
    """tabs: list of (order, worksheet, parsed title). Mutates meets. -> list of problem strings."""
    problems = []
    tabs = with_sheet_details(tabs)
    for order, ws, parsed in tabs:
        parsed = dict(parsed, title=ws.title)
        if parsed.get("dateNote"):
            print(f"note: tab {ws.title!r}: {parsed['dateNote']}.")
        if parsed["defaultedMS"]:
            hit = hs_level_from_meet_column(ws)
            if hit:
                lvl, label, _ = hit
                parsed.update(division="HS", level=lvl, levelLabel=label, race=label, defaultedMS=False,
                              levelNote=f"read as high school {label} from the Meet column")
        if parsed.get("levelNote"):
            print(f"note: tab {ws.title!r} {parsed['levelNote']}.")
        if parsed["defaultedMS"]:
            print(f"note: tab {ws.title!r} has a date but no level (Varsity, JV, Freshman, Open, or MS); treated as middle school.")
        if not parsed["meet"]:
            problems.append(f"tab {ws.title!r} was NOT added:\n    - the tab name has a date, but no meet name left after the date, distance and level.")
            continue
        try:
            field, probs, skipped = field_from_tab(ws, parsed["division"])
        except Exception as ex:
            problems.append(f"tab {ws.title!r} was NOT added:\n    - couldn't read the tab ({type(ex).__name__}: {ex}).")
            continue
        for s in skipped:
            print(f"note: tab {ws.title!r} skipped {s}.")
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
            meta["location"] = parsed.get("venue") or replace.get("location")
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
                meta["id"] = B.slug(f"hs-{meta['season']}-{canon}-{parsed['levelLabel']}")
        meets.append(B.finalize(field, meta, quirks))
    return problems
