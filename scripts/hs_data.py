"""High school results (data/hs/): one CSV per race plus races.csv (one row per race), official team scores
and the MS<->HS runner matches. Imported by build_data.py; produces meet dicts in the same shape as the MS races
(division "HS"), plus the name merges used to build one runner page per person."""
import csv, os, re, json, itertools, difflib, datetime, collections

import build_data as B

HS = os.path.join(B.ROOT, "data", "hs")

# The Heights' own team names in HS results. Must be the whole school field, so "Dixie Heights" is not us.
HS_HEIGHTS_RE = re.compile(r"^(the\s+)?heights(\s+schoo+l)?(\s+boys\s+[a-c])?$", re.I)

# HS-only spelling variants of the same Heights runner (official HS spellings win; the most-used one is kept).
# One line per variant. Variants that are only *possibly* the same runner are deliberately left out
# (see NOT_MERGED below) so they stay separate until confirmed.
HS_NAME_ALIASES = {
    "Timothy Vaughn": "Timothy Vaughan",
    "Chrictopher Chmura": "Christopher Chmura",
    "Peyton Petersen": "Peyton Peterson",
    "Luke McLorey": "Luke McClorey",
    "Joseph McLorey": "Joseph McClorey",
    "Joe McClorey": "Joseph McClorey",
    "Joesph Cremers": "Joseph Cremers",
    "Johnny Ortman": "John Ortman",
    "Phillip Allen": "Philip Allen",
    "Nick Monahan": "Nicholas Monahan",
    "Creegan Cahill": "Cregan Cahill",
    "Gregan Cahill": "Cregan Cahill",
    "Gregory Lipovsky": "Greg Lipovsky",
    # confirmed by the team (Oct 2026)
    "Nick O'Neill": "Nicholas O'Neil",
    "Nicholas Monaghan": "Nicholas Monahan",
    "Nick Morahan": "Nicholas Monahan",
}
# Listed in hs_name_variants.csv but kept separate: each is also a "needs confirmation" MS-HS pair.
NOT_MERGED = []  # all earlier "possible same runner" pairs were confirmed by the team (Oct 2026) and merged

# Same-name HS runners that really are two people (both "Will Cannon": class of 2025, also listed as William
# Cannon, and class of 2027), so they are never re-joined despite never racing on the same day.
KEEP_SPLIT = {"Will Cannon"}

# One school listed two ways inside the same race (other team's runners only; used for scoring and de-duplication).
HS_SCHOOL_ALIASES = {"Cornerstone Christian Academy (Middleburg)": "Cornerstone Christian Academy"}

# Official team-score lists that use an old or short school name. Only used to pair an official row with the team
# in the results; the official name is still shown. (Renamed schools: J.E.B. Stuart -> Justice 2017, Wilson ->
# Jackson-Reed 2022, T.C. Williams -> Alexandria City 2021, Robert E. Lee -> John R. Lewis 2019,
# George Mason -> Meridian 2021.)
HS_OFFICIAL_TEAM_ALIASES = [
    (r"^(j\.? ?e\.? ?b\.? ?stuart|jeb stuart|stuart, jeb)\b", "Justice"),
    (r"^wilson$", "Jackson-Reed"),
    (r"^t\. ?c\. williams\b", "Alexandria City"),
    (r"^lee r e$", "John R. Lewis HS"),
    (r"^george mason\b", "Meridian"),
    (r"^magruder\b", "Col. Zadok Magruder"),
    (r"^mcnamara\b", "Bishop McNamara"),
    (r"^glenelg boys small$", "Glenelg Country"),
    (r"^episcopal (high school|hs va)$", "Episcopal"),
    (r"^philadelphia-montgomery\b", "Phil-Mont Christian Academy"),
    (r"^ss$", "St Stephens and St Agnes"),
]


def load_scoring_overrides():
    p = os.path.join(HS, "scoring_overrides.json")
    return json.load(open(p, encoding="utf-8")).get("races", {}) if os.path.exists(p) else {}


SCORING_OVERRIDES = None


def apply_scoring_override(meta, field):
    """Per-race scoring rules from data/hs/scoring_overrides.json (keyed by race CSV name), used only where a stated
    or clearly evidenced meet rule reproduces the official team scores."""
    global SCORING_OVERRIDES
    if SCORING_OVERRIDES is None:
        SCORING_OVERRIDES = load_scoring_overrides()
    o = SCORING_OVERRIDES.get(meta["csv"])
    if not o:
        return
    meta["scoreIncomplete"] = bool(o.get("incomplete_teams_score"))
    meta["scoreByOverall"] = bool(o.get("score_by_overall_place"))
    meta["scoreExactSchool"] = bool(o.get("separate_listed_teams"))
    if o.get("partial_teams"):
        meta["_scorePartialOk"] = {B.team_key(t) for t in o["partial_teams"]}
    if o.get("official_teams_only") and meta.get("_official"):
        schools = sorted({e["school"] for e in field})
        keep = set()
        for r in meta["_official"]:
            c = B.match_team(r["Team"], [x for x in schools if B.team_key(x) not in keep])
            if c:
                keep.add(B.team_key(c))
        meta["_scoreTeamsOnly"] = keep
    if o.get("note"):
        meta["scoreRuleNote"] = o["note"]


def explain_official(m):
    """Short reason on an HS race page when the calculated team scores don't all match the official ones."""
    o = m["scoring"].get("official")
    if not o or o["matched"] == len(o["teams"]):
        return
    bad = [r for r in o["teams"] if not r["match"]]
    cnt = collections.Counter(B.team_key(e["school"]) for e in m["results"] if isinstance(e["place"], int) and e["secs"])
    schools = sorted({e["school"] for e in m["results"]})
    short = []
    for r in o["teams"]:
        c = r["computedTeam"] or B.match_team(r["team"], schools)
        if not c or cnt[B.team_key(c)] < 5:
            short.append(r["team"])
    diffs = [abs(r["computedScore"] - r["score"]) for r in bad if r["computedScore"] is not None]
    why = o.get("why")
    if not why:
        if short:
            why = (f"{', '.join(short[:4])}{' and others' if len(short) > 4 else ''} {'is' if len(short) == 1 else 'are'} in the official "
                   f"team scores but {'has' if len(short) == 1 else 'have'} fewer than 5 runners in the posted results, so the posted "
                   "results are missing finishers and not every calculated score can match.")
        elif diffs and max(diffs) == 0:
            why = "Every calculated score matches the official one; only some places differ (tie-breaks)."
        elif diffs and max(diffs) <= 5:
            why = (f"The calculated scores are within {max(diffs)} point{'s' if max(diffs) > 1 else ''} of the official ones. The meet's "
                   "scoring counted a few runners differently from the posted list (for example runners entered as individuals rather "
                   "than for their team, or corrections made after the list was posted); which runners can't be told from the posted results.")
        else:
            why = (f"The official scores can't be reproduced from the posted results (differences of up to {max(diffs) if diffs else '?'} points). "
                   "The meet scored from its own entries or rules, which the posted results don't show.")
    if m.get("doubled"):
        why += (" The source lists nearly every runner twice; scoring that doubled list matches none of the official scores, so the "
                "official scores were tallied on the single list used here.")
    o["why"] = why


def official_team_alias(team, schools):
    """Field school name an official team row refers to, when only an old/short name links them."""
    for pat, to in HS_OFFICIAL_TEAM_ALIASES:
        if re.search(pat, team.strip(), re.I) and to in schools:
            return to
    return None

# Woodward Relays 2008 prints truncated two-man team names; map each to the runner's HS name.
RELAY_NAME_FIX = {"Tim Hadley": "Timothy Hadley", "Tim Vaughn": "Timothy Vaughan", "Sebastian Luli": "Sebastian Lulli",
                  "Greg Had": "Gregory Hadley"}

# MS runner id -> HS name, for MS-HS pairs not in the confident list of ms_hs_matches.csv.
# To merge a "needs confirmation" pair once it's confirmed, add ONE line here, e.g.
#   "nick-monahan": "Nicholas Monaghan",
EXTRA_MS_HS_MERGES = {
    # confirmed by the team (Oct 2026)
    "anthony-giantino": "Anthony Giuntini",
    "tomoyo-kanno": "Tamoya Kano",
    "luca-tocci": "Lucca Tocci",
    "steven-mccreary": "Stephen McCreary",
    "david-krohn": "David Krohn",
}

# Runner pages to join into one person regardless of the listed grades (ids as the build would otherwise make them;
# the first id is kept, the others redirect to it). Confirmed by the team (Oct 2026): every Will / William Cannon
# result is the same runner, even though the listed grades point to two different class years.
CONFIRMED_SAME_PERSON = [
    ["william-cannon", "will-cannon-2022", "will-cannon-2024", "will-cannon"],
]

ROMAN = {"II", "III", "IV", "V", "VI"}


def proper_case(name):
    """'Benjamin MAGEE' -> 'Benjamin Magee', 'KEVIN UNG' -> 'Kevin Ung', 'MCDONALD' -> 'McDonald'; keeps III, C.J., AJ."""
    out = []
    for tok in name.split(" "):
        letters = re.sub(r"[^A-Za-z]", "", tok)
        if len(letters) > 2 and letters.isupper() and tok.strip(".,") not in ROMAN and "." not in tok:
            t = tok.lower().title()  # handles O'Brien, Smith-Jones
            t = re.sub(r"\bMc([a-z])", lambda m: "Mc" + m.group(1).upper(), t)
            t = t.replace("'S", "'s") if t.endswith("'S") else t
            out.append(t)
        else:
            out.append(tok)
    return " ".join(out)


def parse_date(s):
    m = re.match(r"([A-Z][a-z]{2,8})\.? (\d{1,2}), (\d{4})", s or "")
    if not m:
        return None
    try:
        return datetime.datetime.strptime(f"{m.group(1)[:3]} {m.group(2)} {m.group(3)}", "%b %d %Y").date().isoformat()
    except ValueError:
        return None


def parse_distance(s):
    """-> (metres, miles, key) ; key groups races for records ('5K', '3 mi', ...)."""
    t = (s or "").lower()
    if "relay" in t:
        if "2.5 km" in t:
            return 2500.0, 2500 / 1609.344, "2.5K"
        m = re.search(r"(\d+(?:\.\d+)?)\s*mile", t)
        return (float(m.group(1)) * 1609.344, float(m.group(1)), f"{m.group(1)} mi") if m else (None, None, None)
    m = re.search(r"(\d+(?:\.\d+)?)\s*mi(le)?\b", t)
    if m and "m (" not in t:
        mi = float(m.group(1)); return mi * 1609.344, mi, f"{mi:g} mi"
    m = re.search(r"(\d+(?:\.\d+)?)\s*km", t)
    if m:
        km = float(m.group(1)); return km * 1000, km * 1000 / 1609.344, f"{km:g}K"
    m = re.search(r"(\d{4})\s*m\b", t)
    if m:
        metres = float(m.group(1))
        if metres == 5000:
            return 5000.0, 5000 / 1609.344, "5K"
        for mi in (2, 3):
            if abs(metres - mi * 1609.344) < 2:
                return mi * 1609.344, float(mi), f"{mi} mi"
        return metres, metres / 1609.344, f"{int(metres):,} m"
    if re.search(r"\b5\s*k\b", t):
        return 5000.0, 5000 / 1609.344, "5K"
    return None, None, None


def level_of(s):
    """-> (level, label, note). level is one of Varsity, JV, Freshman, Open, Section, Relay."""
    L, race, meet = s["Level"], s["Race"].strip(), s["Meet"]
    if "WCAC" in meet.upper() and re.search(r"varsity\s*b\b", race, re.I):
        return "JV", "Varsity B", "WCAC's \u201cVarsity B\u201d race is its second-team race, so it is listed at the JV level."
    if L.startswith("Relay"):
        if "2-man" in L:
            return "Relay", "2-man relay (6 mi)", None
        sec = L.split("\u2013")[-1].strip()
        return "Relay", f"5\u00d72.5K relay ({sec})", None
    if L.startswith("Open (individual"):
        return "Open", "Individual 2.5K", ("Individual 2.5 km race held alongside the RVA Relays. Places are as posted; "
                                          "this race was posted in sections, so overall place numbers skip.")
    if L.startswith("Open (Boys overall"):
        return "Open", "Open (varsity & JV together)", ("Varsity and JV ran together in one race (COVID season). "
                                                       "The source marks each runner's division; places here are overall.")
    if L.startswith("Section"):
        return "Section", re.sub(r"\s+", " ", race).title().replace("Under-Class", "Under-Class"), "The results don't say which level this section was."
    if L.startswith("Unlabeled"):
        return "Freshman", "Freshman (likely)", None
    base = L.split(" (")[0]
    base = {"Varsity": "Varsity", "JV": "JV", "Freshman": "Freshman", "Open": "Open"}.get(base, base)
    if base == "Open" and "mixed" in L:
        return "Open", "Open", None
    return base, L, None


RACE_DROP = re.compile(r"\b(boys|men's|cc|run|individual results:|kilometer|meter|championships?|wcac|cross country|rva|relays?boys|relays|5k|5000m?)\b", re.I)


def race_short(s, level, label):
    r = s["Race"].strip()
    if r in ("(untitled)",) or level == "Section":
        return label
    r = re.sub(r"^\d{4}\s+", "", r)
    for w in s["Meet"].split():
        if len(w) > 4 and w.isalpha():
            r = re.sub(rf"\b{re.escape(w)}\b", "", r, flags=re.I)
    r = re.sub(r"12\.5", "", r)
    r = RACE_DROP.sub("", r)
    r = re.sub(r"\s*-\s*$", "", re.sub(r"\s+", " ", r)).strip(" -:")
    r = re.sub(r"\s+-(\s|$)", " ", r).strip()
    if r.isupper():
        r = r.title().replace("Jv", "JV")
    if level == "Relay":
        r = (r + " relay").strip() if r and "relay" not in r.lower() else (r or "Relay")
    return r or label


def entry(row, grade_fix):
    raw = re.sub(r"\s+", " ", (row.get("Name") or "").strip())
    leg = None
    m = re.search(r"\s*\(leg (\d)\)$", raw)
    if m:
        leg, raw = int(m.group(1)), raw[:m.start()]
    name = proper_case(raw)
    g = (row.get("Grade") or "").strip().strip("'").strip("-")
    e = B.make_entry(row.get("Place"), g or None, name, row.get("School"), row.get("Time"), None)
    e["name"] = name
    e.pop("note", None)
    e["heights"] = bool(HS_HEIGHTS_RE.match(e["school"].strip()))
    if e["school"] in HS_SCHOOL_ALIASES:
        e["listedSchool"], e["school"] = e["school"], HS_SCHOOL_ALIASES[e["school"]]
    if e["heights"]:
        new = HS_NAME_ALIASES.get(name, name)
        new = re.sub(r"\bLuli$", "Lulli", new)
        if new != name:
            e["note"] = f"Listed as \u201c{raw}\u201d"
            e["name"] = new
            e["listed"] = name
        elif name != raw:
            e["note"] = f"Listed as \u201c{raw}\u201d"
        if raw in grade_fix:
            e["note"] = (e.get("note", "") + "; " if e.get("note") else "") + f"Listed as grade {e['grade']}"
            e["grade"] = grade_fix[raw]
    if leg:
        e["leg"] = leg
    return e


def _nu(u):
    return re.sub(r"^https?://(www\.)?", "", u or "").strip()


def load_official():
    by_key, by_url = {}, collections.defaultdict(list)
    for r in csv.DictReader(open(os.path.join(HS, "official_team_scores.csv"), encoding="utf-8")):
        by_key.setdefault(r["Meet"], []).append(r)
        by_url[(_nu(r["SourceURL"]), r["Level"])].append(r)
    return by_key, by_url


def split_blocks(rows):
    blocks, cur = [], []
    for r in rows:
        if cur and int(r["Place"]) <= int(cur[-1]["Place"]):
            blocks.append(cur); cur = []
        cur.append(r)
    return blocks + ([cur] if cur else [])


def relay_teams(field, two_man):
    teams = []
    if two_man:
        for e in field:
            names = [n.strip() for n in e["name"].split(",")]
            unnamed = e["name"].lower().startswith("relay")
            teams.append({"place": e["place"], "team": e["school"], "time": e["time"], "secs": e["secs"], "heights": e["heights"],
                          "legs": [] if unnamed else [{"name": RELAY_NAME_FIX.get(n, n) if e["heights"] else n, "listed": n} for n in names],
                          "label": e["name"] if unnamed else None})
        return teams
    cur = None
    for e in field:
        if e["name"].lower() == "(team total)":
            cur = {"place": e["place"], "team": e["school"], "time": e["time"], "secs": e["secs"], "heights": e["heights"], "legs": []}
            teams.append(cur)
        elif cur is not None:
            cur["legs"].append(e)
    return teams


def build_hs_meets(quirks):
    notes = json.load(open(os.path.join(HS, "race_notes.json")))
    rnotes, gfix = notes.get("races", {}), notes.get("grade_fixes", {})
    off_key, off_url = load_official()
    rows = list(csv.DictReader(open(os.path.join(HS, "races.csv"), encoding="utf-8")))
    pending, url_users = [], collections.defaultdict(list)
    for s in rows:
        base = os.path.basename(s["CSV"])
        field = [entry(r, gfix.get(base, {})) for r in csv.DictReader(open(os.path.join(HS, "races", base), encoding="utf-8"))]
        for x in notes.get("extra_rows", {}).get(base, []):  # e.g. a DQ the source lists apart from the finishers
            e = entry(x, {})
            e["note"] = x.get("note") or e.get("note")
            if e["time"] == "DQ":
                e["dq"] = True
            field.append(e)
        drops = notes.get("drop_rows", {}).get(base, [])
        if drops:
            gone = [e for e in field if any(str(e["place"]) == str(x["Place"]) and e["name"] == x["Name"] for x in drops)]
            if len(gone) != len(drops):
                quirks.append(f"HS {base}: drop_rows matched {len(gone)} of {len(drops)} rows.")
            field = [e for e in field if e not in gone]
            dp = [g["place"] for g in gone if isinstance(g["place"], int)]
            for e in field:
                if isinstance(e["place"], int):
                    e["place"] -= sum(1 for x in dp if x < e["place"])
        level, label, lnote = level_of(s)
        lov = notes.get("level_overrides", {}).get(base)
        old_short = race_short(s, level, label) if lov else None
        if lov:
            level, label, lnote = lov["level"], lov["label"], lov["note"]
        metres, miles, dkey = parse_distance(s["Distance"])
        season = int(s["Season"])
        meet = re.sub(r"^\d{4}\s+", "", s["Meet"]).strip()
        short = lov["short"] if lov else race_short(s, level, label)
        rn = list(rnotes.get(base, []))
        if lnote:
            rn.insert(0, lnote)
        if "per pace" in s["Distance"]:
            rn.append("The results don't state the distance; 5K is inferred from the posted paces.")
        if metres is None and level != "Relay":
            rn.append("The results don't list the race distance, so no paces are shown.")
        # Some racecast pages list every finisher twice in a row. Drop the second copy (same time, school and
        # grade, same or near-identical name) and renumber places.
        def _same(a, b):
            if not a["secs"] or a["secs"] != b["secs"] or B.team_key(a["school"]) != B.team_key(b["school"]) or a["grade"] != b["grade"]:
                return False
            return a["name"] == b["name"] or difflib.SequenceMatcher(None, a["name"].lower(), b["name"].lower()).ratio() >= 0.8
        pairs = sum(1 for a, b in zip(field, field[1:]) if a["name"] == b["name"] and a["secs"] and a["secs"] == b["secs"] and B.team_key(a["school"]) == B.team_key(b["school"]))
        if pairs > 0.3 * len(field):
            ded = []
            for e in field:
                prev = next((d for d in reversed(ded[-4:]) if not d.get("_dup") and _same(d, e)), None)
                if prev:
                    prev["_dup"] = True
                    if e["name"] != prev["name"] and e["heights"]:
                        rn.append(f"\u201c{e['name']}\u201d is listed right after \u201c{prev['name']}\u201d with the same time; treated as the same runner.")
                    continue
                ded.append(e)
            for e in ded:
                e.pop("_dup", None)
            n_before = len(field)
            places = [e["place"] for e in ded if isinstance(e["place"], int)]
            for i, e in enumerate([e for e in ded if isinstance(e["place"], int)]):
                e["place"] = i + 1
            field = ded
            meta_doubled = True
            rn.append(f"The source page lists almost every finisher twice ({n_before} rows for {len(field)} runners). "
                      "The duplicates are removed here and places renumbered.")
            quirks.append(f"HS {season} {meet} \u2013 {short}: source lists finishers twice ({n_before} rows); de-duplicated to {len(field)} and renumbered.")
        else:
            meta_doubled = False
        # duplicate Heights rows (same race, same time, near-identical name), e.g. "John Ortman" / "Johnny Ortmann"
        dropped = []
        hs_rows = [e for e in field if e["heights"]]
        for a, b in itertools.combinations(hs_rows, 2):
            if a in dropped or b in dropped or not a["secs"] or a["secs"] != b["secs"] or a.get("leg") or b.get("leg"):
                continue
            la, lb = a["name"].split()[-1].lower(), b["name"].split()[-1].lower()
            if difflib.SequenceMatcher(None, la, lb).ratio() >= 0.8 and a["name"].split()[0][:3].lower() == b["name"].split()[0][:3].lower():
                dropped.append(b)
                rn.append(f"The source lists \u201c{b['name']}\u201d ({ordinal(b['place'])}) right after \u201c{a['name']}\u201d ({ordinal(a['place'])}) with the same time "
                          f"({a['time']}); it is the same runner listed twice, so the duplicate row is left out.")
                quirks.append(f"HS {season} {meet} \u2013 {short}: dropped duplicate row \u201c{b['name']}\u201d (same time as {a['name']}).")
        field = [e for e in field if e not in dropped]
        status = s["TeamScores"]
        if status.startswith("official (mocorunning meet printout; racecast lists computed only)"):
            rn.append("Racecast's team scores for this race are its own calculation; the official scores shown are from the meet printout on mocorunning.com.")
        src_url = s["SourceURL"]
        src_name = "racecast.co" if "racecast" in src_url else "mocorunning.com" if "mocorunning" in src_url else s["Source"]
        meta = {
            "id": B.slug(f"hs-{season}-{meet}-{short}"), "season": season, "division": "HS",
            "level": level, "levelLabel": label, "race": s["Race"].strip(), "meetName": meet,
            "name": f"{meet} \u2013 {short}", "dateLabel": None, "date": parse_date(s["Date"]),
            "tab": None, "origin": "hs", "distance": miles, "distanceM": metres, "distKey": dkey,
            "distanceSource": "official" if miles else None, "statedDistance": None, "impliedDistance": None, "officialDistance": None,
            "officialName": f"{s['Meet']} \u2013 {s['Race'].strip()}",
            "location": s["Location"] or None,
            "source": {"name": src_name, "url": src_url, "original": s["OriginalSource"] or None},
            "correction": None, "note": " ".join(rn) or None, "notes": rn, "completeness": s["Completeness"],
            "teamScoreStatus": status, "csv": base, "_tabOrder": 0, "doubled": meta_doubled,
            "aliasIds": [B.slug(f"hs-{season}-{meet}-{old_short}")] if lov else None,
        }
        if level == "Relay":
            pending.append((meta, field, None)); continue
        off = off_key.get(f"{s['Season']} {s['Meet']} \u2013 {s['Race']}")
        ukey = None
        if not off:
            u = _nu(s["OfficialTeamSource"] or (src_url if "mocorunning" in src_url else ""))
            if u and (u, s["Level"]) in off_url:
                ukey = (u, s["Level"]); url_users[ukey].append(len(pending))
        pending.append((meta, field, off))
    # one official list shared by several races (Landon 2011 Brown/White): split by place restarts and give each
    # race the block that matches its own calculated scores best
    for ukey, idx in url_users.items():
        blocks = split_blocks(off_url[ukey])
        if len(idx) == 1:
            pending[idx[0]] = (*pending[idx[0]][:2], off_url[ukey] if len(blocks) == 1 else max(blocks, key=len)); continue
        def fit(i, blk):
            st, _ = B.score_race(pending[i][1])
            comp = {B.team_key(t["team"]): t["score"] for t in st}
            return sum(1 for o in blk if any(abs(v - int(o["Score"])) <= 0 and B.match_team(o["Team"], [k]) for k, v in comp.items()))
        best = max(itertools.permutations(range(len(blocks)), len(idx)), key=lambda p: sum(fit(i, blocks[b]) for i, b in zip(idx, p)))
        for i, b in zip(idx, best):
            pending[i] = (*pending[i][:2], blocks[b])
        quirks.append(f"HS: one official team-score list covers {len(idx)} races ({', '.join(pending[i][0]['name'] for i in idx)}); split by place restarts and matched to each race by its scores.")
    meets = []
    for meta, field, off in pending:
        if meta["level"] == "Relay":
            meets.append(finalize_relay(meta, field)); continue
        listed = {}
        if off:
            schools = {e["school"] for e in field}
            off2 = []
            for o in off:
                a = official_team_alias(o["Team"], schools)
                if a:
                    listed[(int(o["Place"]), int(o["Score"]))] = o["Team"]
                    o = dict(o, Team=a)
                off2.append(o)
            off = off2
        meta["_official"] = off
        apply_scoring_override(meta, field)
        m = B.finalize(field, meta, quirks)
        for r in (m["scoring"].get("official") or {}).get("teams", []):
            if (r["place"], r["score"]) in listed:
                r["team"] = listed[(r["place"], r["score"])]
        explain_official(m)
        if m["levelLabel"] == "Individual 2.5K":
            # posted in sections (and run alongside the relays), so a team score from the listed runners means nothing
            for e in m["results"]:
                e.pop("teamPlace", None)
            m["scoring"]["standings"], m["scoring"]["partialTeams"] = [], []
            m["scoring"]["caveats"] = ["No team scores: this individual race was run alongside the relays and posted in sections."]
        o = m["scoring"].get("official")
        meets.append(m)
    seen = collections.Counter()
    for m in meets:
        seen[m["id"]] += 1
        if seen[m["id"]] > 1:
            m["id"] += f"-{seen[m['id']]}"
    return meets


def ordinal(n):
    if not isinstance(n, int):
        return str(n)
    return f"{n}{'th' if 10 <= n % 100 <= 20 else {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')}"


def finalize_relay(meta, field):
    two_man = "2-man" in meta["levelLabel"]
    teams = relay_teams(field, two_man)
    results = []  # Heights runners, one row per leg (or per team member for the 2-man relay)
    leg_m = 2500.0 if not two_man else None
    for t in teams:
        for i, lg in enumerate(t["legs"]):
            if two_man:
                lg = {"name": lg["name"], "listed": lg["listed"], "place": t["place"], "grade": None, "school": t["team"], "time": t["time"],
                      "secs": t["secs"], "pace": None, "paceSecs": None, "heights": t["heights"], "leg": None}
                if lg["listed"] != lg["name"]:
                    lg["note"] = f"Listed as \u201c{lg['listed']}\u201d"
                t["legs"][i] = lg
            else:
                lg["place"] = t["place"]
                if lg["secs"]:
                    ps = int(lg["secs"] / (leg_m / 1609.344)); lg["paceSecs"], lg["pace"] = ps, f"{ps // 60}:{ps % 60:02d}"
            lg["relay"] = {"team": t["team"], "teamPlace": t["place"], "teamTime": t["time"], "leg": lg.get("leg")}
            if t["heights"]:
                lg["heights"] = True
                results.append(lg)
    hp = 0
    for t in teams:
        if t["heights"]:
            hp += 1; t["hplace"] = hp
    meta.update({"relay": {"teams": teams, "twoMan": two_man}, "scoring": {"standings": [], "partialTeams": [], "caveats": [], "sheetNotes": []},
                 "fieldSize": len(teams), "heightsCount": len(results), "hasPace": not two_man, "hasGrade": any(e.get("grade") for e in results),
                 "results": results})
    meta["notes"] = meta["notes"] + ["Relay results: team placings are shown with each runner's leg. Relay legs don't count toward personal records or the records pages."]
    meta["note"] = " ".join(meta["notes"])
    return meta


def load_matches():
    """MS runner id -> set of HS names (after HS spelling merges), confident rows only, plus EXTRA_MS_HS_MERGES."""
    out = collections.defaultdict(set)
    for r in csv.DictReader(open(os.path.join(HS, "ms_hs_matches.csv"), encoding="utf-8")):
        if r["Status"].strip() != "confident":
            continue
        for part in r["HS_names"].split("|"):
            n = re.sub(r"\s*\(x\d+\)\s*$", "", part.strip())
            n = proper_case(n)
            out[r["MS_id"]].add(HS_NAME_ALIASES.get(n, n))
    for k, v in EXTRA_MS_HS_MERGES.items():
        out[k].add(v)
    return out
