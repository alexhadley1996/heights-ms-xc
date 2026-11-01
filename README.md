# Heights Cross Country — results site (draft)

Middle school (MS) and high school (HS) results for The Heights School Cavaliers, one site.

Plain HTML/CSS/JS, no build step. Works on GitHub Pages as-is (`.nojekyll` included).

Pages: `index.html` (home, both divisions), `meets.html` (all races; filter `?div=MS|HS&level=Varsity|JV|Freshman|Open|Section|Relay&season=2010`),
`meet.html?id=…` (one race: Heights runners or full field, team scores), `runners.html` (roster; `?div=MS|HS|both`),
`runner.html?id=…` (one page per person: middle school and high school races, Division and Level columns, MS PB pace, HS PRs by distance),
`results.html` (All Results: every Heights individual result, MS + HS; sort by any column, shift-click for a secondary sort; filters for division, level, seasons, grade, meet, distance, runner, PRs only, best per runner, DQs/relay legs; CSV export; the view is kept in the URL),
`top.html` (top times, separate for HS — by distance, mostly 5K — and MS — by pace), `compare.html?r=id1,id2` (works across divisions),
`history.html` (seasons table for both divisions + placeholders).
Every results table is sortable: click a column heading (again to reverse).

Race distances are read from the divisor in each tab's pace formula (the script loads the workbook twice: values and formulas).
Team scores are calculated by the script (standard XC scoring) and compared with official published team scores where those were found (Official column on each meet page).

## Refreshing the data
1. In Google Sheets, File → Download → Microsoft Excel (.xlsx) for “Heights MS XC Results”.
2. `pip install openpyxl` then `python3 scripts/build_data.py path/to/file.xlsx`
3. This rewrites `data/results.js` (used by every page: race list, runners and their results), `data/results.json`
   (same content) and `data/races/<race id>.js` (full field + team scores for one race, loaded only by that race's page).

New meet tabs are picked up automatically; add a line to `MEET_INFO` in the script to set a nicer name/date/distance,
and `NAME_ALIASES` to merge spelling variants of a runner's name.

## Supplemental data and corrections (`data/supplemental/`)
The sheet stays the primary source; the build also merges:
- `manifest.json` + `races/*.csv` — races found online (racecast.co / Athletic.net / mocorunning.com) that are not in the sheet
  (2006 Sept 15, 2011–2019, 2021 CLS, 2022 FHS). Columns: Place,Grade,Name,Time,Pace,School,Meet. Distance is the official one;
  paces are recalculated from it. Each meet page links its source. Per-race rule flags: `score_no_displacer_limit`, `score_rule_note`.
- `sheet_overrides.json` — keyed by site race id: official date/location/distance/source for sheet races, plus fixes:
  `replace_field_with` (use the official field instead of the sheet's), `a_team_list`/`b_team_list` (2022 States: runners the
  meet re-scored as B teams keep their places but don't score; ties follow the official order), `fill_from_official` (keep sheet rows, add missing
  finishers, e.g. 2021 States, 2007 Landon Invitational), `score_exclude_grades` (e.g. 2025 States didn't score 5th graders).
  The sheet's distance is kept; an official distance within 0.01 mi counts as agreeing (e.g. 2024 Landon 1.6 vs 1.599).
- `replaces_supplemental` in an override: the sheet tab now covers that online race, so the online copy is not shown; its
  source link and official team scores are kept on the sheet race, and grades missing in the sheet are filled from it.
  Team-score summary tabs (header Race/Rank/Team/Score) are skipped automatically.
- `official_team_scores.csv` — official team places/scores (Season,Meet,Team,Place,Score,SourceURL).
In `MEET_INFO`: `keep_races` drops extra lists in a tab (the 2022 States B listing), `dist_from_team` marks a distance
supplied by the team (2010 Landon, 2.2 mi). Runners with a time but no place (e.g. 2025 States, Sean Carney, no bib) are
slotted in by time and later places renumbered. Missing Heights grades are inferred from the runner's other races (italic).
A runner name that appears in seasons/grades that can't be one runner is split into separate runner pages (HS pieces that never
raced on the same day are re-joined as a likely grade typo, except names in `KEEP_SPLIT`);
`identityReview` in results.json lists merges/splits worth a human check.

## High school data (`data/hs/`)
Built by the same script (`scripts/hs_data.py`, called from `build_data.py`); the sheet is not involved.
- `races.csv` — one row per race (season, date, meet, race, level, distance, location, sources, CSV file name); `races/*.csv` — full
  fields (Place,Grade,Name,Time,Pace,School,Meet,Level,SourceURL). Paces are recalculated from the listed distance.
- `official_team_scores.csv` — official team scores (racecast “as scored by meet” and mocorunning.com meet printouts). On HS race
  pages the official standings come first and the site's own calculation is shown as a crosscheck (✓ / ≠).
- `scoring_overrides.json` — per-race HS scoring rules, used only where the rule is what the meet did and it reproduces the
  official scores (overall-place scoring, only official teams score, an official team with under 5 posted runners still takes
  team places, A/B teams listed under two spellings scored separately). Each entry records its evidence; the note is shown on the race page.
  Old/short school names in official lists (J.E.B. Stuart → Justice, Wilson → Jackson-Reed, …) are paired via
  `HS_OFFICIAL_TEAM_ALIASES` in `scripts/hs_data.py` (HS only; the official name is still shown). Unmatched races get a short reason.
- `race_notes.json` — notes shown on a race page (data issues from the research notes), per-race grade fixes,
  `extra_rows` (e.g. Tim Hadley's 2007 Episcopal DQ: shown as DQ, no place/time, not counted in records) and
  `level_overrides` (a mislabelled race's level/label; the old race id redirects — 2010 MD/DC Small Schools is Varsity).
- One meet page per meet: all races of a meet (varsity, JV, freshman, …) are sections of the same `meet.html?id=<meet>`
  page (`groups` in results.json). An old race id (`meet.html?id=<race id>`) redirects to its meet page, scrolled to that race
  (`#<race id>`); `raceAliases` maps renamed race ids. `SEASON_NOTES` in build_data.py adds a note to a season (2020: COVID).
- One row per runner per race: `drop_rows` in race_notes.json removes a duplicate listing (2013 Landon Freshman: Edy McCormack,
  Griffin Woofter); `ROW_SPLITS` in build_data.py marks a row that belongs to a different runner with the same name (brothers:
  Robert Spence at 2019 Landon, Noah Schneider at 2016 Norwood). The build stops with an error if any runner still appears
  twice in one race.
- Runner merges confirmed by the team: `HS_NAME_ALIASES` (spellings), `EXTRA_MS_HS_MERGES` (MS→HS), `CONFIRMED_SAME_PERSON`
  (pages joined regardless of listed grades, e.g. all Will/William Cannon pages). Old runner ids redirect (`runnerAliases`).
- Athlete search: the box in the header of every page (type part of a name, any spelling the runner is listed under;
  Enter opens the first match). The Runners page also has a name filter (`runners.html?q=…`).
- `ms_hs_matches.csv` — MS↔HS runner matches. Only `confident` rows are merged. To merge a “needs confirmation” pair, add one line to
  `EXTRA_MS_HS_MERGES` in `scripts/hs_data.py` (`"ms-runner-id": "HS Name"`).
- `hs_name_variants.csv` — HS spelling variants; the ones merged are listed in `HS_NAME_ALIASES` in `scripts/hs_data.py`.
- Levels: Varsity, JV, Freshman, Open, Section (level not given — section name shown, e.g. Landon 2011 Brown/White), Relay.
  WCAC “Varsity B” is listed at JV level; Holy Cross 2020 (V/JV together) is Open; relays show team placings with legs and
  never count toward PRs or Top Times. Heights = school field exactly “The Heights”/“The Heights School”/“Heights School”
  (+ relay “Boys A/B/C”), so “Dixie Heights” is not counted. ALL-CAPS names are put in normal case.
- Sources that list every finisher twice (several 2025 racecast pages) are de-duplicated (same time, grade and school, school
  compared without punctuation) and places renumbered (noted on the page). Scoring the doubled list matches none of the
  official scores, so the official scores were tallied on the single list.
- Missing grades are inferred within each runner's whole career (MS 5–8, HS 9–12), shown in italics.

## Preview locally
`python3 -m http.server` in this folder, then open http://localhost:8000

Text marked “Placeholder” (home page About, History page, footer contact) still needs to be written.

## Branding sources (checked 2026-10-30)
- Colors: school colors red and white — https://heights.edu/about/quick-facts/ ("School Color: Red"), https://en.wikipedia.org/wiki/The_Heights_School_(Maryland) and https://www.wcacsports.com/members/theheights/index ("Red / White").
- Exact red `#841313` (and darker `#6e1010`): heights.edu theme stylesheet https://heights.edu/app/themes/heights-school/dist/styles/main_feb5b753.css (primary link/button/nav color).
- Mascot / team name: Cavalier / Cavaliers — heights.edu Quick Facts, Wikipedia, WCAC.
- Logo (provided by the team, Nov 2026): `assets/heights-logo-original.png` (red H over crossed sabers, white background). Made from it: `heights-logo.png` (transparent, red; footer), `heights-logo-white.png` (transparent, white; red header and home banner), favicons `favicon.ico` (16/32/48), `favicon-16/32/48.png`, `apple-touch-icon.png` (180), `icon-192.png`, `icon-512.png` (`site.webmanifest`). The school crest is no longer used.
- `assets/heights-men-fully-alive.svg` — https://heights.edu/app/themes/heights-school/resources/assets/images/mfa.svg (unaltered, not currently used)
- Fonts on heights.edu: Adobe Garamond SC (self-hosted) and Typekit `tablet-gothic` / `adobe-garamond-pro` (licensed). This site uses free EB Garamond (Google Fonts) in small caps as a stand-in, plus system sans-serif.
