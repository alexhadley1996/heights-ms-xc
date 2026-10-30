# Heights MS Cross Country — results site (draft)

Plain HTML/CSS/JS, no build step. Works on GitHub Pages as-is (`.nojekyll` included).

Pages: `index.html` (home), `meets.html` (all races by season), `meet.html?id=…` (one race: Heights runners or full field),
`runners.html` (searchable roster), `runner.html?id=…` (every race for one runner + PB), `top.html` (top times/paces),
`compare.html?r=id1,id2` (compare runners, or one runner's progression), `history.html` (placeholders).
Every results table is sortable: click a column heading (again to reverse).

Race distances are read from the divisor in each tab's pace formula (the script loads the workbook twice: values and formulas).
Team scores are calculated by the script (standard XC scoring) and compared with official published team scores where those were found (Official column on each meet page).

## Refreshing the data
1. In Google Sheets, File → Download → Microsoft Excel (.xlsx) for “Heights MS XC Results”.
2. `pip install openpyxl` then `python3 scripts/build_data.py path/to/file.xlsx`
3. This rewrites `data/results.js` (used by the site) and `data/results.json`.

New meet tabs are picked up automatically; add a line to `MEET_INFO` in the script to set a nicer name/date/distance,
and `NAME_ALIASES` to merge spelling variants of a runner's name.

## Supplemental data and corrections (`data/supplemental/`)
The sheet stays the primary source; the build also merges:
- `manifest.json` + `races/*.csv` — races found online (racecast.co / Athletic.net / mocorunning.com) that are not in the sheet
  (2006 Sept 15, 2011–2019, 2021 CLS, 2022 FHS). Columns: Place,Grade,Name,Time,Pace,School,Meet. Distance is the official one;
  paces are recalculated from it. Each meet page links its source. Per-race rule flags: `score_no_displacer_limit`, `score_rule_note`.
- `sheet_overrides.json` — keyed by site race id: official date/location/distance/source for sheet races, plus fixes:
  `replace_field_with` (use the official field, e.g. 2022 States Race 1), `fill_from_official` (keep sheet rows, add missing
  finishers, e.g. 2021 States, 2007 Landon Invitational), `score_exclude_grades` (e.g. 2025 States didn't score 5th graders).
  The sheet's distance is kept; an official distance within 0.01 mi counts as agreeing (e.g. 2024 Landon 1.6 vs 1.599).
- `official_team_scores.csv` — official team places/scores (Season,Meet,Team,Place,Score,SourceURL).
A runner name that appears in seasons/grades that can't be one middle-schooler is split into separate runner pages;
`identityReview` in results.json lists merges/splits worth a human check.

## Preview locally
`python3 -m http.server` in this folder, then open http://localhost:8000

Text marked “Placeholder” (home page About, History page, footer contact) still needs to be written.

## Branding sources (checked 2026-10-30)
- Colors: school colors red and white — https://heights.edu/about/quick-facts/ ("School Color: Red"), https://en.wikipedia.org/wiki/The_Heights_School_(Maryland) and https://www.wcacsports.com/members/theheights/index ("Red / White").
- Exact red `#841313` (and darker `#6e1010`): heights.edu theme stylesheet https://heights.edu/app/themes/heights-school/dist/styles/main_feb5b753.css (primary link/button/nav color).
- Mascot / team name: Cavalier / Cavaliers — heights.edu Quick Facts, Wikipedia, WCAC.
- `assets/heights-crest.svg` — https://heights.edu/app/themes/heights-school/resources/assets/images/crest.svg (unaltered)
- `assets/heights-men-fully-alive.svg` — https://heights.edu/app/themes/heights-school/resources/assets/images/mfa.svg (unaltered, not currently used)
- `assets/favicon-32.png`, `icon-192.png`, `apple-touch-icon.png` — heights.edu site icons https://heights.edu/app/uploads/2019/10/cropped-Admiss6-{32x32,192x192,180x180}.png (unaltered)
- Fonts on heights.edu: Adobe Garamond SC (self-hosted) and Typekit `tablet-gothic` / `adobe-garamond-pro` (licensed). This site uses free EB Garamond (Google Fonts) in small caps as a stand-in, plus system sans-serif.
