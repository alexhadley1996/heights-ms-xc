# Heights MS Cross Country — results site (draft)

Plain HTML/CSS/JS, no build step. Works on GitHub Pages as-is (`.nojekyll` included).

Pages: `index.html` (home), `meets.html` (all races by season), `meet.html?id=…` (one race: Heights runners or full field),
`runners.html` (searchable roster), `runner.html?id=…` (one runner's races + PB), `top.html` (top times/paces), `history.html` (placeholders).

## Refreshing the data
1. In Google Sheets, File → Download → Microsoft Excel (.xlsx) for “Heights MS XC Results”.
2. `pip install openpyxl` then `python3 scripts/build_data.py path/to/file.xlsx`
3. This rewrites `data/results.js` (used by the site) and `data/results.json`.

New meet tabs are picked up automatically; add a line to `MEET_INFO` in the script to set a nicer name/date/distance,
and `NAME_ALIASES` to merge spelling variants of a runner's name.

## Preview locally
`python3 -m http.server` in this folder, then open http://localhost:8000

Colors in `assets/styles.css` (`--accent` etc.) are neutral placeholders — swap in official school colors when confirmed.
Text marked “Placeholder” (home page About, History page, footer contact) still needs to be written.

## Branding sources (checked 2026-10-30)
- Colors: school colors red and white — https://heights.edu/about/quick-facts/ ("School Color: Red"), https://en.wikipedia.org/wiki/The_Heights_School_(Maryland) and https://www.wcacsports.com/members/theheights/index ("Red / White").
- Exact red `#841313` (and darker `#6e1010`): heights.edu theme stylesheet https://heights.edu/app/themes/heights-school/dist/styles/main_feb5b753.css (primary link/button/nav color).
- Mascot / team name: Cavalier / Cavaliers — heights.edu Quick Facts, Wikipedia, WCAC.
- `assets/heights-crest.svg` — https://heights.edu/app/themes/heights-school/resources/assets/images/crest.svg (unaltered)
- `assets/heights-men-fully-alive.svg` — https://heights.edu/app/themes/heights-school/resources/assets/images/mfa.svg (unaltered, not currently used)
- `assets/favicon-32.png`, `icon-192.png`, `apple-touch-icon.png` — heights.edu site icons https://heights.edu/app/uploads/2019/10/cropped-Admiss6-{32x32,192x192,180x180}.png (unaltered)
- Fonts on heights.edu: Adobe Garamond SC (self-hosted) and Typekit `tablet-gothic` / `adobe-garamond-pro` (licensed). This site uses free EB Garamond (Google Fonts) in small caps as a stand-in, plus system sans-serif.
