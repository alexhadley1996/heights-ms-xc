# Heights HS XC research (2006–2026): notes

Files:
- csv/: 316 per-race CSVs. Columns: Place,Grade,Name,Time,Pace,School,Meet,Level,SourceURL. Boys only, whole field.
- ALL_FOUND_HS_RACES_summary.csv: one row per race (316).
- official_team_scores.csv: about 4,630 team rows. Sources: racecast "as scored by meet", plus mocorunning meet printouts.
- heights_hs_results_combined.csv: all Heights rows, about 3,050.
- ms_hs_matches.csv: MS to HS runner matches (confident / needs confirmation).
- hs_name_variants.csv: spelling variants of the same name within HS.

Pipeline: scrape_hs.py → build_hs.py → build_moco_hs.py → assemble.py → match.py → hs_variants.py

Data quality:
- Pace = time / listed distance, given as m:ss.s like the MS CSVs. It is blank when the distance isn't listed (some WCAC races, Holy Cross 2020).
- 2007 Episcopal Quad: the official printout (moco 435) has "Handley, Tim 15:35" (probably Timothy Hadley; the time looks wrong) and "Name, Unreadable". Racecast drops both. Official score: Heights 1st with 42. Racecast's computed score is 2nd with 70.
- Woodward Relays (moco 792): moco files it under Sep 8, 2008. The HY-TEK header says 9/11/2006. The roster (Greg Hadley in HS) points to 2008.
- 2019 Bullis Invitational: moco says Sep 7, 2019. The RaceTab header says 8/31/2019. The header date is used.
- 2010 Glory Days: on moco, the boys freshman team table sits under the "Seeded Girls" title. It was matched to Freshman by Heights finisher places.
- 2020 Holy Cross 1: the source lists Joe McClorey as grade 9. He was grade 11 in Holy Cross 2.
- WCAC "Varsity B" is labeled JV (Varsity B). The 2021 WCAC 3rd boys race has no title in the source (probably freshman).
- Racecast's Heights grade data sometimes doesn't match the moco printout. Check mismatches if needed.
- Missing places in source printouts: 2018 WCAC frosh, 2023 WCAC Varsity B and frosh, 2026 VA Showcase JV.
- RVA "Extras" races were posted in sections (e.g., M–Z), so overall places have gaps.

## Pre-2006 history (added Nov 2026)
- Batch 1 (2005 HS, 2004): 11 races from racecast + mocorunning printouts.
- Batch 2 (2003, 2002): 8 races: St. Johns Inv. 2003 and Georgetown Prep Classic 2003 (full fields), Sandy Spring 2003 and BW Private Championships 2002/2003 (partial fields; official team scores shown, not recalculated — race_notes.json "partial_scores"), and three Washington Post team-score-only meets.
- "Mike Fagan" (2002 gr 12; 2003 no grade) is kept apart from "Michael Fagan" (2004 gr 12) until the team confirms.

## History batch 3 (2001-2000), Nov 2026
- 2001: Mercersburg Invitational varsity (racecast full field; official scores from the DyeStat printout, Heights 12th/327);
  Georgetown Prep Classic Div I varsity, Div I JV and freshman (MileSplit raw HY-TEK printouts, official scores: 12th/329, 20th/582, 22nd/598).
- 2000: Georgetown Prep Classic Div I varsity (16th/372), Div I JV and freshman (no Heights team score) from the DyeStat printout via mocorunning 3492;
  team scores only: Ballenger Creek Park meet Oct 11 (Heights 1st, 26; WaPo via mocorunning 9680) and Washington Independent Schools Championship
  Nov 4 at Glenelg (4th of 13, 124; mocorunning 7625).
- Merges: AJ Guarente = Andrew Guarente (confirmed by the team). Spelling/nickname variants (same team, consecutive seasons, grades consistent):
  Dan Shusinsky→Dan Sushinsky, Danny Mccarthy/Dan MC Carthy→Dan McCarthy, Allejandro→Alejandro Lulli, Dan Henebrey→Dan Henebery (spelling in 3 of 4 sources),
  Charley→Charles Gallagher, Mike Galagher→Mike Gallagher, James Acsvedo→James Acevedo, Jim→James Streeter, Matt Fuley (2001 gr 9)→Matt Foley (2004 gr 12).
- 2001 GP Classic "Mike Fagan" (gr 9) is put on Michael Fagan's page via race_notes name_fixes (2001 Mercersburg lists "Michael Fagan"; gr 9 in 2001 = senior in 2004).
  The 2002-03 "Mike Fagan" (2002 listed gr 12) stays separate pending the team's answer.
- 2003 Baltimore-Washington small-school race: distance set to 5K by the team.
