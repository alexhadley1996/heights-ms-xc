# How to update the Heights XC website

You only need the **GitHub website** and **Google Sheets**. Nothing to install.

Every time you change a file on GitHub, the website rebuilds itself. It's live about **3 minutes** later.
If a file has a mistake, the website stays as it was, and GitHub tells you what to fix (see [part 4](#4-if-something-goes-wrong)).

The project's page on GitHub: **https://github.com/alexhadley1996/heights-ms-xc** (sign in first).
The website: **https://alexhadley1996.github.io/heights-ms-xc/**

---

## 1. Add a new race (this season)

Each race is **one small file** in the **`race-results`** folder. (Varsity and JV are two races, so two files.)

### One-time setup: make a "Race entry" Google Sheet (5 minutes)

1. On GitHub, open the **`race-results`** folder, then click **`_TEMPLATE.csv`**.
2. Click the **Download raw file** button (the down-arrow icon above the file, on the right).
3. In Google Sheets, create a **Blank spreadsheet** and name it **Race entry**.
4. Click **File → Import → Upload**, pick the file you downloaded, choose **Replace current sheet**, then click **Import data**.
5. Click the column **E** letter (the Time column), then **Format → Number → Plain text**. This stops Sheets from changing times like 17:42 into clock times.

You'll keep this sheet and add one tab per race.

### Each race

1. In your **Race entry** sheet, right-click the tab at the bottom and click **Duplicate**. Rename the new tab after the race, e.g. **2027-09-18 Bullis Varsity**.
2. Fill in the top rows (column B):

   | Row | What to write | Example |
   |---|---|---|
   | Meet | The meet's name | Bullis Invitational |
   | Date | The race day | 2027-09-18 (or Sep 18, 2027) |
   | Venue | Where it was run | Bullis School |
   | Level | **MS** (middle school) or **HS** (high school) | HS |
   | Race | Varsity, JV, Freshman or Open (anything extra is fine, e.g. "JV 2A") | Varsity |
   | Distance | e.g. 5K, 3 miles, 2 mi, 3000m | 5K |
   | Source | Link to the results page (optional) | https://www.athletic.net/... |
   | Notes | Anything worth saying on the race page (optional) | Muddy course |

3. Under the **Place, Name, Grade, School, Time** row, add one row per runner.
   - **Best: copy the whole race from the results website** (every school), so the site can work out team scores. Open the race on **Athletic.net** (meet → Results → the race), **mocorunning.com** or **racecast.co**, select the results table with your mouse, copy it, and paste it under the **Place** row. Then check each column sits under the right heading. Extra columns such as Pace or Points are fine. You can even paste the site's own heading row in place of ours; it understands Athlete, Team and Yr.
   - **Short on time?** Type only our runners. Leave **School** empty for them.
   - **Time** like `17:42.3` or `17:42`. Put **DNF** (and no place) for a runner who didn't finish.
   - **Grade** is a number (9–12 for HS, 5–8 for MS). It can be empty.
   - Our school must be written **The Heights** (or left empty).
4. Click **File → Download → Comma Separated Values (.csv)**. This downloads only the tab you're on.
5. On GitHub, open the **`race-results`** folder.
6. Click **Add file → Upload files**.
7. Drag the downloaded file onto the page (or click **choose your files**).
8. Click the green **Commit changes** button.
9. Wait about 3 minutes, then open the website. The race shows on the **Meets** page and on each runner's page. Races from the same meet (same Meet name and year) share one meet page.

**Shortcut (no download):** in Sheets, select everything on the tab and copy it. On GitHub, open **`race-results`**, click **Add file → Create new file**, paste, and type a name ending in **.csv** (e.g. `2027-09-18-bullis-varsity.csv`). Then click **Commit changes…** and **Commit changes**.

### Middle school races in the team sheet ("Heights MS XC Results")

The middle school races come from the team's Google Sheet **Heights MS XC Results**. Pick **one** of these:

- **Automatic (recommended):** share that sheet so the website can read it. Open the sheet, click **Share**, and under **General access** change **Restricted** to **Anyone with the link**, with role **Viewer**. Click **Done**. After that, the site checks the sheet **every morning** and picks up new tabs and fixes by itself. Name new tabs like the old ones, e.g. **2027 - Bullis**. (The sheet then works like the website: anyone with the link can view it, but only people you invite can edit it.)
- **By hand:** in the sheet, click **File → Download → Microsoft Excel (.xlsx)**. On GitHub, open **`data`** → **`ms-sheet`**, click **Add file → Upload files**, drop the file in, and click **Commit changes**. It replaces the old copy, as long as the file is still named `Heights MS XC Results.xlsx`.

### Tabs the site should ignore

Some tabs are scratch work (a team-score table, or a copy of a race with some grades left out), not a race of their own. Open **data** → **ms-sheet** → **skip-tabs.csv**, click the pencil, add a line with the exact tab name, and commit. Tabs named **Sheet** plus a number, or containing **no 8th** or **team score**, are skipped automatically.

Put each middle school race in **either** the team sheet **or** a `race-results` file, not both, or it will show up twice.

---

## 2. Add older results

Use exactly the same steps as part 1. The **Date** decides which season the race goes in.

- **Adding a race to a meet that's already on the site?** Copy the **Meet** name exactly as the website shows it (same spelling, without the year). The new race then joins that meet's page.
- Put where you found it in **Source** (a link). If it came from a paper printout, say so in **Notes**.
- If the race is already on the site (same meet, year and race), the site won't add it again. It tells you instead.
- Old results often only list our runners. That's fine: type only ours.

---

## 3. Fix a misspelled name, or merge two runners into one

All name fixes are in one list: **`data/name-fixes.csv`**.

1. On GitHub, open **`data`** → **`name-fixes.csv`**.
2. Click the **pencil** button (**Edit this file**).
3. Go to the very end and add one line:

   `Name as it was listed,Correct name,Level,Note`

   For example: `Jon Smith,John Smith,HS,typo in 2027 Bullis results`
   - **Level:** **HS**, **MS**, or **Both** (use Both when the same boy is spelled differently in middle school and high school).
   - **Note** is just for you. It can be empty.
4. Click **Commit changes…**, then **Commit changes**.

This is also how you **merge two runner pages** for the same boy. Put the name you want to get rid of first and the name to keep second.

A boy who ran middle school and high school under **exactly the same name** is joined into one runner page automatically.

To **undo** a fix, edit the file again and delete that line.

---

## 4. If something goes wrong

**Check whether the update worked**

1. On GitHub, click the **Actions** tab (at the top of the project page).
2. Click **Update site** on the left. The newest run is at the top:
   - **Green check**: it worked. The website updates within a few minutes.
   - **Yellow dot**: still running. Wait a minute.
   - **Red X**: something needs fixing. Click the run. Its **Summary** page says, in plain words, which file and which row has the problem (e.g. "row 14 (John Smith): can't read the time '1742'").

**Fix it**

- **A race file has a mistake:** fix it in your Race entry sheet, download it again, and upload it again with the **same file name** (it replaces the old one). Or open the file on GitHub and click the **pencil** to fix it right there. The other races still go live in the meantime; only the broken file is left out.
- **Remove a race:** open its file in **`race-results`**, click the **⋯** button (top right), click **Delete file**, then click **Commit changes**.
- **Undo any change:** open the file, click **History**, and find the version from before the change. Click **⋯ → View file**, then click **Copy raw file**. Go back to the current file, click the **pencil**, select everything, paste, and click **Commit changes**.

**Other problems**

- **The website looks old after a green check:** wait 5 minutes, then refresh the page (on a phone, close the tab and reopen it). You can also force an update: **Actions → Update site → Run workflow → Run workflow**.
- **A yellow banner in Actions says the scheduled workflow is disabled:** GitHub pauses the morning check after 60 days with no changes. Click **Enable workflow**.
- **The red X says it couldn't save (permission):** go to **Settings → Actions → General**. Under **Workflow permissions**, choose **Read and write permissions** and click **Save**. Then **Actions → Update site → Run workflow**.
- **The Summary says the Google Sheet has a problem:** the website keeps the last good copy of the sheet. Fix the tab it names in **Heights MS XC Results**, and the next morning's check (or **Run workflow**) picks it up.

Everything is saved in the project's history. Nothing you do here can permanently break the website.
