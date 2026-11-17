# How to update the Heights XC website

You only need the **GitHub website** and **Google Sheets**. Nothing to install.

The results live in one Google Sheet, **Heights XC Results**. Each race is its own tab. The website reads that sheet **every morning at about 5:17 am Eastern** (6:17 am while daylight saving time is in effect). It is live a few minutes later. To publish sooner, use **Run workflow** (step 4 below).

The project's page on GitHub: **https://github.com/alexhadley1996/heights-ms-xc** (sign in first).
The website: **https://alexhadley1996.github.io/heights-ms-xc/**

The sheet must be shared as **Anyone with the link → Viewer**. Open the sheet, click **Share**, and under **General access** change **Restricted** to **Anyone with the link**, with role **Viewer**. Only people you invite can edit it.

---

## 1. Add a race

1. Open the **Heights XC Results** sheet.
2. Right-click a tab at the bottom and click **Duplicate**.
3. Rename the new tab so the name includes the **date**, the **meet**, the **level**, and the **distance**. Example: **2027-09-18 Bullis Varsity 5K**.
4. Paste the results however the source shows them. The site figures out the columns. A normal table (Place, Name, Grade, School, Time, in any order) works, and so does text copied as one column from Athletic.net or a results file. Check the Actions summary for any skipped rows.
5. You can still type the headings yourself: **Place**, **Name**, **Grade**, **School**, **Time**. Extra columns such as Pace or Bib are ignored.
   - **Best:** copy the whole race (every school), so the site can work out team scores.
   - **Short on time?** Type only our runners. Leave **School** empty for them; an empty School counts as **The Heights**.
   - **Time** like `17:42.3` or `17:42`. If Sheets turns 17:42 into a clock time, that is fine. Put **DNF** for a runner who didn't finish.
   - **Grade** is a number (9–12 for high school, 5–8 for middle school). It can be empty.
6. Leave the tab there. Do not download anything.

### Tab names

The tab name is how the site knows what the race is. A few examples:

| Tab name | What it becomes |
|---|---|
| `2027-09-18 Bullis Varsity 5K` | High school Varsity, 5K, September 18, 2027 |
| `Bullis JV 5K 9/18/27` | High school JV, 5K, same day |
| `Bullis Freshman 3 miles Sep 18 2027` | High school Freshman, 3 miles |
| `9.18.2027 - Bullis Invitational - MS - 2 mi` | Middle school, 2 miles |
| `Oct 3, 2027 WCAC Varsity 5000m` | High school Varsity, 5K (5000 m) |
| `2027-09-18 Bullis Varsity 2A 5K` | High school Varsity; "2A" stays on the race label |
| `Lake Needwood 11426 5k Varsity` | High school Varsity, 5K, November 4, 2026. A compact date like `11426` means 11/4/26 (the fall reading, not January 14) |

- **Varsity**, **Var**, or a lone **V** → high school Varsity.
- **JV** → high school JV.
- **Freshman**, **Frosh**, or **Fr** → high school Freshman.
- **Open** → high school Open.
- **MS** or **Middle School** → middle school.
- Words like **2A** or **Small Schools** stay on the race name. They do not change the level.
- If the name has a date but no level, the site treats it as middle school and says so on the update summary.
- Races with the same meet name and the same day share one meet page (Varsity, JV, Freshman, and middle school together).
- A high school tab whose meet and year match a meet already on the site is added to that meet's page. If that page already has the same level, the sheet replaces that race and the old official team scores are kept.
- A compact date with no slashes is fine: `9426` is September 4, 2026, and `11426` is November 4, 2026. When the number could be two dates, the site picks the one in the cross country season (August–November) and says so on the summary.
- Dashes, dots, and slashes all work (`2027-09-18`, `9.18.2027`, `9/18/27`).
- The older middle school tabs, named like **2026 - Bullis**, keep working as they are. You do not have to rename them.

### Publish it now

You do not have to wait until morning.

1. On GitHub, click the **Actions** tab.
2. Click **Update site** on the left.
3. Click **Run workflow**, leave the branch as **main**, and click the green **Run workflow** button.
4. Wait until the run has a green check (about 3 minutes), then refresh the website.

A tab with a mistake is left out. The Summary of the run names the tab and the row. The other races still go live.

---

## 2. Fix a misspelled name, or merge two runners into one

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

## 3. Tabs the site should ignore

Some tabs are scratch work (a team-score table, or a copy of a race with some grades left out), not a race of their own.

Open **data** → **ms-sheet** → **skip-tabs.csv**, click the pencil, add a line with the exact tab name, and commit. Tabs named **Sheet** plus a number, or containing **no 8th** or **team score**, are skipped automatically.

---

## 4. If something goes wrong

**Check whether the update worked**

1. On GitHub, click the **Actions** tab (at the top of the project page).
2. Click **Update site** on the left. The newest run is at the top:
   - **Green check**: it worked. The website updates within a few minutes.
   - **Yellow dot**: still running. Wait a minute.
   - **Red X**: something needs fixing. Click the run. Its **Summary** page says, in plain words, which tab and which row has the problem (e.g. "row 14 (John Smith): can't read the time '1742'").

**Fix it**

- **A tab has a mistake:** fix it in the Google Sheet. The other races still go live; only the broken tab is left out. Then **Run workflow** again, or wait for the morning check.
- **Remove a race:** delete that tab in the sheet (or add its name to **skip-tabs.csv**).
- **Undo any change:** in the sheet, use **File → Version history**. On GitHub, open a file, click **History**, and restore the older copy if you need to.

**Other problems**

- **The website looks old after a green check:** wait 5 minutes, then refresh the page (on a phone, close the tab and reopen it). You can also force an update: **Actions → Update site → Run workflow → Run workflow**.
- **A yellow banner in Actions says the scheduled workflow is disabled:** GitHub pauses the morning check after 60 days with no changes. Click **Enable workflow**.
- **The red X says it couldn't save (permission):** go to **Settings → Actions → General**. Under **Workflow permissions**, choose **Read and write permissions** and click **Save**. Then **Actions → Update site → Run workflow**.
- **The Summary says the Google Sheet was not downloaded:** the sheet is not shared as **Anyone with the link → Viewer**. Fix the share setting. The site keeps using the copy it already has.

Everything is saved in the project's history. Nothing you do here can permanently break the website.

---

## Alternative: upload a CSV

You can still add one race as a file in the **`race-results`** folder, instead of a sheet tab. Open **`race-results/_TEMPLATE.csv`** for the layout (Meet, Date, Level, Race, Distance, then Place, Name, Grade, School, Time). Upload the file on GitHub and commit it. A sheet tab and a CSV for the same race will both try to appear, so use one or the other.
