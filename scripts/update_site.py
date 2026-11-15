#!/usr/bin/env python3
"""Run by GitHub Actions (.github/workflows/update-site.yml) on every change, every morning, and on demand.

1. Optionally downloads the Google Sheet "Heights MS XC Results" as .xlsx (works only if the sheet is shared as
   "Anyone with the link - Viewer"); a download that isn't a valid workbook is ignored.
2. Rebuilds the site data (scripts/build_data.py).
3. If the new sheet breaks the build, puts the previous sheet back, rebuilds, and reports the problem.
4. Writes a plain-English summary for the Actions page and tells the workflow whether to commit.

Outputs (for the workflow): commit=true|false, ok=true|false.
Can also be run by hand: python3 scripts/update_site.py [--no-sheet]
"""
import hashlib, io, os, subprocess, sys, urllib.request, zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SHEET_DIR = os.path.join(ROOT, "data", "ms-sheet")
SHEET_FILE = "Heights MS XC Results.xlsx"
# The team's Google Sheet "Heights MS XC Results" (from its web address: docs.google.com/spreadsheets/d/<this>/edit)
SHEET_ID = os.environ.get("MS_SHEET_ID", "1pBwR0G2JnCTRZnEclSB361fIp0pTEarX0lJjakd9jvU")
EXPORT_URL = os.environ.get("MS_SHEET_URL") or f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"

summary = []


def say(line=""):
    print(line, flush=True)
    summary.append(line)


def finish(commit, ok):
    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a") as f:
            f.write(f"commit={'true' if commit else 'false'}\nok={'true' if ok else 'false'}\n")
    path = os.environ.get("GITHUB_STEP_SUMMARY")
    if path:
        with open(path, "a", encoding="utf-8") as f:
            f.write("\n".join(summary) + "\n")
    sys.exit(0 if os.environ.get("GITHUB_OUTPUT") or ok else 1)


def workbook_fingerprint(data):
    """Hash of every cell (values and formulas) so a re-download of an unchanged sheet counts as unchanged."""
    import openpyxl
    h = hashlib.sha256()
    for data_only in (True, False):
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=data_only)
        for ws in wb:
            h.update(("\x00TAB " + ws.title).encode())
            for row in ws.iter_rows(values_only=True):
                h.update(repr(row).encode())
    return h.hexdigest(), len(wb.sheetnames)


def current_sheet_path():
    files = sorted(f for f in os.listdir(SHEET_DIR) if f.lower().endswith(".xlsx")) if os.path.isdir(SHEET_DIR) else []
    if not files:
        return os.path.join(SHEET_DIR, SHEET_FILE)
    sys.path.insert(0, os.path.join(ROOT, "scripts"))
    import build_data
    return build_data.find_ms_sheet()


def sync_sheet():
    """-> (path, previous bytes) if the sheet changed, else (None, None)."""
    try:
        with urllib.request.urlopen(urllib.request.Request(EXPORT_URL, headers={"User-Agent": "heights-xc-site"}), timeout=90) as r:
            data = r.read()
    except Exception as ex:  # 401/403 when the sheet isn't shared publicly, or a network hiccup
        say(f"- Google Sheet: not downloaded ({getattr(ex, 'code', '') or type(ex).__name__}). If you want automatic "
            "updates from the sheet, share it as **Anyone with the link → Viewer** (see HOW-TO-UPDATE.md). "
            "Using the copy already in `data/ms-sheet/`.")
        return None, None
    if not data.startswith(b"PK"):
        say("- Google Sheet: the download wasn't a spreadsheet (the sheet is probably not shared as "
            "**Anyone with the link → Viewer**). Using the copy already in `data/ms-sheet/`.")
        return None, None
    try:
        new_fp, ntabs = workbook_fingerprint(data)
    except (zipfile.BadZipFile, Exception) as ex:
        say(f"- Google Sheet: downloaded file couldn't be opened ({type(ex).__name__}); using the copy already in `data/ms-sheet/`.")
        return None, None
    path = current_sheet_path()
    old = open(path, "rb").read() if os.path.exists(path) else None
    if old is not None:
        try:
            if workbook_fingerprint(old)[0] == new_fp:
                say("- Google Sheet: no changes since the last update.")
                return None, None
        except Exception:
            pass
    os.makedirs(SHEET_DIR, exist_ok=True)
    with open(path, "wb") as f:
        f.write(data)
    say(f"- Google Sheet: **changed** ({ntabs} tabs) — downloaded a fresh copy.")
    return path, old


def build():
    p = subprocess.run([sys.executable, os.path.join(ROOT, "scripts", "build_data.py")], cwd=ROOT, capture_output=True, text=True)
    sys.stdout.write(p.stdout)
    sys.stderr.write(p.stderr)
    return p.returncode, p.stdout, p.stderr


def headline(stdout):
    keep = [l for l in stdout.splitlines() if l.startswith(("Meet pages:", "MS:", "HS:", "Added from race-results/", "note:"))]
    return ["- " + l for l in keep]


def problem_text(stdout, stderr):
    lines = stderr.strip().splitlines()
    if "PROBLEMS" in stderr:
        lines = stderr[stderr.index("PROBLEMS"):].strip().splitlines()
    elif lines and "Traceback" in stderr:
        lines = [l for l in lines if l.strip()][-3:]
    return "\n".join(lines[-40:]) or "(no details; open the build step's log)"


def main():
    event = os.environ.get("GITHUB_EVENT_NAME", "manual")
    use_sheet = "--no-sheet" not in sys.argv and os.environ.get("SYNC_SHEET", "1") != "0"
    say("## Heights XC site update")
    changed_path, old_bytes = sync_sheet() if use_sheet else (None, None)
    if event == "schedule" and not changed_path:
        say("- Nothing new to build this morning.")
        finish(commit=False, ok=True)

    code, out, err = build()
    sheet_rejected = None
    if code not in (0, 3) and changed_path and old_bytes is not None:
        sheet_rejected = problem_text(out, err)
        with open(changed_path, "wb") as f:
            f.write(old_bytes)
        say("- The new copy of the Google Sheet caused an error, so the previous copy was used instead.")
        code, out, err = build()

    if code in (0, 3):
        say("")
        say("✅ **The website data was rebuilt.**" if code == 0 and not sheet_rejected else
            "⚠️ **The website was updated, but something was left out — see below.**")
        summary.extend(headline(out))
        if code == 3:
            say("")
            say("### Files that were NOT added (fix them and upload again)")
            say("```")
            say(problem_text(out, err))
            say("```")
        if sheet_rejected:
            say("")
            say("### The Google Sheet has a problem (its new changes are not on the site yet)")
            say("```")
            say(sheet_rejected)
            say("```")
        finish(commit=True, ok=(code == 0 and not sheet_rejected))

    say("")
    say("❌ **The website was NOT changed** (it still shows the last good version). The problem:")
    say("```")
    say(problem_text(out, err))
    say("```")
    say("Fix the file named above (or undo your last change from its History page), and the site will rebuild by itself.")
    finish(commit=False, ok=False)


if __name__ == "__main__":
    main()
