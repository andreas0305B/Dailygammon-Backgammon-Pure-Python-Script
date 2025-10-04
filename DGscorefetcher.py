"""
DailyGammon Score Synchronizer
------------------------------

This script synchronizes tournament match results between an Excel results table
and DailyGammon (DG). It automates the process of filling in missing match IDs,
fetching match results, and updating scores into the correct table cells.

This script processes match results for a specific league and writes them to Excel. 
This script can also run across multiple leagues when used together with the wrapper script.


Usage:
    - Manual mode (default):
        Simply run the script without arguments. 
        Example: python dailygammon.py
        -> Uses default league "4d" hardwired in the script
        -> Keeps Excel workbook open for manual review

    - Command line / wrapper mode:
        Provide the league as the first argument and optionally '--auto' as the second.
        Example: python dailygammon.py 2b --auto
        -> Processes league "2b"
        -> Closes Excel workbook automatically (needed when running multiple leagues in sequence)

This makes it possible to run the script across multiple leagues 
without changing the source code manually.


Core Concepts:
--------------

1. Match ID Handling
   - If a match_id cell is empty, the script searches DG for invitations
     initiated by the "player" and inserts the found ID into Excel.
   - If a match_id cell is filled, it may be either:
        a) Automatically inserted earlier (normal case)
        b) Manually entered by a moderator (manual ID)
   - Manual IDs are detected by comparing player/opponent order between Excel
     and DG: if reversed, the entry is considered manual.

2. Manual Match IDs
   - Stored separately in `matches_by_hand`.
   - They carry a `switched=True` flag, meaning that for DG lookups the roles
     of player and opponent must be swapped to retrieve results.
   - When writing scores back to Excel, the swapped results are re-switched
     so the table remains consistent from the perspective of the Excel player.

3. Caching
   - Each match_id is requested from DG at most once.
   - A simple dict (`html_cache`) maps { match_id -> html } to reduce load.

4. Idempotence
   - Running the script multiple times does not duplicate work.
   - IDs are inserted only if cells are empty; scores are written only if
     the cell does not already contain a final result (e.g., "11").

5. Score Writing
   - For each resolved match, the correct Excel row and columns are located
     via player/opponent name mapping.
   - Exact (case-insensitive) name matches are preferred.
   - If no exact match is found, a heuristic rule is applied:
       * Check whether one name appears as a substring of the other.
   - If the heuristic is inconclusive, the match is skipped for safety.

6. Safety Rules
   - The script never overwrites an existing score of 11.
   - If names cannot be reliably mapped, the match is skipped instead of
     risking a wrong write.

"""


# ============================================================
# Script Purpose:
# This script automatically updates match results for a DailyGammon league season.
# It connects to DailyGammon with your login credentials, collects all match IDs,
# downloads intermediate/final scores, and writes them into the Excel results file.
#
# Workflow in summary:
#   1. Login to DailyGammon with your credentials
#   2. Read the player list from the Excel "Players" sheet
#   3. Detect already known matches from the "Links" sheet
#   4. Find and insert missing match IDs automatically
#   5. Update "Matches" sheet with intermediate results
#   6. For finished matches, set the final winner score to 11
#
# Excel file requirement:
# - Requires Excel file "<season>th_Backgammon-championships_<league>.xlsm"
#   The corresponding Excel file (e.g. "34th_Backgammon-championships_4d.xlsm")
#   must be located in the same folder as this script.
#
# - Excel sheets used:
#       * "Players" → base player list
#       * "Links"   → references to match IDs
#       * "Matches" → current scores
# - Important: Scores are only updated if the match is not yet marked as finished (11).
#
# Before running, configure:
#   - Your User ID and Password (variables: payload["login"], payload["password"])
#   - Current Season number (variable: saison_nummer, e.g. "34")
#   - League (variable: liga, e.g. "4d")
#
# Required Python libraries:
#   requests, beautifulsoup4, xlwings
#
# If not installed, run:
#   pip install requests beautifulsoup4 xlwings
#
# ============================================================

import requests
from bs4 import BeautifulSoup
import openpyxl
import re
import xlwings as xw

# --- Login Data ---
load_dotenv(dotenv_path="a.env")  # lokale .env laden
login_url = "http://dailygammon.com/bg/login"

# Zuerst versuchen, aus Streamlit-Secrets zu laden, sonst .env / Umgebungsvariablen
try:
    DG_LOGIN = st.secrets["dailygammon"]["login"]
    DG_PW = st.secrets["dailygammon"]["password"]
except Exception:
    DG_LOGIN = os.getenv("DG_LOGIN", "")
    DG_PW = os.getenv("DG_PW", "")

# Prüfen, ob Login-Daten vorhanden sind
if not DG_LOGIN or not DG_PW:
    st.error("❌ Keine Login-Daten gefunden. Bitte in a.env oder .streamlit/secrets.toml eintragen.")
    st.stop()

# Debug-Ausgabe (Login maskieren, Passwort nicht ausgeben!)
masked_login = DG_LOGIN[0] + "*" * (len(DG_LOGIN) - 2) + DG_LOGIN[-1] if len(DG_LOGIN) > 2 else DG_LOGIN

payload = {
    "login": DG_LOGIN,
    "password": DG_PW,
    "save": "1"
}

BASE_URL = "http://dailygammon.com/bg/game/{}/0/list"
saison_nummer = "34"                                                      # Current Saison input

import sys

if len(sys.argv) > 1:
    liga = sys.argv[1] 
else:
    liga = "4d"                                                           # Current Liga, if not started by wrapper script

# detect if script was called from wrapper with '--auto'
AUTO_MODE = "--auto" in sys.argv


file = f"{saison_nummer}th_Backgammon-championships_{liga}.xlsm"
season = f"{saison_nummer}th-season-{liga}"

print("="*50)
print(f"▶ Script started – collecting links and data for {season}")
print(f"📂 Results saved in Excel file: {file}")
print("="*50)

# -----------------------------------------------------
# --- Read players from Excel ---
# -----------------------------------------------------
# EXCEL FILE ACCESS (OPENPYXL READ-ONLY PHASE):
# - We first use openpyxl to read the "Players" sheet without opening Excel.
# - The file must exist in the same folder and contain a proper "Players" sheet.
# - Player IDs are extracted from hyperlinks in the first column (format: .../bg/user/<id>).
# - If hyperlinks are missing, that player won't get an ID and will be skipped later.
# -----------------------------------------------------


wb_meta = openpyxl.load_workbook(file, data_only=True)
ws_players = wb_meta["Players"]

players = []
player_ids = {}
for row in ws_players.iter_rows(min_row=2, max_col=1, values_only=False):
    cell = row[0]
    if cell.value:
        name = str(cell.value).strip()
        players.append(name)
# - The script assumes each player cell may contain a hyperlink to their DailyGammon user page.
        if cell.hyperlink:
            url = cell.hyperlink.target
            player_id = url.rsplit("/", 1)[-1]
            player_ids[name] = player_id

wb_meta.close()

# -----------------------------------------------------
# --- Login session ---
# -----------------------------------------------------
# -----------------------------------------------------
# Function: login_session
# Purpose:
#   Opens a persistent HTTP session with DailyGammon,
#   logs in with your credentials, and returns the session
#   so all following requests are authenticated.
# -----------------------------------------------------


def login_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0"})
    resp = s.post(login_url, data=payload, timeout=30)
    resp.raise_for_status()
    return s

session = login_session()

# -----------------------------------------------------
# --- Collect matches per player ---
# -----------------------------------------------------
# -----------------------------------------------------
# Function: get_player_matches
# Purpose:
#   Collects all matches for a specific player in the
#   given season. It scrapes the DailyGammon user page
#   and extracts:
#     - Opponent name
#     - Opponent ID
#     - Match ID
#
# - Filters table rows by the 'season' string to avoid pulling old matches.
# -----------------------------------------------------

def get_player_matches(session: requests.Session, player_id, season):
    url = f"http://www.dailygammon.com/bg/user/{player_id}"
    r = session.get(url)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    player_matches = []
    for row in soup.find_all("tr"):
        text = row.get_text(" ", strip=True)
        if season not in text:
            continue
        opponent_link = row.find("a", href=re.compile(r"/bg/user/\d+"))
        match_link = row.find("a", href=re.compile(r"/bg/game/\d+/0/"))
        if not opponent_link or not match_link:
            continue
        opponent_name = opponent_link.text.strip()
        opponent_id = re.search(r"/bg/user/(\d+)", opponent_link["href"]).group(1)
        match_id = re.search(r"/bg/game/(\d+)/0/", match_link["href"]).group(1)
        player_matches.append((opponent_name, opponent_id, match_id))
    return player_matches

# -----------------------------------------------------
# --- Helper functions: fetch HTML & extract scores ---
# -----------------------------------------------------
# -----------------------------------------------------
# Function: fetch_list_html
# Purpose:
#   Downloads the HTML page for a specific match ID.
#   Returns the HTML content or None if the request failed.
# -----------------------------------------------------

def fetch_list_html(session: requests.Session, match_id: int) -> str | None:
    url = BASE_URL.format(match_id)
    try:
        resp = session.get(url, timeout=30)
        if not resp.ok or "Please Login" in resp.text:
            return None
        return resp.text
    except requests.RequestException:
        return None

# -----------------------------------------------------
# Function: extract_latest_score
# Purpose:
#   Parses the match HTML page and extracts the latest
#   visible score row for the two players.
#   Returns player names + current scores.
#
# PARSE LATEST SCORE FROM MATCH PAGE:
# - Scans table rows from bottom to top (reversed) to find the most recent score line.
# - Assumes the pattern "<Name> : <Score>" is present on both left and right columns.
# -----------------------------------------------------

def extract_latest_score(html: str, players_list: list[str]):
    soup = BeautifulSoup(html, "html.parser")
    for row in reversed(soup.find_all("tr")):
        text = row.get_text(" ", strip=True)
        if not any(p in text for p in players_list):
            continue
        cells = row.find_all("td")
        if len(cells) >= 3:
            left_text = cells[1].get_text(" ", strip=True)
            right_text = cells[2].get_text(" ", strip=True)
            left_match = re.match(r"(.+?)\s*:\s*(\d+)", left_text)
            right_match = re.match(r"(.+?)\s*:\s*(\d+)", right_text)
            if left_match and right_match:
                left_name, left_score = left_match.groups()
                right_name, right_score = right_match.groups()
                return left_name.strip(), right_name.strip(), int(left_score), int(right_score)
    return None

# -----------------------------------------------------
# Function: map_scores_for_excel
# Purpose:
#   Aligns scores from DailyGammon with the correct order
#   in the Excel sheet.
#   Handles switched cases (player order reversed for manual added matches).
#
# NAME/SCORE ALIGNMENT TO EXCEL:
# - The Excel grid expects "excel_player" vs "excel_opponent" in a fixed orientation.
# - 'switched_flag=True' means the match was manually entered with reversed order
#   (excel_player appears on the right on DailyGammon), so we swap scores here.
# - If names match exactly (case-insensitive), we map directly; otherwise we use a
#   small heuristic (substring check) as a fallback. If unsure, return None (skip write).
# -----------------------------------------------------


def map_scores_for_excel(player, opponent, left_name, right_name, left_score, right_score, switched_flag):
    ln = left_name.strip().lower()
    rn = right_name.strip().lower()
    pn = player.strip().lower()
    on = opponent.strip().lower()

    if switched_flag:
        return right_score, left_score

    if ln == pn and rn == on:
        return left_score, right_score
    if ln == on and rn == pn:
        return right_score, left_score

    # Fallback heuristic if names differ slightly
    if pn in ln or pn in rn or on in ln or on in rn:
        if pn in ln:
            return left_score, right_score
        if pn in rn:
            return right_score, left_score
    return None

# -----------------------------------------------------
# --- Open Excel workbook via xlwings ---
# -----------------------------------------------------
# EXCEL WRITING PHASE (XLWINGS):
# - From this point on, we interact with Excel via xlwings (live Excel instance).
# -----------------------------------------------------


wb_xw = xw.Book(file)
ws_links = wb_xw.sheets["Links"]
ws_matches = wb_xw.sheets["Matches"]

# Extract players/columns from "Links"
# "LINKS" SHEET LAYOUT ASSUMPTION:
# - Column A (from row 2 down) lists row player names.
# - Row 1 (from column B rightwards) lists opponent names (as columns).
# - Cells at (row_player, col_opponent) hold the match ID (and hyperlink).

row_players_links = []
r = 2
while (v := ws_links.range(f"A{r}").value):
    row_players_links.append(str(v).strip())
    r += 1

col_opponents_links = []
c = 2
while (v := ws_links.range(1, c).value):
    col_opponents_links.append(str(v).strip())
    c += 1

col_index_links = {name: 2 + i for i, name in enumerate(col_opponents_links)}

# -----------------------------------------------------
# --- Data structures ---
# -----------------------------------------------------
matches = {}
matches_by_hand = {}
match_id_to_excel = {}
html_cache = {}
finished_by_id = {}

# -----------------------------------------------------
# Step 1: Check existing links
# Purpose:
#   Go through the "Links" sheet and verify which matches
#   already have a match ID entered. If the IDs are present,
#   confirm whether the player/opponent order is correct.
#   Marks switched matches if detected.
# -----------------------------------------------------

for i, player_name in enumerate(row_players_links, start=2):
    for opp in col_opponents_links:
        if player_name == opp:
            continue
        c = col_index_links.get(opp)
        val = ws_links.range(i, c).value
        if not val:
            continue
        try:
            match_id = int(val)
        except Exception:
            match_id = int(str(val).strip())
        if match_id not in html_cache:
            html_cache[match_id] = fetch_list_html(session, match_id)
        html = html_cache[match_id]
        if not html:
            matches[(player_name, opp)] = match_id
            match_id_to_excel[match_id] = (player_name, opp, False)
            continue
        score_info = extract_latest_score(html, [player_name, opp])
        if not score_info:
            matches[(player_name, opp)] = match_id
            match_id_to_excel[match_id] = (player_name, opp, False)
            continue
        left_name, right_name, _, _ = score_info
        ln = left_name.lower(); rn = right_name.lower()
        pn = player_name.lower(); on = opp.lower()
        if ln == pn and rn == on:
            matches[(player_name, opp)] = match_id
            match_id_to_excel[match_id] = (player_name, opp, False)

# MANUAL/SWITCHED CASE:
# - DailyGammon lists "opponent vs player", but Excel expects "player vs opponent".
# - We record 'switched=True' for this match_id so all later writes swap correctly.

        elif ln == on and rn == pn:
            matches_by_hand[(player_name, opp)] = (match_id, True)
            match_id_to_excel[match_id] = (player_name, opp, True)
            print(f"Found manual inserted match detected: {player_name} vs {opp} with match ID {match_id}.")
        else:
            matches[(player_name, opp)] = match_id
            match_id_to_excel[match_id] = (player_name, opp, False)
            print(f"⚠️ Unclear order for match ID {match_id}: DG shows '{left_name}' vs '{right_name}'")

# -----------------------------------------------------
# Step 2: Fill missing match IDs
# Purpose:
#   For each player, check which opponents still have no
#   match ID in the "Links" sheet. Search for the match on
#   DailyGammon and insert it automatically into the table.
#   Also detects "switched" matches (player/opponent reversed).
#
# STEP 2 RATIONALE:
# - For any missing (player, opponent) cell, we look up the player's page to find
#   their active matches for this season and backfill the match ID into "Links".
# - We also attach a hyperlink to the specific match list page for quick access.
# - Existing cells are left untouched; only empty cells get filled.
# -----------------------------------------------------


for player in players:
    pid = player_ids.get(player)
    if not pid:
        continue
    missing = [opp for opp in players if opp != player and (player, opp) not in matches and (player, opp) not in matches_by_hand]
    if not missing:
        continue
    player_matches = get_player_matches(session, pid, season=season)
    for opponent_name, opponent_id, match_id in player_matches:
        key = (player, opponent_name)
        if key in matches or key in matches_by_hand:
            continue
        mid_int = int(match_id)
        switched_flag = False
        if mid_int in match_id_to_excel:
            _, _, switched_flag = match_id_to_excel[mid_int]
        matches[key] = mid_int
        match_id_to_excel[mid_int] = (player, opponent_name, switched_flag)
        try:
            row_idx = row_players_links.index(player) + 2
        except ValueError:
            continue
        c = col_index_links.get(opponent_name)
        if not c or opponent_name == player:
            continue
        cell = ws_links.range(row_idx, c)

# - We only write if the cell is empty to avoid overwriting manual adjustments.
# - If you ever need to refresh a wrong ID, clear the cell first, then rerun.

        if not cell.value:
            cell.value = str(match_id)
            cell.api.Hyperlinks.Add(
                Anchor=cell.api,
                Address=f"http://www.dailygammon.com/bg/game/{match_id}/0/list#end",
                TextToDisplay=str(match_id)
            )
            print(f"Detected missing match between {player} and {opponent_name} — match ID={match_id} has been auto-added to the table")

wb_xw.save()
print("✅ Match IDs updated (auto + manual detection)")

# -----------------------------------------------------
# Step 3: Collect finished matches
# Purpose:
#   For every player, fetch their export page.
#   If a match is marked as finished, extract the winner.
#   Results are stored in a dictionary for later processing.
#
# FINISHED MATCH DETECTION:
# - We open each player's page and follow "export" links for matches of this season.
# - The winner is inferred from a simple textual rule (position of "Wins" on the line).
# - 'finished_by_id' maps match_id -> winner_name for later use in Phase 2.
# -----------------------------------------------------

for player in players:
    pid = player_ids.get(player)
    if not pid:
        continue
    url = f"http://www.dailygammon.com/bg/user/{pid}"
    try:
        r = session.get(url, timeout=30)
        r.raise_for_status()
    except requests.RequestException:
        continue
    soup = BeautifulSoup(r.text, "html.parser")
    for row in soup.find_all("tr"):
        text = row.get_text(" ", strip=True)
        if season not in text:
            continue
        export_link = row.find("a", href=re.compile(r"/bg/export/\d+"))
        match_link = row.find("a", href=re.compile(r"/bg/game/\d+/0/"))
        opponent_link = row.find("a", href=re.compile(r"/bg/user/\d+"))
        if not export_link or not match_link or not opponent_link:
            continue
        try:
            match_id = int(re.search(r"/bg/game/(\d+)/0/", match_link["href"]).group(1))
        except Exception:
            continue
        opponent_name = opponent_link.text.strip()
        export_url = f"http://www.dailygammon.com/bg/export/{match_id}"
        try:
            resp_export = session.get(export_url, timeout=30)
            text_lines = resp_export.text.splitlines()
        except requests.RequestException:
            continue
        winner = None

# - 'mid_threshold 24' is a rough character-position cutoff to decide whether the "Wins"
#   belongs to the left or right player on the export line.

        mid_threshold = 24
        for line in text_lines:
            if "and the match" in line and "Wins" in line:
                pos = line.find("Wins")
                winner = player if pos < mid_threshold else opponent_name
                break
        if winner:
            finished_by_id[match_id] = winner

# -----------------------------------------------------
# Phase 1: Write intermediate scores
# Purpose:
#   For each match, download the latest score and update
#   the "Matches" sheet in Excel.
#   IMPORTANT: If a score of 11 is already present,
#   the match is considered finished and will not be overwritten.
# -----------------------------------------------------

print("🔎 Phase 1: Writing intermediate scores for matches...")
players_in_matches = []
row_counter = 4
while True:
    nm = ws_matches.range((row_counter, 1)).value
    if not nm:
        break
    players_in_matches.append(str(nm).strip())
    row_counter += 1
col_start = 2

# EXCEL WRITE HELPER (INTERMEDIATE SCORES):
# - Translates (excel_player, excel_opponent) to row/column indices in "Matches".
# - For each opponent, we reserve two columns: left=excel_player's score, right=excel_opponent's score.
# - Safety: if either cell already equals 11, we skip to preserve final results.
# - Scores are already correctly oriented by 'map_scores_for_excel'; no swapping here.

def write_score_to_excel(excel_player, excel_opponent, player_score, opponent_score, switched_flag):
    try:
        r_idx = players_in_matches.index(excel_player) + 4
        c_base = players_in_matches.index(excel_opponent)
    except ValueError:
        print(f"⚠️ Player not found in Excel sheet: {excel_player} vs {excel_opponent}")
        return False
    c_left = col_start + c_base * 2
    c_right = c_left + 1

    # Do not overwrite already finished (11) scores!
    # - Once a match is finished (11), intermediate updates must never overwrite that cell.

    left_cell_val = ws_matches.range((r_idx, c_left)).value
    right_cell_val = ws_matches.range((r_idx, c_right)).value
    if left_cell_val == 11 or right_cell_val == 11:
        return False

    ws_matches.range((r_idx, c_left)).value = player_score
    ws_matches.range((r_idx, c_right)).value = opponent_score
    return True

# - Pull HTML from cache if available; otherwise fetch fresh.

for match_id, (excel_player, excel_opponent, switched_flag) in list(match_id_to_excel.items()):
    html = html_cache.get(match_id)
    if not html:
        html = fetch_list_html(session, match_id)
        html_cache[match_id] = html
    if not html:
        continue
    result = extract_latest_score(html, players_in_matches)
    if not result:
        continue
    left_name, right_name, left_score, right_score = result

    # Map scores based on player names

    mapped = map_scores_for_excel(excel_player, excel_opponent, left_name, right_name, left_score, right_score, switched_flag)
    if mapped is None:
        continue
    excel_player_score, excel_opponent_score = mapped
    write_score_to_excel(excel_player, excel_opponent, excel_player_score, excel_opponent_score, switched_flag)

wb_xw.save()
print("✅ Phase 1: completed")

# -----------------------------------------------------
# Phase 2: Final results - Set winners to 11 points
# Purpose:
#   For matches identified as finished, write the final
#   winner score (11 points) into the correct player cell
#   in the "Matches" sheet.
# -----------------------------------------------------

print("🔎 Phase 2: Final results (set winner = 11) ...")
for match_id, winner_name in finished_by_id.items():
    info = match_id_to_excel.get(match_id)
    if not info:
        continue
    excel_player, excel_opponent, switched_flag = info
    try:
        r_idx = players_in_matches.index(excel_player) + 4
        c_base = players_in_matches.index(excel_opponent)
    except ValueError:
        continue
    c_left = col_start + c_base * 2
    c_right = c_left + 1

    # Write 11 to the correct winner cell

    winner_lower = winner_name.strip().lower()
    if switched_flag:
        if winner_lower == excel_player.lower():
            ws_matches.range((r_idx, c_right)).value = 11
        elif winner_lower == excel_opponent.lower():
            ws_matches.range((r_idx, c_left)).value = 11
    else:
        if winner_lower == excel_player.lower():
            ws_matches.range((r_idx, c_left)).value = 11
        elif winner_lower == excel_opponent.lower():
            ws_matches.range((r_idx, c_right)).value = 11

wb_xw.save()
# Close workbook automatically only if called from wrapper
if AUTO_MODE:
    wb_xw.close()
print("🏁 Script finished successfully")
print("="*50)



