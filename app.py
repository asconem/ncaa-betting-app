from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
import re
import logging
from logging.handlers import RotatingFileHandler
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'ncaa-ats-analysis-key'

# --- 1. DATA CLEANING & MAPPING ---
def clean_team_name(name):
    """Strips rankings and records/conferences from ESPN text"""
    if not name:
        return name
    # Removes records/conferences: e.g., (15-92-4Away) or (20-29-2Big Ten)
    name = re.sub(r'\(\d+-\d+.*?\)[\w\s]*$', '', name)
    # Removes leading rankings: e.g., '9 ' from '9 Nebraska'
    name = re.sub(r'^\d+\s+', '', name)
    return name.strip()

# Complete mapping to ensure ESPN names match TeamRankings names
TEAM_NAME_MAPPING = {
    'Long Island University': 'LIU',
    'Long Island': 'LIU',
    'Massachusetts': 'UMass',
    'Arkansas-Pine Bluff': 'AR-Pine Bluff',
    'Abilene Christian': 'Abl Christian',
    'Alabama State': 'Alabama St',
    'Appalachian State': 'App State',
    'Coastal Carolina': 'Coastal Car',
    'East Carolina': 'E Carolina',
    'Loyola Maryland': 'Loyola MD',
    'IU Indianapolis': 'IU Indy',
    'Saint Francis': 'St Francis PA',
    'Middle Tennessee': 'Middle Tenn',
    'New Hampshire': 'New Hampshire',
    'Eastern Michigan': 'E Michigan',
    'Portland State': 'Portland St',
    'Sam Houston': 'Sam Houston',
    'SF Austin': 'SF Austin'
}

# --- 2. ATS DATA PARSER (Team | Record | Cover% | MOV | ATS +/-) ---
def load_ats_data_from_text(text):
    """Parses raw ATS data without expecting headers"""
    ats_dict = {}
    for line in text.strip().split('\n'):
        line = line.strip()
        if not line: continue
        # Matches: Team Name | Record | Cover% | MOV | ATS +/-
        match = re.search(r'^(.+?)\s+(\d+-\d+-?\d*)\s+([\d.]+%)\s+([-+]?\d+\.?\d*)\s+([-+]?\d+\.?\d*)$', line)
        if match:
            team_name = match.group(1).strip()
            ats_dict[team_name] = {
                'cover_pct': match.group(3),
                'ats_pm': float(match.group(5))
            }
    return ats_dict

# --- 3. SCHEDULE PARSER (ESPN Desktop) ---
def parse_espn_schedule_from_text(text):
    """Extracts games and cleans team names from ESPN text blocks"""
    lines = [l.strip() for l in text.strip().split('\n') if l.strip()]
    games = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if re.match(r'^\d{1,2}:\d{2}\s*[AP]M$', line):
            time, away_team, home_team, spread = line, None, None, None
            j = i + 1
            while j < min(i+40, len(lines)):
                check_line = lines[j]
                # If we find a record line, the team name was on the line above it
                if re.match(r'^\(\d+-\d+.*?\)$', check_line):
                    if not away_team: away_team = clean_team_name(lines[j-1])
                    elif not home_team: home_team = clean_team_name(lines[j-1])

                if "Spread:" in check_line:
                    match = re.search(r'Spread:([A-Z0-9&\-]+)\s+([-+]?\d+\.?\d*)', check_line)
                    if match:
                        spread = {'original_abbrev': match.group(1), 'value': match.group(2),
                                 'display': f"{match.group(1)} {match.group(2)}"}
                if 'Gamecast' in check_line: break
                j += 1
            if away_team and home_team:
                games.append({'Away': away_team, 'Home': home_team, 'Time': time, 'Market': spread or 'N/A'})
            i = j
        else: i += 1
    return games

# --- 4. ABBREVIATION & SPREAD LOGIC ---
def derive_abbreviation(team_name):
    """Fallback to capture the 4-letter abbrev often used in spreads"""
    abbrev_map = {'Nebraska': 'NEB', 'Virginia': 'UVA', 'Arkansas': 'ARK', 'Syracuse': 'SYR', 'Massachusetts': 'MASS'}
    return abbrev_map.get(team_name, team_name[:4].upper())

def flip_spread_if_needed(market, away_team, home_team, away_cover, home_cover):
    """Sets the Market perspective to the team with the higher cover percentage"""
    if not market or market == 'N/A' or not away_cover or not home_cover:
        return 'N/A'
    a_pct, h_pct = float(away_cover.replace('%','')), float(home_cover.replace('%',''))
    orig_val = float(market['value'])
    orig_abbrev = market['original_abbrev'].upper()
    away_abbrev = derive_abbreviation(away_team).upper()
    orig_is_away = (orig_abbrev == away_abbrev)

    if a_pct > h_pct:
        final_val = orig_val if orig_is_away else -orig_val
        return f"{away_abbrev} {final_val:+g}"
    else:
        home_abbrev = derive_abbreviation(home_team).upper()
        final_val = orig_val if not orig_is_away else -orig_val
        return f"{home_abbrev} {final_val:+g}"

# --- 5. ROUTES & EXECUTION ---
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        os.makedirs('static', exist_ok=True)
        games = parse_espn_schedule_from_text(request.form.get('espn_schedule', ''))
        ats = load_ats_data_from_text(request.form.get('teamrankings_ats', ''))

        if not games:
            flash('Failed to parse schedule. Check ESPN text format.', 'error')
            return redirect(url_for('index'))

        chart_rows, unmapped = [], set()
        for game in games:
            away, home = game['Away'], game['Home']
            a_map, h_map = TEAM_NAME_MAPPING.get(away, away), TEAM_NAME_MAPPING.get(home, home)
            a_data, h_data = ats.get(a_map), ats.get(h_map)

            if not a_data: unmapped.add(away)
            if not h_data: unmapped.add(home)

            a_pct, h_pct = (a_data['cover_pct'], h_data['cover_pct']) if (a_data and h_data) else (None, None)
            avg_conf, ats_pm = '', ''
            if a_pct and h_pct:
                ap, hp = float(a_pct.replace('%','')), float(h_pct.replace('%',''))
                avg_conf = f"{abs(ap - hp):.1f}"
                target = a_data if ap > hp else h_data
                ats_pm = f"{target['ats_pm']:+g}"

            chart_rows.append({
                'Away': away, 'Home': home, 'Time': game['Time'], 'Avg Conf': avg_conf,
                'A Cover %': a_pct or '', 'H Cover %': h_pct or '', 'ATS +/-': ats_pm,
                'Market': flip_spread_if_needed(game['Market'], away, home, a_pct, h_pct)
            })

        if unmapped: flash(f"Unmapped: {', '.join(sorted(unmapped))}", 'error')
        return render_template('index.html', chart_rows=chart_rows, games_count=len(chart_rows), teams_count=len(ats))
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
