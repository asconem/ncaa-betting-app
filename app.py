from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this-in-production'

# Team name mapping - ESPN Name -> TeamRankings Name
TEAM_NAME_MAPPING = {}

def load_team_name_mapping():
    """Load complete team name mappings for all 365 Division I teams"""
    global TEAM_NAME_MAPPING
    
    TEAM_NAME_MAPPING = {
        'Arkansas-Pine Bluff': 'AR-Pine Bluff', 'Abilene Christian': 'Abl Christian',
        'Alabama State': 'Alabama St', 'Alcorn State': 'Alcorn St', 'Appalachian State': 'App State',
        'Arizona State': 'Arizona St', 'Arkansas State': 'Arkansas St', 'Ball State': 'Ball St',
        'Bethune-Cookman': 'Bethune', 'Boise State': 'Boise St', 'Boston University': 'Boston U',
        'Central Arkansas': 'C Arkansas', 'Central Connecticut': 'C Connecticut', 'Central Michigan': 'C Michigan',
        'CSU Bakersfield': 'CS Bakersfield', 'Cal State Bakersfield': 'CS Bakersfield',
        'Cal State Fullerton': 'CS Fullerton', 'CSU Northridge': 'CS Northridge',
        'Cal State Northridge': 'CS Northridge', 'Cal Baptist': 'Cal Baptist', 'California Baptist': 'Cal Baptist',
        'Charleston Southern': 'Charleston So', 'Chicago State': 'Chicago St', 'Cleveland State': 'Cleveland St',
        'Coastal Carolina': 'Coastal Car', 'Colorado State': 'Colorado St', 'Coppin State': 'Coppin St',
        'East Carolina': 'E Carolina', 'Eastern Illinois': 'E Illinois', 'Eastern Kentucky': 'E Kentucky',
        'Eastern Michigan': 'E Michigan', 'East Tennessee State': 'E Tennessee St', 'East Texas A&M': 'E Texas A&M',
        'Eastern Washington': 'E Washington', 'Fairleigh Dickinson': 'F Dickinson', 'Florida Gulf Coast': 'FGCU',
        'Florida International': 'Florida Intl', 'Florida State': 'Florida St', 'Fresno State': 'Fresno St',
        'George Washington': 'G Washington', 'Gardner-Webb': 'Gardner-Webb', 'Georgia Southern': 'Georgia So',
        'Georgia State': 'Georgia St', 'Hawaii': "Hawai'i", 'Houston Christian': 'Hou Christian',
        'IU Indianapolis': 'IU Indy', 'Idaho State': 'Idaho St', 'UIC': 'Illinois Chicago',
        'Illinois State': 'Illinois St', 'Indiana State': 'Indiana St', 'Iowa State': 'Iowa St',
        'James Madison': 'J Madison', 'Jackson State': 'Jackson St', 'Jacksonville State': 'Jacksonville St',
        'Kansas State': 'Kansas St', 'Kennesaw State': 'Kennesaw St', 'Kent State': 'Kent St',
        'Long Beach State': 'Long Beach St', 'Loyola Chicago': 'Loyola Chi', 'Loyola Maryland': 'Loyola MD',
        'Loyola Marymount': 'Loyola Mymt', 'Maryland Eastern Shore': 'Maryland ES', 'Miami (OH)': 'Miami OH',
        'Miami (FL)': 'Miami', 'Michigan State': 'Michigan St', 'Middle Tennessee': 'Middle Tenn',
        'Mississippi Valley State': 'Miss Valley St', 'Ole Miss': 'Mississippi', 'Mississippi State': 'Mississippi St',
        'Missouri State': 'Missouri St', 'Montana State': 'Montana St', 'Morehead State': 'Morehead St',
        'Morgan State': 'Morgan St', "Mount St. Mary's": "Mt St Mary's", 'Murray State': 'Murray St',
        'North Alabama': 'N Alabama', 'Northern Arizona': 'N Arizona', 'Northern Colorado': 'N Colorado',
        'North Dakota State': 'N Dakota St', 'North Florida': 'N Florida', 'Northern Illinois': 'N Illinois',
        'Northern Iowa': 'N Iowa', 'Northern Kentucky': 'N Kentucky', 'North Texas': 'N Texas',
        'North Carolina A&T': 'NC A&T', 'UNC Asheville': 'NC Asheville', 'NC Central': 'NC Central',
        'UNC Greensboro': 'NC Greensboro', 'UNC Wilmington': 'NC Wilmington', 'Northwestern State': 'NW State',
        'Norfolk State': 'Norfolk St', 'Ohio State': 'Ohio St', 'Oklahoma State': 'Oklahoma St',
        'Oregon State': 'Oregon St', 'Penn State': 'Penn St', 'Portland State': 'Portland St',
        'Prairie View A&M': 'Prairie View', 'Purdue Fort Wayne': 'Purdue FW', 'Queens University': 'Queens',
        'South Alabama': 'S Alabama', 'South Carolina State': 'S Carolina St', 'South Dakota State': 'S Dakota St',
        'South Florida': 'S Florida', 'Southern Illinois': 'S Illinois', 'Southern Indiana': 'S Indiana',
        'Southern Utah': 'S Utah', 'Southeastern Louisiana': 'SE Louisiana', 'Southeast Missouri State': 'SE Missouri St',
        'Stephen F. Austin': 'SF Austin', 'SIU Edwardsville': 'SIU Edward', 'Sacramento State': 'Sacramento St',
        "Saint Mary's (CA)": "Saint Mary's", 'Sam Houston State': 'Sam Houston', 'San Diego State': 'San Diego St',
        'San Jose State': 'San Jose St', 'St. Bonaventure': 'St Bonaventure', 'St. Francis (PA)': 'St Francis PA',
        "St. John's": "St John's", 'St. Thomas': 'St Thomas', 'Tarleton State': 'Tarleton St',
        'Tennessee Tech': 'Tenn Tech', 'Tennessee State': 'Tennessee St', 'Texas A&M-Corpus Christi': 'Texas A&M-CC',
        'Texas Southern': 'Texas So', 'Texas State': 'Texas St', 'UC San Diego': 'UCSD',
        'UC Santa Barbara': 'UCSB', 'UMass': 'UMass', 'UT Rio Grande Valley': 'UT Rio Grande',
        'Utah State': 'Utah St', 'Weber State': 'Weber St', 'Western Carolina': 'W Carolina',
        'West Georgia': 'W Georgia', 'Western Illinois': 'W Illinois', 'Western Kentucky': 'W Kentucky',
        'Western Michigan': 'W Michigan', 'Wichita State': 'Wichita St', 'Wright State': 'Wright St',
        'Washington State': 'Washington St', 'Youngstown State': 'Youngstown St',
        'Air Force': 'Air Force', 'Akron': 'Akron', 'Alabama': 'Alabama', 'Alabama A&M': 'Alabama A&M',
        'Albany': 'Albany', 'UAlbany': 'Albany', 'American': 'American', 'Arizona': 'Arizona',
        'Arkansas': 'Arkansas', 'Army': 'Army', 'Auburn': 'Auburn', 'Austin Peay': 'Austin Peay',
        'BYU': 'BYU', 'Baylor': 'Baylor', 'Bellarmine': 'Bellarmine', 'Belmont': 'Belmont',
        'Binghamton': 'Binghamton', 'Boston College': 'Boston College', 'Bowling Green': 'Bowling Green',
        'Bradley': 'Bradley', 'Brown': 'Brown', 'Bryant': 'Bryant', 'Bucknell': 'Bucknell',
        'Buffalo': 'Buffalo', 'Butler': 'Butler', 'Cal Poly': 'Cal Poly', 'California': 'California',
        'Campbell': 'Campbell', 'Canisius': 'Canisius', 'Charleston': 'Charleston', 'Charlotte': 'Charlotte',
        'Chattanooga': 'Chattanooga', 'Cincinnati': 'Cincinnati', 'Clemson': 'Clemson', 'Colgate': 'Colgate',
        'Colorado': 'Colorado', 'Columbia': 'Columbia', 'Cornell': 'Cornell', 'Creighton': 'Creighton',
        'Dartmouth': 'Dartmouth', 'Davidson': 'Davidson', 'Dayton': 'Dayton', 'DePaul': 'DePaul',
        'Delaware': 'Delaware', 'Delaware State': 'Delaware St', 'Denver': 'Denver', 'Detroit Mercy': 'Detroit Mercy',
        'Drake': 'Drake', 'Drexel': 'Drexel', 'Duke': 'Duke', 'Duquesne': 'Duquesne',
        'Elon': 'Elon', 'Evansville': 'Evansville', 'Fairfield': 'Fairfield', 'Florida': 'Florida',
        'Florida A&M': 'Florida A&M', 'Florida Atlantic': 'Florida Atlantic', 'Fordham': 'Fordham',
        'Furman': 'Furman', 'George Mason': 'George Mason', 'Georgetown': 'Georgetown', 'Georgia': 'Georgia',
        'Georgia Tech': 'Georgia Tech', 'Gonzaga': 'Gonzaga', 'Grambling': 'Grambling',
        'Grand Canyon': 'Grand Canyon', 'Green Bay': 'Green Bay', 'Hampton': 'Hampton', 'Harvard': 'Harvard',
        'High Point': 'High Point', 'Hofstra': 'Hofstra', 'Holy Cross': 'Holy Cross', 'Houston': 'Houston',
        'Howard': 'Howard', 'Idaho': 'Idaho', 'Illinois': 'Illinois', 'Incarnate Word': 'Incarnate Word',
        'Indiana': 'Indiana', 'Iona': 'Iona', 'Iowa': 'Iowa', 'Jacksonville': 'Jacksonville',
        'Kansas': 'Kansas', 'Kansas City': 'Kansas City', 'Kentucky': 'Kentucky', 'LIU': 'LIU',
        'LSU': 'LSU', 'La Salle': 'La Salle', 'Lafayette': 'Lafayette', 'Lamar': 'Lamar',
        'Le Moyne': 'Le Moyne', 'Lehigh': 'Lehigh', 'Liberty': 'Liberty', 'Lindenwood': 'Lindenwood',
        'Lipscomb': 'Lipscomb', 'Little Rock': 'Little Rock', 'Longwood': 'Longwood', 'Louisiana': 'Louisiana',
        'Louisiana Tech': 'Louisiana Tech', 'Louisville': 'Louisville', 'Maine': 'Maine', 'Manhattan': 'Manhattan',
        'Marist': 'Marist', 'Marquette': 'Marquette', 'Marshall': 'Marshall', 'Maryland': 'Maryland',
        'McNeese': 'McNeese', 'Memphis': 'Memphis', 'Mercer': 'Mercer', 'Mercyhurst': 'Mercyhurst',
        'Merrimack': 'Merrimack', 'Michigan': 'Michigan', 'Milwaukee': 'Milwaukee', 'Minnesota': 'Minnesota',
        'Missouri': 'Missouri', 'Monmouth': 'Monmouth', 'Montana': 'Montana', 'NJIT': 'NJIT',
        'NC State': 'NC State', 'Navy': 'Navy', 'Nebraska': 'Nebraska', 'Nevada': 'Nevada',
        'New Hampshire': 'New Hampshire', 'New Haven': 'New Haven', 'New Mexico': 'New Mexico',
        'New Mexico State': 'New Mexico St', 'New Orleans': 'New Orleans', 'Niagara': 'Niagara',
        'Nicholls': 'Nicholls', 'North Carolina': 'North Carolina', 'North Dakota': 'North Dakota',
        'Northeastern': 'Northeastern', 'Northwestern': 'Northwestern', 'Notre Dame': 'Notre Dame',
        'Oakland': 'Oakland', 'Ohio': 'Ohio', 'Oklahoma': 'Oklahoma', 'Old Dominion': 'Old Dominion',
        'Omaha': 'Omaha', 'Oral Roberts': 'Oral Roberts', 'Oregon': 'Oregon', 'Pacific': 'Pacific',
        'Penn': 'Penn', 'Pepperdine': 'Pepperdine', 'Pittsburgh': 'Pittsburgh', 'Portland': 'Portland',
        'Presbyterian': 'Presbyterian', 'Princeton': 'Princeton', 'Providence': 'Providence', 'Purdue': 'Purdue',
        'Queens': 'Queens', 'Quinnipiac': 'Quinnipiac', 'Radford': 'Radford', 'Rhode Island': 'Rhode Island',
        'Rice': 'Rice', 'Richmond': 'Richmond', 'Rider': 'Rider', 'Robert Morris': 'Robert Morris',
        'Rutgers': 'Rutgers', 'SC Upstate': 'SC Upstate', 'SMU': 'SMU', 'Sacred Heart': 'Sacred Heart',
        "Saint Joseph's": "Saint Joseph's", 'Saint Louis': 'Saint Louis', "Saint Peter's": "Saint Peter's",
        'Samford': 'Samford', 'San Diego': 'San Diego', 'San Francisco': 'San Francisco',
        'Santa Clara': 'Santa Clara', 'Seattle': 'Seattle', 'Seton Hall': 'Seton Hall', 'Siena': 'Siena',
        'South Carolina': 'South Carolina', 'South Dakota': 'South Dakota', 'Southern': 'Southern',
        'Southern Miss': 'Southern Miss', 'Stanford': 'Stanford', 'Stetson': 'Stetson',
        'Stonehill': 'Stonehill', 'Stony Brook': 'Stony Brook', 'Syracuse': 'Syracuse', 'TCU': 'TCU',
        'Temple': 'Temple', 'Tennessee': 'Tennessee', 'Texas': 'Texas', 'Texas A&M': 'Texas A&M',
        'Texas Tech': 'Texas Tech', 'The Citadel': 'The Citadel', 'Toledo': 'Toledo', 'Towson': 'Towson',
        'Troy': 'Troy', 'Tulane': 'Tulane', 'Tulsa': 'Tulsa', 'UAB': 'UAB',
        'UC Davis': 'UC Davis', 'UC Irvine': 'UC Irvine', 'UC Riverside': 'UC Riverside', 'UCF': 'UCF',
        'UCLA': 'UCLA', 'UConn': 'UConn', 'UL Monroe': 'UL Monroe', 'UMBC': 'UMBC',
        'UMass Lowell': 'UMass Lowell', 'UNLV': 'UNLV', 'USC': 'USC', 'UT Arlington': 'UT Arlington',
        'UT Martin': 'UT Martin', 'UTEP': 'UTEP', 'UTSA': 'UTSA', 'Utah': 'Utah',
        'Utah Tech': 'Utah Tech', 'Utah Valley': 'Utah Valley', 'VCU': 'VCU', 'VMI': 'VMI',
        'Valparaiso': 'Valparaiso', 'Vanderbilt': 'Vanderbilt', 'Vermont': 'Vermont', 'Villanova': 'Villanova',
        'Virginia': 'Virginia', 'Virginia Tech': 'Virginia Tech', 'Wagner': 'Wagner', 'Wake Forest': 'Wake Forest',
        'Washington': 'Washington', 'West Virginia': 'West Virginia', 'William & Mary': 'William & Mary',
        'Winthrop': 'Winthrop', 'Wisconsin': 'Wisconsin', 'Wofford': 'Wofford', 'Wyoming': 'Wyoming',
        'Xavier': 'Xavier', 'Yale': 'Yale',
    }

def parse_espn_schedule_from_text(text):
    """Parse ESPN schedule from pasted text"""
    lines = text.strip().split('\n')
    format_type = detect_schedule_format(lines)
    
    if format_type == 'mobile':
        return parse_mobile_format(lines)
    else:
        return parse_desktop_format(lines)

def detect_schedule_format(lines):
    """Detect whether schedule is from mobile or desktop format"""
    for i in range(min(100, len(lines) - 2)):
        line = lines[i].strip()
        if line and i+1 < len(lines):
            next_line = lines[i+1].strip()
            
            if re.match(r'^\(\d+-\d+.*?\)$', next_line):
                return 'desktop'
            
            if re.match(r'^\d+-\d+$', next_line):
                if i+2 < len(lines):
                    if lines[i+2].strip() == '' or re.match(r'^\d+-\d+$', lines[i+3].strip() if i+3 < len(lines) else ''):
                        return 'mobile'
    
    return 'desktop'

def parse_desktop_format(lines):
    """Parse desktop format schedule"""
    games = []
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        
        if re.match(r'^\d{1,2}:\d{2}\s*[AP]M$', line):
            time = line
            away_team = None
            home_team = None
            spread = None
            
            j = i + 1
            while j < min(i+40, len(lines)):
                check_line = lines[j].strip()
                
                if away_team and home_team and (re.match(r'^\d{1,2}:\d{2}\s*[AP]M$', check_line) or 'Gamecast' in check_line):
                    break
                
                if away_team and home_team:
                    if check_line.startswith('Spread:'):
                        spread_match = re.match(r'^Spread:([A-Z0-9&\-]+)\s+([-+]?\d+\.?\d*)$', check_line)
                        if spread_match:
                            team_abbrev = spread_match.group(1)
                            spread_value = spread_match.group(2)
                            spread = {'original_abbrev': team_abbrev, 'value': spread_value, 'display': f"{team_abbrev} {spread_value}"}
                    j += 1
                    continue
                
                if re.match(r'^\(\d+-\d+.*?(Home|Away)\)$', check_line) and not away_team and j >= 2:
                    if j-2 >= 0 and j-1 >= 0:
                        potential_away = lines[j-2].strip()
                        potential_home = lines[j-1].strip()
                        if potential_away and potential_home and not re.match(r'^\d', potential_away):
                            away_team = potential_away
                            home_team = potential_home
                            j += 1
                            continue
                
                if check_line and j+1 < len(lines) and not away_team:
                    next_line = lines[j+1].strip()
                    if re.match(r'^\(\d+-\d+.*?\)$', next_line) and not re.match(r'^\(\d+-\d+.*?(Home|Away)\)$', next_line):
                        away_team = check_line
                        j += 2
                        continue
                
                if check_line and j+1 < len(lines) and away_team and not home_team:
                    next_line = lines[j+1].strip()
                    if re.match(r'^\(\d+-\d+.*?\)$', next_line) and not re.match(r'^\(\d+-\d+.*?(Home|Away)\)$', next_line):
                        home_team = check_line
                        j += 2
                        continue
                
                j += 1
            
            if away_team and home_team:
                market_display = spread['display'] if spread else 'PK'
                games.append({'Away': away_team, 'Home': home_team, 'Time': time, 'Market': market_display})
            
            i = j
        else:
            i += 1
    
    return games

def parse_mobile_format(lines):
    """Parse mobile format schedule"""
    games = []
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        
        if re.match(r'^\d{1,2}:\d{2}\s*[AP]M$', line):
            time = line
            
            if i+5 < len(lines):
                away_team = lines[i+1].strip()
                home_team = lines[i+3].strip()
                
                if away_team and home_team:
                    spread = None
                    j = i + 5
                    while j < min(i+20, len(lines)):
                        check_line = lines[j].strip()
                        if check_line == 'Spread:' and j+1 < len(lines):
                            spread_line = lines[j+1].strip()
                            spread_match = re.match(r'^([A-Z0-9&\-]+)\s+([-+]?\d+\.?\d*)$', spread_line)
                            if spread_match:
                                team_abbrev = spread_match.group(1)
                                spread_value = spread_match.group(2)
                                spread = f"{team_abbrev} {spread_value}"
                            break
                        j += 1
                    
                    market_display = spread if spread else 'PK'
                    games.append({'Away': away_team, 'Home': home_team, 'Time': time, 'Market': market_display})
            
            i += 6
        else:
            i += 1
    
    return games

def load_ats_data_from_text(text):
    """Parse ATS data from TeamRankings paste"""
    ats_dict = {}
    lines = text.strip().split('\n')
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        match = re.match(r'^(\d+)\s+(.+?)\s+(\d+-\d+-?\d*)\s+([\d.]+%)\s+([-+]?\d+\.?\d*)$', line)
        if match:
            rank = match.group(1)
            team_name = match.group(2).strip()
            ats_record = match.group(3)
            cover_pct = match.group(4)
            ats_pm = float(match.group(5))
            
            ats_dict[team_name] = {'rank': rank, 'record': ats_record, 'cover_pct': cover_pct, 'ats_pm': ats_pm}
            continue
        
        match2 = re.match(r'^(.+?)\s+(\d+-\d+-?\d*)\s+([\d.]+%)\s+[-+]?\d+\.?\d*\s+([-+]?\d+\.?\d*)$', line)
        if match2:
            team_name = match2.group(1).strip()
            ats_record = match2.group(2)
            cover_pct = match2.group(3)
            ats_pm = float(match2.group(4))
            
            ats_dict[team_name] = {'rank': 'N/A', 'record': ats_record, 'cover_pct': cover_pct, 'ats_pm': ats_pm}
    
    return ats_dict

def find_team_cover_pct(team_name, ats_dict, name_mapping):
    """Find team's cover percentage with name mapping"""
    if team_name in ats_dict:
        return ats_dict[team_name]['cover_pct']
    
    if team_name in name_mapping:
        mapped_name = name_mapping[team_name]
        if mapped_name in ats_dict:
            return ats_dict[mapped_name]['cover_pct']
    
    return None

def find_team_ats_plus_minus(team_name, ats_dict, name_mapping):
    """Find team's ATS +/- with name mapping"""
    if team_name in ats_dict:
        return ats_dict[team_name]['ats_pm']
    
    if team_name in name_mapping:
        mapped_name = name_mapping[team_name]
        if mapped_name in ats_dict:
            return ats_dict[mapped_name]['ats_pm']
    
    return None

def flip_spread_if_needed(market, away_team, home_team, away_cover, home_cover):
    """Flip spread to show perspective of team with higher cover %"""
    if not away_cover or not home_cover or market == 'PK':
        return market
    
    away_pct = float(away_cover.replace('%', ''))
    home_pct = float(home_cover.replace('%', ''))
    
    if away_pct == home_pct:
        return market
    
    match = re.match(r'^([A-Z0-9&\-]+)\s+([-+]?\d+\.?\d*)$', market)
    if not match:
        return market
    
    team_abbrev = match.group(1)
    spread_value = float(match.group(2))
    
    higher_cover_is_away = away_pct > home_pct
    spread_is_away = True
    
    if higher_cover_is_away and not spread_is_away:
        new_spread = -spread_value
        return f"{team_abbrev} {new_spread:+.1f}".replace('+', '')
    elif not higher_cover_is_away and spread_is_away:
        new_spread = -spread_value
        return f"{team_abbrev} {new_spread:+.1f}".replace('+', '')
    
    return market

def create_daily_chart(games, ats_dict, name_mapping):
    """Create the daily betting analysis chart"""
    chart_rows = []
    
    for game in games:
        away_team = game['Away']
        home_team = game['Home']
        
        away_cover = find_team_cover_pct(away_team, ats_dict, name_mapping)
        home_cover = find_team_cover_pct(home_team, ats_dict, name_mapping)
        
        away_ats_pm = find_team_ats_plus_minus(away_team, ats_dict, name_mapping)
        home_ats_pm = find_team_ats_plus_minus(home_team, ats_dict, name_mapping)
        
        avg_conf = ''
        play_team_ats = ''
        if away_cover and home_cover:
            away_pct = float(away_cover.replace('%', ''))
            home_pct = float(home_cover.replace('%', ''))
            avg_conf = abs(away_pct - home_pct)
            avg_conf = f"{avg_conf:.1f}"
            
            if away_pct > home_pct and away_ats_pm is not None:
                play_team_ats = f"+{away_ats_pm:.1f}" if away_ats_pm >= 0 else f"{away_ats_pm:.1f}"
            elif home_pct > away_pct and home_ats_pm is not None:
                play_team_ats = f"+{home_ats_pm:.1f}" if home_ats_pm >= 0 else f"{home_ats_pm:.1f}"
            elif away_pct == home_pct:
                play_team_ats = "-"
        
        adjusted_market = flip_spread_if_needed(game['Market'], away_team, home_team, away_cover, home_cover)
        
        chart_rows.append({
            'Away': away_team,
            'Home': home_team,
            'Market': adjusted_market,
            'A Cover %': away_cover if away_cover else '',
            'H Cover %': home_cover if home_cover else '',
            'Avg Conf': avg_conf,
            'ATS +/-': play_team_ats,
            'Time': game['Time']
        })
    
    return chart_rows

def create_xlsx_file(chart_rows, filename):
    """Create XLSX file with color coding"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Chart"
    
    headers = ['Away', 'Home', 'Market', 'A Cover %', 'H Cover %', 'Avg Conf', 'ATS +/-', 'Time']
    ws.append(headers)
    
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    bold_border = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))
    
    for cell in ws[1]:
        cell.border = bold_border
    
    for row_data in chart_rows:
        row_values = [row_data['Away'], row_data['Home'], row_data['Market'], 
                     row_data['A Cover %'], row_data['H Cover %'], 
                     row_data['Avg Conf'], row_data['ATS +/-'], row_data['Time']]
        ws.append(row_values)
        
        current_row = ws.max_row
        
        for col_num in range(1, 9):
            cell = ws.cell(current_row, col_num)
            cell.border = bold_border
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        if row_data['Avg Conf']:
            try:
                conf_val = float(row_data['Avg Conf'])
                fill_color = None
                
                if conf_val >= 50:
                    fill_color = green_fill
                elif conf_val >= 30:
                    fill_color = yellow_fill
                
                if fill_color:
                    for col_num in range(1, 9):
                        ws.cell(current_row, col_num).fill = fill_color
            except ValueError:
                pass
    
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 12
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 12
    
    wb.save(filename)

# Ensure static directory exists when app starts
os.makedirs('static', exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        espn_schedule = request.form.get('espn_schedule', '')
        teamrankings_ats = request.form.get('teamrankings_ats', '')
        
        if not espn_schedule or not teamrankings_ats:
            flash('Please provide both ESPN schedule and TeamRankings ATS data', 'error')
            return redirect(url_for('index'))
        
        try:
            games = parse_espn_schedule_from_text(espn_schedule)
            ats_dict = load_ats_data_from_text(teamrankings_ats)
            
            if not games:
                flash('Could not parse any games from ESPN schedule', 'error')
                return redirect(url_for('index'))
            
            if not ats_dict:
                flash('Could not parse ATS data from TeamRankings', 'error')
                return redirect(url_for('index'))
            
            chart_rows = create_daily_chart(games, ats_dict, TEAM_NAME_MAPPING)
            
            output_filename = f'daily_chart_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
            output_path = os.path.join('static', output_filename)
            create_xlsx_file(chart_rows, output_path)
            
            flash(f'Successfully generated chart with {len(chart_rows)} games!', 'success')
            return render_template('index.html', chart_rows=chart_rows, download_file=output_filename, games_count=len(games), teams_count=len(ats_dict))
        
        except Exception as e:
            flash(f'Error processing data: {str(e)}', 'error')
            return redirect(url_for('index'))
    
    return render_template('index.html')

@app.route('/download/<filename>')
def download(filename):
    """Download the generated XLSX file"""
    return send_file(os.path.join('static', filename), as_attachment=True)

if __name__ == '__main__':
    os.makedirs('static', exist_ok=True)
    load_team_name_mapping()
    app.run(debug=True)
