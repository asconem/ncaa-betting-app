from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
import re
import logging
from logging.handlers import RotatingFileHandler
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this-in-production'

# Configure logging
if not os.path.exists('logs'):
    os.makedirs('logs')

logger = logging.getLogger('ncaa_ats_app')
logger.setLevel(logging.INFO)

# File handler with rotation (10MB max, keep 7 backup files)
file_handler = RotatingFileHandler('logs/app.log', maxBytes=10*1024*1024, backupCount=7)
file_handler.setLevel(logging.INFO)

# Console handler
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# Format: timestamp - level - message
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Team name mapping - ESPN Name -> TeamRankings Name
TEAM_NAME_MAPPING = {}

def load_team_name_mapping():
    """Load complete team name mappings for all 365 Division I teams"""
    global TEAM_NAME_MAPPING
    
    TEAM_NAME_MAPPING = {
        'Arkansas-Pine Bluff': 'AR-Pine Bluff', 'Abilene Christian': 'Abl Christian',
        'Alabama State': 'Alabama St', 'Alcorn State': 'Alcorn St', 'Appalachian State': 'App State',
        'App State': 'App State',
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
        'Georgia State': 'Georgia St', 'Hawaii': "Hawai'i", "Hawai'i": "Hawai'i", "Hawai'i": "Hawai'i", 'Houston Christian': 'Hou Christian',
        'IU Indianapolis': 'IU Indy', 'Idaho State': 'Idaho St', 'UIC': 'Illinois Chicago',
        'Illinois State': 'Illinois St', 'Indiana State': 'Indiana St', 'Iowa State': 'Iowa St',
        'James Madison': 'J Madison', 'Jackson State': 'Jackson St', 'Jacksonville State': 'Jacksonville St',
        'Kansas State': 'Kansas St', 'Kennesaw State': 'Kennesaw St', 'Kent State': 'Kent St',
        'Long Beach State': 'Long Beach St', 'Loyola Chicago': 'Loyola Chi', 'Loyola Maryland': 'Loyola MD',
        'Loyola Marymount': 'Loyola Mymt', 'Maryland Eastern Shore': 'Maryland ES',
        'Maryland-Eastern Shore': 'Maryland ES', 'UMES': 'Maryland ES',
        'Miami (OH)': 'Miami OH',
        'Miami (FL)': 'Miami', 'Michigan State': 'Michigan St', 'Middle Tennessee': 'Middle Tenn',
        'Mississippi Valley State': 'Miss Valley St', 'Ole Miss': 'Mississippi', 'Mississippi State': 'Mississippi St',
        'Missouri State': 'Missouri St', 'Montana State': 'Montana St', 'Morehead State': 'Morehead St',
        'Morgan State': 'Morgan St', "Mount St. Mary's": "Mt St Mary's", "Mt. St. Mary's": "Mt St Mary's",
        "Mount St Mary's": "Mt St Mary's", 'Murray State': 'Murray St',
        'North Alabama': 'N Alabama', 'Northern Arizona': 'N Arizona', 'Northern Colorado': 'N Colorado',
        'North Dakota State': 'N Dakota St', 'North Florida': 'N Florida', 'Northern Illinois': 'N Illinois',
        'Northern Iowa': 'N Iowa', 'Northern Kentucky': 'N Kentucky', 'North Texas': 'N Texas',
        'North Carolina A&T': 'NC A&T', 'UNC Asheville': 'NC Asheville', 'NC Central': 'NC Central',
        'North Carolina Central': 'NC Central',
        'UNC Greensboro': 'NC Greensboro', 'UNC Wilmington': 'NC Wilmington', 'Northwestern State': 'NW State',
        'Norfolk State': 'Norfolk St', 'Ohio State': 'Ohio St', 'Oklahoma State': 'Oklahoma St',
        'Oregon State': 'Oregon St', 'Penn State': 'Penn St', 'Portland State': 'Portland St',
        'Prairie View A&M': 'Prairie View', 'Purdue Fort Wayne': 'Purdue FW', 'Queens University': 'Queens',
        'South Alabama': 'S Alabama', 'South Carolina State': 'S Carolina St', 'South Dakota State': 'S Dakota St',
        'South Florida': 'S Florida', 'Southern Illinois': 'S Illinois', 'Southern Indiana': 'S Indiana',
        'Southern Utah': 'S Utah', 'Southeastern Louisiana': 'SE Louisiana', 'SE Louisiana': 'SE Louisiana',
        'Southeast Missouri State': 'SE Missouri St',
        'Pennsylvania': 'Penn',
        'Stephen F. Austin': 'SF Austin', 'Stephen F Austin': 'SF Austin',
        'SIU Edwardsville': 'SIU Edward', 'Sacramento State': 'Sacramento St',
        "Saint Mary's (CA)": "Saint Mary's", 'Sam Houston State': 'Sam Houston', 
        "St. Mary's": "Saint Mary's", "St. Mary's (CA)": "Saint Mary's", "Saint Mary's": "Saint Mary's", 'San Diego State': 'San Diego St',
        'San Jose State': 'San Jose St', 'San José State': 'San Jose St',
        'St. Bonaventure': 'St Bonaventure', 'St. Francis (PA)': 'St Francis PA', 'Saint Francis': 'St Francis PA',
        "St. John's": "St John's", "St John's": "St John's", "Saint John's": "St John's",
        'St. Thomas': 'St Thomas', 'St. Thomas-Minnesota': 'St Thomas', 'Tarleton State': 'Tarleton St',
        'Tennessee Tech': 'Tenn Tech', 'Tennessee State': 'Tennessee St', 'Texas A&M-Corpus Christi': 'Texas A&M-CC',
        'Texas Southern': 'Texas So', 'Texas State': 'Texas St', 'UC San Diego': 'UCSD',
        'UC Santa Barbara': 'UCSB', 'UMass': 'UMass', 'UT Rio Grande Valley': 'UT Rio Grande',
        'Utah State': 'Utah St', 'Weber State': 'Weber St', 'Western Carolina': 'W Carolina',
        'West Georgia': 'W Georgia', 'Western Illinois': 'W Illinois', 'Western Kentucky': 'W Kentucky',
        'Western Michigan': 'W Michigan', 'Wichita State': 'Wichita St', 'Wright State': 'Wright St',
        'Washington State': 'Washington St', 'Youngstown State': 'Youngstown St',
        'Air Force': 'Air Force', 'Akron': 'Akron', 'Alabama': 'Alabama', 'Alabama A&M': 'Alabama A&M',
        'Albany': 'Albany', 'UAlbany': 'Albany', 'American': 'American', 'American University': 'American',
        'Arizona': 'Arizona',
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
        'Rutgers': 'Rutgers', 'SC Upstate': 'SC Upstate', 'South Carolina Upstate': 'SC Upstate',
        'SMU': 'SMU', 'Sacred Heart': 'Sacred Heart',
        "Saint Joseph's": "Saint Joseph's", 'Saint Louis': 'Saint Louis', "Saint Peter's": "Saint Peter's",
        "St. Joseph's": "Saint Joseph's", 'St. Louis': 'Saint Louis', "St. Peter's": "Saint Peter's",
        "Saint Joe's": "Saint Joseph's", "St. Joe's": "Saint Joseph's",
        'Samford': 'Samford', 'San Diego': 'San Diego', 'San Francisco': 'San Francisco',
        'Santa Clara': 'Santa Clara', 'Seattle': 'Seattle', 'Seattle U': 'Seattle', 'Seattle University': 'Seattle',
        'Seton Hall': 'Seton Hall', 'Siena': 'Siena',
        'South Carolina': 'South Carolina', 'South Dakota': 'South Dakota', 'Southern': 'Southern',
        'Southern Miss': 'Southern Miss', 'Stanford': 'Stanford', 'Stetson': 'Stetson',
        'Stonehill': 'Stonehill', 'Stony Brook': 'Stony Brook', 'Syracuse': 'Syracuse', 'TCU': 'TCU',
        'Temple': 'Temple', 'Tennessee': 'Tennessee', 'Texas': 'Texas', 'Texas A&M': 'Texas A&M',
        'Texas Tech': 'Texas Tech', 'The Citadel': 'The Citadel', 'Toledo': 'Toledo', 'Towson': 'Towson',
        'Troy': 'Troy', 'Tulane': 'Tulane', 'Tulsa': 'Tulsa', 'UAB': 'UAB',
        'UC Davis': 'UC Davis', 'UC Irvine': 'UC Irvine', 'UC Riverside': 'UC Riverside', 'UCF': 'UCF',
        'UCLA': 'UCLA', 'UConn': 'UConn', 'UL Monroe': 'UL Monroe', 'UMBC': 'UMBC',
        'Connecticut': 'UConn', 'Massachusetts': 'UMass', 'Massachusetts Lowell': 'UMass Lowell',
        'Maryland Baltimore County': 'UMBC', 'Virginia Commonwealth': 'VCU',
        'Central Florida': 'UCF', 'Southern Methodist': 'SMU', 'Southern California': 'USC',
        'Louisiana State': 'LSU', 'Brigham Young': 'BYU', 'Texas Christian': 'TCU',
        'North Carolina State': 'NC State', 'N.C. State': 'NC State',
        'Miami FL': 'Miami', 'Miami (Ohio)': 'Miami OH',
        'UMass Lowell': 'UMass Lowell', 'UNLV': 'UNLV', 'USC': 'USC', 'UT Arlington': 'UT Arlington',
        'UT Martin': 'UT Martin', 'UTEP': 'UTEP', 'UTSA': 'UTSA', 'Utah': 'Utah',
        'Utah Tech': 'Utah Tech', 'Utah Valley': 'Utah Valley', 'VCU': 'VCU', 'VMI': 'VMI',
        'Valparaiso': 'Valparaiso', 'Vanderbilt': 'Vanderbilt', 'Vermont': 'Vermont', 'Villanova': 'Villanova',
        'Virginia': 'Virginia', 'Virginia Tech': 'Virginia Tech', 'Wagner': 'Wagner', 'Wake Forest': 'Wake Forest',
        'Washington': 'Washington', 'West Virginia': 'West Virginia', 'William & Mary': 'William & Mary',
        'William and Mary': 'William & Mary',
        'Winthrop': 'Winthrop', 'Wisconsin': 'Wisconsin', 'Wofford': 'Wofford', 'Wyoming': 'Wyoming',
        'Xavier': 'Xavier', 'Yale': 'Yale',
    }

# Load the mapping when the module is imported
load_team_name_mapping()

def derive_abbreviation(team_name):
    """Derive team abbreviation from full name - EXACT copy from skill"""
    abbrev_map = {
        'Air Force': 'AF', 'Akron': 'AKR', 'Alabama State': 'ALST', 'Middle Tennessee': 'MTSU',
        'Michigan': 'MICH', 'Harvard': 'HARV', 'Penn State': 'PSU', 'Arizona': 'ARIZ',
        'UConn': 'CONN', 'Chattanooga': 'UTC', 'South Carolina State': 'SCST',
        'South Alabama': 'USA', 'Jacksonville State': 'JXST', 'Bethune-Cookman': 'BCU',
        'Ohio': 'OHIO', 'Louisiana Tech': 'LT', 'Indiana State': 'INST', 'Hofstra': 'HOF',
        'Temple': 'TEM', 'Tennessee Tech': 'TNTC', 'South Carolina Upstate': 'UPST',
        'Howard': 'HOW', 'Stetson': 'STET', 'North Florida': 'UNF', 'Wofford': 'WOF',
        'William & Mary': 'WM', 'Bowling Green': 'BGSU', 'Indiana': 'IU', 'Utah State': 'USU',
        'St. Bonaventure': 'SBU', 'Western Michigan': 'WMU', 'UNC Wilmington': 'UNCW',
        'North Carolina Central': 'NCC', 'Kent State': 'KENT', 'Old Dominion': 'ODU',
        'Northern Illinois': 'NIU', 'Louisiana': 'UL', 'UL Monroe': 'ULM', 'Texas State': 'TXST',
        'Southern Miss': 'USM', 'Southeast Missouri State': 'SEMO', 'Southern Indiana': 'USI',
        'Loyola Chicago': 'LUC', 'Santa Clara': 'SCU', 'James Madison': 'JMU',
        'Georgia Southern': 'GASO', 'Omaha': 'OMA', 'Portland State': 'PRST',
        'UC Riverside': 'UCR', 'Sacramento State': 'SAC', 'Northeastern': 'NE',
        'Northwestern': 'NU', 'Central Michigan': 'CMU', 'Liberty': 'LIB',
        'Chicago State': 'CHST', 'Saint Francis': 'SFPA', 'Georgia State': 'GAST',
        'Denver': 'DEN', 'Marshall': 'MRSH', 'Grand Canyon': 'GCU', 'UT Martin': 'UTM',
        'Utah Tech': 'UTU', 'Florida Gulf Coast': 'FGCU', 'Samford': 'SAM',
        'Youngstown State': 'YSU', 'Toledo': 'TOL', 'Valparaiso': 'VALP',
        'Cleveland State': 'CLE', 'Villanova': 'VILL', 'La Salle': 'LAS', 'Bryant': 'BRY',
        'Virginia Tech': 'VT', 'Bellarmine': 'BELL', 'Notre Dame': 'ND', 'VMI': 'VMI',
        'Richmond': 'RICH', 'UMBC': 'UMBC', 'George Washington': 'GW',
        'Eastern Washington': 'EWU', 'Central Arkansas': 'CARK', 'Loyola Maryland': 'LOM',
        'Duquesne': 'DUQ', "Mount St. Mary's": 'MSMT', 'Maryland': 'MD',
        'UNC Asheville': 'UNCA', 'UNC Greensboro': 'UNCG', 'Western Carolina': 'WCU',
        'Wyoming': 'WYO', 'Sam Houston': 'SHSU', 'Maine': 'ME', 'Merrimack': 'MRMK',
        'Dayton': 'DAY', 'Marquette': 'MARQ', 'UC Irvine': 'UCI', 'Utah Valley': 'UVU',
        'UMass Lowell': 'UML', 'Bradley': 'BRAD', 'North Dakota': 'UND',
        'South Dakota': 'SDAK', 'Northern Colorado': 'UNCO', 'Southern Illinois': 'SIU',
        'Prairie View A&M': 'PVAM', 'Creighton': 'CREI', 'South Florida': 'USF',
        'Oklahoma State': 'OKST', 'Kansas City': 'UMKC', 'TCU': 'TCU', 'Lipscomb': 'LIP',
        'Belmont': 'BEL', 'Jackson State': 'JKST', 'Auburn': 'AUB', 'Arkansas State': 'ARST',
        "Saint Mary's": 'SMC', 'Campbell': 'CAM', 'Weber State': 'WEB', 'Southern Utah': 'SUU',
        'Washington State': 'WSU', 'Buffalo': 'BUF', 'Bucknell': 'BUCK',
        'Tarleton State': 'TAR', 'Tarleton St': 'TAR', 'Sacred Heart': 'SHU',
        'Dartmouth': 'DART', 'Wright State': 'WRST', 'Evansville': 'EVAN',
        'Mississippi Valley State': 'MVSU', 'Mississippi Valley St': 'MVSU',
        'Morehead State': 'MORE', 'Morehead St': 'MORE', 'Montana State': 'MTST',
        'Montana St': 'MTST', 'Mississippi State': 'MSST', 'Mississippi St': 'MSST',
        'Tulane': 'TULN', 'Boise State': 'BOIS', 'Georgetown': 'GTWN',
        'Long Island University': 'LIU', 'Long Island': 'LIU', 'Oral Roberts': 'ORU',
        'UC San Diego': 'UCSD', 'Loyola Marymount': 'LMU', 'East Tennessee State': 'ETSU',
        'East Tennessee St': 'ETSU', 'North Carolina A&T': 'NCAT', 'Abilene Christian': 'ACU',
        'Pacific': 'PACI', 'Queens University': 'QUC', 'Queens': 'QUC',
        'California Baptist': 'CBU', 'Cal Baptist': 'CBU', 'Southern': 'SOU',
        'Drexel': 'DREX', 'Florida International': 'FIU', 'Marist': 'MARI',
        'Lamar': 'LAM', 'Delaware': 'DELA', 'Alabama': 'ALA', 'Illinois': 'ILL',
        'North Carolina': 'UNC', 'Duke': 'DUKE', 'Kentucky': 'UK', 'Kansas': 'KU',
        'Gonzaga': 'GONZ', 'UCLA': 'UCLA', 'USC': 'USC', 'Michigan State': 'MSU',
        'Ohio State': 'OSU', 'Florida State': 'FSU', 'NC State': 'NCST', 'Virginia': 'UVA',
        'Purdue': 'PUR', 'Wisconsin': 'WISC', 'Iowa': 'IOWA', 'Texas': 'TEX',
        'Baylor': 'BAY', 'Tennessee': 'TENN', 'Arkansas': 'ARK', 'LSU': 'LSU',
        'Florida': 'FLA', 'Georgia': 'UGA', 'Missouri': 'MIZZ', 'Missouri State': 'MOST',
        'Missouri St': 'MOST', 'Texas A&M': 'TAMU', 'East Texas A&M': 'ETAM',
        'Mississippi': 'MISS', 'Vanderbilt': 'VAN', 'South Carolina': 'SC',
        'Iowa State': 'ISU', 'Kansas State': 'KSU', 'West Virginia': 'WVU',
        'Oklahoma': 'OU', 'Texas Tech': 'TTU', 'Arizona State': 'ASU', 'Colorado': 'COLO',
        'Utah': 'UTAH', 'Oregon': 'ORE', 'Washington': 'WASH', 'Stanford': 'STAN',
        'California': 'CAL', 'Oregon State': 'ORST', 'Syracuse': 'SYR', 'Louisville': 'LOU',
        'Pittsburgh': 'PITT', 'Miami': 'MIA', 'Miami (OH)': 'M-OH', 'Miami OH': 'M-OH',
        'Boston College': 'BC', 'Clemson': 'CLEM', 'Wake Forest': 'WAKE',
        'Georgia Tech': 'GT', 'Providence': 'PROV', 'Xavier': 'XAV', 'Butler': 'BUT',
        'Seton Hall': 'HALL', "St. John's": 'SJU', 'St. Thomas-Minnesota': 'STMN',
        'DePaul': 'DEP', 'BYU': 'BYU', 'Houston': 'HOU', 'Cincinnati': 'CIN',
        'Memphis': 'MEM', 'SMU': 'SMU', 'UCF': 'UCF', 'Wichita State': 'WICH',
        'San Diego State': 'SDSU', 'Colorado State': 'CSU', 'Fresno State': 'FRES',
        'Nevada': 'NEV', 'New Mexico': 'UNM', 'San Jose State': 'SJSU', 'UNLV': 'UNLV',
        'Le Moyne': 'LEM', 'Ball State': 'BALL', "Saint Joseph's": 'JOES',
        'App State': 'APP', 'UTSA': 'UTSA', 'Tennessee State': 'TNST',
        'William & Mary': 'W&M', 'Robert Morris': 'RMU', 'New Hampshire': 'UNH',
        'Fairfield': 'FAIR', 'Stonehill': 'STON', 'Quinnipiac': 'QUIN', 'Yale': 'YALE',
        'Vermont': 'UVM', 'Cornell': 'COR', 'Monmouth': 'MONM', 'Lafayette': 'LAF',
        'Pennsylvania': 'PENN', 'Mercyhurst': 'MERC', 'Fordham': 'FORD', 'Colgate': 'COLG',
        'Alabama A&M': 'AAMU', 'Coastal Carolina': 'CCU', 'Florida Atlantic': 'FAU',
        'Charleston': 'CHSN', 'Charleston Southern': 'CHSO', 'Charlotte': 'CLT',
        'North Texas': 'UNT', 'Longwood': 'LONG', 'American University': 'AMER',
        'The Citadel': 'CIT', 'Presbyterian': 'PRES', 'San José State': 'SJSU',
        'San Diego': 'USD', 'Long Beach State': 'LBSU', 'Cal State Bakersfield': 'CSUB',
        'Coppin State': 'COPP', 'Norfolk State': 'NORF', 'Tulsa': 'TLSA',
        'Rutgers': 'RUTG', 'Massachusetts': 'MASS', 'Elon': 'ELON',
        'Central Connecticut': 'CCSU', 'Davidson': 'DAV', 'Furman': 'FUR', 'Iona': 'IONA',
        'Binghamton': 'BING', 'High Point': 'HPU', 'Lindenwood': 'LIN',
        'Murray State': 'MUR', 'Northern Arizona': 'NAU', 'North Alabama': 'UNA',
        'Rhode Island': 'URI', 'Seattle U': 'SEA', 'SIU Edwardsville': 'SIUE',
        'South Dakota State': 'SDST', 'Western Kentucky': 'WKU', 'Winthrop': 'WIN',
        'Holy Cross': 'HC', 'Oakland': 'OAKL', 'Montana': 'MONT',
    }
    
    if team_name in abbrev_map:
        return abbrev_map[team_name]
    
    # Fallback
    words = team_name.replace('University of ', '').replace('College of ', '').split()
    if len(words) == 1:
        return words[0][:4].upper()
    if len(words) <= 3:
        return ''.join(w[0] for w in words).upper()
    return words[0][:4].upper()

def flip_spread_if_needed(market, away_team, home_team, away_cover, home_cover):
    """Flip spread to show higher cover % team's perspective - EXACT copy from skill"""
    if not market or market == 'N/A' or not away_cover or not home_cover:
        return 'N/A'
    
    if isinstance(market, str):
        return market
    
    away_pct = float(away_cover.replace('%', ''))
    home_pct = float(home_cover.replace('%', ''))
    
    if away_pct == home_pct:
        return market['display']
    
    original_abbrev = market['original_abbrev']
    spread_value = float(market['value'])
    
    is_away_higher = away_pct > home_pct
    away_abbrev = derive_abbreviation(away_team)
    home_abbrev = derive_abbreviation(home_team)
    
    def abbrev_matches(abbrev1, abbrev2):
        a1, a2 = abbrev1.upper(), abbrev2.upper()
        if a1 == a2:
            return True
        if a1.startswith(a2) or a2.startswith(a1):
            return True
        return False
    
    original_refers_to_away = abbrev_matches(original_abbrev, away_abbrev)
    
    if is_away_higher:
        if original_refers_to_away:
            return f"{away_abbrev} {spread_value:+g}"
        else:
            return f"{away_abbrev} {-spread_value:+g}"
    else:
        if not original_refers_to_away:
            return f"{home_abbrev} {spread_value:+g}"
        else:
            return f"{home_abbrev} {-spread_value:+g}"

def parse_espn_schedule_from_text(text):
    """Parse ESPN schedule - returns spread as dictionary"""
    lines = text.strip().split('\n')
    format_type = detect_schedule_format(lines)
    
    if format_type == 'mobile':
        return parse_mobile_format(lines)
    else:
        return parse_desktop_format(lines)

def detect_schedule_format(lines):
    """Detect mobile vs desktop format"""
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
    """Parse desktop format - RETURNS SPREAD AS DICT"""
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
                            spread = {
                                'original_abbrev': team_abbrev,
                                'value': spread_value,
                                'display': f"{team_abbrev} {spread_value}"
                            }
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
                games.append({
                    'Away': away_team,
                    'Home': home_team,
                    'Time': time,
                    'Market': spread if spread else 'N/A'
                })
            
            i = j
        else:
            i += 1
    
    return games

def parse_mobile_format(lines):
    """Parse mobile format - RETURNS SPREAD AS DICT"""
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
                                spread = {
                                    'original_abbrev': team_abbrev,
                                    'value': spread_value,
                                    'display': f"{team_abbrev} {spread_value}"
                                }
                            break
                        j += 1
                    
                    games.append({
                        'Away': away_team,
                        'Home': home_team,
                        'Time': time,
                        'Market': spread if spread else 'N/A'
                    })
            
            i += 6
        else:
            i += 1
    
    return games

def load_ats_data_from_text(text):
    """Parse ATS data from TeamRankings"""
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
    """Find cover % with name mapping"""
    if team_name in ats_dict:
        return ats_dict[team_name]['cover_pct']
    if team_name in name_mapping:
        mapped_name = name_mapping[team_name]
        if mapped_name in ats_dict:
            return ats_dict[mapped_name]['cover_pct']
    return None

def find_team_ats_plus_minus(team_name, ats_dict, name_mapping):
    """Find ATS +/- with name mapping"""
    if team_name in ats_dict:
        return ats_dict[team_name]['ats_pm']
    if team_name in name_mapping:
        mapped_name = name_mapping[team_name]
        if mapped_name in ats_dict:
            return ats_dict[mapped_name]['ats_pm']
    return None

def create_daily_chart(games, ats_dict, name_mapping):
    """Create daily chart with flipped spreads and track unmapped teams"""
    chart_rows = []
    unmapped_teams = set()
    
    for game in games:
        away_team = game['Away']
        home_team = game['Home']
        
        away_cover = find_team_cover_pct(away_team, ats_dict, name_mapping)
        home_cover = find_team_cover_pct(home_team, ats_dict, name_mapping)
        away_ats_pm = find_team_ats_plus_minus(away_team, ats_dict, name_mapping)
        home_ats_pm = find_team_ats_plus_minus(home_team, ats_dict, name_mapping)
        
        # Track unmapped teams
        if not away_cover:
            unmapped_teams.add(away_team)
        if not home_cover:
            unmapped_teams.add(home_team)
        
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
    
    return chart_rows, list(unmapped_teams)

def create_xlsx_file(chart_rows, filename):
    """Create XLSX with color coding"""
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

# Ensure static directory exists
os.makedirs('static', exist_ok=True)

def cleanup_old_files(directory='static', hours=24):
    """Delete files older than specified hours"""
    try:
        now = datetime.now()
        deleted_count = 0
        for filename in os.listdir(directory):
            if filename.endswith('.xlsx'):
                filepath = os.path.join(directory, filename)
                file_modified = datetime.fromtimestamp(os.path.getmtime(filepath))
                if (now - file_modified).total_seconds() > hours * 3600:
                    os.remove(filepath)
                    deleted_count += 1
        if deleted_count > 0:
            logger.info(f"Cleanup: deleted {deleted_count} old file(s)")
    except Exception as e:
        logger.error(f"Error during file cleanup: {e}")

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        logger.info("Chart generation request received")
        
        # Clean up old files first
        cleanup_old_files()
        
        espn_schedule = request.form.get('espn_schedule', '')
        teamrankings_ats = request.form.get('teamrankings_ats', '')
        
        # Input validation
        if not espn_schedule or not teamrankings_ats:
            logger.warning("Request missing ESPN or TeamRankings data")
            flash('Please provide both ESPN schedule and TeamRankings ATS data', 'error')
            return redirect(url_for('index'))
        
        # Size limits: 500KB per input (way more than needed for legitimate use)
        MAX_INPUT_SIZE = 500000  # 500KB in bytes
        if len(espn_schedule) > MAX_INPUT_SIZE:
            logger.warning(f"ESPN data too large: {len(espn_schedule)} characters")
            flash(f'ESPN schedule data is too large ({len(espn_schedule):,} characters). Maximum allowed is {MAX_INPUT_SIZE:,} characters.', 'error')
            return redirect(url_for('index'))
        
        if len(teamrankings_ats) > MAX_INPUT_SIZE:
            logger.warning(f"TeamRankings data too large: {len(teamrankings_ats)} characters")
            flash(f'TeamRankings ATS data is too large ({len(teamrankings_ats):,} characters). Maximum allowed is {MAX_INPUT_SIZE:,} characters.', 'error')
            return redirect(url_for('index'))
        
        try:
            # Parse ESPN schedule
            try:
                games = parse_espn_schedule_from_text(espn_schedule)
                if not games:
                    logger.error("ESPN schedule parsing returned no games")
                    flash('Could not parse any games from ESPN schedule. Make sure you copied from the date through the last Gamecast button.', 'error')
                    return redirect(url_for('index'))
                logger.info(f"Successfully parsed {len(games)} games from ESPN schedule")
            except Exception as e:
                logger.error(f"ESPN parsing error: {str(e)}")
                flash(f'ESPN parsing error: {str(e)}. Please check your ESPN schedule format.', 'error')
                return redirect(url_for('index'))
            
            # Parse TeamRankings data
            try:
                ats_dict = load_ats_data_from_text(teamrankings_ats)
                if not ats_dict:
                    logger.error("TeamRankings parsing returned no data")
                    flash('Could not parse ATS data from TeamRankings. Make sure you copied the entire table.', 'error')
                    return redirect(url_for('index'))
                logger.info(f"Successfully parsed {len(ats_dict)} teams from TeamRankings")
            except Exception as e:
                logger.error(f"TeamRankings parsing error: {str(e)}")
                flash(f'TeamRankings parsing error: {str(e)}. Please check your ATS data format.', 'error')
                return redirect(url_for('index'))
            
            # Create chart and get unmapped teams
            try:
                chart_rows, unmapped_teams = create_daily_chart(games, ats_dict, TEAM_NAME_MAPPING)
                logger.info(f"Chart created with {len(chart_rows)} rows")
            except Exception as e:
                logger.error(f"Chart generation error: {str(e)}")
                flash(f'Chart generation error: {str(e)}', 'error')
                return redirect(url_for('index'))
            
            # Create Excel file
            try:
                output_filename = f'daily_chart_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
                output_path = os.path.join('static', output_filename)
                create_xlsx_file(chart_rows, output_path)
                logger.info(f"Excel file created: {output_filename}")
            except Exception as e:
                logger.error(f"File creation error: {str(e)}")
                flash(f'File creation error: {str(e)}', 'error')
                return redirect(url_for('index'))
            
            # Show success message
            flash(f'Successfully generated chart with {len(chart_rows)} games!', 'success')
            logger.info(f"Chart generation successful - {len(chart_rows)} games, {len(ats_dict)} teams tracked")
            
            # Warn about unmapped teams
            if unmapped_teams:
                teams_list = ', '.join(sorted(unmapped_teams))
                logger.warning(f"Unmapped teams ({len(unmapped_teams)}): {teams_list}")
                flash(f'Warning: Could not find ATS data for {len(unmapped_teams)} team(s): {teams_list}', 'error')
            
            return render_template('index.html', 
                                 chart_rows=chart_rows, 
                                 download_file=output_filename, 
                                 games_count=len(games), 
                                 teams_count=len(ats_dict),
                                 unmapped_teams=unmapped_teams)
        
        except Exception as e:
            logger.error(f"Unexpected error: {str(e)}")
            flash(f'Unexpected error: {str(e)}', 'error')
            return redirect(url_for('index'))
    
    return render_template('index.html')

@app.route('/download/<filename>')
def download(filename):
    """Download generated file"""
    return send_file(os.path.join('static', filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
