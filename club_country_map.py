import json, sys
from collections import Counter, defaultdict
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

with open('clubs_new.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

club_counter = Counter()
for team, players in data.items():
    for player, club in players.items():
        if club:
            club_counter[club.strip()] += 1

CLUB_COUNTRY = {
    # England
    'Manchester City': 'England', 'Arsenal': 'England', 'Crystal Palace': 'England',
    'Manchester United': 'England', 'Liverpool': 'England', 'Sunderland': 'England',
    'Aston Villa': 'England', 'Chelsea': 'England', 'Fulham': 'England',
    'Newcastle': 'England', 'Brighton': 'England', 'Tottenham Hotspur': 'England',
    'Wolves': 'England', 'Wolverhampton Wanderers': 'England', 'West Ham': 'England',
    'Middlesbrough': 'England', 'Southampton': 'England', 'Ipswich Town': 'England',
    'Burnley': 'England', 'Norwich City': 'England', 'Nottingham Forest': 'England',
    'Leeds United': 'England', 'Brentford': 'England', 'Bournemouth': 'England',
    'Everton': 'England', 'Leicester City': 'England', 'Watford': 'England',
    'Hull City': 'England', 'Derby County': 'England', 'Coventry City': 'England',
    'Swansea City': 'England', 'Stoke City': 'England', 'Stoke': 'England',
    'Birmingham': 'England', 'Portsmouth': 'England', 'Swansea': 'England',
    'Wrexham': 'England', 'Wrexham AFC': 'England', 'Sheffield United': 'England',
    'Rotherham United': 'England', 'Millwall FC': 'England', 'Peterborough United': 'England',
    'Port Vale': 'England', 'Charlton Athletic': 'England', 'West Ham United': 'England',
    'Sunderland AFC': 'England', 'Nottingham Forest F.C.': 'England',
    'AFC Bournemouth': 'England', 'Brighton & Hove Albion': 'England',
    'Newcastle United': 'England', 'Wolverhampton': 'England', 'Spurs': 'England',
    'Braintree Town': 'England',
    # Germany
    'Bayern Munich': 'Germany', 'Borussia Dortmund': 'Germany', 'Bayer Leverkusen': 'Germany',
    'Eintracht Frankfurt': 'Germany', 'RB Leipzig': 'Germany', 'Stuttgart': 'Germany',
    'Hoffenheim': 'Germany', 'Borussia Mönchengladbach': 'Germany', 'Schalke 04': 'Germany',
    'VfB Stuttgart': 'Germany', 'SC Freiburg': 'Germany', 'Mainz 05': 'Germany',
    'Mainz': 'Germany', 'Augsburg': 'Germany', 'VfL Wolfsburg': 'Germany',
    'Wolfsburg': 'Germany', 'St Pauli': 'Germany', 'FC St. Pauli': 'Germany',
    'Werder Bremen': 'Germany', 'Hannover 96': 'Germany', 'TSG Hoffenheim': 'Germany',
    'Borussia Monchengladbach': 'Germany', 'FSV Mainz': 'Germany',
    'Karlsruher SC': 'Germany', 'Hamburger SV': 'Germany', 'Fortuna Düsseldorf': 'Germany',
    'Bayern München': 'Germany', 'Holstein Kiel': 'Germany', 'Union Berlin': 'Germany',
    'Frankfurt': 'Germany', 'TSG 1899 Hoffenheim': 'Germany',
    'Hamburg': 'Germany', 'HSV': 'Germany', 'FC Cosmos Koblenz': 'Germany',
    # France
    'PSG': 'France', 'Lille': 'France', 'Strasbourg': 'France', 'OGC Nice': 'France',
    'Lens': 'France', 'Olympique de Marseille': 'France', 'Olympique Lyonnais': 'France',
    'Monaco': 'France', 'Lyon': 'France', 'AS Monaco': 'France', 'Rennes': 'France',
    'Auxerre': 'France', 'AJ Auxerre': 'France', 'Lorient': 'France',
    'Marseille': 'France', 'Stade Rennais': 'France', 'SC Bastia': 'France',
    'FC Sochaux': 'France', 'AS Nancy-Lorraine': 'France', 'Angers SCO': 'France',
    'Toulouse': 'France', 'Reims': 'France', 'Paris St Germain': 'France',
    'Olympique Marseille': 'France', 'Olympique Marseilla': 'France',
    'Saint-Étienne': 'France', 'Nantes': 'France', 'Le Havre': 'France',
    'Le Havre AC': 'France', 'Paris FC': 'France', 'Montpellier': 'France',
    'Saint-Etienne': 'France', 'RC Strasbourg': 'France', 'Nice': 'France',
    # Spain
    'Barcelona': 'Spain', 'Atlético de Madrid': 'Spain', 'Real Madrid': 'Spain',
    'Real Betis': 'Spain', 'Villarreal': 'Spain', 'Real Sociedad': 'Spain',
    'Athletic Club': 'Spain', 'Athletic Bilbao': 'Spain', 'Sevilla': 'Spain', 'Celta de Vigo': 'Spain',
    'Osasuna': 'Spain', 'Rayo Vallecano': 'Spain', 'Villarreal CF': 'Spain',
    'FC Barcelona': 'Spain', 'FC Sevilla': 'Spain', 'Real Oviedo': 'Spain',
    'Valencia': 'Spain', 'Cultural Leonesa': 'Spain',
    'Girona': 'Spain', 'Elche': 'Spain', 'Castellón': 'Spain',
    'Granada': 'Spain', 'Mallorca': 'Spain', 'RCD Mallorca': 'Spain',
    'Atlético Madrid': 'Spain', 'Levante': 'Spain', 'Espanyol': 'Spain',
    'Athletic Club': 'Spain',
    # Italy
    'AC Milan': 'Italy', 'Juventus': 'Italy', 'Atalanta': 'Italy',
    'Internazionale': 'Italy', 'AS Roma': 'Italy', 'Torino': 'Italy',
    'Napoli': 'Italy', 'Inter Milan': 'Italy', 'Venezia': 'Italy',
    'Bologna': 'Italy', 'Inter': 'Italy', 'Atalanta BC': 'Italy', 'Genoa': 'Italy',
    'Cremonese': 'Italy', 'Sassuolo': 'Italy', 'Sassuolo Calcio': 'Italy',
    'US Sassuolo': 'Italy', 'UC Sampdoria': 'Italy', 'Pisa': 'Italy', 'Parma': 'Italy',
    'Udinese': 'Italy', 'Como': 'Italy', 'Como 1907': 'Italy', 'Frosinone': 'Italy',
    'Cagliari': 'Italy', 'Fiorentina': 'Italy', 'Hellas Verona': 'Italy', 'Roma': 'Italy',
    # Netherlands
    'PSV Eindhoven': 'Netherlands', 'PSV': 'Netherlands', 'Feyenoord': 'Netherlands',
    'Ajax': 'Netherlands', 'PEC Zwolle': 'Netherlands', 'SC Telstar': 'Netherlands',
    'RKC Waalwijk': 'Netherlands', 'FC Volendam': 'Netherlands',
    'Almere City FC': 'Netherlands', 'NEC Nijmegen': 'Netherlands',
    'VVV-Venlo': 'Netherlands', 'Sparta Rotterdam': 'Netherlands',
    'FC Den Bosch': 'Netherlands', 'Heracles Almelo': 'Netherlands',
    'Nijmegen': 'Netherlands', 'Twente': 'Netherlands', 'Utrecht': 'Netherlands',
    'Ajax Amsterdam': 'Netherlands', 'Feynoord': 'Netherlands',
    # Portugal
    'Sporting CP': 'Portugal', 'Benfica': 'Portugal', 'SL Benfica': 'Portugal',
    'FC Porto': 'Portugal', 'Braga': 'Portugal', 'SC Braga': 'Portugal',
    'Porto': 'Portugal', 'FC Vizela': 'Portugal', 'Gil Vicente': 'Portugal',
    'Chaves': 'Portugal', 'Torreense': 'Portugal', 'Vitória de Guimarães': 'Portugal',
    'Farense': 'Portugal', 'Estrela Amadora': 'Portugal', 'Casa Pia': 'Portugal',
    'Sporting': 'Portugal', 'CD Tondela': 'Portugal',
    # Saudi Arabia
    'Al-Hilal': 'Saudi Arabia', 'Al-Ahli': 'Saudi Arabia', 'Al-Nassr': 'Saudi Arabia',
    'Al-Qadsiah': 'Saudi Arabia', 'Al-Ittihad': 'Saudi Arabia', 'Al-Ettifaq': 'Saudi Arabia',
    'Al Hilal': 'Saudi Arabia', 'Abha Club': 'Saudi Arabia', 'Al-Shabab': 'Saudi Arabia',
    'Al-Fayha FC': 'Saudi Arabia', 'Al-Ula': 'Saudi Arabia', 'Neom': 'Saudi Arabia',
    'Al Qadsiah': 'Saudi Arabia',
    # Egypt
    'Al Ahly': 'Egypt', 'Pyramids': 'Egypt', 'Zamalek': 'Egypt', 'El Gouna': 'Egypt',
    'Enppi': 'Egypt', 'Zed': 'Egypt', 'Al Najma': 'Egypt',
    # Qatar
    'Al-Duhail': 'Qatar', 'Al-Sadd': 'Qatar', 'Al Rayyan': 'Qatar',
    'Al Gharafa': 'Qatar', 'Al-Wakrah': 'Qatar', 'Al Wakrah': 'Qatar',
    'Al Arabi': 'Qatar', 'Al-Sailiya': 'Qatar', 'Qatar SC': 'Qatar', 'Al Shamal': 'Qatar',
    # UAE
    'Al Ain': 'UAE', 'Baniyas': 'UAE', 'Al Jazira': 'UAE', 'Al-Dhafra': 'UAE',
    'Kalba': 'UAE', 'Dibba': 'UAE', 'Shabab Al-Ahli': 'UAE', 'Al Bataeh': 'UAE',
    'Al-Samiya': 'Jordan', 'Al-Wahda': 'UAE', 'Al-Ain': 'UAE',
    # Turkey
    'Fenerbahçe A.Ş.': 'Turkey', 'Galatasaray A.Ş.': 'Turkey', 'Fenerbahçe': 'Turkey',
    'Galatasaray': 'Turkey', 'Besiktas': 'Turkey', 'Istanbul Basaksehir': 'Turkey',
    'Çaykur Rizespor': 'Turkey', 'Beşiktaş A.Ş.': 'Turkey', 'Kasımpaşa A.Ş.': 'Turkey',
    'Trabzonspor': 'Turkey', 'Konyaspor': 'Turkey', 'Kayserispor': 'Turkey',
    'Trabzonspor Kulübü': 'Turkey', 'İstanbul Başakşehir': 'Turkey',
    'Çaykur Rizespor A.Ş.': 'Turkey', 'Beşiktaş J.K.': 'Turkey',
    'Galatasaray S.K.': 'Turkey', 'Galatasaray SK': 'Turkey', 'Samsunspor': 'Turkey',
    'Fenerbahce SK': 'Turkey', 'Alanyaspor': 'Turkey', 'Kasimpasa': 'Turkey',
    'Iğdır': 'Turkey',
    # Iran
    'Persepolis': 'Iran', 'Tractor': 'Iran', 'Esteghlal': 'Iran',
    'Esteghlal FC': 'Iran', 'Sepahan': 'Iran', 'Malavan': 'Iran', 'Foolad': 'Iran',
    # Morocco
    'RS Berkane': 'Morocco', 'AS FAR': 'Morocco', 'Raja Casablanca': 'Morocco',
    # Algeria
    'Kabylie': 'Algeria', 'USM Alger': 'Algeria', 'Alger': 'Algeria',
    # Tunisia
    'Club Africain': 'Tunisia', 'Etoile Sahel': 'Tunisia', 'CS Sfaxien': 'Tunisia',
    'Esperance': 'Tunisia', 'US Monastir': 'Tunisia', 'Espérance de Tunis': 'Tunisia',
    # South Africa
    'Mamelodi Sundowns': 'South Africa', 'Orlando Pirates': 'South Africa',
    'Siwelele FC': 'South Africa', 'Kaizer Chiefs': 'South Africa',
    'Polokwane City': 'South Africa',
    # Brazil
    'Flamengo': 'Brazil', 'Palmeiras': 'Brazil', 'Atlético Mineiro': 'Brazil',
    'Grêmio': 'Brazil', 'Internacional': 'Brazil', 'Botafogo': 'Brazil',
    'Santos': 'Brazil', 'Corinthians': 'Brazil', 'Sao Paulo': 'Brazil',
    'Red Bull Bragantino': 'Brazil', 'Fluminense': 'Brazil', 'Vasco da Gama': 'Brazil',
    'Esporte Clube Bahia': 'Brazil',
    # Argentina
    'River Plate': 'Argentina', 'Independiente': 'Argentina', 'Huracán': 'Argentina',
    'San Lorenzo': 'Argentina', 'Boca Juniors': 'Argentina', 'Estudiantes': 'Argentina',
    'Lanus': 'Argentina', 'Talleres': 'Argentina', 'Independiente Rivadavia': 'Argentina',
    'Racing Club': 'Argentina', 'Rosario Central': 'Argentina', 'Velez Sarsfield': 'Argentina',
    # Mexico
    'Chivas': 'Mexico', 'Club América': 'Mexico', 'Toluca': 'Mexico',
    'Cruz Azul': 'Mexico', 'Club Tijuana': 'Mexico', 'Pumas UNAM': 'Mexico',
    'UNAM Pumas': 'Mexico', 'Tigres UANL': 'Mexico', 'Tigres': 'Mexico',
    'Monterrey': 'Mexico', 'Santos Laguna': 'Mexico', 'Atlas': 'Mexico',
    'Pachuca': 'Mexico', 'Club Leon FC': 'Mexico', 'FC Juarez': 'Mexico',
    'Mazatlan FC': 'Mexico', 'Pumas': 'Mexico', 'CF América': 'Mexico',
    # USA
    'Chicago Fire': 'USA', 'Inter Miami': 'USA', 'Orlando City': 'USA',
    'Philadelphia Union': 'USA', 'FC Dallas': 'USA', 'Columbus Crew': 'USA',
    'Atlanta United': 'USA', 'New York City': 'USA', 'Minnesota United': 'USA',
    'LAFC': 'USA', 'Los Angeles FC': 'USA', 'Colorado Springs': 'USA',
    'Colorado Rapids': 'USA', 'Nashville SC': 'USA', 'Portland Timbers': 'USA',
    'San Diego': 'USA', 'San Diego FC': 'USA', 'Real Salt Lake City': 'USA',
    'Seattle Sounders': 'USA', 'Charlotte FC': 'USA', 'FC Cincinnati': 'USA',
    'El Paso Locomotive FC': 'USA', 'New England Revolution': 'USA',
    'New York City FC': 'USA', 'Miami FC': 'USA', 'Minnesota United FC': 'USA',
    'Columbus': 'USA', 'Dallas': 'USA',
    # Canada
    'Toronto': 'Canada', 'Vancouver Whitecaps': 'Canada', 'Toronto FC': 'Canada',
    # Australia
    'Melbourne City': 'Australia', 'Western Sydney Wanderers': 'Australia',
    'Sydney': 'Australia', 'Newcastle Jets': 'Australia',
    # New Zealand
    'Auckland FC': 'New Zealand', 'Wellington Phoenix': 'New Zealand',
    # Scotland
    'Celtic': 'Scotland', 'Rangers': 'Scotland', 'Hearts': 'Scotland',
    'Hibernian': 'Scotland', 'Heart of Midlothian': 'Scotland',
    'Kilmarnock': 'Scotland', 'Motherwell FC': 'Scotland',
    # Sweden
    'AIK Fotball AB': 'Sweden', 'Malmö': 'Sweden', 'Norrkoping': 'Sweden',
    'IFK Norrköping': 'Sweden', 'Mjallby': 'Sweden',
    # Norway
    'Sarpsborg': 'Norway', 'Viking FK': 'Norway', 'Viking': 'Norway',
    'Bodø/Glimt': 'Norway', 'Molde FK': 'Norway',
    # Denmark
    'Midtjylland': 'Denmark', 'Bröndby IF': 'Denmark', 'Randers': 'Denmark',
    'AGF': 'Denmark', 'Silkeborg': 'Denmark', 'FC Copenhagen': 'Denmark',
    'Copenhagen': 'Denmark', 'FC Nordsjaelland': 'Denmark',
    'Nordsjaelland': 'Denmark', 'København': 'Denmark',
    # Switzerland
    'BSC Young Boys': 'Switzerland', 'FC Lugano': 'Switzerland',
    'FC Zürich': 'Switzerland', 'Servette Geneva': 'Switzerland',
    'Stade Nyonnais': 'Switzerland', 'Young Boys Berne': 'Switzerland',
    'Lugano': 'Switzerland', 'St. Gallen': 'Switzerland', 'BSC Young Boys II': 'Switzerland',
    # Austria
    'RB Salzburg': 'Austria', 'Austria Wien': 'Austria', 'LASK': 'Austria',
    'Wolfsberger AC': 'Austria', 'Grazer': 'Austria',
    # Belgium
    'Club Brugge': 'Belgium', 'Anderlecht': 'Belgium', 'Genk': 'Belgium',
    'KAA Gent': 'Belgium', 'Standard Liege': 'Belgium', 'Sint-Truiden': 'Belgium',
    'Royale Union Saint-Gilloise': 'Belgium', 'Union Saint-Gilloise': 'Belgium',
    'SK Beveren': 'Belgium', 'Royal Charleroi SC': 'Belgium', 'Sporting Charleroi': 'Belgium',
    'Royal Charleroi': 'Belgium', 'Cercle Brugge KSV': 'Belgium', 'Dender': 'Belgium',
    'KRC Genk': 'Belgium', 'Mechelen': 'Belgium', 'SV Zulte Waregem': 'Belgium',
    'Royal Antwerp': 'Belgium', 'Saint-Gilloise': 'Belgium',
    # Serbia
    'Crvena Zvezda': 'Serbia',
    # Croatia
    'HNK Rijeka': 'Croatia', 'Dinamo Zagreb': 'Croatia', 'Hajduk Split': 'Croatia',
    'Lokomotiva Zagreb': 'Croatia', 'Slaven Belupo': 'Croatia', 'Rijeka': 'Croatia',
    # Czech Republic
    'Slavia Prague': 'Czech Republic', 'Viktoria Plzen': 'Czech Republic',
    'Sparta Prague': 'Czech Republic', 'Hradec Kralove': 'Czech Republic',
    'FC Slovan Liberec': 'Czech Republic', 'FC Viktoria Plzen': 'Czech Republic',
    'Viktoria Plzeň': 'Czech Republic',
    # Greece
    'PAOK': 'Greece', 'Olympiacos': 'Greece', 'AEK Athens': 'Greece',
    'Larissa': 'Greece', 'Panathinaikos': 'Greece', 'Panathinaikos F.C.': 'Greece',
    'Atromitos': 'Greece', 'AE Kifisia': 'Greece', 'AE Kifisias': 'Greece',
    # Cyprus
    'AEL Limassol': 'Cyprus', 'Pafos': 'Cyprus', 'Pafos FC': 'Cyprus',
    'APOEL': 'Cyprus', 'Omonia': 'Cyprus', 'AEK Larnaca': 'Cyprus',
    'Apollon Limassol': 'Cyprus',
    # Israel
    'Maccabi Haifa': 'Israel', 'Maccabi Tel Aviv': 'Israel', 'Ironi Kiryat Shmona': 'Israel',
    # Russia
    'Dynamo Moscow': 'Russia', 'Krasnodar': 'Russia', 'Zenit': 'Russia',
    'Lokomotiv Moscow': 'Russia', 'Spartak Moscow': 'Russia',
    'Dynamo Makhachkala': 'Russia', 'Rostov': 'Russia',
    'FC Pari Nizhniy Novgorod': 'Russia', 'Akron Tolyatti': 'Russia',
    'Dinamo Moscow': 'Russia',
    # Japan
    'FC Tokyo': 'Japan', 'Kashima Antlers': 'Japan', 'Albirex Niigata': 'Japan',
    'Sanfrecce Hiroshima': 'Japan',
    # South Korea
    'Jeonbuk': 'South Korea', 'Ulsan': 'South Korea', 'Gangwon': 'South Korea',
    'Daejeon': 'South Korea', 'Daejeon Citizen': 'South Korea', 'Seoul': 'South Korea',
    # Malaysia
    'Selangor': 'Malaysia', 'Terengganu FC': 'Malaysia',
    # Indonesia
    'Persib Bandung': 'Indonesia',
    # Uzbekistan
    'Pakhtakor': 'Uzbekistan', 'Nasaf': 'Uzbekistan', 'Navbahor': 'Uzbekistan',
    'Surkhon': 'Uzbekistan', 'Dinamo Samarqand': 'Uzbekistan',
    'AGMK': 'Uzbekistan', 'Bukhara': 'Uzbekistan',
    # Azerbaijan
    'Neftchi': 'Azerbaijan', 'Turan Tovuz': 'Azerbaijan',
    # Kazakhstan
    'FC Astana': 'Kazakhstan',
    # Armenia
    'FC Noah': 'Armenia',
    # Iraq
    'Al-Zawraa': 'Iraq', 'Al-Shorta': 'Iraq', 'Al-Karma': 'Iraq',
    'Al-Quwa Al-Jawiya': 'Iraq', 'Al Talaba': 'Iraq', 'Al-Talaba': 'Iraq',
    'Al-Nasr': 'Iraq',
    # Jordan
    'Al-Faisaly': 'Jordan', 'Al-Wehdat': 'Jordan', 'Al-Hussein': 'Jordan',
    # Ghana
    'Accra Hearts of Oak': 'Ghana',
    # Romania
    'FC Universitatea Cluj': 'Romania', 'FCSB': 'Romania',
    # Hungary
    'Ferencváros TC': 'Hungary', 'Puskás Akadémia': 'Hungary', 'ETO': 'Hungary',
    # Poland
    'Jagiellonia Bialystok': 'Poland', 'Lechia Gdańsk': 'Poland',
    'Widzew Łódź': 'Poland', 'Cracovia': 'Poland', 'Pogoń Szczecin': 'Poland',
    # Slovakia
    'FC Tatran Prešov': 'Slovakia', 'Slovan Bratislava': 'Slovakia',
    # Bulgaria
    'Ludogorets Razgrad': 'Bulgaria',
    # Ireland
    'Shamrock Rovers': 'Ireland', "St. Patrick's Athletic": 'Ireland',
    "St. Patrick’s Athletic": 'Ireland',
    # Finland
    'SJK': 'Finland',
    # Ecuador
    'LDU Quito': 'Ecuador',
    # Costa Rica
    'Deportivo Saprissa': 'Costa Rica',
    # Panama
    'CD Plaza Amador': 'Panama',
    # Venezuela
    'Deportivo La Guaira': 'Venezuela', 'Academia Puerto Cabello': 'Venezuela',
    # Paraguay
    'Cerro Porteño': 'Paraguay', 'Olimpia': 'Paraguay',
    # Uruguay
    'Club Nacional': 'Uruguay',
    # Chile
    'CD Cobresal': 'Chile', 'CD Universidad Catholica': 'Chile',
    'Universidad de Concepcion': 'Chile', 'CD Universidad Catolica': 'Chile',
    # Colombia
    'Atletico Nacional': 'Colombia',
    # Haiti
    'Violette AC': 'Haiti',
    # Honduras
    'CD Marathon': 'Honduras',
    # Thailand
    'Port': 'Thailand',
    # China
    'Zhejiang FC': 'China',
    # Slovenia
    'NK Maribor': 'Slovenia',
    # Extra entries / alternate spellings
    'Freiburg': 'Germany', 'Al Nassr': 'Saudi Arabia', 'Gaziantep FK': 'Turkey',
    'FC ⁠Augsburg': 'Germany',
    # Extra Colombia clubs
    'Racing de Santander': 'Spain',
    'Athletico Paranaense': 'Brazil',
    # Corrected entries
    'Girone': 'Spain',                     # Girona FC (misspelling in source data)
    'Montana': 'Bulgaria',                 # Montana FC
    'Machida Zelvia': 'Japan',

    # Standardiserte navn (etter standardiser_klubber.py 2026-06-18)
    'Al-Arabi':                  'Qatar',
    'Al-Gharafa':                'Qatar',
    'Al-Shamal':                 'Qatar',
    'Al Fateh':                  'Saudi Arabia',
    'Angers':                    'France',
    'Beşiktaş':                  'Turkey',
    'Boavista':                  'Portugal',
    'Bochum':                    'Germany',
    'Borac Banja Luka':          'Bosnia and Herzegovina',
    'Colorado Springs Switchbacks': 'USA',
    'El Paso Locomotive':        'USA',
    'Fagiano Okayama':           'Japan',
    'Gimcheon Sangmu':           'South Korea',
    'Mafra':                     'Portugal',
    'Seongnam':                  'South Korea',
    'Sturm Graz':                'Austria',
    'Suwon':                     'South Korea',
    'Ulsan HD':                  'South Korea',
    'Vissel Kobe':               'Japan',
    'Austin FC':                 'USA',
}

country_counter = defaultdict(int)
unknown = []
for club, count in club_counter.items():
    country = CLUB_COUNTRY.get(club)
    if country:
        country_counter[country] += count
    else:
        unknown.append((club, count))

print('UNKNOWN clubs:')
for c, n in sorted(unknown, key=lambda x: -x[1]):
    print(f'  {n}\t{repr(c)}')
print(f'\nTotal unknown players: {sum(n for _,n in unknown)}')
print(f'Total mapped players: {sum(country_counter.values())}')
