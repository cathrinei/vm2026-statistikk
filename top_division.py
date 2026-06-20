# Clubs confirmed to play in the top division of their country (as of 2025/26 season)
TOP_DIVISION_CLUBS = {
    # England - Premier League
    'Manchester City', 'Arsenal', 'Crystal Palace', 'Manchester United', 'Liverpool',
    'Aston Villa', 'Chelsea', 'Fulham', 'Newcastle', 'Brighton', 'Tottenham Hotspur',
    'West Ham', 'Brentford', 'Bournemouth', 'Everton',
    'Wolverhampton Wanderers', 'Wolves',
    'West Ham United', 'Brighton & Hove Albion', 'Newcastle United',
    'Nottingham Forest', 'Nottingham Forest F.C.', 'AFC Bournemouth', 'Spurs', 'Burnley',
    'Leeds United', 'Sunderland', 'Sunderland AFC',  # promoted from Championship 2024/25

    # Germany - Bundesliga
    'Bayern Munich', 'Borussia Dortmund', 'Bayer Leverkusen', 'RB Leipzig',
    'Eintracht Frankfurt', 'VfB Stuttgart', 'Stuttgart', 'TSG Hoffenheim', 'Hoffenheim',
    'Borussia Mönchengladbach', 'Borussia Monchengladbach', 'SC Freiburg',
    'Mainz 05', 'Mainz', 'Augsburg', 'VfL Wolfsburg', 'Wolfsburg', 'Werder Bremen',
    'FC St. Pauli', 'St Pauli', 'Union Berlin', 'Bayern München',
    'TSG 1899 Hoffenheim', 'Hamburger SV', 'HSV', 'Hamburg',  # HSV promoted 2024/25

    # France - Ligue 1
    'PSG', 'Lille', 'OGC Nice', 'Lens', 'Monaco', 'AS Monaco', 'Rennes', 'Stade Rennais',
    'Auxerre', 'AJ Auxerre', 'Nantes', 'Marseille', 'Olympique de Marseille', 'Angers SCO',
    'Olympique Marseille', 'Olympique Marseilla', 'Lyon', 'Olympique Lyonnais',
    'Strasbourg', 'RC Strasbourg', 'Le Havre', 'Le Havre AC', 'Toulouse',
    'Lorient', 'Paris St Germain', 'Nice', 'Paris FC', 'Brest', 'Stade Brestois',

    # Spain - La Liga
    'Barcelona', 'Atlético de Madrid', 'Real Madrid', 'Real Betis', 'Villarreal',
    'Real Sociedad', 'Athletic Bilbao', 'Sevilla', 'Celta de Vigo', 'Osasuna',
    'Rayo Vallecano', 'Villarreal CF', 'FC Barcelona', 'FC Sevilla', 'Valencia',
    'Girona', 'Girone', 'Mallorca', 'RCD Mallorca', 'Atlético Madrid', 'Espanyol',
    'Athletic Club', 'Elche', 'Levante',

    # Italy - Serie A
    'AC Milan', 'Juventus', 'Atalanta', 'Internazionale', 'AS Roma', 'Torino',
    'Napoli', 'Inter Milan', 'Bologna', 'Inter', 'Atalanta BC', 'Genoa', 'Lazio', 'SS Lazio', 'Lecce',
    'Parma', 'Udinese', 'Como', 'Como 1907', 'Cagliari', 'Fiorentina',
    'Hellas Verona', 'Roma', 'Cremonese', 'Pisa', 'Sassuolo', 'Sassuolo Calcio', 'US Sassuolo',

    # Netherlands - Eredivisie
    'PSV Eindhoven', 'PSV', 'Feyenoord', 'Ajax', 'NEC Nijmegen',
    'Sparta Rotterdam', 'Heracles Almelo', 'Twente', 'Utrecht',
    'Ajax Amsterdam', 'Feynoord', 'PEC Zwolle', 'SC Telstar', 'FC Volendam',  # Volendam promoted 2025/26

    # Portugal - Primeira Liga
    'Sporting CP', 'Benfica', 'SL Benfica', 'FC Porto', 'Braga', 'SC Braga',
    'Porto', 'Gil Vicente', 'Vitória de Guimarães', 'Estrela Amadora',
    'Casa Pia', 'Sporting', 'CD Tondela', 'Boavista',

    # Saudi Arabia - Saudi Pro League
    'Al-Hilal', 'Al-Ahli', 'Al-Nassr', 'Al-Qadsiah', 'Al-Ittihad', 'Al-Ettifaq',
    'Al Hilal', 'Abha Club', 'Al-Shabab', 'Al-Fayha FC', 'Al Qadsiah', 'Al Nassr',
    'Al Najma', 'Neom', 'Al Fateh', 'Al-Fateh',

    # Egypt - Egyptian Premier League
    'Al Ahly', 'Pyramids', 'Zamalek', 'El Gouna', 'Enppi', 'Zed',

    # Qatar - Qatar Stars League
    'Al-Duhail', 'Al-Sadd', 'Al Rayyan', 'Al Gharafa', 'Al-Gharafa', 'Al-Wakrah', 'Al Wakrah',
    'Al Arabi', 'Al-Arabi', 'Al-Sailiya', 'Qatar SC', 'Al Shamal', 'Al-Shamal',

    # UAE - UAE Pro League
    'Al Ain', 'Baniyas', 'Al Jazira', 'Al-Dhafra', 'Shabab Al-Ahli', 'Al-Wahda',
    'Al-Ain', 'Kalba',  # Dibba & Al Bataeh relegated 2025-26

    # Turkey - Süper Lig
    'Fenerbahçe A.Ş.', 'Galatasaray A.Ş.', 'Fenerbahçe', 'Galatasaray', 'Besiktas',
    'Istanbul Basaksehir', 'Çaykur Rizespor', 'Beşiktaş A.Ş.', 'Kasımpaşa A.Ş.',
    'Trabzonspor', 'Konyaspor', 'Kayserispor', 'Trabzonspor Kulübü', 'İstanbul Başakşehir',
    'Çaykur Rizespor A.Ş.', 'Beşiktaş J.K.', 'Galatasaray S.K.', 'Galatasaray SK',
    'Samsunspor', 'Fenerbahce SK', 'Alanyaspor', 'Kasimpasa', 'Gaziantep FK',

    # Iran - Persian Gulf Pro League
    'Persepolis', 'Tractor', 'Esteghlal', 'Esteghlal FC', 'Sepahan', 'Foolad', 'Malavan',

    # Morocco - Botola Pro
    'RS Berkane', 'AS FAR', 'Raja Casablanca',

    # Algeria - Ligue Professionnelle 1
    'Kabylie', 'USM Alger', 'Alger',

    # Tunisia - Ligue 1
    'Club Africain', 'Etoile Sahel', 'CS Sfaxien', 'Esperance', 'US Monastir',
    'Espérance de Tunis',

    # South Africa - Premier Soccer League
    'Mamelodi Sundowns', 'Orlando Pirates', 'Kaizer Chiefs', 'Polokwane City',

    # Brazil - Série A
    'Flamengo', 'Palmeiras', 'Atlético Mineiro', 'Grêmio', 'Internacional', 'Botafogo',
    'Corinthians', 'Sao Paulo', 'Red Bull Bragantino', 'Fluminense',
    'Vasco da Gama', 'Esporte Clube Bahia', 'Santos',

    # Argentina - Primera División
    'River Plate', 'Independiente', 'Huracán', 'San Lorenzo', 'Boca Juniors',
    'Estudiantes', 'Lanus', 'Talleres', 'Independiente Rivadavia', 'Racing Club',
    'Rosario Central', 'Velez Sarsfield',

    # Mexico - Liga MX
    'Chivas', 'Club América', 'Toluca', 'Cruz Azul', 'Club Tijuana', 'Pumas UNAM',
    'UNAM Pumas', 'Tigres UANL', 'Tigres', 'Monterrey', 'Santos Laguna', 'Atlas',
    'Pachuca', 'Club Leon FC', 'FC Juarez', 'Mazatlan FC', 'Pumas', 'CF América',

    # USA - MLS
    'Chicago Fire', 'Inter Miami', 'Orlando City', 'Philadelphia Union', 'FC Dallas',
    'Columbus Crew', 'Atlanta United', 'New York City', 'Minnesota United', 'LAFC',
    'Los Angeles FC', 'Colorado Rapids', 'Nashville SC', 'Portland Timbers',
    'San Diego FC', 'Real Salt Lake City', 'Seattle Sounders', 'Charlotte FC',
    'FC Cincinnati', 'New England Revolution', 'New York City FC', 'Minnesota United FC',
    'Columbus', 'Dallas', 'Austin FC',  # Colorado Springs is USL Championship (L2)

    # Canada - Canadian MLS clubs
    'Toronto', 'Vancouver Whitecaps', 'Toronto FC',

    # Australia - A-League
    'Melbourne City', 'Western Sydney Wanderers', 'Sydney', 'Newcastle Jets',

    # New Zealand - A-League / NZ National League
    'Auckland FC', 'Wellington Phoenix',

    # Scotland - Scottish Premiership
    'Celtic', 'Rangers', 'Hearts', 'Hibernian', 'Heart of Midlothian', 'Kilmarnock',
    'Motherwell FC',

    # Sweden - Allsvenskan
    'AIK Fotball AB', 'Malmö', 'Mjallby',

    # Norway - Eliteserien
    'Sarpsborg', 'Viking FK', 'Viking', 'Bodø/Glimt', 'Molde FK',

    # Denmark - Superligaen
    'Midtjylland', 'Bröndby IF', 'AGF', 'Silkeborg', 'FC Copenhagen', 'Copenhagen',
    'FC Nordsjaelland', 'Nordsjaelland', 'København', 'Randers',

    # Switzerland - Super League
    'BSC Young Boys', 'FC Zürich', 'Servette Geneva', 'St. Gallen', 'FC Lugano', 'Lugano',

    # Austria - Bundesliga
    'RB Salzburg', 'Austria Wien', 'LASK', 'Wolfsberger AC', 'Grazer', 'Sturm Graz',

    # Belgium - Pro League / First Division A
    'Club Brugge', 'Anderlecht', 'Genk', 'KAA Gent', 'Standard Liege', 'Sint-Truiden',
    'Royale Union Saint-Gilloise', 'Union Saint-Gilloise', 'Royal Charleroi SC',
    'Sporting Charleroi', 'Royal Charleroi', 'Cercle Brugge KSV', 'Dender',
    'KRC Genk', 'Mechelen', 'Royal Antwerp', 'Saint-Gilloise', 'SV Zulte Waregem',

    # Serbia - SuperLiga
    'Crvena Zvezda',

    # Bosnia and Herzegovina - Premier liga BiH
    'Borac Banja Luka', 'FK Borac Banja Luka',

    # Croatia - HNL
    'HNK Rijeka', 'Dinamo Zagreb', 'Hajduk Split', 'Lokomotiva Zagreb', 'Rijeka',

    # Czech Republic - Czech First League
    'Slavia Prague', 'Sparta Prague', 'Viktoria Plzen', 'Viktoria Plzeň',
    'FC Slovan Liberec', 'FC Viktoria Plzen', 'Hradec Kralove', 'FC Tatran Prešov',

    # Greece - Super League
    'PAOK', 'Olympiacos', 'AEK Athens', 'Panathinaikos', 'Panathinaikos F.C.',
    'Atromitos', 'AE Kifisia', 'AE Kifisias', 'Larissa',

    # Cyprus - First Division
    'AEL Limassol', 'Pafos', 'Pafos FC', 'APOEL', 'Omonia', 'AEK Larnaca',
    'Apollon Limassol',

    # Israel - Israeli Premier League
    'Maccabi Haifa', 'Maccabi Tel Aviv', 'Ironi Kiryat Shmona',

    # Russia - Premier League
    'Dynamo Moscow', 'Krasnodar', 'Zenit', 'Lokomotiv Moscow', 'Spartak Moscow',
    'Rostov', 'Dinamo Moscow', 'FC Pari Nizhniy Novgorod', 'Akron Tolyatti',
    'Dynamo Makhachkala',

    # Japan - J1 League
    'FC Tokyo', 'Kashima Antlers', 'Sanfrecce Hiroshima',
    'Machida Zelvia', 'Fagiano Okayama', 'Vissel Kobe',

    # South Korea - K League 1
    'Jeonbuk', 'Ulsan', 'Ulsan HD', 'Gangwon', 'Seoul', 'Daejeon', 'Daejeon Hana Citizen', 'Daejeon Citizen',

    # Malaysia - Super League
    'Selangor', 'Terengganu FC',

    # Indonesia - Liga 1
    'Persib Bandung',

    # Uzbekistan - Super League
    'Pakhtakor', 'Nasaf', 'Navbahor', 'AGMK', 'Surkhon', 'Dinamo Samarqand', 'Bukhara',

    # Azerbaijan - Premier League
    'Neftchi', 'Turan Tovuz',

    # Kazakhstan - Premier League
    'FC Astana',

    # Armenia - Armenian Premier League
    'FC Noah',

    # Iraq - Iraqi Premier League
    'Al-Zawraa', 'Al-Shorta', 'Al-Quwa Al-Jawiya', 'Al Talaba', 'Al-Talaba', 'Al-Talaka',
    'Al-Karma',

    # Jordan - Jordan Premier League
    'Al-Hussein', 'Al-Wehdat', 'Al-Faisaly', 'Al-Nasr',
    'Al-Samiya',

    # Ghana - Ghana Premier League
    'Accra Hearts of Oak',

    # Romania - Liga 1
    'FCSB', 'FC Universitatea Cluj',

    # Hungary - OTP Bank Liga
    'Ferencváros TC', 'Puskás Akadémia', 'ETO',

    # Poland - Ekstraklasa
    'Jagiellonia Bialystok', 'Lechia Gdańsk', 'Widzew Łódź', 'Cracovia',
    'Pogoń Szczecin',

    # Slovakia - Super Liga
    'Slovan Bratislava',

    # Bulgaria - First League
    'Ludogorets Razgrad', 'Montana',

    # Slovenia - Prva liga
    'NK Maribor',

    # Ireland - League of Ireland Premier Division
    "Shamrock Rovers", "St. Patrick's Athletic",

    # Finland - Veikkausliiga
    'SJK',

    # Chile - Primera División
    'CD Cobresal', 'CD Universidad Catolica', 'CD Universidad Católica',
    'Universidad de Concepcion', 'Universidad de Concepción',

    # Ecuador - Serie A Ecuador
    'LDU Quito',

    # Panama - Liga Panameña de Fútbol
    'CD Plaza Amador',  # mestere Clausura 2026

    # Costa Rica - Primera División
    'Deportivo Saprissa',

    # Venezuela - Primera División
    'Deportivo La Guaira', 'Academia Puerto Cabello',

    # Paraguay - Primera División
    'Cerro Porteño', 'Olimpia',

    # Uruguay - Primera División
    'Club Nacional',

    # Colombia - Categoría Primera A
    'Atletico Nacional',

    # Haiti - Ligue Haïtienne
    'Violette AC',

    # Honduras - Liga Nacional
    'CD Marathon',

    # Thailand - Thai League 1
    'Port',

    # China - Chinese Super League
    'Zhejiang FC',
}
