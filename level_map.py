# Club → (level, league, country) for clubs NOT in top division
# Level 1 = top division (handled by TOP_DIVISION_CLUBS)

# English → Norwegian country name translations
LAND_NO = {
    'Algeria':       'Algerie',
    'Austria':       'Østerrike',
    'Azerbaijan':    'Aserbajdsjan',
    'Belgium':       'Belgia',
    'Brazil':        'Brasil',
    'China':         'Kina',
    'Croatia':       'Kroatia',
    'Cyprus':        'Kypros',
    'Czech Republic':'Tsjekkia',
    'Denmark':       'Danmark',
    'France':        'Frankrike',
    'Germany':       'Tyskland',
    'Greece':        'Hellas',
    'Hungary':       'Ungarn',
    'Iraq':          'Irak',
    'Ireland':       'Irland',
    'Italy':         'Italia',
    'Kazakhstan':    'Kasakhstan',
    'Mexico':        'Mexico',
    'Morocco':       'Marokko',
    'Netherlands':   'Nederland',
    'New Zealand':   'Ny-Zealand',
    'Norway':        'Norge',
    'Poland':        'Polen',
    'Romania':       'Romania',
    'Russia':        'Russland',
    'Saudi Arabia':  'Saudi-Arabia',
    'Scotland':      'Skottland',
    'South Africa':  'Sør-Afrika',
    'South Korea':   'Sør-Korea',
    'Spain':         'Spania',
    'Sweden':        'Sverige',
    'Switzerland':   'Sveits',
    'Turkey':        'Tyrkia',
    'Uzbekistan':    'Usbekistan',
    'Unknown':       'Ukjent',
    # Manglende oversettelser
    'Cape Verde':               'Kapp Verde',
    'Ivory Coast':              'Elfenbenskysten',
    'Bosnia and Herzegovina':   'Bosnia og Hercegovina',
    'United States':            'USA',
    'United States of America': 'USA',
    'Republic of Korea':        'Sør-Korea',
    'DR Congo':                 'DR Kongo',
    'Czechia':                  'Tsjekkia',
    'Curacao':                  'Curaçao',
}
LEVEL_MAP = {
    # England – Championship (L2)
    'Ipswich Town':         (2, 'Championship',         'England'),
    'Leicester City':       (2, 'Championship',         'England'),
    'Southampton':          (2, 'Championship',         'England'),
    'Birmingham':           (2, 'Championship',         'England'),
    'Charlton Athletic':    (2, 'Championship',         'England'),
    'Coventry City':        (2, 'Championship',         'England'),
    'Derby County':         (2, 'Championship',         'England'),
    'Hull City':            (2, 'Championship',         'England'),
    'Middlesbrough':        (2, 'Championship',         'England'),
    'Millwall FC':          (2, 'Championship',         'England'),
    'Norwich City':         (2, 'Championship',         'England'),
    'Portsmouth':           (2, 'Championship',         'England'),
    'Sheffield United':     (2, 'Championship',         'England'),
    'Stoke':                (2, 'Championship',         'England'),
    'Stoke City':           (2, 'Championship',         'England'),
    'Swansea':              (2, 'Championship',         'England'),
    'Swansea City':         (2, 'Championship',         'England'),
    'Watford':              (2, 'Championship',         'England'),
    'Wrexham':              (2, 'Championship',         'England'),
    'Wrexham AFC':          (2, 'Championship',         'England'),
    # England – League One (L3)
    'Peterborough United':  (3, 'League One',           'England'),
    'Port Vale':            (3, 'League One',           'England'),
    'Rotherham United':     (3, 'League One',           'England'),
    # England – National League South (L6)
    'Braintree Town':       (6, 'National League South','England'),

    # Germany – 2. Bundesliga (L2)
    'Holstein Kiel':        (2, '2. Bundesliga',        'Germany'),
    'Fortuna Düsseldorf':   (2, '2. Bundesliga',        'Germany'),
    'Hannover 96':          (2, '2. Bundesliga',        'Germany'),
    'Karlsruher SC':        (2, '2. Bundesliga',        'Germany'),
    'Schalke 04':           (2, '2. Bundesliga',        'Germany'),
    'Bochum':               (2, '2. Bundesliga',        'Germany'),
    'VfL Bochum':           (2, '2. Bundesliga',        'Germany'),
    # Germany – Regionalliga (L4)
    'FC Cosmos Koblenz':    (4, 'Regionalliga',         'Germany'),

    # France – Ligue 2 (L2)
    'Montpellier':          (2, 'Ligue 2',              'France'),
    'Saint-Etienne':        (2, 'Ligue 2',              'France'),
    'Saint-Étienne':        (2, 'Ligue 2',              'France'),
    'Reims':                (2, 'Ligue 2',              'France'),
    'SC Bastia':            (2, 'Ligue 2',              'France'),
    'AS Nancy-Lorraine':    (2, 'Ligue 2',              'France'),
    # France – National (L3)
    'FC Sochaux':           (3, 'National',             'France'),

    # Spain – Segunda División (L2)
    'Granada':              (2, 'Segunda División',     'Spain'),
    'Racing de Santander':  (2, 'Segunda División',     'Spain'),
    'Castellón':            (2, 'Segunda División',     'Spain'),
    'Real Oviedo':          (2, 'Segunda División',     'Spain'),
    'Cultural Leonesa':     (2, 'Segunda División',     'Spain'),

    # Italy – Serie B (L2)
    'Venezia':              (2, 'Serie B',              'Italy'),
    'Frosinone':            (2, 'Serie B',              'Italy'),
    'UC Sampdoria':         (2, 'Serie B',              'Italy'),

    # Netherlands – Eerste Divisie (L2)
    'Almere City FC':       (2, 'Eerste Divisie',       'Netherlands'),
    'RKC Waalwijk':         (2, 'Eerste Divisie',       'Netherlands'),
    'FC Den Bosch':         (2, 'Eerste Divisie',       'Netherlands'),
    'VVV-Venlo':            (2, 'Eerste Divisie',       'Netherlands'),

    # Portugal – Liga Portugal 2 (L2)
    'Farense':              (2, 'Liga Portugal 2',      'Portugal'),
    'FC Vizela':            (2, 'Liga Portugal 2',      'Portugal'),
    'Chaves':               (2, 'Liga Portugal 2',      'Portugal'),
    'Torreense':            (2, 'Liga Portugal 2',      'Portugal'),
    # Portugal – Liga 3 (L3)
    'Mafra':                (3, 'Liga 3',               'Portugal'),
    'CD Mafra':             (3, 'Liga 3',               'Portugal'),

    # Switzerland – Challenge League (L2)
    'Stade Nyonnais':       (2, 'Challenge League',     'Switzerland'),
    # Switzerland – Promotion League (L3)
    'BSC Young Boys II':    (3, 'Promotion League',     'Switzerland'),


    # Sweden – Superettan (L2)
    'Norrkoping':           (2, 'Superettan',           'Sweden'),
    'IFK Norrkoping':       (2, 'Superettan',           'Sweden'),
    'IFK Norrköping':       (2, 'Superettan',           'Sweden'),

    # UAE – UAE Division 1 (L2)
    'Al Bataeh':            (2, 'UAE Division 1',       'UAE'),
    'Dibba':                (2, 'UAE Division 1',       'UAE'),   # Dibba Al-Fujairah

    # Saudi Arabia – First Division (L2)
    'Al-Ula':               (2, 'First Division',       'Saudi Arabia'),

    # Belgium – Challenger Pro League (L2)
    'SK Beveren':           (2, 'Challenger Pro League','Belgium'),

    # Turkey – TFF 1. Lig (L2)
    'Iğdır':                (2, 'TFF 1. Lig',           'Turkey'),

    # South Africa – National First Division (L2)
    'Siwelele FC':          (2, 'National First Division', 'South Africa'),

    # Japan – J2 League (L2)
    'Albirex Niigata':      (2, 'J2 League',            'Japan'),

    # USA – USL Championship (L2)
    'El Paso Locomotive FC':(2, 'USL Championship',     'USA'),
    'Miami FC':             (2, 'USL Championship',     'USA'),
    'Colorado Springs':     (2, 'USL Championship',     'USA'),
}
