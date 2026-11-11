"""
Module: Geography, Physiography, Soils, Hydrology & Canals (Slides 7 - 18)
Author: Academic Research Presentation Engine (ARPE)
Target: Haryana FSO & State Screening Tests (HPSC / HSSC)
"""

from pptx.util import Inches

def build_geography_slides(builder):
    prs = builder.prs
    blank = builder.blank_layout

    # -------------------------------------------------------------------------
    # Slide 7: Geographical Profile & Location
    # -------------------------------------------------------------------------
    s7 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s7)
    builder._add_header_and_footer(
        s7,
        tag="GEOGRAPHICAL METRICS",
        title="Geographical Coordinates, Territorial Extent & Border States",
        takeaway="Haryana covers 44,212 sq km (1.34% of India) and is landlocked by 4 States & 2 UTs.",
        citation="Survey of India; Statistical Abstract of Haryana 2023-24.",
        slide_num=7
    )
    c1 = [
        "Latitudinal Extent: 27°39' N to 30°55' N latitude (Span: ~3°16').",
        "Longitudinal Extent: 74°27' E to 77°36' E longitude (Span: ~3°09').",
        "Total Geographical Area: 44,212 sq km (17,070 sq miles), accounting for 1.34% of India's total land area.",
        "National Area Rank: 20th among Indian States (21st including UT of Ladakh/J&K reorganisation).",
        "Elevation Profile: Ranging from 200 m (656 ft) in southwestern depressions to 1,514 m in the northern Morni Hills."
    ]
    builder._create_card(s7, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Coordinates & Physical Extent", c1, builder.COLOR_COBALT)

    c2 = [
        "Bordering States (4 States): Punjab (North-West), Himachal Pradesh (North), Uttar Pradesh (East), and Rajasthan (South & South-West).",
        "Bordering Union Territories (2 UTs): New Delhi (surrounded on three sides: North, West, South) and Chandigarh (Joint Capital in North).",
        "Longest Inter-State Border: Rajasthan shares the longest boundary (~1,262 km) with 7 Haryana districts (Sirsa, Fatehabad, Hisar, Bhiwani, Mahendragarh, Rewari, Nuh).",
        "Shortest Inter-State Border: Himachal Pradesh shares ~11 km border with Panchkula, Ambala, and Yamunanagar.",
        "Landlocked Nature: Completely enclosed with zero international border and no coastal outlet.",
        "Districts with No Inter-State Boundary: Rohtak and Charkhi Dadri (touch only internal Haryana districts)."
    ]
    builder._create_card(s7, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Boundaries & Border Dynamics", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 8: Physiographic Classification (Dr. Jasbir Singh)
    # -------------------------------------------------------------------------
    s8 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s8)
    builder._add_header_and_footer(
        s8,
        tag="PHYSIOGRAPHIC ZONING",
        title="Physiography of Haryana: Dr. Jasbir Singh's 8-Fold Classification",
        takeaway="In 'Agricultural Geography of Haryana', Dr. Jasbir Singh classified the terrain into 8 distinct relief units.",
        citation="Dr. Jasbir Singh, 'An Agricultural Geography of Haryana' (Vishal Publications); HPSC Screening Benchmark.",
        slide_num=8
    )
    headers = ["Physiographic Unit", "Geographical Coverage", "Topographical Nature", "Key Soil / Ecological Feature"]
    rows = [
        ["1. Shivalik Hills", "Panchkula, Northern Ambala & Yamunanagar", "Tertiary folded hills; altitude 400m - 1514m", "Subject to severe gully erosion; Pinjore gravel zone"],
        ["2. Piedmont Plain (Ghar/Kandi)", "Foothills of Shivaliks (Panchkula, Ambala)", "Piedmont alluvial fans formed by seasonal torrents", "Coarse sand, boulders; dissected by 'Chos' (torrents)"],
        ["3. Alluvial Plain (Bangar)", "Central Haryana (Kaithal, Karnal, Jind, Rohtak)", "Flat, older alluvium plain; elevation 220m - 250m", "Highly fertile loam; heart of Green Revolution"],
        ["4. Flood Plains (Khadar/Naili/Bet)", "Yamuna, Ghaggar, and Markanda flood belts", "Low-lying newer alluvium formed by seasonal overflows", "Khadar (Yamuna), Naili (Ghaggar), Bet (Markanda)"],
        ["5. Sandy Plains & Sand Dunes", "Sirsa, Fatehabad, Hisar, Bhiwani, Mahendragarh", "Aeolian undulating plains with shifting sand dunes (Tibas)", "Semi-arid; high wind erosion; calcareous dry soils"],
        ["6. Aravalli Outliers", "Gurugram, Mewat (Nuh), Rewari, Mahendragarh", "Denuded Precambrian residual metamorphic ridges", "Quartzite, slate, granite outcrops; rocky thin soils"],
        ["7. Drifting Sand Zone (Taal)", "Southwestern desert transition belt (Bhiwani/Sirsa)", "Shallow depressions between sand dunes", "Saline water accumulation during episodic rain"],
        ["8. Marshy / Tarai Tract", "Southern Sirsa (Sanisa depression / seasonal marshes)", "Low-lying water accumulating basins", "Heavy clayey soils; prone to temporary waterlogging"]
    ]
    widths = [Inches(2.5), Inches(3.2), Inches(3.2), Inches(2.833)]
    builder._create_table(s8, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 9: Shivalik Range & Morni Hills (Northern Topography)
    # -------------------------------------------------------------------------
    s9 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s9)
    builder._add_header_and_footer(
        s9,
        tag="NORTHERN HIGHLANDS",
        title="The Shivalik Range, Morni Hills & Karoh Peak Profile",
        takeaway="Karoh Peak (1,514m) in Morni Hills, Panchkula is the highest elevation point of Haryana.",
        citation="Geological Survey of India (GSI) Northern Region; Department of Environment & Climate Change, Haryana.",
        slide_num=9
    )
    c1 = [
        "Geological Origin: Outer-most Himalayan foothills formed during late Cenozoic era by fluvial deposits folded during Himalayan orogeny.",
        "Hypsometric Peak: Karoh Peak (करोह चोटी) situated in Morni Hills, Panchkula district at an elevation of 1,514 metres (4,967 ft) - widely asked in HPSC/HSSC.",
        "Morni Hills Ecosystem: The only hill station of Haryana, home to Tikkar Taal (twin lakes: Chhota Tikkar and Bada Tikkar separated by hillock).",
        "Drainage Origin: Cradle of seasonal foothill streams: Tangri, Ghaggar tributaries, and Markanda torrents.",
        "Vegetation: Sub-tropical chir pine (Pinus roxburghii), mixed deciduous forests, and rich wild medicinal flora."
    ]
    builder._create_card(s9, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Shivalik Relief & Morni Hills", c1, builder.COLOR_COBALT)

    c2 = [
        "Piedmont Plain (Ghar & Kandi): Narrow alluvial strip (10-15 km wide) lying immediately south of Shivaliks across Panchkula, Ambala, and Yamunanagar.",
        "Local Terminology Trap: Known locally as 'Ghar' (घर) in Kalka/Raipur Rani and 'Kandi' (कांडी) in hilly margins.",
        "Torrents ('Chos'): Cut deep ravines and seasonal dry beds called Chos that cause suddenflash floods during monsoon.",
        "Chhachhrauli Phenomenon: Located in Yamunanagar district; receives highest average annual rainfall in the state (>1100 mm), famously dubbed 'The Cherrapunji of Haryana'.",
        "Geological Assets: Boulder gravel, sandstone beds, and lime mortar kankar deposits extensively quarried for regional construction."
    ]
    builder._create_card(s9, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Ghar, Kandi & The Cherrapunji of Haryana", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 10: The Aravalli Outliers & Dhosi Hill (Southern Topography)
    # -------------------------------------------------------------------------
    s10 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s10)
    builder._add_header_and_footer(
        s10,
        tag="SOUTHERN RELIEF",
        title="Aravalli Outliers, Residual Metamorphic Ridges & Dhosi Hill",
        takeaway="Dhosi Hill (652m) in Mahendragarh is the highest peak of Aravallis in Haryana and the hermitage of Rishi Chyavana.",
        citation="Archaeological Survey of India (ASI); Gazetteers of Gurgaon & Mahendragarh.",
        slide_num=10
    )
    c1 = [
        "Precambrian Metamorphic Antiquity: One of the oldest fold mountain chains in the world (Delhi Supergroup: Alwar and Ajabgarh series).",
        "Geographical Spread: Discontinuous relict hills entering through Rajasthan into Mahendragarh, Rewari, Gurugram, Nuh (Mewat), and Faridabad.",
        "Highest Elevation Peak: Dhosi Hill (धोसी की पहाड़ी) at 652 metres (2,139 ft) situated at the border of Mahendragarh (near Narnaul) and Jhunjhunu (Rajasthan).",
        "Extinct Volcanic Crater: An isolated volcanic plug of late Proterozoic era with a crater-like depression at the summit.",
        "Ecological Shield: Acts as a critical green natural barrier preventing eastward expansion of the Thar Desert into Indo-Gangetic plains."
    ]
    builder._create_card(s10, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Dhosi Hill & Aravalli Ridges", c1, builder.COLOR_GOLD)

    c2 = [
        "Vedic Hermitage of Chyavana Rishi: Ancient seat where sage Chyavana (son of sage Bhrigu) formulated the legendary Ayurvedic preparation 'Chyawanprash' using wild Amla and herbs.",
        "Surajkund Sun Temple & Amphitheatre: Built by Tomar King Anangpal I in 10th century CE in the Aravalli foothills of Faridabad.",
        "Kotla & Khalilpur Depressions: Natural drainage basins at the foot of Mewat Aravalli hills collecting seasonal runoff.",
        "Mining & Conservation Ban: Supreme Court of India banned commercial mining in Aravalli ridges across Faridabad and Gurugram to restore groundwater aquifers.",
        "Aravalli Green Wall Project: Ambitious national 1,400 km long and 5 km wide afforestation green belt spanning Haryana, Rajasthan, Gujarat, and Delhi."
    ]
    builder._create_card(s10, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Historical, Medicinal & Ecological Role", c2, builder.COLOR_PURPLE)

    # -------------------------------------------------------------------------
    # Slide 11: Soil Classification & Agro-Ecological Zones
    # -------------------------------------------------------------------------
    s11 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s11)
    builder._add_header_and_footer(
        s11,
        tag="SOIL PEDOLOGY",
        title="Soil Classification, Texture Profile & Agro-Ecological Zones",
        takeaway="Soils vary from light sandy loams in the south-west to fertile alluvial loams in central plains and rocky loams in Shivaliks.",
        citation="ICAR-Central Soil Salinity Research Institute (CSSRI), Karnal; Agriculture Department of Haryana.",
        slide_num=11
    )
    c1 = [
        "Very Light Soils (Sandy): Found in Sirsa, Fatehabad, Bhiwani, and Mahendragarh. High permeability, deficient in nitrogen and organic matter, prone to wind erosion.",
        "Light Soils (Sandy Loam): Spans dry parts of Hisar, Rewari, and Jhajjar. Highly responsive to tubewell/canal irrigation and phosphatic fertilizers.",
        "Medium Soils (Loam & Silt Loam): Dominant in Kaithal, Jind, Sonipat, Panipat, and Rohtak. The backbone of wheat, basmati paddy, and sugarcane cropping.",
        "Moderately Heavy & Heavy Soils (Clayey): Trapped along Ghaggar basin (Naili tract) and parts of Yamunanagar and Karnal; rich in nutrients but slow drainage."
    ]
    builder._create_card(s11, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Soil Textural Classes", c1, builder.COLOR_COBALT)

    c2 = [
        "Agro-Climatic Zone I (Eastern / Wet Zone): Panchkula, Ambala, Yamunanagar, Kurukshetra, Karnal, Kaithal, Sonipat, Panipat. Sub-humid to semi-arid, high irrigation, paddy-wheat rotation.",
        "Agro-Climatic Zone II (Western / Dry Zone): Sirsa, Fatehabad, Hisar, Bhiwani, Charkhi Dadri, Mahendragarh, Rewari, Jhajjar. Arid to semi-arid, mustard-pearl millet-cotton cultivation.",
        "Organic Carbon Status: Generally low to medium (<0.5%) due to intensive double-cropping and crop residue burning.",
        "Micronutrient Deficiencies: Zinc (Zn) and Iron (Fe) widespread across central plains; Sulphur (S) deficit pronounced in oilseed mustard belts."
    ]
    builder._create_card(s11, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Agro-Ecological Distribution", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 12: Soil Salinity, Waterlogging & CSSRI Karnal
    # -------------------------------------------------------------------------
    s12 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s12)
    builder._add_header_and_footer(
        s12,
        tag="FSO PEDOLOGY SPECIAL",
        title="Soil Salinity, Alkalinity (Kallar / Rehi) & CSSRI Reclamation",
        takeaway="CSSRI Karnal pioneered gypsum application and sub-surface drainage (SSD) to reclaim sodic soils.",
        citation="ICAR-Central Soil Salinity Research Institute (CSSRI), Karnal (Est. 1969); Haryana Reclamation Corporation.",
        slide_num=12
    )
    headers = ["Diagnostic Property", "Saline Soils (Solonchak)", "Alkali / Sodic Soils (Kallar/Rehi)", "CSSRI Reclamation Technology"]
    rows = [
        ["pH Range", "< 8.5 (Neutral to mildly alkaline)", "> 8.5 (Highly alkaline, up to 10.5)", "Gypsum (CaSO4.2H2O) requirement based on Soil Testing"],
        ["Electrical Conductivity (EC)", "> 4.0 dS/m (High soluble salts)", "< 4.0 dS/m (Low neutral soluble salts)", "Leaching with good quality canal water through deep furrows"],
        ["Exchangeable Sodium (ESP)", "< 15% (Low sodium saturation)", "> 15% (Excessive exchangeable Na+)", "Sub-Surface Drainage (SSD) with corrugated perforated PVC pipes"],
        ["Physical Soil Structure", "Flocculated, good aeration, permeable", "Dispersed, crusting, impermeable hard pan", "Growing salt-tolerant crops: CSR-30 Basmati, KRL-210 Wheat"],
        ["Major Affected Districts", "Rohtak, Jhajjar, Hisar, Jind, Sonipat", "Karnal, Kurukshetra, Kaithal, Panipat", "Bio-drainage using Eucalyptus plantation along canal seepage"]
    ]
    widths = [Inches(2.5), Inches(3.0), Inches(3.0), Inches(3.233)]
    builder._create_table(s12, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 13: River Systems & Drainage Divide
    # -------------------------------------------------------------------------
    s13 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s13)
    builder._add_header_and_footer(
        s13,
        tag="HYDROLOGICAL DIVIDE",
        title="River Systems of Haryana: Northern & Southern Drainage Networks",
        takeaway="Haryana forms the great hydrological water divide between the Indus Basin and the Ganga Basin.",
        citation="Central Ground Water Board (CGWB); Irrigation and Water Resources Department, Haryana.",
        slide_num=13
    )
    c1 = [
        "Yamuna River: The only perennial river; marks the entire eastern border with Uttar Pradesh; drains into the Bay of Bengal (Ganga System).",
        "Ghaggar River: Main seasonal river of northern Haryana; originates in Dagshai (HP); passes into Rajasthan desert (Indus / Inland System).",
        "Markanda River: Ancient Aruna; originates in Paonta/Dharti Dhar hills (HP); famous for seasonal fury and heavy siltation.",
        "Tangri (Dangri) River: Originates in Morni Hills; joins Markanda near Mullana (Ambala).",
        "Saraswati River: Ancient sacred river; originating near Adi Badri (Yamunanagar); revived by the Haryana Saraswati Heritage Development Board."
    ]
    builder._create_card(s13, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Northern Drainage System (Shivalik Origin)", c1, builder.COLOR_COBALT)

    c2 = [
        "Sahibi River: Chief river of southern Haryana; originates in Kotputli/Bairath hills (Rajasthan); flows through Rewari and Gurugram into Najafgarh drain.",
        "Krishnavati (Kasaunti) River: Originates in Neem Ka Thana (Rajasthan); enters Mahendragarh; historically joins Sahibi.",
        "Dohan River: Originates near Mandoli (Rajasthan); mythological origin from the eye of Rishi Bhrigu's wife Divyapoloma; seasonal dry stream in Mahendragarh.",
        "Indori River: Originates in Nuh hills near old fort of Indore; divides into two branches near Rewari and meets Sahibi.",
        "Inland Drainage: Southern streams disappear into the sandy tracts of Mahendragarh and Rewari without reaching any sea."
    ]
    builder._create_card(s13, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Southern Drainage System (Aravalli Origin)", c2, builder.COLOR_GOLD)

    # -------------------------------------------------------------------------
    # Slide 14: Perennial Lifeline: The Yamuna River
    # -------------------------------------------------------------------------
    s14 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s14)
    builder._add_header_and_footer(
        s14,
        tag="PERENNIAL WATERWAY",
        title="The Yamuna River: Entry, Course, Barrages & Basin Dynamics",
        takeaway="Yamuna enters Haryana at Kalesar (Yamunanagar) and exits at Hasanpur (Palwal), spanning ~320 km.",
        citation="Yamuna Basin Organization (Central Water Commission); Haryana Irrigation Department.",
        slide_num=14
    )
    c1 = [
        "Glacial Source: Originates from Yamunotri Glacier at Bandarpunch peak in Uttarkashi district (Uttarakhand) at ~6,387 m elevation.",
        "Point of Entry: Enters Haryana through Kalesar National Park forest in Yamunanagar district.",
        "Natural Boundary Line: Flows continuously southward defining the ~320 km interstate border between Haryana (Yamunanagar, Karnal, Panipat, Sonipat, Faridabad, Palwal) and Uttar Pradesh.",
        "Point of Exit: Leaves Haryana near Hasanpur village in Palwal district to flow into Mathura/Agra (UP).",
        "Floodplain (Khadar): Creates fertile flood-deposited newer alluvium belt across eastern Yamunanagar and Karnal."
    ]
    builder._create_card(s14, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "River Course & Demarcation", c1, builder.COLOR_COBALT)

    c2 = [
        "Tajewala Barrage: Historical stone diversion structure constructed by the British in 1873 to divert water into Western and Eastern Yamuna Canals.",
        "Hathnikund Barrage: Modern concrete barrage constructed in 1999 (operational 2002) 3 km upstream of Tajewala to handle super-floods (>8 lakh cusecs).",
        "Okhla Barrage: Located in Delhi/Faridabad border; feeds Agra Canal and Gurgaon Canal providing irrigation to southern districts.",
        "Kishau, Renuka & Lakhwar Dams: Upstream storage projects on tributaries (Tons, Giri, Yamuna) designed to ensure lean-season flow to Haryana.",
        "FSO Ecological Note: High industrial effluent contamination from Sonipat-Panipat-Gurugram drains (Dhanauri, Najafgarh) impacting downstream irrigation water."
    ]
    builder._create_card(s14, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Barrages, Diversions & Projects", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 15: Ghaggar, Markanda & Saraswati Rivers
    # -------------------------------------------------------------------------
    s15 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s15)
    builder._add_header_and_footer(
        s15,
        tag="NORTHERN HYDROLOGY",
        title="Ghaggar, Markanda (Aruna) & The Saraswati Heritage Project",
        takeaway="Ghaggar creates the 'Naili' flood plain; Markanda merges into Saraswati at Sanisa Lake near Pehowa.",
        citation="Haryana Saraswati Heritage Development Board (HSHDB); State Remote Sensing Application Centre (HARSAC).",
        slide_num=15
    )
    headers = ["River Parameter", "Ghaggar River", "Markanda River (Aruna)", "Saraswati River Project"]
    rows = [
        ["Origin Point", "Dagshai hills (Himachal Pradesh) at ~1,927m", "Dharti Dhar / Paonta hills (HP)", "Adi Badri foothills (Yamunanagar district)"],
        ["Entry into Haryana", "Near Kalka (Panchkula district)", "Near Ambala district border", "Adi Badri / Kathgarh cultural corridor"],
        ["Districts Traversed", "Panchkula, Ambala, Kaithal, Fatehabad, Sirsa", "Ambala and Kurukshetra", "Yamunanagar, Kurukshetra, Kaithal"],
        ["Major Tributaries", "Tangri, Markanda, Kaushalya (dam at Pinjore)", "Runn, Begna, Nakti, and Tangri river", "Chautang (ancient Drishadvati) historically"],
        ["Key Reservoirs / Sinks", "Ottu Barrage / Weir (Sirsa); dries in Thar Desert", "Sanisa Lake (near Pehowa / Kurukshetra)", "Palaeochannel excavations by HSHDB & ONGC"],
        ["Exam Terminology", "Flood plain called 'Naili'; non-perennial", "Known as 'Aruna' in Mahabharata & Puranas", "Mughalwali (Yamunanagar) site of active water revival"]
    ]
    widths = [Inches(2.5), Inches(3.2), Inches(3.0), Inches(3.033)]
    builder._create_table(s15, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 16: Southern Drainage: Sahibi, Krishnavati, Dohan & Indori
    # -------------------------------------------------------------------------
    s16 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s16)
    builder._add_header_and_footer(
        s16,
        tag="SOUTHERN HYDROLOGY",
        title="Southern River Systems: Sahibi, Krishnavati, Dohan & Indori",
        takeaway="Sahibi is the largest southern river, entering near Rewari and discharging into Najafgarh Jheel.",
        citation="Groundwater Directorate of Haryana; Narnaul & Rewari District Gazetteers.",
        slide_num=16
    )
    c1 = [
        "Chief Southern Artery: Originates in Bairath / Sewar hills in Jaipur district, Rajasthan.",
        "Entry into Haryana: Enters Rewari district near Kotkasim border, flows through Pataudi (Gurugram) and Jhajjar.",
        "Discharge Sink: Empties into Najafgarh drain (historical Najafgarh lake) in Southwest Delhi.",
        "Masani Barrage: Built across Sahibi river in 1989 near Masani village (Rewari) on NH-48 for flood moderation and groundwater recharge.",
        "Sub-Basin Significance: Recharges vital deep aquifers in the water-stressed arid Rewari-Gurugram industrial belt."
    ]
    builder._create_card(s16, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Sahibi River & Masani Barrage", c1, builder.COLOR_COBALT)

    c2 = [
        "Dohan River: Mythologically linked to Rishi Bhrigu's hermitage at Dhosi Hill; originates near Mandoli village; flows northward into Mahendragarh before getting lost in shifting sands.",
        "Krishnavati (Kasaunti) River: Originates from Dariba hills in Neem Ka Thana (Rajasthan); enters south of Narnaul; flows past Mahendragarh to disappear in Jhajjar depressions.",
        "Indori River: Takes birth near ancient ruined fort of Indore in Nuh (Mewat) hills; splits into two courses near Malab, one joining Sahibi near Rewari.",
        "Ephemeral Nature: All four streams are completely non-perennial and flow only during extreme cloudbursts in monsoon season.",
        "Groundwater Depletion Alert: Entire basin categorized as 'Over-exploited' / 'Dark Zone' by CGWB due to extensive tubewell extraction."
    ]
    builder._create_card(s16, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Dohan, Krishnavati & Indori Rivers", c2, builder.COLOR_GOLD)

    # -------------------------------------------------------------------------
    # Slide 17: Canal Irrigation: Western Yamuna & Bhakra Networks
    # -------------------------------------------------------------------------
    s17 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s17)
    builder._add_header_and_footer(
        s17,
        tag="IRRIGATION LIFELINES",
        title="Canal Irrigation Networks: Western Yamuna Canal & Bhakra System",
        takeaway="WYC is Haryana's oldest canal (1355 CE); Tohana in Fatehabad is renowned as the 'City of Canals'.",
        citation="Irrigation & Water Resources Department, Haryana; Bhakra Beas Management Board (BBMB).",
        slide_num=17
    )
    c1 = [
        "Historical Antiquity: Originally excavated by Sultan Firoz Shah Tughlaq in 1355 CE to supply water to hunting lodges in Hisar.",
        "Mughal Restoration: Reconditioned by Akbar (1568) and remodeled as Nahr-i-Bihisht by Shah Jahan under Ali Mardan Khan to bring water to Delhi.",
        "British Modernization: Fully rebuilt in 1879 originating from Tajewala Barrage (now Hathnikund).",
        "Branches & Command: Main Branch divides into Delhi Branch, Hansi Branch, Bhalaut Branch, Barwala Branch, and Sunder Branch.",
        "Command Districts: Irrigates Yamunanagar, Karnal, Kaithal, Rohtak, Sonipat, Panipat, and Jind."
    ]
    builder._create_card(s17, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Western Yamuna Canal (WYC)", c1, builder.COLOR_COBALT)

    c2 = [
        "Bhakra Nangal Canal System: Originates from Nangal Dam on Sutlej river; enters Haryana at Tohana (Fatehabad district) - known as 'Canal City'.",
        "Main Branches: Narwana Branch and Barwala Link Canal connecting Bhakra waters to Western Yamuna Canal system.",
        "Jawaharlal Nehru (JLN) Lift Canal: Multi-stage lift irrigation system engineered under CM Bansi Lal to pump water uphill to arid sandy terrains of Mahendragarh and Rewari.",
        "Loharu (Indira Gandhi) & Siwani Lift Canals: Lift canal systems designed to irrigate sand dune tracts of Bhiwani and Charkhi Dadri.",
        "SYL Canal Dispute: 214 km Sutlej-Yamuna Link canal (122 km in Punjab, 92 km in Haryana) constructed by Haryana but blocked by Punjab."
    ]
    builder._create_card(s17, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Bhakra System & Lift Irrigation Schemes", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 18: Lakes, Reservoirs & Sacred Sarovars
    # -------------------------------------------------------------------------
    s18 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s18)
    builder._add_header_and_footer(
        s18,
        tag="AQUATIC GEOGRAPHY",
        title="Prominent Lakes, Wetland Reservoirs & Sacred Sarovars",
        takeaway="Damdama is the largest natural lake; Brahma Sarovar is the sacred epicenter of Solar Eclipse pilgrimages.",
        citation="Haryana Tourism Corporation; Archaeological Survey of India (ASI); Kurukshetra Development Board (KDB).",
        slide_num=18
    )
    headers = ["Water Body / Sarovar", "District Location", "Origin / Structural Type", "Key Exam Fact / Screening Point"]
    rows = [
        ["Damdama Lake", "Sohna (Gurugram)", "Natural depression in Aravalli hills (Est. 1947)", "Largest natural lake in Haryana (~3,000 acres); fed by rain runoff"],
        ["Badkhal Lake", "Faridabad", "Man-made bundh constructed in 1947", "Built under CM Bansi Lal / Nehru for agricultural irrigation; name from 'Be-dakhal'"],
        ["Brahma Sarovar", "Thanesar (Kurukshetra)", "Ancient sacred tank mentioned by Al-Biruni in 'Kitab-ul-Hind'", "Largest man-made bathing tank in Asia (1800m x 1400m); Solar Eclipse dip"],
        ["Sannihit Sarovar", "Kurukshetra", "Meeting point of 7 sacred Saraswatis", "Believed to hold waters of all sacred teerthas on Amavasya; Lord Vishnu seat"],
        ["Karna Lake & Chakravarty", "Karnal (NH-44)", "Artificial lake named after warrior Karna", "Major roadside tourism hub of Haryana Tourism; Uchana complex"],
        ["Tilyar Lake", "Rohtak", "Spread over 132 acres with Zoo", "Major tourist complex and recreation hub of central Haryana"],
        ["Bhindawas & Khaparwas", "Jhajjar district", "Freshwater wetland lake", "Bhindawas is Haryana's largest wetland lake and designated Ramsar Site (2021)"],
        ["Blue Bird Lake", "Hisar", "Artificial recreation lake & wetland", "Famous migratory bird wetland near deer park and civil aerodrome"]
    ]
    widths = [Inches(2.5), Inches(2.2), Inches(3.2), Inches(3.833)]
    builder._create_table(s18, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    print("Geography slides (7-18) built successfully.")
