"""
Module: Climate, Forest Cover, National Parks, Sanctuaries & Breeding Centres (Slides 19 - 24)
Author: Academic Research Presentation Engine (ARPE)
Target: Haryana FSO & State Screening Tests (HPSC / HSSC)
"""

from pptx.util import Inches

def build_environment_slides(builder):
    prs = builder.prs
    blank = builder.blank_layout

    # -------------------------------------------------------------------------
    # Slide 19: Climate, Seasons & Rainfall Isohyets
    # -------------------------------------------------------------------------
    s19 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s19)
    builder._add_header_and_footer(
        s19,
        tag="METEOROLOGICAL PROFILE",
        title="Climatic Zones, Temperature Extremes & Rainfall Isohyets",
        takeaway="Climate is sub-tropical continental monsoon; Chhachhrauli is the 'Cherrapunji of Haryana' (>1100 mm).",
        citation="India Meteorological Department (IMD) Chandigarh; State Climate Change Centre, Haryana.",
        slide_num=19
    )
    c1 = [
        "Classification: Sub-tropical continental, semi-arid to sub-humid monsoon climate with extreme continental variations.",
        "Temperature Extremes: Searing summer peaks (>46°C - 48°C) in Hisar, Bhiwani, Sirsa, and Narnaul; freezing winter lows (<1°C - 2°C) with occasional ground frost (Pala/पाला) damaging potato and mustard crops.",
        "Loo Winds (लू): Hot, dry, dust-laden western winds blowing across southwestern plains from May to June.",
        "Western Disturbances (Mawat / मावठ): Winter precipitation in January-February brought by Mediterranean depressions; critical lifeline for Rabi wheat and gram.",
        "HPSC Climate Question: Southwestern fringe borders Thar Desert; northeastern fringe borders humid sub-Himalayas."
    ]
    builder._create_card(s19, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Temperature & Seasonal Dynamics", c1, builder.COLOR_COBALT)

    c2 = [
        "Isohyet Gradient: Average annual rainfall declines sharply from ~1,100 mm in the northeast to <300 mm in the southwest.",
        "Highest Rainfall Belt: Panchkula, Ambala, and Yamunanagar. Chhachhrauli tehsil (Yamunanagar) records the highest rainfall in the state (>1100 mm), earning the title 'The Cherrapunji of Haryana'.",
        "Lowest Rainfall Belt: Sirsa, Bhiwani, and Mahendragarh (average <300 mm; frequent drought episodes).",
        "Monsoon Window: South-West Monsoon (Arabian Sea and Bay of Bengal branches) active from early July to September, delivering ~80% of total annual rain.",
        "Food Safety & Storage Implication: High humidity (>80%) during July-August in eastern granary districts (Karnal, Kurukshetra) accelerates aflatoxin (Aspergillus flavus) proliferation in stored grains."
    ]
    builder._create_card(s19, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Rainfall Distribution & Isohyet Map", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 20: Forest Cover Analysis (ISFR Latest Report)
    # -------------------------------------------------------------------------
    s20 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s20)
    builder._add_header_and_footer(
        s20,
        tag="FORESTRY METRICS",
        title="Forest Cover Assessment: India State of Forest Report (ISFR)",
        takeaway="Haryana has ~3.63% forest cover (1,603 sq km); Panchkula leads with 43.87% while Palwal/Fatehabad trail.",
        citation="India State of Forest Report (ISFR) 2021/latest, Forest Survey of India (FSI), MoEFCC.",
        slide_num=20
    )
    headers = ["Forest Assessment Metric", "Area (sq km)", "Percentage of State Area", "High-Yield Screening Exam Details"]
    rows = [
        ["Very Dense Forest (VDF)", "28.00 sq km", "0.06%", "Confined to Kalesar National Park (Yamunanagar) and high Morni ridgelines"],
        ["Moderately Dense Forest (MDF)", "451.00 sq km", "1.02%", "Distributed in Panchkula, Yamunanagar, Ambala, and Aravalli scrub tracts"],
        ["Open Forest (OF)", "1,124.00 sq km", "2.54%", "Largest forest category; agro-forestry belts, canal bundhs, roadside strips"],
        ["Total Recorded Forest Cover", "1,603.00 sq km", "3.63%", "One of India's lowest forest percentages; National Forest Policy target is 33%"],
        ["Tree Cover (Outside Recorded Forest)", "1,425.00 sq km", "3.22%", "Extensive Poplar and Eucalyptus farm forestry in Yamunanagar/Karnal"],
        ["Total Green Cover (Forest + Tree)", "3,028.00 sq km", "6.85%", "Combined forest and agro-forestry canopy in Haryana"],
        ["District with Highest Forest Cover", "Panchkula: 390.70 sq km", "43.87%", "Leading district in absolute area and proportional percentage"],
        ["District with Lowest Forest Cover", "Palwal (13.7 sq km) & Fatehabad", "0.71% - 0.73%", "Intensively cultivated agricultural plains with minimal natural canopy"]
    ]
    widths = [Inches(2.8), Inches(2.2), Inches(2.0), Inches(4.733)]
    builder._create_table(s20, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 21: National Parks: Sultanpur & Kalesar
    # -------------------------------------------------------------------------
    s21 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s21)
    builder._add_header_and_footer(
        s21,
        tag="PROTECTED AREAS",
        title="National Parks of Haryana: Sultanpur (Ramsar) & Kalesar (Sal)",
        takeaway="Haryana has exactly 2 National Parks: Sultanpur (Gurugram) and Kalesar (Yamunanagar).",
        citation="Wildlife Wing, Department of Forests & Wildlife, Haryana; Ramsar Convention on Wetlands.",
        slide_num=21
    )
    c1 = [
        "Geographical Location: Farrukhnagar block, Gurugram district; spread over ~1.43 sq km (352 acres).",
        "Historical Pioneer: Declared a sanctuary in 1971 through efforts of famous British ornithologist Peter Scott and Dr. Salim Ali; upgraded to National Park in 1991.",
        "Ramsar Site Recognition: Designated as a Ramsar Wetland of International Importance in August 2021.",
        "Avian Paradise: Critical wintering ground for over 250 migratory bird species flying the Central Asian Flyway (Siberian crane, Greater flamingo, Bar-headed goose, Northern pintail).",
        "Key Flora: Acacia nilotica (Kikar), Ziziphus, Ber, and aquatic emergent marsh vegetation."
    ]
    builder._create_card(s21, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "1. Sultanpur National Park (Gurugram)", c1, builder.COLOR_COBALT)

    c2 = [
        "Geographical Location: Eastern foothills of Shivaliks, Yamunanagar district; spread over ~46.82 sq km (11,570 acres).",
        "Upgraded to National Park: December 2003 (surrounded by contiguous Kalesar Wildlife Sanctuary).",
        "Interstate Forest Corridor: Contiguous with Rajaji National Park (Uttarakhand) and Simbalbara National Park (Himachal Pradesh).",
        "Dominant Flora: Pure stands of mature Sal (Shorea robusta) forest, Semul, Khair, and Shisham; the only natural Sal forest in Haryana.",
        "Faunal Diversity: Wild Asian elephants (visiting from Rajaji), Leopards, Barking deer, Sambar, Goral, King Cobra, and Red Junglefowl.",
        "Historical Monument: Ancient British-era Kalesar Dak Bungalow (1901) and historical Shiva temple."
    ]
    builder._create_card(s21, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "2. Kalesar National Park (Yamunanagar)", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 22: Wildlife Sanctuaries Network
    # -------------------------------------------------------------------------
    s22 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s22)
    builder._add_header_and_footer(
        s22,
        tag="SANCTUARIES NETWORK",
        title="Wildlife Sanctuaries of Haryana: 8 Sanctuaries & 2 Reserves",
        takeaway="Abubshahar (Sirsa) is largest (11,530 ha); Chhilchhila (Kurukshetra) is smallest (28 ha); Bhindawas is Ramsar.",
        citation="Chief Wildlife Warden, Haryana; MoEFCC Gazette Notifications.",
        slide_num=22
    )
    headers = ["Sanctuary / Reserve", "District Location", "Area (Hectares)", "Key Protected Fauna & Exam Points"]
    rows = [
        ["1. Abubshahar Sanctuary", "Sirsa (Mandi Dabwali border)", "11,530 ha", "Largest sanctuary in Haryana; prime habitat for Blackbuck (Kala Hiran) & Nilgai"],
        ["2. Bhindawas Sanctuary", "Jhajjar district", "411.55 ha", "Man-made freshwater wetland; Ramsar Site (2021); ~250 bird species; largest wetland lake"],
        ["3. Khaparwas Sanctuary", "Jhajjar (1.5 km from Bhindawas)", "82.70 ha", "Wetland sanctuary recognized for waterfowl protection; notified 1991"],
        ["4. Kalesar Wildlife Sanctuary", "Yamunanagar district", "5,435 ha", "Adjoins Kalesar NP; dense Sal & mixed forest; Leopard, Wild Boar, Chital corridor"],
        ["5. Bir Shikargah Sanctuary", "Pinjore (Panchkula district)", "767.30 ha", "Houses the famous Vulture Conservation Breeding Centre (VCBC); Sambar, Barking Deer"],
        ["6. Khol Hi-Raitan Sanctuary", "Panchkula (Shivalik belt)", "4,883 ha", "Critical Shivalik eco-fragile corridor connecting Bir Shikargah and Morni Hills"],
        ["7. Nahar Wildlife Sanctuary", "Kosli tehsil (Rewari district)", "211.35 ha", "Famous for Blackbuck, Jackal, Fox, and Peafowl in dry scrub deciduous forest"],
        ["8. Chhilchhila Sanctuary", "Kurukshetra (near Pehowa)", "28.92 ha", "Smallest wildlife sanctuary in Haryana; vital shallow lake nesting site for winter ducks"],
        ["Conservation Reserves", "Bir Bara Ban (Jind) & Saraswati (Kaithal/Kurukshetra)", "Bir Bara Ban: 419 ha\nSaraswati: 4,452 ha", "Bir Bara Ban (Jind town green lung); Saraswati Reserve (Bir Sonty) 2nd largest forest patch"]
    ]
    widths = [Inches(2.5), Inches(2.2), Inches(1.5), Inches(5.533)]
    builder._create_table(s22, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 23: Specialized Breeding Centres
    # -------------------------------------------------------------------------
    s23 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s23)
    builder._add_header_and_footer(
        s23,
        tag="EX-SITU CONSERVATION",
        title="Specialized Wildlife Breeding Centres: Location & Species Focus",
        takeaway="Pinjore hosts Asia's first Jatayu Vulture Conservation Breeding Centre (VCBC).",
        citation="Bombay Natural History Society (BNHS); Central Zoo Authority (CZA); Haryana Wildlife Department.",
        slide_num=23
    )
    c1 = [
        "Jatayu Vulture Conservation Breeding Centre (VCBC): Pinjore, Panchkula (Est. 2001 in collaboration with BNHS). Pioneer in captive breeding of 3 critically endangered Gyps species: Oriental white-backed, Long-billed, and Slender-billed vultures.",
        "Diclofenac Ban Catalyst: Research at Pinjore proved veterinary non-steroidal anti-inflammatory drug (NSAID) Diclofenac caused catastrophic visceral gout and renal failure in vultures, leading to a nationwide ban.",
        "Pheasant Breeding Centre: Morni, Panchkula district (Est. 1992-93). Specializes in Red Junglefowl and Kalij pheasant conservation.",
        "Red Junglefowl Genetic Purity: Rigorous DNA testing to prevent hybridization with domestic poultry."
    ]
    builder._create_card(s23, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Avian Breeding Hubs (Panchkula)", c1, builder.COLOR_COBALT)

    c2 = [
        "Chinkara Breeding Centre: Kairu, Bhiwani district. Established for breeding Indian Gazelle (Gazella bennettii) adapted to dry sand dune habitats.",
        "Peacock & Chinkara Breeding Centre: Jhabua, Rewari district. Dedicated to Indian Peafowl (Pavo cristatus) and Chinkara protection in Aravalli scrub.",
        "Crocodile Breeding Centre (Bhor Saidan): Bhor Saidan village near Pehowa, Kurukshetra district (Est. 1981-82). Houses Mugger crocodile (Crocodylus palustris).",
        "Blackbuck Breeding Centre: Pipli Mini Zoo, Kurukshetra and Rohtak Zoo.",
        "Mini Zoos of Haryana: Pipli Zoo (Kurukshetra), Rohtak Zoo, and Bhiwani Zoo under CZA regulatory standards."
    ]
    builder._create_card(s23, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Faunal Breeding & Reptile Centres", c2, builder.COLOR_GOLD)

    # -------------------------------------------------------------------------
    # Slide 24: Environmental Governance & Aravalli Green Wall
    # -------------------------------------------------------------------------
    s24 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s24)
    builder._add_header_and_footer(
        s24,
        tag="GREEN GOVERNANCE",
        title="Aravalli Green Wall Project, Urban Oxy-Vans & Climate Action",
        takeaway="Aravalli Green Wall Project aims to afforest a 5 km buffer along the 1,400 km Aravalli belt.",
        citation="Ministry of Environment, Forest & Climate Change (MoEFCC); Haryana Forest Vision 2030.",
        slide_num=24
    )
    c1 = [
        "Aravalli Green Wall Initiative: Inaugurated at Tikli village in Gurugram on International Day of Forests (March 2023).",
        "National Buffer Scale: 1,400 km long and 5 km wide green belt covering Haryana, Rajasthan, Gujarat, and Delhi.",
        "Core Objectives: Arrest eastward sand ingress from Thar desert; rejuvenate degraded Aravalli water recharge zones; combat regional desertification under UNCCD commitments.",
        "Species Selection: Indigenous drought-resilient trees: Kikar, Dhak (Butea monosperma), Ronjh, Neem, Peepal, and wild Ber.",
        "Aravalli Safari Park: Planned 10,000-acre project in Gurugram and Nuh districts; slated as the world's largest curated safari outside Africa."
    ]
    builder._create_card(s24, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Aravalli Green Wall & Eco-Restoration", c1, builder.COLOR_TEAL)

    c2 = [
        "Oxy-Van (ऑक्सी वन) Initiative: Urban micro-forests conceptualized during post-COVID environmental push. Flagship 800-acre Oxy-Van established at Karnal along Western Yamuna Canal.",
        "Panchkula Oxy-Van: Bir Ghaggar forest patch developed with 7 dedicated thematic groves (Panchavati, Navagraha, Nakshatra, Sugandh Vatika).",
        "Pran Vayu Devta Pension Scheme: Unique state scheme providing annual lifelong pension (~Rs 2,750+) to caretakers of healthy trees older than 75 years.",
        "Air Quality Action (CAQM): Stringent Commission for Air Quality Management regulations in 14 NCR districts of Haryana regulating stubble burning (Parali / पराली) and industrial boilers.",
        "Bio-Energy Policy: Extensive setup of 2G Ethanol Plant by IOCL at Panipat utilizing paddy straw to tackle seasonal smog."
    ]
    builder._create_card(s24, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Oxy-Vans, Pran Vayu Pension & Stubble Management", c2, builder.COLOR_PURPLE)

    print("Environment slides (19-24) built successfully.")
