"""
Module: Ancient, Medieval, Modern History & Freedom Struggle (Slides 25 - 40)
Author: Academic Research Presentation Engine (ARPE)
Target: Haryana FSO & State Screening Tests (HPSC / HSSC)
"""

from pptx.util import Inches

def build_history_slides(builder):
    prs = builder.prs
    blank = builder.blank_layout

    # -------------------------------------------------------------------------
    # Slide 25: Vedic Heritage & Kurukshetra Epic
    # -------------------------------------------------------------------------
    s25 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s25)
    builder._add_header_and_footer(
        s25,
        tag="VEDIC ANTIQUITY",
        title="Vedic Civilization: Brahmavarta, Saraswati Basin & Kurukshetra",
        takeaway="Manu Smriti defines Brahmavarta as the land between sacred rivers Saraswati & Drishadvati.",
        citation="Rigveda; Manu Smriti (Chapter 2, Verse 17); Archaeological Survey of India (ASI).",
        slide_num=25
    )
    c1 = [
        "Brahmavarta (ब्रह्मावर्त): Defined in Manu Smriti as the holy tract created by the gods between the divine rivers Saraswati and Drishadvati (Chautang).",
        "Rigvedic Composition: The majority of Rigvedic hymns were composed by Vedic seers (Bharadwaja, Vashistha, Vishvamitra) on the banks of the sacred Saraswati river.",
        "Brahmarshi Desha: Tract including Kurukshetra, Matsya, Panchala, and Surasena regarded as the highest ethical standard.",
        "Kurukshetra Naming: Named after King Kuru of Bharata dynasty, who ploughed the field with a golden plough across 48 kos (48 Kos Parikrama) to cultivate eight moral virtues (Satya, Dharma, Shaucha, etc.).",
        "Jyotisar (ज्योतिसर): The sacred banyan tree site where Lord Krishna delivered the celestial message of the Bhagavad Gita to Arjuna before the 18-day Mahabharata War."
    ]
    builder._create_card(s25, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Brahmavarta & Vedic Geography", c1, builder.COLOR_COBALT)

    c2 = [
        "48 Kos Parikrama: Circumambulation encompassing over 134 sacred tirthas spread across Kurukshetra, Kaithal, Karnal, and Jind districts.",
        "Karnal Connection: Named after Anga-raj Karna, famous for unshakeable charity; Karna Lake and Karna Kot.",
        "Prithudaka (Pehowa): Ancient pilgrimage town on Saraswati river named after King Prithu; premier center in northern India for Pitru Tarpan / ancestral shradh rites.",
        "Ban Ganga (Narkatari): Site near Kurukshetra where Arjuna shot an arrow into the earth to quench Bhishma Pitamah's thirst on the bed of arrows.",
        "Abhimanyu Pur (Amin): Ancient village where 16-year-old warrior Abhimanyu was trapped in the Chakravyuha; recently renamed Abhimanyupur."
    ]
    builder._create_card(s25, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Epic Mahabharata Landscape", c2, builder.COLOR_GOLD)

    # -------------------------------------------------------------------------
    # Slide 26: Rakhigarhi: Largest Harappan Metropolis
    # -------------------------------------------------------------------------
    s26 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s26)
    builder._add_header_and_footer(
        s26,
        tag="INDUS CIVILIZATION",
        title="Rakhigarhi (Hisar): Largest Harappan Metropolis in South Asia",
        takeaway="Spanning ~550 hectares across 11 mounds, Rakhigarhi is larger than Mohenjo-daro.",
        citation="Archaeological Survey of India (Dr. Amarendra Nath & Prof. Vasant Shinde); Cell / Science DNA Studies (2019).",
        slide_num=26
    )
    c1 = [
        "Geographical Setting: Located in Narnaund tehsil of Hisar district in the dry palaeochannel of the ancient Drishadvati (Chautang) river.",
        "Discovery & Excavations: First identified by Acharya Bhagwan Dev (1960s); extensively excavated by ASI under Dr. Amarendra Nath (1997-2000) and Prof. Vasant Shinde (Deccan College, 2011-2017).",
        "Magnitude & Scale: Comprises 11 distinct archaeological mounds covering ~550 hectares, officially surpassing Mohenjo-daro and Harappa as the largest Mature Harappan site.",
        "Urban Masterplanning: Mud-brick defensive perimeter walls, multi-roomed houses with planned bathrooms, baked brick soak pits, covered street drainage, and terracotta drainage pipes.",
        "Mound 7 Cemetery: Yielded over 60 burials (primary, secondary, and symbolic burials) with rich grave pottery and jasper/steatite beads."
    ]
    builder._create_card(s26, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Excavation History & Urban Scale", c1, builder.COLOR_COBALT)

    c2 = [
        "Landmark aDNA Discovery (Cell, 2019): Ancient DNA extracted from female skeleton (I6113) at Mound 7 proved: Indigenous South Asian ancestry with zero Steppe pastoralist or Anatolian farmer genetic contribution.",
        "Monumental Granary: Massive multi-chambered mud-brick structure with lime and decomposed grass plaster, engineered for grain storage and preservation against moisture.",
        "Jewellery Workshop: Uncovered lapidary beads of carnelian, agate, jasper, gold foils, copper mirrors, and micro-steatite beads indicating major trade hub.",
        "Gold Smelting & Metallurgy: Furnaces with crucibles containing minute traces of refined gold and copper slag.",
        "World-Class Site Museum: Central ASI and State Government site museum with on-site hostel, interpretation center, and experiential gallery."
    ]
    builder._create_card(s26, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Scientific Discoveries & DNA Breakthrough", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 27: Banawali, Bhirrana & Kunal (Fatehabad Cluster)
    # -------------------------------------------------------------------------
    s27 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s27)
    builder._add_header_and_footer(
        s27,
        tag="HARAPPAN CLUSTER",
        title="Banawali, Bhirrana & Kunal: The Fatehabad Civilizational Cradle",
        takeaway="Bhirrana is India's oldest Harappan site (~7500 BCE); Banawali yielded clay plough & barley.",
        citation="ASI Excavation Reports (R.S. Bisht, L.S. Rao, J.S. Khatri & M. Acharya); Nature Scientific Reports.",
        slide_num=27
    )
    headers = ["Archaeological Site", "District & River Basin", "Chronological Antiquity", "Key Excavated Artefacts & Screening Traps"]
    rows = [
        ["Bhirrana (भीड़ाना)", "Fatehabad (Ancient Saraswati bed)", "Hakra Ware phase (c. 7500 BCE - 6000 BCE)", "Oldest Harappan settlement in South Asia (C-14 dating); red ware pot shard featuring graffito of 'Dancing Girl' mimicking Mohenjo-daro bronze idol"],
        ["Banawali (बनावली)", "Fatehabad (14 km from Fatehabad town)", "Pre-Harappan, Mature, and Post-Harappan", "Excavated by Dr. R.S. Bisht (1974); terracotta toy plough (मिट्टी का हल); rich caches of barley grains (जौ); radial town planning layout without rigid grid"],
        ["Kunal (कुणाल)", "Fatehabad (Ratia tehsil on Saraswati)", "Early/Pre-Harappan (c. 4000 BCE)", "Excavated by J.S. Khatri & M. Acharya; two exquisite royal silver tiaras/crowns (चांदी के मुकुट); gold disc ornaments; bone tools; pit-dwellings preceding brick houses"],
        ["Agroha (अग्रोहा)", "Hisar (NH-9 highway)", "Post-Harappan to Kushan/Gupta period", "Capital of legendary King Agrasen (Agrawal community); coins bearing legend 'Agrodaka'; terracotta tablet with 7 musical notes (Sa Re Ga Ma Pa Dha Ni)"]
    ]
    widths = [Inches(2.4), Inches(2.3), Inches(2.8), Inches(4.233)]
    builder._create_table(s27, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 28: Mitathal, Siswal, Farmana & Balu
    # -------------------------------------------------------------------------
    s28 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s28)
    builder._add_header_and_footer(
        s28,
        tag="PROTO-HISTORIC SITES",
        title="Mitathal, Siswal, Farmana (Daksh Khera) & Balu",
        takeaway="Farmana revealed 70 Harappan burials; Balu yielded earliest evidence of garlic cultivation.",
        citation="Department of Archaeology & Museums, Haryana; Deccan College & Research Institute.",
        slide_num=28
    )
    c1 = [
        "Mitathal (Bhiwani): Excavated by Dr. Suraj Bhan (1968). Yielded Late Harappan copper hoard including copper axes, parashu, and rings; shows transition to Ochre Coloured Pottery (OCP).",
        "Siswal (Hisar): Excavated by Dr. Suraj Bhan (1970). Type-site for the distinct 'Siswal Culture' characterized by Kalibangan-I related painted pottery in non-alluvial dry zones.",
        "Tigrana (Bhiwani): Recent excavations revealing large-scale terracotta and semi-precious stone bead manufacturing workshop.",
        "Naurangabad (Bhiwani): Ancient Prakritnaket; coins, seals, and terracotta moulds of the Yaudheya Republic (योध्येय गणराज्य) bearing legend 'Yaudheya Ganasya Jaya'."
    ]
    builder._create_card(s28, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Bhiwani & Hisar Proto-Historic Network", c1, builder.COLOR_COBALT)

    c2 = [
        "Farmana / Daksh Khera (Rohtak): Joint excavation by Deccan College (Prof. Vasant Shinde) and Research Institute for Humanity and Nature (Japan). Massive 70-grave cemetery, copper ornaments, and jasper stamp seals.",
        "Balu (Kaithal): Excavated by Kurukshetra University (Dr. U.V. Singh). Crucial archaeobotanical find: Earliest evidence of cultivated Garlic (Allium sativum), green gram, and mustard seeds.",
        "Sugh / Shrughna (Yamunanagar): Ancient fortified port city on Yamuna; punch-marked coins, terracotta figurines, and Ashokan Buddhist heritage.",
        "Bhadra & Daulatpur (Kurukshetra): Late Harappan sites yielding painted grey ware (PGW) and clay seals."
    ]
    builder._create_card(s28, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Farmana, Balu & Agricultural Antiquity", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 29: Ashokan Inscriptions: Topra Pillar & Buddhist Sugh
    # -------------------------------------------------------------------------
    s29 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s29)
    builder._add_header_and_footer(
        s29,
        tag="MAURYAN & BUDDHIST ERA",
        title="The Topra Ashokan Pillar, Chaneti Stupa & Buddhist Heritage",
        takeaway="Firoz Shah Tughlaq moved the Topra Ashokan Pillar (Yamunanagar) to Kotla, Delhi in 1356 CE.",
        citation="Inscriptions of Asoka (Corpus Inscriptionum Indicarum); Shams-i Siraj Afif's 'Tarikh-i Firoz Shahi'.",
        slide_num=29
    )
    c1 = [
        "Topra Kalan (Yamunanagar): Original site of Emperor Ashoka's monumental monolithic polished sandstone pillar (3rd century BCE) bearing 7 Major Pillar Edicts in Brahmi script.",
        "Only Pillar with 7 Edicts: Topra is the only Ashokan pillar containing the entire Seventh Pillar Edict summarizing Dhamma policy (other pillars contain only 6).",
        "Transport to Delhi (1356 CE): Sultan Firoz Shah Tughlaq transported this multi-ton pillar on a 42-wheeled carriage pulled by 8,400 men and floated it down the Yamuna on large boats to install it atop Firoz Shah Kotla, Delhi.",
        "Topra Ashoka Edicts Park: Newly established memorial park in Yamunanagar housing exact replicas of all 14 Ashokan Rock Edicts and Topra pillar replica."
    ]
    builder._create_card(s29, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Topra Ashokan Pillar Edicts", c1, builder.COLOR_GOLD)

    c2 = [
        "Chaneti Buddhist Stupa (Yamunanagar): Located 3 km from Jagadhri; built in 3rd century BCE during Ashoka's reign using concentric baked bricks, mimicking Sanchi hemisphere style.",
        "Sugh (Ancient Shrughna): Major Buddhist monastery mentioned by Chinese pilgrim Xuanzang (Hiuen Tsang); recorded 5 monasteries with over 700 Buddhist monks in Shrughna.",
        "Karnal Inscription: Kharosthi script inscription discovered at Karnal dating to 1st century BCE / CE.",
        "Adi Badri Stupa: Excavated Buddhist monastery and stupa structures in Yamunanagar foothills highlighting cultural confluence.",
        "HPSC History Trap: Remember, Topra was originally in Ambala district; after 1989 reorganization, it falls squarely in Yamunanagar district."
    ]
    builder._create_card(s29, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Chaneti Stupa & Shrughna Monasteries", c2, builder.COLOR_PURPLE)

    # -------------------------------------------------------------------------
    # Slide 30: Pushyabhuti Dynasty & Emperor Harshavardhana
    # -------------------------------------------------------------------------
    s30 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s30)
    builder._add_header_and_footer(
        s30,
        tag="IMPERIAL HARSHA",
        title="Pushyabhuti Dynasty: Harshavardhana, Thanesar & Xuanzang",
        takeaway="Harsha ruled from Thanesar (Kurukshetra) before shifting his imperial capital to Kannauj.",
        citation="Banabhatta's 'Harshacharita'; Xuanzang's 'Si-Yu-Ki'; Archaeological Survey of India.",
        slide_num=30
    )
    c1 = [
        "Dynasty Founder: Pushyabhuti (worshipper of Lord Shiva) founded dynasty at Sthanishvara (modern Thanesar in Kurukshetra).",
        "Prabhakarvardhana: First powerful king; adopted title 'Paramabhattaraka Maharajadhiraja'; called 'Huna-harina-kesari' (a lion to the Huna deer).",
        "Accession of Harsha (606 CE): Ascended throne at age 16 following the murder of elder brother Rajyavardhana by Gauda King Shashanka.",
        "Shifting Capital: Unified crowns of Thanesar and Maukhari (Kannauj) and shifted imperial capital to Kannauj for administrative centrality.",
        "Literary Sovereign: Harsha himself authored three acclaimed Sanskrit plays: Ratnavali, Priyadarsika, and Nagananda."
    ]
    builder._create_card(s30, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Pushyabhutis of Thanesar", c1, builder.COLOR_COBALT)

    c2 = [
        "Banabhatta (Court Poet): Authored 'Harshacharita' (first historical biography in Sanskrit) and 'Kadambari'; described the prosperity of Srikantha Janapada (Haryana).",
        "Visit of Xuanzang (Hiuen Tsang): Traveled across Haryana (634-644 CE); stayed in Thanesar; praised social harmony, agricultural fertility, and Buddhist viharas.",
        "Defeat by Pulakeshin II: Harsha's southward expansion was halted at the Narmada River by Chalukya king Pulakeshin II (recorded in Aihole Inscription by Ravikirti).",
        "Harsha-ka-Tila (Thanesar): Excavated multi-period cultural mound in Thanesar revealing Kushana, Gupta, Harsha, Rajput, and Sultanate architectural layers.",
        "Sheikh Chilli's Tomb: Adjacent Sufi monument in Thanesar often called 'Taj Mahal of Haryana'."
    ]
    builder._create_card(s30, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Court Culture & Archaeological Legacy", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 31: Tomars, Chauhans & The Battles of Tarain
    # -------------------------------------------------------------------------
    s31 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s31)
    builder._add_header_and_footer(
        s31,
        tag="RAJPUT ERA",
        title="Tomar Dynasty, Prithviraj Chauhan & The Battles of Tarain (1191-92)",
        takeaway="First Battle of Tarain (1191): Prithviraj won; Second Battle (1192): Ghori won, founding Delhi Sultanate.",
        citation="Minhaj-i-Siraj's 'Tabaqat-i-Nasiri'; Chand Bardai's 'Prithviraj Raso'; ASI Monuments.",
        slide_num=31
    )
    headers = ["Historical Event / Dynasty", "Year / Period", "Protagonists / Builders", "Key Screening Points & Monuments"]
    rows = [
        ["Tomar Dynasty Rule", "8th - 12th Century CE", "Anangpal Tomar I & II", "Founded 'Dhillika' (Delhi); built Surajkund Sun Temple & Dam in Faridabad; Pehowa Inscription of Bhoja"],
        ["Asigarh Fort (Hansi)", "12th Century CE", "Prithviraj Chauhan III", "Also called 'Sword Fort' (Asi = Sword); strategic military stronghold of Chauhan kingdom"],
        ["First Battle of Tarain", "1191 CE (Tarori, Karnal)", "Prithviraj Chauhan vs Muhammad Ghori", "Prithviraj's Rajput confederacy routed Ghori; Govind Rai wounded Ghori; Ghori fled to Ghazni"],
        ["Second Battle of Tarain", "1192 CE (Tarori, Karnal)", "Muhammad Ghori vs Prithviraj Chauhan", "Ghori used 10,000 mounted archers; Rajput forces defeated; Prithviraj captured near Saraswati river; paved way for Muslim rule"],
        ["Qutb-ud-din Aibak Campaign", "1192 - 1206 CE", "First Mamluk Sultan", "Crushed local resistance in Hansi, Kohram, and Rewari; established Turkish military outposts"]
    ]
    widths = [Inches(2.5), Inches(1.8), Inches(2.8), Inches(4.633)]
    builder._create_table(s31, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 32: Delhi Sultanate: Razia Sultan, Balban & Firoz Shah
    # -------------------------------------------------------------------------
    s32 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s32)
    builder._add_header_and_footer(
        s32,
        tag="SULTANATE ERA",
        title="Delhi Sultanate: Razia Sultan at Kaithal, Balban & Firoz Shah",
        takeaway="Razia Sultan was assassinated near Kaithal (1240 CE); Firoz Shah Tughlaq founded Hisar & Fatehabad.",
        citation="Ziauddin Barani's 'Tarikh-i-Firoz Shahi'; Archaeological Survey of India.",
        slide_num=32
    )
    c1 = [
        "Razia Sultan's Assassination (14 Oct 1240): Razia (first female Muslim monarch of Delhi) and husband Malik Altunia were captured and murdered by rebel forces near Kaithal; historical Tomb of Razia Sultan exists in Kaithal town.",
        "Ghiyas-ud-din Balban: Subdued ferocious Mewati guerrilla rebels in southern Haryana; deforested Gurgaon/Mewat forests around Delhi; built a cordon of frontier forts at Hansi, Rewari, and Narnaul.",
        "Invasion of Timur (1398 CE): Plundered Sirsa, Fatehabad, Aharwan, Tohana, and Asandh en route to Delhi; faced fierce resistance from Jat and Ahir peasants.",
        "Yaudheya & Khanzada Chiefs: Semi-autonomous regional chieftains maintaining de facto control over south Haryana countryside."
    ]
    builder._create_card(s32, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Razia Sultan, Balban & Timur's Scourge", c1, builder.COLOR_COBALT)

    c2 = [
        "Foundation of Hisar-e-Firoza (1354 CE): Built by Sultan Firoz Shah Tughlaq as a fortified royal township; features Gujri Mahal (built for his beloved milkmaid Gujri), Lat ki Masjid, and Diwan-e-Aam.",
        "Lat (Pillar) of Firoz Shah: Ashokan-style sandstone monolithic pillar section erected in the courtyard of Hisar Fort complex.",
        "Foundation of Fatehabad: Founded in 1351 CE and named after his newborn son Fateh Khan; houses Humayun Mosque and an Ashokan pillar fragment.",
        "Firozpur Jhirka: Established in southern Mewat hills as an administrative post.",
        "Agricultural Canal Legacy: Excavated original Western Yamuna Canal to supply water from Yamuna river to royal gardens in Hisar."
    ]
    builder._create_card(s32, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Firoz Shah Tughlaq's Urban & Canal Architecture", c2, builder.COLOR_GOLD)

    # -------------------------------------------------------------------------
    # Slide 33: The Three Epic Battles of Panipat
    # -------------------------------------------------------------------------
    s33 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s33)
    builder._add_header_and_footer(
        s33,
        tag="DECISIVE CONFLICTS",
        title="The Three Battles of Panipat (1526, 1556, 1761): Strategic Analysis",
        takeaway="Panipat altered Indian history thrice: 1526 (Mughal advent), 1556 (Mughal restoration), 1761 (Maratha decimation).",
        citation="Baburnama; Abul Fazl's 'Akbarnama'; J.N. Sarkar's 'Fall of the Mughal Empire'; Panipat Museum.",
        slide_num=33
    )
    headers = ["Parameter", "First Battle of Panipat", "Second Battle of Panipat", "Third Battle of Panipat"]
    rows = [
        ["Date of Conflict", "21 April 1526", "5 November 1556", "14 January 1761 (Makar Sankranti)"],
        ["Opposing Forces", "Zahir-ud-din Babur vs Sultan Ibrahim Lodi", "Akbar (Bairam Khan) vs Hemu (Vikramaditya)", "Ahmad Shah Abdali vs Sadashivrao Bhau (Marathas)"],
        ["Military Innovation", "Tulughma flanking tactic & Rumi cart-gun carriage (Araba)", "Artillery bombardment; stray arrow hit Hemu's eye while atop 'Hawai' elephant", "Abdali's mobile swivel camel-mounted guns (Zamburak) & cavalry"],
        ["Casualties / Decisive Event", "Ibrahim Lodi killed on battlefield (Tomb in Panipat)", "Hemu captured unconscious & beheaded by Bairam Khan", "28,000+ Marathas slain; Vishwasrao & Jankoji Scindia martyred"],
        ["Historical Outcome", "Fall of Delhi Sultanate; foundation of Mughal Empire", "Decisive consolidation of Mughal rule under Akbar", "Shattered Maratha dream of Pan-Indian rule; paved way for British"],
        ["Memorial Landmark", "Kabuli Bagh Mosque (built by Babur in Panipat)", "Hemu's Samadhi Sthal at Shodapur village", "Kala Amb (Black Mango Tree) Memorial in Panipat"]
    ]
    widths = [Inches(2.0), Inches(3.2), Inches(3.2), Inches(3.333)]
    builder._create_table(s33, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 34: Hemu (Hemchandra Vikramaditya) of Rewari
    # -------------------------------------------------------------------------
    s34 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s34)
    builder._add_header_and_footer(
        s34,
        tag="LAST HINDU EMPEROR",
        title="Hemu (Samrat Hemchandra Vikramaditya): Rise from Rewari to Delhi",
        takeaway="Hemu won 22 consecutive battles and was crowned Emperor at Purana Qila, Delhi (Oct 7, 1556).",
        citation="Vincent Smith, 'Akbar the Great Mogul'; K.K. Bhardwaj, 'Hemu: Napoleon of Medieval India'.",
        slide_num=34
    )
    c1 = [
        "Humble Beginnings: Born in a Brahmin/Dhussar merchant family at Machhari village (Alwar) and brought up in Qutabpur, Rewari.",
        "Mercantile to Military Genius: Started as a saltpetre trader (vital ingredient for gunpowder); recognized for brilliant logistics by Sultan Islam Shah Suri.",
        "Prime Minister & Commander-in-Chief: Appointed Wazir and supreme military general by Sultan Adil Shah Suri; won 22 consecutive battles against Afghan rebels and Mughal commanders.",
        "Capture of Delhi & Agra (1556): Smashed Mughal Governor Tardi Beg Khan at the Battle of Tughlaqabad (Delhi) on 6 October 1556.",
        "Imperial Coronation: Assumed imperial title 'Samrat Hemchandra Vikramaditya' at Purana Qila, Delhi on 7 October 1556; re-established indigenous Hindu sovereignty."
    ]
    builder._create_card(s34, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Ascent of a Native Hero", c1, builder.COLOR_GOLD)

    c2 = [
        "The Tragic Turn at Panipat: On 5 November 1556, Hemu's forces were routing the Mughal army under Bairam Khan when a rogue stray arrow pierced Hemu's right eye, causing him to fall unconscious in his howdah.",
        "Mughal Vengeance: Captured unconscious and brought before the 13-year-old Akbar; beheaded by Bairam Khan; Hemu's head was sent to Kabul and torso hung outside Delhi gate.",
        "Father's Martyrdom: Hemu's 80-year-old father Puran Das refused to convert and was martyred at Alwar.",
        "Historical Stature: Described by historians as the 'Napoleon of Medieval India' and the last Hindu sovereign to hold Delhi's imperial throne.",
        "Memorial: Hemu's Samadhi Sthal situated at Shodapur village near Panipat."
    ]
    builder._create_card(s34, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Battle of Panipat II & Historical Legacy", c2, builder.COLOR_PURPLE)

    # -------------------------------------------------------------------------
    # Slide 35: Battle of Karnal (1739) & George Thomas (Jahazi Sahib)
    # -------------------------------------------------------------------------
    s35 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s35)
    builder._add_header_and_footer(
        s35,
        tag="18TH CENTURY TURMOIL",
        title="Battle of Karnal (1739) & The Adventurer Raja: George Thomas at Hansi",
        takeaway="Nadir Shah defeated Muhammad Shah at Karnal; Irish adventurer George Thomas ruled Hansi (1797-1801).",
        citation="William Irvine's 'Later Mughals'; Capt. Francklin's 'Military Memoirs of Mr. George Thomas'.",
        slide_num=35
    )
    c1 = [
        "Battle Date: 24 February 1739 at Karnal plains.",
        "Opponents: Nadir Shah (Shah of Persia / Afsharid dynasty) vs Mughal Emperor Muhammad Shah 'Rangeela'.",
        "Military Disaster: In just 3 hours of combat, Nadir Shah's disciplined musketeers annihilated the superior-numbered Mughal army; Burhan-ul-Mulk captured.",
        "Plunder of Delhi: Nadir Shah marched to Delhi, ordered horrific general massacre (Qatl-e-Aam), and looted the Peacock Throne (Takht-e-Taus) and the Koh-i-Noor diamond.",
        "Treaty of Surji-Anjangaon (30 Dec 1803): Following Second Anglo-Maratha War, Daulat Rao Scindia ceded Haryana territory to British East India Company."
    ]
    builder._create_card(s35, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Battle of Karnal (1739) & Collapse", c1, builder.COLOR_CRIMSON)

    c2 = [
        "George Thomas ('Jahazi Sahib'): Irish deserter who rose through mercenary ranks under Begum Samru of Sardhana and Maratha commander Appa Khande Rao.",
        "Independent Kingdom at Hansi (1797-1801): Established independent rule carving out 800 villages; rebuilt Asigarh Fort at Hansi as his royal capital.",
        "Mint & Artillery Foundry: Established his own currency mint issuing 'Sikka-e-Sahib' coins and cast European-style cannons in Hansi.",
        "Downfall: Besieged by French General Pierre Perron commanding Maratha forces; surrendered Hansi in 1802; died en route to Calcutta at Berhampore.",
        "Hansi Heritage: George Thomas's residence 'Jahaz Kothi' (so called because it resembled a ship) preserved as heritage monument in Hisar."
    ]
    builder._create_card(s35, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "George Thomas: The Jahazi Sahib of Hansi", c2, builder.COLOR_COBALT)

    # -------------------------------------------------------------------------
    # Slide 36: 1857 Uprising: Ambala Cantt Outbreak
    # -------------------------------------------------------------------------
    s36 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s36)
    builder._add_header_and_footer(
        s36,
        tag="1857 WAR OF INDEPENDENCE",
        title="The 1857 Uprising: Outbreak at Ambala Cantonment (10 May 1857)",
        takeaway="The uprising broke out at Ambala Cantt 9 hours BEFORE Meerut (8-9 hours earlier on 10 May 1857).",
        citation="State Archives Department, Haryana; Prof. K.C. Yadav, 'The Revolt of 1857 in Haryana'.",
        slide_num=36
    )
    c1 = [
        "Pioneering Revolt Timing: Mutiny broke out in the 5th and 60th Native Infantry at Ambala Cantonment at ~9:00 AM on Sunday, 10 May 1857 (almost 9 hours prior to the Meerut rebellion at 5:00 PM).",
        "The Traitor Sham Singh: Sepoy Sham Singh betrayed the mutineers by leaking plan details to British officers, allowing commander General Anson to disarm the regiments.",
        "Ambala War Memorial: Grand 22-acre international-standard memorial museum constructed on NH-44 at Ambala Cantt to honor unsung 1857 martyrs.",
        "Strategic Ambala Cantt: Established in 1843 (transferred from Karnal after malaria epidemic); northern military headquarters of British India."
    ]
    builder._create_card(s36, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Ambala Cantt: The Forgotten Epicenter", c1, builder.COLOR_COBALT)

    c2 = [
        "Regional Flare-Ups across Haryana:",
        "• Rohtak & Jhajjar: Rebel sepoys led by Risaldar Bisarat Ali routed British district officers.",
        "• Hisar & Hansi: Led by Muhammad Azim (Prince of Delhi) and Hukam Chand Jain; British deputy commissioner John Wedderburn killed.",
        "• Sirsa: 4th Irregular Cavalry rebelled; British fled to Bikaner.",
        "• Panipat: Led by Imam Ali Qalandar and local Gujjar/Jat villagers cutting Grand Trunk Road communication.",
        "• Mewat / Gurgaon: Sadruddin Mewati led thousands of Meos against British forces; burnt police outposts."
    ]
    builder._create_card(s36, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "District-Wide Spread of Resistance", c2, builder.COLOR_CRIMSON)

    # -------------------------------------------------------------------------
    # Slide 37: 1857 Heroes: Rao Tula Ram, Raja Nahar Singh & Nawab of Jhajjar
    # -------------------------------------------------------------------------
    s37 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s37)
    builder._add_header_and_footer(
        s37,
        tag="MARTYRS OF 1857",
        title="Key Leaders of 1857: Rao Tula Ram, Raja Nahar Singh & Abdul Rahman",
        takeaway="Rao Tula Ram liberated Rewari; Raja Nahar Singh ('Lion of Ballabhgarh') was hanged in Chandni Chowk.",
        citation="National Archives of India; Punjab Mutiny Records (1857-58); Haryana Sahitya Akademi.",
        slide_num=37
    )
    headers = ["1857 Rebel Leader", "Region / Princely State", "Resistance Milestones", "Fate & Commemoration"]
    rows = [
        ["Rao Tula Ram (with cousin Gopal Dev)", "Rewari & Ahirwal region", "Liberated Rewari; established cannon factory at Rampura Fort; funded Mughal treasury at Delhi", "Escaped to Iran/Afghanistan to seek Russian aid; died in Kabul (23 Sep 1863); Haryana Shaheedi Diwas"],
        ["Raja Nahar Singh ('Ballabhgarh Sher')", "Ballabhgarh (Faridabad)", "Guarded Delhi's southern frontier; defied British commands; protected refugees", "Arrested deceitfully; publicly hanged at Chandni Chowk on 9 Jan 1858; Ballabhgarh metro named in honor"],
        ["Nawab Abdul Rahman Khan", "Jhajjar princely state", "Mobilized cavalry and infantry in support of Emperor Bahadur Shah Zafar", "Tried by military commission; hanged outside Red Fort on 23 Dec 1857; Jhajjar territory annexed"],
        ["Nawab Ahmad Ali", "Farrukhnagar (Gurugram)", "Co-ordinated Meo and Ahir peasant resistance against British relief columns", "Hanged at Kotwali Chandni Chowk on 23 Jan 1858"],
        ["Lala Hukam Chand Jain & Muneer Beg", "Hansi (Hisar district)", "Drafted plans to overthrow British garrisons across western Haryana", "Hanged simultaneously from the same tree in Hansi in 1858"]
    ]
    widths = [Inches(2.5), Inches(2.2), Inches(3.6), Inches(3.433)]
    builder._create_table(s37, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 38: Battle of Naseepur & Haryana Shaheedi Diwas
    # -------------------------------------------------------------------------
    s38 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s38)
    builder._add_header_and_footer(
        s38,
        tag="MARTYRS' DAY",
        title="Battle of Naseepur (16 Nov 1857) & Haryana Shaheedi Diwas (23 Sep)",
        takeaway="Haryana observes Shaheedi Diwas on 23 September in sacred memory of Rao Tula Ram.",
        citation="Department of Information, Public Relations and Languages, Haryana.",
        slide_num=38
    )
    c1 = [
        "Battlefield Setting: Naseepur plain situated 3 km north of Narnaul town (Mahendragarh district).",
        "Date of Epic Battle: 16 November 1857.",
        "Allied Forces: Joint patriotic command under Rao Tula Ram, Gopal Dev, Gen. Abdus Samad Khan (Jhajjar), and Jodhpur Legion mutineers (5,000+ fighters).",
        "Opposing British Column: Col. Gerrard leading combined force of 1st European Bengal Fusiliers, Sikh cavalry (Patiala/Jind), and Multani horse.",
        "Battle Climax: Fierce hand-to-hand combat; British Commander Col. Gerrard was shot dead on the field; Captain Morehead wounded; allied forces eventually succumbed to heavy artillery fire."
    ]
    builder._create_card(s38, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "The Bloodbath at Naseepur", c1, builder.COLOR_CRIMSON)

    c2 = [
        "Haryana Shaheedi Diwas (23 September): Observed as a gazetted State public holiday every year in Haryana to honor Rao Tula Ram who breathed his last in Kabul on 23 September 1863.",
        "Post-1857 Disenfranchisement: Under the Queen's Proclamation of 1858, Haryana was ripped away from North-Western Provinces and merged with Punjab as collective punishment.",
        "Princely Rewards for Loyalist Chiefs: Princely states that aided the British (Patiala, Nabha, Jind, Kapurthala) were rewarded with confiscated territories of Jhajjar, Dadri, and Narnaul.",
        "Exemplary Ahirwal Valour: Legacy continued through the Battle of Rezang La (1962) where 13 Kumaon (Ahir troops) made supreme sacrifice under Major Shaitan Singh."
    ]
    builder._create_card(s38, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Shaheedi Diwas & Territorial Repercussions", c2, builder.COLOR_GOLD)

    # -------------------------------------------------------------------------
    # Slide 39: Sir Chhotu Ram (Deenbandhu) & Peasant Empowerment
    # -------------------------------------------------------------------------
    s39 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s39)
    builder._add_header_and_footer(
        s39,
        tag="PEASANT CRUSADER",
        title="Sir Chhotu Ram (Deenbandhu / Rahbar-i-Azam): Peasant Legislation",
        takeaway="Sir Chhotu Ram passed revolutionary agrarian debt cancellation acts & conceptualized Bhakra Dam.",
        citation="Dr. Raghuvendra Tanwar, 'Sir Chhotu Ram: A Biography'; Punjab Legislative Assembly Debates.",
        slide_num=39
    )
    c1 = [
        "Early Life & Identity: Born Ram Richhpal on 24 November 1881 at Garhi Sampla (Rohtak district); earned law degree from Agra College.",
        "Jat Gazette (1916): Launched premier weekly newspaper 'Jat Gazette' from Rohtak in Urdu/Hindi to articulate rural peasant grievances.",
        "Unionist Party (Zamindara League 1923): Founded non-communal agrarian party with Sir Fazl-i-Hussain representing Hindu, Sikh, and Muslim agriculturists.",
        "Bhakra Dam Visionary: Spearheaded water agreement with the Raja of Bilaspur (1944) laying the legal and financial foundation for Bhakra Dam.",
        "Iconic Statues: Prime Minister Narendra Modi unveiled a monumental 64-foot statue of Sir Chhotu Ram at Garhi Sampla in 2018."
    ]
    builder._create_card(s39, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Biographical & Political Landmarks", c1, builder.COLOR_COBALT)

    c2 = [
        "Landmark Pro-Farmer Legislative Acts:",
        "1. Punjab Relief of Indebtedness Act 1934: Liquidated predatory moneylender debts; capped compound interest; protected oxen and agricultural tools from auction.",
        "2. Punjab Debtor's Protection Act 1936: Exempted ancestral agricultural land and standing crops from court attachment.",
        "3. Punjab Restitution of Mortgaged Lands Act 1938: Restored hundreds of thousands of mortgaged agricultural acres back to destitute farmers without cash payment.",
        "4. Punjab Agricultural Produce Markets Act 1939: Established institutional Mandis (Market Committees) eliminating fraudulent merchant weights (Dami).",
        "Titles: Conferred 'Rao Bahadur' (1919), Knighthood ('Sir' in 1937), 'Rahbar-i-Azam' by grateful peasants."
    ]
    builder._create_card(s39, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Legislative Magna Carta for Farmers", c2, builder.COLOR_GOLD)

    # -------------------------------------------------------------------------
    # Slide 40: National Freedom Movement: Lala Lajpat Rai & Neki Ram
    # -------------------------------------------------------------------------
    s40 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s40)
    builder._add_header_and_footer(
        s40,
        tag="FREEDOM STRUGGLE",
        title="Freedom Struggle Stalwarts: Lala Lajpat Rai & Pandit Neki Ram Sharma",
        takeaway="Pt. Neki Ram Sharma founded 'Sandesh'; declared 'I want the entire land of India'.",
        citation="National Archives; State Archives Haryana; 'Who's Who of Haryana Freedom Fighters'.",
        slide_num=40
    )
    headers = ["Freedom Fighter / Pioneer", "Native Region", "Major Publications / Institutions", "Iconic Contribution / Exam Quotation"]
    rows = [
        ["Lala Lajpat Rai ('Punjab Kesari')", "Hisar (Karma-bhoomi 1886-92)", "Founded Hisar Arya Samaj (1886) & Hisar Congress (1887)", "Practised law at Hisar Court; served on Hisar Municipal Committee; established national protest networks"],
        ["Pandit Neki Ram Sharma ('Haryana Kesari')", "Kelanga village (Rohtak/Bhiwani)", "Weekly Newspaper: 'Sandesh' (संदेश) from Bhiwani", "Opposed British recruitment in WWI; famously told Viceroy: 'Mujhe to sampoorna Bharat ki bhoomi chahiye'"],
        ["Lala Murli Dhar ('Grand Old Man of Punjab')", "Ambala (Born in Palwal)", "Founding member of Indian National Congress (1885 Bombay)", "Jailed in 1886; renounced title 'Rai Bahadur' & 'Kaisar-i-Hind' medal in protest against Jallianwala Bagh massacre (1919)"],
        ["Deshbandhu Gupta", "Panipat", "Daily Tej (Delhi) & nationalist journalism", "Championed integration of Haryana and Delhi in 1932; leading figure in Constituent Assembly"],
        ["Smt. Kasturba & Mahatma Gandhi Visits", "Palwal & Rohtak", "First arrested in Haryana at Palwal Railway Station (10 Apr 1919)", "En route to Punjab against Rowlatt Act; Gandhi Smarak established at Palwal railway platform"]
    ]
    widths = [Inches(2.5), Inches(2.2), Inches(3.2), Inches(3.833)]
    builder._create_table(s40, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    print("History slides (25-40) built successfully.")
