"""
Module: District Encyclopedic Digest, Schemes, Food Safety & Exam Traps (Slides 61 - 65)
Author: Academic Research Presentation Engine (ARPE)
Target: Haryana FSO & State Screening Tests (HPSC / HSSC)
"""

from pptx.util import Inches

def build_districts_and_traps_slides(builder):
    prs = builder.prs
    blank = builder.blank_layout

    # -------------------------------------------------------------------------
    # Slide 61: District Master Matrix I: Northern & Central Districts
    # -------------------------------------------------------------------------
    s61 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s61)
    builder._add_header_and_footer(
        s61,
        tag="DISTRICT ENCYCLOPEDIA I",
        title="District Profile Matrix Part I: Northern & Central Districts",
        takeaway="Quick-revision encyclopedic matrix covering administrative hubs from Panchkula to Rohtak.",
        citation="Directorate of Land Records, Haryana; District Gazetteers.",
        slide_num=61
    )
    headers = ["District", "Creation Date", "Historical / Nickname", "Key Screening Points & Monuments"]
    rows = [
        ["Panchkula", "15 Aug 1995", "'City of Five Canals' (Panch Kula)", "Morni Hills (Karoh 1514m); Pinjore Gardens (Yadavindra Gardens built by Fidai Khan); Mansa Devi Temple; HMT Pinjore; VCBC Vulture Centre"],
        ["Ambala", "1 Nov 1966 (Original)", "'Scientific City' / 'Mixie City'", "Ambala Cantt (Est. 1843); 1857 Uprising outbreak (May 10); birthplace of Sushma Swaraj; Gurdwara Manji Sahib; Muslim shrines"],
        ["Yamunanagar", "1 Nov 1989", "'Paper City' (Abdullapur Mandi)", "Hathnikund & Tajewala; Kalesar NP; Saraswati Sugar Mills; Chaneti Stupa; Topra Ashokan pillar site; Jagadhri Railway Workshop"],
        ["Kurukshetra", "23 Jan 1973", "'Dharmakshetra' / Holy City", "Brahma Sarovar & Sannihit Sarovar; Jyotisar (Gita sermon); Sheikh Chilli's Tomb; Kurukshetra University (1st Univ in Haryana, 1956); Kalpana Chawla Planetarium"],
        ["Kaithal", "1 Nov 1989", "'Kapishthala' (Hanuman birthplace)", "Tomb of Razia Sultan (1240); Falgu Tirth (Pharal); Anjani Teela; Neem Sahib Gurdwara; Balu archaeological site"],
        ["Karnal", "1 Nov 1966 (Original)", "'Rice Bowl' / 'City of Daanveer Karna'", "NDRI (1955); CSSRI (1969); IIWBR; NBAGR; Karna Lake; Kalpana Chawla birthplace (1962); Liberty Shoes; Taraori Basmati export hub"],
        ["Panipat", "1 Nov 1989\n(Recreated 1992)", "'City of Weavers' / 'Textile City'", "3 Battles of Panipat (1526, 1556, 1761); Kabuli Bagh Mosque; Kala Amb; Bu Ali Shah Qalandar Dargah; Altaf Hussain Hali tomb; IOCL Refinery"],
        ["Sonipat", "22 Dec 1972", "'Swarnaprastha' / Golden City", "NIFTEM Kundli; OP Jindal Global Univ; Sports University Rai (1st Sports Univ); Deenbandhu Chhotu Ram Univ (DCRUST Murthal); Atlas Cycles (historical)"],
        ["Rohtak", "1 Nov 1966 (Original)", "'Rohitaka' / Education Hub", "MDU Rohtak (Est. 1976); Pt. BD Sharma Post Graduate Institute of Medical Sciences (PGIMS); Tilyar Lake; Garhi Sampla (Chhotu Ram); Rewari-Rohtak eco zone"]
    ]
    widths = [Inches(1.8), Inches(1.8), Inches(2.8), Inches(5.333)]
    builder._create_table(s61, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 62: District Master Matrix II: Western & Southern Districts
    # -------------------------------------------------------------------------
    s62 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s62)
    builder._add_header_and_footer(
        s62,
        tag="DISTRICT ENCYCLOPEDIA II",
        title="District Profile Matrix Part II: Western & Southern Districts",
        takeaway="Quick-revision encyclopedic matrix covering administrative hubs from Hisar to Charkhi Dadri.",
        citation="Directorate of Land Records, Haryana; District Gazetteers.",
        slide_num=62
    )
    headers = ["District", "Creation Date", "Historical / Nickname", "Key Screening Points & Monuments"]
    rows = [
        ["Hisar", "1 Nov 1966 (Original)", "'City of Steel' / Hisar-e-Firoza", "Founded by Firoz Shah Tughlaq (1354); Gujri Mahal; CCSHAU (1970); CIRB (1985); LUVAS (2010); Jindal Steel; Rakhigarhi; Agroha Dham; Blue Bird Lake"],
        ["Fatehabad", "15 Jul 1997", "'Pink City of Haryana'", "Founded by Firoz Shah (named after son Fateh Khan); Bhirrana (oldest site ~7500 BCE); Banawali & Kunal; Gorakhpur Nuclear Power Plant (2800 MW)"],
        ["Sirsa", "1 Sep 1975", "'Saraswati Nagari' / 'Arnyavanya'", "Largest district by area (4,277 sq km); lowest population density (303); Abubshahar Sanctuary (largest); Ottu Barrage on Ghaggar; Tara Baba Kutiya"],
        ["Jind", "1 Nov 1966 (Original)", "'Heart of Haryana' / Jayantipuri", "Named after Jayanti Devi; 'Home of Murrah Buffalo'; Bir Bara Ban reserve; Dhamtan Sahib Gurdwara; Safidon (ancient Sarpadevi where Janamejaya performed snake sacrifice)"],
        ["Bhiwani", "22 Dec 1972", "'Chhoti Kashi' / 'Mini Cuba'", "City of 300+ temples ('Chhoti Kashi'); Bhiwani Boxing Club; Mitathal & Naurangabad; Devsar & Star Monument; Board of School Education Haryana (BSEH)"],
        ["Charkhi Dadri", "1 Dec 2016", "22nd District of Haryana", "Carved from Bhiwani; flexible sandstone (कम्पा पत्थर) at Kaliana; famous Phogat sisters (wrestling); 1996 mid-air plane crash site"],
        ["Mahendragarh", "1 Nov 1966 (Original)", "Mineral Capital of Haryana", "Headquarters at Narnaul (only district whose HQ is not in its own name!); Dhosi Hill (652m, Chyavana Rishi); Birbal ka Chhatta; Jal Mahal; Antri-Beharipur iron"],
        ["Rewari", "1 Nov 1989", "'Brass City' / Ahirwal Epicenter", "Rao Tula Ram (1857 hero); Hemu's hometown; Heritage Steam Locomotive Shed (1893, Fairy Queen); Kund slate stone; Bawal IMT; first Arya Samaj branch in Haryana (1880)"],
        ["Gurugram", "1 Nov 1966 (Original)", "'Millennium City' / Guru Drona seat", "Highest per capita income; Maruti Suzuki; Sheetla Mata Temple; Sultanpur NP; National Brain Research Centre (NBRC) Manesar; International Solar Alliance (ISA) HQ"],
        ["Faridabad", "2 Aug 1979", "Industrial City (Baba Farid)", "Most populous & dense district (2442/sq km); Surajkund crafts mela; Badkhal lake; Raja Nahar Singh Palace (Ballabhgarh); Escorts & JCB"],
        ["Palwal", "15 Aug 2008", "'Apelva' / Cotton City", "Panchvati Temple; Mahatma Gandhi first arrested here (10 Apr 1919); Hodal Chameli forest; lowest recorded forest percentage (~0.71%)"],
        ["Nuh (Mewat)", "4 Apr 2005", "Historical Mewat / Satyamevpuram", "Highest sex ratio (907) & lowest literacy (54.08%); Kotla lake & Sheikh Musa tomb; Ratvai dance; Chhui Mal pond; Country's first Mobile Court (Punhana, 2007)"]
    ]
    widths = [Inches(1.8), Inches(1.8), Inches(2.8), Inches(5.333)]
    builder._create_table(s62, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 63: Landmark Welfare & Digital Governance Schemes
    # -------------------------------------------------------------------------
    s63 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s63)
    builder._add_header_and_footer(
        s63,
        tag="STATE SCHEMES & E-GOVERNANCE",
        title="Landmark Initiatives: Parivar Pehchan Patra, MFMB & Bhavantar",
        takeaway="Haryana is 1st state to implement Bhavantar Bharpayee Yojana (price deficit compensation).",
        citation="Department of Information Technology, Electronics & Communication (DITECH); CM Window Dashboard.",
        slide_num=63
    )
    headers = ["Flagship Scheme / Portal", "Launch Year", "Core Objective", "Screening Exam Question Focus"]
    rows = [
        ["Parivar Pehchan Patra (PPP)", "2020", "8-digit unique biometric/digital family ID for every Haryana household", "Mandatory single identifier for availing all state DBT, pensions (Old Age), राशन कार्ड, and public scholarship schemes"],
        ["Meri Fasal Mera Byora (MFMB)", "2019", "Unified digital portal for self-registration of cultivated crop acreage", "Facilitates transparent MSP procurement by FCI/HAFED, crop loss compensation under PMFBY, and financial subsidies"],
        ["Bhavantar Bharpayee Yojana (BBY)", "1 Jan 2018", "Price Deficit Financing Scheme for horticulture & agricultural crops", "First state in India to launch BBY; compensates farmers if wholesale market price drops below designated benchmark price; covers 21+ horticulture crops & Bajra"],
        ["Mhara Gaon Jagmag Gaon (MGJG)", "1 July 2015", "24-hour uninterrupted three-phase electricity to rural Haryana villages", "Launched from Dyalpura (Kurukshetra); linked to voluntary smart meter installation and power bill payment; >5,800+ villages covered"],
        ["Mukhyamantri Antyodaya Parivar Utthan", "2021", "Targeted upliftment of poorest families earning <Rs 1.80 Lakh/year", "Provides customized skill training, subsidized entrepreneurship loans, and self-employment linkages under Antyodaya Gramin Melas"],
        ["CM Window & Jan Samvad", "25 Dec 2014", "Public grievance redressal system launched on Good Governance Day", "Direct digital tracking of citizen complaints monitored at Chief Minister's Office (CMO) level"]
    ]
    widths = [Inches(2.5), Inches(1.5), Inches(3.8), Inches(3.933)]
    builder._create_table(s63, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 64: Food Safety Regulation, Enforcement & Mobile Testing Labs
    # -------------------------------------------------------------------------
    s64 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s64)
    builder._add_header_and_footer(
        s64,
        tag="FSO ENFORCEMENT DIRECTIVE",
        title="Food Safety Administration in Haryana: Infrastructure & Enforcement",
        takeaway="Food Safety Department functions under Food and Drugs Administration (FDA) Haryana, Panchkula.",
        citation="Food Safety and Standards Act 2006; Food and Drugs Administration (FDA) Haryana Annual Action Plan.",
        slide_num=64
    )
    c1 = [
        "Apex Regulatory Body: Food and Drugs Administration (FDA), Haryana (Headquartered at Sector-14, Panchkula).",
        "Administrative Hierarchy:",
        "• Commissioner of Food Safety (State Secretary / DG Health level).",
        "• Designated Officers (DO) at District Level (Administrative jurisdiction over licensing and prosecution).",
        "• Food Safety Officers (FSO) at Field Level (Inspection, sampling, seizures, and compounding).",
        "Food Testing Infrastructure:",
        "• State Food Testing Laboratory, Karnal: Premier NABL-accredited facility testing adulteration, pesticides, and microbial pathogens.",
        "• State Drug & Food Testing Lab, Chandigarh / Panchkula.",
        "• Referral Testing: Central Food Laboratory (CFL) Kolkata and NIFTEM Kundli."
    ]
    builder._create_card(s64, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Statutory Enforcement Structure", c1, builder.COLOR_COBALT)

    c2 = [
        "Mobile Food Testing Labs ('Food Safety on Wheels'):",
        "• Fleet of customized mobile analytical vans operating across rural mandis, school mid-day meals, and dairy hubs.",
        "• Conducts instant on-spot testing for milk fat/SNF, synthetic milk (detergent, urea), starch in paneer, and adulterated edible oils.",
        "Eat Right India Campaigns in Haryana:",
        "• Eat Right Mela and Walkathons organized at Kurukshetra, Panchkula, and Gurugram.",
        "• Clean Street Food Hubs: Certified hygienic street food markets established at Sector-14 Panchkula and Gurugram.",
        "Fortification Mandate:",
        "• Fortified Wheat Flour (Atta fortified with Iron, Folic Acid, B12) distributed through Public Distribution System (PDS) across all 22 districts to combat stunting and anemia."
    ]
    builder._create_card(s64, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Mobile Testing & Eat Right Initiatives", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 65: Top 30 High-Frequency Exam Traps & Rapid-Fire One-Liners
    # -------------------------------------------------------------------------
    s65 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s65)
    builder._add_header_and_footer(
        s65,
        tag="HIGH-YIELD REVISION",
        title="Top 30 High-Frequency Exam Traps: Rapid-Fire Screening Checklist",
        takeaway="Must-memorize high-yield factual checkpoints most frequently tested in HPSC & HSSC exams.",
        citation="Compilation of Repeated Questions from HPSC HCS, FSO, ADO, and HSSC CET Examinations (2015-2024).",
        slide_num=65
    )
    c1 = [
        "1. Cherrapunji of Haryana: Chhachhrauli (Yamunanagar).",
        "2. Highest Elevation Peak: Karoh Peak in Morni Hills (1,514 metres).",
        "3. Highest Aravalli Peak: Dhosi Hill in Mahendragarh (652 metres).",
        "4. Oldest Harappan Settlement: Bhirrana (Fatehabad, ~7500 BCE).",
        "5. Largest Harappan Settlement: Rakhigarhi (Hisar, ~550 hectares).",
        "6. Clay Plough Excavated at: Banawali (Fatehabad, by R.S. Bisht).",
        "7. Royal Silver Crowns Excavated at: Kunal (Fatehabad).",
        "8. Earliest Garlic Evidence: Balu (Kaithal).",
        "9. Topra Ashokan Pillar Moved by: Sultan Firoz Shah Tughlaq in 1356 CE.",
        "10. 1857 Uprising Started at: Ambala Cantt (10 May 1857, 9 AM).",
        "11. Shaheedi Diwas Observed on: 23 September (death of Rao Tula Ram).",
        "12. First Chief Minister: Pt. Bhagwat Dayal Sharma (Beri, Jhajjar).",
        "13. First Non-Congress CM: Rao Birender Singh (Vishal Haryana Party).",
        "14. First Female Assembly Speaker: Shanno Devi (1966).",
        "15. First Indian Chief Justice: Sir Shadi Lal (Lahore High Court, 1920)."
    ]
    builder._create_card(s65, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Rapid-Fire Checkpoints: 1 to 15", c1, builder.COLOR_COBALT)

    c2 = [
        "16. 100% Rural Electrification Achieved on: 29 Nov 1970 (under Bansi Lal).",
        "17. Old Age Pension Introduced by: Chaudhary Devi Lal in 1987 (Rs 100).",
        "18. District with Highest Forest %: Panchkula (43.87%).",
        "19. District with Lowest Forest %: Palwal (~0.71%) and Fatehabad.",
        "20. Largest Wildlife Sanctuary: Abubshahar (Sirsa, 11,530 ha).",
        "21. Smallest Wildlife Sanctuary: Chhilchhila (Kurukshetra, 28 ha).",
        "22. Two Ramsar Wetland Sites: Sultanpur NP (Gurugram) & Bhindawas (Jhajjar).",
        "23. Vulture Breeding Centre Located at: Bir Shikargah, Pinjore (Panchkula).",
        "24. Home / Capital of Murrah Buffalo: Jind district.",
        "25. National Institute of Food Tech (NIFTEM) at: Kundli (Sonipat).",
        "26. Centre of Excellence for Vegetables: Gharaunda (Karnal).",
        "27. Flexible Sandstone (कम्पा पत्थर) at: Kaliana (Charkhi Dadri).",
        "28. Slate Stone Quarry Hub: Kund village (Rewari).",
        "29. Only District with HQ Outside Name: Mahendragarh (HQ at Narnaul).",
        "30. Beti Bachao Beti Padhao Launched on: 22 Jan 2015 from Panipat."
    ]
    builder._create_card(s65, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Rapid-Fire Checkpoints: 16 to 30", c2, builder.COLOR_GOLD)

    print("Districts & Traps slides (61-65) built successfully.")
