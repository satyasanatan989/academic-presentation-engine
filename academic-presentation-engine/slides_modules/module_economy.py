"""
Module: Industries, Clusters, Minerals & Demographics (Slides 52 - 56)
Author: Academic Research Presentation Engine (ARPE)
Target: Haryana FSO & State Screening Tests (HPSC / HSSC)
"""

from pptx.util import Inches

def build_economy_slides(builder):
    prs = builder.prs
    blank = builder.blank_layout

    # -------------------------------------------------------------------------
    # Slide 52: Major Industrial Belts: Gurugram, Faridabad, Panipat & Yamunanagar
    # -------------------------------------------------------------------------
    s52 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s52)
    builder._add_header_and_footer(
        s52,
        tag="INDUSTRIAL GEOGRAPHY",
        title="Industrial Powerhouses: Gurugram, Faridabad, Panipat & Yamunanagar",
        takeaway="Gurugram produces ~50% of India's passenger cars; Panipat is the 'Textile / Weaver City'.",
        citation="Department of Industries & Commerce, Haryana; Haryana State Industrial Development Corp (HSIIDC).",
        slide_num=52
    )
    headers = ["Industrial City / Hub", "Key Economic Identity", "Flagship Companies / Units", "Key Manufacturing Output"]
    rows = [
        ["Gurugram (Millennium City)", "Automobile & IT Software Capital of North India", "Maruti Suzuki (Manesar plant), Hero MotoCorp, DLF Cyber City, Google, Microsoft (>500 Fortune 500 offices)", "Passenger cars, two-wheelers, financial technology, enterprise software, auto ancillary parts"],
        ["Faridabad (Industrial Hub)", "Oldest & largest heavy manufacturing city in Haryana", "Escorts Group, JCB India, Yamaha Motors, Whirlpool, Goodyear Tyres, Havells", "Tractors, earthmoving machinery, refrigerators, industrial rubber tyres, electrical switchgears"],
        ["Panipat (Textile / Weaver City)", "'City of Weavers' & National Handloom Capital", "IOCL Panipat Refinery (largest in South Asia), National Fertilizer Limited (NFL), Shodapur loom units", "Military acrylic blankets (supplies 75% to armed forces), recycled yarn (shoddy wool), carpets, petrochemicals"],
        ["Yamunanagar (Paper / Timber City)", "National Timber Market & Paper Hub (Abdullapur Mandi)", "Ballarpur Industries Limited (BILT), Saraswati Sugar Mills, Jagadhri Railway Workshop (Est. 1952)", "Writing and printing paper, plywood (supplies 60% of India's plywood), railway wagons, refined sugar"]
    ]
    widths = [Inches(2.8), Inches(2.8), Inches(3.2), Inches(2.933)]
    builder._create_table(s52, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 53: Specialized Industrial Clusters
    # -------------------------------------------------------------------------
    s53 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s53)
    builder._add_header_and_footer(
        s53,
        tag="SPECIALIZED CLUSTERS",
        title="Specialized Manufacturing Clusters: Rewari, Ambala, Bahadurgarh & Kundli",
        takeaway="Rewari is the 'Brass City' (Pital Nagari); Ambala manufactures ~35% of India's scientific apparatus.",
        citation="Micro, Small & Medium Enterprises (MSME) Development Institute; Haryana Enterprise Promotion Board.",
        slide_num=53
    )
    c1 = [
        "Rewari ('Brass City' / पीतल नगरी): Century-old artisanal brass casting and sheet rolling; prominent manufacturer of brass and copper utensils, brass electrical components, and Tilla-jooti (embroidered footwear).",
        "Jagadhri (Yamunanagar): Renowned for stainless steel and brass household utensils; extensive rolling mills.",
        "Ambala ('Scientific City' & 'Mixie City'): Accounts for ~35% - 40% of India's total scientific laboratory equipment (microscopes, glassware, physics apparatus); also famous for domestic kitchen mixer-grinder manufacturing.",
        "Bahadurgarh ('Gateway of Haryana', Jhajjar): Major footwear hub in India (Relaxo, Action, Aero shoes); sanitaryware capital (Hindware / HSIL plants)."
    ]
    builder._create_card(s53, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Rewari, Ambala & Bahadurgarh", c1, builder.COLOR_COBALT)

    c2 = [
        "IMT Manesar (Gurugram): Integrated Model Township spread over 4,500 acres; houses Maruti Suzuki, Denso, Honda Motorcycle, and National Brain Research Centre (NBRC).",
        "IMT Bawal (Rewari): Heavy automotive component and Japanese industrial corridor (JETRO) cluster on NH-48.",
        "IMT Kharkhoda (Sonipat): Massive 3,200-acre upcoming mega-manufacturing zone; new mega-plant of Maruti Suzuki (capacity 10 lakh cars/year) and Suzuki two-wheelers.",
        "Kundli & Rai (Sonipat): Food processing, stainless steel fabrication, and plastics manufacturing.",
        "Dharuhera (Rewari): Major automobile and industrial hub (Hero MotoCorp flagship assembly plant)."
    ]
    builder._create_card(s53, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "IMT Corridors: Manesar, Bawal & Kharkhoda", c2, builder.COLOR_GOLD)

    # -------------------------------------------------------------------------
    # Slide 54: Mineral Resources of Haryana
    # -------------------------------------------------------------------------
    s54 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s54)
    builder._add_header_and_footer(
        s54,
        tag="MINERAL ASSETS",
        title="Mineral Resources of Haryana: Mahendragarh Dominance & Distribution",
        takeaway="Mahendragarh is the undisputed mineral capital of Haryana, holding deposits of iron, slate & calcite.",
        citation="Department of Mines & Geology, Haryana; Indian Minerals Yearbook (IBM).",
        slide_num=54
    )
    headers = ["Mineral Commodity", "Primary District / Geological Belt", "Key Mining Localities", "Industrial Application / Exam Fact"]
    rows = [
        ["Iron Ore (Magnetite/Hematite)", "Mahendragarh district", "Antri-Beharipur, Dhanota-Dhancholi, Chhapra", "Low to medium grade deposits associated with Aravalli quartzite ridges"],
        ["Slate Stone (स्लेट पत्थर)", "Rewari & Mahendragarh", "Kund village (Rewari district) - World famous", "Kund slate is exported globally for architectural roofing, paving, and chalkboards"],
        ["Limestone & Kankar", "Panchkula (Morni), Rohtak, Bhiwani", "Malla, Kharag, and Tundapathar (Panchkula)", "Feedstock for Surajpur Cement Factory (ACC - historical) and Charkhi Dadri cement plant"],
        ["Marble & Calcite", "Mahendragarh (Narnaul)", "Bayal, Musnota, Mandi, Sirohi", "Variegated marble (black, white, green banded) and pure crystalline calcite for filler industries"],
        ["Quartz & Silica Sand", "Gurugram, Faridabad, Mahendragarh", "Aravalli ridges, Badarpur sand pits (Faridabad)", "Glass manufacturing and construction concrete casting"],
        ["China Clay (Kaolin)", "Gurugram district", "Sikanderpur and Alwar border", "Sanitaryware, ceramic tiles, and porcelain insulators"],
        ["Flexible Sandstone (कम्पा पत्थर)", "Charkhi Dadri (Kaliana village)", "Kaliana hillock", "Unique geological curiosity: naturally bending/flexible stone (Itacolumite) found only here in India!"]
    ]
    widths = [Inches(2.5), Inches(2.2), Inches(3.2), Inches(3.833)]
    builder._create_table(s54, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 55: Census 2011 Master Demographics
    # -------------------------------------------------------------------------
    s55 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s55)
    builder._add_header_and_footer(
        s55,
        tag="DEMOGRAPHIC PROFILE",
        title="Census 2011 Master Profile: Population, Density, Growth & Literacy",
        takeaway="Total population: 2,53,51,462 (2.09% of India); Faridabad is most dense (2,442/sq km).",
        citation="Census of India 2011, Directorate of Census Operations, Haryana.",
        slide_num=55
    )
    headers = ["Census Demographic Parameter", "Absolute Value / Figure", "Highest District (Maximum)", "Lowest District (Minimum)"]
    rows = [
        ["Total Population", "2,53,51,462 (2.53 Crore)\n(18th most populous in India)", "Faridabad (18,09,733)\nfollowed by Hisar", "Panchkula (5,61,293)\nfollowed by Charkhi Dadri (est) / Fatehabad"],
        ["Decadal Growth Rate (2001-2011)", "19.90% (National: 17.70%)", "Gurugram (+73.14%)\n(Massive urban migration)", "Jhajjar (+8.90%)\nfollowed by Ambala (+11.23%)"],
        ["Population Density", "573 persons / sq km\n(National: 382/sq km)", "Faridabad (2,442 persons/sq km)\nfollowed by Gurugram (1,204)", "Sirsa (303 persons/sq km)\n(Large arid area, low settlement)"],
        ["Overall Literacy Rate", "75.55% (National: 74.04%)\nMale: 84.06% | Female: 65.94%", "Gurugram (84.70%)\nfollowed by Panchkula (81.88%)", "Nuh / Mewat (54.08%)\nfollowed by Fatehabad (67.92%)"],
        ["Male Literacy", "84.06%", "Rewari (91.44%)\nfollowed by Mahendragarh (89.72%)", "Nuh / Mewat (69.94%)\nfollowed by Fatehabad (76.14%)"],
        ["Female Literacy", "65.94%", "Gurugram (77.98%)\nfollowed by Panchkula (75.99%)", "Nuh / Mewat (36.60%)\n(Severe socio-educational deficit)"],
        ["Urban-Rural Spread", "Rural: 65.12% (1.65 Cr)\nUrban: 34.88% (88.42 Lakh)", "Most Urbanized: Faridabad (79.51%)\nfollowed by Gurugram (68.82%)", "Least Urbanized: Nuh / Mewat (11.39%)\nfollowed by Mahendragarh (14.41%)"]
    ]
    widths = [Inches(3.0), Inches(2.8), Inches(3.0), Inches(2.933)]
    builder._create_table(s55, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 56: Sex Ratio Dynamics & 'Beti Bachao, Beti Padhao'
    # -------------------------------------------------------------------------
    s56 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s56)
    builder._add_header_and_footer(
        s56,
        tag="GENDER DYNAMICS",
        title="Sex Ratio Dynamics, Child Sex Ratio & Beti Bachao Beti Padhao",
        takeaway="Overall Sex Ratio in 2011 was 879 (lowest in India); PM launched 'BBBP' from Panipat in 2015.",
        citation="Ministry of Women & Child Development (MoWCD); Health Department, Haryana (Civil Registration System).",
        slide_num=56
    )
    c1 = [
        "2011 Overall Sex Ratio: 879 females per 1,000 males (lowest among all 28 Indian states in Census 2011; national average was 943).",
        "Highest Sex Ratio (2011): Nuh / Mewat (907 females/1,000 males), followed by Fatehabad (902) and Rewari (898).",
        "Lowest Sex Ratio (2011): Gurugram (854 females/1,000 males) and Sonipat (856), reflecting severe urban selective abortions.",
        "Child Sex Ratio (0-6 Years, 2011): A catastrophic 834 females per 1,000 males (national average 918); Jhajjar (782) and Mahendragarh (775) were India's worst-affected districts.",
        "Screening Question: The district with highest overall sex ratio in 2011 is Nuh (907), while lowest is Gurugram (854)."
    ]
    builder._create_card(s56, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Census 2011 Sex Ratio Realities", c1, builder.COLOR_CRIMSON)

    c2 = [
        "Historic Launch of 'Beti Bachao, Beti Padhao' (BBBP):",
        "• Launched by Prime Minister Narendra Modi on 22 January 2015 from Panipat, Haryana.",
        "• Joint flagship mission of Ministry of Women & Child Development, Ministry of Health, and Ministry of Education.",
        "• Strict Enforcement: Severe crackdown on illegal sex determination under PC-PNDT Act (1994) with special task forces, covert sting operations, and reward schemes for informants.",
        "Aapki Beti Hamari Beti Scheme: Financial incentive deposited into LIC account on the birth of first girl child in SC/BPL families and second girl child in all families.",
        "Civil Registration System (CRS) Resurgence: Sex ratio at birth (SRB) improved significantly to ~915-920+ in recent years."
    ]
    builder._create_card(s56, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Beti Bachao Beti Padhao & Recovery", c2, builder.COLOR_TEAL)

    print("Economy slides (52-56) built successfully.")
