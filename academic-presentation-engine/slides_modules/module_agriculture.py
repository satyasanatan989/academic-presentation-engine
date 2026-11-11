"""
Module: Agriculture, Dairying, Food Technology & Institutes (Slides 46 - 51)
Author: Academic Research Presentation Engine (ARPE)
Target: Haryana FSO & State Screening Tests (HPSC / HSSC) - SPECIAL CORE FOCUS
"""

from pptx.util import Inches

def build_agriculture_slides(builder):
    prs = builder.prs
    blank = builder.blank_layout

    # -------------------------------------------------------------------------
    # Slide 46: Agrarian Structure & Cropping Seasons
    # -------------------------------------------------------------------------
    s46 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s46)
    builder._add_header_and_footer(
        s46,
        tag="FSO AGRONOMY CORE",
        title="Agrarian Profile: Cropping Seasons, Land Utilization & Production",
        takeaway="Haryana contributes ~15% of wheat & ~10% of rice to the National Central Foodgrain Pool.",
        citation="Directorate of Agriculture & Farmers Welfare, Haryana; State Statistical Abstract 2023-24.",
        slide_num=46
    )
    c1 = [
        "Net Sown Area: ~3.52 million hectares (>80% of total geographical area), reflecting extraordinary land cultivation efficiency (national average ~43%).",
        "Cropping Intensity: Approximately 185% - 190%, driven by intensive tubewell and canal irrigation networks.",
        "Total Foodgrain Production: Surpasses 180-185 lakh metric tonnes annually.",
        "National Contribution: Despite occupying only 1.34% of India's landmass, Haryana is the 2nd largest contributor of wheat and 3rd largest of rice to the Central Pool (PDS / FCI).",
        "Gross Irrigated Area: ~89% of gross cropped area is under assured irrigation (Tubewells ~55%, Canals ~44%)."
    ]
    builder._create_card(s46, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Land Utilization & Grain Dominance", c1, builder.COLOR_COBALT)

    c2 = [
        "Rabi Season (Winter / Asadhi - असाढ़ी):",
        "• Sowing: Oct - Dec; Harvesting: April - May.",
        "• Major Crops: Wheat (Gehun), Mustard (Sarson), Gram (Chana), Barley (Jau), Potato.",
        "Kharif Season (Monsoon / Sawani - सावणी):",
        "• Sowing: June - July; Harvesting: Sep - Nov.",
        "• Major Crops: Paddy (Dhan), Cotton (Narma/Kapas), Pearl Millet (Bajra), Sugarcane, Maize.",
        "Zaid Season (Summer / Jayad):",
        "• Sowing: March - June (between Rabi and Kharif).",
        "• Major Crops: Fodder (Jowar/Barseem), Watermelon, Muskmelon, Cucumber, Seasonal vegetables."
    ]
    builder._create_card(s46, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Three Cropping Seasons of Haryana", c2, builder.COLOR_GOLD)

    # -------------------------------------------------------------------------
    # Slide 47: Crop Specialization: Basmati Rice, Wheat, Mustard & Bajra
    # -------------------------------------------------------------------------
    s47 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s47)
    builder._add_header_and_footer(
        s47,
        tag="CROP GEOGRAPHY",
        title="Crop Belts: Taraori Basmati Rice, Wheat, Mustard, Bajra & Cotton",
        takeaway="Karnal is the 'Rice Bowl of Haryana'; Taraori Basmati accounts for over 60% of India's Basmati exports.",
        citation="Agricultural & Processed Food Products Export Development Authority (APEDA); HAFED.",
        slide_num=47
    )
    headers = ["Crop Category", "Leading Districts", "Key Agro-Economic Feature", "FSO Food Science / Quality Standard"]
    rows = [
        ["Basmati Rice (Dhan)", "Karnal, Kaithal, Kurukshetra, Panipat (Rice Bowl)", "Taraori Basmati (HBC-19) world famous for aroma & elongation; accounts for >60% of national Basmati exports", "Geographical Indication (GI) tagged; 2-acetyl-1-pyrroline (2-AP) compound confers distinctive aroma"],
        ["Wheat (Gehun)", "Sirsa, Jind, Hisar, Karnal, Fatehabad", "Sirsa leads in absolute wheat output; high-yielding varieties WH-1105, HD-2967, HD-3086 dominant", "Subject to moisture limit (<12%) and foreign matter inspection under FSSAI cereals regulation"],
        ["Rapeseed & Mustard (Rai/Sarson)", "Mahendragarh, Bhiwani, Rewari, Charkhi Dadri", "Mahendragarh is the premier mustard producing hub; cold-pressed oil extraction in local expellers", "Erucic acid regulation (<2.0% in canola types; FSSAI blending limits); pungency due to allyl isothiocyanate"],
        ["Pearl Millet (Bajra)", "Bhiwani, Mahendragarh, Rewari, Hisar", "Staple dryland cereal; highly drought-resilient; Haryana covers Bajra under Bhavantar Bharpayee Yojana", "Rich in iron (Fe) and dietary fiber; classified as Nutri-Cereal under International Year of Millets"],
        ["Cotton (Narma / Kapas)", "Sirsa, Fatehabad, Hisar, Jind ('White Gold')", "Sirsa is the undisputed cotton capital of Haryana; Bt-Cotton dominates >95% acreage", "Cottonseed oil is refined for edible consumption; rich in tocopherols (Vitamin E)"],
        ["Sugarcane (Ganna)", "Yamunanagar, Karnal, Ambala, Sonipat", "Yamunanagar leads in production; Saraswati Sugar Mills (Yamunanagar) is one of Asia's largest mills", "Sucrose recovery percentage (~9.5% - 10.5%); jaggery (Gur) manufacturing cluster at Shahbad"]
    ]
    widths = [Inches(2.2), Inches(2.8), Inches(3.5), Inches(3.233)]
    builder._create_table(s47, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 48: Livestock & Dairying: Murrah Buffalo & Hariana Cattle
    # -------------------------------------------------------------------------
    s48 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s48)
    builder._add_header_and_footer(
        s48,
        tag="LIVESTOCK EXCELLENCE",
        title="Murrah Buffalo ('Black Gold') & Indigenous Hariana Cattle",
        takeaway="Murrah buffalo is Haryana's 'Black Gold' (Kala Sona); Jind is known as the 'Home of Murrah'.",
        citation="National Bureau of Animal Genetic Resources (NBAGR), Karnal; Animal Husbandry Department, Haryana.",
        slide_num=48
    )
    c1 = [
        "Origin & Homeland: Rohtak, Jind, Hisar, and Fatehabad districts; Jind is reverently dubbed 'The Capital / Home of Murrah'.",
        "Physical Conformation: Jet-black color, tightly curled spiral horns (like जलेबी), wedge-shaped body, short legs, thin neck, and deep capacious udder.",
        "Exceptional Milk Yield: 2,000 - 3,500 kg per lactation; champions exceed 25-30 liters/day.",
        "High Milk Fat Content: Average fat 7.0% - 8.5%; preferred for ghee, paneer, and mawa production.",
        "Global Sire Export: Murrah bulls exported globally (Brazil, Italy, Bulgaria, Philippines) to upgrade native water buffalo stocks.",
        "Economic Title: Reverently hailed as 'Black Gold' (काला सोना) and 'The Tractor of Rural Economy'."
    ]
    builder._create_card(s48, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Murrah Buffalo ('Black Gold')", c1, builder.COLOR_COBALT)

    c2 = [
        "Hariana Cattle Breed (Bos indicus):",
        "• Native dual-purpose cattle breed originating in Rohtak, Hisar, Jind, and Gurugram.",
        "• Bullocks renowned nationwide for unmatched draught power, endurance, and swift road transport.",
        "• Coat: White or light grey; long narrow face, flat forehead with convex center.",
        "Sahiwal & Tharparkar Propagation:",
        "• Promoted under Rashtriya Gokul Mission; Sahiwal cattle breeding farm at Government Livestock Farm (GLF) Hisar.",
        "Government Livestock Farm (GLF), Hisar:",
        "• One of the largest livestock breeding farms in Asia; established in 1809 as a camel breeding stud by Major James Lumsdaine.",
        "Gau Seva Aayog: Statutory Haryana Gau Seva Aayog headquartered at Panchkula regulating ~600+ Gaushalas."
    ]
    builder._create_card(s48, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Hariana Cattle & GLF Hisar", c2, builder.COLOR_GOLD)

    # -------------------------------------------------------------------------
    # Slide 49: Premier Dairy & Veterinary Research Institutes
    # -------------------------------------------------------------------------
    s49 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s49)
    builder._add_header_and_footer(
        s49,
        tag="RESEARCH ECOSYSTEM",
        title="Premier Research Institutes: NDRI, CIRB, LUVAS & NBAGR",
        takeaway="NDRI shifted from Bangalore to Karnal in 1955; cloned world-famous buffalo 'Garima' & 'Pratham'.",
        citation="ICAR-NDRI Karnal; ICAR-CIRB Hisar; LUVAS Hisar; ICAR-NBAGR Karnal.",
        slide_num=49
    )
    headers = ["Research Institute", "Headquarters", "Establishment Year", "Pioneering Research Achievements (FSO Must-Know)"]
    rows = [
        ["ICAR - National Dairy Research Institute (NDRI)", "Karnal", "Est. 1923 (Bangalore)\nShifted to Karnal: 1955", "Deemed University; developed Karan-Swiss and Karan-Fries cattle crossbreeds; cloned buffalo 'Pratham' (1990) & 'Garima' (2009) using Hand-Guided Cloning; developed rapid chemical test kits for milk adulteration (detergents, starch, urea)"],
        ["ICAR - Central Institute for Research on Buffaloes (CIRB)", "Hisar (Sirsa Road)", "1 February 1985", "Premier institute for buffalo genetics, embryo transfer technology (ETT), and cryopreservation of semen; cloned superior Murrah bull 'Hisar-Gaurav' and 'Assamese' buffalo"],
        ["Lala Lajpat Rai University of Vet & Animal Sciences (LUVAS)", "Hisar", "2010 (Carved from CCSHAU)", "State veterinary university; veterinary pathology, microbiology, zoonotic diseases, and animal disease diagnostic laboratories"],
        ["ICAR - National Bureau of Animal Genetic Resources (NBAGR)", "Karnal", "1984 (Shifted to Karnal 1985)", "Nodal national agency for registration and characterization of new indigenous livestock and poultry breeds in India"],
        ["Central Sheep Breeding Farm (CSBF)", "Hisar", "1969-70", "Premier central farm breeding high-quality Corriedale and Sonadi rams to boost mutton and fine wool production"]
    ]
    widths = [Inches(3.0), Inches(1.8), Inches(2.2), Inches(4.733)]
    builder._create_table(s49, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 50: Food Technology & Processing: NIFTEM, Mega Food Parks & HAFED
    # -------------------------------------------------------------------------
    s50 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s50)
    builder._add_header_and_footer(
        s50,
        tag="FOOD PROCESSING INFRASTRUCTURE",
        title="Food Science Hubs: NIFTEM Kundli, Mega Food Parks & HAFED",
        takeaway="NIFTEM Kundli (Sonipat) is an Institute of National Importance under Ministry of Food Processing Industries.",
        citation="Ministry of Food Processing Industries (MoFPI); HAFED Annual Report; NIFTEM Act, 2021.",
        slide_num=50
    )
    c1 = [
        "National Institute of Food Technology Entrepreneurship & Management (NIFTEM):",
        "• Located at Kundli, Sonipat district (adjacent to Delhi border on NH-44).",
        "• Established by Ministry of Food Processing Industries (MoFPI); declared an 'Institute of National Importance' by Parliament under NIFTEM Act 2021.",
        "• State-of-the-art Pilot Plants: Dairy, Fruits & Vegetables, Meat & Poultry, Cereal & Extrusion, and Packaging Technology.",
        "• Certified Food Testing Laboratories: NABL-accredited and FSSAI-notified referral laboratory testing pesticide residues, mycotoxins, heavy metals, and adulterants.",
        "• Incubation & Entrepreneurship: Provides testing and pilot-scale scaleup facilities for food processing MSMEs across North India."
    ]
    builder._create_card(s50, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "NIFTEM Kundli: Academic Epicenter", c1, builder.COLOR_COBALT)

    c2 = [
        "Mega Food Parks & Industrial Hubs:",
        "• Mega Food Park at Rai, Sonipat: Houses large-scale fruit and vegetable dehydration, cold storage, and tetra-pack lines.",
        "• Mega Food Park at Barhi, Sonipat: Focuses on bakery, confectionery, and ready-to-eat (RTE) processing.",
        "• Mega Food Park at Dabwali, Sirsa: Promoted by private/cooperative sector for kinnow and citrus processing.",
        "• HAFED Mega Food Park at IMT Rohtak: Advanced milling, cold chains, and multi-crop automated packaging.",
        "HAFED (Haryana State Co-op Supply & Marketing Fed):",
        "• Largest apex cooperative body in Haryana (Est. 1 Nov 1966).",
        "• Operates automatic basmati rice mills at Taraori, Ding, and Rania; cold-pressed mustard oil mills at Rewari and Narnaul; cattle feed plants at Rohtak and Saktakhera."
    ]
    builder._create_card(s50, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Mega Food Parks & HAFED Cooperative Network", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 51: Indo-Israel Centres of Excellence in Horticulture
    # -------------------------------------------------------------------------
    s51 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s51)
    builder._add_header_and_footer(
        s51,
        tag="HORTICULTURE INNOVATION",
        title="Indo-Israel Centres of Excellence: Precision Farming & Processing",
        takeaway="Gharaunda (Karnal) houses the Centre of Excellence for Vegetables; Mangiana (Sirsa) for Fruits.",
        citation="Mission for Integrated Development of Horticulture (MIDH); Indo-Israel Agricultural Project (IIAP).",
        slide_num=51
    )
    headers = ["Centre of Excellence (CoE)", "District Location", "Focal Commodity", "Advanced Technology Implemented"]
    rows = [
        ["1. CoE for Vegetables", "Gharaunda (Karnal)", "Tomato, Capsicum, Cucumber, Seedling nursery", "Naturally ventilated polyhouses, climate-controlled hi-tech nurseries, soilless media (cocopeat), fertigation drip lines"],
        ["2. CoE for Fruits", "Mangiana (Sirsa district)", "Kinnow, Pomegranate, Guava, Olive", "Canopy management, high-density meadow orchards (guava), micro-irrigation, post-harvest fruit waxing and sorting"],
        ["3. CoE for Sub-Tropical Fruits", "Ladwa (Kurukshetra)", "Mango, Pear, Litchi, Plum", "Elite mother block preservation, rejuvenation of senile orchards, fruit fly trap management"],
        ["4. CoE for Integrated Beekeeping", "Ramnagar (Kurukshetra)", "Honey & Bee Products (Apiculture)", "Scientific queen bee rearing, honey processing and extraction facility, automated wax comb foundation plant"],
        ["5. CoE for Citrus Fruits", "Bhuna (Fatehabad district)", "Sweet Orange (Malta), Lime, Kinnow", "Viral-free certified citrus rootstock propagation, drip-fertigation monitoring"],
        ["6. CoE for Dairy / Protected Cultivation", "Hisar (CCSHAU campus) & Hodal (Palwal)", "Dairy automation & Floriculture", "Hydroponic automated fodder production, protected cultivation of gerbera and rose cut-flowers"]
    ]
    widths = [Inches(3.2), Inches(2.2), Inches(3.0), Inches(3.333)]
    builder._create_table(s51, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    print("Agriculture slides (46-51) built successfully.")
