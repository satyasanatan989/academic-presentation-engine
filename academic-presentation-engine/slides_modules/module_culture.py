"""
Module: Folk Theatre, Dances, Music, Attire, Fairs & Sports (Slides 57 - 60)
Author: Academic Research Presentation Engine (ARPE)
Target: Haryana FSO & State Screening Tests (HPSC / HSSC)
"""

from pptx.util import Inches

def build_culture_slides(builder):
    prs = builder.prs
    blank = builder.blank_layout

    # -------------------------------------------------------------------------
    # Slide 57: Folk Theatre (Saang / Swang) & Literary Masters
    # -------------------------------------------------------------------------
    s57 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s57)
    builder._add_header_and_footer(
        s57,
        tag="FOLK THEATRE (SAANG)",
        title="Folk Theatre (Saang): Kishan Lal Bhat, Deep Chand & Dada Lakhmi Chand",
        takeaway="Deep Chand Bahman is the 'Kalidasa / Shakespeare of Haryana'; Lakhmi Chand is 'Surya Kavi'.",
        citation="Haryana Sahitya Akademi; Haryana Granth Akademi; Folk Cultural Studies.",
        slide_num=57
    )
    headers = ["Folk Theatre Master / Legend", "Native Village & Period", "Artistic Title / Honor", "Pioneering Contribution & Famous Compositions"]
    rows = [
        ["Kishan Lal Bhat", "Meerut (Active c. 1730 CE)", "Father / Originator of Saang format", "Formalized open-air musical dramatic performance staged on wooden platform (Takhtaposh) with Sarangi and Nagada accompaniment"],
        ["Deep Chand Bahman", "Khanda village (Sonipat district)\nActive WWI era", "'Kalidasa of Haryana'\n'Shakespeare of Haryana'", "Modernized Haryanvi Saang; introduced disciplined orchestra, musical meters (Alha, Chhand); recruited youth during WWI; authored 'Raja Harishchandra', 'Gopichand'"],
        ["Pandit Lakhmi Chand ('Dada Lakhmi')", "Jatti Kalan village (Sonipat)\n(1903 - 1948)", "'Surya Kavi' (सूर्य कवि)\n'Sufi Saint of Haryana'", "Unlettered genius; composed over 25 classic Saangs (Nal-Damyanti, Satyavan-Savitri, Seth Taramal, Puranmal) and 500+ philosophical Ragnis; Research Chair in MDU Rohtak"],
        ["Pandit Mange Ram", "Sisana village (Sonipat district)", "Chief Disciple of Lakhmi Chand", "Authored acclaimed poetic Saangs: 'Pingla-Bhartari', 'Dushyant-Shakuntala'; mastered classical meter"],
        ["Baje Bhagat", "Sisana village (Sonipat district)", "Renowned Ragni & Bhajan Master", "Authored popular rural folk compositions: 'Raja Mordhwaj', 'Heer Ranjha'; moral aphorisms"],
        ["Ahmad Baksh Thanesari", "Thanesar (Kurukshetra)", "Pioneering Epic Adaptor", "Translated Valmiki Ramayana into Haryanvi Saang format; authored 'Gopi Chand' and 'Padmavat'"]
    ]
    widths = [Inches(2.5), Inches(2.2), Inches(2.3), Inches(4.733)]
    builder._create_table(s57, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 58: Folk Dances of Haryana
    # -------------------------------------------------------------------------
    s58 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s58)
    builder._add_header_and_footer(
        s58,
        tag="PERFORMING ARTS",
        title="Folk Dances of Haryana: Dhamal, Phag, Khoria, Loor & Ratvai",
        takeaway="Dhamal dates back to Mahabharata; Khoria is performed when groom departs; Loor means 'girl'.",
        citation="North Zone Cultural Centre (NZCC); Department of Art & Cultural Affairs, Haryana.",
        slide_num=58
    )
    headers = ["Folk Dance Form", "Cultural Occasion / Season", "Gender / Performer Type", "Distinctive Feature & Screening Trap"]
    rows = [
        ["1. Dhamal Dance (धमाल)", "Post-harvest moonlit nights; Spring", "Exclusively Men (Ahirwal/Mahendragarh/Jhajjar)", "Oldest dance form dating back to Mahabharata era; performed with Daph, Dholak, and wooden Chimta to celebrate good crop"],
        ["2. Phag Dance (फाग)", "Phalgun month (around Holi festival)", "Men and women together (mixed)", "Celebratory rhythm; witty banter with colored water sprinkling and rhythmic clapping"],
        ["3. Khoria Dance (खोरिया)", "Wedding night after Groom's Barat departs", "Exclusively Women in domestic courtyard", "Parody and joyful drama performed by women waiting for bride's arrival; men strictly prohibited"],
        ["4. Loor Dance (लूर)", "Holi season in Bangru linguistic belt", "Exclusively Girls / Young Women", "'Loor' means girl in Bangru dialect; question-answer lyric format between two facing lines of dancers"],
        ["5. Chhathi Dance (छठी)", "6th day after the birth of male child", "Women of the household & neighborhood", "Devotional thanksgiving to Bemata; boiled gram (Bakli / बाकली) distributed as prasad"],
        ["6. Gugga Dance (गुग्गा)", "Bhadrapada month (Gugga Navami)", "Male devotees / Bhagats", "Religious ritual dance; devotees self-flagellate with iron chains (Chhad/छड़) to honor snake deity Gugga Pir"],
        ["7. Ratvai Dance (रतवाई)", "Monsoon season in Mewat region", "Meo men and women", "Accompanied by Deru (small drum) and Bansuri; fast swirling circular formations"],
        ["8. Daph Dance (डफ)", "Spring harvest / Baisakhi celebration", "Male dancers playing large circular frame drums", "First state folk dance officially represented in Republic Day Parade (1969) under CM Bansi Lal"]
    ]
    widths = [Inches(2.4), Inches(2.3), Inches(2.8), Inches(4.233)]
    builder._create_table(s58, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 59: Traditional Attire, Folk Instruments & Fairs
    # -------------------------------------------------------------------------
    s59 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s59)
    builder._add_header_and_footer(
        s59,
        tag="MATERIAL CULTURE",
        title="Traditional Attire, Folk Musical Instruments & Iconic Fairs",
        takeaway="Surajkund Mela started in 1987; Khandwa is the turban of rural honor; Daaman is the 52-yard skirt.",
        citation="Surajkund Mela Authority; Haryana Tourism; State Cultural Affairs Directorate.",
        slide_num=59
    )
    c1 = [
        "Male Attire: Khandwa (खंडवा - ceremonial turban symbolizing dignity and respect), Pagri / Safa, Kurta, and Dhoti.",
        "Female Attire: Daaman (दामण - pleated flared skirt traditionally made from 52 yards / बावन गज of heavy fabric), Kurti (fitted blouse), and Chundri (चूंदड़ी - printed veil).",
        "Traditional Ornaments:",
        "• Head & Forehead: Borla, Tika, Shringarpatti.",
        "• Nose: Nath, Koka, Purli.",
        "• Neck: Hansli (solid silver torque), Kathla, Galshree, Haar.",
        "• Hands & Waist: Pahunchi, Chhan-Kangan, Tagdi (silver waist chain).",
        "• Feet: Kari, Ramjhol, Pajeb."
    ]
    builder._create_card(s59, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Traditional Attire & Folk Ornaments", c1, builder.COLOR_COBALT)

    c2 = [
        "Folk Musical Instruments: Ektara, Sarangi (associated with Saang), Been (snake charmer flute), Algoza (twin flutes), Deru, Nagada, Tasha, Khartal, and Chimta.",
        "Surajkund International Crafts Mela (Faridabad):",
        "• Annual fair held in February since 1987 in Surajkund Aravalli complex.",
        "• Features designated Theme State and Partner Country to promote rural crafts.",
        "International Gita Mahotsav: Held annually in Kurukshetra (Brahma Sarovar) on Shukla Ekadashi of Margashirsha (Mokshada Ekadashi).",
        "Other Iconic Fairs: Kapal Mochan Mela (Yamunanagar), Sheetla Mata Mela (Gurugram), Falgu Mela (Pharal, Kaithal - Somvati Amavasya), Mango Mela (Pinjore Gardens, July)."
    ]
    builder._create_card(s59, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Musical Instruments & Landmark Fairs", c2, builder.COLOR_GOLD)

    # -------------------------------------------------------------------------
    # Slide 60: Sports Capital of India: Olympic Medals & Cash Policy
    # -------------------------------------------------------------------------
    s60 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s60)
    builder._add_header_and_footer(
        s60,
        tag="SPORTS POWERHOUSE",
        title="Sports Capital of India: Olympic Medal Factory & Cash Award Policy",
        takeaway="Haryana awards India's highest Olympic cash prizes (Gold: 6 Cr, Silver: 4 Cr, Bronze: 2.5 Cr).",
        citation="Department of Sports & Youth Affairs, Haryana; Sports Policy 2015 / 2021; Indian Olympic Association.",
        slide_num=60
    )
    headers = ["Sports Discipline", "Legendary Champions from Haryana", "Native Village / District", "Historic Milestone / Olympic Achievement"]
    rows = [
        ["Athletics (Javelin Throw)", "Subedar Major Neeraj Chopra (VSM)", "Khandra village (Panipat district)", "Independent India's 1st Track & Field Olympic Gold (Tokyo 2020: 87.58m); Olympic Silver (Paris 2024: 89.45m); World Champion"],
        ["Wrestling (Kushti)", "Sakshi Malik, Bajrang Punia,\nPhogat Sisters, Ravi Dahiya, Aman Sehrawat", "Mokhra (Rohtak), Khudan (Jhajjar),\nBalali (Charkhi Dadri), Nahri (Sonipat)", "Sakshi: 1st Indian female wrestler Olympic medal (Rio 2016 Bronze); Dahiya: Tokyo 2020 Silver; Aman: Paris 2024 Bronze"],
        ["Shooting", "Manu Bhaker", "Goria village (Jhajjar district)", "First Indian athlete post-independence to win 2 medals at a single Olympic Games (Paris 2024: 10m Air Pistol individual & mixed team)"],
        ["Boxing", "Vijender Singh, Amit Panghal,\nNitu Ghanghas", "Kaluwas (Bhiwani - 'Mini Cuba'),\nMayna (Rohtak), Dhanana (Bhiwani)", "Vijender: India's 1st Olympic Boxing medal (Beijing 2008 Bronze); Bhiwani Boxing Club (BBC) founded by Coach Jagdish Singh"],
        ["Field Hockey", "Rani Rampal, Savita Punia,\nSandeep Singh, Surinder Kaur", "Shahbad Markanda (Kurukshetra district)", "Shahbad is renowned as the 'Nursery / Cradle of Indian Women's Hockey' & 'Sansarpur of Haryana'"]
    ]
    widths = [Inches(2.5), Inches(2.6), Inches(2.8), Inches(3.833)]
    builder._create_table(s60, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    print("Culture & Sports slides (57-60) built successfully.")
