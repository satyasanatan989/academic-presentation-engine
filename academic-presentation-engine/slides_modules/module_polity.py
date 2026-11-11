"""
Module: Polity, Governance, Judiciary & Administration (Slides 41 - 45)
Author: Academic Research Presentation Engine (ARPE)
Target: Haryana FSO & State Screening Tests (HPSC / HSSC)
"""

from pptx.util import Inches

def build_polity_slides(builder):
    prs = builder.prs
    blank = builder.blank_layout

    # -------------------------------------------------------------------------
    # Slide 41: Legislative Framework: Haryana Vidhan Sabha
    # -------------------------------------------------------------------------
    s41 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s41)
    builder._add_header_and_footer(
        s41,
        tag="LEGISLATIVE STRUCTURE",
        title="Haryana Legislative Assembly: Composition, Evolution & Reservation",
        takeaway="Unicameral house of 90 seats; exactly 17 seats are reserved for Scheduled Castes (SC).",
        citation="Constitution of India (Articles 168-177); Haryana Vidhan Sabha Secretariat.",
        slide_num=41
    )
    c1 = [
        "Unicameral Architecture: Single legislative chamber (Haryana Vidhan Sabha) located in the Capitol Complex, Chandigarh (designed by Le Corbusier).",
        "Evolution of Assembly Seats:",
        "• At State Formation (1 Nov 1966): 54 seats (10 reserved for SC).",
        "• First General Elections (March 1967): Increased to 81 seats (15 reserved for SC).",
        "• Delimitation of 1977: Increased to 90 seats (17 reserved for SC); remains constant until 2026 freeze.",
        "Proposed New Assembly Building: 10 acres of land allocated near Railway Station light point in Chandigarh for expanded delimitation capacity (>126 seats)."
    ]
    builder._create_card(s41, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Assembly Structure & Seat Evolution", c1, builder.COLOR_COBALT)

    c2 = [
        "17 Reserved Constituencies (SC):",
        "1. Kalka/Mulana (Ambala)  2. Sadhaura (Yamunanagar)",
        "3. Shahbad (Kurukshetra)  4. Guhla (Kaithal)",
        "5. Nilokheri (Karnal)     6. Israna (Panipat)",
        "7. Kharkhoda (Sonipat)    8. Narwana (Jind)",
        "9. Uklana (Hisar)        10. Bawanikhera (Bhiwani)",
        "11. Kalanaur (Rohtak)    12. Jhajjar (Jhajjar)",
        "13. Bawal (Rewari)       14. Pataudi (Gurugram)",
        "15. Hodal (Palwal)       16. Rania/Kalanwali (Sirsa)",
        "17. Ratia (Fatehabad)",
        "Districts with ZERO Reserved Seats: Panchkula, Charkhi Dadri, Mahendragarh, and Nuh (Mewat)."
    ]
    builder._create_card(s41, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "17 Reserved Constituencies (SC) & Exceptions", c2, builder.COLOR_TEAL)

    # -------------------------------------------------------------------------
    # Slide 42: Parliamentary Representation (Lok Sabha & Rajya Sabha)
    # -------------------------------------------------------------------------
    s42 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s42)
    builder._add_header_and_footer(
        s42,
        tag="PARLIAMENTARY PROFILE",
        title="Parliamentary Representation: 10 Lok Sabha & 5 Rajya Sabha Seats",
        takeaway="Haryana sends 10 MPs to Lok Sabha (2 reserved: Ambala & Sirsa) and 5 MPs to Rajya Sabha.",
        citation="Election Commission of India (ECI); Lok Sabha Secretariat.",
        slide_num=42
    )
    headers = ["Constituency Type", "Total Seats", "Specific Constituencies", "Key Screening & Electoral Details"]
    rows = [
        ["Lok Sabha (Lower House)", "10 Seats", "1. Ambala (SC)  2. Kurukshetra\n3. Karnal       4. Sonipat\n5. Rohtak       6. Sirsa (SC)\n7. Hisar        8. Bhiwani-Mahendragarh\n9. Gurgaon     10. Faridabad", "2 Reserved Seats for SC:\n• Ambala (SC)\n• Sirsa (SC)\nGurgaon is largest electorate; Sonipat & Karnal major agricultural hubs."],
        ["Rajya Sabha (Upper House)", "5 Seats", "Elected by 90 MLAs of Haryana Vidhan Sabha through proportional representation (STV)", "Staggered 6-year tenure; one-third members retire every two years; represents state interest at Center."],
        ["Electoral Demographics", "22 Districts", "Over 2.01 Crore registered voters in Haryana electorate", "Highest voter population in Badshahpur assembly (Gurugram); lowest in Narnaul."],
        ["Electoral Boundaries", "Delimitation 2008", "Merged Mahendragarh with Bhiwani; carved out modern Gurgaon and Faridabad PC", "Assembly segments per Lok Sabha constituency: Exactly 9 Assembly segments each (9 x 10 = 90)."]
    ]
    widths = [Inches(2.5), Inches(1.5), Inches(4.0), Inches(3.733)]
    builder._create_table(s42, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 43: Executive Leadership: Chief Ministers of Haryana
    # -------------------------------------------------------------------------
    s43 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s43)
    builder._add_header_and_footer(
        s43,
        tag="EXECUTIVE LEADERSHIP",
        title="Chief Ministers of Haryana: Historic Tenures & Landmark Milestones",
        takeaway="Pt. BD Sharma was 1st CM; Rao Birender Singh was 1st non-Congress CM; Bhajan Lal served longest.",
        citation="Haryana Vidhan Sabha Secretariat; General Administration Department (GAD).",
        slide_num=43
    )
    headers = ["Chief Minister", "Party / Alignment", "Tenure Milestone", "Key Historical / Exam Fact"]
    rows = [
        ["Pt. Bhagwat Dayal Sharma", "Indian National Congress", "1 Nov 1966 - 23 Mar 1967", "First Chief Minister of Haryana; hailed from Beri (Jhajjar); later Governor of Odisha & MP."],
        ["Rao Birender Singh", "Vishal Haryana Party (VHP)", "24 Mar 1967 - 2 Nov 1967", "First non-Congress CM; first elected male Speaker of Vidhan Sabha; founded regional Vishal Haryana Party (Symbol: Rising Sun)."],
        ["Bansi Lal (3 Non-Consecutive Eras)", "Congress / Haryana Vikas Party", "1968-75, 1985-87, 1996-99\n(~4,268 total days)", "'Architect of Modern Haryana'; 100% rural electrification by Nov 1970; lift irrigation; alcohol prohibition (1996-98)."],
        ["Banarsi Das Gupta", "Congress", "1975-77 & 1990", "Served as Speaker, Deputy CM, and Chief Minister of Haryana."],
        ["Chaudhary Devi Lal ('Tau')", "Janata Party / Lok Dal", "1977-79 & 1987-89", "'King Maker'; twice Deputy Prime Minister of India; introduced Old Age Pension & farm loan waiver."],
        ["Bhajan Lal ('Ph.D of Politics')", "Congress / Haryana Janhit Congress", "1979-86 & 1991-96\n(Longest cumulative tenure: 3,952 days)", "Longest-serving Chief Minister of Haryana; mastered coalition engineering and non-Jat consolidation."],
        ["Om Prakash Chautala", "INLD", "Sworn in 5 times (1989, 90, 90, 91, 1999-2005)", "Highest number of oath ceremonies as CM (5 times); shortest single tenure (6 days in July 1990)."],
        ["Bhupinder Singh Hooda", "Congress", "5 Mar 2005 - 19 Oct 2014\n(Continuous 9 yrs 7 mos)", "Longest continuous single stretch as Chief Minister; IMT Manesar and education hubs."]
    ]
    widths = [Inches(2.5), Inches(2.2), Inches(2.6), Inches(4.433)]
    builder._create_table(s43, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    # -------------------------------------------------------------------------
    # Slide 44: Role of Governors & Political Titans (Devi Lal / Bansi Lal)
    # -------------------------------------------------------------------------
    s44 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s44)
    builder._add_header_and_footer(
        s44,
        tag="GUBERNATORIAL & LEADERSHIP",
        title="Governors of Haryana & Legacies of Devi Lal and Bansi Lal",
        takeaway="Dharma Vira was 1st Governor; B.N. Chakravarty served longest; O.P. Verma served shortest (5 days).",
        citation="Raj Bhavan Haryana; National Portal of India.",
        slide_num=44
    )
    c1 = [
        "First Governor: Shri Dharma Vira (1 Nov 1966 - 14 Sep 1967).",
        "Longest-Serving Governor: B.N. Chakravarty (15 Sep 1967 - 26 Mar 1976; 8+ years; died in office; B.N. Chakravarty Canal named after him).",
        "Shortest-Serving Governor: Justice O.P. Verma (2 July 2004 - 7 July 2004; barely 5 days).",
        "First Female Speaker: Smt. Shanno Devi (first female speaker of any state legislature in India!).",
        "President's Rule in Haryana (3 Times):",
        "1. 21 Nov 1967 - 21 May 1968 (Rao Birender Singh regime fall)",
        "2. 30 Apr 1977 - 21 Jun 1977 (Post-Emergency dissolution)",
        "3. 6 Apr 1991 - 23 Jul 1991 (Om Prakash Chautala regime fall)"
    ]
    builder._create_card(s44, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Gubernatorial Profile & President's Rule", c1, builder.COLOR_COBALT)

    c2 = [
        "Chaudhary Devi Lal ('Jan Nayak' / 'Tau'):",
        "• Born in Teja Khera (Sirsa); organized historic 'Nyaya Yudh' in 1985-87.",
        "• Deputy Prime Minister of India (1989-1991) under V.P. Singh & Chandra Shekhar.",
        "• Introduced pioneering Old Age Pension (वृद्धावस्था पेंशन) of Rs 100/month in 1987 (now Rs 3,000/month).",
        "• Waived rural farmer and artisan loans up to Rs 10,000.",
        "Chaudhary Bansi Lal ('Iron Man of Haryana'):",
        "• Born in Golagarh (Bhiwani); modernized state infrastructure.",
        "• Spearheaded 100% rural electrification achieved on 29 Nov 1970 (felicitated by Indira Gandhi at Ucchani, Karnal).",
        "• Pioneered tourism naming after birds (Blue Jay, Skylark, Magpie, Parakeet)."
    ]
    builder._create_card(s44, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Devi Lal vs Bansi Lal: Modern Architects", c2, builder.COLOR_GOLD)

    # -------------------------------------------------------------------------
    # Slide 45: Punjab & Haryana High Court & Panchayati Raj
    # -------------------------------------------------------------------------
    s45 = prs.slides.add_slide(blank)
    builder._add_canvas_background(s45)
    builder._add_header_and_footer(
        s45,
        tag="JUDICIARY & LOCAL GOVERNANCE",
        title="Punjab & Haryana High Court & Panchayati Raj Governance",
        takeaway="High Court began in Lahore (1919); Haryana was 2nd state to mandate educational criteria for Panchayats.",
        citation="Punjab & Haryana High Court; Haryana Panchayati Raj Act, 1994; Supreme Court (Rajbala v. State of Haryana, 2015).",
        slide_num=45
    )
    c1 = [
        "Judicial Lineage: Established at Lahore on 21 March 1919 under Letters Patent issued by King George V.",
        "First Indian Chief Justice: Sir Shadi Lal of Rewari (appointed Chief Justice of Lahore High Court in 1920).",
        "Post-Partition Shift: Shifted to Shimla (Peterhoff) in 1947; relocated to Le Corbusier-designed building in Chandigarh in 1955.",
        "First Chief Justice post-1966: Justice Mehar Singh was Chief Justice when Haryana was carved out on 1 Nov 1966.",
        "Sanctioned Strength: 85 judges (64 permanent, 21 additional); shared jurisdiction over Punjab, Haryana, and UT Chandigarh."
    ]
    builder._create_card(s45, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Punjab & Haryana High Court", c1, builder.COLOR_COBALT)

    c2 = [
        "Haryana Panchayati Raj Act 1994: Passed on 22 April 1994 enforcing 73rd Constitutional Amendment (Three-tier: Gram Panchayat, Panchayat Samiti, Zila Parishad).",
        "Historic Amendment (2015): Mandated minimum educational qualification to contest Panchayat polls (10th pass for General males, 8th pass for SC males & women, 5th pass for SC women).",
        "Mandatory Sanitation & Defaulter Check: Must have a functional toilet at home; no power bill arrears; no cooperative loan default.",
        "Upheld by Supreme Court: Validated in landmark case *Rajbala v. State of Haryana (2015)*.",
        "Panchayat 2.0 Reforms: Inter-District Council, Gram Darshan portal, and 50% reservation for women in Panchayati Raj institutions."
    ]
    builder._create_card(s45, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Panchayati Raj & Educational Criteria Milestone", c2, builder.COLOR_TEAL)

    print("Polity slides (41-45) built successfully.")
