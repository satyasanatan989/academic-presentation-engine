"""
Haryana GK Master Screening Presentation Generator - Royal Oxford Keynote Edition
Comprehensive 65-Slide Curriculum for Haryana Food Safety Officer (FSO) & State Screening Tests
(HPSC HCS, HPSC FSO, HPSC ADO, HPSC Assistant Professor, HSSC CET)

Design Architecture:
- Canvas: Luminous Slate-White (#F8FAFC)
- Header: Deep Royal Oxford Navy (#0F1D36) with Champagne Gold (#D4A017) rule line
- Grounded Footer: Deep Royal Oxford Navy capsule with exam citations & slide counter
- Typography: 100% Times New Roman across every title, subtitle, card, table, and footer
- High-Contrast Elevated Cards: Native DrawingML vector containers with color-coded accent strips
- Tables: Explicit column budgeting with middle vertical centering and word-wrapping
"""

import sys
import os
import subprocess
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

class HaryanaGKPresentationBuilder:
    FONT_FAMILY = "Times New Roman"

    # Palette
    COLOR_CANVAS = RGBColor(248, 250, 252)         # Luminous Slate White #F8FAFC
    COLOR_HEADER_BG = RGBColor(15, 29, 54)         # Deep Royal Oxford Navy #0F1D36
    COLOR_HEADER_BAR = RGBColor(212, 160, 23)      # Metallic Champagne Gold #D4A017
    
    COLOR_CARD_WHITE = RGBColor(255, 255, 255)     # Elevated Pure White #FFFFFF
    COLOR_CARD_ALT = RGBColor(241, 245, 249)       # Subtle Slate Tint #F1F5F9
    COLOR_BORDER_SLATE = RGBColor(203, 213, 225)   # Subtle Slate Outline #CBD5E1
    
    COLOR_COBALT = RGBColor(29, 78, 216)           # Royal Cobalt Blue #1D4ED8
    COLOR_COBALT_LIGHT = RGBColor(239, 246, 255)   # Light Ice Blue #EFF6FF
    COLOR_GOLD = RGBColor(180, 83, 9)              # Deep Amber Gold #B45309
    COLOR_GOLD_LIGHT = RGBColor(254, 243, 199)     # Soft Amber Tint #FEF3C7
    COLOR_GOLD_BRIGHT = RGBColor(252, 211, 77)     # Keynote Champagne Gold #FCD34D
    COLOR_CRIMSON = RGBColor(185, 28, 28)          # Academic Crimson #B91C1C
    COLOR_CRIMSON_LIGHT = RGBColor(254, 242, 242)  # Light Rose Tint #FEF2F2
    COLOR_TEAL = RGBColor(13, 148, 136)            # Clinical Teal Accent #0D9488
    COLOR_TEAL_LIGHT = RGBColor(240, 253, 250)     # Light Teal Tint #F0FDFA
    COLOR_PURPLE = RGBColor(109, 40, 217)          # Royal Purple #6D28D9
    COLOR_PURPLE_LIGHT = RGBColor(245, 243, 255)   # Light Purple Tint #F5F3FF

    COLOR_TEXT_TITLE = RGBColor(255, 255, 255)     # Pure White Header Text
    COLOR_TEXT_SUBTITLE = RGBColor(186, 230, 253)  # Luminous Ice Blue
    COLOR_TEXT_DARK = RGBColor(15, 23, 42)         # Deep Charcoal Body Text
    COLOR_TEXT_NAVY = RGBColor(15, 30, 60)         # Deep Oxford Navy Headings
    COLOR_TEXT_MUTED = RGBColor(71, 85, 105)       # Slate Gray
    COLOR_TEXT_FOOTER = RGBColor(203, 213, 225)    # Silver Slate for Dark Footer
    COLOR_WHITE = RGBColor(255, 255, 255)

    TOTAL_SLIDES = 65

    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]

    def _add_canvas_background(self, slide):
        content_bg_path = PROJECT_ROOT / "outputs" / "haryana_assets" / "haryana_content_background.jpg"
        if content_bg_path.exists():
            bg = slide.shapes.add_picture(str(content_bg_path), 0, 0, self.prs.slide_width, self.prs.slide_height)
        else:
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
            bg.fill.solid()
            bg.fill.fore_color.rgb = self.COLOR_CANVAS
            bg.line.fill.background()

            top_zone = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, Inches(1.65))
            top_zone.fill.solid()
            top_zone.fill.fore_color.rgb = self.COLOR_HEADER_BG
            top_zone.line.fill.background()

            gold_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.65), self.prs.slide_width, Pt(2.5))
            gold_line.fill.solid()
            gold_line.fill.fore_color.rgb = self.COLOR_HEADER_BAR
            gold_line.line.fill.background()
        return bg

    def _add_header_and_footer(self, slide, tag: str, title: str, takeaway: str, citation: str, slide_num: int):
        # 1. Badge Pill
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.20), Inches(4.8), Inches(0.32))
        pill.fill.solid()
        pill.fill.fore_color.rgb = self.COLOR_HEADER_BG
        pill.line.color.rgb = self.COLOR_HEADER_BAR
        pill.line.width = Pt(1.0)
        ptf = pill.text_frame
        ptf.word_wrap = True
        ptf.margin_top = Pt(2)
        pp = ptf.paragraphs[0]
        pp.alignment = PP_ALIGN.CENTER
        pp.text = tag.upper()
        pp.font.name = self.FONT_FAMILY
        pp.font.size = Pt(9.5)
        pp.font.bold = True
        pp.font.color.rgb = self.COLOR_GOLD_BRIGHT

        # 2. Main Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(11.733), Inches(0.60))
        ttf = title_box.text_frame
        ttf.word_wrap = True
        ttf.margin_left = ttf.margin_right = ttf.margin_top = ttf.margin_bottom = 0
        tp = ttf.paragraphs[0]
        tp.text = title
        tp.font.name = self.FONT_FAMILY
        tp.font.size = Pt(21)
        tp.font.bold = True
        tp.font.color.rgb = self.COLOR_TEXT_TITLE

        # 3. Subtitle / Exam Takeaway
        sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.18), Inches(11.733), Inches(0.40))
        stf = sub_box.text_frame
        stf.word_wrap = True
        stf.margin_left = stf.margin_right = stf.margin_top = stf.margin_bottom = 0
        sp = stf.paragraphs[0]
        sp.text = f"Key Exam Takeaway: {takeaway}"
        sp.font.name = self.FONT_FAMILY
        sp.font.size = Pt(11)
        sp.font.italic = True
        sp.font.color.rgb = self.COLOR_TEXT_SUBTITLE

        # 4. Grounded Footer Capsule
        foot_capsule = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.80), Inches(11.733), Inches(0.46))
        foot_capsule.fill.solid()
        foot_capsule.fill.fore_color.rgb = self.COLOR_HEADER_BG
        foot_capsule.line.color.rgb = self.COLOR_BORDER_SLATE
        foot_capsule.line.width = Pt(1.0)

        # Footer Citation Text
        c_box = slide.shapes.add_textbox(Inches(0.95), Inches(6.82), Inches(10.2), Inches(0.40))
        ctf = c_box.text_frame
        ctf.word_wrap = True
        ctf.margin_top = Pt(4)
        cp = ctf.paragraphs[0]
        cp.text = citation
        cp.font.name = self.FONT_FAMILY
        cp.font.size = Pt(8.5)
        cp.font.color.rgb = self.COLOR_TEXT_FOOTER

        # Divider line in footer
        div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(11.35), Inches(6.88), Pt(1.5), Inches(0.30))
        div.fill.solid()
        div.fill.fore_color.rgb = self.COLOR_HEADER_BAR
        div.line.fill.background()

        # Counter Box
        num_box = slide.shapes.add_textbox(Inches(11.45), Inches(6.82), Inches(0.95), Inches(0.40))
        ntf = num_box.text_frame
        ntf.margin_top = Pt(4)
        np = ntf.paragraphs[0]
        np.alignment = PP_ALIGN.RIGHT
        np.text = f"{slide_num}/{self.TOTAL_SLIDES}"
        np.font.name = self.FONT_FAMILY
        np.font.size = Pt(9.5)
        np.font.bold = True
        np.font.color.rgb = self.COLOR_GOLD_BRIGHT

    def _create_card(self, slide, left, top, width, height, title: str, bullet_points: list, accent_color=None, bg_color=None):
        if accent_color is None:
            accent_color = self.COLOR_COBALT
        if bg_color is None:
            bg_color = self.COLOR_CARD_WHITE

        # Card container
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = self.COLOR_BORDER_SLATE
        card.line.width = Pt(1.0)

        # Top Accent Stripe
        stripe = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, Inches(0.36))
        stripe.fill.solid()
        stripe.fill.fore_color.rgb = accent_color
        stripe.line.fill.background()

        # Title text
        tb = slide.shapes.add_textbox(left + Inches(0.12), top + Inches(0.04), width - Inches(0.24), Inches(0.30))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_top = Pt(2)
        p = tf.paragraphs[0]
        p.text = title.upper()
        p.font.name = self.FONT_FAMILY
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = self.COLOR_WHITE
        p.alignment = PP_ALIGN.LEFT

        # Bullet body
        body_box = slide.shapes.add_textbox(left + Inches(0.20), top + Inches(0.45), width - Inches(0.40), height - Inches(0.52))
        btf = body_box.text_frame
        btf.word_wrap = True
        btf.margin_left = btf.margin_right = btf.margin_top = btf.margin_bottom = 0

        for i, bp in enumerate(bullet_points):
            p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
            p.font.name = self.FONT_FAMILY
            p.font.size = Pt(10)
            p.space_after = Pt(4)
            p.line_spacing = 1.15

            if ":" in bp:
                parts = bp.split(":", 1)
                r_head = p.add_run()
                r_head.text = f"• {parts[0].strip()}: "
                r_head.font.name = self.FONT_FAMILY
                r_head.font.bold = True
                r_head.font.color.rgb = self.COLOR_TEXT_NAVY

                r_body = p.add_run()
                r_body.text = parts[1].strip()
                r_body.font.name = self.FONT_FAMILY
                r_body.font.color.rgb = self.COLOR_TEXT_DARK
            else:
                r = p.add_run()
                r.text = f"• {bp.strip()}"
                r.font.name = self.FONT_FAMILY
                r.font.color.rgb = self.COLOR_TEXT_DARK

    def _create_table(self, slide, left, top, width, height, headers: list, rows: list, col_widths: list):
        num_rows = len(rows) + 1
        num_cols = len(headers)
        table_shape = slide.shapes.add_table(num_rows, num_cols, left, top, width, height)
        table = table_shape.table

        for i, w in enumerate(col_widths):
            table.columns[i].width = w

        # Header Row
        for col_idx, header_text in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = self.COLOR_HEADER_BG
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.text = str(header_text)
            p.font.name = self.FONT_FAMILY
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = self.COLOR_GOLD_BRIGHT

        # Data Rows
        for row_idx, row_data in enumerate(rows):
            bg_col = self.COLOR_CARD_WHITE if row_idx % 2 == 0 else self.COLOR_CARD_ALT
            for col_idx, cell_value in enumerate(row_data):
                cell = table.cell(row_idx + 1, col_idx)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg_col
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT if col_idx > 0 else PP_ALIGN.CENTER
                p.text = str(cell_value)
                p.font.name = self.FONT_FAMILY
                p.font.size = Pt(9.5)
                p.font.color.rgb = self.COLOR_TEXT_DARK
                if col_idx == 0:
                    p.font.bold = True
                    p.font.color.rgb = self.COLOR_TEXT_NAVY

    def _create_callout_box(self, slide, left, top, width, height, title: str, text: str, border_color=None, bg_color=None):
        if border_color is None:
            border_color = self.COLOR_GOLD
        if bg_color is None:
            bg_color = self.COLOR_GOLD_LIGHT

        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        box.fill.solid()
        box.fill.fore_color.rgb = bg_color
        box.line.color.rgb = border_color
        box.line.width = Pt(1.5)

        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0.12)
        
        p0 = tf.paragraphs[0]
        p0.text = title.upper()
        p0.font.name = self.FONT_FAMILY
        p0.font.size = Pt(10.5)
        p0.font.bold = True
        p0.font.color.rgb = self.COLOR_GOLD
        p0.space_after = Pt(2)

        p1 = tf.add_paragraph()
        p1.text = text
        p1.font.name = self.FONT_FAMILY
        p1.font.size = Pt(9.5)
        p1.font.color.rgb = self.COLOR_TEXT_DARK
        p1.line_spacing = 1.15

    # =========================================================================
    # SLIDE BUILDERS (SLIDES 1 - 65)
    # =========================================================================

    def build_slide_1(self):
        """Slide 1: Master Title Slide"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        
        # Royal Oxford Dark Background for Title Slide
        title_bg_path = PROJECT_ROOT / "outputs" / "haryana_assets" / "haryana_title_background.jpg"
        if title_bg_path.exists():
            bg = slide.shapes.add_picture(str(title_bg_path), 0, 0, self.prs.slide_width, self.prs.slide_height)
        else:
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
            bg.fill.solid()
            bg.fill.fore_color.rgb = self.COLOR_HEADER_BG
            bg.line.fill.background()

        # Gold decorative frame
        frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.6), self.prs.slide_width - Inches(1.2), self.prs.slide_height - Inches(1.2))
        frame.fill.background()
        frame.line.color.rgb = self.COLOR_HEADER_BAR
        frame.line.width = Pt(2.0)

        inner_frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(0.7), self.prs.slide_width - Inches(1.4), self.prs.slide_height - Inches(1.4))
        inner_frame.fill.background()
        inner_frame.line.color.rgb = self.COLOR_HEADER_BAR
        inner_frame.line.width = Pt(0.75)

        # Badge
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.8), Inches(1.1), Inches(5.733), Inches(0.45))
        badge.fill.solid()
        badge.fill.fore_color.rgb = self.COLOR_COBALT
        badge.line.color.rgb = self.COLOR_GOLD_BRIGHT
        badge.line.width = Pt(1.5)
        btf = badge.text_frame
        bp = btf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        bp.text = "HPSC & HSSC SCREENING EXAM COMPREHENSIVE DOSSIER"
        bp.font.name = self.FONT_FAMILY
        bp.font.size = Pt(11)
        bp.font.bold = True
        bp.font.color.rgb = self.COLOR_WHITE

        # Main Title
        tb = slide.shapes.add_textbox(Inches(1.0), Inches(1.75), Inches(11.333), Inches(1.6))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        p1.text = "HARYANA GENERAL KNOWLEDGE"
        p1.font.name = self.FONT_FAMILY
        p1.font.size = Pt(36)
        p1.font.bold = True
        p1.font.color.rgb = self.COLOR_GOLD_BRIGHT

        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.text = "Master Preparation Keynote for Food Safety Officer (FSO), HCS & Screening Tests"
        p2.font.name = self.FONT_FAMILY
        p2.font.size = Pt(18)
        p2.font.color.rgb = self.COLOR_TEXT_SUBTITLE
        p2.space_before = Pt(8)

        # 4 Core Pillar Containers
        pillars = [
            ("I. HISTORY & ARCHAEOLOGY", "Vedic Roots, Rakhigarhi, 3 Battles of Panipat, 1857 Heroes & Chhotu Ram", self.COLOR_COBALT),
            ("II. GEOGRAPHY & HYDROLOGY", "Shivalik to Aravalli, Yamuna-Ghaggar, Canals, Soils, Forests & Sanctuaries", self.COLOR_TEAL),
            ("III. AGRICULTURE & FOOD ECOSYSTEM", "FSO Special: Murrah Buffalo, NDRI, NIFTEM, Mega Food Parks & Crop Belts", self.COLOR_GOLD),
            ("IV. POLITY, ECONOMY & DISTRICTS", "90 Vidhan Sabha, Census 2011, 22 Districts Dossier & Top 30 Exam Traps", self.COLOR_PURPLE),
        ]
        col_w = Inches(2.7)
        gap = Inches(0.24)
        left_start = Inches(1.0)
        top_pos = Inches(3.6)
        h = Inches(2.3)

        for i, (p_title, p_desc, col) in enumerate(pillars):
            l = left_start + i * (col_w + gap)
            c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, top_pos, col_w, h)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor(22, 38, 70)
            c.line.color.rgb = col
            c.line.width = Pt(1.5)

            ctf = c.text_frame
            ctf.word_wrap = True
            ctf.margin_left = ctf.margin_right = ctf.margin_top = ctf.margin_bottom = Inches(0.12)
            
            cp1 = ctf.paragraphs[0]
            cp1.alignment = PP_ALIGN.CENTER
            cp1.text = p_title
            cp1.font.name = self.FONT_FAMILY
            cp1.font.size = Pt(10.5)
            cp1.font.bold = True
            cp1.font.color.rgb = self.COLOR_GOLD_BRIGHT
            cp1.space_after = Pt(8)

            cp2 = ctf.add_paragraph()
            cp2.alignment = PP_ALIGN.CENTER
            cp2.text = p_desc
            cp2.font.name = self.FONT_FAMILY
            cp2.font.size = Pt(9.5)
            cp2.font.color.rgb = self.COLOR_TEXT_FOOTER
            cp2.line_spacing = 1.2

        # Bottom Subtext
        bot_box = slide.shapes.add_textbox(Inches(1.0), Inches(6.15), Inches(11.333), Inches(0.45))
        btf = bot_box.text_frame
        bp = btf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        bp.text = "Based on Comprehensive Cross-Examination Analysis of 10+ Years HPSC / HSSC Screening Question Papers"
        bp.font.name = self.FONT_FAMILY
        bp.font.size = Pt(10.5)
        bp.font.italic = True
        bp.font.color.rgb = self.COLOR_GOLD_BRIGHT

    def build_slide_2(self):
        """Slide 2: Exam Blueprint & Screening Weightage Analysis"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide)
        self._add_header_and_footer(
            slide,
            tag="SCREENING EXAM BLUEPRINT",
            title="Weightage & Question Distribution in Haryana Screening Tests",
            takeaway="Haryana GK carries 20-25% decisive weightage in HPSC FSO and ADO screening tests.",
            citation="Analysis of HPSC FSO 2023, HPSC ADO 2022-23, HPSC HCS Pre 2021-2024, and HSSC CET Group C/D Exam Patterns.",
            slide_num=2
        )

        headers = ["Exam Category", "Conducting Body", "Total Questions", "Haryana GK Weightage", "Core High-Frequency Focus Areas"]
        rows = [
            ["Food Safety Officer (FSO)", "HPSC", "100 MCQs", "20-25 Questions", "Agriculture, Murrah Buffalo, NDRI/NIFTEM, Food Laws, Soil & Geography"],
            ["HCS (Executive Branch)", "HPSC", "100 MCQs (Paper 1)", "15-20 Questions", "Ancient History, Harappan Sites, 1857 Revolt, Vidhan Sabha Polity, Schemes"],
            ["Agriculture Development Officer", "HPSC", "100 MCQs", "20 Questions", "Soils, Canal Irrigation (WYC/Bhakra), Crop Seasons, Agro-Climatic Zones"],
            ["Assistant Professor (College)", "HPSC", "100 MCQs", "15-20 Questions", "History, Folk Literature (Saang), Census 2011, Demographics, Personalities"],
            ["HSSC CET (Group C / D)", "HSSC", "100 MCQs", "25 Questions (Fixed 25%)", "District Profiles, Folk Dances, Rivers, Sports Stars, Wildlife Sanctuaries"]
        ]
        widths = [Inches(2.5), Inches(1.5), Inches(1.5), Inches(2.0), Inches(4.233)]
        self._create_table(slide, Inches(0.8), Inches(1.9), Inches(11.733), Inches(2.7), headers, rows, widths)

        # Callout summary
        self._create_callout_box(
            slide, Inches(0.8), Inches(4.8), Inches(11.733), Inches(1.75),
            title="FSO Screening Strategic Mandate",
            text="In the Food Safety Officer (FSO) screening test, subject knowledge (Food Tech, Microbiology, Chemistry, FSS Act 2006) accounts for ~80% of technical marks, while Haryana GK forms the mandatory non-technical section (20%). Candidate qualification cutoff hinges on this section because technical scores are closely grouped. High-yield topics include: Dairy Research Infrastructure (NDRI, CIRB, LUVAS), Basmati Rice & Mustard Zones, Mega Food Parks, Soil Salinity Reclamation (CSSRI), and Government Agro-Nutrition Initiatives."
        )

    def build_slide_3(self):
        """Slide 3: Historical Genesis & Demand for Haryana"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide)
        self._add_header_and_footer(
            slide,
            tag="HISTORICAL GENESIS",
            title="The Political Journey to Statehood: 1858 to 1966",
            takeaway="Haryana was tagged to Punjab in 1858 as British collective punishment for active participation in the 1857 revolt.",
            citation="Government of Haryana Gazetteer; Report of the Punjab Boundary Commission (Shah Commission, 1966).",
            slide_num=3
        )

        c1 = [
            "Colonial Annexation (1858): After the 1857 First War of Independence, the British detached Haryana from North-Western Provinces and merged it with Punjab to penalize the rebellious peasantry.",
            "First Formal Demand (1926): All India Muslim League Delhi session, Pirzada Muhammad Hussain demanded separation of Delhi and Haryana territory from Punjab.",
            "Deshbandhu Gupta's Advocacy (1932): Formulated a clear plea for 'Greater Delhi' integrating Ambala Division, Meerut, and Agra into a culturally unified linguistic zone.",
            "Dinbandhu Sir Chhotu Ram: Stood for peasant rights through the Unionist Party (1923), safeguarding rural agrarian interests under the Punjab Legislative Council."
        ]
        self._create_card(slide, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Phase I: Pre-Independence Demands", c1, self.COLOR_COBALT)

        c2 = [
            "Sachar Formula (1 Oct 1949): Proposed by CM Bhim Sen Sachar. Divided Punjab into 'Punjabi Zone' (Gurmukhi script) and 'Hindi Zone' (Devanagari script). Rejected due to language friction.",
            "Fazal Ali Commission (SRC 1953-55): States Reorganisation Commission rejected separate Haryana, citing administrative synergy, causing massive discontent among Hindi speakers.",
            "Regional Formula (1956): Established Hindi and Punjabi Regional Committees in Punjab Assembly. Failed to resolve deep developmental and cultural disparities.",
            "Sant Fateh Singh's Agitation (1965): Threat of self-immolation for Punjabi Suba forced Central Government to form Parliamentary Committee headed by Lok Sabha Speaker Hukam Singh."
        ]
        self._create_card(slide, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Phase II: Post-Independence Formulas", c2, self.COLOR_GOLD)

    def build_slide_4(self):
        """Slide 4: Shah Commission & 18th Constitutional Amendment"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide)
        self._add_header_and_footer(
            slide,
            tag="FORMATION OF THE STATE",
            title="Shah Commission & 18th Amendment: Birth of 17th State",
            takeaway="Haryana was created under the Punjab Reorganisation Act 1966 via the 18th Constitutional Amendment.",
            citation="Punjab Reorganisation Act, 1966 (Act No. 31 of 1966); Ministry of Law & Justice, Government of India.",
            slide_num=4
        )

        c1 = [
            "Constitution Date: 23 April 1966 by the Government of India under the chairmanship of Justice J.C. Shah (Supreme Court Judge).",
            "Commission Members: Justice J.C. Shah (Chairman), S. Dutt (Member), and M.M. Philip (Member).",
            "Mandate: Demarcate the Hindi-speaking areas of the composite State of Punjab to carve out a new State.",
            "Report Submission: 31 May 1966. Recommended creation of Haryana with Ambala, Gurgaon, Hissar, Rohtak, Mahendragarh, Jind, and Narwana tehsils.",
            "Kharar & Chandigarh Controversy: Commission recommended Kharar tehsil (including Chandigarh) to Haryana by 2:1 majority (S. Dutt dissented). Center later made Chandigarh a Joint Capital Union Territory."
        ]
        self._create_card(slide, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.7), "Shah Commission Key Parameters", c1, self.COLOR_COBALT)

        c2 = [
            "18th Constitutional Amendment Act 1966: Clarified parliamentary power under Article 3 to form a new State or Union Territory by uniting parts of existing states.",
            "Punjab Reorganisation Act 1966: Passed by Parliament in September 1966; received Presidential assent on 18 September 1966.",
            "Official Formation: 1 November 1966 as the 17th State of the Republic of India.",
            "Area Ceded: Approximately 35.18% of the composite Punjab area (44,212 sq km) was transferred to Haryana.",
            "Shared Institutions: Common High Court (Punjab & Haryana High Court at Chandigarh), Joint Capital (Chandigarh), and shared Bhakra Beas Management Board (BBMB)."
        ]
        self._create_card(slide, Inches(6.8), Inches(1.9), Inches(5.733), Inches(4.7), "Constitutional & Legislative Framework", c2, self.COLOR_TEAL)

    def build_slide_5(self):
        """Slide 5: Administrative Evolution: Original 7 to 22 Districts"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide)
        self._add_header_and_footer(
            slide,
            tag="ADMINISTRATIVE EVOLUTION",
            title="District Genesis: From Original 7 ('GHAR JK') to 22 Districts",
            takeaway="Haryana began on Nov 1, 1966 with 7 districts; Charkhi Dadri was carved out as the 22nd district in Dec 2016.",
            citation="Department of Revenue & Disaster Management, Haryana; State Statistical Abstract 2023-24.",
            slide_num=5
        )

        headers = ["Chronological Phase", "Districts Created", "Carved Out From", "Formation Date", "Exam Significance"]
        rows = [
            ["Original 7 (At Birth)", "Gurgaon, Hissar, Ambala, Rohtak, Jind, Karnal, Mahendragarh", "Punjab Reorganisation Act", "1 Nov 1966", "Acronym 'GHAR JK'. Largest at birth: Hissar; Smallest: Jind."],
            ["Phase 2 (1972-73)", "Bhiwani & Sonipat (1972)\nKurukshetra (1973)", "Hissar/Rohtak\nKarnal", "22 Dec 1972\n23 Jan 1973", "Bhiwani created by Bansi Lal; Kurukshetra named after King Kuru."],
            ["Phase 3 (1975-79)", "Sirsa (1975)\nFaridabad (1979)", "Hissar\nGurgaon", "1 Sep 1975\n2 Aug 1979", "Sirsa became largest district after Bhiwani split. Faridabad industrial hub."],
            ["Phase 4 (1989 Quad)", "Kaithal, Panipat, Rewari, Yamunanagar", "Karnal, Rohtak, Gurgaon, Ambala", "1 Nov 1989", "Created simultaneously on Haryana Day. Panipat merged back & re-created 1992!"],
            ["Phase 5 (1995-97)", "Panchkula (1995)\nFatehabad & Jhajjar (1997)", "Ambala\nHissar & Rohtak", "15 Aug 1995\n15 Jul 1997", "Panchkula planned satellite; Fatehabad historical city of Firoz Shah."],
            ["Phase 6 (2005-16)", "Nuh/Mewat (2005), Palwal (2008),\nCharkhi Dadri (2016)", "Gurgaon/Faridabad\nBhiwani", "4 Apr 2005\n15 Aug 2008\n1 Dec 2016", "Nuh created as Satyamevpuram. Charkhi Dadri is the 22nd & newest district."]
        ]
        widths = [Inches(1.8), Inches(2.6), Inches(1.8), Inches(1.5), Inches(4.033)]
        self._create_table(slide, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.7), headers, rows, widths)

    def build_slide_6(self):
        """Slide 6: State Symbols, Emblems & Official Identity"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide)
        self._add_header_and_footer(
            slide,
            tag="STATE IDENTITY",
            title="Official State Symbols & Emblems of Haryana",
            takeaway="Black Francolin (Kala Teetar) is State Bird; Peepal is State Tree; Lotus is State Flower.",
            citation="Haryana Government Notification No. 12/4/85-Edu; State Wildlife Wing.",
            slide_num=6
        )

        c1 = [
            "Scientific Name: Francolinus francolinus.",
            "Local Name: Kala Teetar (काला तीतर).",
            "Habitat: Shrublands, cultivated fields of Sirsa, Hisar, Bhiwani, and Kurukshetra.",
            "Ecological Role: Natural pest predator in wheat and mustard fields.",
            "HPSC Trap: Do not confuse with Great Indian Bustard (Rajasthan) or Monal (Uttarakhand)."
        ]
        self._create_card(slide, Inches(0.8), Inches(1.9), Inches(2.75), Inches(4.7), "State Bird: Black Francolin", c1, self.COLOR_COBALT)

        c2 = [
            "Scientific Name: Antilope cervicapra.",
            "Local Name: Krishna Mriga / Kala Hiran (काला हिरण).",
            "Primary Range: Abubshahar Wildlife Sanctuary (Sirsa), Bishnoi settlements in Fatehabad and Hisar.",
            "Cultural Protector: Bishnoi community upholds absolute sacred protection.",
            "IUCN Status: Least Concern (Schedule I of Wildlife Protection Act 1972)."
        ]
        self._create_card(slide, Inches(3.75), Inches(1.9), Inches(2.75), Inches(4.7), "State Animal: Blackbuck", c2, self.COLOR_TEAL)

        c3 = [
            "Scientific Name: Ficus religiosa.",
            "Local Name: Peepal (पीपल) / Bodhi Vriksha.",
            "Sacred Significance: Mentioned in Rigveda and Bhagavad Gita (Lord Krishna calls himself Ashvattha among trees).",
            "Medicinal & Oxygen Value: Continuous oxygen producer; rich in flavonoids and tannins used in traditional medicine.",
            "HPSC Trap: State Tree is Peepal, NOT Khejri (which is Rajasthan's State Tree)."
        ]
        self._create_card(slide, Inches(6.70), Inches(1.9), Inches(2.75), Inches(4.7), "State Tree: Peepal", c3, self.COLOR_GOLD)

        c4 = [
            "State Flower: Lotus (Nelumbo nucifera) / Kamal. Sacred aquatic herb prevalent in village ponds and lakes.",
            "Official State Seal: Circular seal featuring Lion Capital of Ashoka with a rising sun emerging behind lotus blooms, flanked by ears of wheat.",
            "Official Language: Hindi (Devanagari script). Punjabi was declared 2nd official language in 2010 (under Bhupinder Singh Hooda; earlier Bansi Lal made Tamil 2nd language in 1969 to counter Punjabi).",
            "State Sport: Wrestling (Kushti) - Traditional soil akhadas producing 60%+ of national Olympic wrestling medals."
        ]
        self._create_card(slide, Inches(9.65), Inches(1.9), Inches(2.88), Inches(4.7), "Flower, Seal & Language", c4, self.COLOR_PURPLE)

    def generate_all_slides(self):
        print("Generating Module 1: Orientation & Formation (Slides 1-6)...")
        self.build_slide_1()
        self.build_slide_2()
        self.build_slide_3()
        self.build_slide_4()
        self.build_slide_5()
        self.build_slide_6()

        print("Generating Module 2: Geography, Soils & Hydrology (Slides 7-18)...")
        from slides_modules.module_geography import build_geography_slides
        build_geography_slides(self)

        print("Generating Module 3: Climate, Forests & Sanctuaries (Slides 19-24)...")
        from slides_modules.module_environment import build_environment_slides
        build_environment_slides(self)

        print("Generating Module 4: History & Freedom Struggle (Slides 25-40)...")
        from slides_modules.module_history import build_history_slides
        build_history_slides(self)

        print("Generating Module 5: Polity, Judiciary & Governance (Slides 41-45)...")
        from slides_modules.module_polity import build_polity_slides
        build_polity_slides(self)

        print("Generating Module 6: Agriculture, Dairying & Food Tech (Slides 46-51)...")
        from slides_modules.module_agriculture import build_agriculture_slides
        build_agriculture_slides(self)

        print("Generating Module 7: Economy, Minerals & Census 2011 (Slides 52-56)...")
        from slides_modules.module_economy import build_economy_slides
        build_economy_slides(self)

        print("Generating Module 8: Folk Culture, Dances & Sports (Slides 57-60)...")
        from slides_modules.module_culture import build_culture_slides
        build_culture_slides(self)

        print("Generating Module 9: District Dossier, Schemes & Top Traps (Slides 61-65)...")
        from slides_modules.module_districts_traps import build_districts_and_traps_slides
        build_districts_and_traps_slides(self)

        output_dir = PROJECT_ROOT.parent / "samples"
        output_dir.mkdir(parents=True, exist_ok=True)
        pptx_path = output_dir / "haryana_gk_master_screening_deck.pptx"
        self.prs.save(str(pptx_path))
        print(f"\n[SUCCESS] Master Deck of {len(self.prs.slides)} slides saved to: {pptx_path}")
        return pptx_path


def main():
    print("=================================================================")
    print("HARYANA GK MASTER SCREENING PRESENTATION GENERATOR (ROYAL OXFORD)")
    print("65-Slide High-Yield Keynote for HPSC FSO & State Screening Tests")
    print("=================================================================\n")
    
    builder = HaryanaGKPresentationBuilder()
    pptx_path = builder.generate_all_slides()
    print("Slide deck compilation complete!")


if __name__ == "__main__":
    main()

