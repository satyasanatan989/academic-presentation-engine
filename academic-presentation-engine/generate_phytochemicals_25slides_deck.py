"""
Biological Functionality of Phytochemicals - 25-Slide Master Keynote Edition
Comprehensive 25-Slide Academic & Medical Presentation on:
'Biological Functionality of Cruciferous Vegetables, Tropical, Subtropical and Temperate Fruits, Herbs and Spices'

Design Architecture (Botanical Clinical Luxury - Modern Scientific Keynote):
- Primary Canvas: Luminous Slate-White (#F8FAFC)
- Header Canopy: Deep Botanical Navy-Emerald (#081C1B) with dual metallic rules (Champagne Gold #D4A017 and Chlorophyll Mint #10B981)
- Typography: 100% Times New Roman across all slides, titles, cards, tables, and footers
- Elevated Card System: Pure White (#FFFFFF) rounded containers with color-coded thematic rails
- AI Flowcharts & Infographics: Embedded high-yield vector flowcharts, enzyme cascades, and metabolic diagrams
- Museum-Grade Botanical Plates: High-resolution botanical photography in royal mats with gold hairline borders
- Clinical & Mechanistic Depth: Full biochemical pathways, receptor dynamics, human dosages, and peer-reviewed citations
"""

import sys
import os
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

PROJECT_ROOT = Path(__file__).resolve().parent
REPO_ROOT = PROJECT_ROOT.parent
sys.path.insert(0, str(PROJECT_ROOT))

ASSETS_DIR = PROJECT_ROOT / "outputs" / "PRJ-PHYTOCHEMICALS-001" / "assets"

class Phytochemicals25DeckBuilder:
    FONT_FAMILY = "Times New Roman"

    # Botanical Clinical Luxury Palette
    COLOR_CANVAS = RGBColor(248, 250, 252)          # Slate White #F8FAFC
    COLOR_HEADER_BG = RGBColor(8, 28, 27)           # Deep Botanical Emerald Navy #081C1B
    COLOR_HEADER_GOLD = RGBColor(212, 160, 23)      # Metallic Champagne Gold #D4A017
    COLOR_HEADER_MINT = RGBColor(16, 185, 129)      # Chlorophyll Mint #10B981

    COLOR_CARD_WHITE = RGBColor(255, 255, 255)      # Pure Elevated White #FFFFFF
    COLOR_CARD_ALT = RGBColor(241, 245, 249)        # Subtle Slate Tint #F1F5F9
    COLOR_BORDER_SLATE = RGBColor(203, 213, 225)    # Subtle Slate Outline #CBD5E1

    # Pathway & Categorical Thematic Colors
    COLOR_EMERALD = RGBColor(5, 150, 105)           # Chlorophyll Emerald #059669
    COLOR_COBALT = RGBColor(29, 78, 216)            # Royal Cobalt Blue #1D4ED8
    COLOR_AMBER = RGBColor(217, 119, 6)             # Bioactive Amber #D97706
    COLOR_CRIMSON = RGBColor(185, 28, 28)           # Academic Crimson #B91C1C
    COLOR_PURPLE = RGBColor(124, 58, 237)           # Royal Purple #7C3AED
    COLOR_TEAL = RGBColor(13, 148, 136)             # Clinical Teal #0D9488
    COLOR_GOLD_BRIGHT = RGBColor(252, 211, 77)      # Keynote Bright Gold #FCD34D

    COLOR_TEXT_TITLE = RGBColor(255, 255, 255)      # Header Pure White
    COLOR_TEXT_SUBTITLE = RGBColor(204, 251, 241)   # Soft Mint Ice
    COLOR_TEXT_DARK = RGBColor(15, 23, 42)          # Deep Charcoal Body Text (#0F172A)
    COLOR_TEXT_NAVY = RGBColor(8, 28, 27)           # Deep Botanical Navy Headings
    COLOR_TEXT_MUTED = RGBColor(71, 85, 105)        # Slate Muted Text
    COLOR_TEXT_FOOTER = RGBColor(203, 213, 225)     # Silver Slate

    TOTAL_SLIDES = 25

    SLIDE_METADATA = {
        1:  {"domain": "EXECUTIVE KEYNOTE", "theme": COLOR_HEADER_GOLD},
        2:  {"domain": "PHYTOCHEMICAL TAXONOMY", "theme": COLOR_EMERALD},
        3:  {"domain": "BRASSICACEAE CASCADE", "theme": COLOR_EMERALD},
        4:  {"domain": "KEAP1-NRF2 MECHANISM", "theme": COLOR_COBALT},
        5:  {"domain": "EPIGENETICS & HDAC", "theme": COLOR_PURPLE},
        6:  {"domain": "I3C & DIM ESTROGEN RATIO", "theme": COLOR_CRIMSON},
        7:  {"domain": "TROPICAL FRUITS & PROTEASES", "theme": COLOR_AMBER},
        8:  {"domain": "BROMELAIN & PAPAIN FLOW", "theme": COLOR_AMBER},
        9:  {"domain": "MANGIFERIN & GLYCEMIA", "theme": COLOR_TEAL},
        10: {"domain": "SUBTROPICAL CITRUS PMFs", "theme": COLOR_AMBER},
        11: {"domain": "eNOS & ENDOTHELIAL NO", "theme": COLOR_COBALT},
        12: {"domain": "PUNICALAGIN & UROLITHIN A", "theme": COLOR_CRIMSON},
        13: {"domain": "TEMPERATE ROSACEAE BIOACTIVES", "theme": COLOR_EMERALD},
        14: {"domain": "QUERCETIN & MAST CELLS", "theme": COLOR_COBALT},
        15: {"domain": "BERRY ANTHOCYANINS & BBB", "theme": COLOR_PURPLE},
        16: {"domain": "ZINGIBERACEAE PHARMACOLOGY", "theme": COLOR_AMBER},
        17: {"domain": "CURCUMIN-PIPERINE SURGE", "theme": COLOR_CRIMSON},
        18: {"domain": "GINGEROLS & 5-HT3 MOTILITY", "theme": COLOR_TEAL},
        19: {"domain": "ALLIUM ALLICIN CASCADE", "theme": COLOR_EMERALD},
        20: {"domain": "LAMIACEAE ANTIOXIDANTS", "theme": COLOR_EMERALD},
        21: {"domain": "FOOD MATRIX SYNERGIES", "theme": COLOR_COBALT},
        22: {"domain": "HUMAN CLINICAL DOSAGE", "theme": COLOR_TEAL},
        23: {"domain": "CULINARY PHARMACOKINETICS", "theme": COLOR_AMBER},
        24: {"domain": "MULTI-TARGET SUMMARY", "theme": COLOR_PURPLE},
        25: {"domain": "SCIENTIFIC BIBLIOGRAPHY", "theme": COLOR_HEADER_GOLD}
    }

    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]

    def _add_canvas_background(self, slide, slide_num: int):
        meta = self.SLIDE_METADATA.get(slide_num, self.SLIDE_METADATA[1])

        # 1. Base Luminous Slate-White Canvas
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = self.COLOR_CANVAS
        bg.line.fill.background()

        # 2. Left Spine Architectural Accent Rail
        spine = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.14), self.prs.slide_height)
        spine.fill.solid()
        spine.fill.fore_color.rgb = meta["theme"]
        spine.line.fill.background()

        # 3. Solid Deep Botanical Navy Header Zone
        top_zone = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, Inches(1.72))
        top_zone.fill.solid()
        top_zone.fill.fore_color.rgb = self.COLOR_HEADER_BG
        top_zone.line.fill.background()

        # 4. Primary Metallic Champagne Gold Divider Rule
        gold_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.72), self.prs.slide_width, Pt(2.5))
        gold_line.fill.solid()
        gold_line.fill.fore_color.rgb = self.COLOR_HEADER_GOLD
        gold_line.line.fill.background()

        # 5. Secondary Mint Theme Accent Hairline
        mod_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.72) + Pt(2.5), self.prs.slide_width, Pt(1.5))
        mod_line.fill.solid()
        mod_line.fill.fore_color.rgb = meta["theme"]
        mod_line.line.fill.background()

        return bg

    def _add_header_and_footer(self, slide, tag: str, title: str, takeaway: str, citation: str, slide_num: int):
        meta = self.SLIDE_METADATA.get(slide_num, self.SLIDE_METADATA[1])

        # Category Badge Pill (Left)
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.20), Inches(4.5), Inches(0.32))
        pill.fill.solid()
        pill.fill.fore_color.rgb = self.COLOR_HEADER_BG
        pill.line.color.rgb = self.COLOR_HEADER_GOLD
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

        # Right Theme Tracker Pill
        mod_pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(0.20), Inches(4.733), Inches(0.32))
        mod_pill.fill.solid()
        mod_pill.fill.fore_color.rgb = RGBColor(12, 42, 40)
        mod_pill.line.color.rgb = meta["theme"]
        mod_pill.line.width = Pt(1.0)
        mtf = mod_pill.text_frame
        mtf.word_wrap = True
        mtf.margin_top = Pt(2)
        mp = mtf.paragraphs[0]
        mp.alignment = PP_ALIGN.CENTER
        mp.text = f"MODULE {slide_num:02d}/{self.TOTAL_SLIDES} • {meta['domain'][:30]}"
        mp.font.name = self.FONT_FAMILY
        mp.font.size = Pt(9.0)
        mp.font.bold = True
        mp.font.color.rgb = self.COLOR_GOLD_BRIGHT

        # Main Title Box
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(11.733), Inches(0.60))
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = self.FONT_FAMILY
        p.font.size = Pt(20.5)
        p.font.bold = True
        p.font.color.rgb = self.COLOR_TEXT_TITLE

        # Subtitle / Key Takeaway Box
        sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.15), Inches(11.733), Inches(0.48))
        stf = sub_box.text_frame
        stf.word_wrap = True
        stf.margin_left = stf.margin_top = stf.margin_right = stf.margin_bottom = 0
        sp = stf.paragraphs[0]
        sp.text = takeaway
        sp.font.name = self.FONT_FAMILY
        sp.font.size = Pt(11.5)
        sp.font.color.rgb = self.COLOR_TEXT_SUBTITLE

        # Grounded Footer Zone (Deep Navy Capsule)
        foot_zone = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.92), Inches(11.733), Inches(0.40))
        foot_zone.fill.solid()
        foot_zone.fill.fore_color.rgb = self.COLOR_HEADER_BG
        foot_zone.line.color.rgb = meta["theme"]
        foot_zone.line.width = Pt(0.75)

        # Footer Left Citation Text
        c_box = slide.shapes.add_textbox(Inches(0.95), Inches(6.93), Inches(9.8), Inches(0.38))
        ctf = c_box.text_frame
        ctf.word_wrap = True
        ctf.margin_left = ctf.margin_top = ctf.margin_right = ctf.margin_bottom = 0
        cp = ctf.paragraphs[0]
        cp.text = citation
        cp.font.name = self.FONT_FAMILY
        cp.font.size = Pt(8.5)
        cp.font.color.rgb = self.COLOR_TEXT_FOOTER

        # Footer Right Slide Counter
        pg_box = slide.shapes.add_textbox(Inches(10.8), Inches(6.93), Inches(1.6), Inches(0.38))
        ptf2 = pg_box.text_frame
        ptf2.word_wrap = True
        ptf2.margin_left = ptf2.margin_top = ptf2.margin_right = ptf2.margin_bottom = 0
        pp2 = ptf2.paragraphs[0]
        pp2.alignment = PP_ALIGN.RIGHT
        pp2.text = f"SLIDE {slide_num:02d} / {self.TOTAL_SLIDES:02d}"
        pp2.font.name = self.FONT_FAMILY
        pp2.font.size = Pt(8.5)
        pp2.font.bold = True
        pp2.font.color.rgb = self.COLOR_GOLD_BRIGHT

    def create_card(self, slide, left, top, width, height, title: str, subtitle: str = None, accent_color=None):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = self.COLOR_CARD_WHITE
        card.line.color.rgb = self.COLOR_BORDER_SLATE
        card.line.width = Pt(1.0)

        color = accent_color or self.COLOR_EMERALD

        # Top Accent Rail
        rail = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left + Inches(0.12), top, width - Inches(0.24), Pt(3.0))
        rail.fill.solid()
        rail.fill.fore_color.rgb = color
        rail.line.fill.background()

        # Title
        t_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.12), width - Inches(0.4), Inches(0.35))
        tf = t_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = self.FONT_FAMILY
        p.font.size = Pt(12.5)
        p.font.bold = True
        p.font.color.rgb = self.COLOR_TEXT_NAVY

        start_y = top + Inches(0.48)
        if subtitle:
            s_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.44), width - Inches(0.4), Inches(0.25))
            stf = s_box.text_frame
            stf.word_wrap = True
            stf.margin_left = stf.margin_top = stf.margin_right = stf.margin_bottom = 0
            sp = stf.paragraphs[0]
            sp.text = subtitle
            sp.font.name = self.FONT_FAMILY
            sp.font.size = Pt(9.5)
            sp.font.bold = True
            sp.font.color.rgb = color
            start_y = top + Inches(0.72)

        content_height = (top + height) - start_y - Inches(0.15)
        c_box = slide.shapes.add_textbox(left + Inches(0.2), start_y, width - Inches(0.4), content_height)
        ctf = c_box.text_frame
        ctf.word_wrap = True
        ctf.margin_left = ctf.margin_top = ctf.margin_right = ctf.margin_bottom = 0

        return ctf

    def append_bullet(self, tf, bold_prefix: str, text: str, font_size: float = 9.2, space_after: float = 4.0):
        if len(tf.paragraphs) == 1 and tf.paragraphs[0].text == "":
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.space_after = Pt(space_after)

        r1 = p.add_run()
        r1.text = bold_prefix + " "
        r1.font.name = self.FONT_FAMILY
        r1.font.size = Pt(font_size)
        r1.font.bold = True
        r1.font.color.rgb = self.COLOR_TEXT_NAVY

        r2 = p.add_run()
        r2.text = text
        r2.font.name = self.FONT_FAMILY
        r2.font.size = Pt(font_size)
        r2.font.bold = False
        r2.font.color.rgb = self.COLOR_TEXT_DARK

    def add_image_panel(self, slide, left, top, width, height, image_name: str, caption: str = None):
        img_path = ASSETS_DIR / image_name
        if not img_path.exists():
            return None

        # Surrounding Frame Container
        frame = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        frame.fill.solid()
        frame.fill.fore_color.rgb = self.COLOR_CARD_WHITE
        frame.line.color.rgb = self.COLOR_BORDER_SLATE
        frame.line.width = Pt(1.0)

        # Image inset
        pad = Inches(0.08)
        img_w = width - (pad * 2)
        img_h = height - (pad * 2) - (Inches(0.3) if caption else 0)
        slide.shapes.add_picture(str(img_path), left + pad, top + pad, img_w, img_h)

        if caption:
            c_box = slide.shapes.add_textbox(left + pad, top + height - Inches(0.28), img_w, Inches(0.25))
            ctf = c_box.text_frame
            ctf.word_wrap = True
            p = ctf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.text = caption
            p.font.name = self.FONT_FAMILY
            p.font.size = Pt(8.5)
            p.font.color.rgb = self.COLOR_TEXT_MUTED

    def create_table(self, slide, left, top, width, height, headers: list, col_widths: list, rows_data: list):
        rows_count = len(rows_data) + 1
        cols_count = len(headers)
        table_shape = slide.shapes.add_table(rows_count, cols_count, left, top, width, height)
        table = table_shape.table

        for i, w in enumerate(col_widths):
            table.columns[i].width = w

        for col_idx, header_text in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = self.COLOR_HEADER_BG
            tf = cell.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = Inches(0.08)
            p = tf.paragraphs[0]
            p.text = header_text
            p.alignment = PP_ALIGN.CENTER
            p.font.name = self.FONT_FAMILY
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = self.COLOR_GOLD_BRIGHT

        for row_idx, row_values in enumerate(rows_data):
            bg_color = self.COLOR_CARD_WHITE if (row_idx % 2 == 0) else self.COLOR_CARD_ALT
            for col_idx, val in enumerate(row_values):
                cell = table.cell(row_idx + 1, col_idx)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg_color
                tf = cell.text_frame
                tf.word_wrap = True
                tf.margin_left = tf.margin_right = Inches(0.08)
                p = tf.paragraphs[0]
                p.text = str(val)
                p.font.name = self.FONT_FAMILY
                p.font.size = Pt(9.0)
                if col_idx == 0:
                    p.alignment = PP_ALIGN.LEFT
                    p.font.bold = True
                    p.font.color.rgb = self.COLOR_TEXT_NAVY
                else:
                    p.alignment = PP_ALIGN.LEFT
                    p.font.bold = False
                    p.font.color.rgb = self.COLOR_TEXT_DARK

        return table_shape

    # -------------------------------------------------------------------------
    # 25 INDIVIDUAL SLIDE BUILDERS
    # -------------------------------------------------------------------------
    def build_slide_01_cover(self):
        slide = self.prs.slides.add_slide(self.blank_layout)

        # Full dark Oxford Navy-Emerald keynote canvas
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = self.COLOR_HEADER_BG
        bg.line.fill.background()

        # Double Gold & Mint Accent Rules
        rule1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.1), Inches(11.733), Pt(2.5))
        rule1.fill.solid()
        rule1.fill.fore_color.rgb = self.COLOR_HEADER_GOLD
        rule1.line.fill.background()

        # Category Badge
        cat = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.65), Inches(5.8), Inches(0.35))
        cat.fill.solid()
        cat.fill.fore_color.rgb = RGBColor(14, 46, 44)
        cat.line.color.rgb = self.COLOR_HEADER_GOLD
        cat.line.width = Pt(1.0)
        ctf = cat.text_frame
        p = ctf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = "ADVANCED NUTRITIONAL BIOCHEMISTRY & PHARMACOLOGY"
        p.font.name = self.FONT_FAMILY
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = self.COLOR_GOLD_BRIGHT

        # Main Title Box
        tbox = slide.shapes.add_textbox(Inches(0.8), Inches(1.35), Inches(11.733), Inches(1.6))
        tf = tbox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = "BIOLOGICAL FUNCTIONALITY OF CRUCIFEROUS VEGETABLES, FRUITS, HERBS & SPICES"
        p.font.name = self.FONT_FAMILY
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = self.COLOR_TEXT_TITLE

        p2 = tf.add_paragraph()
        p2.text = "A 25-Slide Comprehensive Master Curriculum: Enzymatic Cascades, Cellular Receptors, Epigenetic Signatures, Bioavailability Surges & Food Matrix Synergies"
        p2.font.name = self.FONT_FAMILY
        p2.font.size = Pt(12.5)
        p2.font.color.rgb = self.COLOR_TEXT_SUBTITLE

        # 4 Core Pillar Cards
        card_w = Inches(2.78)
        card_h = Inches(3.3)
        gap = Inches(0.2)
        start_x = Inches(0.8)
        top_y = Inches(3.2)

        pillars = [
            ("BRASSICACEAE", "Glucosinolate-Isothiocyanate Axis", self.COLOR_EMERALD, [
                ("Enzyme Trigger:", "Myrosinase thioglucosidase converts glucoraphanin to active sulforaphane."),
                ("Nrf2 Activation:", "Electrophilic modification of Keap1 Cys151 releases Nrf2 for antioxidant defense."),
                ("Epigenetic Scope:", "Pan-HDAC inhibition and maintenance of p21/Waf1 tumor suppression."),
                ("Estrogen Health:", "I3C/DIM shifts estrogen hydroxylation toward protective 2-OHE1 metabolite.")
            ]),
            ("TROPICAL & SUBTROPICAL", "Proteases, PMFs & Urolithins", self.COLOR_AMBER, [
                ("Proteolytic Enzymes:", "Bromelain & Papain downregulate COX-2, PGE2 and degrade inflammatory fibrin."),
                ("Citrus Bioflavonoids:", "Nobiletin and Hesperidin phosphorylate eNOS Ser1177, expanding nitric oxide."),
                ("Mitophagy Inducer:", "Pomegranate punicalagins convert via gut microbiome into Urolithin A."),
                ("Metabolic Target:", "Mangiferin activates AMPK to suppress hepatic gluconeogenesis.")
            ]),
            ("TEMPERATE SPECIES", "Anthocyanins & Flavonols", self.COLOR_PURPLE, [
                ("Blood-Brain Barrier:", "Vaccinium anthocyanins cross BBB via GLUT1/OATP to induce BDNF neurogenesis."),
                ("Mast Cell Guard:", "Quercetin inhibits IgE-mediated histamine degranulation and IL-6 release."),
                ("Sirtuin Activator:", "Resveratrol activates SIRT1-PGC-1alpha mitochondrial biogenesis."),
                ("Cardiovascular Protection:", "Improves flow-mediated dilation (FMD) and attenuates LDL oxidation.")
            ]),
            ("HERBS & SPICES", "Curcuminoids & Organosulfur", self.COLOR_CRIMSON, [
                ("Pharmacokinetic Surge:", "Piperine co-administration inhibits UGT/P-gp, boosting curcumin by +2000%."),
                ("Allicin Synthesis:", "Alliinase shears alliin upon physical crushing to generate diallyl thiosulfinates."),
                ("Anti-Inflammatory Target:", "Curcumin and Gingerols directly inhibit NF-kappaB p65 nuclear translocation."),
                ("Gastric Motility:", "Gingerols antagonize 5-HT3 receptors in chemoreceptor trigger zone.")
            ])
        ]

        for i, (p_title, p_sub, p_col, bullets) in enumerate(pillars):
            cx = start_x + i * (card_w + gap)
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, top_y, card_w, card_h)
            card.fill.solid()
            card.fill.fore_color.rgb = self.COLOR_CARD_WHITE
            card.line.color.rgb = p_col
            card.line.width = Pt(1.5)

            rail = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx + Inches(0.1), top_y, card_w - Inches(0.2), Pt(3))
            rail.fill.solid()
            rail.fill.fore_color.rgb = p_col
            rail.line.fill.background()

            ctf = slide.shapes.add_textbox(cx + Inches(0.15), top_y + Inches(0.12), card_w - Inches(0.3), Inches(0.5)).text_frame
            ctf.word_wrap = True
            p = ctf.paragraphs[0]
            p.text = p_title
            p.font.name = self.FONT_FAMILY
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = self.COLOR_TEXT_NAVY

            p_s = ctf.add_paragraph()
            p_s.text = p_sub
            p_s.font.name = self.FONT_FAMILY
            p_s.font.size = Pt(8.5)
            p_s.font.color.rgb = p_col
            p_s.font.bold = True

            btf = slide.shapes.add_textbox(cx + Inches(0.15), top_y + Inches(0.65), card_w - Inches(0.3), card_h - Inches(0.75)).text_frame
            btf.word_wrap = True
            for b_head, b_desc in bullets:
                self.append_bullet(btf, b_head, b_desc, font_size=8.5, space_after=3.5)

        # Bottom Rule & Reference
        b_rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.8), Inches(11.733), Pt(1.5))
        b_rule.fill.solid()
        b_rule.fill.fore_color.rgb = self.COLOR_HEADER_GOLD
        b_rule.line.fill.background()

        f_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.35))
        ftf = f_box.text_frame
        p = ftf.paragraphs[0]
        p.text = "Academic Standards: Harvard Medical School / Nature Reviews Molecular Cell Biology / Annual Review of Nutrition"
        p.font.name = self.FONT_FAMILY
        p.font.size = Pt(8.5)
        p.font.color.rgb = self.COLOR_TEXT_FOOTER

    def build_slide_02_taxonomy(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 2)
        self._add_header_and_footer(
            slide,
            tag="PHYTOCHEMICAL TAXONOMY",
            title="Structural Taxonomy of Plant Bioactives: 5 Primary Chemical Classes",
            takeaway="Plant secondary metabolites fall into distinct chemical structural classes that govern their cellular target affinity, solubility, and metabolic fate.",
            citation="Taxonomy Reference: Crozier, A., et al., Plant Secondary Metabolites: Occurrence, Structure and Role in the Human Diet, Wiley-Blackwell.",
            slide_num=2
        )

        w = Inches(2.2)
        h = Inches(4.75)
        gap = Inches(0.18)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        taxa = [
            ("GLUCOSINOLATES", "Sulfur-Nitrogen Glycosides", self.COLOR_EMERALD, [
                ("Structure:", "Beta-D-thioglucoside group linked to a sulfonated aldoxime and variable side chain."),
                ("Key Compounds:", "Glucoraphanin, Sinigrin, Glucobrassicin."),
                ("Botanical Sources:", "Broccoli, Kale, Brussels sprouts, Wasabi, Mustard greens."),
                ("Bioactive Form:", "Enzymatically hydrolyzed into Isothiocyanates (Sulforaphane, AITC) and Indoles (I3C).")
            ]),
            ("POLYPHENOLS", "Aromatic Hydroxyl Matrices", self.COLOR_PURPLE, [
                ("Structure:", "One or more aromatic benzene rings bound to multiple phenolic hydroxyl (-OH) groups."),
                ("Subclasses:", "Flavonoids (Quercetin, Hesperidin), Stilbenes (Resveratrol), Phenolic acids (Chlorogenic acid)."),
                ("Key Sources:", "Berries, Grapes, Citrus, Green tea, Apples."),
                ("Core Role:", "Direct ROS scavenging, metal chelation, and modulation of MAPK signaling.")
            ]),
            ("CAROTENOIDS", "Tetraterpenoid Polyenes", self.COLOR_AMBER, [
                ("Structure:", "40-carbon polyene conjugated hydrocarbon backbone; lipophilic pigments."),
                ("Carotenes vs Xanthophylls:", "Beta-carotene/Lycopene (hydrocarbon) vs Lutein/Zeaxanthin/Astaxanthin (oxygenated)."),
                ("Key Sources:", "Tomatoes, Carrots, Guava, Watermelon."),
                ("Target Action:", "Singlet oxygen quenching and macular pigment optical density preservation.")
            ]),
            ("ORGANOSULFUR", "Volatile Thiosulfinates", self.COLOR_CRIMSON, [
                ("Structure:", "Alkyl and alkenyl sulfoxides, sulfides, and di/trisulfide linkages."),
                ("Key Compounds:", "Allicin, Diallyl disulfide (DADS), Diallyl trisulfide (DATS), S-allylcysteine (SAC)."),
                ("Botanical Sources:", "Garlic (Allium sativum), Onions, Leeks, Shallots."),
                ("Core Role:", "Generates hydrogen sulfide (H2S) gaseous transmitter; activates eNOS and inhibits HMG-CoA reductase.")
            ]),
            ("TERPENOIDS", "Isoprene Polymeric Chains", self.COLOR_TEAL, [
                ("Structure:", "Formed from repetitive 5-carbon isoprene ($C_5H_8$) units via mevalonate pathway."),
                ("Subclasses:", "Monoterpenes (Limonene), Diterpenes (Carnosol), Triterpenes (Ursolic acid)."),
                ("Key Sources:", "Citrus peel, Rosemary, Oregano, Thyme."),
                ("Target Action:", "Membrane lipid fluidization, antimicrobial disruption, and suppression of inflammatory cascades.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(taxa):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=8.8, space_after=4.0)

    def build_slide_03_brassicaceae_cascade(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 3)
        self._add_header_and_footer(
            slide,
            tag="ENZYMATIC REACTION CASCADE",
            title="The Glucosinolate-Myrosinase Cascade: Sulforaphane vs. Inactive Nitriles",
            takeaway="Physical cellular disruption releases plant myrosinase; temperature and epithiospecifier protein (ESP) determine active isothiocyanate yield.",
            citation="Biochemical Mechanism: Fahey, J. W., et al., PNAS; Talalay, P., Chem. Res. Toxicol., 2012.",
            slide_num=3
        )

        # Left 2 Cards: Pathway Details
        w_card = Inches(3.75)
        top_y = Inches(1.95)
        h_card = Inches(4.75)

        ctf1 = self.create_card(slide, Inches(0.8), top_y, w_card, h_card, "STEPWISE HYDROLYSIS CASCADE", "Plant Defense Chemistry", self.COLOR_EMERALD)
        self.append_bullet(ctf1, "1. Cellular Disruption:", "Chewing, chopping, or blending ruptures plant cell vacuoles, releasing compartmentalized Glucoraphanin.", font_size=9.2)
        self.append_bullet(ctf1, "2. Myrosinase Cleavage:", "Thioglucoside glucohydrolase (EC 3.2.1.147) cleaves the thioglucosidic bond, liberating D-glucose and an unstable thiohydroximate-O-sulfonate intermediate.", font_size=9.2)
        self.append_bullet(ctf1, "3. Spontaneous Lossen Shift:", "At neutral pH and in the absence of active ESP, the aglycone intermediate spontaneously rearranges into Sulforaphane (SFN - active isothiocyanate, -N=C=S).", font_size=9.2)
        self.append_bullet(ctf1, "4. The ESP Shunt (Trap):", "Active Epithiospecifier Protein directs conversion toward Sulforaphane Nitrile, which possesses zero Nrf2 inductive activity!", font_size=9.2)

        ctf2 = self.create_card(slide, Inches(4.8), top_y, w_card, h_card, "THERMAL SENSITIVITY KINETICS", "Culinary Pharmacokinetics", self.COLOR_AMBER)
        self.append_bullet(ctf2, "Thermal Inactivation:", "Myrosinase is completely denatured at temperatures above 60°C (140°F) in under 3 minutes of boiling.", font_size=9.2)
        self.append_bullet(ctf2, "Boiling Disaster:", "Boiling broccoli leads to 80-90% loss of isothiocyanate yield due to thermal enzyme death and glucosinolate leaching into water.", font_size=9.2)
        self.append_bullet(ctf2, "Steaming Optimum:", "Light steaming for 3-4 minutes maintains glucoraphanin integrity while selectively denaturing heat-labile ESP, maximizing SFN production.", font_size=9.2)
        self.append_bullet(ctf2, "The Mustard Seed Hack:", "Adding raw ground brown mustard seed (Brassica juncea) provides exogenous heat-stable myrosinase, restoring 100% SFN yield in cooked broccoli!", font_size=9.2)

        # Right Card / Infographic Graphic
        self.add_image_panel(slide, Inches(8.8), top_y, Inches(3.733), h_card, "flowchart_glucosinolate_myrosinase.png", "Figure 1: Myrosinase Cascade & Reaction Vectors")

    def build_slide_04_nrf2_mechanism(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 4)
        self._add_header_and_footer(
            slide,
            tag="CELLULAR TRANSLOCATION",
            title="Keap1-Nrf2-ARE Pathway: Master Regulator of Phase II Cellular Detoxification",
            takeaway="Sulforaphane covalently modifies Keap1 sensor cysteines, freeing Nrf2 to enter the nucleus and initiate transcription of over 200 cytoprotective genes.",
            citation="Molecular Biology: Dinkova-Kostova, A. T., et al., Nat. Rev. Drug Discov.; Yamamoto, M., et al., Physiol. Rev.",
            slide_num=4
        )

        top_y = Inches(1.95)
        h_card = Inches(4.75)
        w_card = Inches(3.75)

        ctf1 = self.create_card(slide, Inches(0.8), top_y, w_card, h_card, "BASAL SUPPRESSION TO RELEASE", "Electrophilic Cysteine Sensing", self.COLOR_COBALT)
        self.append_bullet(ctf1, "1. Basal State Repression:", "In unstressed cells, Nrf2 is tethered in the cytoplasm by Keap1 homodimer, directing it to Cul3-Rbx1 E3 ubiquitin ligase for 26S proteasomal degradation (half-life ~20 min).", font_size=9.2)
        self.append_bullet(ctf1, "2. Cys151 Electrophilic Attack:", "Sulforaphane's electrophilic isothiocyanate carbon (-N=C=S) reacts covalently with reactive thiol groups on Keap1 Cys151, inducing conformational distortion.", font_size=9.2)
        self.append_bullet(ctf1, "3. Cul3 Dissociation:", "Keap1 conformational shift disrupts ubiquitin transfer; newly synthesized Nrf2 bypasses degradation and accumulates in cytosol.", font_size=9.2)
        self.append_bullet(ctf1, "4. Nuclear Translocation:", "Stabilized Nrf2 translocates across the nuclear membrane via its nuclear localization signal (NLS).", font_size=9.2)

        ctf2 = self.create_card(slide, Inches(4.8), top_y, w_card, h_card, "TRANSCRIPTIONAL GENE BATTERY", "Phase II Enzyme Induction", self.COLOR_EMERALD)
        self.append_bullet(ctf2, "Maf Heterodimerization:", "Inside the nucleus, Nrf2 binds to small Maf proteins (sMaf), forming a high-affinity heterodimeric transcription factor complex.", font_size=9.2)
        self.append_bullet(ctf2, "ARE Promoter Binding:", "The complex binds the cis-acting Antioxidant Response Element (ARE: 5'-TGACNNNGC-3') in promoter regions.", font_size=9.2)
        self.append_bullet(ctf2, "HO-1 Induction:", "Upregulates Heme Oxygenase-1 (HO-1), catabolizing pro-oxidant heme into biliverdin, iron, and anti-inflammatory carbon monoxide (CO).", font_size=9.2)
        self.append_bullet(ctf2, "NQO1 & GST:", "Dramatically induces NAD(P)H:quinone oxidoreductase-1 (NQO1) and Glutathione S-Transferases (GST), clearing toxic quinones and electrophiles.", font_size=9.2)

        # Right Visual Infographic Panel
        self.add_image_panel(slide, Inches(8.8), top_y, Inches(3.733), h_card, "flowchart_nrf2_pathway.png", "Figure 2: Nrf2-Keap1-ARE Induction Cascade")

    def build_slide_05_epigenetics_hdac(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 5)
        self._add_header_and_footer(
            slide,
            tag="EPIGENETIC REGULATION",
            title="Epigenetic Modulations: Histone Deacetylase (HDAC) & DNA Methylation",
            takeaway="Cruciferous bioactives reverse aberrant epigenetic silencing in cancer and aging by acting as competitive pan-HDAC and DNMT inhibitors.",
            citation="Epigenetics Research: Dashwood, R. H., & Ho, E., Cancer Lett.; Myzak, M. C., et al., Cancer Res., 2006.",
            slide_num=5
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        top_y = Inches(1.95)
        start_x = Inches(0.8)

        cards = [
            ("HISTONE DEACETYLASE (HDAC) INHIBITION", "Chromatin Relaxation & Gene Re-expression", self.COLOR_PURPLE, [
                ("Mechanism of Action:", "Sulforaphane metabolites (SFN-cysteine and SFN-N-acetylcysteine) fit into the catalytic pocket of HDAC enzymes, chelating active-site zinc ions ($Zn^{2+}$)."),
                ("Target Enzymes:", "Pan-HDAC inhibition across Class I (HDAC1, 2, 3, 8) and Class II (HDAC4, 5, 7, 9) enzymes."),
                ("Histone Hyperacetylation:", "Suppression of HDAC activity causes global hyperacetylation of core histones H3 and H4, shifting chromatin from heterochromatin to open euchromatin."),
                ("Tumor Suppressor Reactivation:", "Unlocks silenced promoter regions of p21/Waf1 and p16/INK4a, triggering $G_1/S$ cell cycle arrest in transformed pre-malignant cells.")
            ]),
            ("DNA METHYLTRANSFERASE (DNMT) ATTENUATION", "Promoter Demethylation Kinetics", self.COLOR_COBALT, [
                ("Catalytic Blockade:", "Both Sulforaphane and 3,3'-Diindolylmethane (DIM) downregulate expression and enzymatic activity of DNMT1 and DNMT3b."),
                ("CpG Island Demethylation:", "Reduces hypermethylation of dense CpG island clusters located within promoter regions of protective genes."),
                ("Nrf2 Promoter Re-opening:", "Re-activates silenced Nrf2 expression in aging or hypermethylated prostate and colon tissues, restoring endogenous redox defense."),
                ("Synergistic Epigenome Remodeling:", "Simultaneous HDAC inhibition and DNMT suppression produces a synergistic transcriptional release unmatched by synthetic single-target drugs.")
            ]),
            ("TELOMERASE & microRNA MODULATION", "Telomere Length & Non-Coding RNAs", self.COLOR_TEAL, [
                ("hTERT Repression:", "Downregulates human telomerase reverse transcriptase (hTERT) transcription in malignant cells, accelerating senescence in aberrant clones."),
                ("microRNA Network Tuning:", "Modulates key oncogenic and tumor suppressor microRNAs: upregulates miR-140 and miR-200 while downregulating oncogenic miR-21 and miR-155."),
                ("Translational Control:", "Normalizes ribosomal translation of inflammatory cytokines (TNF-alpha, IL-1beta) in human monocytes."),
                ("Heritability:", "Exerts transgenerational epigenetic protection against carcinogen-induced tissue mutations in animal lineage studies.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_06_i3c_dim(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 6)
        self._add_header_and_footer(
            slide,
            tag="HORMONE METABOLISM",
            title="Indole-3-Carbinol (I3C) and 3,3'-Diindolylmethane (DIM): Estrogen Ratio Shift",
            takeaway="Gastric acid condensation transforms dietary I3C into bioactive DIM, shifting estrogen catabolism toward protective 2-OHE1.",
            citation="Endocrinology Source: Bradlow, H. L., et al., Ann. N.Y. Acad. Sci.; Thomson, C. A., et al., Breast Cancer Res. Treat.",
            slide_num=6
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        top_y = Inches(1.95)
        start_x = Inches(0.8)

        cards = [
            ("GASTRIC CONDENSATION CASCADE", "Formation of Polycyclic Condensation Products", self.COLOR_CRIMSON, [
                ("Parent Precursor:", "Glucobrassicin in Brassica vegetables is hydrolyzed by myrosinase into unstable Indole-3-Carbinol (I3C)."),
                ("Gastric Acid Condensation:", "In the low pH environment of the stomach ($pH < 2.5$), I3C molecules undergo spontaneous acid-catalyzed oligomerization."),
                ("Primary Metabolite:", "Over 60% of I3C condenses into 3,3'-Diindolylmethane (DIM); remaining portions yield indolo[3,2-b]carbazole (ICZ)."),
                ("Clinical Trap:", "Patients taking Proton Pump Inhibitors (PPIs) have elevated gastric pH, which blocks I3C condensation into bioactive DIM!")
            ]),
            ("CYP450 ENZYME REGULATION & 2:16-OHE1 RATIO", "Protective vs. Genotoxic Estrogen Pathways", self.COLOR_EMERALD, [
                ("CYP1A1 / CYP1A2 Induction:", "DIM potently induces cytochrome P450 1A1/1A2 through AhR-dependent signaling pathways."),
                ("The 2-OHE1 Pathway (Protective):", "Drives hepatic 2-hydroxylation of estrone ($E_1$) and estradiol ($E_2$), generating 2-hydroxyestrone (2-OHE1), a non-genotoxic, weakly estrogenic catabolite."),
                ("Suppression of 16alpha-OHE1 (Genotoxic):", "Suppresses CYP1B1, reducing formation of 16alpha-hydroxyestrone, a potent estrogen receptor agonist linked to DNA adducts."),
                ("Clinical Biomarker:", "Elevating the urinary 2:16-OHE1 ratio (>2.0) correlates with significantly reduced risk of hormone-dependent breast and prostate neoplasms.")
            ]),
            ("ANDROGEN RECEPTOR & PROSTATE HEALTH", "Nuclear Export & PSA Downregulation", self.COLOR_COBALT, [
                ("AR Downregulation:", "DIM binds androgen receptors, suppressing AR nuclear translocation and downregulating Prostate-Specific Antigen (PSA)."),
                ("Akt/NF-kappaB Blockade:", "Inhibits Akt phosphorylation and blocks IkappaBalpha kinase (IKK), shutting down downstream inflammatory transcription."),
                ("Induction of Apoptosis:", "Selectively triggers mitochondrial apoptosis in pre-neoplastic epithelial cells via Bcl-2 downregulation and Bax activation."),
                ("Clinical Human Trials:", "Daily oral supplementation with 100-300 mg bioavailable microencapsulated DIM demonstrated stabilization of serum PSA and normalized 2:16-OHE1 ratios.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_07_tropical_fruits(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 7)
        self._add_header_and_footer(
            slide,
            tag="TROPICAL BIOACTIVES",
            title="Tropical Fruits: Proteolytic Enzymes, Mangiferin & Carotenoid Profiles",
            takeaway="Tropical fruits yield unique sulfhydryl endopeptidases, polyphenolic xanthones, and carotenoids supporting joint, digestive, and metabolic health.",
            citation="Food Chemistry: Maurer, H. R., Cell. Mol. Life Sci.; Pavan, R., et al., Biotechnol. Res. Int.",
            slide_num=7
        )

        w = Inches(2.78)
        h = Inches(4.75)
        gap = Inches(0.2)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        fruits = [
            ("PINEAPPLE (BROMELAIN)", "Sulfhydryl Endopeptidase Complex", self.COLOR_AMBER, [
                ("Botanical Origin:", "Stem and fruit of Ananas comosus (Bromeliaceae)."),
                ("Active Fractions:", "Mixture of thiol endopeptidases, peroxidase, acid phosphatase, and cellulase."),
                ("Anti-Inflammatory:", "Inhibits thromboxane synthase; downregulates PGE2 synthesis and degrades circulating immune complexes."),
                ("Tissue Fibrinolysis:", "Cleaves fibrin deposits in inflamed tissues, reducing edema and restoring microvascular capillary permeability.")
            ]),
            ("PAPAYA (PAPAIN & CHYMOPAPAIN)", "Broad-Spectrum Proteolytic Matrix", self.COLOR_EMERALD, [
                ("Botanical Origin:", "Latex and pulp of Carica papaya (Caricaceae)."),
                ("Proteolytic Capacity:", "Hydrolyzes peptide bonds involving basic amino acids (arginine, lysine) across a broad pH range (3.0 - 10.5)."),
                ("Digestive Support:", "Accelerates gastric emptying of heavy protein meals, alleviating functional dyspepsia."),
                ("Wound Debridement:", "Topical and systemic enzymatic clearing of necrotic exudates without damaging viable cellular tissue.")
            ]),
            ("MANGO (MANGIFERIN)", "Glucosylxanthone Bioactive", self.COLOR_TEAL, [
                ("Botanical Origin:", "Pulp, peel, and kernel of Mangifera indica (Anacardiaceae)."),
                ("Unique Structure:", "C-glucosylxanthone with high chemical stability against gastrointestinal enzymatic degradation."),
                ("Metabolic Powerhouse:", "Activates AMP-activated protein kinase (AMPK), stimulating GLUT4 glucose transporter translocation."),
                ("Cardioprotection:", "Attenuates ischemia-reperfusion injury and suppresses vascular adhesion molecule (VCAM-1) expression.")
            ]),
            ("GUAVA & PASSION FRUIT", "Lycopene & Piceatannol Synergy", self.COLOR_CRIMSON, [
                ("Guava (Psidium guajava):", "Contains more lycopene per 100g than fresh tomatoes (~5.2 mg/100g), with exceptional bioavailability without oil cooking."),
                ("Passion Fruit (Passiflora edulis):", "Exceptionally rich in Piceatannol, a hydroxylated resveratrol analogue with superior metabolic stability."),
                ("Insulin Sensitivity:", "Piceatannol clinical trials demonstrate marked reduction in HOMA-IR in overweight adults."),
                ("Vascular Tone:", "Suppresses endothelin-1 (ET-1) release, promoting systemic arterial compliance.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(fruits):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.0, space_after=4.0)

    def build_slide_08_bromelain_flowchart(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 8)
        self._add_header_and_footer(
            slide,
            tag="ENZYMATIC FLOWCHART",
            title="Bromelain & Papain: Anti-Inflammatory Cascade & Fibrinolytic Degradation",
            takeaway="Sulfhydryl endopeptidases clear fibrin matrices, downregulate arachidonic acid inflammatory cascades, and accelerate soft tissue recovery.",
            citation="Enzymology Review: Fitzhugh, D. J., et al., Clin. Immunol.; Rathnavelu, V., et al., Biomed. Rep.",
            slide_num=8
        )

        w_card = Inches(5.7)
        top_y = Inches(1.95)
        h_card = Inches(4.75)

        ctf1 = self.create_card(slide, Inches(0.8), top_y, w_card, h_card, "STEPWISE MECHANISTIC CASCADE", "Multimodal Anti-Inflammatory Vectors", self.COLOR_AMBER)
        self.append_bullet(ctf1, "Stage 1: Systemic Absorption:", "Intact bromelain molecules (~33 kDa) cross the intestinal barrier via paracellular and pinocytic pathways; bind alpha-2-macroglobulin in plasma.", font_size=9.2)
        self.append_bullet(ctf1, "Stage 2: Fibrinolysis Activation:", "Directly cleaves intravascular and interstitial fibrin deposits; converts plasminogen to active plasmin, eliminating inflammatory tissue edema.", font_size=9.2)
        self.append_bullet(ctf1, "Stage 3: Downregulation of PGE2 & COX-2:", "Selectively inhibits prostaglandin E2 (PGE2) synthesis while preserving protective PGI2 (prostacyclin), avoiding the gastric mucosal damage of NSAIDs.", font_size=9.2)
        self.append_bullet(ctf1, "Stage 4: Cytokine Attenuation:", "Decreases secretion of pro-inflammatory cytokines IL-1beta, IL-6, and TNF-alpha by activated macrophages.", font_size=9.2)
        self.append_bullet(ctf1, "Stage 5: Clinical Endpoints:", "Accelerates surgical wound healing, reduces acute sports injury hematomas, and relieves osteoarthritic knee stiffness.", font_size=9.2)

        ctf2 = self.create_card(slide, Inches(6.8), top_y, w_card, h_card, "CLINICAL TRANSLATION & DOSAGE PARAMETERS", "Evidence-Based Application", self.COLOR_CRIMSON)
        self.append_bullet(ctf2, "Digestive vs Systemic Timing:", "• With Meals: Acts strictly as a digestive enzyme in the gut lumen to hydrolyze dietary proteins.\n• Between Meals (Empty Stomach): Absorbed systemically into the bloodstream to execute anti-inflammatory fibrinolysis.", font_size=9.2)
        self.append_bullet(ctf2, "Standardized Potency Units:", "Measured in Gelatin Digesting Units (GDU) or Milk Clotting Units (MCU). Therapeutic potency requires >=2,400 GDU/gram.", font_size=9.2)
        self.append_bullet(ctf2, "Target Clinical Dosage:", "500 - 1,000 mg daily (in divided doses taken 1 hour before or 2 hours after meals).", font_size=9.2)
        self.append_bullet(ctf2, "Safety Profile:", "Exceptional tolerability; caution advised in patients on anticoagulant therapy (warfarin, apixaban) due to intrinsic antiplatelet activity.", font_size=9.2)

    def build_slide_09_mangiferin_metabolism(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 9)
        self._add_header_and_footer(
            slide,
            tag="METABOLIC BIOCHEMISTRY",
            title="Mangiferin & Glycemic Homeostasis: AMPK Activation & Gluconeogenesis",
            takeaway="Mangiferin mimics physical exercise signaling by phosphorylating AMPK, enhancing insulin-independent glucose uptake into skeletal muscle.",
            citation="Metabolism Research: Sellamuthu, P. S., et al., J. Pharm. Pharmacol.; Na, L., et al., Br. J. Nutr.",
            slide_num=9
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        top_y = Inches(1.95)
        start_x = Inches(0.8)

        cards = [
            ("AMPK PHOSPHORYLATION & GLUT4", "Exercise-Mimetic Pathway", self.COLOR_TEAL, [
                ("Catalytic Activation:", "Mangiferin stimulates phosphorylation of AMP-activated protein kinase (AMPK) at Thr172 in skeletal myocytes."),
                ("AS160 Phosphorylation:", "Phosphorylated AMPK inactivates Akt substrate of 160 kDa (AS160), releasing intracellular Rab-GTPase vesicles."),
                ("GLUT4 Translocation:", "Triggers rapid translocation of Glucose Transporter-4 (GLUT4) from intracellular storage vesicles to the plasma membrane."),
                ("Insulin-Independent Uptake:", "Facilitates peripheral glucose disposal independent of insulin receptor binding, bypassing insulin resistance!")
            ]),
            ("HEPATIC GLUCONEOGENESIS SUPPRESSION", "Downregulation of PEPCK and G6Pase", self.COLOR_EMERALD, [
                ("Transcription Suppression:", "Inhibits FoxO1 nuclear retention in hepatocytes, suppressing key rate-limiting gluconeogenic genes."),
                ("PEPCK Blockade:", "Downregulates Phosphoenolpyruvate Carboxykinase (PEPCK), cutting off the conversion of oxaloacetate to phosphoenolpyruvate."),
                ("Glucose-6-Phosphatase:", "Suppresses Glucose-6-Phosphatase (G6Pase), preventing the release of free glucose into the systemic bloodstream."),
                ("End-Result:", "Reduces fasting blood glucose spikes and lowers hepatic steatosis in metabolic dysfunction-associated steatotic liver disease (MASLD).")
            ]),
            ("LIPID OXIDATION & PPAR-alpha AXIS", "Triglyceride & Lipid Clearing", self.COLOR_AMBER, [
                ("PPAR-alpha Activation:", "Upregulates Peroxisome Proliferator-Activated Receptor Alpha (PPAR-alpha) expression in hepatic tissue."),
                ("Mitochondrial CPT-1:", "Induces Carnitine Palmitoyltransferase-1 (CPT-1), accelerating the entry of fatty acids into mitochondria for beta-oxidation."),
                ("VLDL Export Attenuation:", "Reduces hepatic de novo lipogenesis and suppresses assembly of ApoB-100 containing VLDL particles."),
                ("Clinical Trials:", "12-week randomized trials in dyslipidemic patients showed significant reductions in serum triglycerides (-15.8%) and free fatty acids.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_10_citrus_pmfs(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 10)
        self._add_header_and_footer(
            slide,
            tag="SUBTROPICAL FLAVONOIDS",
            title="Subtropical Citrus Flavonoids & Polymethoxyflavones (PMFs)",
            takeaway="Lipophilic polymethoxyflavones (Nobiletin, Tangeretin) easily cross biological membranes, regulating hepatic lipid secretion and neuroprotection.",
            citation="Citrus Science: Manthey, J. A., et al., J. Agric. Food Chem.; Mulvihill, E. E., et al., Diabetes.",
            slide_num=10
        )

        w = Inches(2.78)
        h = Inches(4.75)
        gap = Inches(0.2)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        cards = [
            ("NOBILETIN", "Hexa-Methoxyflavone Powerhouse", self.COLOR_AMBER, [
                ("Structure:", "5,6,7,8,3',4'-hexamethoxyflavone found in sweet orange and mandarin peels."),
                ("Superior Bioavailability:", "All six phenolic positions are O-methylated, rendering it highly lipophilic and resistant to intestinal glucuronidation."),
                ("Circadian Clock Reset:", "Directly binds ROR-alpha/gamma nuclear receptors, enhancing circadian rhythm amplitude and metabolic vigor."),
                ("Neuroprotection:", "Crosses the blood-brain barrier; restores CREB phosphorylation and rescues memory deficits in tauopathy models.")
            ]),
            ("TANGERETIN", "Penta-Methoxyflavone Regulator", self.COLOR_EMERALD, [
                ("Structure:", "5,6,7,8,4'-pentamethoxyflavone; intensely concentrated in tangerine and clementine flavedo."),
                ("ApoB Clearance:", "Suppresses microsomal triglyceride transfer protein (MTP), preventing hepatic assembly of atherogenic ApoB particles."),
                ("Cell Adhesion:", "Up-regulates E-cadherin expression on endothelial and epithelial membranes, maintaining barrier integrity."),
                ("Microglial Quenching:", "Crosses into central nervous system, suppressing microglial TNF-alpha and iNOS expression.")
            ]),
            ("HESPERIDIN & HESPERETIN", "Microvascular Glycoside Matrix", self.COLOR_COBALT, [
                ("Structure:", "Hesperetin-7-O-rutinoside, abundant in sweet oranges and lemons."),
                ("Intestinal Hydrolysis:", "Deglycosylated by gut microbiota alpha-rhamnosidase into active aglycone Hesperetin."),
                ("Microvascular Resistance:", "Strengthens capillary resistance, lowers capillary fragility, and prevents fluid extravasation in chronic venous disease."),
                ("Endothelial FMD:", "Significantly improves Flow-Mediated Dilation (FMD) in double-blind clinical trials of hypertensive adults.")
            ]),
            ("NARINGENIN", "Bitter Grapefruit Bioactive", self.COLOR_CRIMSON, [
                ("Botanical Source:", "Predominant flavonoid in grapefruit (Citrus paradisi) and sour orange."),
                ("PPAR-gamma & AMPK:", "Dual agonist activating both PPAR-alpha and PPAR-gamma, accelerating hepatic fatty acid oxidation."),
                ("Drug-Nutrient Warning:", "Inhibits intestinal CYP3A4 and P-glycoprotein, drastically increasing serum concentrations of statins, calcium channel blockers, and cyclosporine!"),
                ("Clinical Window:", "Requires a mandatory 4-hour temporal separation from CYP3A4-metabolized pharmaceuticals.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.0, space_after=4.0)

    def build_slide_11_enos_activation(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 11)
        self._add_header_and_footer(
            slide,
            tag="VASCULAR ENDOTHELIUM",
            title="Endothelial Nitric Oxide (eNOS) Activation by Citrus Flavonoids",
            takeaway="Hesperetin and Nobiletin stimulate Akt-mediated Ser1177 phosphorylation of eNOS, driving vascular relaxation and blood pressure normalization.",
            citation="Cardiovascular Physiology: Rizza, S., et al., J. Clin. Endocrinol. Metab.; Valls, R. M., et al., Eur. J. Nutr.",
            slide_num=11
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        top_y = Inches(1.95)
        start_x = Inches(0.8)

        cards = [
            ("PHOSPHORYLATION AT SER1177", "Post-Translational Enzyme Activation", self.COLOR_COBALT, [
                ("Upstream Kinase Cascade:", "Citrus bioflavonoids activate the intracellular Phosphoinositide 3-Kinase (PI3K) / Akt kinase signaling axis."),
                ("Ser1177 Phosphorylation:", "Akt directly phosphorylates endothelial Nitric Oxide Synthase (eNOS) at Serine 1177, shifting the enzyme into its high-activity state."),
                ("Dephosphorylation at Thr495:", "Simultaneously dephosphorylates the inhibitory Threonine 495 residue, removing constitutive basal enzyme suppression."),
                ("Substrate Utilization:", "Enhances eNOS affinity for its physiological substrate L-arginine, boosting catalytic conversion to L-citrulline and Nitric Oxide ($NO^\bullet$).")
            ]),
            ("cGMP & VASCULAR SMOOTH MUSCLE RELAXATION", "The Vasodilation Mechanics", self.COLOR_EMERALD, [
                ("Gas Diffusion:", "Endothelial $NO^\bullet$ gas diffuses rapidly across the internal elastic lamina into adjacent vascular smooth muscle cells."),
                ("Guanylyl Cyclase Binding:", "Binds the prosthetic heme group of Soluble Guanylyl Cyclase (sGC), accelerating enzymatic turnover by 400-fold."),
                ("cGMP Generation:", "Converts GTP to cyclic Guanosine Monophosphate (cGMP), activating cGMP-dependent Protein Kinase G (PKG)."),
                ("Calcium Sequestering:", "PKG phosphorylates phospholamban and opens potassium channels, driving $Ca^{2+}$ out of the cytosol and inducing smooth muscle relaxation and vasodilation.")
            ]),
            ("PREVENTION OF eNOS UNCOUPLING", "Preserving Tetrahydrobiopterin ($BH_4$)", self.COLOR_AMBER, [
                ("The Uncoupling Hazard:", "Under severe oxidative stress, peroxynitrite oxidizes tetrahydrobiopterin ($BH_4$) to inactive $BH_2$, uncoupling eNOS into a superoxide generator!"),
                ("Flavonoid Scavenging:", "Citrus flavonoids directly quench intracellular superoxide ($O_2^{\bullet-}$), preventing the destructive formation of peroxynitrite ($ONOO^-$)."),
                ("Restoring $BH_4$ Levels:", "Upregulates GTP cyclohydrolase I (GTPCH-1), the rate-limiting enzyme in de novo $BH_4$ biosynthesis."),
                ("Clinical Outcome:", "Restores physiological endothelial Flow-Mediated Dilation (FMD) and lowers systolic blood pressure by 4-6 mmHg in clinical trials.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_12_punicalagin_urolithin(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 12)
        self._add_header_and_footer(
            slide,
            tag="MICROBIOME-MITOCHONDRIAL AXIS",
            title="Pomegranate Ellagitannins to Urolithin A: Mitophagy Induction & Muscle Health",
            takeaway="Dietary punicalagins are transformed by gut commensal bacteria into Urolithin A, which selectively clears dysfunctional mitochondria.",
            citation="Mitophagy Discovery: Ryu, D., et al., Nature Medicine; Andreux, P. A., et al., Nat. Metab., 2019.",
            slide_num=12
        )

        top_y = Inches(1.95)
        h_card = Inches(4.75)
        w_card = Inches(3.75)

        ctf1 = self.create_card(slide, Inches(0.8), top_y, w_card, h_card, "GUT MICROBIAL BIOTRANSFORMATION", "From Punicalagin to Urolithin A", self.COLOR_CRIMSON)
        self.append_bullet(ctf1, "1. Giant Ellagitannin Ingestion:", "Pomegranate peel and arils contain massive polyphenolic Punicalagins (MW 1,084 Da), which are too large for systemic absorption in the small intestine.", font_size=9.2)
        self.append_bullet(ctf1, "2. Hydrolysis to Ellagic Acid:", "In the gut lumen, punicalagins are hydrolyzed into Ellagic Acid, which passes into the colon.", font_size=9.2)
        self.append_bullet(ctf1, "3. Commensal Bacterial Metabolism:", "Specific human gut commensals (Gordonibacter urolithinfaciens & Ellagibacter isohannegifaciens) sequentially lactonize and dehydroxylate ellagic acid.", font_size=9.2)
        self.append_bullet(ctf1, "4. The Producer Phenotype:", "Only ~40% of humans naturally possess the gut microbiome consortium required to generate high levels of Urolithin A (Metabotype A)!", font_size=9.2)

        ctf2 = self.create_card(slide, Inches(4.8), top_y, w_card, h_card, "PINK1-PARKIN MITOPHAGY CASCADE", "Mitochondrial Quality Control", self.COLOR_PURPLE)
        self.append_bullet(ctf2, "Depolarized Mitochondria:", "Damaged, aging mitochondria lose their inner membrane potential (Delta-Psi-m), accumulating ROS and leaking cytochrome c.", font_size=9.2)
        self.append_bullet(ctf2, "PINK1 Stabilization:", "Urolithin A triggers stabilization of PTEN-induced kinase 1 (PINK1) on the outer mitochondrial membrane of damaged organelles.", font_size=9.2)
        self.append_bullet(ctf2, "Parkin Recruitment:", "Stabilized PINK1 phosphorylates ubiquitin and recruits the E3 ubiquitin ligase Parkin, polyubiquitinating outer membrane proteins.", font_size=9.2)
        self.append_bullet(ctf2, "Autophagosome Engulfment:", "Recruits p62/SQSTM1 and LC3-II autophagosomal membranes, routing defective mitochondria to lysosomes for clean catabolic degradation and renewal.", font_size=9.2)

        # Right Infographic Panel
        self.add_image_panel(slide, Inches(8.8), top_y, Inches(3.733), h_card, "flowchart_urolithin_mitophagy.png", "Figure 3: Urolithin A Biogenesis & Mitophagy")

    def build_slide_13_temperate_fruits(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 13)
        self._add_header_and_footer(
            slide,
            tag="TEMPERATE PHYTOCHEMISTRY",
            title="Temperate Fruits: Rosaceae & Vitaceae Secondary Metabolite Profiles",
            takeaway="Temperate fruits contain specialized dihydrochalcones, proanthocyanidins, and stilbenes targeting intestinal SGLT transporters and vascular compliance.",
            citation="Botanical Review: Hyson, D. A., Adv. Nutr.; Manach, C., et al., Am. J. Clin. Nutr.",
            slide_num=13
        )

        w = Inches(2.78)
        h = Inches(4.75)
        gap = Inches(0.2)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        fruits = [
            ("APPLES (PHLORIDZIN & QUERCETIN)", "Dihydrochalcone & Flavonol Duo", self.COLOR_EMERALD, [
                ("Phloridzin Mechanism:", "Selective natural competitive inhibitor of Sodium-Glucose Cotransporter 1 (SGLT1) and SGLT2."),
                ("Glycemic Moderation:", "Slows postprandial glucose absorption across the intestinal brush border membrane."),
                ("Apple Peel Quercetin:", "Concentrated in the skin as quercetin-3-galactoside; upregulates endothelial nitric oxide."),
                ("Oligomeric Procyanidins:", "Inhibits pancreatic lipase and supports beneficial Akkermansia muciniphila gut bloom.")
            ]),
            ("GRAPES (RESVERATROL & TANNINGS)", "Stilbene Sirtuin Activator", self.COLOR_PURPLE, [
                ("Botanical Origin:", "Skin of red grapes (Vitis vinifera) synthesize trans-resveratrol as a phytoalexin against fungal rot."),
                ("SIRT1 Activation:", "Allosterically activates NAD+-dependent deacetylase Sirtuin-1 (SIRT1), mimicking caloric restriction."),
                ("PGC-1alpha Deacetylation:", "SIRT1 deacetylates PGC-1alpha, driving mitochondrial biogenesis in skeletal muscle."),
                ("Endothelial Fibrinolysis:", "Stimulates tissue-type plasminogen activator (t-PA) synthesis in vascular endothelium.")
            ]),
            ("TART CHERRIES (ANTHOCYANINS)", "Anti-Hyperuricemic Analgesic", self.COLOR_CRIMSON, [
                ("Cyanidin-3-Glucoside:", "Abundant anthocyanin responsible for the deep ruby pigmentation of Montmorency cherries."),
                ("COX-1 & COX-2 Inhibition:", "Exhibits in vitro cyclooxygenase inhibitory potency comparable to ibuprofen and naproxen."),
                ("Uric Acid Transport:", "Inhibits URAT1 renal transporter, accelerating urinary excretion of serum uric acid."),
                ("Exercise Recovery:", "Significantly reduces delayed onset muscle soreness (DOMS) and creatine kinase post-marathon.")
            ]),
            ("PEARS & PLUMS (CHLOROGENIC ACID)", "Hydroxycinnamic Acid Matrix", self.COLOR_AMBER, [
                ("5-Caffeoylquinic Acid:", "Primary hydroxycinnamic acid ester in European and Asian pear varieties."),
                ("Hepatic G6Pase Inhibition:", "Suppresses glucose-6-phosphatase, regulating morning fasting hepatic glucose output."),
                ("Plum Proanthocyanidins:", "Dense polymeric flavan-3-ols with superior hydroxyl radical scavenging capacity."),
                ("Bone Health:", "Dried plum polyphenols suppress osteoclastogenesis via downregulation of NFATc1 signaling.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(fruits):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.0, space_after=4.0)

    def build_slide_14_quercetin_mast_cells(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 14)
        self._add_header_and_footer(
            slide,
            tag="IMMUNE MODULATION",
            title="Quercetin: Mast Cell Membrane Stabilization & Anti-Allergic Signaling",
            takeaway="Quercetin prevents calcium influx into sensitized mast cells, halting histamine degranulation and leukotriene biosynthesis.",
            citation="Immunology Research: Weng, Z., et al., PLoS ONE; Mlcek, J., et al., Molecules, 2016.",
            slide_num=14
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        top_y = Inches(1.95)
        start_x = Inches(0.8)

        cards = [
            ("MAST CELL CALCIUM INFLUX BLOCKADE", "Membrane Stabilization Dynamics", self.COLOR_COBALT, [
                ("FceRI Receptor Crosslinking:", "When allergens cross-link IgE antibodies bound to high-affinity FceRI receptors, it triggers intracellular signaling."),
                ("Phospholipase C-gamma Blockade:", "Quercetin directly inhibits phosphorylation of Lyn kinase and Syk kinase, preventing downstream activation of PLC-gamma."),
                ("IP3 & Calcium Mobilization:", "Blocks generation of inositol 1,4,5-trisphosphate (IP3), completely arresting calcium release from the endoplasmic reticulum."),
                ("Degranulation Arrest:", "Without the requisite cytosolic $Ca^{2+}$ spike, secretory granules cannot fuse with the plasma membrane, preventing histamine dump!")
            ]),
            ("LEUKOTRIENE & CYTOKINE SUPPRESSION", "Downstream Inflammatory Mediator Shutdown", self.COLOR_EMERALD, [
                ("5-Lipoxygenase (5-LOX) Inhibition:", "Directly inhibits 5-lipoxygenase enzyme, suppressing production of bronchoconstrictive leukotrienes ($LTA_4$, $LTB_4$, $LTC_4$)."),
                ("COX-2 Downregulation:", "Suppresses inducible Cyclooxygenase-2 expression, eliminating pro-inflammatory Prostaglandin D2 ($PGD_2$)."),
                ("IL-6 and TNF-alpha:", "Inhibits nuclear translocation of NF-kappaB p65, blocking transcription of inflammatory interleukins in mast cells and basophils."),
                ("Clinical Synergy:", "Co-administration with Vitamin C (Ascorbic Acid) prevents quercetin autoxidation, multiplying its biological half-life and cellular potency.")
            ]),
            ("VASCULAR CELL ADHESION MOLECULES", "Endothelial Extravasation Attenuation", self.COLOR_PURPLE, [
                ("ICAM-1 & VCAM-1 Suppression:", "Quercetin blocks TNF-alpha-induced expression of Vascular Cell Adhesion Molecule-1 (VCAM-1) and ICAM-1 on endothelial cells."),
                ("Leukocyte Rolling Block:", "Prevents the rolling, tethering, and transmigration of circulating eosinophils and neutrophils into inflamed nasal and bronchial mucosa."),
                ("Allergic Rhinitis Efficacy:", "Double-blind clinical trials demonstrate marked reduction in sneezing, rhinorrhea, and ocular pruritus within 2 weeks of supplementation."),
                ("Formulation Optimization:", "Standard quercetin aglycone has low absorption (<2%); phospholipid-bound Quercetin Phytosome increases bioavailability by 20-fold!")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_15_berry_anthocyanins_bbb(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 15)
        self._add_header_and_footer(
            slide,
            tag="NEUROPROTECTIVE AXIS",
            title="Temperate Berries: Anthocyanin Transit Across Blood-Brain Barrier (BBB)",
            takeaway="Berry anthocyanins cross the endothelial BBB via GLUT1/OATP transporters, activating the CREB-BDNF neuroplasticity pathway.",
            citation="Neurobiology Source: Youdim, K. A., et al., Free Radic. Biol. Med.; Devore, E. E., et al., Ann. Neurol.",
            slide_num=15
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        top_y = Inches(1.95)
        start_x = Inches(0.8)

        cards = [
            ("BLOOD-BRAIN BARRIER TRANSIT", "Transporter-Mediated Endothelial Crossing", self.COLOR_PURPLE, [
                ("Intact Glucoside Permeation:", "Cyanidin-3-glucoside and Delphinidin-3-glucoside cross cerebral microvascular endothelial cells intact without prior deglycosylation."),
                ("GLUT1 & OATP Transporters:", "Exploits the Glucose Transporter-1 (GLUT1) and Organic Anion Transporting Polypeptide (OATP1A2) to traverse the tight BBB."),
                ("Hippocampal Localization:", "Autoradiography confirms selective bioaccumulation of anthocyanin metabolites in the hippocampus and cerebral cortex."),
                ("Pericyte & Astrocyte Shield:", "Fortifies tight junction proteins (Claudin-5, Occludin, ZO-1), reducing BBB permeability and preventing microvascular leakage.")
            ]),
            ("CREB-BDNF PLASTICITY CASCADE", "Synaptic Remodeling & Long-Term Potentiation", self.COLOR_COBALT, [
                ("ERK1/2 & Akt Signaling:", "Anthocyanins activate Extracellular Signal-Regulated Kinase (ERK1/2) and Akt pathways in pyramidal neurons."),
                ("CREB Phosphorylation:", "Triggers phosphorylation of cAMP Response Element-Binding Protein (CREB) at Serine 133."),
                ("BDNF Transcription:", "Drives robust transcription and secretion of Brain-Derived Neurotrophic Factor (BDNF)."),
                ("Synaptogenesis:", "Stimulates TrkB receptor signaling, upregulating PSD-95 and Synaptophysin to enhance hippocampal Long-Term Potentiation (LTP) and memory formation.")
            ]),
            ("MICROGLIAL QUENCHING & HUMAN TRIALS", "Halting Neuroinflammation in Aging", self.COLOR_EMERALD, [
                ("Microglial Deactivation:", "Reverses the primed, hyper-reactive phenotype of aging microglia, suppressing excessive nitric oxide ($NO^\bullet$) and IL-1beta release."),
                ("Amyloid-beta Aggregation:", "Inhibits fibrillization of amyloid-beta ($A\beta_{1-42}$) and tau hyperphosphorylation via GSK-3beta inhibition."),
                ("The Nurses' Health Study:", "Prospective cohort study of 16,010 women demonstrated that high blueberry and strawberry intake delayed cognitive aging by up to 2.5 years."),
                ("Target Intake:", "1/2 to 1 cup of fresh or wild frozen blueberries daily (~250-500 mg total polyphenols) provides therapeutic neuroprotection.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_16_curcuminoids_pharmacology(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 16)
        self._add_header_and_footer(
            slide,
            tag="CULINARY SPICES",
            title="Zingiberaceae Pharmacology: Curcuminoids & Multi-Target Molecular Actions",
            takeaway="Curcumin is a pleiotropic polyphenol that directly inhibits NF-kappaB p65 nuclear translocation, downregulating hundreds of pro-inflammatory genes.",
            citation="Pharmacological Reviews: Aggarwal, B. B., et al., Ann. N.Y. Acad. Sci.; Kunnumakkara, A. B., et al., Br. J. Pharmacol.",
            slide_num=16
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        top_y = Inches(1.95)
        start_x = Inches(0.8)

        cards = [
            ("THE CURCUMINOID TRIAD", "Chemical Fractions of Turmeric Rhizome", self.COLOR_AMBER, [
                ("Botanical Origin:", "Rhizome of Curcuma longa (Zingiberaceae); constitutes 3 - 5% of raw turmeric dry weight."),
                ("Curcumin (Diferuloylmethane):", "Principal bioactive (77%); symmetric bis-alpha,beta-unsaturated beta-diketone structure."),
                ("Demethoxycurcumin (DMC):", "Secondary fraction (17%); possesses single methoxy group, higher metabolic stability than curcumin."),
                ("Bisdemethoxycurcumin (BDMC):", "Tertiary fraction (3%); lacks methoxy substitutions, exhibits superior chemical stability in alkaline physiological pH.")
            ]),
            ("NF-kappaB & IKK DIRECT INHIBITION", "Shutting Down Master Inflammatory Hub", self.COLOR_CRIMSON, [
                ("Direct IKK Blockade:", "Curcumin directly binds and inhibits IkappaBalpha Kinase (IKKbeta), preventing the phosphorylation and degradation of IkappaBalpha."),
                ("Sequestration in Cytoplasm:", "Without IkappaB degradation, NF-kappaB heterodimer (p50/p65) remains trapped in an inactive state in the cytoplasm."),
                ("Downstream Transcriptional Silence:", "Shuts down transcription of COX-2, 5-LOX, iNOS, TNF-alpha, IL-1beta, IL-6, and Matrix Metalloproteinases (MMP-9)."),
                ("Anti-Catabolic Effect:", "Prevents inflammatory chondrocyte apoptosis and cartilage breakdown in osteoarthritic synovial joints.")
            ]),
            ("DIRECT CYCLOOXYGENASE-2 & 5-LOX BLOCK", "Dual Arachidonic Acid Cascade Inhibition", self.COLOR_EMERALD, [
                ("Dual Eicosanoid Inhibitor:", "Unlike pharmaceutical NSAIDs that only inhibit COX, curcumin simultaneously blocks both Cyclooxygenase (COX-2) and 5-Lipoxygenase (5-LOX)."),
                ("No Gastric Ulceration:", "Does not inhibit gastroprotective COX-1; causes zero gastric mucosal ulceration or renal hemodynamic impairment."),
                ("Free Radical Quenching:", "Its phenolic and beta-diketone functional groups donate electrons to scavenge peroxyl and hydroxyl radicals directly."),
                ("Metal Chelation:", "Chelates pro-oxidant transition metals ($Fe^{2+}$, $Cu^{2+}$), suppressing Fenton-mediated lipid peroxidation.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_17_piperine_surge(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 17)
        self._add_header_and_footer(
            slide,
            tag="PHARMACOKINETIC SYNERGY",
            title="The Curcumin Bioavailability Solution: The 2000% Piperine Clinical Surge",
            takeaway="Co-administration of 20 mg black pepper Piperine inhibits hepatic glucuronidation, increasing human serum curcumin AUC by +2000%.",
            citation="Clinical Pharmacokinetics: Shoba, G., et al., Planta Med., 1998; Anand, P., et al., Mol. Pharmaceutics.",
            slide_num=17
        )

        top_y = Inches(1.95)
        h_card = Inches(4.75)
        w_card = Inches(3.75)

        ctf1 = self.create_card(slide, Inches(0.8), top_y, w_card, h_card, "THE BIOAVAILABILITY BARRIER", "Why Pure Curcumin Fails Alone", self.COLOR_CRIMSON)
        self.append_bullet(ctf1, "1. Extreme Hydrophobicity:", "Water solubility of pure crystalline curcumin is virtually zero (<0.1 mg/mL); forms unabsorbable aggregates in watery gut chyme.", font_size=9.2)
        self.append_bullet(ctf1, "2. Massive First-Pass Metabolism:", "Intestinal and hepatic Phase II enzymes rapidly attach glucuronic acid and sulfate molecules to curcumin's hydroxyl groups.", font_size=9.2)
        self.append_bullet(ctf1, "3. UGT1A1 / UGT1A8 Clearance:", "UDP-glucuronosyltransferases convert active free curcumin into inert curcumin-glucuronide in under 30 minutes.", font_size=9.2)
        self.append_bullet(ctf1, "4. P-Glycoprotein Efflux:", "Active P-gp membrane pumps pump absorbed curcumin back out of enterocytes into the intestinal lumen for fecal excretion.", font_size=9.2)

        ctf2 = self.create_card(slide, Inches(4.8), top_y, w_card, h_card, "THE PIPERINE BIO-ENHANCER", "Mechanism of the +2000% Clinical Surge", self.COLOR_AMBER)
        self.append_bullet(ctf2, "Piperine Action:", "Alkaloid isolated from black pepper (Piper nigrum). Acts as a potent non-competitive inhibitor of hepatic and intestinal UGT enzymes.", font_size=9.2)
        self.append_bullet(ctf2, "P-gp Efflux Blockade:", "Transiently paralyzes enterocyte P-glycoprotein efflux pumps, allowing free curcumin molecules to enter mesenteric circulation.", font_size=9.2)
        self.append_bullet(ctf2, "The Landmark Trial:", "Co-administration of 20 mg piperine with 2,000 mg curcumin in healthy human volunteers increased serum bioavailability by 2,000% (20-fold AUC surge).", font_size=9.2)
        self.append_bullet(ctf2, "Half-Life Extension:", "Prolonged elimination half-life from undetectable to over 8 hours of sustained therapeutic systemic blood levels.", font_size=9.2)

        # Right Infographic Visual Panel
        self.add_image_panel(slide, Inches(8.8), top_y, Inches(3.733), h_card, "flowchart_curcumin_piperine.png", "Figure 4: Curcumin + Piperine Pharmacokinetics")

    def build_slide_18_gingerols_motility(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 18)
        self._add_header_and_footer(
            slide,
            tag="GASTROINTESTINAL NEUROBIOLOGY",
            title="Gingerols & Shogaols: Prokinetic Actions & 5-HT3 Receptor Antagonism",
            takeaway="Ginger bioactives antagonize serotonin 5-HT3 receptors in the gut and brainstem, speeding gastric emptying and blocking nausea pathways.",
            citation="Neurogastroenterology: Hu, M. L., et al., Am. J. Physiol. Gastrointest. Liver Physiol.; Walstab, J., et al., Neurogastroenterol. Motil.",
            slide_num=18
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        top_y = Inches(1.95)
        start_x = Inches(0.8)

        cards = [
            ("SEROTONIN 5-HT3 ANTAGONISM", "Central & Peripheral Antiemetic Action", self.COLOR_TEAL, [
                ("5-HT3 Receptor Target:", "6-Gingerol and 6-Shogaol act as competitive antagonists of the ionotropic Serotonin 5-HT3 receptor."),
                ("Vagal Afferent Nerve Block:", "Blocks serotonin released by enterochromaffin cells from exciting vagal afferent nerve endings in the duodenal mucosa."),
                ("Area Postrema / CTZ Blockade:", "Crosses into the brainstem to suppress 5-HT3 signaling in the Chemoreceptor Trigger Zone (CTZ) and nucleus tractus solitarius."),
                ("Clinical Applications:", "Clinically proven to reduce chemotherapy-induced nausea and vomiting (CINV), motion sickness, and morning sickness of pregnancy.")
            ]),
            ("GASTRIC EMPTYING & PROKINETIC VELOCITY", "Cholinergic Receptor Stimulation", self.COLOR_EMERALD, [
                ("Antral Contraction Velocity:", "Stimulates muscarinic $M_3$ acetylcholine receptors on gastric antral smooth muscle cells, increasing contraction force."),
                ("Pyloric Sphincter Relaxation:", "Accelerates trans-pyloric transit of chyme into the duodenum, reducing gastric retention time by up to 35%."),
                ("Functional Dyspepsia Relief:", "Double-blind clinical trials demonstrate significant improvement in early satiety, postprandial fullness, and epigastric pain."),
                ("Thermal Shift to Shogaols:", "Drying or heating fresh ginger dehydrates 6-Gingerol into 6-Shogaol, which exhibits double the antiemetic and analgesic potency!")
            ]),
            ("TRPV1 DESENSITIZATION & PAIN MODULATION", "Analgesic Actions on Sensory Neurons", self.COLOR_AMBER, [
                ("Capsaicin Receptor Binding:", "Gingerols bind Transient Receptor Potential Vanilloid 1 (TRPV1) channels on peripheral nociceptive C-fibers."),
                ("Channel Desensitization:", "Initial activation is followed by profound functional desensitization and intracellular calcium exhaustion in pain fibers."),
                ("Substance P Depletion:", "Depletes substance P and Calcitonin Gene-Related Peptide (CGRP) from primary sensory nerve terminals."),
                ("Clinical Efficacy in Migraine & Dysmenorrhea:", "Equal in clinical efficacy to 250 mg Mefenamic acid or Sumatriptan in multiple randomized clinical pain trials.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_19_allium_cascade(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 19)
        self._add_header_and_footer(
            slide,
            tag="ALLIUM CASCADE",
            title="Allium Organosulfur Chemistry: The Allicin Synthesis & Hydrogen Sulfide Signaling",
            takeaway="Mechanical crushing of garlic activates alliinase, producing volatile allicin which decomposes into cardioprotective polysulfides generating H2S gas.",
            citation="Cardiovascular Pharmacology: Benavides, G. A., et al., PNAS; Ried, K., et al., Eur. J. Clin. Nutr.",
            slide_num=19
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        top_y = Inches(1.95)
        start_x = Inches(0.8)

        cards = [
            ("ALLIIN TO ALLICIN ENZYMATIC SYNTHESIS", "Compartmentalized Defense Chemistry", self.COLOR_EMERALD, [
                ("Intact Clove State:", "Alliin (S-allyl-L-cysteine sulfoxide) resides in the mesophyll cell cytoplasm; Alliinase enzyme is sequestered in vacuoles."),
                ("Mechanical Cleavage:", "Crushing, mincing, or chewing ruptures compartmental barriers; pyridoxal phosphate (PLP)-dependent alliinase hydrolyzes alliin."),
                ("Intermediate Condensation:", "Alliinase produces allylsulfenic acid, which spontaneously condenses into diallyl thiosulfinate (Allicin) in seconds."),
                ("The 10-Minute Resting Rule:", "Culinary mandate: Garlic must rest at room temperature for 10 minutes post-crushing to allow complete allicin formation before heat exposure!")
            ]),
            ("DECOMPOSITION TO POLYSULFIDES & AJOENE", "Oil-Soluble & Water-Soluble Transformation", self.COLOR_AMBER, [
                ("Thermal Instability:", "Allicin is highly volatile and reactive; in cooking, it rapidly decomposes into lipophilic di- and trisulfides."),
                ("DADS and DATS:", "Yields Diallyl Disulfide (DADS) and Diallyl Trisulfide (DATS), powerful Phase II enzyme and Nrf2 inducers."),
                ("Ajoene Condensation:", "In oil or alcohol media, allicin rearranges into E- and Z-Ajoene, a potent inhibitor of platelet aggregation (anti-thrombotic)."),
                ("Aged Garlic Extract (SAC):", "Long-term aqueous extraction yields stable, non-irritating S-Allylcysteine (SAC), the standardized bioactive for blood pressure trials.")
            ]),
            ("HYDROGEN SULFIDE ($H_2S$) & VASODILATION", "Endogenous Gaseous Signaling Molecule", self.COLOR_COBALT, [
                ("Erythrocyte Interaction:", "Polysulfides (DATS/DADS) react with intracellular reduced glutathione (GSH) in red blood cells, liberating Hydrogen Sulfide ($H_2S$) gas."),
                ("$K_{ATP}$ Channel Opening:", "$H_2S$ gas activates ATP-sensitive potassium channels ($K_{ATP}$) on vascular smooth muscle, causing membrane hyperpolarization and vasodilation."),
                ("HMG-CoA Reductase Inhibition:", "Organosulfur metabolites competitively inhibit hepatic HMG-CoA reductase, downregulating endogenous cholesterol synthesis."),
                ("Clinical Meta-Analyses:", "Standardized aged garlic extract (600-1,200 mg daily) reliably reduces systolic blood pressure by 8-10 mmHg and diastolic by 5-6 mmHg in hypertensive cohorts.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_20_lamiaceae_antioxidants(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 20)
        self._add_header_and_footer(
            slide,
            tag="CULINARY HERBS",
            title="Lamiaceae Herbs: Rosmarinic Acid, Carnosic Acid & Thymoquinone",
            takeaway="Culinary mint-family herbs contain potent lipophilic diterpenes and caffeic acid dimers that arrest lipid peroxidation and inhibit neuroinflammation.",
            citation="Phytochemistry Reviews: Ngo, Y. T. H., et al., J. Agric. Food Chem.; Darakhshan, S., et al., Pharmacol. Res.",
            slide_num=20
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        top_y = Inches(1.95)
        start_x = Inches(0.8)

        cards = [
            ("ROSMARINIC ACID (ROSEMARY & OREGANO)", "Caffeic Acid Dimer Antioxidant", self.COLOR_EMERALD, [
                ("Chemical Nature:", "Ester of caffeic acid and 3,4-dihydroxyphenyllactic acid; highly bioavailable polyphenol."),
                ("Complement System Inhibition:", "Directly inhibits the C3b convertase enzyme, attenuating excessive complement activation in inflammatory tissue injuries."),
                ("Membrane Lipid Shield:", "Inserts into lipid bilayers, scavenging peroxyl radicals ($LOO^\bullet$) to stop self-propagating lipid peroxidation chains."),
                ("Topical & Systemic Antiallergic:", "Suppresses IgE elevation and protects lung epithelial tissue against diesel exhaust particle toxicity.")
            ]),
            ("CARNOSIC ACID & CARNOSOL (ROSEMARY)", "Electrophilic Diterpene Defense", self.COLOR_AMBER, [
                ("Pro-Drug Antioxidant:", "Carnosic acid is an abietane diterpene that acts as a 'pro-electrophilic drug' activated specifically at sites of inflammation."),
                ("Conversion to Carnosol:", "Upon reacting with reactive oxygen species, carnosic acid oxidizes into carnosol, quenching oxidants without becoming toxic itself."),
                ("Nrf2 & Neuroprotection:", "Passes the blood-brain barrier; upregulates neuroprotective enzymes in dopaminergic and cholinergic neurons."),
                ("Food Preservation Power:", "Natural carnosic acid extracts are utilized globally as natural antioxidants to prevent lipid rancidity in cooking oils and meat.")
            ]),
            ("THYMOQUINONE (NIGELLA SATIVA)", "Black Cumin Quinone Bioactive", self.COLOR_PURPLE, [
                ("Structure:", "Bioactive monoterpene quinone isolated from black seed oil (Nigella sativa)."),
                ("PPAR-gamma & Antioxidant:", "Strong agonist of PPAR-gamma; upregulates Superoxide Dismutase (SOD), Catalase (CAT), and Glutathione (GSH)."),
                ("Anti-Asthmatic Bronchodilation:", "Relaxes tracheal smooth muscle via non-selective block of histamine $H_1$ receptors and calcium channels."),
                ("Clinical Metabolic Efficacy:", "Significantly reduces fasting blood glucose, HbA1c, total cholesterol, and high-sensitivity C-reactive protein (hs-CRP) in metabolic syndrome.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_21_synergies(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 21)
        self._add_header_and_footer(
            slide,
            tag="FOOD MATRIX SYNERGIES",
            title="Whole-Food Matrix Synergies: Multi-Component Bioavailability Enhancement",
            takeaway="Whole foods contain natural biochemical cofactors that exponentially improve absorption, enzymatic reactivation, and chemical stability.",
            citation="Nutritional Synergy: Liu, R. H., Am. J. Clin. Nutr.; Jacobs, D. R., et al., Public Health Nutr.",
            slide_num=21
        )

        w = Inches(2.78)
        h = Inches(4.75)
        gap = Inches(0.2)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        synergies = [
            ("LIPID MICELLIZATION", "Fat + Carotenoids & Curcumin", self.COLOR_AMBER, [
                ("Bioavailability Bottleneck:", "Lipophilic bioactives (Curcumin, Lycopene, Carotenes) have negligible aqueous solubility (<0.1 mg/mL)."),
                ("The Synergy Pairing:", "Consuming tomatoes or turmeric with healthy lipids (Extra Virgin Olive Oil, Avocado, Ghee)."),
                ("Physiological Mechanism:", "Lipids stimulate gallbladder release of bile salts and pancreatic lipase, incorporating bioactives into mixed micelles for 4-to-6x intestinal enterocyte absorption."),
                ("Clinical Evidence:", "Cooking tomatoes in olive oil increases trans-to-cis lycopene isomerization and human serum lycopene AUC by over 400%!")
            ]),
            ("THERMAL ENZYME RESCUE", "Cooked Broccoli + Mustard Seed", self.COLOR_EMERALD, [
                ("Bioavailability Bottleneck:", "Boiling or steaming broccoli for >5 minutes thermally inactivates endogenous myrosinase, preventing glucoraphanin conversion to sulforaphane."),
                ("The Synergy Pairing:", "Sprinkling raw ground brown mustard seed, daikon radish, or wasabi onto cooked broccoli prior to consumption."),
                ("Physiological Mechanism:", "Raw Cruciferous seeds provide heat-stable exogenous plant myrosinase, immediately hydrolyzing intact intact glucoraphanin into active Sulforaphane."),
                ("Clinical Evidence:", "Human urinary metabolite excretion studies show a complete 100% restoration of sulforaphane yield in cooked vegetables.")
            ]),
            ("REDOX STABILIZATION", "Green Tea + Citrus Ascorbic Acid", self.COLOR_TEAL, [
                ("Bioavailability Bottleneck:", "Over 80% of green tea catechins (EGCG, EGC, ECG) degrade in the alkaline environment ($pH > 7.0$) of the human small intestine."),
                ("The Synergy Pairing:", "Infusing fresh lemon, lime, or citrus juice directly into freshly brewed green tea."),
                ("Physiological Mechanism:", "Citric acid and Ascorbic Acid lower luminal pH and donate electrons, maintaining green tea catechins in their stable reduced bioactive form."),
                ("Clinical Evidence:", "In vitro and human digestive models demonstrate a 6-fold increase in recovered bioaccessible EGCG across the intestinal barrier.")
            ]),
            ("MINERAL CHELATION", "Non-Heme Iron + Vitamin C", self.COLOR_CRIMSON, [
                ("Bioavailability Bottleneck:", "Plant non-heme iron ($Fe^{3+}$) forms insoluble ferric hydroxide complexes that cannot cross enterocyte membranes."),
                ("The Synergy Pairing:", "Pairing dark leafy greens (Spinach, Kale) or legumes with raw bell peppers or lemon juice."),
                ("Physiological Mechanism:", "Ascorbic acid reduces insoluble ferric iron ($Fe^{3+}$) to soluble ferrous iron ($Fe^{2+}$), which directly binds DMT1 enterocyte transporters."),
                ("Clinical Evidence:", "Adding 100 mg Vitamin C to plant-based meals increases fractional iron absorption by over 300%, overcoming inhibitory dietary phytates.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(synergies):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.0, space_after=4.0)

    def build_slide_22_clinical_dosages(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 22)
        self._add_header_and_footer(
            slide,
            tag="CLINICAL TRANSLATION",
            title="Evidence-Based Human Clinical Dosage & Therapeutic Endpoint Matrix",
            takeaway="Translating in vitro mechanisms into clinical practice requires validated daily therapeutic human dosages verified in randomized clinical trials.",
            citation="Clinical Trials: Cochrane Database of Systematic Reviews; Natural Medicines Comprehensive Database.",
            slide_num=22
        )

        headers = ["Phytochemical Compound", "Botanical Source", "Human Daily Dose (Clinical)", "Primary Clinical Endpoint", "Human Trial Validation"]
        col_widths = [Inches(2.5), Inches(2.2), Inches(2.4), Inches(2.6), Inches(2.033)]
        rows_data = [
            ["Sulforaphane (SFN)", "Broccoli Seed Extracts / Sprouts", "30 - 60 mg SFN (or 200 mg Glucoraphanin)", "Phase II enzyme induction, GST elevation, reduced air pollutant benzene adducts", "Randomized Controlled Trial (RCT)"],
            ["Curcumin + Piperine", "Curcuma longa + Piper nigrum", "1,000 - 2,000 mg + 20 mg Piperine", "Pain score reduction in knee osteoarthritis (WOMAC), downregulates hs-CRP", "Meta-Analysis (29 RCTs)"],
            ["Urolithin A", "Pomegranate extract / fermentation", "500 - 1,000 mg pure Urolithin A", "Mitochondrial clearance, +12% muscle endurance, biomarkers of mitophagy", "Nature Metabolism RCT"],
            ["Hesperidin / Citrus PMFs", "Orange flavedo / peel extracts", "500 mg Hesperidin daily", "Endothelial flow-mediated dilation (FMD +2.2%), diastolic BP reduction", "Double-Blind RCT"],
            ["Anthocyanins (Cyanidin)", "Wild Blueberries / Tart Cherries", "250 - 500 mg standardized", "Cognitive processing speed, reduction of serum uric acid, reduced DOMS", "Prospective Cohort & RCT"],
            ["Gingerols & Shogaols", "Zingiber officinale dried rhizome", "1,000 - 1,500 mg standardized", "Antiemetic efficacy in CINV & pregnancy, acceleration of gastric emptying", "Cochrane Systematic Review"],
            ["Aged Garlic Extract (SAC)", "Allium sativum fermented", "600 - 1,200 mg (2.4 mg S-Allylcysteine)", "Reduction of central systolic blood pressure (-8.8 mmHg), arterial stiffness", "Hypertension Clinical Trials"]
        ]

        self.create_table(slide, Inches(0.8), Inches(1.95), Inches(11.733), Inches(4.75), headers, col_widths, rows_data)

    def build_slide_23_culinary_pharmacokinetics(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 23)
        self._add_header_and_footer(
            slide,
            tag="CULINARY PHARMACOKINETICS",
            title="Thermal Processing Kinetics & Domestic Preparation Guidelines",
            takeaway="Cooking methods dramatically alter bioactive concentrations: Steaming preserves glucosinolates; boiling leaches polyphenols; stewing unlocks carotenoids.",
            citation="Food Science Reference: Palermo, M., et al., J. Sci. Food Agric.; Miglio, C., et al., J. Agric. Food Chem.",
            slide_num=23
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        top_y = Inches(1.95)
        start_x = Inches(0.8)

        cards = [
            ("STEAMING VS. BOILING VS. MICROWAVING", "Retention of Water-Soluble Bioactives", self.COLOR_EMERALD, [
                ("Boiling Leaching Hazard:", "Boiling leafy vegetables causes 50 - 75% leaching of water-soluble polyphenols, anthocyanins, and glucosinolates into discarded cooking water."),
                ("Steaming Supremacy:", "Steaming for 3 - 5 minutes preserves 85 - 95% of total phenolic content and protects myrosinase from rapid thermal denaturation."),
                ("Microwaving with Minimal Water:", "Short-duration microwaving (<2 min) with minimal water retains high bioactive levels; high-power prolonged cooking degrades flavonoid rings."),
                ("Sous-Vide Precision:", "Vacuum packaging at low temperature (55-60°C) allows complete retention of volatile organosulfur compounds and phenolic antioxidants.")
            ]),
            ("OIL MATRIX & THERMAL DISSOLUTION", "Unlocking Lipophilic Carotenoids", self.COLOR_AMBER, [
                ("Cell Wall Breakdown:", "Plant carotenoids are trapped within crystalline chromoplast protein complexes; gentle thermal cooking ruptures pectin and cellulose walls."),
                ("Trans-to-Cis Isomerization:", "Heat promotes isomerization of all-trans-lycopene into cis-isomers, which exhibit dramatically higher human intestinal micellization."),
                ("EVOO Phenolic Cross-Partitioning:", "Cooking in Extra Virgin Olive Oil creates a bi-directional transfer: oil polyphenols enrich the vegetables while fat-soluble carotenoids dissolve into the oil."),
                ("Culinary Rule:", "Cook tomatoes, carrots, and pumpkin in healthy fats; eat berries and citrus raw or freeze-dried.")
            ]),
            ("CHOPPING, RESTING & PRE-TREATMENT RULES", "Enzyme Activation Protocols", self.COLOR_COBALT, [
                ("The Brassica 40-Minute Rule:", "Chop or shred cruciferous vegetables (broccoli, cabbage) and allow to sit raw at room temperature for 30 - 40 minutes before cooking to maximize SFN formation!"),
                ("The Allium 10-Minute Rest:", "Crush or mince fresh garlic and let rest on the cutting board for 10 minutes before heating; preserves alliinase activity from immediate heat destruction."),
                ("Freezing Rupture Benefit:", "Freezing fresh berries forms intracellular ice crystals that shear cell walls, yielding higher extractable anthocyanin bioavailability in smoothies."),
                ("Whole Citrus Zest Integration:", "Grate organic citrus flavedo zest into dishes; the outer peel contains 10-fold higher polymethoxyflavone levels than the inner juice.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_24_multi_target_summary(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 24)
        self._add_header_and_footer(
            slide,
            tag="MULTI-TARGET PHARMACOLOGY",
            title="Comprehensive Multi-Target Molecular Mechanism Summary Matrix",
            takeaway="Phytochemicals exert synergistic therapeutic benefits by acting simultaneously across multiple molecular and epigenetic pathways.",
            citation="Systems Biology: Surh, Y. J., Nat. Rev. Cancer; Aggarwal, B. B., et al., Biochem. Pharmacol.",
            slide_num=24
        )

        w = Inches(2.2)
        h = Inches(4.75)
        gap = Inches(0.18)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        systems = [
            ("1. ANTIOXIDANT / NRF2", "Cellular Redox Defense", self.COLOR_EMERALD, [
                ("Key Bioactives:", "Sulforaphane, Curcumin, Resveratrol, Carnosic Acid."),
                ("Molecular Mechanism:", "Modifies Keap1 C151, freeing Nrf2 to enter nucleus and bind Antioxidant Response Elements (ARE)."),
                ("Target Enzymes:", "HO-1, NQO1, Glutamate-Cysteine Ligase (GCLC), Glutathione Peroxidase."),
                ("Clinical Outcome:", "Direct and indirect quenching of oxidative stress across cardiac, renal, and neural cells.")
            ]),
            ("2. ANTI-INFLAMMATORY", "NF-kappaB & Eicosanoids", self.COLOR_CRIMSON, [
                ("Key Bioactives:", "Curcumin, Gingerols, Quercetin, Bromelain, Allicin."),
                ("Molecular Mechanism:", "Directly inhibits IKKbeta kinase, preventing phosphorylation of IkappaB and trapping NF-kappaB in cytoplasm."),
                ("Target Enzymes:", "COX-2, 5-LOX, iNOS, TNF-alpha, IL-1beta, IL-6, and MMP-9."),
                ("Clinical Outcome:", "Resolves chronic low-grade inflammation in osteoarthritis, metabolic syndrome, and vascular disease.")
            ]),
            ("3. EPIGENETIC TUNING", "HDAC & DNMT Remodeling", self.COLOR_PURPLE, [
                ("Key Bioactives:", "Sulforaphane, I3C, 3,3'-DIM, Epigallocatechin gallate (EGCG)."),
                ("Molecular Mechanism:", "Chelates catalytic zinc in Class I/II HDACs; suppresses DNMT1/3b promoter hypermethylation."),
                ("Target Genes:", "Re-activates p21/Waf1, p16/INK4a, and silenced Nrf2 promoter CpG islands."),
                ("Clinical Outcome:", "Restores cellular differentiation, cell cycle arrest, and prevents age-related epigenetic drift.")
            ]),
            ("4. METABOLIC ENERGETICS", "AMPK & Sirtuin Activation", self.COLOR_TEAL, [
                ("Key Bioactives:", "Mangiferin, Resveratrol, Nobiletin, Phloridzin."),
                ("Molecular Mechanism:", "Phosphorylates AMPK Thr172; allosterically activates NAD+-dependent deacetylase SIRT1."),
                ("Target Effectors:", "Stimulates GLUT4 translocation, downregulates PEPCK/G6Pase, induces PGC-1alpha biogenesis."),
                ("Clinical Outcome:", "Normalizes fasting blood glucose, enhances lipid clearance, and boosts skeletal muscle endurance.")
            ]),
            ("5. VASCULAR & MITOPHAGY", "eNOS & Mitochondrial Health", self.COLOR_COBALT, [
                ("Key Bioactives:", "Hesperidin, Urolithin A, Anthocyanins, Garlic Polysulfides."),
                ("Molecular Mechanism:", "Phosphorylates eNOS Ser1177; triggers PINK1-Parkin autophagosomal clearance of old mitochondria."),
                ("Target Effectors:", "Elevates Nitric Oxide ($NO^\bullet$) and cGMP; selectively degrades depolarized mitochondria."),
                ("Clinical Outcome:", "Reduces systemic vascular resistance (MAP -6 mmHg) and restores muscular endurance in aging.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(systems):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=8.8, space_after=4.0)

    def build_slide_25_bibliography(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 25)
        self._add_header_and_footer(
            slide,
            tag="SCIENTIFIC LEDGER",
            title="Authoritative Academic Bibliography & Primary Literature Citations",
            takeaway="Curated ledger of peer-reviewed seminal studies establishing the pharmacological and clinical evidence base for botanical bioactives.",
            citation="Academic Citations verified via Europe PMC, PubMed, and Crossref APIs.",
            slide_num=25
        )

        w = Inches(5.7)
        h = Inches(4.75)
        top_y = Inches(1.95)

        ctf1 = self.create_card(slide, Inches(0.8), top_y, w, h, "PRIMARY PAPERS (PART I: CELLULAR & ENZYMATIC)", "Peer-Reviewed References", self.COLOR_HEADER_GOLD)
        self.append_bullet(ctf1, "1. Fahey, J. W., et al. (1997):", "'Broccoli sprouts: An exceptionally rich source of inducers of enzymes that protect against chemical carcinogens.' PNAS, 94(19): 10367-10372. DOI: 10.1073/pnas.94.19.10367.", font_size=8.5, space_after=3.0)
        self.append_bullet(ctf1, "2. Dinkova-Kostova, A. T., et al. (2017):", "'The role of Keap1-Nrf2 signaling in physiological homeostasis and chemical carcinogenesis.' Nat. Rev. Drug Discov., 16: 671-692. DOI: 10.1038/nrd.2017.135.", font_size=8.5, space_after=3.0)
        self.append_bullet(ctf1, "3. Dashwood, R. H., & Ho, E. (2007):", "'Dietary histone deacetylase inhibitors: From bench to clinical application.' Cancer Lett., 250(1): 8-16. DOI: 10.1016/j.canlet.2006.09.020.", font_size=8.5, space_after=3.0)
        self.append_bullet(ctf1, "4. Bradlow, H. L., et al. (1996):", "'Multifunctional aspects of the actions of indole-3-carbinol and its derivatives.' Ann. N.Y. Acad. Sci., 768: 180-200. DOI: 10.1111/j.1749-6632.1995.tb12120.x.", font_size=8.5, space_after=3.0)
        self.append_bullet(ctf1, "5. Ryu, D., et al. (2016):", "'Urolithin A induces mitophagy and prolongs lifespan in C. elegans and increases muscle function in rodents.' Nat. Med., 22(8): 879-888. DOI: 10.1038/nm.4132.", font_size=8.5, space_after=3.0)
        self.append_bullet(ctf1, "6. Andreux, P. A., et al. (2019):", "'The mitophagy activator urolithin A is safe and induces a molecular signature of improved mitochondrial and cellular health in humans.' Nat. Metab., 1: 595-603. DOI: 10.1038/s42255-019-0073-4.", font_size=8.5, space_after=3.0)

        ctf2 = self.create_card(slide, Inches(6.8), top_y, w, h, "PRIMARY PAPERS (PART II: CLINICAL PHARMACOLOGY)", "Human Trial Validations", self.COLOR_EMERALD)
        self.append_bullet(ctf2, "7. Shoba, G., et al. (1998):", "'Influence of piperine on the pharmacokinetics of curcumin in animals and human volunteers.' Planta Med., 64(4): 353-356. DOI: 10.1055/s-2006-957450.", font_size=8.5, space_after=3.0)
        self.append_bullet(ctf2, "8. Aggarwal, B. B., et al. (2006):", "'Curcumin: the Indian solid gold.' Adv. Exp. Med. Biol., 595: 1-75. DOI: 10.1007/978-0-387-46401-5_1.", font_size=8.5, space_after=3.0)
        self.append_bullet(ctf2, "9. Benavides, G. A., et al. (2007):", "'Hydrogen sulfide mediates the vasoactivity of garlic.' PNAS, 104(46): 17977-17982. DOI: 10.1073/pnas.0705710104.", font_size=8.5, space_after=3.0)
        self.append_bullet(ctf2, "10. Devore, E. E., et al. (2012):", "'Dietary intakes of berries and flavonoids in relation to cognitive decline.' Ann. Neurol., 72(1): 135-143. DOI: 10.1002/ana.23594.", font_size=8.5, space_after=3.0)
        self.append_bullet(ctf2, "11. Rizza, S., et al. (2011):", "'Oral administration of hesperidin improves endothelial dysfunction in subjects with metabolic syndrome.' J. Clin. Endocrinol. Metab., 96(5): E782-E792. DOI: 10.1210/jc.2010-2334.", font_size=8.5, space_after=3.0)
        self.append_bullet(ctf2, "12. Hu, M. L., et al. (2011):", "'Effect of ginger on gastric motility and symptoms of functional dyspepsia.' World J. Gastroenterol., 17(1): 105-110. DOI: 10.3748/wjg.v17.i1.105.", font_size=8.5, space_after=3.0)

    # -------------------------------------------------------------------------
    # COMPILATION
    # -------------------------------------------------------------------------
    def build_all(self, output_pptx_path: Path):
        print(f"Compiling 25-Slide Phytochemical Master Deck...")
        builders = [
            self.build_slide_01_cover,
            self.build_slide_02_taxonomy,
            self.build_slide_03_brassicaceae_cascade,
            self.build_slide_04_nrf2_mechanism,
            self.build_slide_05_epigenetics_hdac,
            self.build_slide_06_i3c_dim,
            self.build_slide_07_tropical_fruits,
            self.build_slide_08_bromelain_flowchart,
            self.build_slide_09_mangiferin_metabolism,
            self.build_slide_10_citrus_pmfs,
            self.build_slide_11_enos_activation,
            self.build_slide_12_punicalagin_urolithin,
            self.build_slide_13_temperate_fruits,
            self.build_slide_14_quercetin_mast_cells,
            self.build_slide_15_berry_anthocyanins_bbb,
            self.build_slide_16_curcuminoids_pharmacology,
            self.build_slide_17_piperine_surge,
            self.build_slide_18_gingerols_motility,
            self.build_slide_19_allium_cascade,
            self.build_slide_20_lamiaceae_antioxidants,
            self.build_slide_21_synergies,
            self.build_slide_22_clinical_dosages,
            self.build_slide_23_culinary_pharmacokinetics,
            self.build_slide_24_multi_target_summary,
            self.build_slide_25_bibliography
        ]

        for idx, builder_fn in enumerate(builders, start=1):
            print(f"  -> Building Slide {idx:02d}/25: {builder_fn.__name__}...")
            builder_fn()

        output_pptx_path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(output_pptx_path))
        print(f"[SUCCESS] 25-Slide presentation saved to: {output_pptx_path}")


def main():
    samples_dir = REPO_ROOT / "samples"
    pptx_path = samples_dir / "phytochemicals_25slides_master_deck.pptx"

    print("=================================================================")
    print("PHYTOCHEMICALS 25-SLIDE MASTER KEYNOTE DECK GENERATOR")
    print("=================================================================")
    builder = Phytochemicals25DeckBuilder()
    builder.build_all(pptx_path)


if __name__ == "__main__":
    main()
