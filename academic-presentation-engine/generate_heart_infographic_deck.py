"""
Human Heart Organ Infographic Presentation Generator - Royal Oxford Medical Edition
Comprehensive 12-Slide High-Yield Medical & Physiological Infographic Deck
Covers: Anatomical Architecture, 4 Chambers, Valves, Dual Circulation, Electrical Conduction,
Cardiac Cycle, Systemic Organ Perfusion (Brain, Kidneys, Liver, Muscle), Neuroendocrine (ANP/BNP),
Coronary Metabolism, Clinical Hemodynamics & Biomarkers, and Pathology Cascades.

Design Standards:
- Typography: 100% Times New Roman across all titles, cards, tables, and footers.
- Canvas: Luminous Slate-White (#F8FAFC) content area with Deep Oxford Navy (#0F1D36) Header Canopy.
- High-Contrast Elevated Cards: Native DrawingML vector containers with color-coded thematic strips.
- Table Layouts: Explicit column budgeting with middle vertical centering.
- 100% Compatible with Microsoft PowerPoint, Canva Pro, and Google Slides.
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

PROJECT_ROOT = Path(__file__).resolve().parent
REPO_ROOT = PROJECT_ROOT.parent
sys.path.insert(0, str(PROJECT_ROOT))

class HeartInfographicDeckBuilder:
    FONT_FAMILY = "Times New Roman"

    # Color Palette
    COLOR_CANVAS = RGBColor(248, 250, 252)         # Luminous Slate White #F8FAFC
    COLOR_HEADER_BG = RGBColor(15, 29, 54)         # Deep Royal Oxford Navy #0F1D36
    COLOR_HEADER_BAR = RGBColor(212, 160, 23)      # Metallic Champagne Gold #D4A017
    
    COLOR_CARD_WHITE = RGBColor(255, 255, 255)     # Elevated Pure White #FFFFFF
    COLOR_CARD_ALT = RGBColor(241, 245, 249)       # Slate Tint #F1F5F9
    COLOR_BORDER_SLATE = RGBColor(203, 213, 225)   # Border Outline #CBD5E1

    # Physiological Thematic Colors
    COLOR_ARTERIAL_RED = RGBColor(185, 28, 28)     # Oxygenated Arterial Red #B91C1C
    COLOR_VENOUS_BLUE = RGBColor(29, 78, 216)      # Deoxygenated Venous Blue #1D4ED8
    COLOR_ELECTRICAL_GOLD = RGBColor(217, 119, 6)  # Conduction Amber Gold #D97706
    COLOR_PERFUSION_TEAL = RGBColor(13, 148, 136)  # Organ Perfusion Teal #0D9488
    COLOR_NEURO_PURPLE = RGBColor(109, 40, 217)    # Autonomic Purple #6D28D9
    COLOR_METABOLIC_ROSE = RGBColor(225, 29, 72)   # Myocardial Rose #E11D48

    COLOR_TEXT_TITLE = RGBColor(255, 255, 255)
    COLOR_TEXT_SUBTITLE = RGBColor(186, 230, 253)
    COLOR_TEXT_DARK = RGBColor(15, 23, 42)         # Deep Charcoal (#0F172A)
    COLOR_TEXT_NAVY = RGBColor(15, 30, 60)         # Deep Navy (#0F1E3C)
    COLOR_TEXT_MUTED = RGBColor(71, 85, 105)
    COLOR_TEXT_FOOTER = RGBColor(203, 213, 225)
    COLOR_GOLD_BRIGHT = RGBColor(252, 211, 77)

    TOTAL_SLIDES = 12

    SLIDE_THEMES = {
        1:  {"name": "KEYNOTE INFOGRAPHIC ARCHITECTURE", "color": COLOR_ARTERIAL_RED},
        2:  {"name": "ANATOMICAL BLUEPRINT & 4 CHAMBERS", "color": COLOR_ARTERIAL_RED},
        3:  {"name": "VALVE MECHANICS & FLUID DYNAMICS", "color": COLOR_VENOUS_BLUE},
        4:  {"name": "DUAL CIRCULATION: PULMONARY & SYSTEMIC", "color": COLOR_VENOUS_BLUE},
        5:  {"name": "CARDIAC CONDUCTION & ECG VECTORS", "color": COLOR_ELECTRICAL_GOLD},
        6:  {"name": "CARDIAC CYCLE & HEMODYNAMICS", "color": COLOR_ELECTRICAL_GOLD},
        7:  {"name": "SYSTEMIC ORGAN PERFUSION INFOGRAPHIC", "color": COLOR_PERFUSION_TEAL},
        8:  {"name": "AUTONOMIC & ENDOCRINE REGULATION (ANP/BNP)", "color": COLOR_NEURO_PURPLE},
        9:  {"name": "CORONARY FLOW & METABOLIC BIOENERGETICS", "color": COLOR_METABOLIC_ROSE},
        10: {"name": "CLINICAL HEMODYNAMICS & BIOMARKERS", "color": COLOR_PERFUSION_TEAL},
        11: {"name": "PATHOLOGY CASCADES & ORGAN FAILURE", "color": COLOR_ARTERIAL_RED},
        12: {"name": "MASTER 5-PILLAR PHYSIOLOGY SYNTHESIS", "color": COLOR_HEADER_BAR}
    }

    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]

    def _add_canvas_background(self, slide, slide_num: int):
        theme = self.SLIDE_THEMES.get(slide_num, self.SLIDE_THEMES[1])

        # 1. Base Luminous Slate-White Canvas
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = self.COLOR_CANVAS
        bg.line.fill.background()

        # 2. Left Spine Architectural Rail
        spine = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.14), self.prs.slide_height)
        spine.fill.solid()
        spine.fill.fore_color.rgb = theme["color"]
        spine.line.fill.background()

        # 3. Solid Deep Royal Oxford Navy Architectural Header Zone
        top_zone = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, Inches(1.72))
        top_zone.fill.solid()
        top_zone.fill.fore_color.rgb = self.COLOR_HEADER_BG
        top_zone.line.fill.background()

        # 4. Primary Metallic Champagne Gold Divider Rule
        gold_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.72), self.prs.slide_width, Pt(2.5))
        gold_line.fill.solid()
        gold_line.fill.fore_color.rgb = self.COLOR_HEADER_BAR
        gold_line.line.fill.background()

        # 5. Secondary Theme Accent Hairline
        mod_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.72) + Pt(2.5), self.prs.slide_width, Pt(1.5))
        mod_line.fill.solid()
        mod_line.fill.fore_color.rgb = theme["color"]
        mod_line.line.fill.background()

        return bg

    def _add_header_and_footer(self, slide, tag: str, title: str, takeaway: str, citation: str, slide_num: int):
        theme = self.SLIDE_THEMES.get(slide_num, self.SLIDE_THEMES[1])

        # Top Category Badge Pill
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.20), Inches(4.5), Inches(0.32))
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

        # Right Theme Tracker Pill
        mod_pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(0.20), Inches(4.733), Inches(0.32))
        mod_pill.fill.solid()
        mod_pill.fill.fore_color.rgb = RGBColor(20, 36, 66)
        mod_pill.line.color.rgb = theme["color"]
        mod_pill.line.width = Pt(1.0)
        mtf = mod_pill.text_frame
        mtf.word_wrap = True
        mtf.margin_top = Pt(2)
        mp = mtf.paragraphs[0]
        mp.alignment = PP_ALIGN.CENTER
        mp.text = f"INFOGRAPHIC {slide_num:02d}/{self.TOTAL_SLIDES} • {theme['name'][:30]}"
        mp.font.name = self.FONT_FAMILY
        mp.font.size = Pt(9.0)
        mp.font.bold = True
        mp.font.color.rgb = self.COLOR_GOLD_BRIGHT

        # Slide Main Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(11.733), Inches(0.60))
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = self.FONT_FAMILY
        p.font.size = Pt(21)
        p.font.bold = True
        p.font.color.rgb = self.COLOR_TEXT_TITLE

        # Subtitle / Clinical Takeaway
        sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.15), Inches(11.733), Inches(0.48))
        stf = sub_box.text_frame
        stf.word_wrap = True
        stf.margin_left = stf.margin_top = stf.margin_right = stf.margin_bottom = 0
        sp = stf.paragraphs[0]
        sp.text = takeaway
        sp.font.name = self.FONT_FAMILY
        sp.font.size = Pt(11.5)
        sp.font.color.rgb = self.COLOR_TEXT_SUBTITLE

        # Grounded Footer Zone (Deep Oxford Navy Capsule)
        foot_zone = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.92), Inches(11.733), Inches(0.40))
        foot_zone.fill.solid()
        foot_zone.fill.fore_color.rgb = self.COLOR_HEADER_BG
        foot_zone.line.color.rgb = theme["color"]
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

        color = accent_color or self.COLOR_ARTERIAL_RED

        # Accent top rail
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
        p.font.size = Pt(13)
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

    def append_bullet(self, tf, bold_prefix: str, text: str, font_size: float = 10.0, space_after: float = 4.0):
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

    def create_table(self, slide, left, top, width, height, headers: list, col_widths: list, rows_data: list):
        rows_count = len(rows_data) + 1
        cols_count = len(headers)
        table_shape = slide.shapes.add_table(rows_count, cols_count, left, top, width, height)
        table = table_shape.table

        for i, w in enumerate(col_widths):
            table.columns[i].width = w

        # Header Row
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

        # Data Rows
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
                p.font.size = Pt(9.5)
                if col_idx == 0:
                    p.alignment = PP_ALIGN.LEFT
                    p.font.bold = True
                    p.font.color.rgb = self.COLOR_TEXT_NAVY
                else:
                    p.alignment = PP_ALIGN.LEFT
                    p.font.bold = False
                    p.font.color.rgb = self.COLOR_TEXT_DARK

        return table_shape

    # -------------------------------------------------------------
    # SLIDE BUILDERS (1 to 12)
    # -------------------------------------------------------------
    def build_slide_1_cover(self):
        slide = self.prs.slides.add_slide(self.blank_layout)

        # Full dark Oxford Navy title canvas
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = self.COLOR_HEADER_BG
        bg.line.fill.background()

        # Double Gold Accent Rules
        rule1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.1), Inches(11.733), Pt(2.5))
        rule1.fill.solid()
        rule1.fill.fore_color.rgb = self.COLOR_HEADER_BAR
        rule1.line.fill.background()

        # Category Badge
        cat = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.65), Inches(5.8), Inches(0.35))
        cat.fill.solid()
        cat.fill.fore_color.rgb = RGBColor(24, 43, 73)
        cat.line.color.rgb = self.COLOR_HEADER_BAR
        cat.line.width = Pt(1.0)
        ctf = cat.text_frame
        p = ctf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = "HUMAN CARDIOVASCULAR PHYSIOLOGY • MEDICAL INFOGRAPHIC"
        p.font.name = self.FONT_FAMILY
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = self.COLOR_GOLD_BRIGHT

        # Title
        tbox = slide.shapes.add_textbox(Inches(0.8), Inches(1.35), Inches(11.733), Inches(1.6))
        tf = tbox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = "THE HUMAN HEART: ARCHITECTURE, HEMODYNAMICS & SYSTEMIC ORGAN PERFUSION"
        p.font.name = self.FONT_FAMILY
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = self.COLOR_TEXT_TITLE

        p2 = tf.add_paragraph()
        p2.text = "A Complete Illustrated Infographic Blueprint: Anatomical Chambers, Valvular Mechanics, Electrical Vectors, Systemic Organ Delivery & Clinical Biomarkers"
        p2.font.name = self.FONT_FAMILY
        p2.font.size = Pt(13)
        p2.font.color.rgb = self.COLOR_TEXT_SUBTITLE

        # 4 Core Vital Metric Cards across bottom
        card_w = Inches(2.78)
        card_h = Inches(3.3)
        gap = Inches(0.2)
        start_x = Inches(0.8)
        top_y = Inches(3.2)

        pillars = [
            ("MECHANICAL PUMP", "7,500 L / Day Fluid Displacement", self.COLOR_ARTERIAL_RED, [
                ("Cardiac Output:", "4.5 - 5.5 L/min at rest; surges to 20-25 L/min during peak physical exertion."),
                ("Stroke Volume:", "~70 mL blood ejected per beat into the ascending aorta and pulmonary trunk."),
                ("Lifetime Workload:", "Beats ~100,000 times/day and over 2.5 billion times in an average lifespan."),
                ("Hemodynamic Role:", "Maintains the systemic arterial pressure gradient (MAP >65 mmHg) powering microvascular filtration.")
            ]),
            ("ELECTRICAL OSCILLATOR", "Autonomous Pacemaker Grid", self.COLOR_ELECTRICAL_GOLD, [
                ("SA Node Pacemaker:", "Spontaneous phase 4 diastolic depolarization at 60-100 actions potentials/min."),
                ("AV Nodal Gate:", "Imposes a critical 0.12-second delay allowing complete ventricular end-diastolic filling."),
                ("Fast Conduction:", "Purkinje fiber network deploys electrical vector at 2-4 m/s apex-to-base."),
                ("Syncytium:", "Gap junctions in intercalated discs allow simultaneous all-or-nothing syncytial contraction.")
            ]),
            ("ORGAN PERFUSION ENGINE", "Lifeline to All Vital Organs", self.COLOR_PERFUSION_TEAL, [
                ("Brain Delivery:", "15% of Cardiac Output (~750 mL/min); neurons suffer necrosis after 4-5 min arrest."),
                ("Renal Filtration:", "20-25% of CO (~1,200 mL/min) powering GFR and electrolyte homeostasis."),
                ("Hepato-Splanchnic:", "25% of CO nourishing liver detox, digestion, and nutrient absorption."),
                ("Metabolic Delivery:", "Delivers 1,000 mL O2/min and cleanses cellular metabolic byproducts.")
            ]),
            ("ENDOCRINE GLAND", "Natriuretic Peptide Regulation", self.COLOR_NEURO_PURPLE, [
                ("ANP Hormone:", "Secreted by atrial myocytes in response to atrial distension and hypervolemia."),
                ("BNP Hormone:", "Secreted by ventricular myocardium under wall tension; diagnostic gold standard for heart failure."),
                ("Target Action:", "Promotes renal natriuresis, diuresis, systemic vasodilation, and RAAS suppression."),
                ("Fluid Homeostasis:", "Acts as the intrinsic physiological brake preventing lethal circulatory overload.")
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

        # Bottom Rule
        b_rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.8), Inches(11.733), Pt(1.5))
        b_rule.fill.solid()
        b_rule.fill.fore_color.rgb = self.COLOR_HEADER_BAR
        b_rule.line.fill.background()

        f_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.35))
        ftf = f_box.text_frame
        p = ftf.paragraphs[0]
        p.text = "Authoritative Source Standards: Guyton & Hall Medical Physiology (14th Ed.) | Braunwald's Heart Disease | Harrison's Principles of Internal Medicine"
        p.font.name = self.FONT_FAMILY
        p.font.size = Pt(8.5)
        p.font.color.rgb = self.COLOR_TEXT_FOOTER

    def build_slide_2_chambers(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 2)
        self._add_header_and_footer(
            slide,
            tag="ANATOMICAL BLUEPRINT",
            title="Four-Chamber Architecture, Myocardial Walls & Pressure Gradients",
            takeaway="Asymmetrical muscular architecture matches differing hemodynamic workloads: LV generates 5x RV peak systolic pressure.",
            citation="Anatomical References: Netter's Atlas of Human Anatomy (8th Ed.) | Guyton & Hall Medical Physiology, Ch. 9.",
            slide_num=2
        )

        w = Inches(2.78)
        h = Inches(4.75)
        gap = Inches(0.2)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        chambers = [
            ("RIGHT ATRIUM (RA)", "Low-Pressure Venous Collector", self.COLOR_VENOUS_BLUE, [
                ("Inflow Sources:", "Receives deoxygenated systemic venous return from Superior Vena Cava (head/arms), Inferior Vena Cava (body/legs), and Coronary Sinus (myocardium)."),
                ("Normal Pressure:", "Central Venous Pressure (CVP) ranges 2 - 6 mmHg."),
                ("Internal Structures:", "Fossa ovalis (remnant of fetal foramen ovale), crista terminalis, and pectinate muscles lining the right auricle."),
                ("Electrophysiology:", "Houses Sinoatrial (SA) node in high posterior wall near SVC junction and Atrioventricular (AV) node near triangle of Koch.")
            ]),
            ("RIGHT VENTRICLE (RV)", "Low-Resistance Pulmonary Pump", self.COLOR_VENOUS_BLUE, [
                ("Geometry:", "Crescent-shaped cavity wrapped around the muscular interventricular septum."),
                ("Pressure Profile:", "Systolic 15 - 30 mmHg; End-diastolic 2 - 8 mmHg."),
                ("Outflow Conduit:", "Pumps blood through Pulmonary Semilunar Valve into the low-resistance pulmonary trunk toward alveolar capillary beds."),
                ("Wall Thickness:", "Thin muscular wall (3 - 5 mm) calibrated for high capacitance and low impedance of the pulmonary vascular circuit.")
            ]),
            ("LEFT ATRIUM (LA)", "High-Oxygen Inflow Reservoir", self.COLOR_ARTERIAL_RED, [
                ("Inflow Sources:", "Receives fully oxygenated blood (98% SaO2) from the lungs via 4 Pulmonary Veins (2 superior, 2 inferior; valve-less)."),
                ("Normal Pressure:", "Mean pressure 4 - 12 mmHg (surrogate for Left Ventricular End-Diastolic Pressure)."),
                ("Endocrine Secretion:", "Atrial distension stimulates release of Atrial Natriuretic Peptide (ANP)."),
                ("Clinical Significance:", "Left atrial enlargement predisposes to Atrial Fibrillation and thromboembolic stroke formation in left atrial appendage.")
            ]),
            ("LEFT VENTRICLE (LV)", "High-Pressure Systemic Powerhouse", self.COLOR_ARTERIAL_RED, [
                ("Geometry & Wall:", "Bullet-shaped conical chamber with thick myocardium (10 - 15 mm, 3x thicker than RV)."),
                ("Pressure Profile:", "Peak systolic 100 - 140 mmHg; End-diastolic 3 - 12 mmHg."),
                ("Outflow Conduit:", "Overcomes Systemic Vascular Resistance (SVR) to propel blood through Aortic Valve into the Ascending Aorta."),
                ("Wall Layers:", "Endocardium (smooth inner endothelium), Myocardium (helical spiraling muscle fibers producing wringing systolic ejection), Epicardium.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(chambers):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.0, space_after=4.0)

    def build_slide_3_valves(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 3)
        self._add_header_and_footer(
            slide,
            tag="VALVULAR FLUID MECHANICS",
            title="Heart Valve Mechanics: Unidirectional Flow, Fibrous Skeleton & Auscultation",
            takeaway="Passive pressure-operated flap valves guarantee strictly unidirectional blood flow; closure events produce the classic Lub-Dub (S1/S2) heart sounds.",
            citation="Cardiology Reference: Braunwald's Heart Disease: A Textbook of Cardiovascular Medicine, 12th Ed., Ch. 68.",
            slide_num=3
        )

        w = Inches(2.78)
        h = Inches(4.75)
        gap = Inches(0.2)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        valves = [
            ("TRICUSPID VALVE", "Right Atrioventricular (AV) Gate", self.COLOR_VENOUS_BLUE, [
                ("Anatomy:", "Three fibrous cusps (Anterior, Posterior, Septal) spanning ~7-9 cm² orifice."),
                ("Tethering System:", "Chordae tendineae connect cusps to papillary muscles in the RV wall."),
                ("Closing Mechanics:", "During ventricular systole, rising RV pressure forces cusps together; papillary muscle contraction tensions chordae to prevent prolapse into RA."),
                ("Auscultation:", "Component of S1 heart sound; best heard at the lower left sternal border (4th-5th intercostal space).")
            ]),
            ("PULMONARY VALVE", "Right Ventricular Outflow Gate", self.COLOR_VENOUS_BLUE, [
                ("Structure:", "Three pocket-shaped semilunar cusps (Anterior, Left, Right) with central nodules of Arantius."),
                ("Operation:", "Opens when RV pressure exceeds pulmonary trunk diastolic pressure (>10-15 mmHg)."),
                ("Diastolic Seal:", "Snaps shut when RV relaxes and pulmonary arterial column recoils backward, filling the pocket cusps."),
                ("Auscultation:", "Closure forms P2 component of S2 sound; physiological inspiratory splitting occurs due to increased right-sided venous return.")
            ]),
            ("MITRAL / BICUSPID VALVE", "Left Atrioventricular (AV) Gate", self.COLOR_ARTERIAL_RED, [
                ("Anatomy:", "Two large, durable cusps (Anterior and Posterior) anchored to prominent anterolateral and posteromedial papillary muscles."),
                ("Peak Workload:", "Withstands peak left ventricular systolic pressure (120-140 mmHg) without allowing regurgitation."),
                ("Closure Dynamics:", "Snaps shut at the very onset of ventricular isovolumetric contraction; produces the loud initial M1 component of S1 ('LUB')."),
                ("Clinical Trap:", "Mitral valve prolapse (MVP) and rheumatic mitral stenosis represent the most frequent valvular pathologies worldwide.")
            ]),
            ("AORTIC VALVE", "High-Pressure Systemic Gateway", self.COLOR_ARTERIAL_RED, [
                ("Structure:", "Three thick semilunar cusps (Right, Left, Posterior/Non-coronary) seated within the aortic root."),
                ("Sinuses of Valsalva:", "Anatomical pouches behind the right and left cusps house the ostia of the Right and Left Main Coronary Arteries."),
                ("Diastolic Closure:", "Closure produces the crisp A2 component of S2 ('DUB'); diastolic recoil directly perfuses the coronary vascular bed!"),
                ("Hemodynamic Metric:", "Normal orifice area is 3.0 - 4.0 cm²; severe aortic stenosis (<1.0 cm²) drastically elevates LV afterload.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(valves):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.0, space_after=4.0)

    def build_slide_4_circulation(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 4)
        self._add_header_and_footer(
            slide,
            tag="HEMODYNAMIC CIRCUITS",
            title="Dual Circulation: Pulmonary Circuit vs. Systemic High-Pressure Vascular Circuit",
            takeaway="William Harvey's closed dual circuit: Right heart drives gas exchange at 1/6th systemic pressure; Left heart propels life-sustaining perfusion to peripheral beds.",
            citation="Physiology Source: Hall, J. E., Guyton and Hall Textbook of Medical Physiology, 14th Ed., Elsevier, Unit IV.",
            slide_num=4
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        circuits = [
            ("PULMONARY CIRCUIT (LOW-PRESSURE)", "Alveolar Gas Exchange Loop", self.COLOR_VENOUS_BLUE, [
                ("Vascular Path:", "Right Ventricle -> Pulmonary Trunk -> Left & Right Pulmonary Arteries -> Lobar/Segmental Arteries -> Alveolar Capillaries -> 4 Pulmonary Veins -> Left Atrium."),
                ("Hemodynamic Profile:", "Low pressure, low resistance (Mean PAP ~14 mmHg; Systolic/Diastolic 25/10 mmHg). Pulmonary vascular resistance (PVR) is 1/10th of SVR."),
                ("Gas Exchange:", "Deoxygenated blood (pO2 ~40 mmHg, pCO2 ~46 mmHg) offloads CO2 across the 0.5-μm alveolar-capillary membrane and binds O2 to reach pO2 ~100 mmHg (98% SaO2)."),
                ("Unique Physiology:", "Hypoxic Pulmonary Vasoconstriction (HPV): Pulmonary vessels constrict in unventilated lung zones to shunt blood to ventilated alveoli, maximizing V/Q matching.")
            ]),
            ("SYSTEMIC CIRCUIT (HIGH-PRESSURE)", "Whole-Body Peripheral Perfusion Loop", self.COLOR_ARTERIAL_RED, [
                ("Vascular Path:", "Left Ventricle -> Aortic Root -> Ascending/Arch/Descending Aorta -> Elastic Arteries -> Muscular Conduit Arteries -> Arterioles -> Capillaries -> Venules -> Vena Cavae."),
                ("Hemodynamic Profile:", "High pressure, high resistance (Systolic 120 mmHg, Diastolic 80 mmHg, Mean Arterial Pressure MAP ~93 mmHg)."),
                ("Stopcock Regulation:", "Arterioles (rich in smooth muscle) serve as primary resistance vessels, dynamically adjusting vascular tone via autonomic and local metabolic factors."),
                ("Nutrient Delivery:", "Unloads O2, glucose, fatty acids, amino acids, electrolytes, and hormones; sweeps away lactic acid, urea, creatinine, and cellular metabolites.")
            ]),
            ("MICROCIRCULATION & STARLING FORCES", "Capillary Exchange Dynamics", self.COLOR_PERFUSION_TEAL, [
                ("Capillary Bed Surface:", "Over 10 billion capillaries providing a massive ~500-700 m² surface area for diffusion."),
                ("Starling Equation:", "Net Fluid Movement $J_v = K_f [(P_c - P_i) - \sigma (\pi_c - \pi_i)]$. Hydrostatic pressure ($P_c$) drives filtration; oncotic pressure ($\pi_c$) drives reabsorption."),
                ("Arterial Capillary End:", "$P_c \approx 32\text{ mmHg} > \pi_c \approx 25\text{ mmHg}$; net filtration pushes fluid, electrolytes, and nutrients into interstitial space."),
                ("Venous Capillary End:", "$P_c \approx 15\text{ mmHg} < \pi_c \approx 25\text{ mmHg}$; net absorption draws 85-90% of fluid back; remaining 10-15% cleared via lymphatic vessels back to venous angle.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(circuits):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_5_conduction(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 5)
        self._add_header_and_footer(
            slide,
            tag="ELECTROPHYSIOLOGY & CONDUCTION",
            title="The Cardiac Conduction System & Surface ECG Vector Mapping",
            takeaway="Rhythmic cardiac pumping relies on an intrinsic autonomous conduction hierarchy; defects lead to lethal arrhythmias or heart block.",
            citation="Electrophysiology Source: Zipes, Jalife & Stevenson, Cardiac Electrophysiology: From Cell to Bedside, 8th Ed.",
            slide_num=5
        )

        w = Inches(2.2)
        h = Inches(4.75)
        gap = Inches(0.18)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        nodes = [
            ("1. SA NODE", "Primary Pacemaker", self.COLOR_ELECTRICAL_GOLD, [
                ("Location:", "Subepicardial at SVC-RA junction."),
                ("Intrinsic Rate:", "60 - 100 impulses/min."),
                ("Phase 4 Drift:", "Funny channels ($I_f$) carry slow inward Na+ influx; T-type Ca2+ channels trigger threshold at -40 mV."),
                ("Vagal Dominance:", "Resting vagal tone suppresses intrinsic 100 bpm to resting 60-80 bpm.")
            ]),
            ("2. ATRIAL TRACTS", "Interatrial Wavefront", self.COLOR_ELECTRICAL_GOLD, [
                ("Conduction Paths:", "Anterior, Middle (Wenckebach), and Posterior (Thorel) internodal pathways."),
                ("Bachmann Bundle:", "Carries wave from RA across septum to LA ensuring synchronized bi-atrial contraction."),
                ("ECG Correlation:", "Corresponds to the P WAVE (atrial depolarization; normal duration <0.12s, amplitude <2.5 mm).")
            ]),
            ("3. AV NODE", "Physiological Gatekeeper", self.COLOR_ARTERIAL_RED, [
                ("Location:", "Triangle of Koch in interatrial septum."),
                ("Critical Delay:", "Imposes a 0.12s delay due to slow Ca2+ conduction and fewer gap junctions."),
                ("Function:", "Allows complete ventricular filling prior to systole and protects ventricles from rapid atrial flutter/fibrillation."),
                ("ECG Correlation:", "Maps to the isoelectric PR INTERVAL (normal 0.12 - 0.20s).")
            ]),
            ("4. BUNDLE OF HIS", "Septal Distributor", self.COLOR_VENOUS_BLUE, [
                ("Anatomy:", "Penetrates the central fibrous skeleton of the heart into muscular interventricular septum."),
                ("Bifurcation:", "Splits into Right Bundle Branch (RBB) and Left Bundle Branch (LBB; subdivides into Anterior & Posterior fascicles)."),
                ("Speed:", "High conduction velocity (~2 m/s).")
            ]),
            ("5. PURKINJE SYSTEM", "Apex-to-Base Ejection", self.COLOR_PERFUSION_TEAL, [
                ("Network:", "Large specialized subendocardial myocytes with abundant gap junctions."),
                ("Velocity:", "Fastest conduction in heart: 2 - 4 m/s."),
                ("Apex-to-Base:", "Triggers ventricular depolarization from cardiac apex toward outflow base."),
                ("ECG Correlation:", "QRS COMPLEX (<0.12s) = ventricular depolarization; ST segment & T WAVE = ventricular repolarization.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(nodes):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=8.8, space_after=4.0)

    def build_slide_6_cardiac_cycle(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 6)
        self._add_header_and_footer(
            slide,
            tag="CARDIAC CYCLE & HEMODYNAMICS",
            title="The Cardiac Cycle: Wiggers Diagram & Pressure-Volume Relationships",
            takeaway="Wiggers diagram integrates ECG, ventricular pressure, aortic pressure, ventricular volume, and phonocardiogram into 4 phases.",
            citation="Hemodynamics Reference: Mohrman, D. E. & Heller, L. J., Cardiovascular Physiology, 10th Ed., McGraw-Hill.",
            slide_num=6
        )

        w = Inches(2.78)
        h = Inches(4.75)
        gap = Inches(0.2)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        phases = [
            ("PHASE 1: ISOVOLUMETRIC CONTRACTION", "Tension Rise with Constant Volume", self.COLOR_ARTERIAL_RED, [
                ("Valvular State:", "ALL FOUR VALVES CLOSED (Mitral closes as LV pressure exceeds LA; Aortic valve remains closed)."),
                ("Timeframe:", "Begins immediately following QRS complex (~0.05 seconds duration)."),
                ("Pressure Jump:", "LV pressure soars from ~8 mmHg to 80 mmHg with zero change in chamber volume (EDV ~120 mL)."),
                ("Phonocardiogram:", "Mitral and Tricuspid closure vibrations generate First Heart Sound (S1 'LUB').")
            ]),
            ("PHASE 2: VENTRICULAR EJECTION", "Rapid & Reduced Systolic Outflow", self.COLOR_ARTERIAL_RED, [
                ("Valve Action:", "Aortic valve opens once LV pressure eclipses aortic diastolic pressure (>80 mmHg)."),
                ("Rapid Ejection:", "Initial 1/3 of ejection propels 70% of stroke volume; LV pressure peaks at 120 mmHg."),
                ("Reduced Ejection:", "Later 2/3 shows decelerating outflow; aortic pressure declines as runoff exceeds entry."),
                ("Volume Shift:", "Chamber volume plummets from EDV (120 mL) to End-Systolic Volume (ESV ~50 mL). Stroke Volume ejected = 70 mL.")
            ]),
            ("PHASE 3: ISOVOLUMETRIC RELAXATION", "Pressure Collapse with Constant Volume", self.COLOR_VENOUS_BLUE, [
                ("Valve Action:", "ALL FOUR VALVES CLOSED (Aortic valve snaps shut as LV pressure drops below aortic pressure)."),
                ("Incisura / Dicrotic Notch:", "Brief pressure rebound in aorta caused by elastic recoil of aortic root cusps."),
                ("Pressure Plunge:", "LV pressure falls precipitously from ~100 mmHg down to near 0-5 mmHg."),
                ("Phonocardiogram:", "Aortic and Pulmonic valve closure generates Second Heart Sound (S2 'DUB').")
            ]),
            ("PHASE 4: VENTRICULAR FILLING", "Rapid Passive Inflow & Atrial Kick", self.COLOR_VENOUS_BLUE, [
                ("Valve Action:", "Mitral valve opens when LV pressure drops below Left Atrial pressure (<8-10 mmHg)."),
                ("Rapid Filling (80%):", "Blood stored in LA rushes passively into relaxing ventricle; suction effect of myocyte elastic recoil."),
                ("Atrial Kick (20%):", "Atrial systole (P wave) delivers final 15-20% boost to ventricular end-diastolic volume (EDV ~120 mL)."),
                ("Clinical Trap:", "Loss of atrial kick in Atrial Fibrillation can cause acute decompensation in stiff, hypertrophied ventricles.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(phases):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.0, space_after=4.0)

    def build_slide_7_systemic_perfusion(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 7)
        self._add_header_and_footer(
            slide,
            tag="SYSTEMIC ORGAN USES & PERFUSION",
            title="Systemic Organ Perfusion: How the Heart Serves Brain, Kidneys, Liver & Muscle",
            takeaway="Cardiac output distribution matches metabolic demand and organ hierarchy: Brain demands continuous perfusion; Kidneys demand massive flow for filtration.",
            citation="Clinical Physiology: Boron, W. F. & Boulpaep, E. L., Medical Physiology, 3rd Ed., Elsevier, Ch. 24.",
            slide_num=7
        )

        w = Inches(2.78)
        h = Inches(4.75)
        gap = Inches(0.2)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        organs = [
            ("THE BRAIN", "15% Cardiac Output (~750 mL/min)", self.COLOR_ARTERIAL_RED, [
                ("Conduit Vessels:", "Internal Carotid & Vertebral arteries form the anastomotic Circle of Willis at brain base."),
                ("Metabolic Demand:", "Consumes 20% of total resting body oxygen despite accounting for only 2% of body mass!"),
                ("Cerebral Autoregulation:", "Maintains constant blood flow across MAP range 60 - 160 mmHg via myogenic tone and local CO2/H+ vasoactivity."),
                ("Ischemic Vulnerability:", "Zero anaerobic capacity: syncope occurs within 6-8s of circulatory arrest; irreversible neuronal death occurs in 4-5 minutes.")
            ]),
            ("THE KIDNEYS", "20 - 25% Cardiac Output (~1,200 mL/min)", self.COLOR_PERFUSION_TEAL, [
                ("Massive Flow Ratio:", "Receives the highest blood flow per gram of tissue of any organ in the human body (~400 mL/min/100g)."),
                ("Primary Purpose:", "Not for oxygen delivery, but to power Glomerular Filtration Rate (GFR ~125 mL/min = 180 L plasma filtered daily!)."),
                ("Renal Autoregulation:", "Myogenic response and Tubuloglomerular Feedback (TGF via Macula Densa) preserve GFR at MAP 80-180 mmHg."),
                ("Endocrine Feedback:", "Hypoperfusion triggers Renin release from juxtaglomerular cells -> Angiotensin II -> Aldosterone (RAAS cascade).")
            ]),
            ("LIVER & GI TRACT", "25% Cardiac Output (~1,300 mL/min)", self.COLOR_ELECTRICAL_GOLD, [
                ("Dual Blood Supply:", "Hepatic Artery (25% oxygenated flow from celiac trunk) + Portal Vein (75% venous inflow draining stomach, intestines, spleen)."),
                ("Functional Use:", "Carries absorbed carbohydrates, amino acids, and lipids directly to hepatocytes for metabolic processing, storage, and first-pass clearance."),
                ("Blood Reservoir:", "Hepatic sinusoids and splanchnic venules hold ~500-1,000 mL blood; sympathetic vasoconstriction mobilizes this pool during hemorrhage."),
                ("Bile & Protein Synthesis:", "Drives albumin synthesis (maintains plasma oncotic pressure) and coagulation factor production (I, II, VII, IX, X).")
            ]),
            ("SKELETAL MUSCLE & SKIN", "20% at Rest -> 85% in Heavy Exercise", self.COLOR_VENOUS_BLUE, [
                ("Dynamic Hyperaemia:", "Resting flow (~1,000 mL/min) surges up to 15,000 - 20,000 mL/min during maximal exercise!"),
                ("Local Vasodilation:", "Muscle contractions release adenosine, lactate, K+, nitric oxide (NO), and H+, dilating precapillary sphincters."),
                ("Cardiac Compensation:", "Heart rate rises from 70 to 180 bpm; stroke volume rises from 70 to 110 mL; CO expands from 5 to 22 L/min."),
                ("Cutaneous Shunt:", "Skin blood flow adjusts dynamically: dilates for radiative heat dissipation, constricts in cold to conserve core temperature.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(organs):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.0, space_after=4.0)

    def build_slide_8_neuroendocrine(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 8)
        self._add_header_and_footer(
            slide,
            tag="NEUROENDOCRINE INTEGRATION",
            title="Autonomic Control of the Heart & Endocrine Secretion of ANP and BNP",
            takeaway="The heart is both a receptor of neural regulation and an active endocrine gland releasing natriuretic peptides to protect against volume overload.",
            citation="Endocrinology Source: Gardner, D. G., Natriuretic Peptides in Cardiovascular Homeostasis, J. Clin. Endocrinol. Metab.",
            slide_num=8
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        pillars = [
            ("SYMPATHETIC STIMULATION", "Fight-or-Flight Acceleration", self.COLOR_ARTERIAL_RED, [
                ("Neural Pathway:", "Postganglionic sympathetic fibers originate from T1-T4 cervical & upper thoracic ganglia (cardiac accelerator nerves)."),
                ("Receptor Dynamics:", "Releases Norepinephrine (and adrenal Epinephrine) acting on $\beta_1$-adrenergic Gs-protein coupled receptors."),
                ("Four Cardinal Effects:", "• Positive Chronotropy: $\\uparrow$ SA node firing rate ($I_f$ current).\n• Positive Inotropy: $\\uparrow$ myocardial contractility (Ca2+ release via RyR2).\n• Positive Dromotropy: $\\uparrow$ AV nodal conduction speed.\n• Positive Lusitropy: $\\uparrow$ speed of relaxation (phospholamban phosphorylation)."),
                ("Cardiac Reserve:", "Enables cardiac output to escalate 4-fold to 5-fold during acute physiological stress.")
            ]),
            ("PARASYMPATHETIC BRAKE", "Rest-and-Digest Deceleration", self.COLOR_VENOUS_BLUE, [
                ("Neural Pathway:", "Preganglionic fibers travel via the Right & Left Vagus Nerves (Cranial Nerve X)."),
                ("Receptor Dynamics:", "Releases Acetylcholine (ACh) onto $M_2$ muscarinic Gi-protein coupled receptors in SA and AV nodes."),
                ("Physiological Effects:", "• Negative Chronotropy: Activates $I_{K,ACh}$ channels causing hyperpolarization; reduces heart rate.\n• Negative Dromotropy: Slows AV nodal conduction; prolonged high vagal tone can induce AV block.\n• Minimal Ventricular Inotropy: Ventricles have scarce direct parasympathetic innervation."),
                ("Resting Vagal Tone:", "Prevents the heart from racing at its intrinsic SA node rate of 100 bpm, keeping healthy resting heart rate between 60 - 80 bpm.")
            ]),
            ("HEART AS AN ENDOCRINE GLAND", "ANP & BNP Natriuretic Peptides", self.COLOR_NEURO_PURPLE, [
                ("Atrial Natriuretic Peptide (ANP):", "28-amino-acid peptide synthesized and stored in granules within atrial myocytes. Secreted upon atrial mechanical stretch caused by hypervolemia."),
                ("B-Type Natriuretic Peptide (BNP):", "32-amino-acid peptide produced primarily in ventricular myocytes in response to ventricular wall stress, pressure overload, and volume expansion."),
                ("Receptors & Second Messenger:", "Bind to NPR-A receptors containing guanylyl cyclase domains, boosting intracellular cyclic GMP (cGMP)."),
                ("Physiological Counter-Regulation:", "• Renal: Increases GFR and inhibits Na+ reabsorption, driving natriuresis and diuresis.\n• Vascular: Induces systemic arterial vasodilation, lowering BP.\n• Endocrine: Potently suppresses Renin, Aldosterone, and Endothelin.\n• Diagnostic: NT-proBNP is the universal biomarker for Heart Failure.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(pillars):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_9_coronary_metabolism(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 9)
        self._add_header_and_footer(
            slide,
            tag="CORONARY FLOW & METABOLISM",
            title="Coronary Perfusion Dynamics & Myocardial Energy Substrates",
            takeaway="The heart is an omnivorous metabolic engine extracting 75% of arterial O2; Left ventricular myocardium receives perfusion almost exclusively during diastole.",
            citation="Cardiovascular Metabolism: Opie, L. H., Heart Physiology: From Cell to Circulation, 4th Ed., Lippincott Williams & Wilkins.",
            slide_num=9
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        cards = [
            ("CORONARY ARTERIAL ANATOMY", "Blood Supply to the Pump", self.COLOR_METABOLIC_ROSE, [
                ("Arterial Origins:", "Arise from the left and right sinuses of Valsalva immediately above the aortic valve cusps."),
                ("Left Coronary System:", "Left Main breaks into Left Anterior Descending (LAD - 'widowmaker', supplies anterior wall & septum) and Circumflex (LCx, supplies lateral LV)."),
                ("Right Coronary System:", "Right Coronary Artery (RCA) runs in atrioventricular groove, giving off Marginal branch and Posterior Descending Artery (PDA in 85% right-dominant hearts)."),
                ("Nodal Perfusion:", "SA nodal artery arises from RCA (60%) or LCx (40%); AV nodal artery arises from RCA in 90% of individuals.")
            ]),
            ("PHASIC CORONARY FLOW DYNAMICS", "The Diastolic Perfusion Imperative", self.COLOR_ARTERIAL_RED, [
                ("Systolic Compression:", "During ventricular systole, powerful intramyocardial pressure exceeds intracoronary pressure, squeezing intramural vessels shut (flow drops to near zero!)."),
                ("Diastolic Window:", "80% of Left Ventricular coronary blood flow occurs during DIASTOLE, when the myocardium relaxes and aortic diastolic recoil forces blood into coronary ostia."),
                ("Clinical Hazard of Tachycardia:", "Extreme heart rates (>140 bpm) shorten diastolic filling time far more than systolic time, severely compromising coronary perfusion and triggering myocardial ischemia."),
                ("Autoregulation:", "Local metabolites (Adenosine, NO, H+, pO2 drop) regulate coronary arteriolar resistance to match myocardial O2 demand.")
            ]),
            ("METABOLIC SUBSTRATE PREFERENCE", "ATP Production in Cardiac Mitochondria", self.COLOR_PERFUSION_TEAL, [
                ("Mitochondrial Density:", "Mitochondria occupy 30 - 40% of adult cardiomyocyte volume, reflecting absolute reliance on aerobic respiration."),
                ("Substrate Hierarchy at Rest:", "• Free Fatty Acids (FFA): 60 - 70% of ATP generated via $\beta$-oxidation.\n• Glucose & Pyruvate: 20 - 25% of energy needs.\n• Lactate & Ketone Bodies: 10% under resting conditions."),
                ("Metabolic Flexibility:", "During exercise, myocardium readily extracts and oxidizes circulating lactate produced by working skeletal muscles!"),
                ("Extreme O2 Extraction:", "Myocardium extracts 70 - 80% of available arterial oxygen at rest (coronary sinus pO2 is only ~20 mmHg); thus, increased workload requires increasing blood flow.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cards):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_10_biomarkers_table(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 10)
        self._add_header_and_footer(
            slide,
            tag="CLINICAL HEMODYNAMICS & BIOMARKERS",
            title="Hemodynamic Reference Values & Diagnostic Cardiac Biomarkers",
            takeaway="Quantitative clinical monitoring of pressure, flow, wall stress, and myocyte necrosis guides critical care management.",
            citation="Clinical Reference: ESC/ACC Guidelines for Diagnosis and Management of Heart Failure and Myocardial Infarction.",
            slide_num=10
        )

        headers = ["Parameter / Biomarker", "Normal Reference Range", "Physiological Significance", "Clinical Diagnostic Role", "Target Organ Affected"]
        col_widths = [Inches(2.5), Inches(2.2), Inches(2.6), Inches(2.4), Inches(2.033)]
        rows_data = [
            ["Cardiac Output (CO)", "4.0 - 8.0 L/min", "Total volume ejected by ventricle per minute (CO = HR x SV)", "Identifies cardiogenic shock, low-output states, sepsis", "All systemic organs"],
            ["Mean Arterial Pressure (MAP)", "70 - 105 mmHg", "Average organ perfusion pressure over complete cycle", "MAP <65 mmHg causes irreversible renal & cerebral ischemia", "Brain, Kidneys, Viscera"],
            ["Central Venous Pressure (CVP)", "2 - 6 mmHg (3 - 8 cm H2O)", "Right atrial filling pressure; assesses intravascular volume", "Elevated in right ventricular failure and hypervolemia", "Liver (congestive hepatopathy)"],
            ["High-Sensitivity Troponin I / T", "<14 ng/L (assay dependent)", "Structural regulatory protein of cardiomyocyte myofibrils", "Gold standard for Acute Myocardial Infarction (AMI)", "Myocardium"],
            ["NT-proBNP / BNP", "<125 pg/mL (<300 pg/mL in acute dyspnea)", "Cleaved neurohormone secreted during ventricular wall strain", "Differentiates cardiac dyspnea from pulmonary disease", "Heart, Kidneys, Vasculature"],
            ["Ejection Fraction (LVEF)", "55% - 70%", "Fraction of end-diastolic volume ejected per beat (SV / EDV)", "Classifies heart failure (HFrEF vs. HFpEF)", "Whole-body circulation"],
            ["Creatine Kinase-MB (CK-MB)", "<5 ng/mL (or <3-5% of total CK)", "Isoenzyme present in high concentration in myocardium", "Detects re-infarction due to rapid 48-72h clearance window", "Cardiac muscle"]
        ]

        self.create_table(slide, Inches(0.8), Inches(1.95), Inches(11.733), Inches(4.75), headers, col_widths, rows_data)

    def build_slide_11_pathology(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 11)
        self._add_header_and_footer(
            slide,
            tag="PATHOPHYSIOLOGICAL FAILURE CASCADES",
            title="Cardiovascular Disease Cascades: From Ischemia to End-Organ Failure",
            takeaway="Failure of the central pump triggers devastating multi-organ failure through forward hypoperfusion and backward venous congestion.",
            citation="Pathology Source: Robbins and Cotran Pathologic Basis of Disease, 10th Ed., Ch. 12 (The Heart).",
            slide_num=11
        )

        w = Inches(3.75)
        h = Inches(4.75)
        gap = Inches(0.24)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        cascades = [
            ("ISCHEMIC CASCADE & INFARCTION", "Coronary Occlusion to Necrosis", self.COLOR_ARTERIAL_RED, [
                ("Trigger:", "Rupture of an unstable atherosclerotic plaque with superimposed platelet thrombosis occludes a coronary artery branch."),
                ("Timeline of Injury:", "• Within 60 seconds: Myocyte ATP drops; systolic contractility ceases immediately.\n• Within 10 minutes: Massive edema, loss of membrane potential, arrhythmia risk.\n• Within 20 - 40 minutes: Irreversible coagulation necrosis begins from subendocardium to subepicardium (wavefront of ischemia)."),
                ("Clinical Outcome:", "ST-elevation MI (STEMI), loss of functional contractile mass, ventricular arrhythmias (V-fib), or free wall rupture.")
            ]),
            ("CONGESTIVE HEART FAILURE (CHF)", "Neurohormonal Maladaptive Spiral", self.COLOR_VENOUS_BLUE, [
                ("Initial Insult:", "Loss of myocyte contractility (HFrEF) or stiff, non-compliant filling (HFpEF) reduces forward cardiac output."),
                ("Vicious Compensatory Loop:", "Hypoperfusion triggers sympathetic hyperactivity and excessive RAAS activation:\n• Chronic Norepinephrine induces myocyte apoptosis and fibrosis.\n• Angiotensin II & Aldosterone promote toxic cardiac remodeling.\n• Fluid retention elevates venous pressures."),
                ("Clinical Manifestation:", "Left-sided failure -> Pulmonary venous congestion, dyspnea, orthopnea, pulmonary edema. Right-sided failure -> Jugular venous distension, hepatosplenomegaly, peripheral edema.")
            ]),
            ("CARDIORENAL & SHOCK CASCADES", "Terminal Multi-Organ Dysfunction", self.COLOR_NEURO_PURPLE, [
                ("Cardiorenal Syndrome (Type 1 & 2):", "Drop in cardiac output + elevated renal venous pressure ($P_{rv}$) compresses renal tubules, collapses GFR, and precipitates acute oliguric kidney injury."),
                ("Cardiogenic Shock:", "Sustained hypotension (SBP <90 mmHg, Cardiac Index <2.2 L/min/m²) despite adequate filling pressures."),
                ("Multi-Organ Failure (MODS):", "• Brain: Hypoxic ischemic encephalopathy, coma.\n• Liver: Centrilobular necrosis ('nutmeg liver' & acute ischemic hepatitis).\n• Gut: Mesenteric ischemia, translocation of endotoxins causing systemic inflammation."),
                ("Mortality:", "Cardiogenic shock carries 40 - 50% in-hospital mortality without rapid mechanical circulatory support (ECMO/Impella).")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(cascades):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=9.2, space_after=4.5)

    def build_slide_12_summary_infographic(self):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_canvas_background(slide, 12)
        self._add_header_and_footer(
            slide,
            tag="MASTER INFOGRAPHIC SYNTHESIS",
            title="The 5 Fundamental Uses of the Heart Organ in the Human Body",
            takeaway="Summary Blueprint: The heart integrates mechanical pumping, gas transport, fluid balance, metabolic cleansing, and endocrine signaling to preserve human life.",
            citation="Comprehensive Medical Physiology Synthesis | Master Keynote Presentation",
            slide_num=12
        )

        w = Inches(2.2)
        h = Inches(4.75)
        gap = Inches(0.18)
        start_x = Inches(0.8)
        top_y = Inches(1.95)

        pillars = [
            ("1. PRESSURE ENGINE", "Continuous Pressure Gradient", self.COLOR_ARTERIAL_RED, [
                ("Function:", "Maintains a constant hydrostatic pressure head (MAP ~93 mmHg) to drive continuous blood flow across high-resistance capillary networks."),
                ("Key Mechanics:", "Elastic arterial recoil (Windkessel effect) cushions high systolic pressure and preserves diastolic forward flow."),
                ("Organ Benefit:", "Guarantees capillary filtration across all vascular beds without stagnant stasis.")
            ]),
            ("2. RESPIRATORY CONDUIT", "O2 & CO2 Gas Transportation", self.COLOR_VENOUS_BLUE, [
                ("Function:", "Pumps entire blood volume through alveolar membranes every minute, delivering 1,000 mL O2/min to peripheral tissues."),
                ("Key Mechanics:", "Maintains hemoglobin saturation at 98% and rapid transit of carbon dioxide (as bicarbonate and carbamino compounds) to lungs."),
                ("Organ Benefit:", "Prevents anaerobic lactic acidosis and cellular energy failure in all vital organs.")
            ]),
            ("3. METABOLIC HIGHWAY", "Nutrient & Waste Cleansing", self.COLOR_ELECTRICAL_GOLD, [
                ("Function:", "Connects intestinal nutrient absorption with cellular metabolism and organs of excretion."),
                ("Key Mechanics:", "Circulates glucose, fatty acids, amino acids, trace minerals, vitamins, and pharmacological compounds to target tissues."),
                ("Organ Benefit:", "Shunts nitrogenous wastes (urea, creatinine) to kidneys and hepatic toxins to liver.")
            ]),
            ("4. ENDOCRINE SENSOR", "Fluid Volume & Tone Governor", self.COLOR_NEURO_PURPLE, [
                ("Function:", "Directly senses cardiovascular stretch and releases ANP and BNP hormones into circulation."),
                ("Key Mechanics:", "Activates particulate guanylyl cyclase, elevating cGMP to relax vascular smooth muscle and accelerate renal salt excretion."),
                ("Organ Benefit:", "Serves as the internal volume protector preventing acute pulmonary edema.")
            ]),
            ("5. THERMAL & IMMUNE LINE", "Body Defense & Heat Buffer", self.COLOR_PERFUSION_TEAL, [
                ("Function:", "Distributes core thermal energy to peripheral cutaneous capillary beds for evaporative and radiative heat loss."),
                ("Key Mechanics:", "Constricts in freezing environments to protect core cardiac and cerebral perfusion."),
                ("Organ Benefit:", "Rapidly transports neutrophils, lymphocytes, platelets, and immunoglobulins to injury sites within seconds.")
            ])
        ]

        for i, (title, sub, col, bullets) in enumerate(pillars):
            cx = start_x + i * (w + gap)
            ctf = self.create_card(slide, cx, top_y, w, h, title, sub, accent_color=col)
            for bold_p, text in bullets:
                self.append_bullet(ctf, bold_p, text, font_size=8.8, space_after=4.0)

    # -------------------------------------------------------------
    # BUILD & EXPORT ALL SLIDES
    # -------------------------------------------------------------
    def build_all(self, output_pptx_path: Path):
        print("Generating Slide 1: Infographic Cover...")
        self.build_slide_1_cover()
        print("Generating Slide 2: Anatomical Blueprint & 4 Chambers...")
        self.build_slide_2_chambers()
        print("Generating Slide 3: Heart Valve Dynamics...")
        self.build_slide_3_valves()
        print("Generating Slide 4: Dual Circulation (Pulmonary & Systemic)...")
        self.build_slide_4_circulation()
        print("Generating Slide 5: Cardiac Conduction & ECG...")
        self.build_slide_5_conduction()
        print("Generating Slide 6: Cardiac Cycle & Wiggers Diagram...")
        self.build_slide_6_cardiac_cycle()
        print("Generating Slide 7: Systemic Organ Perfusion...")
        self.build_slide_7_systemic_perfusion()
        print("Generating Slide 8: Neuroendocrine Regulation (ANP/BNP)...")
        self.build_slide_8_neuroendocrine()
        print("Generating Slide 9: Coronary Microcirculation & Metabolism...")
        self.build_slide_9_coronary_metabolism()
        print("Generating Slide 10: Hemodynamics & Biomarkers Matrix...")
        self.build_slide_10_biomarkers_table()
        print("Generating Slide 11: Disease Cascades & Organ Failure...")
        self.build_slide_11_pathology()
        print("Generating Slide 12: Master 5-Pillar Synthesis Infographic...")
        self.build_slide_12_summary_infographic()

        output_pptx_path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(output_pptx_path))
        print(f"[SUCCESS] Infographic presentation saved to: {output_pptx_path}")


def export_deck_to_pdf_and_gallery(pptx_path: Path, pdf_path: Path, gallery_dir: Path):
    """Uses PowerPoint COM automation to export pristine vector PDF and PNG gallery."""
    import win32com.client
    import pythoncom

    print(f"Exporting PDF via PowerPoint COM -> {pdf_path}")
    gallery_dir.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        presentation = ppt_app.Presentations.Open(str(pptx_path.resolve()), WithWindow=False)
        # 32 = ppSaveAsPDF
        presentation.SaveAs(str(pdf_path.resolve()), 32)
        print(f"[SUCCESS] Exported vector PDF to: {pdf_path}")

        # Export high-res PNGs of all slides
        for idx, slide in enumerate(presentation.Slides):
            slide_png = gallery_dir / f"Slide_{idx+1:02d}.PNG"
            slide.Export(str(slide_png.resolve()), "PNG", 1920, 1080)
        print(f"[SUCCESS] Exported {len(presentation.Slides)} PNG slides to: {gallery_dir}")

        presentation.Close()
    finally:
        ppt_app.Quit()
        pythoncom.CoUninitialize()


def main():
    samples_dir = REPO_ROOT / "samples"
    pptx_path = samples_dir / "heart_organ_infographic_deck.pptx"
    pdf_path = samples_dir / "heart_organ_infographic_deck.pdf"
    gallery_dir = samples_dir / "gallery_heart_infographic"

    print("=================================================================")
    print("HUMAN HEART ORGAN USES INFOGRAPHIC DECK GENERATOR")
    print("=================================================================")
    builder = HeartInfographicDeckBuilder()
    builder.build_all(pptx_path)

    try:
        export_deck_to_pdf_and_gallery(pptx_path, pdf_path, gallery_dir)
    except Exception as e:
        print(f"[WARNING] COM export failed or timed out: {e}")


if __name__ == "__main__":
    main()
