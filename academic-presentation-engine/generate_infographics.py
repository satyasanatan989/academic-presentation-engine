"""
Scientific Infographics and Pathway Flowcharts Generator
Generates publication-quality pathway diagrams and infographics for:
- Glucosinolate-Myrosinase enzymatic cascade (Slide 3)
- Sulforaphane Nrf2-Keap1 Phase II gene induction (Slide 4)
- Ellagitannin-Urolithin A Mitophagy pathway (Slide 9)
- Curcumin-Piperine pharmacokinetics & bioavailability (Slide 13)
- Synergistic Food Pairings & Bioavailability Stacking (Slide 16)
"""

from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.path import Path as MPath

ASSETS_DIR = Path(__file__).resolve().parent / "outputs" / "PRJ-PHYTOCHEMICALS-001" / "assets"
ASSETS_DIR.mkdir(parents=True, exist_ok=True)

# Styling parameters
FONT_FAMILY = "sans-serif"
COLOR_NAVY = "#133826"
COLOR_GREEN = "#2D6A4F"
COLOR_SAGE = "#40916C"
COLOR_LIGHT_BG = "#F6FAF7"
COLOR_BOX_BG = "#FFFFFF"
COLOR_ACCENT = "#D9534F"
COLOR_GOLD = "#D48F18"
COLOR_BORDER = "#B7D1C2"


def generate_glucosinolate_flowchart():
    """Generates Slide 3 Enzymatic Cascade Flowchart."""
    fig, ax = plt.subplots(figsize=(10, 4.2), dpi=250)
    ax.set_facecolor(COLOR_LIGHT_BG)
    fig.patch.set_facecolor(COLOR_LIGHT_BG)

    # Title
    ax.text(5, 3.8, "Enzymatic Myrosinase Activation Cascade in Cruciferae",
            ha="center", va="center", fontsize=13, fontweight="bold", color=COLOR_NAVY, fontfamily=FONT_FAMILY)

    # Box 1: Precursor
    b1 = patches.FancyBboxPatch((0.5, 2.0), 2.2, 1.2, boxstyle="round,pad=0.1", fc="#E8F4EC", ec=COLOR_GREEN, lw=2)
    ax.add_patch(b1)
    ax.text(1.6, 2.7, "Glucoraphanin", ha="center", va="center", fontsize=11, fontweight="bold", color=COLOR_NAVY)
    ax.text(1.6, 2.3, "(Intact Vacuoles)", ha="center", va="center", fontsize=9, color="#4D6253")

    # Plus sign
    ax.text(3.05, 2.6, "+", ha="center", va="center", fontsize=18, fontweight="bold", color=COLOR_GREEN)

    # Box 2: Enzyme
    b2 = patches.FancyBboxPatch((3.4, 2.0), 2.2, 1.2, boxstyle="round,pad=0.1", fc="#E8F4EC", ec=COLOR_GREEN, lw=2)
    ax.add_patch(b2)
    ax.text(4.5, 2.7, "Myrosinase", ha="center", va="center", fontsize=11, fontweight="bold", color=COLOR_NAVY)
    ax.text(4.5, 2.3, "(Cytosolic Cells)", ha="center", va="center", fontsize=9, color="#4D6253")

    # Arrow with trigger label
    ax.annotate("", xy=(6.5, 2.6), xytext=(5.8, 2.6),
                arrowprops=dict(arrowstyle="->", color=COLOR_GREEN, lw=2.5, mutation_scale=15))
    ax.text(6.15, 2.9, "Tissue\nDisruption", ha="center", va="center", fontsize=8.5, fontweight="bold", color=COLOR_ACCENT)

    # Box 3: Intermediate
    b3 = patches.FancyBboxPatch((6.6, 2.0), 2.0, 1.2, boxstyle="round,pad=0.1", fc="#FFF3E0", ec=COLOR_GOLD, lw=2)
    ax.add_patch(b3)
    ax.text(7.6, 2.7, "Unstable Aglycone", ha="center", va="center", fontsize=10.5, fontweight="bold", color="#7A4B00")
    ax.text(7.6, 2.3, "Intermediate", ha="center", va="center", fontsize=9, color="#7A4B00")

    # Bifurcation arrows
    # Top arrow: Neutral pH
    ax.annotate("", xy=(9.0, 3.0), xytext=(8.7, 2.7),
                arrowprops=dict(arrowstyle="->", color=COLOR_GREEN, lw=2, mutation_scale=12))
    ax.text(8.85, 3.2, "pH 6.0-7.5\n(Neutral)", ha="center", va="center", fontsize=8, color=COLOR_GREEN, fontweight="bold")

    # Bottom arrow: Acidic pH
    ax.annotate("", xy=(9.0, 2.1), xytext=(8.7, 2.4),
                arrowprops=dict(arrowstyle="->", color="#888888", lw=2, mutation_scale=12))
    ax.text(8.85, 1.9, "pH < 4.0\n(Acidic / ESP)", ha="center", va="center", fontsize=8, color="#666666")

    # Box 4A: Sulforaphane (Bioactive)
    b4a = patches.FancyBboxPatch((9.2, 2.6), 2.6, 0.9, boxstyle="round,pad=0.1", fc="#D1E7DD", ec=COLOR_GREEN, lw=2.5)
    ax.add_patch(b4a)
    ax.text(10.5, 3.15, "SULFORAPHANE (ITC)", ha="center", va="center", fontsize=10.5, fontweight="bold", color=COLOR_NAVY)
    ax.text(10.5, 2.85, "Phase II Inducer (Active)", ha="center", va="center", fontsize=8.5, color=COLOR_GREEN, fontweight="bold")

    # Box 4B: Nitriles (Inert)
    b4b = patches.FancyBboxPatch((9.2, 1.5), 2.6, 0.8, boxstyle="round,pad=0.1", fc="#F8D7DA", ec="#DC3545", lw=1.5)
    ax.add_patch(b4b)
    ax.text(10.5, 1.95, "Inert Nitriles & Epithionitriles", ha="center", va="center", fontsize=9.5, fontweight="bold", color="#842029")
    ax.text(10.5, 1.7, "Inactive Bioavailability", ha="center", va="center", fontsize=8, color="#842029")

    # Bottom notice card
    callout = patches.FancyBboxPatch((0.5, 0.3), 11.3, 0.85, boxstyle="round,pad=0.1", fc="#FFFFFF", ec=COLOR_BORDER, lw=1)
    ax.add_patch(callout)
    ax.text(0.7, 0.72, "Key Kinetic Rule:", fontsize=9.5, fontweight="bold", color=COLOR_ACCENT)
    ax.text(2.3, 0.72, "Boiling (>60°C) denatures myrosinase (reducing yield by 85%). Gentle steaming (3-4 min) preserves catalytic activity.",
            fontsize=8.5, color=COLOR_NAVY)

    ax.set_xlim(0, 12.2)
    ax.set_ylim(0, 4.2)
    ax.axis("off")
    plt.tight_layout()
    out_file = ASSETS_DIR / "flowchart_glucosinolate_myrosinase.png"
    plt.savefig(out_file, bbox_inches="tight")
    plt.close()
    print(f"Generated: {out_file.name}")


def generate_nrf2_pathway_diagram():
    """Generates Slide 4 Nrf2-Keap1 Signaling Pathway Diagram."""
    fig, ax = plt.subplots(figsize=(10, 4.2), dpi=250)
    ax.set_facecolor(COLOR_LIGHT_BG)
    fig.patch.set_facecolor(COLOR_LIGHT_BG)

    # Title
    ax.text(5, 3.85, "Molecular Mechanism: Sulforaphane Nrf2-Keap1 Pathway",
            ha="center", va="center", fontsize=13, fontweight="bold", color=COLOR_NAVY, fontfamily=FONT_FAMILY)

    # Step 1: Cytosol Repression
    s1 = patches.FancyBboxPatch((0.4, 1.4), 2.5, 2.0, boxstyle="round,pad=0.1", fc="#FFFFFF", ec=COLOR_BORDER, lw=1.5)
    ax.add_patch(s1)
    ax.text(1.65, 3.15, "1. Basal State", ha="center", va="center", fontsize=10.5, fontweight="bold", color=COLOR_NAVY)
    ax.text(1.65, 2.65, "Keap1 Homodimer\nbinds Nrf2", ha="center", va="center", fontsize=9, color="#333333")
    ax.text(1.65, 1.95, "Continuous Cul3\nUbiquitination ->\nProteasome Degradation", ha="center", va="center", fontsize=8, color="#777777")

    # Arrow 1->2
    ax.annotate("", xy=(3.3, 2.4), xytext=(2.9, 2.4),
                arrowprops=dict(arrowstyle="->", color=COLOR_GREEN, lw=2.5, mutation_scale=14))

    # Step 2: Electrophilic Sensing
    s2 = patches.FancyBboxPatch((3.4, 1.4), 2.6, 2.0, boxstyle="round,pad=0.1", fc="#E8F4EC", ec=COLOR_GREEN, lw=2)
    ax.add_patch(s2)
    ax.text(4.7, 3.15, "2. Cys151 Modification", ha="center", va="center", fontsize=10.5, fontweight="bold", color=COLOR_GREEN)
    ax.text(4.7, 2.6, "Sulforaphane (-N=C=S)\nattacks Keap1 Cys151", ha="center", va="center", fontsize=9, color=COLOR_NAVY, fontweight="bold")
    ax.text(4.7, 1.9, "Conformational change;\nUbiquitination halts.\nNrf2 accumulates rapidly.", ha="center", va="center", fontsize=8, color="#4D6253")

    # Arrow 2->3
    ax.annotate("", xy=(6.4, 2.4), xytext=(6.0, 2.4),
                arrowprops=dict(arrowstyle="->", color=COLOR_GREEN, lw=2.5, mutation_scale=14))

    # Step 3: Nuclear Translocation
    s3 = patches.FancyBboxPatch((6.5, 1.4), 2.5, 2.0, boxstyle="round,pad=0.1", fc="#FFFFFF", ec=COLOR_BORDER, lw=1.5)
    ax.add_patch(s3)
    ax.text(7.75, 3.15, "3. Nuclear Translocation", ha="center", va="center", fontsize=10.5, fontweight="bold", color=COLOR_NAVY)
    ax.text(7.75, 2.6, "Free Nrf2 translocates\nacross Nuclear Pore", ha="center", va="center", fontsize=9, color="#333333")
    ax.text(7.75, 1.9, "Heterodimerizes with\nsmall Maf proteins\n(sMaf)", ha="center", va="center", fontsize=8, color="#555555")

    # Arrow 3->4
    ax.annotate("", xy=(9.4, 2.4), xytext=(9.0, 2.4),
                arrowprops=dict(arrowstyle="->", color=COLOR_GREEN, lw=2.5, mutation_scale=14))

    # Step 4: Phase II Gene Expression
    s4 = patches.FancyBboxPatch((9.5, 1.4), 2.8, 2.0, boxstyle="round,pad=0.1", fc="#D1E7DD", ec=COLOR_GREEN, lw=2.5)
    ax.add_patch(s4)
    ax.text(10.9, 3.15, "4. ARE Gene Induction", ha="center", va="center", fontsize=10.5, fontweight="bold", color=COLOR_NAVY)
    ax.text(10.9, 2.6, "Binds ARE Promoters:\n• NQO1 (Quinone detox)\n• GSTs (Electrophile conj)\n• HO-1 (Anti-inflammatory)", ha="center", va="center", fontsize=8.5, color=COLOR_NAVY)
    ax.text(10.9, 1.75, "Cytoprotection & Cancer Chemoprevention", ha="center", va="center", fontsize=8, fontweight="bold", color=COLOR_GREEN)

    # Footnote bar
    fbar = patches.FancyBboxPatch((0.4, 0.3), 11.9, 0.75, boxstyle="round,pad=0.1", fc="#FFFFFF", ec=COLOR_BORDER, lw=1)
    ax.add_patch(fbar)
    ax.text(0.6, 0.68, "Downstream Result:", fontsize=9, fontweight="bold", color=COLOR_GREEN)
    ax.text(2.3, 0.68, "Over 200 antioxidant and Phase II detoxifying enzymes upregulated, elevating cellular glutathione by >200%.",
            fontsize=8.5, color=COLOR_NAVY)

    ax.set_xlim(0, 12.5)
    ax.set_ylim(0, 4.2)
    ax.axis("off")
    plt.tight_layout()
    out_file = ASSETS_DIR / "flowchart_nrf2_pathway.png"
    plt.savefig(out_file, bbox_inches="tight")
    plt.close()
    print(f"Generated: {out_file.name}")


def generate_urolithin_mitophagy_diagram():
    """Generates Slide 9 Urolithin Mitophagy Pathway Diagram."""
    fig, ax = plt.subplots(figsize=(10, 4.2), dpi=250)
    ax.set_facecolor(COLOR_LIGHT_BG)
    fig.patch.set_facecolor(COLOR_LIGHT_BG)

    ax.text(5, 3.85, "Pomegranate Punicalagin to Urolithin A: Mitophagy Activation Pathway",
            ha="center", va="center", fontsize=12.5, fontweight="bold", color=COLOR_NAVY, fontfamily=FONT_FAMILY)

    # Stage 1: Dietary Precursor
    p1 = patches.FancyBboxPatch((0.4, 1.4), 2.5, 2.0, boxstyle="round,pad=0.1", fc="#FFFFFF", ec=COLOR_BORDER, lw=1.5)
    ax.add_patch(p1)
    ax.text(1.65, 3.15, "1. Punicalagin Intake", ha="center", va="center", fontsize=10.5, fontweight="bold", color=COLOR_NAVY)
    ax.text(1.65, 2.65, "Pomegranate Ellagitannin\n(MW ~1084 g/mol)", ha="center", va="center", fontsize=9, color="#333333")
    ax.text(1.65, 1.95, "Spontaneous hydrolysis\nto Ellagic Acid in gut\n(low native uptake)", ha="center", va="center", fontsize=8, color="#666666")

    # Arrow 1->2
    ax.annotate("", xy=(3.3, 2.4), xytext=(2.9, 2.4),
                arrowprops=dict(arrowstyle="->", color=COLOR_GREEN, lw=2.5, mutation_scale=14))

    # Stage 2: Colonic Biotransformation
    p2 = patches.FancyBboxPatch((3.4, 1.4), 2.6, 2.0, boxstyle="round,pad=0.1", fc="#E8F4EC", ec=COLOR_GREEN, lw=2)
    ax.add_patch(p2)
    ax.text(4.7, 3.15, "2. Gut Microbiome", ha="center", va="center", fontsize=10.5, fontweight="bold", color=COLOR_GREEN)
    ax.text(4.7, 2.65, "Gordonibacter urolithinfaciens\ncleaves lactone rings", ha="center", va="center", fontsize=9, color=COLOR_NAVY, fontweight="bold")
    ax.text(4.7, 1.9, "Produces bioactive\nUROLITHIN A (lipophilic,\nhighly bioavailable)", ha="center", va="center", fontsize=8.5, color=COLOR_GREEN)

    # Arrow 2->3
    ax.annotate("", xy=(6.4, 2.4), xytext=(6.0, 2.4),
                arrowprops=dict(arrowstyle="->", color=COLOR_GREEN, lw=2.5, mutation_scale=14))

    # Stage 3: Mitochondrial Targeting
    p3 = patches.FancyBboxPatch((6.5, 1.4), 2.5, 2.0, boxstyle="round,pad=0.1", fc="#FFFFFF", ec=COLOR_BORDER, lw=1.5)
    ax.add_patch(p3)
    ax.text(7.75, 3.15, "3. Mitochondrial Sensing", ha="center", va="center", fontsize=10.5, fontweight="bold", color=COLOR_NAVY)
    ax.text(7.75, 2.65, "Accumulates in aging,\ndepolarized mitochondria", ha="center", va="center", fontsize=9, color="#333333")
    ax.text(7.75, 1.9, "Stabilizes PINK1 kinase\non outer membrane;\nrecruits Parkin E3 ligase", ha="center", va="center", fontsize=8, color="#555555")

    # Arrow 3->4
    ax.annotate("", xy=(9.4, 2.4), xytext=(9.0, 2.4),
                arrowprops=dict(arrowstyle="->", color=COLOR_GREEN, lw=2.5, mutation_scale=14))

    # Stage 4: Mitophagy & Rejuvenation
    p4 = patches.FancyBboxPatch((9.5, 1.4), 2.8, 2.0, boxstyle="round,pad=0.1", fc="#D1E7DD", ec=COLOR_GREEN, lw=2.5)
    ax.add_patch(p4)
    ax.text(10.9, 3.15, "4. Selective Mitophagy", ha="center", va="center", fontsize=10.5, fontweight="bold", color=COLOR_NAVY)
    ax.text(10.9, 2.6, "Autophagosome engulfs\ndefective mitochondria.\nPurges toxic ROS.", ha="center", va="center", fontsize=8.5, color=COLOR_NAVY)
    ax.text(10.9, 1.75, "Increases Muscle Strength & Cellular ATP (+12% in RCTs)", ha="center", va="center", fontsize=8, fontweight="bold", color=COLOR_GREEN)

    # Bottom stat
    sbar = patches.FancyBboxPatch((0.4, 0.3), 11.9, 0.75, boxstyle="round,pad=0.1", fc="#FFFFFF", ec=COLOR_BORDER, lw=1)
    ax.add_patch(sbar)
    ax.text(0.6, 0.68, "Clinical Impact:", fontsize=9, fontweight="bold", color=COLOR_GREEN)
    ax.text(2.0, 0.68, "Urolithin A extends lifespan and muscular endurance by selectively clearing damaged mitochondria without affecting healthy organelles.",
            fontsize=8.5, color=COLOR_NAVY)

    ax.set_xlim(0, 12.5)
    ax.set_ylim(0, 4.2)
    ax.axis("off")
    plt.tight_layout()
    out_file = ASSETS_DIR / "flowchart_urolithin_mitophagy.png"
    plt.savefig(out_file, bbox_inches="tight")
    plt.close()
    print(f"Generated: {out_file.name}")


def generate_curcumin_piperine_diagram():
    """Generates Slide 13 Curcumin & Piperine Pharmacokinetics Diagram."""
    fig, ax = plt.subplots(figsize=(10, 4.2), dpi=250)
    ax.set_facecolor(COLOR_LIGHT_BG)
    fig.patch.set_facecolor(COLOR_LIGHT_BG)

    ax.text(5.5, 3.85, "Pharmacokinetics: Curcumin Alone vs. Curcumin + Piperine Synergy",
            ha="center", va="center", fontsize=12.5, fontweight="bold", color=COLOR_NAVY, fontfamily=FONT_FAMILY)

    # Panel A: Curcumin Alone
    pa = patches.FancyBboxPatch((0.5, 1.4), 5.0, 2.1, boxstyle="round,pad=0.1", fc="#FFF5F5", ec="#E57373", lw=2)
    ax.add_patch(pa)
    ax.text(3.0, 3.25, "CURCUMIN ALONE (Low Bioavailability)", ha="center", va="center", fontsize=11, fontweight="bold", color="#C62828")
    ax.text(3.0, 2.7, "• Rapid intestinal & hepatic Phase II metabolism\n• Hepatic UDP-glucuronosyltransferase (UGT) glucuronidates >99%\n• High P-glycoprotein efflux back into intestinal lumen\n• Result: Serum free curcumin < 5 ng/mL (Undetectable)",
            ha="center", va="center", fontsize=9, color="#444444")
    ax.text(3.0, 1.7, "Bioavailability: < 1% (Extensively Excreted)", ha="center", va="center", fontsize=9.5, fontweight="bold", color="#C62828")

    # Panel B: Curcumin + Piperine
    pb = patches.FancyBboxPatch((6.0, 1.4), 5.5, 2.1, boxstyle="round,pad=0.1", fc="#E8F4EC", ec=COLOR_GREEN, lw=2.5)
    ax.add_patch(pb)
    ax.text(8.75, 3.25, "CURCUMIN + 20mg PIPERINE (2000% Surge)", ha="center", va="center", fontsize=11, fontweight="bold", color=COLOR_NAVY)
    ax.text(8.75, 2.7, "• Piperine non-competitively inhibits UGT & sulfation enzymes\n• Blocks P-glycoprotein ABCB1 efflux pumps in enterocytes\n• Halts rapid hepatic clearance of intact diferuloylmethane\n• Result: Serum AUC rises by 2000% (20-Fold increase)",
            ha="center", va="center", fontsize=9, color="#222222")
    ax.text(8.75, 1.7, "Bioavailability: 20x Higher Serum Concentration (RCT Proven)", ha="center", va="center", fontsize=9.5, fontweight="bold", color=COLOR_GREEN)

    # Comparison summary box
    cbar = patches.FancyBboxPatch((0.5, 0.3), 11.0, 0.8, boxstyle="round,pad=0.1", fc="#FFFFFF", ec=COLOR_BORDER, lw=1)
    ax.add_patch(cbar)
    ax.text(0.7, 0.7, "Mechanistic Summary:", fontsize=9.5, fontweight="bold", color=COLOR_GREEN)
    ax.text(2.6, 0.7, "Piperine safely pauses hepatic first-pass clearance, enabling curcumin to reach systemic inflammatory targets (IKKβ, COX-2, TNF-α).",
            fontsize=8.5, color=COLOR_NAVY)

    ax.set_xlim(0, 12.0)
    ax.set_ylim(0, 4.2)
    ax.axis("off")
    plt.tight_layout()
    out_file = ASSETS_DIR / "flowchart_curcumin_piperine.png"
    plt.savefig(out_file, bbox_inches="tight")
    plt.close()
    print(f"Generated: {out_file.name}")


def generate_synergistic_infographic():
    """Generates Slide 16 Synergistic Food Pairings Infographic."""
    fig, ax = plt.subplots(figsize=(11, 4.2), dpi=250)
    ax.set_facecolor(COLOR_LIGHT_BG)
    fig.patch.set_facecolor(COLOR_LIGHT_BG)

    ax.text(5.5, 3.85, "Culinary Food Matrix Synergies: Proven Bioavailability Stacking",
            ha="center", va="center", fontsize=13, fontweight="bold", color=COLOR_NAVY, fontfamily=FONT_FAMILY)

    pairings = [
        {
            "pair": "Turmeric + Pepper",
            "comp": "Curcumin + Piperine",
            "action": "Inhibits hepatic UGT;\nincreases curcumin\nbioavailability by 2000%.",
            "color": "#D48F18"
        },
        {
            "pair": "Broccoli + Mustard",
            "comp": "Glucoraphanin + Myrosinase",
            "action": "Raw mustard provides\nmyrosinase to rescue\nheat-treated cooked broccoli.",
            "color": "#2D6A4F"
        },
        {
            "pair": "Green Tea + Citrus",
            "comp": "EGCG + Ascorbic Acid",
            "action": "Citrus citric/ascorbic acid\nstabilizes EGCG against\nalkaline intestinal decay.",
            "color": "#40916C"
        },
        {
            "pair": "Tomato + Olive Oil",
            "comp": "Lycopene + Oleic Acid",
            "action": "Triglycerides promote\ntrans-cis lycopene\nmicellar absorption (+45%).",
            "color": "#D9534F"
        }
    ]

    for idx, p in enumerate(pairings):
        x = 0.5 + idx * 2.7
        b = patches.FancyBboxPatch((x, 1.2), 2.5, 2.3, boxstyle="round,pad=0.1", fc="#FFFFFF", ec=p["color"], lw=2)
        ax.add_patch(b)
        
        ax.text(x + 1.25, 3.2, p["pair"], ha="center", va="center", fontsize=11, fontweight="bold", color=COLOR_NAVY)
        ax.text(x + 1.25, 2.8, p["comp"], ha="center", va="center", fontsize=9, fontweight="bold", color=p["color"])
        ax.plot([x + 0.3, x + 2.2], [2.55, 2.55], color=COLOR_BORDER, lw=1)
        ax.text(x + 1.25, 1.95, p["action"], ha="center", va="center", fontsize=8.5, color="#333333")

    # Footnote bar
    fbar = patches.FancyBboxPatch((0.5, 0.3), 10.6, 0.65, boxstyle="round,pad=0.1", fc="#FFFFFF", ec=COLOR_BORDER, lw=1)
    ax.add_patch(fbar)
    ax.text(0.7, 0.62, "Food Synergy Rule:", fontsize=9, fontweight="bold", color=COLOR_GREEN)
    ax.text(2.2, 0.62, "Whole-food combinations overcome hepatic and digestive absorption bottlenecks far more effectively than isolated single supplements.",
            fontsize=8.5, color=COLOR_NAVY)

    ax.set_xlim(0, 11.6)
    ax.set_ylim(0, 4.2)
    ax.axis("off")
    plt.tight_layout()
    out_file = ASSETS_DIR / "infographic_synergistic_pairings.png"
    plt.savefig(out_file, bbox_inches="tight")
    plt.close()
    print(f"Generated: {out_file.name}")


def main():
    print("Generating 5 scientific flowcharts and infographics...")
    generate_glucosinolate_flowchart()
    generate_nrf2_pathway_diagram()
    generate_urolithin_mitophagy_diagram()
    generate_curcumin_piperine_diagram()
    generate_synergistic_infographic()
    print("All infographics generated successfully!")


if __name__ == "__main__":
    main()
