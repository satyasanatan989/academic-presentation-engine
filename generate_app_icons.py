import os
import math
from PIL import Image, ImageDraw, ImageFilter

def create_scholardeck_icon(size=512, is_round=False):
    # Scale factor for crisp super-sampling
    scale = 2
    W = size * scale
    H = size * scale
    
    # 1. Base Dark Canvas
    corner_radius = (W // 2) if is_round else int(W * 0.22)
    
    bg = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    bg_draw = ImageDraw.Draw(bg)
    
    if is_round:
        bg_draw.ellipse([0, 0, W, H], fill=(8, 12, 22, 255))
    else:
        bg_draw.rounded_rectangle([0, 0, W, H], radius=corner_radius, fill=(8, 12, 22, 255))
    
    # Inner Radial Glow (Deep Indigo / Violet Core)
    glow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    glow_draw = ImageDraw.Draw(glow)
    cx, cy = int(W * 0.48), int(H * 0.48)
    max_r = int(W * 0.60)
    for r in range(max_r, 0, -max(2, int(W * 0.015))):
        alpha = int(45 * (1 - r / max_r) ** 1.8)
        glow_draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(79, 70, 229, alpha))
    
    # Secondary Cyan ambient glow at bottom right
    cx2, cy2 = int(W * 0.75), int(H * 0.75)
    r2_max = int(W * 0.40)
    for r in range(r2_max, 0, -max(2, int(W * 0.015))):
        alpha = int(35 * (1 - r / r2_max) ** 1.6)
        glow_draw.ellipse([cx2 - r, cy2 - r, cx2 + r, cy2 + r], fill=(14, 165, 233, alpha))

    # Mask glow to icon boundary
    glow_masked = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    glow_masked.paste(glow, (0, 0), bg)
    img = Image.alpha_composite(bg, glow_masked)
    draw = ImageDraw.Draw(img)

    # 1px subtle inner border
    border_w = max(1, int(W * 0.006))
    if is_round:
        draw.ellipse([border_w, border_w, W - border_w, H - border_w], outline=(255, 255, 255, 35), width=border_w)
    else:
        draw.rounded_rectangle([border_w, border_w, W - border_w, H - border_w], radius=corner_radius, outline=(255, 255, 255, 35), width=border_w)

    # 2. Layered Presentation Deck Cards (3D Angled)
    # Card 3 (Bottom Layer)
    c3_pts = [
        (int(W * 0.18), int(H * 0.42)),
        (int(W * 0.68), int(H * 0.28)),
        (int(W * 0.82), int(H * 0.58)),
        (int(W * 0.32), int(H * 0.72))
    ]
    draw.polygon(c3_pts, fill=(30, 41, 59, 170), outline=(99, 102, 241, 100))

    # Card 2 (Middle Layer - Vibrant Indigo/Cyan)
    c2_pts = [
        (int(W * 0.17), int(H * 0.47)),
        (int(W * 0.70), int(H * 0.33)),
        (int(W * 0.84), int(H * 0.63)),
        (int(W * 0.31), int(H * 0.77))
    ]
    draw.polygon(c2_pts, fill=(37, 99, 235, 190), outline=(56, 189, 248, 140))

    # Card 1 (Hero Front Slide - 16:9 Presentation Format)
    card_x1 = int(W * 0.20)
    card_y1 = int(H * 0.30)
    card_x2 = int(W * 0.80)
    card_y2 = int(H * 0.68)
    card_r = max(4, int(W * 0.035))

    # Drop shadow under hero slide
    shadow_pad = int(W * 0.04)
    card_shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    s_draw = ImageDraw.Draw(card_shadow)
    s_draw.rounded_rectangle([card_x1, card_y1 + shadow_pad, card_x2, card_y2 + shadow_pad * 2], radius=card_r, fill=(0, 0, 0, 180))
    card_shadow = card_shadow.filter(ImageFilter.GaussianBlur(max(2, int(W * 0.025))))
    img = Image.alpha_composite(img, card_shadow)
    draw = ImageDraw.Draw(img)

    # Hero Slide Background
    hero_card = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    hc_draw = ImageDraw.Draw(hero_card)
    hc_draw.rounded_rectangle([card_x1, card_y1, card_x2, card_y2], radius=card_r, fill=(15, 23, 42, 245), outline=(99, 102, 241, 230), width=max(1, int(W * 0.008)))
    
    # Slide Header Accent Line (Gold / Violet)
    hdr_y1 = card_y1 + int((card_y2 - card_y1) * 0.10)
    hdr_y2 = hdr_y1 + max(2, int((card_y2 - card_y1) * 0.07))
    hdr_x1 = card_x1 + int((card_x2 - card_x1) * 0.08)
    hdr_x2 = card_x2 - int((card_x2 - card_x1) * 0.08)
    hc_draw.rounded_rectangle([hdr_x1, hdr_y1, hdr_x2, hdr_y2], radius=max(1, int(W * 0.008)), fill=(99, 102, 241, 255))
    
    # 3 Symmetrical Miniature Content Pillar Columns
    content_y1 = hdr_y2 + int((card_y2 - card_y1) * 0.10)
    content_y2 = card_y2 - int((card_y2 - card_y1) * 0.12)
    total_content_w = hdr_x2 - hdr_x1
    col_gap = int(total_content_w * 0.06)
    col_w = int((total_content_w - (2 * col_gap)) / 3)

    for i in range(3):
        cx1 = hdr_x1 + i * (col_w + col_gap)
        cx2 = cx1 + col_w
        if cx2 > cx1 and content_y2 > content_y1:
            fill_col = (30, 58, 138, 220) if i == 0 else ((15, 118, 110, 220) if i == 1 else (136, 19, 55, 220))
            border_col = (96, 165, 250, 190) if i == 0 else ((52, 211, 153, 190) if i == 1 else (244, 63, 94, 190))
            hc_draw.rounded_rectangle([cx1, content_y1, cx2, content_y2], radius=max(2, int(W * 0.01)), fill=fill_col, outline=border_col, width=max(1, int(W * 0.004)))
            
            # Mini text lines inside columns
            line_w = max(1, int(W * 0.004))
            line_y1 = content_y1 + int((content_y2 - content_y1) * 0.25)
            line_y2 = content_y1 + int((content_y2 - content_y1) * 0.50)
            hc_draw.line([cx1 + int(col_w * 0.15), line_y1, cx2 - int(col_w * 0.15), line_y1], fill=(255, 255, 255, 200), width=line_w)
            hc_draw.line([cx1 + int(col_w * 0.15), line_y2, cx2 - int(col_w * 0.30), line_y2], fill=(255, 255, 255, 140), width=line_w)

    img = Image.alpha_composite(img, hero_card)
    draw = ImageDraw.Draw(img)

    # 3. Brilliant 4-Point Golden AI Prism / Sparkle at Top Right
    sx, sy = int(W * 0.74), int(H * 0.26)
    sparkle_r = int(W * 0.15)

    # Ambient golden bloom
    star_glow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    sg_draw = ImageDraw.Draw(star_glow)
    for r in range(int(sparkle_r * 1.6), 0, -max(2, int(W * 0.015))):
        alpha = int(75 * (1 - r / (sparkle_r * 1.6)) ** 1.8)
        sg_draw.ellipse([sx - r, sy - r, sx + r, sy + r], fill=(245, 158, 11, alpha))
    star_glow = star_glow.filter(ImageFilter.GaussianBlur(max(2, int(W * 0.02))))
    img = Image.alpha_composite(img, star_glow)
    draw = ImageDraw.Draw(img)

    # 4-point Diamond Star Astroid
    star_pts = []
    points_count = 64
    for step in range(points_count):
        theta = step * 2 * math.pi / points_count
        cos_t = math.cos(theta)
        sin_t = math.sin(theta)
        rad = sparkle_r * (abs(cos_t)**4 + abs(sin_t)**4) ** (-0.5) * 0.38
        spike = (abs(cos_t)**12 + abs(sin_t)**12) ** 0.5
        r_actual = rad * 0.28 + (sparkle_r * 0.72) * spike
        px = sx + r_actual * cos_t
        py = sy + r_actual * sin_t
        star_pts.append((px, py))

    draw.polygon(star_pts, fill=(254, 240, 138, 255), outline=(245, 158, 11, 255))
    
    # Core white highlight
    core_r = max(2, int(W * 0.012))
    draw.ellipse([sx - core_r, sy - core_r, sx + core_r, sy + core_r], fill=(255, 255, 255, 255))

    # Downscale with Lanczos
    return img.resize((size, size), Image.Resampling.LANCZOS)

res_dir = r"android-app\app\src\main\res"
sizes = {
    "mipmap-mdpi": 48,
    "mipmap-hdpi": 72,
    "mipmap-xhdpi": 96,
    "mipmap-xxhdpi": 144,
    "mipmap-xxxhdpi": 192
}

for folder, s in sizes.items():
    fpath = os.path.join(res_dir, folder)
    os.makedirs(fpath, exist_ok=True)
    
    # Square / Adaptive
    icon_sq = create_scholardeck_icon(s, is_round=False)
    icon_sq.save(os.path.join(fpath, "ic_launcher.png"), "PNG")
    
    # Round
    icon_rd = create_scholardeck_icon(s, is_round=True)
    icon_rd.save(os.path.join(fpath, "ic_launcher_round.png"), "PNG")
    print(f"Generated {folder} ({s}x{s})")

os.makedirs("outputs", exist_ok=True)
master_icon = create_scholardeck_icon(512, is_round=False)
master_icon.save(r"outputs\scholardeck_app_icon_512.png", "PNG")
master_icon.save(r"samples\scholardeck_app_icon_512.png", "PNG")
print("Saved 512x512 master icon!")
