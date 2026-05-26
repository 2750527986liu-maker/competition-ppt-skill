#!/usr/bin/env python3
"""
Generate original design assets for 助你好梦 PPT
Using PIL to create gradient backgrounds, patterns, and design elements
"""
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import math, random, os

OUT = "/sessions/zealous-eager-fermat/mnt/被子/ppt-assets"
os.makedirs(OUT, exist_ok=True)

# PPT dimensions: 13.33 x 7.5 inches at 150 DPI = 2000x1125 px
W, H = 2000, 1125

def gradient(draw, w, h, c1, c2, direction='vertical'):
    """Create a gradient between two RGB colors"""
    for i in range(h if direction == 'vertical' else w):
        ratio = i / (h if direction == 'vertical' else w)
        r = int(c1[0] + (c2[0] - c1[0]) * ratio)
        g = int(c1[1] + (c2[1] - c1[1]) * ratio)
        b = int(c1[2] + (c2[2] - c1[2]) * ratio)
        if direction == 'vertical':
            draw.line([(0, i), (w, i)], fill=(r, g, b))
        else:
            draw.line([(i, 0), (i, h)], fill=(r, g, b))

def radial_gradient(img, cx, cy, radius, center_color, edge_color):
    """Create a radial gradient overlay"""
    pixels = img.load()
    for y in range(img.height):
        for x in range(img.width):
            dist = math.sqrt((x - cx)**2 + (y - cy)**2)
            ratio = min(dist / radius, 1.0)
            r = int(center_color[0] + (edge_color[0] - center_color[0]) * ratio)
            g = int(center_color[1] + (edge_color[1] - center_color[1]) * ratio)
            b = int(center_color[2] + (edge_color[2] - center_color[2]) * ratio)
            pixels[x, y] = (r, g, b)

# ========================================
# 1. COVER BACKGROUND - Deep gradient with abstract geometric patterns
# ========================================
print("Creating cover background...")
img = Image.new('RGB', (W, H))
draw = ImageDraw.Draw(img)

# Multi-point gradient: deep navy center-top to darker edges
gradient(draw, W, H, (15, 25, 55), (5, 10, 25), 'vertical')

# Add subtle radial glow at top-right
overlay = Image.new('RGBA', (W, H), (0, 0, 0, 0))
odraw = ImageDraw.Draw(overlay)
for r in range(800, 0, -2):
    alpha = int(25 * (1 - r/800))
    odraw.ellipse([W-600-r, -200-r, W-600+r, -200+r], fill=(70, 130, 220, alpha))

# Add another glow at bottom-left
for r in range(600, 0, -2):
    alpha = int(20 * (1 - r/600))
    odraw.ellipse([100-r, H-100-r, 100+r, H-100+r], fill=(180, 160, 100, alpha))

img = Image.alpha_composite(img.convert('RGBA'), overlay).convert('RGB')

# Add abstract geometric lines (like circuit/network pattern)
draw = ImageDraw.Draw(img)
random.seed(42)
points = [(random.randint(0, W), random.randint(0, H)) for _ in range(30)]
for i in range(len(points)):
    for j in range(i+1, len(points)):
        dist = math.sqrt((points[i][0]-points[j][0])**2 + (points[i][1]-points[j][1])**2)
        if dist < 500:
            alpha_val = int(40 * (1 - dist/500))
            draw.line([points[i], points[j]], fill=(100+alpha_val, 140+alpha_val, 200+alpha_val), width=1)

# Add small dots at vertices
for p in points:
    draw.ellipse([p[0]-3, p[1]-3, p[0]+3, p[1]+3], fill=(180, 200, 240))

img.save(os.path.join(OUT, "cover-bg.png"))
print("  cover-bg.png done")

# ========================================
# 2. CONTENT SLIDE BACKGROUND - Light with subtle texture
# ========================================
print("Creating content background...")
img = Image.new('RGB', (W, H), (248, 250, 253))
draw = ImageDraw.Draw(img)

# Subtle gradient from top-left (slightly blue) to bottom-right (warm white)
for y in range(H):
    for x in range(W):
        ratio_x = x / W
        ratio_y = y / H
        r = int(248 - 3 * ratio_x - 2 * ratio_y)
        g = int(250 - 2 * ratio_x - 1 * ratio_y)
        b = int(253 + 2 * ratio_x)
        draw.point((x, y), fill=(r, g, b))

# Add subtle geometric grid pattern in very light blue
for x in range(0, W, 80):
    draw.line([(x, 0), (x, H)], fill=(230, 237, 248), width=1)
for y in range(0, H, 80):
    draw.line([(0, y), (W, y)], fill=(230, 237, 248), width=1)

# Bottom-left accent: abstract blue watercolor splash
overlay = Image.new('RGBA', (W, H), (0, 0, 0, 0))
odraw = ImageDraw.Draw(overlay)
for r in range(400, 0, -1):
    alpha = int(15 * (1 - r/400))
    odraw.ellipse([-100-r, H-50-r, -100+r, H-50+r], fill=(90, 150, 220, alpha))

# Bottom-right accent
for r in range(350, 0, -1):
    alpha = int(12 * (1 - r/350))
    odraw.ellipse([W-80-r, H-30-r, W-80+r, H-30+r], fill=(160, 180, 220, alpha))

img = Image.alpha_composite(img.convert('RGBA'), overlay).convert('RGB')
img.save(os.path.join(OUT, "content-bg.png"))
print("  content-bg.png done")

# ========================================
# 3. SECTION BADGE (PNG with transparency)
# ========================================
print("Creating section badge...")
badge_w, badge_h = 400, 80
badge = Image.new('RGBA', (badge_w, badge_h), (0, 0, 0, 0))
bdraw = ImageDraw.Draw(badge)
# Rounded rectangle
bdraw.rounded_rectangle([0, 0, badge_w-1, badge_h-1], radius=15, fill=(10, 22, 40))
# Gold accent line at bottom
bdraw.rectangle([15, badge_h-6, badge_w-15, badge_h-2], fill=(201, 168, 76))
badge.save(os.path.join(OUT, "section-badge.png"))
print("  section-badge.png done")

# ========================================
# 4. GOLD ACCENT LINE
# ========================================
print("Creating gold accent...")
gold = Image.new('RGBA', (600, 6), (0, 0, 0, 0))
gdraw = ImageDraw.Draw(gold)
# Gold gradient line
for x in range(600):
    ratio = x / 600
    r = int(201 + (232 - 201) * ratio)
    g = int(168 + (213 - 168) * ratio)
    b = int(76 + (140 - 76) * ratio)
    gdraw.line([(x, 0), (x, 5)], fill=(r, g, b, 255))
gold.save(os.path.join(OUT, "gold-line.png"))
print("  gold-line.png done")

# ========================================
# 5. THANK YOU SLIDE BACKGROUND
# ========================================
print("Creating thank-you background...")
img = Image.new('RGB', (W, H))
draw = ImageDraw.Draw(img)
gradient(draw, W, H, (8, 18, 38), (20, 35, 65), 'vertical')

# Add golden radial glow at center
overlay = Image.new('RGBA', (W, H), (0, 0, 0, 0))
odraw = ImageDraw.Draw(overlay)
for r in range(600, 0, -2):
    alpha = int(18 * (1 - r/600))
    odraw.ellipse([W//2-r, H//2-r, W//2+r, H//2+r], fill=(201, 168, 76, alpha))

# Add decorative corner lines
for corner_x, corner_y in [(0, 0), (W, 0), (0, H), (W, H)]:
    dx = 1 if corner_x == 0 else -1
    dy = 1 if corner_y == 0 else -1
    for i in range(200):
        alpha = int(30 * (1 - i/200))
        x1 = corner_x + dx * i
        y1 = corner_y
        x2 = corner_x
        y2 = corner_y + dy * i
        odraw.line([(x1, y1), (x2, y2)], fill=(201, 168, 76, alpha), width=1)

img = Image.alpha_composite(img.convert('RGBA'), overlay).convert('RGB')
img.save(os.path.join(OUT, "thankyou-bg.png"))
print("  thankyou-bg.png done")

# ========================================
# 6. CARD BACKGROUND (white with subtle shadow)
# ========================================
print("Creating card...")
card = Image.new('RGBA', (700, 500), (0, 0, 0, 0))
cdraw = ImageDraw.Draw(card)
# Shadow
cdraw.rounded_rectangle([4, 4, 698, 498], radius=12, fill=(200, 210, 225, 40))
# White card
cdraw.rounded_rectangle([0, 0, 694, 494], radius=12, fill=(255, 255, 255, 245))
# Subtle border
cdraw.rounded_rectangle([0, 0, 694, 494], radius=12, outline=(220, 228, 240), width=1)
card.save(os.path.join(OUT, "card-bg.png"))
print("  card-bg.png done")

# ========================================
# 7. TIMELINE SPECIAL BACKGROUND (paint splatter effect)
# ========================================
print("Creating timeline background...")
img = Image.new('RGB', (W, H), (240, 244, 250))
draw = ImageDraw.Draw(img)

# Add blue paint splatters
random.seed(88)
for _ in range(8):
    cx = random.randint(200, W-200)
    cy = random.randint(200, H-200)
    radius = random.randint(100, 300)
    overlay = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    odraw = ImageDraw.Draw(overlay)
    blue_shade = random.choice([(70, 130, 200), (50, 100, 170), (90, 150, 220), (40, 80, 150)])
    for r in range(radius, 0, -1):
        alpha = int(20 * (1 - r/radius))
        odraw.ellipse([cx-r, cy-r, cx+r, cy+r], fill=(*blue_shade, alpha))
    img = Image.alpha_composite(img.convert('RGBA'), overlay).convert('RGB')

img.save(os.path.join(OUT, "timeline-bg.png"))
print("  timeline-bg.png done")

# ========================================
# 8. DARK SECTION BACKGROUND (for emphasis slides)
# ========================================
print("Creating dark section background...")
img = Image.new('RGB', (W, H))
draw = ImageDraw.Draw(img)
gradient(draw, W, H, (10, 20, 45), (15, 30, 60), 'vertical')

# Subtle diagonal pattern
for i in range(-H, W+H, 60):
    draw.line([(i, 0), (i+H, H)], fill=(20, 35, 65), width=1)

# Gold corner accents
for cx, cy in [(100, 100), (W-100, 100), (100, H-100), (W-100, H-100)]:
    for r in range(80, 0, -1):
        alpha_ratio = 1 - r/80
        c = int(40 + 160 * alpha_ratio)
        draw.ellipse([cx-r, cy-r, cx+r, cy+r], outline=(c, int(c*0.8), int(c*0.4)))

img.save(os.path.join(OUT, "dark-section-bg.png"))
print("  dark-section-bg.png done")

# ========================================
# 9. ICON: sleep/crescent moon
# ========================================
print("Creating moon icon...")
icon = Image.new('RGBA', (200, 200), (0, 0, 0, 0))
idraw = ImageDraw.Draw(icon)
# Crescent moon
idraw.ellipse([30, 20, 170, 180], fill=(201, 168, 76))
idraw.ellipse([60, 10, 200, 170], fill=(10, 22, 40))
# Small stars
for sx, sy in [(160, 40), (175, 70), (150, 90)]:
    idraw.ellipse([sx-4, sy-4, sx+4, sy+4], fill=(201, 168, 76))
icon.save(os.path.join(OUT, "moon-icon.png"))
print("  moon-icon.png done")

print("\nAll assets generated!")
print(f"Output directory: {OUT}")
