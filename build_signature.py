#!/usr/bin/env python3
"""Generate email signature with hand-drawn icons."""
from PIL import Image, ImageDraw, ImageFont

img = Image.new('RGB', (600, 240), 'white')
draw = ImageDraw.Draw(img)

try:
    font_bold = ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf', 17)
    font_normal = ImageFont.truetype('C:/Windows/Fonts/arial.ttf', 13)
    font_small = ImageFont.truetype('C:/Windows/Fonts/arial.ttf', 11)
    font_title = ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf', 19)
except:
    font_bold = font_normal = font_small = font_title = ImageFont.load_default()

x = 25
y = 18

# Name + company
draw.text((x, y), 'Kujtim Bajrami', fill=(26, 58, 92), font=font_title)
y += 24
draw.text((x, y), 'Safe OIBT Control Sàrl', fill=(46, 125, 50), font=font_bold)
y += 22
draw.line([(x, y), (x + 350, y)], fill=(46, 125, 50), width=1)
y += 14

# Use emoji as icons - render them as images
def globe(d, cx, cy, s=14):
    draw_emoji_icon(d, cx, cy, "\U0001F310", s)

def envelope(d, cx, cy, s=14):
    draw_emoji_icon(d, cx, cy, "\U0001F4E8", s)

def phone(d, cx, cy, s=14):
    draw_emoji_icon(d, cx, cy, "\U0001F4DE", s)

def house(d, cx, cy, s=14):
    draw_emoji_icon(d, cx, cy, "\U0001F3E0", s)

def bldg(d, cx, cy, s=14):
    draw_emoji_icon(d, cx, cy, "\U0001F3E2", s)

def draw_emoji_icon(d, cx, cy, emoji, size=14):
    try:
        seg_font = ImageFont.truetype('C:/Windows/Fonts/seguiemj.ttf', size)
    except:
        try:
            seg_font = ImageFont.truetype('C:/Windows/Fonts/segoeui.ttf', size)
        except:
            seg_font = ImageFont.load_default()
    bbox = d.textbbox((0, 0), emoji, font=seg_font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    d.text((cx - tw//2, cy - th//2), emoji, font=seg_font, embedded_color=True)

icon_cx = x + 8
text_x = x + 22

# Web
globe(draw, icon_cx, y + 7)
draw.text((text_x, y), 'www.safeoibtcontrol.com', fill=(60, 60, 60), font=font_normal)
y += 19

# Email
envelope(draw, icon_cx, y + 7)
draw.text((text_x, y), 'info@safeoibtcontrol.com', fill=(60, 60, 60), font=font_normal)
y += 19

# Phone
phone(draw, icon_cx, y + 7)
draw.text((text_x, y), '+41 (0)78 641 00 17', fill=(60, 60, 60), font=font_normal)
y += 22

# Siege
house(draw, icon_cx, y + 5)
draw.text((text_x, y - 2), 'Siege :', fill=(26, 58, 92), font=font_small)
y += 11
draw.text((text_x, y - 2), 'Chem. des Pierrettes 3, 1844 Villeneuve', fill=(80, 80, 80), font=font_small)
y += 14

# Succursale
bldg(draw, icon_cx, y + 5)
draw.text((text_x, y - 2), 'Succursale :', fill=(26, 58, 92), font=font_small)
y += 11
draw.text((text_x, y - 2), 'Chem. de la Fine Goutte 4, 1273 Arzier-le Muids', fill=(80, 80, 80), font=font_small)
y += 16

draw.line([(x, y), (x + 550, y)], fill=(220, 220, 220), width=1)
y += 6

for line in ['Autorisation : K-371674-1  |  CHE-442.583.093', 'RC : 10 Mio  |  UBS CH75 0028 7287 1436 1601 H']:
    draw.text((x, y), line, fill=(100, 100, 100), font=font_small)
    y += 13

output = r'C:\Users\openc\.openclaw\workspace838360131817\signature_safe_oibt.png'
img.save(output, quality=95)
print(f'Done {img.width}x{img.height}')
