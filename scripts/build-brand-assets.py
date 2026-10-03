"""Render the canonical, deliberately simple SVG using the existing Pillow dependency.

Supports this icon's rect, one cubic path and lines only. Fail on unexpected
geometry instead of silently producing mismatched Windows assets.
Run: python scripts/build-brand-assets.py
"""
from pathlib import Path
import re
import xml.etree.ElementTree as ET
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
BRAND = ROOT / "progress_studio/assets/brand"


def build():
    svg = ET.parse(BRAND / "progress_studio.svg").getroot()
    scale = 4
    image = Image.new("RGBA", (1024, 1024))
    draw = ImageDraw.Draw(image)
    def point(x, y): return (float(x) * scale, float(y) * scale)
    def stroke(points, color, width):
        width = int(float(width) * scale)
        draw.line(points, fill=color, width=width, joint="curve")
        radius = width / 2
        # Fill all sampled joins; Pillow's wide subpixel polyline may leave
        # tiny interior gaps when only endpoint caps are drawn.
        for x, y in points:
            draw.ellipse((x-radius, y-radius, x+radius, y+radius), fill=color)
    for element in svg:
        tag = element.tag.rsplit("}", 1)[-1]
        a = element.attrib
        if tag == "title": continue
        if tag == "rect":
            draw.rounded_rectangle((0, 0, 1023, 1023), radius=float(a["rx"])*scale, fill=a["fill"])
        elif tag == "path":
            if not re.fullmatch(r"M [\d. ]+ C [\d. ]+", a["d"]):
                raise ValueError("Unsupported icon path")
            numbers = list(map(float, re.findall(r"[\d.]+", a["d"])))
            if len(numbers) != 8: raise ValueError("Expected one cubic curve")
            control = [point(*numbers[i:i+2]) for i in range(0,8,2)]
            points = []
            for i in range(301):
                t = i/300; u = 1-t
                points.append(tuple(u**3*control[0][j]+3*u*u*t*control[1][j]+3*u*t*t*control[2][j]+t**3*control[3][j] for j in (0,1)))
            stroke(points, a["stroke"], a["stroke-width"])
        elif tag == "line":
            stroke([point(a['x1'],a['y1']),point(a['x2'],a['y2'])],a['stroke'],a['stroke-width'])
        else: raise ValueError(f"Unsupported icon element: {tag}")
    for name in ("progress_studio_icon.png", "progress_studio_brand.png"):
        image.resize((512,512), Image.Resampling.LANCZOS).save(BRAND / name, optimize=True)
    image.save(BRAND / "progress_studio.ico", sizes=[(x,x) for x in (16,20,24,32,40,48,64,128,256)])


if __name__ == "__main__": build()
