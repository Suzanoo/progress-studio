"""Small monochrome command images drawn with existing Pillow; no asset files.

Text remains the command identity. Images are cached on the owning frame for
Tk lifetime safety. If image creation is unavailable, buttons remain text-only.
"""
import tkinter as tk


def _make_icon(master, name, light=False):
    from PIL import Image, ImageDraw, ImageTk
    size = max(16, round(master.winfo_fpixels('1i') / 96 * 16))
    scale = 4
    im = Image.new('RGBA', (64, 64))
    draw = ImageDraw.Draw(im)
    color = '#ffffff' if light else '#34434b'
    def line(points):
        draw.line([(x*scale, y*scale) for x, y in points], fill=color, width=5, joint='curve')
    if name in ('map', 'unmap'):
        draw.rounded_rectangle((4, 20, 36, 44), radius=10, outline=color, width=5)
        draw.rounded_rectangle((28, 20, 60, 44), radius=10, outline=color, width=5)
        if name == 'unmap':
            line([(5, 2), (11, 14)])
    elif name == 'undo':
        draw.arc((12, 12, 56, 56), 220, 90, fill=color, width=5)
        line([(2, 2), (2, 7), (7, 7)])
    elif name == 'clear':
        line([(2, 10), (8, 3), (14, 8), (9, 14), (6, 14), (2, 10)])
        line([(6, 7), (11, 12)])
    elif name == 'refresh':
        draw.ellipse((8, 8, 56, 28), outline=color, width=5)
        line([(2, 5), (2, 12), (7, 14), (14, 12), (14, 5)])
        line([(7, 9), (12, 9), (12, 6)])
    else:
        line([(3, 1), (10, 1), (14, 5), (14, 15), (3, 15), (3, 1)])
        line([(10, 1), (10, 5), (14, 5)])
        for y in (8, 11):
            line([(5, y), (12, y)])
        if name == 'sheet':
            line([(8, 7), (8, 13)])
    return ImageTk.PhotoImage(im.resize((size, size), Image.Resampling.LANCZOS), master=master)


def icon_options(master, name, light=False):
    cache = getattr(master, '_command_icons', None)
    if cache is None:
        cache = master._command_icons = {}
    key = (name, light)
    if key not in cache:
        try:
            cache[key] = _make_icon(master, name, light)
        except (ImportError, OSError, tk.TclError):
            cache[key] = None
    return {'image': cache[key], 'compound': 'left'} if cache[key] is not None else {}
