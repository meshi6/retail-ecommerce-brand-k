"""Generate Art Deco sunburst motif PNGs for the Kiabi x Grupo One template.

The motif: a quarter-circle fan of alternating rays with stepped arc bands —
echoing Buenos Aires Art Deco (Kavanagh / Sol de Mayo) and Parisian Deco ironwork.
"""
import math
from PIL import Image, ImageDraw

# Kiabi palette (from brand book p.20)
NAVY = (4, 0, 55)
ORANGE = (251, 91, 0)
RED = (233, 27, 46)
YELLOW = (255, 218, 0)
GREEN = (32, 177, 74)
BLUE = (0, 110, 251)

SS = 4  # supersampling factor


def fan(size, colors_alpha, n_rays=14, gap_deg=2.2, rings=3, ring_gap_frac=0.045):
    """Quarter fan anchored at the bottom-right corner of a transparent square.

    colors_alpha: list of (rgb, alpha) cycled across rays.
    Rays sweep from 180deg (left) to 270deg (up) around corner (W, H).
    Stepped arc bands cut transparent gaps for the Deco look.
    """
    S = size * SS
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    cx, cy = S, S
    R = S * 1.02
    span = 90.0 / n_rays
    for i in range(n_rays):
        a0 = 180 + i * span + gap_deg / 2
        a1 = 180 + (i + 1) * span - gap_deg / 2
        rgb, alpha = colors_alpha[i % len(colors_alpha)]
        d.pieslice([cx - R, cy - R, cx + R, cy + R], a0, a1, fill=rgb + (alpha,))
    # stepped transparent ring gaps
    out = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    mask = Image.new("L", (S, S), 255)
    md = ImageDraw.Draw(mask)
    for k in range(1, rings + 1):
        r = R * k / (rings + 1)
        gw = R * ring_gap_frac
        md.ellipse([cx - r - gw, cy - r - gw, cx + r + gw, cy + r + gw], fill=0)
        md.ellipse([cx - r + gw, cy - r + gw, cx + r - gw, cy + r - gw], fill=255)
    # inner hub kept solid
    hub = R * 0.16
    md.ellipse([cx - hub, cy - hub, cx + hub, cy + hub], fill=255)
    out = Image.composite(im, out, mask)
    return out.resize((size, size), Image.LANCZOS)


def glyph(size, rgb, n_rays=5):
    """Small title-marker fan: quarter fan anchored bottom-left, solid accent
    alternating with a lighter tint of the same accent."""
    S = size * SS
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    cx, cy = 0, S
    R = S * 0.98
    span = 90.0 / n_rays
    light = tuple(int(c * 0.45 + 255 * 0.55) for c in rgb)
    for i in range(n_rays):
        a0 = 270 + i * span + 1.5
        a1 = 270 + (i + 1) * span - 1.5
        col = rgb if i % 2 == 0 else light
        d.pieslice([cx - R, cy - R, cx + R, cy + R], a0, a1, fill=col + (255,))
    return im.resize((size, size), Image.LANCZOS)


# --- Dark-slide corner fan (cover / closing): blue + orange + yellow glow ---
dark_fan = fan(
    1400,
    [(BLUE, 120), (ORANGE, 130), (BLUE, 60), (YELLOW, 165)],
    n_rays=16,
)
dark_fan.save("build/fan_dark.png")

# --- Light-slide corner fan: very quiet navy/blue tints ---
light_fan = fan(
    1000,
    [(NAVY, 26), (BLUE, 22), (NAVY, 14)],
    n_rays=14,
)
# light slides use it at the top-right corner, rays pointing down-left
light_fan.transpose(Image.FLIP_TOP_BOTTOM).save("build/fan_light.png")

# --- Title glyphs, one per accent ---
for name, rgb in [
    ("blue", BLUE), ("orange", ORANGE), ("red", RED),
    ("green", GREEN), ("yellow", YELLOW), ("navy", NAVY),
]:
    glyph(120, rgb).save(f"build/glyph_{name}.png")

print("motifs written")
