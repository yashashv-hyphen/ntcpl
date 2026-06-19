"""
Amazon 4-shot product image generator.

Produces four Amazon-compliant image types per product:
  main        — pure-white-background compliance shot
  lifestyle   — in-use editorial scene
  macro       — extreme close-up of key feature
  styled_hero — premium branded backdrop shot

Providers:
  fal         — fal.ai (flux/dev), text-to-image, ~$0.003/image (default)
  a1111       — local AUTOMATIC1111 / ComfyUI WebUI server (free)
  pil         — PIL compositor fallback, always available

Usage:
    from amazon_product_images import generateProductImages
    from pathlib import Path

    product = {
        "sku": "kurti-001",
        "product_type": "Anarkali kurti",
        "category": "Apparel",
        "material": "Cotton",
        "colors": ["white", "blue"],
        "features": ["V-neckline", "flared skirt"],
        "usage": "festive wear",
        "style": "ethnic",
    }
    images = generateProductImages(
        product=product,
        reference_image="https://example.com/product.jpg",
        config={"shots": ("main", "lifestyle"), "provider": "fal"},
    )
    # images == {"main": Path("/abs/.../kurti-001_main.png"), "lifestyle": Path(...)}
"""

from __future__ import annotations

import json
import logging
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from dotenv import load_dotenv

load_dotenv()

log = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────

FAL_API_KEY    = os.getenv("FAL_KEY", "").strip()
FAL_IMAGE_MODEL = os.getenv("FAL_IMAGE_MODEL", "fal-ai/flux/dev").strip()
FAL_AVAILABLE  = bool(FAL_API_KEY)

# AUTOMATIC1111 local server (Windows host reachable from WSL2)
_WSL2_HOST = "localhost"
try:
    import subprocess as _sp
    _wsl_ns = _sp.run(
        ["bash", "-c", "cat /etc/resolv.conf | grep nameserver | awk '{print $2}'"],
        capture_output=True, text=True, timeout=2,
    )
    _candidate = _wsl_ns.stdout.strip().split("\n")[0]
    if _candidate:
        _WSL2_HOST = _candidate
except Exception:
    pass
A1111_HOST  = os.getenv("A1111_HOST", _WSL2_HOST)
A1111_PORT  = int(os.getenv("A1111_PORT", "7860"))
A1111_URL   = f"http://{A1111_HOST}:{A1111_PORT}"
A1111_STEPS = int(os.getenv("A1111_STEPS", "1"))   # 1 for sd-turbo, 4 for LCM
A1111_CFG   = float(os.getenv("A1111_CFG", "0.0"))  # 0.0 for sd-turbo, 1.0–2.0 for LCM

AI_PROVIDER         = os.getenv("AI_PROVIDER", "fal").strip().lower()
PRODUCT_IMAGE_SIZE  = int(os.getenv("PRODUCT_IMAGE_SIZE", "2048"))

SHOT_TYPES = ("main", "lifestyle", "infographic", "macro", "styled_hero")

_VALID_PROVIDERS = {"fal", "a1111", "pil", "demo"}

# ── Helpers ───────────────────────────────────────────────────────────────────

def _safe(val: Any, fallback: str = "") -> str:
    """Return val stripped if non-empty / non-uncertain, else fallback."""
    v = str(val or "").strip()
    return fallback if v.lower() in {"", "uncertain", "unknown", "n/a"} else v


def _first(lst: list, fallback: str = "") -> str:
    for item in lst:
        s = _safe(item)
        if s:
            return s
    return fallback


def _join(lst: list, sep: str = ", ", limit: int = 3) -> str:
    parts = [_safe(x) for x in lst if _safe(x)]
    return sep.join(parts[:limit])


def _truncate_prompt(prompt: str, max_chars: int = 1000) -> str:
    """Trim prompt to max_chars."""
    if len(prompt) <= max_chars:
        return prompt
    truncated = prompt[:max_chars]
    last_period = truncated.rfind(".")
    if last_period > max_chars - 80:
        return truncated[: last_period + 1]
    return truncated.rstrip() + "."


# ── Theme profiles ────────────────────────────────────────────────────────────
# Each entry: keywords (any match triggers) → (theme_name, palette_hex_list,
#   lifestyle_setting, lifestyle_props, lifestyle_human_hint, lifestyle_mood,
#   lifestyle_light, infographic_bg_description)
# palette is 3-4 hex codes ordered: dominant, accent1, accent2, [highlight]

_THEME_PROFILES: List[Tuple[Tuple[str, ...], Dict[str, Any]]] = [
    (
        ("wellness", "yoga", "meditation", "aromatherapy", "essential oil", "spa", "ayurved", "herbal", "incense", "agarbatti"),
        {
            "theme": "wellness",
            "palette": ["#1B4332", "#C9A84C", "#FFFAF5", "#F0EAD6"],
            "lifestyle_scene": (
                "warm, earthy wellness altar — the product stands on a round natural teak wood tray "
                "at the centre. A small brass diya with a warm flame glows softly out of focus at the "
                "back-left; fresh marigold and jasmine petals are scattered loosely around the base of "
                "the tray, with one or two whole blooms resting near the edge. "
                "A few loose incense sticks or a rolled white cotton cloth lie casually beside the product. "
                "Terracotta and ivory tones fill the scene; a thin wisp of incense smoke spirals beautifully beside. "
                "Soft golden morning light from a large window at camera-left, "
                "gentle warm fill from the right, single clean amber shadow on the wood. "
                "Meditative, authentic, Ayurvedic luxury. Warm and inviting."
            ),
            "infographic_scene": (
                "the product centered on a round natural light-teak wood tray, shot from a gentle "
                "elevated 3/4 angle on a warm cream (#FFFAF5) backdrop. "
                "A few loose marigold and jasmine petals rest on the tray around the base of the product, "
                "with a thin wisp of incense smoke rising softly out of focus at one side. "
                "Perfectly even soft studio lighting — large softbox from upper-left, fill from the right — "
                "every label and surface detail crisp and legible. "
                "Generous clean open margin on all four sides of the frame, free of clutter, "
                "reserved for text and icon overlays added afterward. No patterns, no lines, no logos."
            ),
        },
    ),
    (
        ("skincare", "beauty", "serum", "moisturizer", "cream", "lotion", "face", "lip", "cosmetic"),
        {
            "theme": "beauty",
            "palette": ["#F7C5C5", "#E8967A", "#FAF3EE", "#C9956C"],
            "lifestyle_scene": (
                "white Carrara marble vanity top — the product stands upright at the centre-left. "
                "A single fresh pink rose petal lies at the base; a small amber-glass dropper bottle "
                "sits softly out of focus at the right edge. "
                "Blush-pink and warm coral tones wash the marble with reflected light. "
                "Overhead key light from a large square softbox, warm fill from camera-right, "
                "creating a glowing highlight along the product edge. Clean beauty, luxe, editorial."
            ),
            "infographic_scene": (
                "centered hero product on a gentle radial gradient from blush pink (#F7C5C5) at "
                "center to ivory (#FAF3EE) at all edges. "
                "A fine gold-foil hexagon-tile pattern floats across the upper portion at 6% opacity. "
                "A warm coral (#E8967A) thin arc stripe sweeps from bottom-left to top-right "
                "behind the product without touching it. Minimal, premium, beauty-editorial."
            ),
        },
    ),
    (
        ("gaming", "controller", "headset", "keyboard", "mouse", "rgb", "esport", "gamer"),
        {
            "theme": "gaming",
            "palette": ["#FF00CC", "#00FFDD", "#1A1A2E", "#0062FF"],
            "lifestyle_scene": (
                "dark gaming battlestation at 2 AM — dual monitors glow behind the product, "
                "their screens casting electric cyan and magenta halos across the deep-charcoal desk. "
                "A mechanical keyboard edge glows neon at camera-right; the desk surface has a subtle "
                "RGB underglow strip reflecting on the matte finish. "
                "Rim lighting: cyan from camera-left, magenta from camera-right, deep shadows everywhere else. "
                "Intense, high-energy, premium esports aesthetic."
            ),
            "infographic_scene": (
                "centered hero product on a deep charcoal (#1A1A2E) background. "
                "A bold diagonal gradient stripe from neon magenta (#FF00CC) to electric cyan (#00FFDD) "
                "cuts across the lower-left to upper-right at 40% opacity — behind the product, not over it. "
                "A subtle hexagonal grid pattern in electric blue (#0062FF) covers the whole background at 5%. "
                "Sharp, high-contrast, premium tech."
            ),
        },
    ),
    (
        ("baby", "infant", "newborn", "toddler", "nursery", "diaper", "pacifier", "teether", "head guard", "back protector"),
        {
            "theme": "baby",
            "palette": ["#FFFAF5", "#FFD166", "#B8E8D4", "#F5C0C0"],
            "lifestyle_scene": (
                "warm, airy nursery morning — the product rests against a plush cream bouclé-knit "
                "cushion on a soft-toned sofa or bed. A low wooden shelf with children's picture books "
                "and a small pastel plush toy sits softly out of focus behind it; a leafy potted plant "
                "stands at the far edge near a window. A chunky cream knitted throw blanket is draped "
                "loosely in the foreground corner. "
                "Airy natural window light from camera-left, pale warm walls, no harsh shadows. "
                "Warm, safe, joyful, and authentic — exactly what every parent hopes to see."
            ),
            "infographic_scene": (
                "centered hero product on a pure soft white (#FFFAF5) background. "
                "Even, bright, shadowless studio lighting from directly above. "
                "A very faint pastel-yellow (#FFF3C4) circular glow softly surrounds the product. "
                "Clean open space on all four sides. Minimal, safe, trustworthy."
            ),
        },
    ),
    (
        ("kitchen", "cookware", "bakeware", "utensil", "appliance", "cutting board", "knife", "pan", "pot",
         "bowl", "tray", "container", "box", "storage", "casserole"),
        {
            "theme": "kitchen",
            "palette": ["#1B4332", "#C9A84C", "#FFFAF5", "#F0EAD6"],
            "lifestyle_scene": (
                "warm kitchen scene — the product placed on a natural light-marble or warm-wood countertop. "
                "A small bunch of fresh herbs (coriander or basil) and a few raw ingredients are "
                "arranged casually nearby. A folded cream linen towel drapes the edge. "
                "Bright soft natural light from an overhead window — warm, inviting, fresh. "
                "Clean, appetizing, authentic home cooking aesthetic."
            ),
            "infographic_scene": (
                "centered hero product on a pure warm white (#FFFAF5) background. "
                "Soft even studio lighting from a large overhead softbox. "
                "Very subtle warm-cream (#F0EAD6) radial gradient directly behind the product. "
                "Clean open space on all four sides. No props, no distracting elements."
            ),
        },
    ),
    (
        ("food", "snack", "nutrition", "supplement", "protein", "vitamin", "health food", "organic", "fruit",
         "chip", "papad", "poha", "grain", "flour", "spice", "masala", "tea", "coffee", "dry fruit"),
        {
            "theme": "food",
            "palette": ["#1B4332", "#C9A84C", "#FFFAF5", "#F5EDE0"],
            "lifestyle_scene": (
                "sun-warmed rustic teak wood surface — the product stands at the center, slightly "
                "elevated, with a small earthen or wooden bowl placed directly beside it holding a "
                "portion of the product's actual contents poured out (loose grains, flakes, or leaves), "
                "a few stray pieces scattered naturally on the wood between the pack and the bowl. "
                "A sprig of fresh mint or coriander garnishes the bowl for appetite appeal; "
                "a rough jute or cream linen cloth is draped casually at one corner. "
                "Warm golden-hour light streams softly from camera-left, "
                "casting rich amber tones and a single clean shadow; diffused fill on the right. "
                "Earthy, wholesome, authentic Indian artisan brand. "
                "Product razor-sharp, background beautifully blurred and inviting."
            ),
            "infographic_scene": (
                "centered hero product on a pure warm white (#FFFAF5) background. "
                "Perfectly even soft studio lighting — large octabox from upper-left, fill from the right. "
                "A very subtle warm cream radial glow (#F5EDE0) from directly behind the product. "
                "Clean open space on all four sides. "
                "No patterns, no stripes, no competing elements."
            ),
        },
    ),
    (
        ("outdoor", "camping", "hiking", "trekking", "garden", "yard", "tool", "hardware"),
        {
            "theme": "outdoor",
            "palette": ["#2D5A27", "#D4821A", "#8B8680", "#87CEEB"],
            "lifestyle_scene": (
                "rugged outdoor setting — the product rests on a large flat granite boulder "
                "surrounded by pine needles and dappled forest light. "
                "Rich green foliage is blurred in the background; a thin ribbon of blue sky "
                "peeks through the canopy at the top. "
                "Strong directional natural sun from camera-right, deep amber shadows, "
                "cool blue sky-fill from above. Rugged, authentic, bold."
            ),
            "infographic_scene": (
                "centered hero product on a dark forest-green (#2D5A27) background. "
                "A thick diagonal amber (#D4821A) stripe crosses from bottom-left to mid-right behind the product. "
                "A stone-grey (#8B8680) angular triangle element sits in the upper-right corner. "
                "Industrial-meets-nature, bold and confident."
            ),
        },
    ),
    (
        ("sports", "fitness", "gym", "running", "cycling", "swim", "exercise", "workout", "athletic"),
        {
            "theme": "fitness",
            "palette": ["#E63946", "#457B9D", "#F1FAEE", "#1D3557"],
            "lifestyle_scene": (
                "high-performance gym at golden hour — the product rests on a polished black rubber "
                "gym floor. Blurred weight racks and neon gym signage glow in the background. "
                "A matching water bottle and wireless earbuds sit at the edge of frame, out of focus. "
                "Overhead LED gym lighting punches down, creating sharp dramatic shadows; "
                "a warm orange window light fills from one side. Energetic, motivational, performance-driven."
            ),
            "infographic_scene": (
                "centered hero product on a deep navy (#1D3557) background. "
                "A bold energetic red (#E63946) diagonal speed-stripe sweeps from lower-left to upper-right, "
                "thick and confident, partially behind the product. "
                "A subtle ocean-blue (#457B9D) curved arc echoes it from the opposite corner. "
                "Dynamic, sporty, premium."
            ),
        },
    ),
    (
        ("apparel", "clothing", "fashion", "shirt", "dress", "kurti", "saree", "kurta", "top", "jacket"),
        {
            "theme": "fashion",
            "palette": ["#F9F0E3", "#C9956C", "#1B3A5C", "#2C4A3E"],
            "lifestyle_scene": (
                "editorial flat-lay on a warm ivory linen surface. "
                "The garment is laid naturally with subtle organic folds; a slim gold watch and a "
                "small leather wallet are placed casually in the lower corner. "
                "A single stem dried flower lies across the bottom edge for texture. "
                "Soft overhead diffused studio light from a large octabox, no harsh shadows, "
                "even warm tone across the whole surface. "
                "Warm ivory, rose-gold, and deep navy tones in the accessories echo the palette. "
                "Aspirational, fashion-forward, editorial."
            ),
            "infographic_scene": (
                "centered hero garment on a warm ivory (#F9F0E3) background, "
                "displayed naturally with gentle organic folds. "
                "A rose-gold (#C9956C) diagonal double-line accent sweeps from the lower-right. "
                "A solid navy (#1B3A5C) rectangular block anchors the lower-left corner. "
                "Elegant, editorial, fashion-brand aesthetic."
            ),
        },
    ),
    (
        ("pet", "dog", "cat", "animal", "collar", "leash", "bowl", "treat"),
        {
            "theme": "pet",
            "palette": ["#C4A882", "#3D6B47", "#8B6355", "#F5F0E8"],
            "lifestyle_scene": (
                "warm home living room — the product sits on a natural oak herringbone floor. "
                "A happy golden retriever (or relevant pet breed) lies just behind it, "
                "nose nearly touching the product, softly blurred. "
                "A plaid throw blanket is draped over a couch edge in the far background. "
                "Warm afternoon window light streams from camera-left, painting golden tones "
                "across the floor and the animal's fur. Loving, warm, authentic."
            ),
            "infographic_scene": (
                "centered hero product on a soft cream (#F5F0E8) background. "
                "A warm-tan (#C4A882) organic paw-print watermark pattern tiles the background at 9% opacity. "
                "Two forest-green (#3D6B47) solid rectangular bands frame the upper and lower edges. "
                "Friendly, warm, trustworthy."
            ),
        },
    ),
    (
        ("electronics", "gadget", "tech", "phone", "laptop", "tablet", "charger", "cable", "speaker", "earphone"),
        {
            "theme": "tech",
            "palette": ["#0D1B2A", "#1E88E5", "#CFD8DC", "#ECEFF1"],
            "lifestyle_scene": (
                "minimalist executive desk — the product sits on a matte black deskpad. "
                "A slim MacBook is open just behind it at camera-right, screen glowing softly; "
                "a premium matte-black pen and a small succulent sit to the left. "
                "Cool-toned studio lighting from overhead with a subtle blue-electric fill from the right, "
                "precise crisp edge highlights on the product, no lens flare. "
                "Sleek, confident, premium tech editorial."
            ),
            "infographic_scene": (
                "centered hero product on a very dark navy (#0D1B2A) background. "
                "Electric blue (#1E88E5) bold geometric lines — two diagonal and one arc — "
                "flow behind the product like a circuit trace. "
                "Light-grey (#CFD8DC) solid rectangular callout zone blocks sit at the four "
                "cardinal edge positions. Crisp, modern, high-tech."
            ),
        },
    ),
    (
        ("jewelry", "ring", "necklace", "bracelet", "earring", "watch", "accessory", "pendant"),
        {
            "theme": "jewelry",
            "palette": ["#C9A84C", "#111111", "#FEFDE8", "#B8777A"],
            "lifestyle_scene": (
                "near-black velvet surface — the piece rests alone or draped naturally. "
                "A single dried rose petal lies a few centimetres away for scale and romance. "
                "One narrow focused spotlight from directly above-left creates a dramatic single shadow; "
                "the gold tones glint and scatter tiny light points across the velvet. "
                "Rich, intimate, precious — the product fills 70% of the frame."
            ),
            "infographic_scene": (
                "centered hero product on a near-black (#111111) background. "
                "Antique gold (#C9A84C) ultra-thin geometric frame lines form a partial rectangle "
                "around the product without touching it. "
                "A soft ivory (#FEFDE8) diagonal gradient swatch fades from the upper-right corner "
                "to mid-background. Ultra-premium luxury editorial."
            ),
        },
    ),
    (
        ("home", "furniture", "decor", "lamp", "pillow", "candle", "vase", "rug", "bedding"),
        {
            "theme": "home",
            "palette": ["#FAF8F5", "#9CAF88", "#C47A5A", "#5C3D2E"],
            "lifestyle_scene": (
                "curated interior vignette — the product is placed on a walnut side table beside "
                "a linen sofa. A small olive-green ceramic pot plant sits at the edge; "
                "a stack of three linen-covered books leans casually nearby. "
                "Warm afternoon light pours through a sheer curtain at camera-left, "
                "casting long soft shadows across the warm-white wall. "
                "Sage-green, terracotta, and walnut tones echo through the room. "
                "Warm, aspirational, interior-design editorial."
            ),
            "infographic_scene": (
                "centered hero product on a warm white (#FAF8F5) background. "
                "Two sage-green (#9CAF88) bold horizontal rectangular bands — one at top, one at bottom — "
                "frame the composition. A terracotta (#C47A5A) circle accent, 220px diameter, "
                "sits in the lower-right behind the product baseline. Organic, warm, editorial."
            ),
        },
    ),
]

# Default theme when no keyword matches
_DEFAULT_THEME: Dict[str, Any] = {
    "theme": "lifestyle",
    "palette": ["#1B4332", "#C9A84C", "#FFFAF5", "#F0EAD6"],
    "lifestyle_scene": (
        "warm natural light-wood or marble surface — the product stands upright at the centre. "
        "A folded cream linen cloth and a small complementary ceramic prop rest softly "
        "out of focus at one edge. "
        "Soft golden natural daylight from a large window at camera-left, "
        "gentle warm reflector fill on the right, single clean directional shadow on the surface. "
        "Warm, premium, aspirational — Indian artisan brand aesthetic."
    ),
    "infographic_scene": (
        "centered hero product on a pure warm white (#FFFAF5) background. "
        "Perfectly even soft studio lighting — large softbox from upper-left, fill from the right. "
        "A subtle warm cream radial glow (#F0EAD6) behind the product. "
        "Clean open space on all four sides. No patterns, no competing elements."
    ),
}

# Appended inline to every lifestyle + infographic prompt — keeps preservation
# guidance near the end so creative direction registers first.
_PRESERVATION_RULE = (
    "IMPORTANT: The product must be pixel-perfect identical to the reference image — "
    "same shape, color, finish, branding, logos, text, label design, and proportions. "
    "Do NOT restyle, recolor, redesign, or relabel the product. "
    "Only the background, scene, lighting, props, and graphic elements may be generated."
)


# ── Inline BG removal (no external deps, works on studio white/grey backdrops) ─

def _remove_bg(img: "Image") -> "Image":
    """Edge-aware erode -> flood-fill -> dilate to remove studio backgrounds safely.

    Returns an RGBA image with the background made transparent.
    Works on white, grey, warm-white backdrops; preserves coloured products —
    including dark/coloured products whose overall tone is close to the
    backdrop's average colour.

    Two leak modes are guarded against:
      1. A single-pixel-wide gap/bridge (e.g. a thin highlight) that lets the
         fill leak deep into the product. Fixed by eroding the candidate mask
         by 1px before flooding, then dilating the result back afterwards.
      2. A product whose shade is close enough to the *average* backdrop
         colour to pass a simple global tolerance check (e.g. a dark-grey
         product photographed on a dark-grey background). A pure global
         colour-distance test would eat straight through it. Fixed by also
         requiring each flood step to stay close to its *immediate neighbour*
         already confirmed as background — real product edges almost always
         show a visible step (shadow/contour/specular line) even when the
         product's overall tone matches the backdrop, so this local-continuity
         check stops the fill at that edge instead of jumping across it.
    """
    from collections import deque

    rgba = img.convert("RGBA")
    pix  = rgba.load()
    w, h = rgba.size

    samples: list = []
    step = max(1, w // 40)
    for x in range(0, w, step):
        samples.append(pix[x, 0][:3])
        samples.append(pix[x, h - 1][:3])
    step = max(1, h // 40)
    for y in range(0, h, step):
        samples.append(pix[0, y][:3])
        samples.append(pix[w - 1, y][:3])
    n   = max(1, len(samples))
    bgR = sum(s[0] for s in samples) // n
    bgG = sum(s[1] for s in samples) // n
    bgB = sum(s[2] for s in samples) // n
    tol_global = int(os.getenv("BG_REMOVE_TOLERANCE", "40"))
    # Max allowed colour step between a pixel and the already-confirmed
    # background neighbour it's growing from. Tighter than tol_global so the
    # fill can't cross a real product edge just because the product's overall
    # tone happens to fall within the global tolerance.
    tol_local = int(os.getenv("BG_REMOVE_LOCAL_TOLERANCE", "22"))

    def _near_global(r: int, g: int, b: int) -> bool:
        return abs(r - bgR) <= tol_global and abs(g - bgG) <= tol_global and abs(b - bgB) <= tol_global

    def _near_local(c1, c2) -> bool:
        return abs(c1[0] - c2[0]) + abs(c1[1] - c2[1]) + abs(c1[2] - c2[2]) <= tol_local

    # Candidate mask: pixels that colour-match the background globally.
    candidate = [[False] * w for _ in range(h)]
    for y in range(h):
        row = candidate[y]
        for x in range(w):
            r, g, b, _ = pix[x, y]
            row[x] = _near_global(r, g, b)

    # Erode by 1px (4-neighbourhood) so single-pixel leaks/bridges are cut
    # before the flood-fill ever sees them.
    eroded = [[False] * w for _ in range(h)]
    for y in range(h):
        for x in range(w):
            if not candidate[y][x]:
                continue
            if (
                x > 0 and candidate[y][x - 1]
                and x < w - 1 and candidate[y][x + 1]
                and y > 0 and candidate[y - 1][x]
                and y < h - 1 and candidate[y + 1][x]
            ):
                eroded[y][x] = True
            elif x == 0 or x == w - 1 or y == 0 or y == h - 1:
                # Border pixels keep their candidate status (no outside
                # neighbour to erode against) so seeding still works.
                eroded[y][x] = True

    # Flood-fill the eroded mask from all four borders. Growth requires both
    # the global backdrop-tone match (via `eroded`) AND local continuity with
    # the neighbour pixel it's expanding from, so a same-toned product region
    # separated by a real edge doesn't get swallowed.
    visited = [[False] * w for _ in range(h)]
    q: deque = deque()

    def _seed(x: int, y: int) -> None:
        if not visited[y][x] and eroded[y][x]:
            visited[y][x] = True
            q.append((x, y))

    for x in range(w):
        _seed(x, 0); _seed(x, h - 1)
    for y in range(h):
        _seed(0, y); _seed(w - 1, y)

    while q:
        cx, cy = q.popleft()
        ccol = pix[cx, cy][:3]
        for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1)):
            nx, ny = cx + dx, cy + dy
            if 0 <= nx < w and 0 <= ny < h and not visited[ny][nx] and eroded[ny][nx]:
                if _near_local(ccol, pix[nx, ny][:3]):
                    visited[ny][nx] = True
                    q.append((nx, ny))

    # Second pass: a studio pedestal/platform is often shaded distinctly
    # darker/lighter than the backdrop behind it, so its colour can fall
    # outside `tol_global` (measured against the single backdrop-average
    # colour) even though it is obviously still background. Build a wider
    # candidate mask for this pass, but keep expansion strictly gated by the
    # same tight `tol_local` step-by-step continuity check, seeded only from
    # the background already confirmed above — so it can bridge into a
    # differently-toned background region without ever hopping across a
    # real product edge (which still produces a > tol_local jump).
    tol_global_2 = tol_global * 2
    candidate2 = [[False] * w for _ in range(h)]
    for y in range(h):
        row2 = candidate2[y]
        for x in range(w):
            r, g, b, _ = pix[x, y]
            row2[x] = abs(r - bgR) <= tol_global_2 and abs(g - bgG) <= tol_global_2 and abs(b - bgB) <= tol_global_2

    q2: deque = deque()
    for y in range(h):
        for x in range(w):
            if visited[y][x]:
                q2.append((x, y))
    while q2:
        cx, cy = q2.popleft()
        ccol = pix[cx, cy][:3]
        for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1)):
            nx, ny = cx + dx, cy + dy
            if 0 <= nx < w and 0 <= ny < h and not visited[ny][nx] and candidate2[ny][nx]:
                if _near_local(ccol, pix[nx, ny][:3]):
                    visited[ny][nx] = True
                    q2.append((nx, ny))

    # Dilate the flood-filled region by 1px, but only re-claim pixels that
    # were already colour-candidates — this restores the true background
    # edge that erosion shaved off, without ever spreading into the product.
    bg_mask = [row[:] for row in visited]
    for y in range(h):
        for x in range(w):
            if visited[y][x]:
                continue
            if not candidate[y][x] and not candidate2[y][x]:
                continue
            if (
                (x > 0 and visited[y][x - 1])
                or (x < w - 1 and visited[y][x + 1])
                or (y > 0 and visited[y - 1][x])
                or (y < h - 1 and visited[y + 1][x])
            ):
                bg_mask[y][x] = True

    for y in range(h):
        row_mask = bg_mask[y]
        for x in range(w):
            if row_mask[x]:
                r, g, b, _ = pix[x, y]
                pix[x, y] = (r, g, b, 0)

    return rgba


def _crop_tight(img: "Image") -> "Image":
    """Crop RGBA image to its tight non-transparent bounding box."""
    if img.mode != "RGBA":
        img = img.convert("RGBA")
    bb = img.getbbox()
    return img.crop(bb) if bb else img


def _resolve_theme(product: Dict) -> Dict[str, Any]:
    """Return the best-matching theme profile for this product."""
    search_str = " ".join([
        _safe(product.get("product_type", "")),
        _safe(product.get("category", "")),
        _safe(product.get("usage", "")),
        _safe(product.get("style", "")),
    ]).lower()

    for keywords, profile in _THEME_PROFILES:
        if any(kw in search_str for kw in keywords):
            return profile
    return _DEFAULT_THEME


# ── Prompt builders ───────────────────────────────────────────────────────────

def _build_main_prompt(product: Dict) -> str:
    ptype    = _safe(product.get("product_type"), "product")
    material = _safe(product.get("material"))
    colors   = _join(product.get("colors", []))
    desc     = " ".join(filter(None, [colors, material, ptype]))
    return (
        f"Award-winning commercial e-commerce product photograph of {desc}. "
        "Pure seamless white background, RGB (255,255,255), no gradient, no shadow on background. "
        "Product perfectly centered, filling 85% of the square frame, shot at a crisp 3/4 angle. "
        "Three-point softbox studio lighting: dominant key from upper-left at 45°, "
        "large fill softbox at camera-right at 50% power, hairlight from behind at 30% power — "
        "revealing every surface texture and edge definition. "
        "Single clean contact shadow directly beneath the product only. "
        "Shot on Hasselblad H6D-100c, 100mm macro lens, f/11, ISO 64, ultra-sharp front-to-back focus. "
        "True-to-life color reproduction, no post-saturation boost. "
        "No text, no props, no logos added, no watermarks. "
        "Hyperrealistic, 8K, commercial product photography master-class quality."
    )


def _build_lifestyle_prompt(product: Dict) -> str:
    theme    = _resolve_theme(product)
    ptype    = _safe(product.get("product_type"), "product")
    material = _safe(product.get("material"))
    colors   = _join(product.get("colors", []))
    features = _join(product.get("features", []), limit=2)
    usage    = _safe(product.get("usage"))
    desc     = " ".join(filter(None, [colors, material, ptype]))
    scene    = theme["lifestyle_scene"]
    palette  = theme.get("palette", [])
    pal_str  = ", ".join(palette[:2]) if palette else "warm neutrals"
    mat_note = (
        f"The {material} finish on the product is crisp and true-to-life — texture, sheen, and label intact. "
        if material and material.lower() not in {"uncertain", "unknown", "n/a"} else ""
    )
    feat_note = f"Clearly visible: {features}. " if features else ""
    usage_note = (
        f"The product is shown in active, natural use — real person, real moment, real life ({usage}). "
        if usage and usage.lower() not in {"uncertain", "unknown", "n/a"} else ""
    )

    return (
        f"Award-winning lifestyle product photograph of the {desc}. "
        f"{scene} "
        f"{usage_note}"
        f"{mat_note}"
        f"{feat_note}"
        "Shot on Canon EOS R5, 50mm f/1.4 — product razor-sharp with creamy shallow depth of field, "
        "background beautifully blurred into warm bokeh. "
        f"Color palette: {pal_str}. "
        "Color grading: warm golden tones, lifted shadows, rich and inviting — not cold, not sterile. "
        "Lighting: soft golden natural light from a large window at camera-left, warm fill on the right. "
        "No text overlays, no watermarks, no artificial logos added to the scene. "
        "Hyperrealistic, 8K, premium Indian brand commercial photography. "
        f"{_PRESERVATION_RULE}"
    )


def _build_infographic_bg_prompt(product: Dict) -> str:
    """Generate a clean background + centered hero product only.

    Text callouts, feature icons, and detail chips are overlaid via PIL after
    generation — this prompt must produce a CLEAN, OPEN background.
    """
    theme    = _resolve_theme(product)
    ptype    = _safe(product.get("product_type"), "product")
    material = _safe(product.get("material"))
    colors   = _join(product.get("colors", []))
    desc     = " ".join(filter(None, [colors, material, ptype]))
    bg_scene = theme["infographic_scene"]

    return (
        f"Premium Amazon product studio image of the {desc} — clean infographic background. "
        f"{bg_scene} "
        "The product is the single centered hero, perfectly sharp, filling 60% of the square frame. "
        "Product lighting: large softbox from upper-left at 45 degrees, gentle fill reflector from the right — "
        "no blown highlights, no color cast, every label and surface detail perfectly clear. "
        "A soft, clean contact shadow sits directly beneath the product only. "
        "The background is intentionally minimal and open — large clear zones on all four sides "
        "to accommodate text and graphic overlays that will be added programmatically. "
        "Square 1:1 format, 2000x2000px, hyperrealistic, razor-sharp product rendering. "
        "STRICT: absolutely zero text, zero icons, zero callouts, zero arrows, zero badges — "
        "background and product only, no annotations whatsoever. "
        f"{_PRESERVATION_RULE}"
    )


def _build_macro_prompt(product: Dict) -> str:
    ptype    = _safe(product.get("product_type"), "product")
    features = product.get("features", [])
    key_feat = _first(features, "surface texture and craftsmanship")
    material = _safe(product.get("material"))
    desc     = " ".join(filter(None, [material, ptype]))
    return (
        f"Extreme close-up macro product photograph revealing {key_feat} of the {desc}. "
        "Tight crop: the feature fills 80% of the frame. "
        "Single raking sidelight from camera-left reveals every micro-texture and edge detail. "
        "Very shallow depth of field — the feature is perfectly razor-sharp, "
        "foreground and background fall off smoothly into creamy bokeh. "
        "Soft neutral gradient background, no distractions. "
        "Shot on Canon MP-E 65mm macro lens, f/4, focus-stacked for maximum detail. "
        "Premium tactile feel, hyperrealistic texture reproduction, no text, 4K."
    )


def _build_styled_hero_prompt(product: Dict) -> str:
    theme         = _resolve_theme(product)
    ptype         = _safe(product.get("product_type"), "product")
    material      = _safe(product.get("material"))
    colors        = product.get("colors", [])
    primary_color = _first(colors, "neutral")
    desc          = " ".join(filter(None, [material, ptype]))
    palette       = theme.get("palette", [])
    p0 = palette[0] if palette else "#F0EDE8"
    p1 = palette[1] if len(palette) > 1 else "#4A7C6B"

    return (
        f"High-end styled hero product photograph of the {desc}. "
        f"The product sits on a clean matte surface with a smooth radial gradient background "
        f"sweeping from {p0} at center to a slightly deeper {p1} at the edges. "
        f"The {primary_color} of the product is gently echoed in the backdrop's mid-tones "
        "for a harmonious, intentional aesthetic. "
        "Camera angle: slightly elevated 3/4 view from the front-left corner. "
        "Lighting: large Profoto softbox overhead-right as key, "
        "V-flat bounce on camera-left as fill, rim light from directly behind — "
        "edges glow cleanly against the gradient. "
        "One crisp natural drop shadow beneath the product for grounding. "
        "A single complementary prop softly out of focus at the extreme edge of frame. "
        "Shot on Sony A7R V, 85mm f/1.4, f/5.6, studio flash, ultra-sharp. "
        "Photorealistic, premium brand aesthetic, no CGI look, no text, no logos added. "
        f"{_PRESERVATION_RULE}"
    )


# ── Category / material adaptive logic ────────────────────────────────────────

_GLOSSY_MATERIALS = {"glass", "mirror", "chrome", "metal", "stainless", "acrylic", "glossy", "shiny", "lacquer"}


def _adapt_by_category(
    category: str, style: str, material: str, shot_type: str, prompt: str
) -> str:
    """Append spec-mandated adaptive clauses based on product category/material."""
    cat_lower = (category + " " + style).lower()
    mat_lower = material.lower()

    # Apparel flat-lay enforcement on main shot
    _APPAREL_KW = ("apparel", "clothing", "shirt", "dress", "kurti", "saree", "kurta", "garment")
    if shot_type == "main" and any(kw in cat_lower for kw in _APPAREL_KW):
        prompt = prompt.rstrip(". ") + ". Garment laid flat on white surface, NOT on a model, centered, no mannequin."

    if shot_type == "main" and any(kw in cat_lower for kw in ("shoe", "footwear", "sneaker", "sandal")):
        prompt = prompt.rstrip(". ") + ". Single shoe facing left at 45-degree angle, white background."

    # Glossy/reflective guard — applies to main shot only
    if shot_type == "main":
        if any(g in mat_lower for g in _GLOSSY_MATERIALS) or "glossy" in cat_lower:
            prompt = prompt.rstrip(". ") + (
                ". Minimal controlled reflection, no visible light sources in the reflection."
            )

    return prompt


def _build_prompt(product: Dict, shot_type: str) -> str:
    """Build the full prompt for a given shot type, with adaptive clauses applied."""
    if shot_type == "main":
        base = _build_main_prompt(product)
    elif shot_type == "lifestyle":
        base = _build_lifestyle_prompt(product)
    elif shot_type == "infographic":
        base = _build_infographic_bg_prompt(product)
    elif shot_type == "macro":
        base = _build_macro_prompt(product)
    elif shot_type == "styled_hero":
        base = _build_styled_hero_prompt(product)
    else:
        raise ValueError(f"Unknown shot_type '{shot_type}'. Must be one of {SHOT_TYPES}.")

    category = _safe(product.get("category", ""))
    style    = _safe(product.get("style", ""))
    material = _safe(product.get("material", ""))
    return _adapt_by_category(category, style, material, shot_type, base)


# ── Low-level API wrappers ────────────────────────────────────────────────────


def _a1111_generate_image_bytes(prompt: str) -> bytes:
    """Generate image via a local ComfyUI server; return PNG bytes."""
    import base64 as _b64
    import random

    client_id = str(random.randint(10000, 99999))

    # Minimal SD-Turbo / LCM workflow
    workflow = {
        "3": {
            "class_type": "KSampler",
            "inputs": {
                "cfg": A1111_CFG if A1111_CFG > 0 else 1.0,
                "denoise": 1,
                "latent_image": ["5", 0],
                "model": ["4", 0],
                "negative": ["7", 0],
                "positive": ["6", 0],
                "sampler_name": "euler",
                "scheduler": "normal",
                "seed": random.randint(0, 2**32),
                "steps": A1111_STEPS,
            },
        },
        "4": {"class_type": "CheckpointLoaderSimple", "inputs": {"ckpt_name": "sd_turbo.safetensors"}},
        "5": {"class_type": "EmptyLatentImage", "inputs": {"batch_size": 1, "height": 512, "width": 512}},
        "6": {"class_type": "CLIPTextEncode", "inputs": {"clip": ["4", 1], "text": prompt}},
        "7": {"class_type": "CLIPTextEncode", "inputs": {"clip": ["4", 1], "text": "blurry, low quality, watermark, text, logo"}},
        "8": {"class_type": "VAEDecode", "inputs": {"samples": ["3", 0], "vae": ["4", 2]}},
        "9": {"class_type": "SaveImage", "inputs": {"filename_prefix": "api_out", "images": ["8", 0]}},
    }

    payload = json.dumps({"prompt": workflow, "client_id": client_id}).encode("utf-8")
    req = Request(
        url=f"{A1111_URL}/prompt",
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urlopen(req, timeout=30) as resp:
            queue_result = json.loads(resp.read().decode("utf-8"))
    except Exception as exc:
        raise ValueError(
            f"ComfyUI at {A1111_URL} unreachable: {exc}. "
            "Make sure ComfyUI is running: python main.py --listen --port 7860"
        ) from exc

    prompt_id = queue_result.get("prompt_id")
    if not prompt_id:
        raise ValueError(f"ComfyUI did not return a prompt_id: {queue_result}")

    # Poll until the job is done
    history_url = f"{A1111_URL}/history/{prompt_id}"
    for _ in range(120):
        time.sleep(2)
        req_h = Request(history_url)
        with urlopen(req_h, timeout=10) as r:
            history = json.loads(r.read().decode("utf-8"))
        if prompt_id in history:
            outputs = history[prompt_id].get("outputs", {})
            for node_out in outputs.values():
                for img_info in node_out.get("images", []):
                    img_req = Request(
                        f"{A1111_URL}/view?filename={img_info['filename']}"
                        f"&subfolder={img_info.get('subfolder','')}"
                        f"&type={img_info.get('type','output')}"
                    )
                    with urlopen(img_req, timeout=30) as ir:
                        return ir.read()
            break

    raise ValueError(f"ComfyUI job {prompt_id} produced no images.")


def _fal_generate_image_bytes(prompt: str) -> bytes:
    """Generate image via fal.ai; return PNG bytes."""
    if not FAL_AVAILABLE:
        raise ValueError("FAL_KEY not set.")
    import fal_client
    import base64 as _b64

    # flux/schnell uses 4 steps; flux/dev needs ~28 for quality output
    _is_schnell = "schnell" in FAL_IMAGE_MODEL.lower()
    _steps = 4 if _is_schnell else 28
    _args: dict = {
        "prompt": prompt,
        "image_size": "square_hd",
        "num_inference_steps": _steps,
        "num_images": 1,
        "enable_safety_checker": False,
    }
    if not _is_schnell:
        _args["guidance_scale"] = 3.5
    result = fal_client.run(FAL_IMAGE_MODEL, arguments=_args)
    images = result.get("images") or []
    if not images:
        raise ValueError("fal.ai returned no images.")

    first = images[0]
    url = first.get("url", "") if isinstance(first, dict) else ""
    if url.startswith("data:"):
        # inline base64 data URI
        b64 = url.split(",", 1)[1]
        return _b64.b64decode(b64)
    if url.startswith("http"):
        req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urlopen(req, timeout=60) as r:
            return r.read()
    raise ValueError(f"fal.ai: unrecognised image format in response: {str(first)[:200]}")


def _draw_vgrad(canvas: Any, top: tuple, bot: tuple) -> None:
    from PIL import ImageDraw as _ID
    d = _ID.Draw(canvas)
    h = canvas.size[1]; w = canvas.size[0]
    for y in range(h):
        t = y / h
        d.line([(0, y), (w - 1, y)],
               fill=tuple(int(top[i] + (bot[i] - top[i]) * t) for i in range(3)))


def _load_font(size: int = 28):
    from PIL import ImageFont as _IF
    for path in (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
    ):
        try:
            return _IF.truetype(path, size)
        except Exception:
            pass
    return _IF.load_default()


def _paste_shadow(canvas: Any, product: Any, x: int, y: int,
                  offset: int = 18, blur: int = 30, alpha: int = 60) -> None:
    from PIL import Image as _Im, ImageFilter as _IF
    rw, rh = product.size
    shad = _Im.new("RGBA", canvas.size, (0, 0, 0, 0))
    mask = product.split()[3] if product.mode == "RGBA" else _Im.new("L", (rw, rh), 255)
    fill = _Im.new("RGBA", (rw, rh), (0, 0, 0, alpha))
    shad.paste(fill, (x + offset, y + offset), mask)
    shad = shad.filter(_IF.GaussianBlur(blur))
    canvas.paste(shad.convert("RGB"), (0, 0), shad.split()[3])
    canvas.paste(product.convert("RGBA"), (x, y),
                 product.split()[3] if product.mode == "RGBA" else None)


def _prep_ref(source: str) -> "Optional[Any]":
    from PIL import Image as _Im
    try:
        raw = _read_bytes(source)
        img = _Im.open(BytesIO(raw)).convert("RGBA")
        img = _remove_bg(img)
        return _crop_tight(img)
    except Exception as exc:
        log.warning("PIL fallback: ref load failed: %s", exc)
        return None


def _hex(hx: str) -> tuple:
    hx = hx.lstrip("#")
    return (int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))


def _pil_fallback_image(
    product: Dict,
    reference_source: str,
    out_path: Path,
    shot_type: str = "main",
) -> Path:
    """High-quality styled PIL fallback — distinct per shot type, BG removal applied.

    main        — pure white compliance shot with clean shadow
    lifestyle   — themed warm gradient, rule-of-thirds product, feature tag
    infographic — branded header, feature callout cards (left), product (right)
    macro       — tight center crop simulating close-up lens, detail badge
    styled_hero — dark moody backdrop, elevated product, gold typography
    """
    from PIL import Image as _Im, ImageDraw as _ID, ImageFilter as _IF

    S      = PRODUCT_IMAGE_SIZE
    ptype  = _safe(product.get("product_type", "Product")) or "Product"
    colors = [_safe(c) for c in product.get("colors", []) if _safe(c)]
    feats  = [_safe(f) for f in product.get("features", []) if _safe(f)]
    mat    = _safe(product.get("material", ""))
    f_title = _load_font(max(32, S // 50))
    f_body  = _load_font(max(22, S // 72))
    theme   = _resolve_theme(product)
    pal     = theme.get("palette", ["#1A2648", "#B48C3C", "#F5F0E8"])

    # ── MAIN — pure white compliance shot ─────────────────────────────────────
    if shot_type == "main":
        canvas = _Im.new("RGB", (S, S), (255, 255, 255))
        if reference_source:
            ref = _prep_ref(reference_source)
            if ref:
                ref.thumbnail((int(S * 0.82), int(S * 0.82)), _Im.LANCZOS)
                rw, rh = ref.size
                _paste_shadow(canvas, ref, (S-rw)//2, (S-rh)//2, offset=12, blur=22, alpha=35)

    # ── LIFESTYLE — themed warm scene, rule-of-thirds product ─────────────────
    elif shot_type == "lifestyle":
        top_c  = _hex(pal[2]) if len(pal) > 2 else (245, 240, 228)
        bot_c  = _hex(pal[0]) if len(pal) > 0 else (218, 200, 168)
        accent = _hex(pal[1]) if len(pal) > 1 else (139, 115, 85)
        canvas = _Im.new("RGB", (S, S), top_c)
        _draw_vgrad(canvas, top_c, bot_c)
        d = _ID.Draw(canvas)
        # Soft oval highlight
        hl = _Im.new("RGBA", (S, S), (0, 0, 0, 0))
        _ID.Draw(hl).ellipse([(S//4, S//6), (S*3//4, S*5//6)], fill=(255, 255, 255, 38))
        hl = hl.filter(_IF.GaussianBlur(S // 8))
        canvas.paste(hl.convert("RGB"), (0, 0), hl.split()[3])
        # Floor strip
        d = _ID.Draw(canvas)
        for y in range(S - S//6, S):
            t = (y - (S - S//6)) / (S // 6)
            d.line([(0, y), (S-1, y)],
                   fill=tuple(max(0, int(c - 30 * t)) for c in bot_c))
        # Product at rule-of-thirds
        if reference_source:
            ref = _prep_ref(reference_source)
            if ref:
                ref.thumbnail((int(S * 0.66), int(S * 0.66)), _Im.LANCZOS)
                rw, rh = ref.size
                _paste_shadow(canvas, ref, int(S*0.42)-rw//2, int(S*0.50)-rh//2,
                               offset=20, blur=35, alpha=50)
                d = _ID.Draw(canvas)
        # Colour / type tag top-right
        tag = (colors[0] if colors else ptype)[:28].upper()
        tb = f_body.getbbox(tag)
        tw, th = tb[2]-tb[0]+24, tb[3]-tb[1]+14
        d.rounded_rectangle([(S-tw-28, 28), (S-28, 28+th)], radius=6, fill=accent)
        d.text((S-tw-16, 35), tag, font=f_body, fill=(255, 255, 255))
        # Title strip bottom-right
        words = ptype.split(); line = ""; tlines = []
        for ww in words:
            t2 = (line + " " + ww).strip()
            bb = f_title.getbbox(t2)
            if bb[2]-bb[0] < S * 0.46:
                line = t2
            else:
                if line: tlines.append(line)
                line = ww
        if line: tlines.append(line)
        tlines = tlines[:2]
        sh = len(tlines) * (f_title.size + 10) + 24
        sy = S - sh - 18
        ov = _Im.new("RGBA", (S//2, sh+4), (*accent, 210))
        canvas.paste(ov.convert("RGB"), (S//2, sy), ov.split()[3])
        d = _ID.Draw(canvas)
        ty = sy + 12
        for tl in tlines:
            d.text((S//2 + 14, ty), tl, font=f_title, fill=(255, 255, 255))
            ty += f_title.size + 10

    # ── INFOGRAPHIC — branded header, feature cards left, product right ────────
    elif shot_type == "infographic":
        canvas = _Im.new("RGB", (S, S), (248, 246, 242))
        navy   = _hex(pal[0]) if len(pal) > 0 else (26, 38, 72)
        gold   = _hex(pal[1]) if len(pal) > 1 else (180, 140, 60)
        d = _ID.Draw(canvas)
        d.rectangle([(0, 0), (S, S//8)], fill=navy)
        d.rectangle([(0, S//8), (S, S//8+4)], fill=gold)
        tb = f_title.getbbox(ptype.upper())
        d.text((S//20, (S//8-(tb[3]-tb[1]))//2), ptype.upper(), font=f_title, fill=(255, 255, 255))
        d.line([(S//2, S//8+4), (S//2, S)], fill=(210, 208, 204), width=2)
        # Feature cards
        items = []
        if mat:       items.append(f"Material: {mat}")
        if colors:    items.append(f"Color: {colors[0]}")
        items += feats[:5]
        if not items: items = [f"Premium {ptype}", "High quality", "Durable"]
        items = items[:5]
        card_x = S//40; base_y = S//8 + S//16
        card_w = S//2 - S//20; card_gap = S//60
        avail_h = S - base_y - S//16
        card_h  = min((avail_h - card_gap*len(items)) // max(1,len(items)), S//8)
        accents = [gold, navy, tuple(min(255,c+40) for c in gold),
                   tuple(max(0,c-20) for c in gold)]
        for i, item in enumerate(items):
            ac = accents[i % len(accents)]
            cy = base_y + i * (card_h + card_gap)
            d.rounded_rectangle([(card_x, cy),(card_x+card_w, cy+card_h)],
                                 radius=6, fill=(255, 255, 255))
            d.rounded_rectangle([(card_x, cy),(card_x+card_w, cy+card_h)],
                                 radius=6, outline=(220, 218, 214), width=1)
            d.rounded_rectangle([(card_x, cy),(card_x+7, cy+card_h)], radius=3, fill=ac)
            nc = card_x + 22; nr = cy + card_h//2
            d.ellipse([(nc-12, nr-12),(nc+12, nr+12)], fill=ac)
            nb = f_body.getbbox(str(i+1))
            d.text((nc-(nb[2]-nb[0])//2, nr-(nb[3]-nb[1])//2),
                   str(i+1), font=f_body, fill=(255, 255, 255))
            d.text((card_x+42, cy+(card_h-f_body.size)//2),
                   item[:42], font=f_body, fill=(42, 42, 42))
        # Product right panel
        if reference_source:
            ref = _prep_ref(reference_source)
            if ref:
                pw = S//2 - S//20; ph = S - S//8 - 4 - S//16
                ref.thumbnail((pw, ph), _Im.LANCZOS)
                rw, rh = ref.size
                px = S//2 + (S//2 - rw)//2; py = S//8 + 4 + (ph - rh)//2
                _paste_shadow(canvas, ref, px, py, offset=16, blur=26, alpha=50)

    # ── MACRO — tight center crop, detail badge ────────────────────────────────
    elif shot_type == "macro":
        top_c = (240, 236, 228); bot_c = (220, 214, 200)
        canvas = _Im.new("RGB", (S, S), top_c)
        _draw_vgrad(canvas, top_c, bot_c)
        if reference_source:
            ref = _prep_ref(reference_source)
            if ref:
                rw, rh = ref.size
                cf = 0.45
                ref = ref.crop((int(rw*cf/2), int(rh*cf/2),
                                rw-int(rw*cf/2), rh-int(rh*cf/2)))
                ref.thumbnail((int(S * 0.92), int(S * 0.92)), _Im.LANCZOS)
                rw, rh = ref.size
                _paste_shadow(canvas, ref, (S-rw)//2, (S-rh)//2, offset=10, blur=18, alpha=40)
        d = _ID.Draw(canvas)
        ac = _hex(pal[0]) if pal else (90, 70, 40)
        badge = "DETAIL VIEW"
        bb = f_body.getbbox(badge)
        bw, bh = bb[2]-bb[0]+28, bb[3]-bb[1]+14
        bx = (S-bw)//2; by = S - bh - 28
        d.rounded_rectangle([(bx, by),(bx+bw, by+bh)], radius=6, fill=ac)
        d.text((bx+14, by+7), badge, font=f_body, fill=(255, 255, 255))
        if mat:
            mb = f_body.getbbox(f"Material: {mat}")
            d.text(((S-(mb[2]-mb[0]))//2, by-bh-10),
                   f"Material: {mat}", font=f_body, fill=(80, 74, 60))

    # ── STYLED HERO — dark premium backdrop, elevated product, gold type ───────
    elif shot_type == "styled_hero":
        top_c = (18, 18, 28); mid_c = (38, 32, 55); bot_c2 = (22, 18, 32)
        canvas = _Im.new("RGB", (S, S), top_c)
        d = _ID.Draw(canvas)
        for y in range(S):
            if y < S//2:
                t = y/(S//2)
                d.line([(0,y),(S-1,y)], fill=tuple(int(top_c[i]+(mid_c[i]-top_c[i])*t) for i in range(3)))
            else:
                t = (y-S//2)/(S//2)
                d.line([(0,y),(S-1,y)], fill=tuple(int(mid_c[i]+(bot_c2[i]-mid_c[i])*t) for i in range(3)))
        hl = _Im.new("RGBA", (S, S), (0, 0, 0, 0))
        _ID.Draw(hl).ellipse([(S//6, S//6),(S*5//6, S*5//6)], fill=(120, 100, 180, 28))
        hl = hl.filter(_IF.GaussianBlur(S//5))
        canvas.paste(hl.convert("RGB"), (0, 0), hl.split()[3])
        gold = _hex(pal[1]) if len(pal) > 1 else (180, 140, 60)
        d = _ID.Draw(canvas)
        d.line([(S//8, S*3//4),(S*7//8, S*3//4)], fill=gold, width=2)
        if reference_source:
            ref = _prep_ref(reference_source)
            if ref:
                ref.thumbnail((int(S*0.72), int(S*0.72)), _Im.LANCZOS)
                rw, rh = ref.size
                _paste_shadow(canvas, ref, (S-rw)//2, int(S*0.50)-rh//2-S//20,
                               offset=22, blur=45, alpha=80)
                d = _ID.Draw(canvas)
        hl_text = ptype.upper()
        hb = f_title.getbbox(hl_text)
        hw = hb[2]-hb[0]; hh = hb[3]-hb[1]
        hx = (S-hw)//2; hy = S*3//4 + 18
        d.text((hx, hy), hl_text, font=f_title, fill=gold)
        d.line([(hx, hy+hh+6),(hx+hw, hy+hh+6)], fill=gold, width=2)
        if colors:
            cb = f_body.getbbox(colors[0].title())
            d.text(((S-(cb[2]-cb[0]))//2, hy+hh+20),
                   colors[0].title(), font=f_body, fill=(180, 175, 200))

    else:
        canvas = _Im.new("RGB", (S, S), (255, 255, 255))
        if reference_source:
            ref = _prep_ref(reference_source)
            if ref:
                ref.thumbnail((int(S*0.80), int(S*0.80)), _Im.LANCZOS)
                rw, rh = ref.size
                _paste_shadow(canvas, ref, (S-rw)//2, (S-rh)//2)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(str(out_path), format="PNG", optimize=True)
    return out_path


def _pil_white_bg_image(product: Dict, reference_source: str, out_path: Path, shot_type: str = "main") -> Path:
    """Alias kept for backward compatibility — delegates to _pil_fallback_image."""
    return _pil_fallback_image(product, reference_source, out_path, shot_type=shot_type)


def _read_bytes(source: str) -> bytes:
    """Read bytes from a local path or HTTP/HTTPS URL."""
    if source.lower().startswith(("http://", "https://")):
        req = Request(source, headers={"User-Agent": "Mozilla/5.0"})
        with urlopen(req, timeout=30) as r:
            return r.read()
    return Path(source).read_bytes()


# ── Reference conditioning ────────────────────────────────────────────────────

def _composite_product_onto_generated(
    generated_path: Path,
    reference_source: str,
    shot_type: str = "lifestyle",
    product: Optional[Dict] = None,
) -> None:
    """Stamp the original reference product onto an AI-generated background.

    This guarantees the real product always appears in the final image,
    regardless of what the AI model hallucinated. Called after every
    successful AI generation.

    Placement per shot type:
      lifestyle    — left-of-center at rule-of-thirds, ~65% frame height
      infographic  — perfectly centered, ~58% of frame
      styled_hero  — centered, ~70% of frame, slight elevation
      macro        — no compositing (close-up detail shot)
      main         — no compositing (AI already generates clean white bg)
    """
    if shot_type in {"main", "macro"}:
        return
    if not reference_source:
        return

    from PIL import Image as _Im, ImageFilter as _IF, ImageDraw as _ID
    try:
        raw = _read_bytes(reference_source)
        ref = _Im.open(BytesIO(raw)).convert("RGBA")
        bg  = _Im.open(str(generated_path)).convert("RGBA")
    except Exception as exc:
        log.warning("Composite: could not open images for %s: %s", generated_path.name, exc)
        return

    # ── Remove studio background and crop to tight product bbox ──────────────
    try:
        ref = _remove_bg(ref)
        ref = _crop_tight(ref)
    except Exception as _bge:
        log.debug("BG removal skipped for composite: %s", _bge)

    W, H = bg.size

    # ── Scale the reference product to fit the target size ────────────────────
    _SCALE = {"lifestyle": 0.62, "infographic": 0.58, "styled_hero": 0.68}
    scale  = _SCALE.get(shot_type, 0.60)
    target = int(min(W, H) * scale)
    ref.thumbnail((target, target), _Im.LANCZOS)
    rw, rh = ref.size

    # ── Choose compositing position ───────────────────────────────────────────
    if shot_type == "lifestyle":
        # Rule of thirds — product sits left-centre
        cx = int(W * 0.42)
        cy = int(H * 0.52)
    else:
        # Centered
        cx = W // 2
        cy = H // 2

    px = cx - rw // 2
    py = cy - rh // 2

    # ── Build a soft drop shadow beneath the product ──────────────────────────
    shadow_layer = _Im.new("RGBA", (W, H), (0, 0, 0, 0))
    shadow_mask  = _Im.new("L", (rw, rh), 0)
    _ID.Draw(shadow_mask).ellipse(
        [(rw // 10, rh * 3 // 4), (rw * 9 // 10, rh + rh // 6)],
        fill=90,
    )
    shadow_mask = shadow_mask.filter(_IF.GaussianBlur(radius=rw // 14))
    shadow_layer.paste(_Im.new("RGB", (rw, rh), (0, 0, 0)), (px, py), shadow_mask)
    shadow_layer = shadow_layer.filter(_IF.GaussianBlur(radius=3))

    # ── Composite: background → shadow → product ──────────────────────────────
    result = bg.copy()
    result = _Im.alpha_composite(result, shadow_layer)
    alpha  = ref.split()[3]
    result.paste(ref, (px, py), alpha)

    try:
        result.convert("RGB").save(str(generated_path), format="PNG")
        log.info("Composited reference product onto %s (%s)", generated_path.name, shot_type)
    except Exception as exc:
        log.warning("Composite: could not save %s: %s", generated_path.name, exc)


def _condition_prompt_with_reference(
    prompt: str,
    reference_description: str,
    shot_type: str = "main",
) -> str:
    """Append reference-product conditioning text to a prompt.

    Note: the actual product is always composited back onto the generated
    image via _composite_product_onto_generated() after generation, so
    this text conditioning is a secondary hint to help the AI model generate
    a matching background and lighting context.
    """
    if not reference_description:
        return prompt

    if shot_type == "main":
        block = (
            "\n\nSTRICT PRODUCT REFERENCE: preserve the exact shape, proportions, "
            "materials, texture, stitching, and color palette of the reference product precisely. "
            f"Product description: {reference_description}"
        )
    elif shot_type == "macro":
        block = (
            f"\n\nProduct reference (for material/texture accuracy only): {reference_description}"
        )
    else:
        # lifestyle, infographic, styled_hero — generate scene/background that
        # suits this product; the actual product will be composited in afterward
        block = (
            f"\n\nProduct being featured: {reference_description}. "
            "Generate ONLY the fitting scene, background, props, and lighting for this product. "
            "Do NOT draw the product itself, or any object resembling it, anywhere in the frame — "
            "leave that focal area visually empty (just out-of-focus background continuing through it). "
            "The real product photo will be composited into that empty area afterward, so an "
            "AI-drawn stand-in there would only create a distorted double-image. "
            "Do not include any human hands, fingers, or body parts in the scene."
        )
    return prompt + block


# ── Single-shot generator ─────────────────────────────────────────────────────

def _generate_one_shot(
    prompt: str,
    provider: str,
    out_path: Path,
    reference_source: str,
    product: Dict,
    reference_description: str = "",
    shot_type: str = "main",
) -> Path:
    """Generate one image shot and save to out_path.

    Tries the requested provider first, falls back to PIL on any failure.
    Returns the path where the image was saved.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # The pure white-background "main" compliance shot must always be produced
    # locally with Pillow from the real reference photo — never via an AI image
    # API (which can hallucinate product details). Short-circuit before any
    # provider call.
    if shot_type == "main":
        return _pil_fallback_image(product, reference_source, out_path, shot_type=shot_type)

    conditioned = _condition_prompt_with_reference(prompt, reference_description, shot_type=shot_type)
    fal_prompt  = _truncate_prompt(conditioned, max_chars=1000)

    p = provider.strip().lower()

    def _save_and_composite(img_bytes: bytes) -> Path:
        """Write generated bytes, then stamp the real product on top."""
        out_path.write_bytes(img_bytes)
        _composite_product_onto_generated(out_path, reference_source, shot_type, product)
        return out_path

    # ── AUTOMATIC1111 / ComfyUI local path ───────────────────────────────────
    if p == "a1111":
        try:
            img_bytes = _a1111_generate_image_bytes(fal_prompt)
            log.info("A1111 image saved: %s (%d bytes)", out_path.name, len(img_bytes))
            return _save_and_composite(img_bytes)
        except Exception as exc:
            log.warning("A1111 image failed for %s: %s — PIL fallback", out_path.name, exc)

    # ── fal.ai path ───────────────────────────────────────────────────────────
    elif p == "fal":
        if not FAL_AVAILABLE:
            log.warning("FAL_KEY not set, falling back to PIL for %s", out_path.name)
        else:
            try:
                img_bytes = _fal_generate_image_bytes(fal_prompt)
                log.info("fal.ai image saved: %s (%d bytes)", out_path.name, len(img_bytes))
                return _save_and_composite(img_bytes)
            except Exception as exc:
                log.warning("fal.ai image failed for %s: %s — PIL fallback", out_path.name, exc)

    # ── PIL path (explicit or fallback) ──────────────────────────────────────
    return _pil_fallback_image(product, reference_source, out_path, shot_type=shot_type)


# ── 4-shot orchestrator ───────────────────────────────────────────────────────

def generateProductImages(
    product: Dict,
    reference_image: str = "",
    config: Optional[Dict] = None,
    reference_description: str = "",
) -> Dict[str, Path]:
    """Generate a set of Amazon-compliant product images (up to 4 shots).

    Args:
        product: Product analysis dict with keys: product_type, category,
                 material, colors, features, usage, style, sku (optional).
        reference_image: Local file path or HTTPS URL of the reference product photo.
                         Downloaded once and reused for all shots.
        config: Optional overrides:
                  shots    — tuple/list of shot types to generate (default: all 4)
                  out_dir  — Path for output files
                  provider — "fal" | "a1111" | "pil" (default: AI_PROVIDER env)
                  prefix   — filename prefix (default: sku or product_type)
        reference_description: Pre-computed text description of the reference image
                               (pass the output of _reference_conditioning_text() from
                               amazon_template_autofill_web). When empty, reference
                               conditioning is skipped.

    Returns:
        Dict mapping shot_type → absolute Path of saved PNG.
        Failed shots fall back to PIL; no shot raises an exception.
    """
    cfg = config or {}
    shots    = tuple(cfg.get("shots", SHOT_TYPES))
    out_dir  = Path(cfg.get("out_dir", os.getenv("AI_GENERATED_IMAGE_DIR", "generated_amazon_images/ai")))
    provider = str(cfg.get("provider", AI_PROVIDER)).strip().lower()
    prefix   = str(cfg.get("prefix", "") or product.get("sku") or product.get("product_type", "product") or "product")
    prefix   = prefix.replace(" ", "_").replace("/", "_")[:60]

    if provider not in _VALID_PROVIDERS:
        raise ValueError(f"Unknown provider '{provider}'. Must be one of {sorted(_VALID_PROVIDERS)}.")

    out_dir.mkdir(parents=True, exist_ok=True)

    # Download / resolve reference image once for all shots
    ref_local: str = ""
    if reference_image:
        if reference_image.lower().startswith(("http://", "https://")):
            try:
                raw = _read_bytes(reference_image)
                tmp_path = out_dir / f"{prefix}_ref.jpg"
                tmp_path.write_bytes(raw)
                ref_local = str(tmp_path)
                log.info("Reference image downloaded: %s", tmp_path.name)
            except Exception as exc:
                log.warning("Could not download reference image %s: %s", reference_image, exc)
        else:
            ref_local = reference_image

    valid_shots = [s for s in shots if s in SHOT_TYPES]
    for skipped in set(shots) - set(valid_shots):
        log.warning("Unknown shot_type '%s' — skipping.", skipped)

    def _run_shot(shot_type: str) -> tuple[str, Path]:
        out_path = out_dir / f"{prefix}_{shot_type}.png"
        try:
            prompt = _build_prompt(product, shot_type)
            log.debug("Prompt for %s/%s: %s", prefix, shot_type, prompt[:120])
            path = _generate_one_shot(
                prompt=prompt,
                provider=provider,
                out_path=out_path,
                reference_source=ref_local,
                product=product,
                reference_description=reference_description,
                shot_type=shot_type,
            )
            return shot_type, path
        except Exception as exc:
            log.error("Shot %s failed for %s: %s — PIL fallback", shot_type, prefix, exc)
            try:
                return shot_type, _pil_fallback_image(product, ref_local, out_path, shot_type=shot_type)
            except Exception as pil_exc:
                log.error("PIL fallback also failed for %s/%s: %s", prefix, shot_type, pil_exc)
                raise

    results: Dict[str, Path] = {}
    # Generate all shots in parallel — each CF call is independent I/O
    max_workers = min(len(valid_shots), int(os.getenv("PRODUCT_IMAGE_WORKERS", "4")))
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(_run_shot, s): s for s in valid_shots}
        for future in as_completed(futures):
            try:
                shot_type, path = future.result()
                results[shot_type] = path
                log.info("Shot %s complete: %s", shot_type, path.name)
            except Exception as exc:
                log.error("Shot %s could not be generated: %s", futures[future], exc)

    return results
