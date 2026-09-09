# -*- coding: utf-8 -*-
"""Generate brand visual assets for physics PPT deliveries.

Inputs (repo samples, no external network calls):
  PPTX/公众号图标.png   raw circular WeChat-channel icon on a white background
  PPTX/背景图.png       AI-generated 16:9 physics-themed background that still
                        carries a baked-in, low-quality icon at the top right

Outputs (regenerable, deterministic):
  assets/brand/sciman-icon.png        RGBA icon, tight circular alpha mask
  assets/brand/sciman-icon-shadow.png RGBA icon on a padded canvas with a soft
                                      baked drop shadow (for direct placement)
  assets/brand/bg-16x9.jpg            1920x1080 background: baked icon patched
                                      out, text-safe darkening applied, new
                                      icon composited at the top right

Usage:
  python tools/generate_brand_assets.py [--check]

--check verifies that the outputs exist and are loadable, without rewriting.
"""

from __future__ import annotations

import argparse
import hashlib
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageEnhance, ImageFilter

REPO_ROOT = Path(__file__).resolve().parents[1]
ICON_SOURCE = REPO_ROOT / "PPTX" / "公众号图标.png"
BG_SOURCE = REPO_ROOT / "PPTX" / "背景图.png"
ASSET_DIR = REPO_ROOT / "assets" / "brand"

ICON_OUT = ASSET_DIR / "sciman-icon.png"
ICON_SHADOW_OUT = ASSET_DIR / "sciman-icon-shadow.png"
ICON_WATERMARK_OUT = ASSET_DIR / "sciman-icon-watermark.png"
BG_OUT = ASSET_DIR / "bg-16x9.jpg"

ICON_SIZE = 1024
SHADOW_CANVAS = 1280
BG_SIZE = (1920, 1080)
# Opaque watermark variant for pages whose own text reaches the icon corner.
WATERMARK_ALPHA = 0.45

# Baked icon location in BG_SOURCE pixel coordinates (1672x941), measured from
# the delivered sample: circle centre and radius with a generous margin.
BAKED_ICON_CENTRE = (1590, 78)
BAKED_ICON_RADIUS = 58

# Final icon placement on the 1920x1080 background (circle centre, diameter).
BG_ICON_CENTRE = (1842, 72)
BG_ICON_DIAMETER = 104


def build_icon_shadow(icon_rgba: Image.Image) -> Image.Image:
    canvas = Image.new("RGBA", (SHADOW_CANVAS, SHADOW_CANVAS), (0, 0, 0, 0))
    margin = (SHADOW_CANVAS - ICON_SIZE) // 2
    shadow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    drawer = ImageDraw.Draw(shadow)
    offset = 18
    drawer.ellipse(
        (
            margin + 10,
            margin + offset,
            margin + ICON_SIZE - 10,
            margin + ICON_SIZE + offset,
        ),
        fill=(6, 20, 16, 110),
    )
    shadow = shadow.filter(ImageFilter.GaussianBlur(26))
    canvas = Image.alpha_composite(canvas, shadow)
    canvas.paste(icon_rgba, (margin, margin), icon_rgba)
    return canvas


def make_transparent_icon() -> Image.Image:
    source = Image.open(ICON_SOURCE).convert("RGB")
    gray = source.convert("L")
    # Non-white coverage defines the circular badge bounds; the sample has a
    # pure white margin so a simple threshold is a reliable circle detector.
    coverage = gray.point(lambda v: 255 if v < 238 else 0)
    bbox = coverage.getbbox()
    if bbox is None:
        raise SystemExit("icon source has no visible content")
    left, top, right, bottom = bbox
    side = max(right - left, bottom - top)
    centre_x = (left + right) // 2
    centre_y = (top + bottom) // 2
    half = side // 2
    square = (
        max(0, centre_x - half),
        max(0, centre_y - half),
        min(source.width, centre_x + half),
        min(source.height, centre_y + half),
    )
    side = min(square[2] - square[0], square[3] - square[1])
    square = (
        square[0],
        square[1],
        square[0] + side,
        square[1] + side,
    )
    icon = source.crop(square).resize((ICON_SIZE, ICON_SIZE), Image.LANCZOS)

    # Supersampled antialiased circular mask, inset 3px to cut the white
    # fringe left by the source's own edge anti-aliasing.
    super_size = ICON_SIZE * 2
    mask = Image.new("L", (super_size, super_size), 0)
    ImageDraw.Draw(mask).ellipse((6, 6, super_size - 6, super_size - 6), fill=255)
    mask = mask.resize((ICON_SIZE, ICON_SIZE), Image.LANCZOS)

    result = icon.convert("RGBA")
    result.putalpha(mask)
    return result


def patch_region(base: Image.Image, rect: tuple[int, int, int, int]) -> Image.Image:
    """Fill `rect` by per-row horizontal interpolation between the pixel
    columns just outside the patch, then blend it in through a feathered
    mask so no seam is visible in the smooth background glow."""
    left, top, right, bottom = rect
    patch = Image.new("RGB", (right - left, bottom - top))
    pixels = patch.load()
    source = base.load()
    span = right - left - 1
    for y in range(top, bottom):
        left_sample = source[max(left - 14, 0), y]
        right_sample = source[min(right + 14, base.width - 1), y]
        for x in range(left, right):
            t = (x - left) / span
            pixels[x - left, y - top] = tuple(
                round(left_sample[c] + (right_sample[c] - left_sample[c]) * t)
                for c in range(3)
            )
    patch = patch.filter(ImageFilter.GaussianBlur(6))
    mask = Image.new("L", patch.size, 0)
    feather = 10
    ImageDraw.Draw(mask).rectangle(
        (feather, feather, patch.width - feather, patch.height - feather), fill=255
    )
    mask = mask.filter(ImageFilter.GaussianBlur(feather))
    base.paste(patch, (left, top), mask)
    return base


def darken_for_text(base: Image.Image) -> Image.Image:
    base = ImageEnhance.Brightness(base).enhance(0.88)
    # Extra darkening towards the bottom third keeps white body text readable
    # over the bright light waves of the source artwork.
    gradient = Image.new("L", (1, base.height))
    for y in range(base.height):
        if y <= 540:
            value = 0
        else:
            value = round(95 * (y - 540) / (base.height - 540))
        gradient.putpixel((0, y), value)
    gradient = gradient.resize(base.size)
    black = Image.new("RGB", base.size, (4, 12, 24))
    return Image.composite(black, base, gradient)


def make_background(icon_shadow: Image.Image) -> Image.Image:
    base = Image.open(BG_SOURCE).convert("RGB").resize(BG_SIZE, Image.LANCZOS)
    scale_x = BG_SIZE[0] / Image.open(BG_SOURCE).width
    scale_y = BG_SIZE[1] / Image.open(BG_SOURCE).height
    centre = (
        round(BAKED_ICON_CENTRE[0] * scale_x),
        round(BAKED_ICON_CENTRE[1] * scale_y),
    )
    radius = round(BAKED_ICON_RADIUS * ((scale_x + scale_y) / 2)) + 45
    rect = (
        max(0, centre[0] - radius),
        max(0, centre[1] - radius),
        min(BG_SIZE[0] - 1, centre[0] + radius),
        min(BG_SIZE[1] - 1, centre[1] + radius),
    )
    base = patch_region(base, rect)
    base = darken_for_text(base)

    target = BG_ICON_DIAMETER * (SHADOW_CANVAS / ICON_SIZE)
    scaled = icon_shadow.resize((round(target), round(target)), Image.LANCZOS)
    pad = (target - BG_ICON_DIAMETER) / 2
    position = (
        round(BG_ICON_CENTRE[0] - BG_ICON_DIAMETER / 2 - pad),
        round(BG_ICON_CENTRE[1] - BG_ICON_DIAMETER / 2 - pad),
    )
    base.paste(scaled, position, scaled)
    return base


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="verify outputs only")
    args = parser.parse_args()

    if args.check:
        missing = [
            p for p in (ICON_OUT, ICON_SHADOW_OUT, ICON_WATERMARK_OUT, BG_OUT)
            if not p.is_file()
        ]
        for path in missing:
            print(f"MISSING {path}")
        for path in (ICON_OUT, ICON_SHADOW_OUT, ICON_WATERMARK_OUT, BG_OUT):
            if path.is_file():
                with Image.open(path) as image:
                    image.verify()
                print(f"OK {path.name} {path.stat().st_size} bytes")
        return 1 if missing else 0

    for source in (ICON_SOURCE, BG_SOURCE):
        if not source.is_file():
            raise SystemExit(f"missing input sample: {source}")

    ASSET_DIR.mkdir(parents=True, exist_ok=True)

    icon = make_transparent_icon()
    icon.save(ICON_OUT)
    icon_shadow = build_icon_shadow(icon)
    icon_shadow.save(ICON_SHADOW_OUT)

    watermark = icon.copy()
    alpha = watermark.getchannel("A").point(lambda v: round(v * WATERMARK_ALPHA))
    watermark.putalpha(alpha)
    watermark.save(ICON_WATERMARK_OUT)

    background = make_background(icon_shadow)
    background.save(BG_OUT, quality=92, optimize=True)

    for path in (ICON_OUT, ICON_SHADOW_OUT, ICON_WATERMARK_OUT, BG_OUT):
        print(f"{path} {path.stat().st_size} bytes sha256={sha256(path)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
