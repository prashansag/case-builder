"""Measured text layout and contrast safeguards for generated presentations."""
from functools import lru_cache
import math
from pathlib import Path

from PIL import ImageFont
from pptx.dml.color import RGBColor
from pptx.util import Pt


@lru_cache(maxsize=64)
def font_metrics(size, bold=False):
    name = "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"
    return ImageFont.truetype(str(Path(__file__).parent / "assets" / "fonts" / name), round(size * 4))


def wrap_text(text, width_points, size, bold=False):
    font = font_metrics(size, bold)
    limit = max(1, width_points - 6) * 4
    lines = []
    for paragraph in str(text or "").split("\n"):
        line = ""
        for word in paragraph.split():
            candidate = f"{line} {word}".strip()
            if font.getlength(candidate) <= limit:
                line = candidate
                continue
            if line:
                lines.append(line)
                line = ""
            # Long URLs and unbroken words must wrap too.
            for char in word:
                if line and font.getlength(line + char) > limit:
                    lines.append(line)
                    line = ""
                line += char
        lines.append(line)
    return lines


def fit_text(text, width, height, size, bold=False):
    width_points, height_points = width / 12700, height / 12700
    minimum = 24 if size >= 26 else min(size, 12)
    for candidate in range(size, minimum - 1, -1):
        lines = wrap_text(text, width_points, candidate, bold)
        capacity = max(1, math.floor((height_points - 2) / (candidate * 1.25)))
        if len(lines) <= capacity:
            return "\n".join(lines), "", candidate
    return "\n".join(lines[:capacity]), "\n".join(lines[capacity:]), minimum


def luminance(color):
    channels = [value / 255 for value in color]
    return sum(
        weight * (value / 12.92 if value <= 0.04045 else ((value + .055) / 1.055) ** 2.4)
        for weight, value in zip((.2126, .7152, .0722), channels)
    )


def contrast_ratio(first, second):
    values = sorted((luminance(first), luminance(second)))
    return (values[1] + .05) / (values[0] + .05)


def readable_color(color, background):
    if contrast_ratio(color, background) >= 4.5:
        return color
    return max(
        (RGBColor(255, 255, 255), RGBColor(17, 24, 39)),
        key=lambda option: contrast_ratio(option, background),
    )


def text_background(slide, left, top, width, height):
    background = slide.background.fill.fore_color.rgb
    for shape in slide.shapes:
        if (shape.left <= left and shape.top <= top
                and shape.left + shape.width >= left + width
                and shape.top + shape.height >= top + height
                and shape.fill.type == 1):
            background = shape.fill.fore_color.rgb
    return background


def style_frame(frame, size):
    frame.margin_left = frame.margin_right = 0
    frame.margin_top = frame.margin_bottom = 0
    frame.word_wrap = True
    for paragraph in frame.paragraphs:
        paragraph.space_before = Pt(0)
        paragraph.space_after = Pt(0)
        paragraph.line_spacing = Pt(size * 1.25)