#!/usr/bin/env python3
"""Non-destructively restyle an existing thesis-defense PPTX.

The goal is to optimize color and formatting while preserving the original
content: all text strings, figures, charts, images, and existing slide objects
remain in the deck. The script inserts background/accent geometry behind the
original objects and normalizes text colors/fonts for a more refined academic
look.
"""

from __future__ import annotations

import argparse
from pathlib import Path


PALETTE_HEX = {
    "ink": "14172F",
    "navy": "1E2361",
    "midnight": "0F1638",
    "paper": "F7F4EC",
    "surface": "FFFFFF",
    "mist": "EEF1F6",
    "gold": "C79A3B",
    "cyan": "42B8C8",
    "red": "B9554D",
    "graphite": "374151",
}

PALETTE = {}
MSO_SHAPE = None
PP_ALIGN = None
Inches = None
Pt = None

TITLE_FONT = "Aptos Display"
BODY_FONT = "Microsoft YaHei"


def add_rect(slide, x, y, w, h, fill, transparency=0, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.fill.transparency = transparency
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    send_to_back(shape)
    return shape


def add_line(slide, x, y, w, h, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    send_to_back(shape, after_background=True)
    return shape


def send_to_back(shape, after_background=False):
    """Move a newly-added shape behind existing content.

    PowerPoint stores z-order in the shape tree. Inserting near the start keeps
    restyling geometry below existing slide content while retaining the slide's
    required non-shape XML records.
    """

    element = shape._element
    tree = element.getparent()
    tree.remove(element)
    target_index = 3 if after_background else 2
    tree.insert(min(target_index, len(tree)), element)


def classify_slide(slide, idx):
    text = " ".join(
        shape.text.strip() for shape in slide.shapes if getattr(shape, "has_text_frame", False) and shape.text
    )
    image_count = sum(1 for shape in slide.shapes if shape.shape_type == 13)
    if idx == 0 or len(text) < 90:
        return "paper_title"
    if image_count >= 3:
        return "figure_gallery"
    if any(token in text for token in ("小结", "总结", "致谢", "研究内容", "Conclusion", "Summary")):
        return "dark_section_marker"
    if any(token in text for token in ("PR", "曲线", "效率", "结果", "数据", "性能")):
        return "data_focus"
    return "split_research_card"


def apply_background(slide, idx, mode, width, height):
    rhythm = classify_slide(slide, idx)
    dark = mode == "dark" or rhythm == "dark_section_marker"
    base = PALETTE["midnight"] if dark else PALETTE["paper"]
    add_rect(slide, 0, 0, width, height, base)

    if rhythm == "paper_title":
        add_rect(slide, Inches(0.36), Inches(0.28), width - Inches(0.72), height - Inches(0.56), PALETTE["surface"], 8)
        add_line(slide, Inches(0.52), Inches(1.02), Inches(2.1), Inches(0.04), PALETTE["gold"])
        add_line(slide, width - Inches(2.22), height - Inches(0.7), Inches(1.7), Inches(0.035), PALETTE["cyan"])
    elif rhythm == "figure_gallery":
        add_rect(slide, Inches(0.24), Inches(0.28), width - Inches(0.48), height - Inches(0.56), PALETTE["surface"], 3)
        add_line(slide, Inches(0.24), Inches(0.28), Inches(0.08), height - Inches(0.56), PALETTE["gold"])
        add_line(slide, Inches(0.48), height - Inches(0.48), width - Inches(0.96), Inches(0.025), PALETTE["mist"])
    elif rhythm == "data_focus":
        add_rect(slide, Inches(0.32), Inches(0.42), width - Inches(0.64), height - Inches(0.78), PALETTE["surface"], 0)
        add_line(slide, Inches(0.32), Inches(0.42), width - Inches(0.64), Inches(0.035), PALETTE["cyan"])
        add_line(slide, Inches(0.32), Inches(0.52), Inches(1.25), Inches(0.035), PALETTE["gold"])
    elif rhythm == "dark_section_marker":
        add_rect(slide, Inches(0.44), Inches(0.52), width - Inches(0.88), height - Inches(1.04), PALETTE["navy"], 15)
        add_line(slide, Inches(0.72), Inches(1.05), Inches(1.8), Inches(0.055), PALETTE["gold"])
        add_line(slide, width - Inches(2.52), height - Inches(0.98), Inches(1.8), Inches(0.035), PALETTE["cyan"])
    else:
        add_rect(slide, Inches(0.28), Inches(0.34), width - Inches(0.56), height - Inches(0.68), PALETTE["surface"], 0)
        add_line(slide, Inches(0.28), Inches(0.34), width - Inches(0.56), Inches(0.04), PALETTE["navy"])
        add_line(slide, Inches(0.28), Inches(0.44), Inches(1.1), Inches(0.035), PALETTE["gold"])

    add_footer(slide, idx, width, height, dark)
    return rhythm


def add_footer(slide, idx, width, height, dark):
    footer_color = PALETTE["navy"] if not dark else PALETTE["midnight"]
    add_line(slide, Inches(0.34), height - Inches(0.28), width - Inches(0.68), Inches(0.035), footer_color)
    number = slide.shapes.add_textbox(width - Inches(0.72), height - Inches(0.36), Inches(0.32), Inches(0.16))
    number.text_frame.text = f"{idx + 1:02d}"
    p = number.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    run = p.runs[0]
    run.font.name = TITLE_FONT
    run.font.size = Pt(7)
    run.font.color.rgb = PALETTE["gold"]


def normalize_text(slide, rhythm):
    dark = rhythm == "dark_section_marker"
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        for paragraph in shape.text_frame.paragraphs:
            is_title = paragraph.level == 0 and len(paragraph.text.strip()) <= 42
            for run in paragraph.runs:
                if not run.text:
                    continue
                run.font.name = TITLE_FONT if is_title else BODY_FONT
                if run.font.size is None:
                    run.font.size = Pt(20 if is_title else 11)
                if is_title:
                    run.font.bold = True
                    run.font.color.rgb = PALETTE["paper"] if dark else PALETTE["ink"]
                else:
                    run.font.color.rgb = PALETTE["mist"] if dark else PALETTE["graphite"]


def load_pptx_dependencies():
    """Load python-pptx only when a deck is actually being processed."""

    global Inches, MSO_SHAPE, PALETTE, PP_ALIGN, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE as _MSO_SHAPE
    from pptx.enum.text import PP_ALIGN as _PP_ALIGN
    from pptx.util import Inches as _Inches, Pt as _Pt

    MSO_SHAPE = _MSO_SHAPE
    PP_ALIGN = _PP_ALIGN
    Inches = _Inches
    Pt = _Pt
    PALETTE = {
        name: RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))
        for name, value in PALETTE_HEX.items()
    }


def restyle(input_path: Path, output_path: Path, mode: str):
    from pptx import Presentation

    load_pptx_dependencies()
    prs = Presentation(str(input_path))
    width, height = prs.slide_width, prs.slide_height
    for idx, slide in enumerate(prs.slides):
        rhythm = apply_background(slide, idx, mode, width, height)
        normalize_text(slide, rhythm)
    prs.save(str(output_path))


def parse_args():
    parser = argparse.ArgumentParser(description="Restyle a thesis-defense PPTX without changing content.")
    parser.add_argument("input", type=Path, help="Source .pptx file")
    parser.add_argument("output", type=Path, help="Destination .pptx file")
    parser.add_argument("--mode", choices=("light", "dark"), default="light", help="Overall academic color mode")
    return parser.parse_args()


def main():
    args = parse_args()
    if not args.input.exists():
        raise SystemExit(f"Input file does not exist: {args.input}")
    if args.input.suffix.lower() != ".pptx" or args.output.suffix.lower() != ".pptx":
        raise SystemExit("Both input and output paths must be .pptx files.")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    restyle(args.input, args.output, args.mode)
    print(f"Saved academic restyle: {args.output}")


if __name__ == "__main__":
    main()
