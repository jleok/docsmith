#!/usr/bin/env python3
"""Extract the visual spec of a resume so the cover letter can match it.

Usage:  python3 inspect_format.py <resume.docx|resume.pdf>

Prints a FORMAT SPEC block: body font, body size, heading/name size,
page size and margins, all in the units docx-js wants (DXA, half-points).
Anything it cannot determine is printed as UNKNOWN so you notice and
fall back deliberately instead of silently guessing.
"""

import re
import sys
import subprocess
import zipfile
from collections import Counter

TWIPS_PER_INCH = 1440
PT_PER_INCH = 72


def dxa(inches):
    return int(round(inches * TWIPS_PER_INCH))


def report(spec):
    print("=== FORMAT SPEC ===")
    for key in ("source", "body_font", "body_size_pt", "name_size_pt",
                "heading_size_pt", "page_w_dxa", "page_h_dxa",
                "margin_top_dxa", "margin_bottom_dxa",
                "margin_left_dxa", "margin_right_dxa"):
        print(f"{key}: {spec.get(key, 'UNKNOWN')}")
    if spec.get("notes"):
        print("notes: " + "; ".join(spec["notes"]))
    print("=== END FORMAT SPEC ===")


# ---------------------------------------------------------------- docx

def inspect_docx(path):
    spec = {"source": "docx", "notes": []}
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        doc = z.read("word/document.xml").decode("utf8", "ignore")
        styles = z.read("word/styles.xml").decode("utf8", "ignore") \
            if "word/styles.xml" in names else ""

    # Fonts: runs that declare one, plus the document default.
    fonts = Counter(re.findall(r'w:ascii="([^"]+)"', doc))
    default_font = None
    m = re.search(r"<w:docDefaults>.*?</w:docDefaults>", styles, re.S)
    if m:
        d = re.search(r'w:ascii="([^"]+)"', m.group(0))
        if d:
            default_font = d.group(1)
    if fonts:
        spec["body_font"] = fonts.most_common(1)[0][0]
        if default_font and default_font != spec["body_font"]:
            spec["notes"].append(
                f"styles.xml default is {default_font}; body runs use "
                f"{spec['body_font']}")
    elif default_font:
        spec["body_font"] = default_font
        spec["notes"].append("no run-level fonts; using docDefaults")

    # Sizes are half-points. Most common = body. Largest = the name.
    sizes = Counter(int(s) for s in re.findall(r'<w:sz w:val="(\d+)"', doc))
    if sizes:
        spec["body_size_pt"] = sizes.most_common(1)[0][0] / 2
        spec["name_size_pt"] = max(sizes) / 2
        ranked = sorted(sizes, reverse=True)
        if len(ranked) > 2:
            spec["heading_size_pt"] = ranked[1] / 2
        spec["notes"].append(
            "sizes seen (pt): " +
            ", ".join(f"{k/2:g}x{v}" for k, v in sizes.most_common()))
    else:
        d = re.search(r'<w:sz w:val="(\d+)"', styles)
        if d:
            spec["body_size_pt"] = int(d.group(1)) / 2
            spec["notes"].append("size from docDefaults only")

    m = re.search(r'<w:pgSz w:w="(\d+)" w:h="(\d+)"', doc)
    if m:
        spec["page_w_dxa"], spec["page_h_dxa"] = int(m.group(1)), int(m.group(2))

    m = re.search(r"<w:pgMar[^>]*/>", doc)
    if m:
        for attr, key in (("top", "margin_top_dxa"),
                          ("bottom", "margin_bottom_dxa"),
                          ("left", "margin_left_dxa"),
                          ("right", "margin_right_dxa")):
            v = re.search(rf'w:{attr}="(-?\d+)"', m.group(0))
            if v:
                spec[key] = int(v.group(1))
    return spec


# ----------------------------------------------------------------- pdf

# PDF generators substitute metric-compatible clones for the common Office
# fonts. The user named the font on the left, so report that.
CLONES = {
    "carlito": "Calibri",
    "caladea": "Cambria",
    "liberationsans": "Arial",
    "liberationserif": "Times New Roman",
    "liberationmono": "Courier New",
    "nimbussans": "Helvetica",
    "nimbusroman": "Times New Roman",
    "timesnewromanpsmt": "Times New Roman",
    "arialmt": "Arial",
}


def strip_subset(name):
    """BCDEEE+Calibri-Bold -> Calibri.  Keeps the family, drops the weight."""
    name = re.sub(r"^[A-Z]{6}\+", "", name)
    name = re.split(r"[-,]", name)[0]
    name = re.sub(r"(MT|PS|Std|Pro)$", "", name) or name
    return CLONES.get(re.sub(r"[\s_]", "", name).lower(), name)


def inspect_pdf(path):
    spec = {"source": "pdf", "notes": []}
    try:
        import pdfplumber
    except ImportError:
        spec["notes"].append("pdfplumber missing; run pdffonts by hand")
        return spec

    with pdfplumber.open(path) as pdf:
        page = pdf.pages[0]
        chars = page.chars
        if not chars:
            spec["notes"].append("no extractable text (scanned?); "
                                 "render it and read the image instead")
            return spec

        fam = Counter(strip_subset(c["fontname"]) for c in chars)
        spec["body_font"] = fam.most_common(1)[0][0]
        if len(fam) > 1:
            spec["notes"].append(
                "font families: " +
                ", ".join(f"{k}x{v}" for k, v in fam.most_common(4)))

        sizes = Counter(round(c["size"] * 2) / 2 for c in chars)
        spec["body_size_pt"] = sizes.most_common(1)[0][0]
        spec["name_size_pt"] = max(sizes)
        ranked = sorted(sizes, reverse=True)
        if len(ranked) > 2:
            spec["heading_size_pt"] = ranked[1]
        spec["notes"].append(
            "sizes seen (pt): " +
            ", ".join(f"{k:g}x{v}" for k, v in sizes.most_common(5)))

        spec["page_w_dxa"] = dxa(page.width / PT_PER_INCH)
        spec["page_h_dxa"] = dxa(page.height / PT_PER_INCH)

        # Only the left edge and the top edge are measurable. The right
        # margin is not, because the last line of a paragraph stops short of
        # it; the bottom margin is not, because the page ends wherever the
        # text ran out. Measure the two that are real and mirror them, which
        # is what resumes do anyway.
        def snap(points):
            """Round to the nearest quarter inch. Resume margins are round
            numbers, and the glyph bounding box sits a hair inside the real
            margin, so snapping recovers the number the author typed."""
            return dxa(round(points / PT_PER_INCH * 4) / 4)

        left = snap(min(c["x0"] for c in chars))
        top = snap(min(c["top"] for c in chars))
        spec["margin_left_dxa"] = spec["margin_right_dxa"] = left
        spec["margin_top_dxa"] = spec["margin_bottom_dxa"] = top
        spec["notes"].append(
            "margins measured from the left and top text edges and mirrored, "
            "snapped to the nearest 0.25in; right and bottom are not "
            "measurable from a pdf")
    return spec


def main():
    if len(sys.argv) != 2:
        sys.exit("usage: inspect_format.py <resume.docx|resume.pdf>")
    path = sys.argv[1]
    low = path.lower()
    if low.endswith(".docx"):
        spec = inspect_docx(path)
    elif low.endswith(".pdf"):
        spec = inspect_pdf(path)
    elif low.endswith(".doc"):
        sys.exit(
            "Legacy .doc is not supported directly. Convert it first with "
            "LibreOffice:\n"
            "  soffice --headless --convert-to docx " + path + "\n"
            "If soffice is not installed: "
            "macOS  brew install --cask libreoffice | "
            "Debian/Ubuntu  sudo apt install libreoffice")
    else:
        sys.exit("expected a .docx or .pdf")
    report(spec)


if __name__ == "__main__":
    main()
