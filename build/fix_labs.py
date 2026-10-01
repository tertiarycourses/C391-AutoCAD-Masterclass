"""C391 labs/: rewrite the WSQ course identity in the Markdown and in every lab PDF.

The PDFs have no buildable source here (README.pdf is ReportLab with SUBSET fonts,
lab-NN-instructions.pdf is LibreOffice output with CID text), so each affected text
block is redacted and re-set in place at the same origin, size, font and colour —
the rest of the page is untouched.
"""
import re
from pathlib import Path
import fitz

OLD = "WSQ Technical Drawing with AutoCAD"
NEW = "AutoCAD Masterclass"
FONTS = {False: "/System/Library/Fonts/Supplemental/Arial.ttf",
         True: "/System/Library/Fonts/Supplemental/Arial Bold.ttf"}
TRIGGER = re.compile(r"WSQ|TGS-\d|\([AK]\d, [AK]\d\)|Version v13")


def text(s):
    s = s.replace(f"{OLD} (TGS-2021005538)", f"{NEW} (C391)")
    s = s.replace(OLD, NEW)
    s = s.replace("TGS-2021005538", "C391")
    s = s.replace("WSQ COURSEWARE | Activity Resource", "C391 | Activity Resource")
    s = s.replace("Version v13 · 21 August 2026", "Version v1.0 · 1 October 2026")
    s = s.replace("Version v13", "Version v1.0")
    s = re.sub(r"\s*\([AK]\d, [AK]\d\)", "", s)
    return s


def fix_md():
    n = 0
    for f in Path("labs").rglob("*.md"):
        old = f.read_text(encoding="utf-8")
        new = text(old)
        if new != old:
            f.write_text(new, encoding="utf-8")
            n += 1
    return n


def fix_pdf(f):
    doc = fitz.open(f)
    changed = False
    for page in doc:
        jobs = []
        for b in page.get_text("dict")["blocks"]:
            lines = b.get("lines", [])
            if not lines:
                continue
            spans = [s for l in lines for s in l["spans"] if s["text"].strip()]
            joined = " ".join("".join(s["text"] for s in l["spans"]).strip() for l in lines)
            joined = re.sub(r"\s{3,}", "  ", joined)
            if not TRIGGER.search(joined):
                continue
            fonts = {(s["font"], round(s["size"], 1), s["color"]) for s in spans}
            assert len(fonts) == 1, (f, joined, fonts)
            font, size, color = fonts.pop()
            jobs.append((fitz.Rect(b["bbox"]), lines[0]["spans"][0]["origin"],
                         text(joined).strip(), "Bold" in font, size, color))
        for rect, *_ in jobs:
            page.add_redact_annot(rect + (0, 1, 0, -1), fill=(1, 1, 1))
        if jobs:
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                                  graphics=fitz.PDF_REDACT_LINE_ART_NONE)
        for rect, origin, new, bold, size, color in jobs:
            if not new:
                continue
            rgb = tuple(((color >> k) & 255) / 255 for k in (16, 8, 0))
            fname = "ArB" if bold else "ArR"
            width = fitz.Font(fontfile=FONTS[bold]).text_length(new, fontsize=size)
            avail = page.rect.width - origin[0] - 48
            assert width <= avail, (f, new, width, avail)
            page.insert_text(origin, new, fontsize=size, fontname=fname,
                             fontfile=FONTS[bold], color=rgb)
            changed = True
    if changed:
        doc.save(f + ".tmp", garbage=3, deflate=True)
        Path(f + ".tmp").replace(f)
    return changed


if __name__ == "__main__":
    n_md = fix_md()
    n_pdf = sum(fix_pdf(str(f)) for f in Path("labs").rglob("*.pdf"))
    print(f"md rewritten: {n_md}, pdf rewritten: {n_pdf}")
