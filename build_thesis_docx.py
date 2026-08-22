#!/usr/bin/env python3
"""Build A4 double-spaced thesis .docx from Thesis_Complete.md with a protocol-style cover."""

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
MD = ROOT / "Thesis_Complete.md"
OUT = ROOT / "Thesis_Complete.docx"


def set_run_font(run, size=12, bold=False, italic=False, name="Times New Roman"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)


def set_paragraph_format(p, *, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=0, first_line=True, space_before=0):
    pf = p.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    pf.line_spacing = 2.0
    if first_line and align == WD_ALIGN_PARAGRAPH.JUSTIFY:
        pf.first_line_indent = Cm(1.25)
    else:
        pf.first_line_indent = Cm(0)


def add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    # actual page break
    p = doc.add_paragraph()
    p.add_run().add_break(docx_break())


def docx_break():
    from docx.enum.text import WD_BREAK
    return WD_BREAK.PAGE


def shade_cell(cell, hex_color="F2F2F2"):
    tc = cell._tePr if hasattr(cell, "_tePr") else cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def add_cover(doc):
    def centered(text, size, bold=False, italic=False, space_before=0, space_after=0):
        p = doc.add_paragraph()
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_before=space_before, space_after=space_after)
        run = p.add_run(text)
        set_run_font(run, size=size, bold=bold, italic=italic)
        return p

    centered("PHAROS UNIVERSITY IN ALEXANDRIA", 16, bold=True, space_after=0)
    centered("FACULTY OF DENTISTRY", 14, bold=True)
    centered("Department of Conservative Dentistry", 13, italic=True, space_after=18)

    centered("RESEARCH PROTOCOL / THESIS COVER", 11, italic=True, space_after=18)

    centered(
        "COMPARATIVE EVALUATION OF WEAR RESISTANCE AND SURFACE ROUGHNESS OF INJECTABLE VERSUS CONVENTIONAL NANOHYBRID COMPOSITE RESINS",
        14,
        bold=True,
        space_before=12,
        space_after=6,
    )
    centered("(In Vitro Study)", 13, italic=True, space_after=24)

    centered("A Thesis submitted in partial fulfillment of the", 12)
    centered("requirements for the degree of Master of Science", 12)
    centered("in", 12)
    centered("Conservative Dentistry", 13, bold=True, space_after=24)

    centered("Submitted by", 12, space_before=12)
    centered("Mohamed Talaat Mohamed AbdelMoaty ElAbd", 13, bold=True, space_after=24)

    centered("Supervisors", 12, bold=True, space_before=12)
    centered("Prof. Dr. …………………………………………", 12)
    centered("Ass. Prof. Dr. …………………………………………", 12, space_after=24)

    centered("2025 / 2026", 13, bold=True, space_before=18)

    p = doc.add_paragraph()
    set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_before=18)
    run = p.add_run(
        "This page is typeset from the protocol cover (title, candidate, faculty, degree, year). "
        "Replace it with the photograph of the approved protocol cover when that file is supplied."
    )
    set_run_font(run, size=10, italic=True)

    doc.add_paragraph().add_run().add_break(docx_break())


def parse_table(lines, start):
    """Return (rows, next_index) for a markdown table starting at start."""
    rows = []
    i = start
    while i < len(lines) and lines[i].startswith("|"):
        raw = [c.strip() for c in lines[i].strip().strip("|").split("|")]
        if all(set(c) <= set("-: ") and c for c in raw):
            i += 1
            continue
        rows.append(raw)
        i += 1
    return rows, i


def add_table(doc, rows):
    if not rows:
        return
    cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for r_idx, row in enumerate(rows):
        for c_idx in range(cols):
            cell = table.cell(r_idx, c_idx)
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.first_line_indent = Cm(0)
            text = row[c_idx] if c_idx < len(row) else ""
            run = p.add_run(text)
            set_run_font(run, size=10, bold=(r_idx == 0))
            set_cell_border(cell)
            if r_idx == 0:
                shade_cell(cell, "E8E8E8")
    spacer = doc.add_paragraph()
    set_paragraph_format(spacer, first_line=False, space_after=6)


def add_heading_styled(doc, text, level):
    p = doc.add_paragraph()
    sizes = {0: 16, 1: 14, 2: 13, 3: 12}
    set_paragraph_format(
        p,
        align=WD_ALIGN_PARAGRAPH.LEFT if level else WD_ALIGN_PARAGRAPH.CENTER,
        first_line=False,
        space_before=12 if level else 0,
        space_after=6,
    )
    run = p.add_run(text)
    set_run_font(run, size=sizes.get(level, 12), bold=True)


def is_body_start(line):
    return line == "# Chapter 1"


def convert_md(doc, text):
    lines = text.splitlines()
    # skip markdown cover block until Chapter 1; front matter after cover is rebuilt in Word
    try:
        start = next(i for i, ln in enumerate(lines) if ln.strip() == "# Chapter 1")
    except StopIteration:
        start = 0

    # Abstract and lists from markdown (between Abstract heading and Chapter 1)
    abs_idx = next((i for i, ln in enumerate(lines) if ln.strip() == "## Abstract"), None)
    if abs_idx is not None:
        add_heading_styled(doc, "Abstract", 0)
        i = abs_idx + 1
        while i < start:
            line = lines[i].rstrip()
            if not line:
                i += 1
                continue
            if line.startswith("## List of tables"):
                add_heading_styled(doc, "List of tables", 1)
                i += 1
                continue
            if line.startswith("## List of abbreviations"):
                add_heading_styled(doc, "List of abbreviations", 1)
                i += 1
                continue
            if line.startswith("- "):
                p = doc.add_paragraph()
                set_paragraph_format(p, first_line=False, space_after=0)
                run = p.add_run(line[2:])
                set_run_font(run)
                i += 1
                continue
            if line.startswith("**") and line.endswith("**") is False:
                # bold lead-in paragraph
                p = doc.add_paragraph()
                set_paragraph_format(p)
                render_inline(p, line)
                i += 1
                continue
            if line.startswith("---"):
                i += 1
                continue
            p = doc.add_paragraph()
            set_paragraph_format(p)
            render_inline(p, line)
            i += 1
        doc.add_paragraph().add_run().add_break(docx_break())

    i = start
    n = len(lines)
    while i < n:
        line = lines[i].rstrip()
        if line.startswith("## Note on sources"):
            break
        if not line:
            i += 1
            continue
        if line.startswith("|"):
            rows, i = parse_table(lines, i)
            add_table(doc, rows)
            continue
        if line.startswith("# ") and not line.startswith("##"):
            add_heading_styled(doc, line[2:].strip(), 0 if "Chapter" in line or line[2:].strip() in {
                "Introduction", "Review of literature", "Materials and methods",
                "Results", "Discussion", "Conclusions and recommendations", "References",
            } else 1)
            i += 1
            continue
        if line.startswith("### "):
            add_heading_styled(doc, line[4:].strip(), 2)
            i += 1
            continue
        if line.startswith("## "):
            add_heading_styled(doc, line[3:].strip(), 1)
            i += 1
            continue
        if line.startswith("---"):
            i += 1
            continue
        if line[0].isdigit() and ". " in line[:4] and line[0:2].replace(".", "").isdigit() or (
            len(line) > 2 and line[0].isdigit() and line[1] == "."
        ):
            # numbered conclusions / objectives / references
            p = doc.add_paragraph()
            set_paragraph_format(p, first_line=False)
            render_inline(p, line)
            i += 1
            continue
        p = doc.add_paragraph()
        set_paragraph_format(p)
        render_inline(p, line)
        i += 1


def render_inline(p, text):
    """Minimal **bold** support."""
    parts = text.split("**")
    for idx, part in enumerate(parts):
        if not part:
            continue
        run = p.add_run(part)
        set_run_font(run, bold=(idx % 2 == 1))


def main():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

    add_cover(doc)
    convert_md(doc, MD.read_text(encoding="utf-8"))
    doc.save(OUT)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
