#!/usr/bin/env python3
"""Build an A4 double-spaced thesis .docx from Thesis_Complete.md."""

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

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


def set_paragraph_format(p, *, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=0,
                         first_line=True, space_before=0, line_spacing=2.0):
    pf = p.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE if line_spacing == 2.0 else WD_LINE_SPACING.SINGLE
    pf.line_spacing = line_spacing
    pf.first_line_indent = Cm(1.25) if first_line and align == WD_ALIGN_PARAGRAPH.JUSTIFY else Cm(0)


def page_break(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)


def shade_cell(cell, hex_color="E8E8E8"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_border(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def add_bottom_border(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_top_border(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "12")
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), "000000")
    pBdr.append(top)
    pPr.append(pBdr)


def add_page_field(paragraph, size=10):
    run = paragraph.add_run()
    set_run_font(run, size=size)
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    run._r.append(fld1)
    run._r.append(instr)
    run._r.append(fld2)


def add_header_and_footer(section):
    """Protocol header (PUA bilingual banner + rule) and footer (address block + page) on every page."""
    section.different_first_page_header_footer = False
    section.odd_and_even_pages_header_footer = False
    section.header_distance = Cm(0.5)
    section.footer_distance = Cm(0.4)

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.space_before = Pt(0)
    hp.paragraph_format.space_after = Pt(2)
    hp.paragraph_format.line_spacing = 1.0
    banner = ROOT / "figures" / "pua_header.png"
    run = hp.add_run()
    if banner.exists():
        run.add_picture(str(banner), width=Cm(16.0))
    add_bottom_border(hp)

    footer = section.footer
    footer.is_linked_to_previous = False
    # clear default paragraph
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(2)
    fp.paragraph_format.space_after = Pt(0)
    fp.paragraph_format.line_spacing = 1.0
    add_top_border(fp)
    lines = [
        "Address: P.O. Box 37, Sidi Gaber, Canal El Mahmoudia Street, Smouha, Alexandria, Egypt",
        "العنوان: صندوق بريد ٣٧ سيدي جابر – شارع قناة المحمودية – سموحة – الإسكندرية – مصر",
        "Phone: +(203) 38 77 026     Fax: +(203) 383 0249",
        "E-mail: Dentistry@pua.edu.eg     Web Site: www.pua.edu.eg",
    ]
    run = fp.add_run(lines[0])
    set_run_font(run, size=8)
    for line in lines[1:]:
        p = footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(line)
        set_run_font(r, size=8)
    pnum = footer.add_paragraph()
    pnum.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pnum.paragraph_format.space_before = Pt(2)
    pnum.paragraph_format.space_after = Pt(0)
    pnum.paragraph_format.line_spacing = 1.0
    add_page_field(pnum, size=10)


def centered(doc, text, size, bold=False, italic=False, space_before=0, space_after=0):
    p = doc.add_paragraph()
    set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False,
                         space_before=space_before, space_after=space_after)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def add_cover(doc):
    """Cover laid out from the approved research protocol (Feb 2025)."""
    def left(text, size=12, bold=False, italic=False, space_before=0, space_after=0):
        p = doc.add_paragraph()
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                             space_before=space_before, space_after=space_after)
        run = p.add_run(text)
        set_run_font(run, size=size, bold=bold, italic=italic)
        return p

    left("Faculty of Dentistry", 13, bold=True)
    left("Department of Restorative Dentistry", 13)
    left("Student Code No. 202203112", 12, bold=True, space_after=12)

    centered(doc, "A Thesis submitted in partial fulfilment of the", 13, bold=True, space_before=12)
    centered(doc, "requirements for the degree of Master of Science", 13, bold=True)
    centered(doc, "in Conservative Dentistry", 13, bold=True, space_after=6)
    centered(doc, "Academic Year 2024–2025 / 2025–2026", 13, bold=True, space_after=12)

    centered(doc, "Name of Candidate", 12, space_before=6)
    centered(doc, "Mohamed Talaat Mohamed AbdelMoaty ElAbd", 14, bold=True, space_after=12)

    left("English Title:", 12, bold=True, italic=True, space_before=8)
    centered(
        doc,
        "COMPARATIVE STUDY OF WEAR RESISTANCE AND SURFACE ROUGHNESS OF "
        "INJECTABLE VERSUS CONVENTIONAL COMPOSITE RESIN — IN VITRO STUDY",
        13, bold=True, space_before=4, space_after=8,
    )
    left("Arabic Title:", 12, bold=True, italic=True)
    centered(
        doc,
        "دراسة مقارنة للتآكل وخشونة سطح الراتينج المركب القابل للحقن والتقليدي – دراسة في المختبر",
        13, bold=True, space_before=4, space_after=10,
    )
    left("Keywords: Surface roughness, wear, injectable composite, conventional composite.", 12, space_after=12)

    centered(doc, "Supervision Committee", 13, bold=True, space_before=8, space_after=6)
    centered(doc, "1. Prof. Wegdan M. Abdel-Fattah", 12)
    centered(doc, "2. Asst. Prof. Emad M. El-Sayed  (Main supervisor)", 12, space_after=6)
    page_break(doc)


def parse_table(lines, start):
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
            set_run_font(run, size=9, bold=(r_idx == 0))
            set_cell_border(cell)
            if r_idx == 0:
                shade_cell(cell)
    spacer = doc.add_paragraph()
    set_paragraph_format(spacer, first_line=False, space_after=6)


def add_heading_styled(doc, text, level):
    p = doc.add_paragraph()
    sizes = {0: 16, 1: 14, 2: 13, 3: 12}
    set_paragraph_format(
        p,
        align=WD_ALIGN_PARAGRAPH.CENTER if level == 0 else WD_ALIGN_PARAGRAPH.LEFT,
        first_line=False,
        space_before=12 if level else 6,
        space_after=6,
    )
    run = p.add_run(text)
    set_run_font(run, size=sizes.get(level, 12), bold=True)


def add_image(doc, path: Path):
    if not path.exists():
        return
    p = doc.add_paragraph()
    set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_before=6, space_after=6)
    run = p.add_run()
    run.add_picture(str(path), width=Inches(5.7))


def render_inline(p, text):
    parts = text.split("**")
    for idx, part in enumerate(parts):
        if not part:
            continue
        run = p.add_run(part)
        set_run_font(run, bold=(idx % 2 == 1))


def convert_md(doc, text):
    lines = text.splitlines()
    start = next((i for i, ln in enumerate(lines) if ln.strip() == "## Supervisors"), 0)
    chapter1 = next(i for i, ln in enumerate(lines) if ln.strip() == "# Chapter 1")

    # Front matter from Supervisors through Chapter 1
    i = start
    while i < chapter1:
        line = lines[i].rstrip()
        if not line or line.startswith("---"):
            i += 1
            continue
        if line.startswith("## "):
            heading = line[3:].strip()
            add_heading_styled(doc, heading, 0 if heading in {"Abstract"} else 1)
            i += 1
            continue
        if line.startswith("- "):
            p = doc.add_paragraph()
            set_paragraph_format(p, first_line=False, space_after=0)
            run = p.add_run(line[2:])
            set_run_font(run)
            i += 1
            continue
        p = doc.add_paragraph()
        set_paragraph_format(p)
        render_inline(p, line)
        i += 1
    page_break(doc)

    i = chapter1
    n = len(lines)
    while i < n:
        line = lines[i].rstrip()
        if not line:
            i += 1
            continue
        if line.startswith("---"):
            i += 1
            continue
        if line.startswith("![") and "](" in line:
            src = line.split("](", 1)[1].rstrip(")")
            add_image(doc, ROOT / src)
            i += 1
            continue
        if line.startswith("|"):
            rows, i = parse_table(lines, i)
            add_table(doc, rows)
            continue
        if line.startswith("# ") and not line.startswith("##"):
            title = line[2:].strip()
            level = 0 if title.startswith("Chapter") or title in {
                "Introduction", "Review of literature", "Materials and methods",
                "Results", "Discussion", "Conclusions and recommendations", "References",
            } else 1
            add_heading_styled(doc, title, level)
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
        numbered = len(line) > 2 and line[0].isdigit() and (line[1] == "." or (line[1].isdigit() and line[2] == "."))
        p = doc.add_paragraph()
        set_paragraph_format(p, first_line=not numbered)
        render_inline(p, line)
        i += 1


def main():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(3.4)
    section.bottom_margin = Cm(3.6)
    add_header_and_footer(section)

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
