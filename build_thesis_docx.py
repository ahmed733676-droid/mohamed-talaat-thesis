#!/usr/bin/env python3
"""Build an A4 double-spaced thesis .docx that follows the approved PUA protocol layout."""

import re
import shutil
import subprocess
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
PDF = ROOT / "Thesis_Complete.pdf"

CHAPTER_TITLES = {
    "introduction": "INTRODUCTION",
    "review of literature": "REVIEW OF LITERATURE",
    "materials and methods": "MATERIALS AND METHODS",
    "results": "RESULTS",
    "discussion": "DISCUSSION",
    "conclusions and recommendations": "CONCLUSIONS AND RECOMMENDATIONS",
    "references": "REFERENCES",
    "abstract": "ABSTRACT",
    "acknowledgements": "ACKNOWLEDGEMENTS",
    "role of supervisors": "ROLE OF SUPERVISORS",
    "supervisors": "ROLE OF SUPERVISORS",
    "contents": "CONTENTS",
    "list of tables": "LIST OF TABLES",
    "list of figures": "LIST OF FIGURES",
    "list of abbreviations": "LIST OF ABBREVIATIONS",
}

TOC_ITEMS = [
    "ROLE OF SUPERVISORS",
    "ACKNOWLEDGEMENTS",
    "LIST OF TABLES",
    "LIST OF FIGURES",
    "LIST OF ABBREVIATIONS",
    "ABSTRACT",
    "INTRODUCTION",
    "REVIEW OF LITERATURE",
    "MATERIALS AND METHODS",
    "RESULTS",
    "DISCUSSION",
    "CONCLUSIONS AND RECOMMENDATIONS",
    "REFERENCES",
]

CITE_RE = re.compile(r"\((\d+(?:\s*,\s*\d+)*)\)")


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
    if line_spacing == 2.0:
        pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
        pf.line_spacing = 2.0
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = line_spacing
    # Protocol body: paragraphs are marked by a first-line indent, not a blank gap.
    if first_line and align == WD_ALIGN_PARAGRAPH.JUSTIFY:
        pf.first_line_indent = Cm(1.27)
    else:
        pf.first_line_indent = Cm(0)
    pf.widow_control = True


def keep_with_next(p):
    p.paragraph_format.keep_with_next = True


def page_break(doc):
    """Insert a page break on its own single-spaced paragraph so it stays on the previous page."""
    p = doc.add_paragraph()
    set_paragraph_format(p, first_line=False, space_before=0, space_after=0, line_spacing=1.0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.add_run().add_break(WD_BREAK.PAGE)
    return p


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


def add_right_tab(paragraph, pos_cm=16.2):
    pPr = paragraph._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(int(pos_cm * 567)))
    tabs.append(tab)
    pPr.append(tabs)


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
        run.add_picture(str(banner), width=Cm(14.8))
    add_bottom_border(hp)

    footer = section.footer
    footer.is_linked_to_previous = False
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


def centered(doc, text, size, bold=False, italic=False, space_before=0, space_after=0,
             line_spacing=1.15):
    p = doc.add_paragraph()
    set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False,
                         space_before=space_before, space_after=space_after,
                         line_spacing=line_spacing)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def add_cover(doc):
    """Cover laid out from the approved research protocol (Feb 2025)."""
    def left(text, size=12, bold=False, italic=False, space_before=0, space_after=0):
        p = doc.add_paragraph()
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                             space_before=space_before, space_after=space_after,
                             line_spacing=1.15)
        run = p.add_run(text)
        set_run_font(run, size=size, bold=bold, italic=italic)
        return p

    def gap(pts=10):
        p = doc.add_paragraph()
        set_paragraph_format(p, first_line=False, space_before=0, space_after=pts, line_spacing=1.0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        return p

    left("Faculty of Dentistry", 12, bold=True)
    left("Department of Restorative Dentistry", 12)
    left("Student Code No. 202203112", 12, bold=True)
    gap(16)

    centered(doc, "A Thesis submitted in partial fulfilment of the", 12, bold=True)
    centered(doc, "requirements for the degree of Master of Science", 12, bold=True)
    centered(doc, "in Conservative Dentistry", 12, bold=True)
    centered(doc, "Academic Year 2024–2025 / 2025–2026", 12, bold=True)
    gap(16)

    centered(doc, "Name of Candidate", 12)
    centered(doc, "Mohamed Talaat Mohamed AbdelMoaty ElAbd", 13, bold=True)
    gap(16)

    left("English Title:", 12, bold=True, italic=True)
    centered(doc, "COMPARATIVE STUDY OF WEAR RESISTANCE", 12, bold=True, space_before=4)
    centered(doc, "AND SURFACE ROUGHNESS OF INJECTABLE VERSUS", 12, bold=True)
    centered(doc, "CONVENTIONAL COMPOSITE RESIN — IN VITRO STUDY", 12, bold=True)
    gap(10)

    left("Arabic Title:", 12, bold=True, italic=True)
    centered(
        doc,
        "دراسة مقارنة للتآكل وخشونة سطح الراتينج المركب القابل للحقن والتقليدي – دراسة في المختبر",
        12, bold=True, space_before=4,
    )
    gap(10)

    left("Keywords: Surface roughness, wear, injectable composite, conventional composite.", 12)
    gap(16)

    centered(doc, "Supervision Committee", 12, bold=True, space_after=6)
    centered(doc, "1. Prof. Wegdan M. Abdel-Fattah", 12, space_after=2)
    last = centered(doc, "2. Asst. Prof. Emad M. El-Sayed  (Main supervisor)", 12, space_after=0)
    last.add_run().add_break(WD_BREAK.PAGE)


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


def set_academic_cell_borders(cell, *, top=None, bottom=None):
    """Three-line academic table: horizontal rules only (no vertical grid)."""
    tcPr = cell._tc.get_or_add_tcPr()
    # replace any existing borders
    for child in list(tcPr):
        if child.tag == qn("w:tcBorders"):
            tcPr.remove(child)
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        spec = top if edge == "top" else bottom if edge == "bottom" else None
        if spec:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(spec))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")
        else:
            el.set(qn("w:val"), "nil")
            el.set(qn("w:sz"), "0")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def split_header_lines(text):
    if " Mean ± SD" in text:
        return [text.replace(" Mean ± SD", "").strip(), "Mean ± SD"]
    return [text]


def write_cell_text(cell, text, *, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, size=10):
    cell.text = ""
    lines = [ln for ln in text.split("\n") if ln != ""] or [""]
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.first_line_indent = Cm(0)
        p.alignment = align
        run = p.add_run(line)
        set_run_font(run, size=size, bold=bold)


def add_table(doc, rows):
    if not rows:
        return
    cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    usable = 15.8
    if cols == 6:
        widths = [3.5, 2.5, 2.9, 2.7, 1.5, 2.7]
    elif cols == 5:
        widths = [4.0, 3.1, 3.1, 3.1, 2.5]
    elif cols == 3:
        widths = [6.2, 5.4, 4.2]
    else:
        widths = [usable / cols] * cols
    for i, w in enumerate(widths):
        for cell in table.columns[i].cells:
            cell.width = Cm(w)

    header = rows[0]
    last = len(rows) - 1
    numeric_cols = set()
    for c_idx, h in enumerate(header):
        key = h.lower()
        if any(tok in key for tok in ("mean", "p-value", "load", "range", "vol%", "mg", "µm", "μm", "%")):
            numeric_cols.add(c_idx)

    for r_idx, row in enumerate(rows):
        for c_idx in range(cols):
            cell = table.cell(r_idx, c_idx)
            raw = row[c_idx] if c_idx < len(row) else ""
            if r_idx == 0:
                raw = "\n".join(split_header_lines(raw))
            align = WD_ALIGN_PARAGRAPH.CENTER if c_idx in numeric_cols else WD_ALIGN_PARAGRAPH.LEFT
            write_cell_text(cell, raw, bold=(r_idx == 0), align=align, size=10)
            top = 18 if r_idx == 0 else None
            bottom = 8 if r_idx == 0 else (18 if r_idx == last else None)
            set_academic_cell_borders(cell, top=top, bottom=bottom)

    # One p-value for the whole comparison: merge the data cells in that column.
    if header and header[-1].lower().startswith("p-value") and len(rows) > 2:
        table.cell(1, cols - 1).merge(table.cell(last, cols - 1))
        write_cell_text(
            table.cell(1, cols - 1),
            rows[1][cols - 1] or next((r[cols - 1] for r in rows[1:] if r[cols - 1].strip()), ""),
            align=WD_ALIGN_PARAGRAPH.CENTER,
            size=10,
        )
        set_academic_cell_borders(table.cell(1, cols - 1), bottom=18)

    spacer = doc.add_paragraph()
    set_paragraph_format(spacer, first_line=False, space_after=8, line_spacing=1.0)


def add_heading_styled(doc, text, level):
    """Protocol headings: chapter titles centred 18 pt caps; subsections left 14 pt bold."""
    p = doc.add_paragraph()
    if level in (0, 1):
        set_paragraph_format(
            p,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            first_line=False,
            space_before=12 if level == 0 else 0,
            space_after=18,
            line_spacing=1.15,
        )
        run = p.add_run(text)
        set_run_font(run, size=18, bold=True)
    else:
        set_paragraph_format(
            p,
            align=WD_ALIGN_PARAGRAPH.LEFT,
            first_line=False,
            space_before=12,
            space_after=0,
        )
        run = p.add_run(text)
        set_run_font(run, size=14, bold=True)
    keep_with_next(p)
    return p


def add_contents(doc, page_map):
    add_heading_styled(doc, "CONTENTS", 0)
    header = doc.add_paragraph()
    set_paragraph_format(header, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                         space_before=6, space_after=0, line_spacing=1.15)
    add_right_tab(header)
    run = header.add_run("Contents")
    set_run_font(run, bold=True, size=12)
    run = header.add_run("\tPage")
    set_run_font(run, bold=True, size=12)
    add_bottom_border(header)

    for title in TOC_ITEMS:
        p = doc.add_paragraph()
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                             space_before=0, space_after=0)
        add_right_tab(p)
        run = p.add_run(f"– {title}")
        set_run_font(run, size=12)
        page = page_map.get(title, "")
        run = p.add_run(f"\t{page}")
        set_run_font(run, size=12)


def add_image(doc, path: Path):
    if not path.exists():
        return
    p = doc.add_paragraph()
    set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False,
                         space_before=6, space_after=6, line_spacing=1.0)
    run = p.add_run()
    run.add_picture(str(path), width=Inches(5.7))


def render_inline(p, text, size=12, italic=False):
    """Render markdown bold and Vancouver citations as superscript parentheses."""
    parts = text.split("**")
    for idx, part in enumerate(parts):
        if not part:
            continue
        bold = idx % 2 == 1
        pos = 0
        for match in CITE_RE.finditer(part):
            before = part[pos:match.start()]
            if before:
                run = p.add_run(before)
                set_run_font(run, size=size, bold=bold, italic=italic)
            cite = p.add_run(match.group(0))
            set_run_font(cite, size=size, bold=False, italic=False)
            cite.font.superscript = True
            pos = match.end()
        rest = part[pos:]
        if rest:
            run = p.add_run(rest)
            set_run_font(run, size=size, bold=bold, italic=italic)


def hanging_indent(p, left_cm=1.27, hang_cm=1.27):
    p.paragraph_format.left_indent = Cm(left_cm)
    p.paragraph_format.first_line_indent = Cm(-hang_cm)


def add_body_paragraph(doc, text, *, numbered=False, caption=False, footnote=False,
                       formula=False, reference=False):
    p = doc.add_paragraph()
    if caption:
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                             space_before=10, space_after=6, line_spacing=1.15)
        keep_with_next(p)
        render_inline(p, text, italic=False)
        return p
    if footnote:
        set_paragraph_format(p, first_line=False, space_before=0, space_after=0,
                             line_spacing=1.15)
        render_inline(p, text, size=10, italic=True)
        return p
    if formula:
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False,
                             space_before=6, space_after=6)
        render_inline(p, text)
        return p
    set_paragraph_format(p, first_line=not numbered and not reference)
    if numbered or reference:
        hanging_indent(p)
    render_inline(p, text)
    return p


def is_numbered(line):
    return len(line) > 2 and line[0].isdigit() and (
        line[1] == "." or (line[1].isdigit() and line[2] == ".")
    )


def is_formula(line):
    stripped = line.strip()
    return (
        stripped.startswith("[")
        or stripped.startswith("ΔRa")
        or stripped.startswith("[(W")
        or "× 100" in stripped
    )


def is_footnote(line):
    return line.startswith("†") or line.startswith("‡") or line.startswith("Values are mean")


def convert_md(doc, text, page_map):
    lines = text.splitlines()
    start = next(
        (i for i, ln in enumerate(lines) if ln.strip() in {"## Role of Supervisors", "## Supervisors"}),
        0,
    )
    chapter1 = next(i for i, ln in enumerate(lines) if ln.strip() == "# Chapter 1")
    n = len(lines)

    i = start
    first_front = True
    while i < chapter1:
        line = lines[i].rstrip()
        if not line or line.startswith("---"):
            i += 1
            continue
        if line.startswith("## "):
            if not first_front:
                page_break(doc)
            first_front = False
            heading = line[3:].strip()
            display = CHAPTER_TITLES.get(heading.lower(), heading.upper())
            if display == "CONTENTS":
                add_contents(doc, page_map)
                i += 1
                continue
            add_heading_styled(doc, display, 0)
            i += 1
            continue
        if line.startswith("- "):
            p = doc.add_paragraph()
            set_paragraph_format(p, first_line=False)
            run = p.add_run(f"– {line[2:]}")
            set_run_font(run)
            i += 1
            continue
        add_body_paragraph(
            doc,
            line,
            numbered=is_numbered(line),
            footnote=is_footnote(line),
        )
        i += 1

    i = chapter1
    pending_chapter = None
    in_references = False
    while i < n:
        line = lines[i].rstrip()
        if not line or line.startswith("---"):
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
            if title.lower().startswith("chapter "):
                pending_chapter = title
            else:
                page_break(doc)
                if pending_chapter:
                    add_heading_styled(doc, pending_chapter.upper(), 0)
                    pending_chapter = None
                display = CHAPTER_TITLES.get(title.lower(), title.upper())
                add_heading_styled(doc, display, 1)
                in_references = display == "REFERENCES"
            i += 1
            continue
        if line.startswith("### "):
            add_heading_styled(doc, line[4:].strip(), 2)
            i += 1
            continue
        if line.startswith("## "):
            add_heading_styled(doc, line[3:].strip(), 2)
            i += 1
            continue
        caption = line.startswith("**Table ") or line.startswith("**Figure ")
        add_body_paragraph(
            doc,
            line,
            numbered=is_numbered(line) and not caption and not in_references,
            caption=caption,
            footnote=is_footnote(line),
            formula=is_formula(line),
            reference=in_references and is_numbered(line),
        )
        i += 1


def build_docx(page_map):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(3.0)
    section.bottom_margin = Cm(3.3)
    add_header_and_footer(section)

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

    add_cover(doc)
    convert_md(doc, MD.read_text(encoding="utf-8"), page_map)
    doc.save(OUT)
    print(f"Wrote {OUT}")


def convert_pdf():
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise SystemExit("LibreOffice (soffice) is required to build the PDF")
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", str(OUT), "--outdir", str(ROOT)],
        check=True,
    )
    print(f"Wrote {PDF}")


def extract_heading_pages():
    import pymupdf

    found = {}
    doc = pymupdf.open(PDF)
    wanted = set(TOC_ITEMS)
    for i, page in enumerate(doc):
        for raw in page.get_text().splitlines():
            key = raw.strip()
            if key in wanted and key not in found:
                found[key] = i + 1
    print("TOC pages:", found)
    return found


def main():
    build_docx({})
    convert_pdf()
    page_map = extract_heading_pages()
    if page_map:
        build_docx(page_map)
        convert_pdf()


if __name__ == "__main__":
    main()
