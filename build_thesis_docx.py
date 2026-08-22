#!/usr/bin/env python3
"""Build a properly styled A4 thesis .docx that follows the approved PUA protocol."""

import re
import shutil
import subprocess
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
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

TABLE_LIST = [
    ("Table 3.1 Composition of the resin composites", "Table 3.1 Brand"),
    ("Table 4.1 Weight before and after brushing and percentage weight loss", "Table 4.1 Mean weight"),
    ("Table 4.2 Absolute weight loss after 10 000 brushing cycles", "Table 4.2 Absolute"),
    ("Table 4.3 Ra before and after brushing and ΔRa", "Table 4.3 Mean surface"),
]

FIGURE_LIST = [
    ("Figure 4.1 Mean percentage weight loss after 10 000 brushing cycles", "Figure 4.1 Mean percentage"),
    ("Figure 4.2 Mean Ra before and after brushing (0.2 µm threshold)", "Figure 4.2 Mean surface"),
    ("Figure 4.3 Mean change in surface roughness (ΔRa)", "Figure 4.3 Mean change"),
]

COMPACT_FRONT = {
    "ROLE OF SUPERVISORS",
    "LIST OF TABLES",
    "LIST OF FIGURES",
    "LIST OF ABBREVIATIONS",
    "CONTENTS",
}

# Protocol look, with enough body room that Word is not letterboxed.
MARGIN_LEFT_CM = 3.0
MARGIN_RIGHT_CM = 2.2
MARGIN_TOP_CM = 2.80
MARGIN_BOTTOM_CM = 2.55
HEADER_DISTANCE_CM = 0.30
FOOTER_DISTANCE_CM = 0.28
HEADER_IMAGE_WIDTH_CM = 13.8
BODY_WIDTH_CM = 15.8
BODY_FIRST_LINE_CM = 1.27

CITE_RE = re.compile(r"\((\d+(?:\s*,\s*\d+)*)\)")


def has_arabic(text):
    return any("\u0600" <= ch <= "\u06FF" for ch in text or "")


def set_run_font(run, size=12, bold=False, italic=False, name="Times New Roman"):
    text = run.text or ""
    cs_name = "Noto Naskh Arabic" if has_arabic(text) else name
    run.font.name = name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:cs"), cs_name)
    rFonts.set(qn("w:eastAsia"), name)
    half = str(int(round(size * 2)))
    run.font.size = Pt(size)
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), half)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)


def set_paragraph_format(p, *, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=0,
                         first_line=False, space_before=0, line_spacing=1.15):
    pf = p.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if abs(line_spacing - 2.0) < 0.01:
        pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
        pf.line_spacing = 2.0
    elif abs(line_spacing - 1.0) < 0.01:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.line_spacing = 1.0
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = line_spacing
    pf.first_line_indent = Cm(BODY_FIRST_LINE_CM) if first_line else Cm(0)
    pf.widow_control = True


def keep_with_next(p):
    p.paragraph_format.keep_with_next = True


def page_break_before(p):
    pPr = p._p.get_or_add_pPr()
    el = OxmlElement("w:pageBreakBefore")
    pPr.append(el)


def add_bottom_border(paragraph, sz="12"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_top_border(paragraph, sz="12"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), sz)
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), "000000")
    pBdr.append(top)
    pPr.append(pBdr)


def add_right_tab(paragraph, pos_cm=15.8, leader="dot"):
    pPr = paragraph._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:leader"), leader)
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


def configure_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    rPr = normal._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), "Times New Roman")
    rFonts.set(qn("w:hAnsi"), "Times New Roman")
    rFonts.set(qn("w:cs"), "Times New Roman")
    rFonts.set(qn("w:eastAsia"), "Times New Roman")
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), "24")
    nf = normal.paragraph_format
    nf.space_before = Pt(0)
    nf.space_after = Pt(0)
    nf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    nf.line_spacing = 2.0
    nf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    nf.first_line_indent = Cm(0)
    nf.widow_control = True

    for style_name, size, align in (("Heading 1", 16, WD_ALIGN_PARAGRAPH.CENTER),
                                   ("Heading 2", 12, WD_ALIGN_PARAGRAPH.LEFT)):
        st = doc.styles[style_name]
        st.font.name = "Times New Roman"
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.font.italic = False
        st.font.underline = False
        pf = st.paragraph_format
        pf.alignment = align
        pf.space_before = Pt(0 if style_name == "Heading 1" else 16)
        pf.space_after = Pt(14 if style_name == "Heading 1" else 8)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.line_spacing = 1.0
        pf.first_line_indent = Cm(0)
        pf.keep_with_next = True


def add_header_and_footer(section):
    section.different_first_page_header_footer = False
    section.odd_and_even_pages_header_footer = False
    section.header_distance = Cm(HEADER_DISTANCE_CM)
    section.footer_distance = Cm(FOOTER_DISTANCE_CM)

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.space_before = Pt(0)
    hp.paragraph_format.space_after = Pt(0)
    hp.paragraph_format.line_spacing = 1.0
    hp.paragraph_format.first_line_indent = Cm(0)
    banner = ROOT / "figures" / "pua_header.png"
    run = hp.add_run()
    if banner.exists():
        # Width only: the banner already carries its own rule, so a second
        # paragraph border is not added (that doubled line was crowding the body).
        run.add_picture(str(banner), width=Cm(HEADER_IMAGE_WIDTH_CM))

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(0)
    fp.paragraph_format.space_after = Pt(0)
    fp.paragraph_format.line_spacing = 1.0
    fp.paragraph_format.first_line_indent = Cm(0)
    add_top_border(fp)
    compact = [
        "Address: P.O. Box 37, Sidi Gaber, Canal El Mahmoudia Street, Smouha, Alexandria, Egypt",
        "العنوان: صندوق بريد ٣٧ سيدي جابر – شارع قناة المحمودية – سموحة – الإسكندرية – مصر",
        "Phone: +(203) 38 77 026    Fax: +(203) 383 0249    E-mail: Dentistry@pua.edu.eg    www.pua.edu.eg",
    ]
    run = fp.add_run(compact[0])
    set_run_font(run, size=8)
    for line in compact[1:]:
        p = footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.first_line_indent = Cm(0)
        r = p.add_run(line)
        set_run_font(r, size=8)
    pnum = footer.add_paragraph()
    pnum.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pnum.paragraph_format.space_before = Pt(3)
    pnum.paragraph_format.space_after = Pt(0)
    pnum.paragraph_format.line_spacing = 1.0
    pnum.paragraph_format.first_line_indent = Cm(0)
    add_page_field(pnum, size=12)


def cover_para(doc, text, *, size=12, bold=False, italic=False,
               align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0):
    p = doc.add_paragraph()
    set_paragraph_format(p, align=align, first_line=False,
                         space_before=space_before, space_after=space_after,
                         line_spacing=1.15)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def add_cover(doc):
    cover_para(doc, "Pharos University in Alexandria", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    cover_para(doc, "Faculty of Dentistry", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    cover_para(doc, "Department of Restorative Dentistry", align=WD_ALIGN_PARAGRAPH.LEFT)
    cover_para(doc, "Student Code No. 202203112", bold=True,
               align=WD_ALIGN_PARAGRAPH.LEFT, space_after=14)

    cover_para(doc, "A Thesis submitted in partial fulfilment of the", bold=True, space_before=8)
    cover_para(doc, "requirements for the degree of Master of Science", bold=True)
    cover_para(doc, "in Conservative Dentistry", bold=True)
    cover_para(doc, "Academic Year 2024–2025 / 2025–2026", bold=True, space_after=14)

    cover_para(doc, "Name of Candidate", space_before=6)
    cover_para(doc, "Mohamed Talaat Mohamed AbdelMoaty ElAbd", size=14, bold=True, space_after=14)

    cover_para(doc, "English Title:", bold=True, italic=True,
               align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4)
    cover_para(doc, "COMPARATIVE STUDY OF WEAR RESISTANCE", bold=True, space_before=4)
    cover_para(doc, "AND SURFACE ROUGHNESS OF INJECTABLE VERSUS", bold=True)
    cover_para(doc, "CONVENTIONAL COMPOSITE RESIN — IN VITRO STUDY", bold=True, space_after=12)

    cover_para(doc, "Arabic Title:", bold=True, italic=True,
               align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4)
    cover_para(
        doc,
        "دراسة مقارنة للتآكل وخشونة سطح الراتينج المركب القابل للحقن والتقليدي – دراسة في المختبر",
        bold=True, space_before=4, space_after=12,
    )

    p = doc.add_paragraph()
    set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                         space_after=14, line_spacing=1.15)
    run = p.add_run("Keywords: ")
    set_run_font(run, size=12, bold=True)
    run = p.add_run("Surface roughness, wear, injectable composite, conventional composite.")
    set_run_font(run, size=12)

    cover_para(doc, "Supervision Committee", bold=True, space_before=6, space_after=8)
    cover_para(doc, "1. Prof. Wegdan M. Abdel-Fattah", space_after=4)
    cover_para(doc, "2. Asst. Prof. Emad M. El-Sayed  (Main supervisor)")


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
    tcPr = cell._tc.get_or_add_tcPr()
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


def set_cell_margins(cell, twips=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for edge in ("top", "left", "bottom", "right"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), str(twips))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def split_header_lines(text):
    if " Mean ± SD" in text:
        return [text.replace(" Mean ± SD", "").strip(), "Mean ± SD"]
    return [text]


def write_cell_text(cell, text, *, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, size=10.5):
    cell.text = ""
    lines = [ln for ln in text.split("\n") if ln != ""] or [""]
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        set_paragraph_format(p, align=align, first_line=False, line_spacing=1.15)
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(line)
        set_run_font(run, size=size, bold=bold)


def add_table(doc, rows):
    if not rows:
        return
    cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    usable = BODY_WIDTH_CM
    if cols == 6:
        # Wide enough that Dimethacrylate / Multifunctional do not split mid-word.
        widths = [2.9, 2.3, 3.4, 3.5, 1.5, 2.2]
    elif cols == 5:
        widths = [4.2, 3.0, 3.0, 3.0, 2.6]
    elif cols == 3:
        widths = [6.2, 5.4, 4.2]
    else:
        widths = [usable / cols] * cols
    cell_size = 10
    cell_pad = 80 if cols == 6 else 100

    tbl = table._tbl
    tblPr = tbl.tblPr
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(int(sum(widths) * 567)))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)

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
        tr = table.rows[r_idx]._tr
        trPr = tr.get_or_add_trPr()
        cant = OxmlElement("w:cantSplit")
        trPr.append(cant)
        if r_idx == 0:
            hdr = OxmlElement("w:tblHeader")
            trPr.append(hdr)
        for c_idx in range(cols):
            cell = table.cell(r_idx, c_idx)
            raw = row[c_idx] if c_idx < len(row) else ""
            if r_idx == 0:
                raw = "\n".join(split_header_lines(raw))
            align = WD_ALIGN_PARAGRAPH.CENTER if c_idx in numeric_cols else WD_ALIGN_PARAGRAPH.LEFT
            write_cell_text(cell, raw, bold=(r_idx == 0), align=align, size=cell_size)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell, twips=cell_pad)
            top = 18 if r_idx == 0 else None
            bottom = 8 if r_idx == 0 else (18 if r_idx == last else None)
            set_academic_cell_borders(cell, top=top, bottom=bottom)

    if header and header[-1].lower().startswith("p-value") and len(rows) > 2:
        pval = next((r[cols - 1] for r in rows[1:] if r[cols - 1].strip()), "")
        table.cell(1, cols - 1).merge(table.cell(last, cols - 1))
        write_cell_text(table.cell(1, cols - 1), pval, align=WD_ALIGN_PARAGRAPH.CENTER, size=cell_size)
        set_academic_cell_borders(table.cell(1, cols - 1), bottom=18)
        set_cell_margins(table.cell(1, cols - 1), twips=cell_pad)

    spacer = doc.add_paragraph()
    set_paragraph_format(spacer, first_line=False, space_after=12, line_spacing=1.0)


def add_heading_styled(doc, text, level, *, new_page=False, style_name=None):
    style = style_name or ("Heading 1" if level in (0, 1) else "Heading 2")
    p = doc.add_paragraph()
    p.style = style
    if new_page:
        page_break_before(p)
    if level in (0, 1):
        after = 6 if text.startswith("CHAPTER ") else 16
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False,
                             space_before=0, space_after=after, line_spacing=1.0)
        run = p.add_run(text)
        set_run_font(run, size=16, bold=True)
    else:
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                             space_before=16, space_after=8, line_spacing=1.0)
        run = p.add_run(text)
        set_run_font(run, size=12, bold=True)
    keep_with_next(p)
    return p


def add_contents(doc, page_map):
    add_heading_styled(doc, "CONTENTS", 0)
    header = doc.add_paragraph()
    set_paragraph_format(header, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                         space_before=10, space_after=8, line_spacing=1.15)
    add_right_tab(header, leader="none")
    run = header.add_run("Contents")
    set_run_font(run, bold=True, size=12)
    run = header.add_run("\tPage")
    set_run_font(run, bold=True, size=12)
    add_bottom_border(header, sz="8")

    for title in TOC_ITEMS:
        p = doc.add_paragraph()
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                             space_before=4, space_after=4, line_spacing=1.15)
        add_right_tab(p, leader="dot")
        run = p.add_run(f"{title}")
        set_run_font(run, size=12)
        page = page_map.get(title, "")
        run = p.add_run(f"\t{page}")
        set_run_font(run, size=12)


def add_leader_list(doc, items, page_map):
    header = doc.add_paragraph()
    set_paragraph_format(header, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                         space_before=10, space_after=8, line_spacing=1.15)
    add_right_tab(header, leader="none")
    run = header.add_run("Title")
    set_run_font(run, bold=True, size=12)
    run = header.add_run("\tPage")
    set_run_font(run, bold=True, size=12)
    add_bottom_border(header, sz="8")
    for title, _needle in items:
        p = doc.add_paragraph()
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                             space_before=6, space_after=6, line_spacing=1.15)
        add_right_tab(p, leader="dot")
        run = p.add_run(title)
        set_run_font(run, size=12)
        page = page_map.get(title, "")
        run = p.add_run(f"\t{page}")
        set_run_font(run, size=12)


def add_image(doc, path: Path):
    if not path.exists():
        return
    p = doc.add_paragraph()
    set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False,
                         space_before=10, space_after=8, line_spacing=1.0)
    run = p.add_run()
    run.add_picture(str(path), width=Inches(5.4))


def render_inline(p, text, size=12, italic=False):
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


def hanging_indent(p, left_cm=1.27, hang_cm=0.63):
    p.paragraph_format.left_indent = Cm(left_cm)
    p.paragraph_format.first_line_indent = Cm(-hang_cm)


def add_body_paragraph(doc, text, *, numbered=False, caption=False, footnote=False,
                       formula=False, reference=False, compact=False, duty=False):
    p = doc.add_paragraph()
    if caption:
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                             space_before=14, space_after=8, line_spacing=1.15)
        keep_with_next(p)
        render_inline(p, text)
        return p
    if footnote:
        set_paragraph_format(p, first_line=False, space_before=4, space_after=10,
                             line_spacing=1.15)
        render_inline(p, text, size=10, italic=True)
        return p
    if formula:
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False,
                             space_before=8, space_after=8, line_spacing=1.15)
        render_inline(p, text)
        return p
    if compact or duty:
        before = 10 if numbered else 0
        after = 6 if numbered else 4
        set_paragraph_format(p, first_line=False, line_spacing=1.15,
                             space_before=before, space_after=after)
        if numbered:
            hanging_indent(p, BODY_FIRST_LINE_CM, 0.63)
        elif duty:
            p.paragraph_format.left_indent = Cm(BODY_FIRST_LINE_CM)
        render_inline(p, text)
        return p
    set_paragraph_format(p, first_line=not numbered and not reference, line_spacing=2.0)
    if numbered or reference:
        hanging_indent(p, 1.27, 0.63)
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
    current_front = ""
    while i < chapter1:
        line = lines[i].rstrip()
        if not line or line.startswith("---"):
            i += 1
            continue
        if line.startswith("## "):
            heading = line[3:].strip()
            display = CHAPTER_TITLES.get(heading.lower(), heading.upper())
            current_front = display
            add_heading_styled(doc, display, 0, new_page=not first_front)
            first_front = False
            if display == "CONTENTS":
                # heading already added; replace by contents block without a second title
                # remove the heading we just made by building contents under it — add_contents
                # would duplicate CONTENTS. Build the list only.
                header = doc.add_paragraph()
                set_paragraph_format(header, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                                     space_before=10, space_after=8, line_spacing=1.15)
                add_right_tab(header, leader="none")
                run = header.add_run("Contents")
                set_run_font(run, bold=True, size=12)
                run = header.add_run("\tPage")
                set_run_font(run, bold=True, size=12)
                add_bottom_border(header, sz="8")
                for title in TOC_ITEMS:
                    p = doc.add_paragraph()
                    set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False,
                                         space_before=4, space_after=4, line_spacing=1.15)
                    add_right_tab(p, leader="dot")
                    run = p.add_run(title)
                    set_run_font(run, size=12)
                    page = page_map.get(title, "")
                    run = p.add_run(f"\t{page}")
                    set_run_font(run, size=12)
            elif display == "LIST OF TABLES":
                add_leader_list(doc, TABLE_LIST, page_map)
            elif display == "LIST OF FIGURES":
                add_leader_list(doc, FIGURE_LIST, page_map)
            i += 1
            continue
        if current_front in {"CONTENTS", "LIST OF TABLES", "LIST OF FIGURES"}:
            i += 1
            continue
        if line.startswith("- "):
            p = doc.add_paragraph()
            set_paragraph_format(p, first_line=False, line_spacing=1.15,
                                 space_before=3, space_after=6)
            hanging_indent(p, BODY_FIRST_LINE_CM, 0.63)
            run = p.add_run(line[2:])
            set_run_font(run)
            i += 1
            continue
        compact = current_front in COMPACT_FRONT
        duty = compact and current_front == "ROLE OF SUPERVISORS" and not is_numbered(line)
        add_body_paragraph(
            doc,
            line,
            numbered=is_numbered(line),
            footnote=is_footnote(line),
            compact=compact,
            duty=duty,
        )
        i += 1

    i = chapter1
    pending_chapter = None
    in_references = False
    first_chapter = True
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
                if pending_chapter:
                    add_heading_styled(doc, pending_chapter.upper(), 0, new_page=True)
                    pending_chapter = None
                else:
                    add_heading_styled(doc, "", 0, new_page=True)
                display = CHAPTER_TITLES.get(title.lower(), title.upper())
                add_heading_styled(doc, display, 1)
                in_references = display == "REFERENCES"
                first_chapter = False
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


def apply_section(section):
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(MARGIN_LEFT_CM)
    section.right_margin = Cm(MARGIN_RIGHT_CM)
    section.top_margin = Cm(MARGIN_TOP_CM)
    section.bottom_margin = Cm(MARGIN_BOTTOM_CM)
    add_header_and_footer(section)


def build_docx(page_map):
    doc = Document()
    configure_styles(doc)
    apply_section(doc.sections[0])
    add_cover(doc)
    body = doc.add_section(WD_SECTION.NEW_PAGE)
    apply_section(body)
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
        text = page.get_text()
        for raw in text.splitlines():
            key = raw.strip()
            if key in wanted and key not in found:
                found[key] = i + 1
        # Caption pages: skip the front-matter list pages themselves.
        if i + 1 <= 8:
            continue
        squashed = " ".join(text.split())
        for title, needle in TABLE_LIST + FIGURE_LIST:
            if title not in found and needle in squashed:
                found[title] = i + 1
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
