#!/usr/bin/env python3
"""Supervision gate for the Mohamed Talaat master's thesis.

Enforces the machine-checkable subset of supervision/THESIS_ACCEPTANCE_SPEC.md.
Exits non-zero while any hard gate fails.

    python3 supervision/supervise_thesis.py --md Thesis_Complete.md --docx Thesis_Complete.docx

Checks that cannot be automated -- whether a reference really exists, whether added
length is real content or padding, whether the prose reads as scholarly English --
are listed in the report as items requiring a human read.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

ROOT = Path(__file__).resolve().parent
WORDS_PER_PAGE = 230  # A4, Times New Roman 12 pt, double spaced, 3 cm left margin

MIN_PAGES = 40
TARGET_PAGES = 45
MIN_REFERENCES = 40

CHAPTER_MINIMA = {
    "Introduction": 1100,
    "Review of literature": 3400,
    "Materials and methods": 1500,
    "Results": 1000,
    "Discussion": 2200,
    "Conclusions and recommendations": 400,
}

REQUIRED_SECTIONS = [
    ("Title page", r"master of science"),
    ("Supervisors", r"^#{1,3}\s*supervisors"),
    ("Acknowledgements", r"^#{1,3}\s*acknowledge"),
    ("Table of contents", r"^#{1,3}\s*(table of contents|contents)\b"),
    ("List of tables", r"^#{1,3}\s*list of tables"),
    ("List of figures", r"^#{1,3}\s*list of figures"),
    ("List of abbreviations", r"^#{1,3}\s*list of abbreviations"),
    ("Abstract", r"^#{1,3}\s*abstract"),
    ("Introduction", r"^#{1,3}\s*(chapter 1|introduction)"),
    ("Review of literature", r"^#{1,3}\s*(chapter 2|review of literature)"),
    ("Materials and methods", r"^#{1,3}\s*(chapter 3|materials and methods)"),
    ("Results", r"^#{1,3}\s*(chapter 4|results)"),
    ("Discussion", r"^#{1,3}\s*(chapter 5|discussion)"),
    ("Conclusions", r"^#{1,3}\s*(chapter 6|conclusions)"),
    ("References", r"^#{1,3}\s*(chapter 7|references)"),
]

# Chapter 5 must demonstrably contain each of these (WRITING_RULES.md sec. 3).
DISCUSSION_ELEMENTS = {
    "interpretation of findings": r"null hypothes|interpret",
    "comparison with literature": r"\bet al\b|agree|in line with|consistent with|contrast",
    "clinical implications": r"clinical",
    "strengths": r"strength",
    "limitations": r"limitation|\blimits?\b",
    "concluding remarks": r"conclu|close of the discussion|summary of the discussion",
}

PLACEHOLDER_PATTERNS = [
    (r"\.{3,}", "ellipsis / dotted fill"),
    (r"\u2026{2,}", "dotted fill"),
    (r"\u2026\s*\u2026", "dotted fill"),
    (r"\bT\.?B\.?D\b", "TBD"),
    (r"\bto be (completed|written|added|determined|entered|inserted)\b", "to be completed"),
    (r"\bas previously (detailed|polished|provided|described)\b", "cross-reference to a lost draft"),
    (r"\bwaiting for\b", "waiting for data"),
    (r"\[(\.\.\.|insert|year|placeholder|equipment|table content)[^\]]*\]", "bracketed placeholder"),
    (r"\bplaceholder\b", "placeholder"),
    (r"\bwill be (written|expanded|completed|provided|verified)\b", "deferred work"),
    (r"\bshould be (copied|inserted|entered|supplied)\b", "deferred work"),
    (r"\bnot (available|present|recorded) in the (working )?(files|repository)\b", "missing data admission"),
    (r"\bwere not invented\b", "missing data admission"),
    (r"\bbelongs? in the laboratory notebook\b", "missing data admission"),
    (r"\bXX+\b", "XX placeholder"),
]

INTERNAL_COMMENTARY_PATTERNS = [
    (r"\bnext agent\b|\bcontinuing agents?\b|\bagent (should|must|will)\b", "agent-to-agent instruction"),
    (r"\brepositor(y|ies)\b", "repository reference"),
    (r"\bgithub\b", "GitHub reference"),
    (r"\bmarkdown\b|\.md\b", "markdown / file reference"),
    (r"\bhandoff\b", "handoff note"),
    (r"\bearlier draft\b|\ban earlier draft\b", "draft-history commentary"),
    (r"\blocked (spreadsheet|analysis)\b", "internal process reference"),
    (r"\bthe files used (here|for this write-up)\b", "internal process reference"),
    (r"\bat the time of assembly\b", "internal process reference"),
    (r"\bthis (write-up|assembly)\b", "internal process reference"),
    (r"\bnote to user\b", "note to user"),
    (r"^\*\*status:\*\*", "status line"),
    (r"\bpandoc\b|\bconverters?\b", "tooling reference"),
    (r"\bcurrent best version\b", "draft-status commentary"),
    (r"\.xlsx\b|\.docx\b", "source-file reference"),
]

AI_FILLER_PATTERNS = [
    r"\bit is important to note\b",
    r"\bit is worth noting\b",
    r"\bplays? a (crucial|vital|key|pivotal) role\b",
    r"\bdelve into\b",
    r"\bin today's world\b",
    r"\bmoreover, it\b",
    r"\bfurthermore, it is\b",
    r"\ba testament to\b",
    r"\bnavigate the (complex|landscape)\b",
    r"\bunderscore[sd]? the importance\b",
    r"\brich tapestry\b",
    r"\bever-evolving\b",
]

STATUS = {"pass": "PASS", "fail": "FAIL", "warn": "WARN", "info": "INFO"}


@dataclass
class Result:
    gate: str
    name: str
    status: str
    detail: str
    items: list[str] = field(default_factory=list)


class Report:
    def __init__(self) -> None:
        self.results: list[Result] = []

    def add(self, gate: str, name: str, status: str, detail: str, items: list[str] | None = None) -> None:
        self.results.append(Result(gate, name, status, detail, items or []))

    @property
    def failures(self) -> list[Result]:
        return [r for r in self.results if r.status == "fail"]

    @property
    def warnings(self) -> list[Result]:
        return [r for r in self.results if r.status == "warn"]

    def render(self) -> str:
        icon = {"pass": "PASS", "fail": "FAIL", "warn": "WARN", "info": "INFO"}
        out = []
        for r in self.results:
            out.append(f"[{icon[r.status]}] {r.gate:<4} {r.name}")
            if r.detail:
                out.append(f"           {r.detail}")
            for it in r.items[:12]:
                out.append(f"             - {it}")
            if len(r.items) > 12:
                out.append(f"             ... and {len(r.items) - 12} more")
        return "\n".join(out)


# ----------------------------------------------------------------------------- helpers

def strip_code_and_meta(text: str) -> str:
    """Remove fenced blocks so tooling snippets are not audited as thesis prose."""
    return re.sub(r"```.*?```", " ", text, flags=re.S)


def word_count(text: str) -> int:
    return len(re.findall(r"\b[\w\u2019'-]+\b", text))


def decimals(value: float) -> int:
    s = repr(float(value))
    return len(s.split(".")[1].rstrip("0")) if "." in s else 0


def dedupe(hits: list[str]) -> list[str]:
    """One report line per (line number, label): greedy patterns overlap heavily."""
    seen, out = set(), []
    for h in hits:
        key = h.split(":", 1)[0]
        if key not in seen:
            seen.add(key)
            out.append(h)
    return out


def line_of(text: str, pos: int) -> int:
    return text.count("\n", 0, pos) + 1


def split_sections(text: str) -> dict[str, str]:
    """Map a canonical section name to its body text."""
    lines = text.splitlines()
    heads: list[tuple[int, str]] = []
    for i, ln in enumerate(lines):
        m = re.match(r"^(#{1,2})\s+(.*\S)\s*$", ln)
        if m:
            heads.append((i, m.group(2).strip()))
    sections: dict[str, str] = {}
    for idx, (i, title) in enumerate(heads):
        end = heads[idx + 1][0] if idx + 1 < len(heads) else len(lines)
        body = "\n".join(lines[i + 1 : end])
        # "# Chapter 4" followed by "# Results" -> attribute the body to "Results"
        if re.fullmatch(r"Chapter \d+", title):
            continue
        sections[title] = sections.get(title, "") + "\n" + body
    return sections


def find_body_start(text: str) -> int:
    """Offset where the thesis body begins (used to exclude front matter from some checks)."""
    m = re.search(r"^#{1,2}\s*(chapter 1|introduction)\b", text, re.I | re.M)
    return m.start() if m else 0


# ----------------------------------------------------------------------------- gates

def check_pages(rep: Report, md: str, docx: Path | None, pdf: Path | None) -> None:
    pages = None
    source = ""
    if pdf and pdf.exists():
        pages, source = pdf_pages(pdf), f"rendered PDF {pdf.name}"
    elif docx and docx.exists():
        rendered = render_pdf(docx)
        if rendered:
            pages, source = pdf_pages(rendered), f"{docx.name} rendered via LibreOffice"

    words = word_count(strip_code_and_meta(md))
    est = words / WORDS_PER_PAGE

    if pages is None:
        why = ("no --docx or --pdf supplied" if not (docx or pdf)
               else "LibreOffice not installed or the conversion failed")
        rep.add(
            "G1", "Page count", "warn",
            f"Estimated {est:.0f} pages from {words:,} words at {WORDS_PER_PAGE} words/page "
            f"({why}). The count submitted must come from a rendered PDF.",
        )
        pages = est
        source = "estimate"

    status = "pass" if pages >= MIN_PAGES else "fail"
    rep.add(
        "G1", "Page count >= 40", status,
        f"{pages:.0f} pages ({source}); minimum {MIN_PAGES}. "
        f"Shortfall {max(0, MIN_PAGES - pages):.0f} pages.",
    )
    rep.add(
        "G2", f"Page count >= {TARGET_PAGES} (safety margin)",
        "pass" if pages >= TARGET_PAGES else "warn",
        f"{pages:.0f} pages. A thesis delivered at exactly 40 drops below the minimum after "
        f"the first round of examiner cuts.",
    )
    rep.add("G1", "Total word count", "info", f"{words:,} words (excluding fenced code blocks)")


def render_pdf(docx: Path) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    tmp = Path(tempfile.mkdtemp(prefix="thesis-render-"))
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", str(docx), "--outdir", str(tmp)],
            check=True, capture_output=True, timeout=600,
        )
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return None
    out = tmp / (docx.stem + ".pdf")
    return out if out.exists() else None


def pdf_pages(pdf: Path) -> int:
    try:
        from pypdf import PdfReader
        return len(PdfReader(str(pdf)).pages)
    except Exception:
        # /Type /Page counting fallback
        data = pdf.read_bytes()
        return max(1, len(re.findall(rb"/Type\s*/Page[^s]", data)))


def check_chapter_volume(rep: Report, md: str) -> None:
    sections = split_sections(md)
    short = []
    detail_rows = []
    for name, minimum in CHAPTER_MINIMA.items():
        body = next((v for k, v in sections.items() if k.lower() == name.lower()), None)
        if body is None:
            short.append(f"{name}: section not found")
            continue
        w = word_count(body)
        detail_rows.append(f"{name}: {w:,} words (~{w / WORDS_PER_PAGE:.1f} pp), minimum {minimum:,}")
        if w < minimum:
            short.append(f"{name}: {w:,} words, needs {minimum:,} (+{minimum - w:,})")
    rep.add("G2", "Chapter volume", "info", "Measured against THESIS_ACCEPTANCE_SPEC.md sec. 2", detail_rows)
    rep.add(
        "G2", "Chapter minima", "fail" if short else "pass",
        f"{len(short)} chapter(s) below the minimum." if short else "All chapters meet the minimum.",
        short,
    )


def check_structure(rep: Report, md: str) -> None:
    missing = []
    for label, pattern in REQUIRED_SECTIONS:
        if not re.search(pattern, md, re.I | re.M):
            missing.append(label)
    rep.add(
        "G-struct", "Required sections present", "fail" if missing else "pass",
        f"{len(missing)} required section(s) missing." if missing else "All required sections present.",
        missing,
    )

    sections = split_sections(md)
    disc = next((v for k, v in sections.items() if k.lower() == "discussion"), "")
    if disc:
        absent = [label for label, pat in DISCUSSION_ELEMENTS.items() if not re.search(pat, disc, re.I)]
        rep.add(
            "G12", "Chapter 5 required elements", "fail" if absent else "pass",
            "Missing: " + ", ".join(absent) if absent else
            "Interpretation, literature comparison, clinical implications, strengths, limitations "
            "and concluding remarks all present.",
        )
    else:
        rep.add("G12", "Chapter 5 required elements", "fail", "Discussion chapter not found.")


def parse_references(md: str) -> list[tuple[int, str]]:
    m = re.search(r"^#{1,2}\s*(chapter 7\s*)?#*\s*references\s*$", md, re.I | re.M)
    if not m:
        return []
    tail = md[m.end():]
    stop = re.search(r"^#{1,2}\s+(?!references)", tail, re.M)
    if stop:
        tail = tail[: stop.start()]
    refs = []
    for rm in re.finditer(r"^\s*(\d{1,3})\.\s+(\S.*)$", tail, re.M):
        refs.append((int(rm.group(1)), rm.group(2).strip()))
    return refs


def citation_numbers_in_order(md: str) -> list[int]:
    """Citation numbers in order of first appearance, from the body only."""
    body = md[find_body_start(md):]
    body = re.sub(r"^#{1,2}\s*(chapter 7\s*)?#*\s*references\s*$.*", "", body, flags=re.I | re.M | re.S)
    order: list[int] = []
    for m in re.finditer(r"\((\d{1,3}(?:\s*[,\u2013-]\s*\d{1,3})*)\)", body):
        group = m.group(1)
        nums: list[int] = []
        ok = True
        for part in re.split(r",", group):
            part = part.strip()
            if re.fullmatch(r"\d{1,3}", part):
                nums.append(int(part))
            elif re.fullmatch(r"\d{1,3}\s*[\u2013-]\s*\d{1,3}", part):
                a, b = re.split(r"[\u2013-]", part)
                lo, hi = int(a), int(b)
                if hi - lo > 40:
                    ok = False
                    break
                nums.extend(range(lo, hi + 1))
            else:
                ok = False
                break
        if not ok or any(n == 0 or n > 200 for n in nums):
            continue
        for n in nums:
            if n not in order:
                order.append(n)
    return order


def check_references(rep: Report, md: str) -> None:
    refs = parse_references(md)
    numbers = [n for n, _ in refs]
    rep.add(
        "G3", "Reference count >= 40", "pass" if len(refs) >= MIN_REFERENCES else "fail",
        f"{len(refs)} references; minimum {MIN_REFERENCES}. "
        f"Shortfall {max(0, MIN_REFERENCES - len(refs))}.",
    )

    if not refs:
        rep.add("G5", "Reference numbering", "fail", "No reference list parsed.")
        return

    expected = list(range(1, len(refs) + 1))
    rep.add(
        "G5", "Sequential numbering 1..N", "pass" if numbers == expected else "fail",
        "Numbered 1..N without gaps or repeats." if numbers == expected
        else f"Numbering irregular: {numbers[:20]}",
    )

    cited = citation_numbers_in_order(md)
    cited_set = set(cited)
    uncited = sorted(set(numbers) - cited_set)
    dangling = sorted(n for n in cited_set if n not in set(numbers))
    rep.add(
        "G4", "Every reference cited in text", "pass" if not uncited else "fail",
        "Every entry is cited." if not uncited else f"Uncited entries: {uncited}",
    )
    rep.add(
        "G4", "Every citation resolves to an entry", "pass" if not dangling else "fail",
        "All citations resolve." if not dangling else f"Citations with no entry: {dangling}",
    )

    ideal = list(range(1, len(cited) + 1))
    if cited == ideal:
        rep.add("G5", "Citations numbered in order of first appearance", "pass",
                f"First-appearance order is 1..{len(cited)}.")
    else:
        bad = [f"position {i + 1}: found ({c}), expected ({e})"
               for i, (c, e) in enumerate(zip(cited, ideal)) if c != e][:8]
        rep.add("G5", "Citations numbered in order of first appearance", "fail",
                "Vancouver requires reference numbers assigned in order of first mention. "
                "Renumber the list to match the text.", bad)

    # duplicates by normalised title
    seen: dict[str, int] = {}
    dups = []
    for n, entry in refs:
        key = re.sub(r"[^a-z0-9]", "", entry.lower())[:90]
        if key in seen:
            dups.append(f"({seen[key]}) and ({n}) appear to be the same source")
        else:
            seen[key] = n
    rep.add("G4", "No duplicate references", "pass" if not dups else "fail",
            "No duplicates detected." if not dups else f"{len(dups)} possible duplicate(s).", dups)

    # Vancouver shape
    malformed = []
    for n, entry in refs:
        has_year = re.search(r"\b(19|20)\d{2}\b", entry)
        journal_style = re.search(r"\b(19|20)\d{2}(;|\s*;)", entry)
        is_standard = re.search(r"\bISO\b|International Organization for Standardization", entry)
        is_product = re.search(r"product information|\baccessed\b", entry, re.I)
        is_book = re.search(r"\b\d+(st|nd|rd|th) ed\b|\bpress\b|\bpublish", entry, re.I)
        if not has_year:
            malformed.append(f"({n}) no year: {entry[:70]}")
        elif not (journal_style or is_standard or is_product or is_book):
            malformed.append(f"({n}) no 'year;volume(issue):pages' block: {entry[:70]}")
        if not entry.endswith("."):
            malformed.append(f"({n}) does not end with a full stop")
    rep.add(
        "G3", "Vancouver formatting shape", "warn" if malformed else "pass",
        f"{len(malformed)} entr(y/ies) need a format review." if malformed
        else "All entries carry a year and a locator block.", malformed,
    )

    rep.add(
        "G4", "Reference reality (manual)", "warn",
        f"{len(refs)} entries require verification against DOI / PubMed. "
        "A script cannot confirm a source exists. See SUPERVISION_REPORT.md sec. 7.",
    )

    styles = set()
    if re.search(r"\(\d{1,3}(,\d{1,3})*\)", md[find_body_start(md):]):
        styles.add("parenthetical")
    if re.search(r"<sup>|\[\^", md):
        styles.add("superscript")
    if len(styles) > 1:
        rep.add("G3", "Citation style consistency", "fail",
                f"Mixed citation styles in use: {sorted(styles)}")
    else:
        rep.add("G3", "Citation style consistency", "info",
                f"Style in use: {sorted(styles) or ['none detected']}. "
                "STATUS_AND_HANDOFF.md specifies superscript; confirm which form the faculty requires.")


def check_placeholders(rep: Report, md: str) -> None:
    text = strip_code_and_meta(md)
    hits = []
    for pattern, label in PLACEHOLDER_PATTERNS:
        for m in re.finditer(pattern, text, re.I | re.M):
            frag = text[max(0, m.start() - 45): m.end() + 45].replace("\n", " ")
            hits.append(f"line {line_of(text, m.start())} [{label}]: ...{frag.strip()}...")
    hits = dedupe(hits)
    rep.add("G6", "No placeholders", "fail" if hits else "pass",
            f"{len(hits)} placeholder site(s) still in the text." if hits else "No placeholders found.", hits)

    meta = []
    for pattern, label in INTERNAL_COMMENTARY_PATTERNS:
        for m in re.finditer(pattern, text, re.I | re.M):
            frag = text[max(0, m.start() - 55): m.end() + 55].replace("\n", " ")
            meta.append(f"line {line_of(text, m.start())} [{label}]: ...{frag.strip()}...")
    meta = dedupe(meta)
    rep.add("G7", "No internal / process commentary", "fail" if meta else "pass",
            f"{len(meta)} passage(s) address the writing process rather than the science."
            if meta else "No internal commentary found.", meta)

    filler = []
    for pattern in AI_FILLER_PATTERNS:
        for m in re.finditer(pattern, text, re.I):
            filler.append(f"line {line_of(text, m.start())}: {m.group(0)}")
    rep.add("style", "No AI filler phrases", "fail" if filler else "pass",
            f"{len(filler)} filler phrase(s)." if filler else "No stock filler phrases found.", filler)


def check_register(rep: Report, md: str) -> None:
    """Flag telegraphic prose: a thesis needs connected sentences, not notes."""
    sections = split_sections(md)
    rows = []
    offenders = []
    for name in CHAPTER_MINIMA:
        body = next((v for k, v in sections.items() if k.lower() == name.lower()), "")
        if not body:
            continue
        prose = "\n".join(
            ln for ln in body.splitlines()
            if not ln.strip().startswith(("|", "-", "*", "#", ">"))
            and not re.match(r"^\s*\d+\.\s", ln)
        )
        prose = re.sub(r"\*\*.*?\*\*", "", prose)
        sents = [s.strip() for s in re.split(r"(?<=[.!?])\s+", prose) if s.strip()]
        if len(sents) < 10:
            continue
        lengths = [word_count(s) for s in sents]
        mean_len = sum(lengths) / len(lengths)
        very_short = sum(1 for n in lengths if n <= 6)
        pct = 100 * very_short / len(lengths)
        rows.append(f"{name}: mean sentence {mean_len:.1f} words; {pct:.0f}% of sentences <= 6 words")
        if mean_len < 14 or pct > 18:
            offenders.append(
                f"{name}: mean {mean_len:.1f} words/sentence, {pct:.0f}% very short "
                f"(target: mean 18-25, under 10% very short)"
            )
    rep.add("style", "Sentence-length profile", "info", "Register diagnostic", rows)
    rep.add(
        "style", "Academic register (not telegraphic)", "warn" if offenders else "pass",
        "Chapters read as clipped notes rather than connected scholarly argument. "
        "Join the fragments into full sentences with explicit logical links."
        if offenders else "Sentence lengths are within the scholarly range.", offenders,
    )


def check_data_fidelity(rep: Report, md: str, locked: dict) -> None:
    text = strip_code_and_meta(md)
    flat = re.sub(r"\s+", " ", text)

    missing = []
    for domain, key in (("weight", "loss_percent"), ("weight", "loss_mg"),
                        ("roughness", "after_um"), ("roughness", "before_um"),
                        ("roughness", "delta_um")):
        for group, vals in locked[domain]["groups"].items():
            mean, sd = vals[key]
            # A locked value is written with the precision it was measured at, which is the
            # greater of the mean's and the SD's decimal places: 0.100 +/- 0.023, not 0.1 +/- 0.023.
            dec = max(decimals(mean), decimals(sd))
            pat = rf"{re.escape(f'{mean:.{dec}f}')}\s*(?:\u00b1|\+/-)\s*{re.escape(f'{sd:.{dec}f}')}"
            if not re.search(pat, flat):
                missing.append(f"{group} {domain}.{key}: {mean} \u00b1 {sd} not found as a mean \u00b1 SD pair")
    rep.add("G8", "Locked mean \u00b1 SD values present", "fail" if missing else "pass",
            f"{len(missing)} locked value pair(s) absent or altered." if missing
            else "Every locked mean \u00b1 SD pair appears in the text.", missing)

    wrong = []
    for probe, label in (
        (str(locked["weight"]["p_value"]), "ANOVA p-value"),
        (str(locked["design"]["brushing_cycles"]), "brushing cycles"),
        (f"{locked['design']['n_per_group']}", "n per group"),
        (str(locked["thresholds"]["plaque_retention_um"]), "plaque threshold"),
    ):
        spaced = re.sub(r"(\d)(?=(\d{3})+$)", r"\1 ", probe) if probe.isdigit() and len(probe) > 4 else probe
        if probe not in flat and spaced not in flat:
            wrong.append(f"{label} ({probe}) not found")
    rep.add("G8", "Locked design and test values present", "fail" if wrong else "pass",
            "; ".join(wrong) if wrong else "Cycles, load, n, p-values and thresholds all present.")

    # Geometry: the abandoned bar must not survive anywhere.
    bar = re.search(r"15\s*mm\s*[x\u00d7]\s*4\s*mm|15\s*[x\u00d7]\s*4\s*[x\u00d7]\s*1\.5", flat, re.I)
    disc = re.search(r"10\s*mm\s*(in\s*)?(diameter|\u00d7|x)", flat, re.I)
    if bar:
        rep.add("G10", "Specimen geometry consistent", "fail",
                "The abandoned 15 x 4 x 1.5 mm bar geometry still appears. The experiment used "
                "10 mm x 1 mm discs. Remove every trace of the bar, including any sentence that "
                "explains the discrepancy -- a thesis states the method used, it does not "
                "adjudicate between drafts.")
    elif not disc:
        rep.add("G10", "Specimen geometry consistent", "fail",
                "Disc geometry (10 mm diameter x 1 mm thick) is not stated.")
    else:
        rep.add("G10", "Specimen geometry consistent", "pass",
                "Disc geometry stated; no surviving reference to the abandoned bar.")


def check_arithmetic(rep: Report, locked: dict) -> None:
    """Reconcile the locked tables against each other. Examiners recompute these."""
    notes, fails = [], []
    for group, v in locked["weight"]["groups"].items():
        before_mg = v["before_g"][0] * 1000
        after_mg = v["after_g"][0] * 1000
        implied = before_mg - after_mg
        reported = v["loss_mg"][0]
        # Means are rounded to 0.1 mg, so the difference carries +/- 0.2 mg.
        if abs(implied - reported) > 0.25:
            fails.append(f"{group}: (before - after) = {implied:.2f} mg but Table 4.3 reports "
                         f"{reported:.2f} mg")
        else:
            notes.append(f"{group}: (before - after) = {implied:.2f} mg vs reported "
                         f"{reported:.2f} mg -- reconciles")

        ratio_of_means = 100 * reported / before_mg
        mean_of_ratios = v["loss_percent"][0]
        gap = abs(ratio_of_means - mean_of_ratios)
        rel = gap / max(mean_of_ratios, 1e-9)
        if rel > 0.25:
            fails.append(f"{group}: percentage loss {mean_of_ratios}% is inconsistent with "
                         f"{reported} mg on a {before_mg:.1f} mg disc ({ratio_of_means:.2f}%)")
        else:
            notes.append(f"{group}: {mean_of_ratios}% (mean of ratios) vs {ratio_of_means:.2f}% "
                         f"(ratio of means) -- consistent")

    for group, v in locked["roughness"]["groups"].items():
        implied = v["after_um"][0] - v["before_um"][0]
        reported = v["delta_um"][0]
        if abs(implied - reported) > 0.0015:
            fails.append(f"{group}: Ra after - Ra before = {implied:.3f} but delta-Ra is "
                         f"reported as {reported:.3f}")
        else:
            notes.append(f"{group}: delta-Ra {reported:.3f} = {v['after_um'][0]:.3f} - "
                         f"{v['before_um'][0]:.3f} -- exact")

    out = locked["weight"]["outlier"]
    implied_mass = out["loss_mg"] / (out["loss_percent"] / 100)
    notes.append(f"outlier: {out['loss_mg']} mg at {out['loss_percent']}% implies a disc of "
                 f"{implied_mass:.1f} mg, above the group mean of "
                 f"{locked['weight']['groups'][out['group']]['before_g'][0] * 1000:.1f} mg -- plausible")

    rep.add("G9", "Table arithmetic reconciles", "fail" if fails else "pass",
            f"{len(fails)} inconsistenc(y/ies)." if fails
            else "Weight, percentage, absolute loss and delta-Ra all reconcile.", fails or notes)

    # Distributional sanity, reported as a verification item rather than a failure.
    checks = []
    fp = locked["weight"]["groups"]["Beautifil Flow Plus X F00"]
    if fp["after_g"][1] < fp["before_g"][1] / 2:
        checks.append(
            f"Beautifil Flow Plus X F00 weight SD falls from {fp['before_g'][1] * 1000:.1f} mg to "
            f"{fp['after_g'][1] * 1000:.1f} mg. This is consistent with the heaviest disc shedding "
            f"20.7 mg onto the group mean, but confirm it against the raw spreadsheet."
        )
    ls = locked["roughness"]["groups"]["Beautifil II LS"]
    if ls["delta_um"][1] > max(ls["before_um"][1], ls["after_um"][1]):
        checks.append(
            f"Beautifil II LS delta-Ra SD ({ls['delta_um'][1]:.3f}) exceeds both the before "
            f"({ls['before_um'][1]:.3f}) and after ({ls['after_um'][1]:.3f}) SDs. That requires a "
            f"negative within-disc correlation. Confirm delta-Ra was computed per disc and not "
            f"from the group means."
        )
    rep.add("G9", "Distributional sanity", "warn" if checks else "pass",
            "Verify against the raw data." if checks else "No distributional anomalies.", checks)


def check_tables_figures(rep: Report, md: str) -> None:
    body = md[find_body_start(md):]

    for kind, gate in (("Table", "G-tab"), ("Figure", "G-fig")):
        order, seen = [], set()
        for m in re.finditer(rf"\b{kind}\s+(\d+\.\d+)", body):
            ref = m.group(1)
            if ref not in seen:
                seen.add(ref)
                order.append(ref)
        if kind == "Figure" and not order:
            rep.add("G-fig", "Figures present", "fail",
                    "The thesis contains no figures. Two quantitative outcome measures need "
                    "graphical presentation, and Chapter 3 needs specimen and apparatus images. "
                    "See THESIS_ACCEPTANCE_SPEC.md sec. 4 for the required set.")
            continue

        def key(r: str) -> tuple[int, int]:
            a, b = r.split(".")
            return int(a), int(b)

        ordered = sorted(order, key=key)
        rep.add(
            f"{gate}", f"{kind}s numbered in order of first mention",
            "pass" if order == ordered else "fail",
            f"First-mention order: {order}. Expected {ordered}." if order != ordered
            else f"{len(order)} {kind.lower()}(s) mentioned in ascending order.",
        )

        captioned = set(re.findall(rf"^\*\*{kind}\s+(\d+\.\d+)\*\*", body, re.M))
        uncaptioned = [r for r in order if r not in captioned]
        rep.add(
            f"{gate}", f"Every {kind.lower()} has a caption",
            "pass" if not uncaptioned else "fail",
            f"No caption found for: {uncaptioned}" if uncaptioned
            else f"All {kind.lower()}s captioned.",
        )


def check_docx(rep: Report, docx: Path | None) -> None:
    if not docx or not docx.exists():
        rep.add("G11", "Formatting", "warn", "No .docx supplied; formatting not verified.")
        return
    try:
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except ImportError:
        rep.add("G11", "Formatting", "warn", "python-docx not installed (pip install python-docx).")
        return

    d = Document(str(docx))
    s = d.sections[0]
    problems = []

    def close(a: float, b: float, tol: float = 0.06) -> bool:
        return abs(a - b) <= tol

    if not (close(s.page_width.cm, 21.0) and close(s.page_height.cm, 29.7)):
        problems.append(f"page is {s.page_width.cm:.1f} x {s.page_height.cm:.1f} cm, not A4")
    if not close(s.left_margin.cm, 3.0, 0.2):
        problems.append(f"left margin {s.left_margin.cm:.1f} cm, specified 3 cm")

    normal = d.styles["Normal"]
    if normal.font.name != "Times New Roman":
        problems.append(f"Normal font is {normal.font.name!r}, not Times New Roman")
    if normal.font.size and abs(normal.font.size.pt - 12) > 0.01:
        problems.append(f"Normal size is {normal.font.size.pt} pt, not 12 pt")

    bad_spacing = bad_font = 0
    body_paras = justified = 0
    for p in d.paragraphs:
        if not p.text.strip():
            continue
        ls = p.paragraph_format.line_spacing
        if ls is not None and abs(float(ls) - 2.0) > 0.01:
            bad_spacing += 1
        if len(p.text.split()) > 25:
            body_paras += 1
            if p.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                justified += 1
        for r in p.runs:
            if r.text.strip() and r.font.name not in (None, "Times New Roman"):
                bad_font += 1
    if bad_spacing:
        problems.append(f"{bad_spacing} paragraph(s) not double spaced")
    if bad_font:
        problems.append(f"{bad_font} run(s) not Times New Roman")
    if body_paras and justified / body_paras < 0.9:
        problems.append(f"only {justified}/{body_paras} long paragraphs are justified")

    rep.add("G11", "Formatting (A4 / TNR 12 / double / justified / 3 cm)",
            "fail" if problems else "pass",
            "; ".join(problems) if problems
            else f"A4, Times New Roman 12 pt, double spaced, {justified}/{body_paras} body "
                 f"paragraphs justified, {s.left_margin.cm:.0f} cm left margin.")

    if len(d.inline_shapes) == 0:
        rep.add("G-fig", "Images embedded in .docx", "fail",
                "The .docx contains no images. Figures 3.1-3.3 and 4.1-4.3 are required.")


# ----------------------------------------------------------------------------- main

def main() -> int:
    ap = argparse.ArgumentParser(description="Supervision gate for the thesis.")
    ap.add_argument("--md", required=True, type=Path, help="assembled thesis markdown")
    ap.add_argument("--docx", type=Path, help="assembled thesis .docx")
    ap.add_argument("--pdf", type=Path, help="rendered PDF (authoritative page count)")
    ap.add_argument("--locked", type=Path, default=ROOT / "locked_data.json")
    ap.add_argument("--json", type=Path, help="write machine-readable results here")
    args = ap.parse_args()

    if not args.md.exists():
        print(f"error: {args.md} not found", file=sys.stderr)
        return 2

    md = args.md.read_text(encoding="utf-8")
    locked = json.loads(args.locked.read_text(encoding="utf-8"))

    rep = Report()
    check_pages(rep, md, args.docx, args.pdf)
    check_chapter_volume(rep, md)
    check_structure(rep, md)
    check_references(rep, md)
    check_placeholders(rep, md)
    check_register(rep, md)
    check_data_fidelity(rep, md, locked)
    check_arithmetic(rep, locked)
    check_tables_figures(rep, md)
    check_docx(rep, args.docx)

    print("=" * 78)
    print(f"THESIS SUPERVISION GATE  --  {args.md.name}")
    print("=" * 78)
    print(rep.render())
    print("=" * 78)
    n_fail, n_warn = len(rep.failures), len(rep.warnings)
    print(f"{n_fail} FAIL   {n_warn} WARN   "
          f"{sum(1 for r in rep.results if r.status == 'pass')} PASS")
    print("VERDICT: " + ("NOT RELEASABLE" if n_fail else
                         "gates pass; manual verification still required"))
    print("=" * 78)

    if args.json:
        args.json.write_text(json.dumps(
            [r.__dict__ for r in rep.results], indent=2, ensure_ascii=False), encoding="utf-8")

    return 1 if n_fail else 0


if __name__ == "__main__":
    sys.exit(main())
