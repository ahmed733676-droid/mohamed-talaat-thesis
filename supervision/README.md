# Supervision

Quality gate for the Mohamed Talaat master's thesis. The writing branch decides when a draft is
finished; this directory decides whether it is acceptable.

## Files

| File | Purpose |
|------|---------|
| `THESIS_ACCEPTANCE_SPEC.md` | Binding acceptance criteria — the definition of "no shortcuts" |
| `supervise_thesis.py` | Automated gate; exits non-zero while any hard gate fails |
| `locked_data.json` | Single source of truth for every number allowed in the thesis |
| `SUPERVISION_REPORT.md` | Review of the current draft, with required corrections |
| `gate_output_review1.txt` | Recorded gate output for the reviewed draft |
| `make_result_figures.py` | Builds Figures 4.1–4.3 from `locked_data.json` |

## Building the results figures

```bash
pip install matplotlib
python3 supervision/make_result_figures.py --outdir figures
```

Writes 300 dpi TIFF and PNG in greyscale, plus `figures/CAPTIONS.md`. Every value is read from
`locked_data.json`, so a figure cannot drift away from its table through re-typing. Regenerate
rather than editing by hand.

## Running the gate

```bash
pip install python-docx pypdf
sudo apt-get install -y --no-install-recommends libreoffice-writer   # authoritative page count

python3 build_thesis_docx.py
soffice --headless --convert-to pdf Thesis_Complete.docx --outdir .
python3 supervision/supervise_thesis.py \
    --md Thesis_Complete.md \
    --docx Thesis_Complete.docx \
    --pdf Thesis_Complete.pdf
```

Without LibreOffice the page count falls back to a word-based estimate at 230 words per A4 page
and is reported as a warning. The estimate is adequate for tracking progress but the count
submitted must come from a rendered PDF.

## What it checks

Length (rendered page count and per-chapter word minima) · required sections · reference count,
Vancouver shape, sequential numbering, first-appearance ordering, cross-citation, duplicates ·
placeholders · internal process commentary · AI filler phrases · sentence-length register ·
locked statistics present and unaltered · table arithmetic reconciliation · specimen geometry
consistency · table and figure numbering and captions · A4 / Times New Roman 12 pt / double
spacing / justification / margins.

## What it cannot check

Whether a reference exists, whether it supports the sentence it is attached to, whether added
length is real content or padding, and whether the prose reads as scholarly English. Those are
manual and they are not optional — see `SUPERVISION_REPORT.md` §7.

Exit code 0 means the automated gates pass, not that the thesis is releasable. Release requires
the sign-off checklist in `THESIS_ACCEPTANCE_SPEC.md` §8.
