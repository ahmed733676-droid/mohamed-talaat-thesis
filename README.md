# Mohamed Talaat — Master's thesis

**Canonical files for supervisors live on `main` only.** Do not open closed pull requests or restored draft names.

| | |
|---|---|
| Candidate | Mohamed Talaat Mohamed AbdelMoaty ElAbd |
| Student code | 202203112 |
| Degree | M.Sc. Conservative Dentistry |
| Faculty | Faculty of Dentistry, Pharos University in Alexandria |
| Title | Comparative study of wear resistance and surface roughness of injectable versus conventional composite resin (in vitro study) |
| Supervisors | Prof. Wegdan M. Abdel-Fattah; Asst. Prof. Emad M. El-Sayed (main supervisor) |

## Files to send the supervisors

GitHub cannot preview Word or PDF. Use the raw download links, not the blob page.

| File | Role |
|---|---|
| [Thesis_Complete.docx](https://github.com/ahmed733676-droid/mohamed-talaat-thesis/raw/main/Thesis_Complete.docx) | **Send this.** Assembled, editable Word thesis. Pages 1–2 are the approved protocol covers. |
| [Thesis_Complete.pdf](https://github.com/ahmed733676-droid/mohamed-talaat-thesis/raw/main/Thesis_Complete.pdf) | Same document as PDF (A4). |
| [Thesis_Complete.md](Thesis_Complete.md) | Chapter source text only. Not the formatted submission. |

Supporting files (not for the supervisors unless they ask):

- `build_thesis_docx.py` — rebuilds Word from the markdown
- `assets/cover/` — protocol PDF, cover page images, official PUA website header logo (`pua_logo_official.png` from pua.edu.eg)
- `WRITING_RULES.md` — writing constraints for later edits
- `STATUS_AND_HANDOFF.md` — internal status

## Do not use or restore

These overlapping drafts were deleted from `main` in [PR #4](https://github.com/ahmed733676-droid/mohamed-talaat-thesis/pull/4). They contradict the locked **10 mm × 1 mm disc** geometry (some still describe bar-shaped specimens):

- `Thesis_Draft_Complete.md`
- `Thesis_Draft_Chapters_1-3.md`
- `Chapter_4_Results.md`
- `Chapter_4_Results_Current.md`
- `Tables_Mastered_Academic.md`

Closed PRs #2 and #3 are superseded. #1 is already merged.

## Locked items (do not reopen)

- Geometry: 10 mm diameter × 1 mm discs, not 15 × 4 × 1.5 mm bars
- Results already in `Thesis_Complete.*`: ANOVA p = 0.329 on percentage weight loss; Kruskal–Wallis p < 0.001 on Ra after brushing; Tables 4.1–4.3; the 11.4 % / 20.7 mg outlier is kept
- Do not invent new numbers, figures, or references
- Do not email `mohammed.talaat09@gmail.com` until Ahmed (`ahmed733676@gmail.com`) approves

## Rebuild Word / PDF after a markdown edit

```
python3 build_thesis_docx.py
soffice --headless --norestore --nolockcheck --convert-to pdf --outdir . Thesis_Complete.docx
```
