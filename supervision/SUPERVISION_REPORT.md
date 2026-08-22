# Supervision Report — Review 1

**Thesis:** Comparative Evaluation of Wear Resistance and Surface Roughness of Injectable
Versus Conventional Nanohybrid Composite Resins (In Vitro Study)
**Candidate:** Mohamed Talaat Mohamed AbdelMoaty ElAbd
**Draft reviewed:** `Thesis_Complete.md` / `Thesis_Complete.docx` at commit `cf6fa7b`
(branch `cursor/complete-thesis-f431`)
**Gate result:** 10 FAIL · 4 WARN · 13 PASS

## Verdict

**NOT RELEASABLE.** Do not submit this draft, print it, or send it to the candidate or the
supervisors as a finished thesis.

The draft is **21 rendered A4 pages against a 40-page minimum** — 48% short. Every chapter is
below its minimum volume, the reference list holds 22 of the required 40 entries, and nine
passages of internal working commentary are still sitting in the thesis body where the examiner
will read them.

This is a review of an honest and largely correct draft that is roughly half-built. The problem
is volume and finish, not competence.

---

## 1. What is already right

Recording this so none of it gets lost in a rewrite:

- **Formatting is exactly to specification.** A4, Times New Roman 12 pt, double spacing
  throughout, 66 of 67 body paragraphs justified, 3 cm left margin. `build_thesis_docx.py`
  produces this reliably and should be kept.
- **The statistics are arithmetically sound.** I recomputed every table against the locked data.
  Absolute weight loss reconciles with the difference of the reported means in all three groups
  (2.60 vs 2.58 mg, 1.10 vs 1.07 mg, 1.00 vs 0.95 mg — within the ±0.2 mg the rounding allows).
  ΔRa equals Ra-after minus Ra-before **exactly** in all three groups. The outlier reconciles
  too: 20.7 mg at 11.4% implies a 181.6 mg disc, which sits credibly above the 160.2 mg group
  mean. Nothing here looks invented.
- **Citation hygiene is clean.** All 22 references are cited, every citation resolves, numbering
  runs 1…22 without gaps, first-appearance order is correct, and there are no duplicates. The
  references that I could recognise are real, well-chosen, and appropriate — Bollen 1997,
  Quirynen & Bollen 1995, Condon & Ferracane 1997, Heintze 2010, Mair 1996, Jones 2004. This is
  a genuine literature base, not padding.
- **The scientific argument is correct and appropriately hedged.** The draft says a
  non-significant ANOVA is a failure to reject rather than proof of equivalence; it refuses to
  read Rajabi's two-body chewing data as equivalent to toothbrush abrasion; it notes that
  Beautifil Flow Plus F00 is a different product from Beautifil Flow Plus X F00; it flags that
  the nanohybrid's higher baseline Ra may partly explain its higher post-brushing Ra. That is
  careful work and it should survive the expansion intact.
- **The outlier was retained and footnoted rather than quietly dropped.** Correct decision.
- **Chapter 5 already contains all six elements** required by `WRITING_RULES.md` §3.

---

## 2. Blocking defects

### B1 — Length: 21 pages against a 40-page minimum

Measured by rendering `Thesis_Complete.docx` through LibreOffice and counting the PDF, not
estimated. 5,075 words total.

| Chapter | Words now | ~Pages now | Minimum words | Deficit |
|---------|----------:|-----------:|--------------:|--------:|
| 1 Introduction | 411 | 1.8 | 1,100 | +689 |
| 2 Review of literature | 1,025 | 4.5 | 3,400 | +2,375 |
| 3 Materials and methods | 589 | 2.6 | 1,500 | +911 |
| 4 Results | 755 | 3.3 | 1,000 | +245 |
| 5 Discussion | 929 | 4.0 | 2,200 | +1,271 |
| 6 Conclusions | 250 | 1.1 | 400 | +150 |

Chapter 2 is the main failure. A Master's review of literature carrying 1,025 words is a summary,
not a review. §4 below sets out where the missing content legitimately comes from.

### B2 — References: 22 of 40

Eighteen short. The existing 22 are good and none should be removed. §5 lists the topic gaps
that real additional literature should fill.

### B3 — Nine passages of internal commentary in the thesis body

This is the most damaging defect after length, because it is visible to the examiner on a first
read and it advertises that the document was assembled rather than written:

| Line | Text | Action |
|------|------|--------|
| 22 | "The photograph was not present in the repository at the time of assembly." | Delete. Insert the real cover. |
| 106 | "volume fraction is not given in the files used here" | Delete. Get the figure from the GC technical datasheet. |
| 149 | "The bar geometry (15 mm × 4 mm × 1.5 mm) that appears in an earlier draft was not the test that was run." | Delete the whole sentence — see B4. |
| 167 | "Batch numbers were not available in the working files and are not invented here." | Delete. Insert the real batch numbers. |
| 173 | "Those details belong in the laboratory notebook and should be inserted from it." | Delete. Insert the details. |
| 196 | "The tests match the locked analysis of the experimental spreadsheet. Software version … should be copied from that spreadsheet's output…" | Delete. Name the software and version. |
| 294 | "The statistical tests match the distributions used in the locked spreadsheet." | Delete — a strength is a property of the study, not of file agreement. |
| 378–380 | Entire "Note on sources used for this assembly" section, with markdown links to the draft files | Delete the section. |

A thesis reports the study. It never discusses which file a number came from, which draft was
superseded, or what the writer could not find.

### B4 — The abandoned bar geometry still appears

Line 149 names the 15 × 4 × 1.5 mm bar in order to explain that it was not used. Delete the
sentence entirely. Chapter 3 states the method that was performed — 10 mm diameter × 1 mm thick
discs in a CAD/CAM Teflon mould — and says nothing about earlier drafts.

Related: `Chapter_4_Results_Current.md` on `main` still opens "Thirty-six **bar-shaped**
specimens". That legacy file contradicts the locked geometry and will poison any future assembly
that reads from it. Archive the superseded drafts under `archive/` so nothing is assembled from
them again (see §6).

### B5 — Three placeholder sites

- Lines 28–31: supervisor names are dotted fills — `Prof. Dr. …………` — plus "*(Names to be entered
  from the approved protocol.)*". Real names and titles must go in, on the cover and on the
  supervisors page.
- Line 22: the cover photograph is still a written instruction to insert a photograph.

### B6 — No figures anywhere

The `.docx` contains zero images. A thesis reporting two quantitative outcome measures across
three groups needs graphical presentation, and Chapter 3 needs apparatus documentation.
`THESIS_ACCEPTANCE_SPEC.md` §4 lists the required set. Figures 4.1–4.3 can be plotted directly
from `locked_data.json` with no new experiment; Figures 3.1–3.3 need photographs from the
candidate.

Figures also carry real page weight: six figures with captions add roughly 3–4 pages.

### B7 — Missing front matter

Absent: **acknowledgements**, **table of contents**, **list of figures**. Each is expected in a
bound Master's thesis and each occupies at least a full page, so this is about 3 pages of
legitimate length as well as a completeness failure.

### B8 — Table 4.3 is presented before Table 4.2

Chapter 4 mentions tables in the order 4.1 → 4.3 → 4.2. Tables must be numbered in order of first
mention. Either move the absolute-weight-loss table after the roughness table, or renumber it —
renumbering is cleaner, since absolute loss belongs with the other weight data in §4.1.

### B9 — Register: the prose is telegraphic, not scholarly

`WRITING_RULES.md` requires natural, precise academic English. The draft avoids AI filler
completely — good — but overcorrects into clipped notes:

| Chapter | Mean sentence length | Sentences ≤ 6 words |
|---------|---------------------:|--------------------:|
| Introduction | 13.6 words | 13% |
| Review of literature | 12.2 words | 18% |
| Materials and methods | 12.1 words | 17% |
| Results | 12.9 words | 23% |
| Discussion | 10.6 words | **27%** |

Target is a mean of 18–25 words with under 10% very short. Examples from the current text:
"Mass falls. Ra rises." "The test is not the same test." "That is the result." "Corrosion and
fatigue run beside both."

A quarter of the Discussion arriving in fragments under seven words does not read as a Master's
candidate reasoning about findings. Join the fragments into connected sentences with explicit
logical links — *because*, *whereas*, *which suggests that*. Note that this rewrite also closes a
meaningful part of the length gap without adding a single new claim, because the connective
tissue is currently missing rather than compressed.

### B10 — ISO 4049 is cited for profilometry

Reference 16 (ISO 4049, *Dentistry — Polymer-based restorative materials*) is attached to the
claim that contact stylus profilometry is a standard method for Ra. ISO 4049 does not specify
surface texture measurement. The correct standards are **ISO 4287** and **ISO 25178** for surface
texture parameters and the definition of Ra. Replace it. ISO 4049 may still be cited legitimately
in Chapter 3 for specimen preparation and curing if that is what was followed.

---

## 3. Items requiring verification against the raw data

Not failures, but they must be confirmed before release.

1. **Beautifil Flow Plus X F00 weight SD drops from 7.4 mg to 3.1 mg after brushing.** This is
   arithmetically consistent with the heaviest disc shedding 20.7 mg and landing on the group
   mean, and my reconstruction supports it. Confirm against the spreadsheet anyway — a more than
   twofold fall in SD is the kind of number an examiner will question, so Chapter 4 should state
   plainly *why* it happens.
2. **Beautifil II LS ΔRa SD (0.066 µm) exceeds both its before-SD (0.054) and after-SD (0.050).**
   That requires a negative within-disc correlation between baseline and final roughness — the
   roughest discs smoothing and the smoothest roughening. It is possible, but confirm that ΔRa
   was computed **per disc** and then averaged, rather than derived from the group means.
3. **Percentage weight loss is a mean of per-disc ratios, not a ratio of the group means.** For
   Beautifil Flow Plus X F00 those differ (1.49% vs 1.61%) because of the outlier. State in
   Chapter 3 §3.7 which was used, so the examiner does not think the table disagrees with itself.
4. **Post-hoc test not named.** Chapter 4 reports pairwise comparisons after Kruskal–Wallis
   without naming the procedure (Dunn? Mann–Whitney with Bonferroni?) or the correction. Name it.
5. **Statistical software and version not stated.** Required in any Methods chapter.
6. **Normality and homogeneity testing not reported.** ANOVA was used for weight loss and
   Kruskal–Wallis for Ra, which implies the distributions were assessed. Report the tests
   (Shapiro–Wilk, Levene) and their outcomes — this justifies the choice of tests, and examiners
   ask about it.

---

## 4. Page budget to clear 40 pages

Target 45–48 pages, not 40, so that examiner-requested cuts cannot drop the thesis below the
minimum.

| Section | Now | Target | Where the content comes from |
|---------|----:|-------:|------------------------------|
| Front matter | 1.9 | 7–8 | Acknowledgements, table of contents, list of figures (all missing); real cover and supervisors |
| 1 Introduction | 1.8 | 5 | Statement of the problem, clinical background on toothbrush abrasion, rationale for choosing these three products, justification of the 10 000-cycle model |
| 2 Review of literature | 4.5 | 15 | §5 below — this is the main lift |
| 3 Materials and methods | 2.6 | 7 | Real laboratory detail (§3 items), full Table 3.1, Figures 3.1–3.3, G*Power inputs, randomisation method |
| 4 Results | 3.3 | 6 | Figures 4.1–4.3, per-group descriptive detail, named post-hoc results, normality outcomes |
| 5 Discussion | 4.0 | 10 | Deeper mechanistic interpretation, comparison against the enlarged literature base, expanded clinical implications |
| 6 Conclusions | 1.1 | 2 | Conclusion per objective, recommendations separated into practice / laboratory / clinical |
| 7 References (40+) | 2.5 | 5 | §5 below |
| **Total** | **21** | **~57** | |

That budget overshoots 45 comfortably, which is the point — it leaves room to cut.

**This must be real content.** Gate G13 in the specification exists precisely to stop the length
gap being closed by restating existing points in new words. Restating the 0.2 µm threshold four
more times fails supervision as surely as being short does.

---

## 5. Reference expansion: 22 → 40+

Keep all 22. Add at least 18 more. Do **not** fabricate a single entry, and do not add a
reference that is not discussed in the text — an uncited reference fails gate G4, and a cited but
undiscussed one fails G13.

Topic gaps in the current review where real literature exists and should be located, retrieved,
read, and verified (DOI or PubMed ID) before insertion:

**Wear mechanisms and testing (Chapter 2.1)**
- Toothbrush abrasion simulator methodology and its validation
- Relative dentine abrasivity of dentifrices and its effect on composite wear
- Two-body versus three-body wear correlation with clinical wear data
- Degree of conversion and its relationship to wear resistance
- Filler silanation and filler–matrix interface failure

**Surface roughness (Chapter 2.2)**
- Bacterial adhesion and biofilm formation on composite surfaces at defined Ra values
- Gloss retention as an outcome measure alongside Ra
- Finishing and polishing protocols and their effect on baseline Ra
- Contact versus non-contact profilometry agreement
- Staining and discolouration as a consequence of surface roughening

**Injectable and highly filled flowable composites (Chapter 2.3)**
- Mechanical characterisation of G-ænial Universal Injectable (flexural strength, hardness,
  elastic modulus)
- Beautifil Flow Plus / Flow Plus X characterisation studies
- The injection-moulding restorative technique and its clinical indications
- Polymerisation shrinkage and shrinkage stress of highly filled flowables
- Clinical trials or case series of injectable composites as definitive restorations

**Giomers and S-PRG technology (Chapter 2.4)**
- S-PRG filler ion release and its mechanism
- Fluoride release and recharge in giomer restoratives
- Surface degradation of giomers in acidic challenge
- Clinical performance of giomer restoratives

**Standards**
- ISO 4287 and/or ISO 25178 (surface texture — replaces the misapplied ISO 4049 at B10)

**Methodological**
- Sample size determination for in-vitro composite studies
- Handling of outliers in small-sample laboratory data

Verification requirement: for each new entry record the DOI or PubMed ID at the moment it is
added. `supervise_thesis.py` can confirm formatting, sequence, and cross-citation, but it cannot
confirm that a paper exists. That check is manual and it is not optional.

---

## 6. Repository hygiene

Five overlapping draft files now describe the same thesis with contradictory content:

| File | Problem |
|------|---------|
| `Thesis_Draft_Chapters_1-3.md` | Bar geometry; `[Equipment list as previously]` placeholders |
| `Thesis_Draft_Complete.md` | Bar geometry; empty results tables; "Insert the detailed text here" |
| `Chapter_4_Results.md` | Superseded |
| `Chapter_4_Results_Current.md` | "Thirty-six **bar-shaped** specimens" — contradicts locked geometry |
| `Tables_Mastered_Academic.md` | Correct, and the source of the locked tables |

Recommendation: move the four superseded files to `archive/` with a one-line note, and keep
`Thesis_Complete.md` as the single working document. This prevents a later assembly from reading
the bar geometry back in — which is exactly how B4 happened.

---

## 7. Checks a script cannot make

`supervise_thesis.py` enforces length, structure, citation mechanics, placeholder and commentary
detection, data fidelity, table arithmetic, and formatting. Three things still require a human:

1. **Reference reality (G4).** Every entry verified against DOI or PubMed, and confirmed to
   actually support the sentence it is attached to. B10 is an example the script cannot catch: the
   citation was well-formed, sequenced correctly, and pointing at the wrong standard.
2. **Padding (G13).** Only a reader can tell new content from restated content.
3. **Register (§6 of the specification).** The sentence-length diagnostic flags telegraphic prose,
   but only a reader can confirm the argument connects.

---

## 8. Required actions, in order

1. Delete all nine internal-commentary passages (B3) and the "Note on sources" section.
2. Delete the bar-geometry sentence (B4); archive the superseded drafts (§6).
3. Insert the real laboratory details from the notebook (§3 items 4–6, and B3 lines 167/173/196).
4. Obtain the supervisor names and the protocol cover from the candidate (B5).
5. Rewrite the existing prose out of telegraphic register into connected academic English (B9).
   Do this **before** expanding — expanding fragmented prose just produces more fragments.
6. Expand Chapter 2 to ≥ 3,400 words against newly verified literature (B1, §5).
7. Expand Chapters 1, 3, 5, 6 to their minima (B1).
8. Build Figures 4.1–4.3 from `locked_data.json`; obtain Figures 3.1–3.3 (B6).
9. Add acknowledgements, table of contents, list of figures (B7).
10. Fix table order (B8) and the ISO 4049 citation (B10).
11. Grow the reference list to 40+, each entry verified (B2, §5).
12. Re-render and re-run the gate. Confirm ≥ 45 pages.

---

## 9. Re-review procedure

```bash
python3 build_thesis_docx.py
soffice --headless --convert-to pdf Thesis_Complete.docx --outdir .
python3 supervision/supervise_thesis.py \
    --md Thesis_Complete.md \
    --docx Thesis_Complete.docx \
    --pdf Thesis_Complete.pdf
```

Exit code 0 means the automated gates pass. It does **not** mean the thesis is releasable —
§7 and the sign-off checklist in `THESIS_ACCEPTANCE_SPEC.md` §8 must also be complete.

Run this before every push. Do not send anything to the candidate or to
`mohammed.talaat09@gmail.com` until Ahmed gives explicit approval
(`WRITING_RULES.md` §7).

---

*Review 1 — supervision branch `cursor/thesis-supervision-qa-e72f`. Re-review will be appended
to this file when the writing branch reports the next draft.*
