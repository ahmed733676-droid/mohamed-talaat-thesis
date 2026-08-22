# Thesis Acceptance Specification

**Thesis:** Comparative Evaluation of Wear Resistance and Surface Roughness of Injectable
Versus Conventional Nanohybrid Composite Resins (In Vitro Study)
**Candidate:** Mohamed Talaat Mohamed AbdelMoaty ElAbd
**Role of this document:** binding acceptance criteria. A draft is *not* releasable until every
mandatory gate below passes. `supervise_thesis.py` enforces the machine-checkable subset.

This specification exists because "finished" and "acceptable" are different things. The writing
branch decides when a draft is finished. This document decides when it is acceptable.

---

## 1. Hard gates (release blockers)

A draft that fails any of these must not be submitted, printed, emailed, or shown to the
supervisors as final.

| # | Gate | Criterion | Why it is a blocker |
|---|------|-----------|---------------------|
| G1 | Length | Rendered PDF ≥ **40 pages** including references, measured on the final A4 layout | Explicit requirement from the candidate |
| G2 | Length margin | Target **45–48 pages** so that examiner-requested cuts cannot drop the thesis below 40 | A thesis delivered at exactly 40 fails after the first round of trimming |
| G3 | References | ≥ **40** references, Vancouver style, sequentially numbered | Explicit requirement |
| G4 | Reference integrity | Every reference real and verifiable; no fabricated, duplicated, or uncited entries | Fabrication is academic misconduct |
| G5 | Citation order | Reference numbers assigned in order of first appearance in the text, 1…N with no gaps | Vancouver rule; examiners check this |
| G6 | No placeholders | Zero `TBD`, `…………`, `[insert]`, `to be completed`, `as previously`, `waiting for` | A placeholder in a submitted thesis is an automatic revision |
| G7 | No internal commentary | Zero references to the repository, drafts, agents, handoff notes, spreadsheets, "the working files", or "was not invented here" | This text is invisible to the writer and fatal to the reader |
| G8 | Data fidelity | Every statistic matches `locked_data.json` exactly; no invented numbers | Non-negotiable per `WRITING_RULES.md` |
| G9 | Arithmetic coherence | Weight loss %, absolute loss (mg), and ΔRa must reconcile with the reported means | Examiners recompute tables |
| G10 | Geometry consistency | Specimen geometry is **10 mm diameter × 1 mm thick discs** everywhere; no surviving reference to the abandoned 15 × 4 × 1.5 mm bar | A contradiction here invalidates Chapter 3 |
| G11 | Formatting | A4; Times New Roman; body 12 pt; double spacing; justified body; left margin 3 cm | Faculty formatting rules |
| G12 | Chapter 5 completeness | Contains interpretation, literature comparison, clinical implications, strengths, limitations, and concluding remarks as identifiable sections | `WRITING_RULES.md` §3 |

## 2. Minimum chapter volume

Page figures assume A4, Times New Roman 12 pt, double spacing, 3 cm left margin
(≈ 230 words of prose per page). Tables and figures displace prose, so a chapter that carries
figures reaches its page target on fewer words.

| Section | Minimum words | Target pages | Notes |
|---------|--------------:|-------------:|-------|
| Front matter | — | 7–8 | Title, supervisors, acknowledgements, table of contents, list of tables, list of figures, list of abbreviations, abstract. Each occupies at least one page regardless of word count. |
| Ch 1 Introduction | 1,100 | 5 | Must end with aim, objectives, and null hypotheses |
| Ch 2 Review of literature | 3,400 | 14–16 | The largest chapter in a Master's thesis |
| Ch 3 Materials and methods | 1,500 | 6–7 | Reproducible by a third party without asking the candidate a single question |
| Ch 4 Results | 1,000 | 5–6 | Plus one figure per outcome measure |
| Ch 5 Discussion | 2,200 | 9–11 | Second largest chapter |
| Ch 6 Conclusions and recommendations | 400 | 2 | |
| Ch 7 References | — | 5 | 40+ entries at double spacing |
| **Total** | **≈ 9,600** | **45–48** | |

Word minima are a floor, not a licence to pad. Gate G13 below exists to stop padding.

| # | Gate | Criterion |
|---|------|-----------|
| G13 | No padding | Added length must be new scientific content: additional studies reviewed, additional method detail, additional interpretation. Restating the same point in new words, or recycling paragraphs between chapters, fails this gate as surely as being short. |

## 3. Required structural elements

### Front matter
1. Title page matching the approved protocol cover
2. Supervisors page with **real names and titles** — not dotted lines
3. Acknowledgements
4. Table of contents with page numbers
5. List of tables
6. List of figures
7. List of abbreviations
8. English abstract (250–350 words, structured: background, aim, methods, results, conclusions, keywords)

### Body
Chapters 1–7 as listed in §2.

### Back matter
- Arabic summary, **if** required by the Faculty of Dentistry, Pharos University in Alexandria.
  Confirm this with the department before binding; several Egyptian faculties require it and its
  absence is discovered late.
- Appendices: raw data tables, ethical approval / institutional clearance if applicable.

## 4. Figures

The current draft contains **no figures**. A Master's thesis reporting two quantitative outcome
measures needs, at minimum:

| Figure | Content |
|--------|---------|
| 3.1 | Teflon mould and disc specimen |
| 3.2 | Toothbrushing simulator with a disc mounted |
| 3.3 | Profilometer tracing setup |
| 4.1 | Bar chart, mean percentage weight loss ± SD, three groups |
| 4.2 | Bar chart or clustered column, mean Ra before and after brushing ± SD |
| 4.3 | Mean ΔRa ± SD, with the 0.2 µm plaque threshold drawn as a reference line |

Figures 4.1–4.3 can be generated from `locked_data.json` and require no new experiment.
Figures 3.1–3.3 require laboratory photographs from the candidate.

Every figure needs a caption below it, a number in order of first mention, and a citation in
the text.

## 5. Reference quality rules

1. **Real only.** Every entry must be traceable to a published source. Verify DOI, PubMed ID, or
   journal page range before it enters the list.
2. **Cited where it supports the claim.** A reference must actually contain the statement it is
   attached to. Standards must be cited for what they specify:
   - Toothbrush wear protocol → ISO/TR 14569-1
   - Dentifrice requirements → ISO 11609
   - Surface texture parameters and Ra definition → ISO 4287 / ISO 25178
   - Polymer-based restorative materials → ISO 4049 (**not** a profilometry reference)
3. **Balanced recency.** Landmark papers may be old — Bollen 1997 and Sexson & Phillips 1951 are
   legitimately foundational. But a 2025/2026 submission needs a substantial body of literature
   from the last five years.
4. **No filler citations.** Do not cite a paper that is not discussed. Padding the list to reach
   40 fails G4 and G13 together.
5. Manufacturer product information is acceptable for composition claims, but must carry the
   product name, manufacturer, city, and an access date.

## 6. Language register

`WRITING_RULES.md` requires natural, precise, formal academic English. Two failure modes both
violate it:

- **Inflated AI prose:** "it is important to note that", "plays a crucial role", "delve into".
- **Telegraphic fragments:** "Mass falls. Ra rises." "The test is not the same test." Strings of
  two- and three-word sentences do not read as a scholarly voice. They read as notes. A thesis
  requires connected argument in complete sentences with explicit logical links.

The target is the register of a *Dental Materials* discussion section: full sentences, subordinate
clauses where the logic needs them, and no ornament.

## 7. Verification procedure

```bash
python3 supervision/supervise_thesis.py --md Thesis_Complete.md --docx Thesis_Complete.docx
```

The script exits non-zero while any hard gate fails. Run it before every push and before any
delivery to the candidate or the supervisors.

Machine checks cannot certify G4 (reference reality), G13 (padding), or §6 (register). Those
require the human read described in `SUPERVISION_REPORT.md` §7.

## 8. Sign-off

Release requires all of:

- [ ] `supervise_thesis.py` exits 0
- [ ] Every reference independently verified against its source
- [ ] Full read-through for register and argument
- [ ] Supervisor names, acknowledgements, and cover confirmed by the candidate
- [ ] Laboratory details (curing light, polishing sequence, storage time, brush, slurry ratio,
      balance model, profilometer model, statistical software) supplied from the laboratory
      notebook and inserted
- [ ] Arabic summary requirement confirmed with the department
- [ ] Final PDF page count re-measured on the exact file being submitted
