# Mohamed Talaat — Master's thesis

Comparative study of wear resistance and surface roughness of injectable versus conventional composite resin (in vitro study). Faculty of Dentistry, Pharos University in Alexandria.

The file to send the supervisors is `Thesis_Complete.docx` (PDF beside it). Pages 1–2 are the approved protocol covers. Chapters 1–7, English and Arabic abstracts, and the Arabic summary are in the same Word file.

## Final assembled files

- [Thesis_Complete.docx](Thesis_Complete.docx) — A4, Times New Roman 12 pt, double spacing, 3 cm left margin
- [Thesis_Complete.pdf](Thesis_Complete.pdf) — same document as PDF
- [Thesis_Complete.md](Thesis_Complete.md) — source text of the chapters

Rebuild Word/PDF after editing the markdown:

```
python3 build_thesis_docx.py
soffice --headless --convert-to pdf Thesis_Complete.docx
```
