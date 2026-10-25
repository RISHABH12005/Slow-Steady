# Slow-Steady — University Document Automation

Slow-Steady reduces repetitive formatting work for university reports and presentations. The first implementation is a local web app that generates an editable JUET summer internship report DOCX from a form.

## MVP status

Implemented:
- FastAPI web interface for internship report details.
- DOCX export using `python-docx`.
- A4 layout, JUET guideline margins, Times New Roman, 12 pt body text, 1.15 line spacing and styled headings.
- Cover page and report sections with placeholders for authentic certificates, student-written content, references and appendices.
- Page-number field and a review checklist.

Not yet implemented:
- PPTX generation.
- Minor I/II and Major I/II-specific report profiles.
- Automatic table of contents and Roman-to-Arabic section numbering.
- PDF guideline import, saved projects or full compliance validation.

## Run locally

Requires Python 3.10 or later.

```bash
python -m venv .venv
# Windows PowerShell:
.venv\Scripts\Activate.ps1
# macOS/Linux:
# source .venv/bin/activate

python -m pip install -r requirements.txt
uvicorn app:app --reload
```

Open http://127.0.0.1:8000.

If `python` is not found on Windows, install Python from python.org with “Add Python to PATH”, or use the Python Launcher command `py -m venv .venv`.

## JUET guideline notes

The supplied Summer Internship Report Guidelines specify A4 portrait pages, Times New Roman, 1-inch top/bottom/right margins, 1.5-inch left margin, 16 pt bold chapter names, 14 pt bold headings, 12 pt normal text and 1.15 line spacing. They specify centered page numbers, lowercase Roman numerals before the table of contents, Arabic numbering starting at 1 after the contents, and a minimum of 40 pages after the index.

The supplied Major Project PPT instructions require nine sections and explicitly prohibit AI-generated images in any part of the presentation. The future PPT generator must preserve this constraint.

## Academic integrity

The generator does not create official certificates, signatures, approvals, project results or verified references. Replace bracketed placeholders with truthful content and confirm current department rules before submission. The generated report is a draft, not a guarantee of compliance.

## Stack

Python, FastAPI, python-docx. Planned: python-pptx, PDF extraction and optional local Ollama integration.

## License

MIT.
