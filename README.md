# Slow-Steady — Academic Document Studio

Slow-Steady is an open-source document automation app for students. It generates editable JUET-focused DOCX report drafts and a text-only Major Project PPTX from reusable project details.

## Current features

- Dark-mode responsive web interface.
- Summer Internship, Minor I, Minor II, Major I and Major II report profiles.
- Major Project presentation generator with 13 slides arranged around the supplied instructions' nine sections.
- DOCX formatting: A4 page setup, 1-inch top/bottom/right margins, 1.5-inch left margin, Times New Roman 12 pt body, 1.15 line spacing, 16 pt Heading 1 and 14 pt Heading 2.
- PPTX generation using editable text slides; no AI-generated images are inserted.
- Basic validation endpoint, profile API and automated tests.
- Render web service configured for automatic deployment from main.

## Deployment

Render service: https://slow-steady.onrender.com  
Health endpoint: https://slow-steady.onrender.com/health  
API docs: https://slow-steady.onrender.com/docs

A deployment being configured does not guarantee the latest build is healthy. Verify the health endpoint and generate sample files before relying on the site.

## Run locally

Requires Python 3.10 or later.

```powershell
git clone https://github.com/RISHABH12005/Slow-Steady.git
cd Slow-Steady
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m uvicorn app:app --reload
```

Open http://127.0.0.1:8000.

Run tests with `python -m pytest -q`.

## API

- `GET /` — dark-mode web app
- `GET /health` — health/status
- `GET /api/profiles` — report and PPT section definitions
- `POST /generate` — form-encoded fields; returns DOCX or PPTX according to `document_type`
- `POST /api/validate` — basic missing-content checks

Accepted `document_type` values: `internship`, `minor1`, `minor2`, `major1`, `major2`, `major_ppt`.

## JUET rules and accuracy notes

The supplied Summer Internship Report Guidelines specify A4 portrait, Times New Roman, 1-inch top/bottom/right margins, 1.5-inch left margin, 16 pt bold chapter names, 14 pt bold headings, 12 pt normal text and 1.15 line spacing. They specify centered bottom page numbers, lowercase Roman numerals before the table of contents, Arabic numbering starting at 1 after it, and at least 40 pages after the index.

The supplied Major Project PPT instructions specify Introduction (1 slide), Problem Statement (1), Literature Review (1–2), Work Plan (1), Individual Contribution (1), Work Done So Far (3–4), Work to be Done (1), Conclusion (1), and References (1 in IEEE format). This generator uses a 13-slide version: two literature review slides and four work-completed slides.

**The supplied PPT instructions prohibit AI-generated images in any part of the presentation.** This generator creates text-only slides. The Minor/Major report structures are starter profiles until checked against their exact official report guidelines; do not treat them as certified JUET formats.

## Important limitations

- The generated DOCX includes a table-of-contents placeholder; insert/update the automatic TOC in Word or LibreOffice.
- Roman numeral front matter numbering and Arabic numbering after the contents still require final validation.
- The internship guideline states at least 40 pages after the index; generated starter drafts do not automatically meet that length.
- Validation checks are basic, not full semantic or layout compliance.
- Inspect generated files, replace placeholders, verify sources, and check current departmental rules before submission.
- The tool must not invent certificates, signatures, approvals, results, citations or individual contributions.

## Stack

Python, FastAPI, python-docx, python-pptx, pytest. No AI API or Ollama is currently required to generate files; local AI assistance can be added later.

## License

MIT.
