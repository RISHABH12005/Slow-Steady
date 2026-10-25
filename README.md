# Slow-Steady — University Document Automation

Slow-Steady is an open-source AI-assisted tool to reduce the repetitive work of preparing university project reports and presentations. Start with JUET's supplied report and PPT examples, enter project details once, and generate editable DOCX and PPTX files that follow selected guidelines.

> The initial profile targets JUET (Jaypee University of Engineering & Technology, Guna). Generated documents are drafts for students to review; students remain responsible for accuracy, originality, and confirming requirements with their department.

## Problem

Students repeatedly format Minor I/II reports, Major I/II reports, internship reports, and presentations. The content differs, but document structures and formatting rules overlap. Slow-Steady aims to automate the repetitive formatting—not fabricate project work or bypass academic requirements.

## Initial scope

- JUET profile for summer internship reports, based on the uploaded *Summer Internship Report Guidelines*.
- Project report profiles based on the supplied Minor I and Minor II examples.
- Internship report cover and front matter based on the supplied internship examples.
- Major Project presentation profile based on the supplied *PPT Instructions Major* and presentation sample.
- Editable `.docx` and `.pptx` export.
- Reusable project details, student/team information, guide, organization, dates, chapters, references, diagrams, screenshots, and contribution tables.
- Rule-based validation with a checklist showing passed, failed, and needs-review rules.

## Important rules captured from the supplied JUET resources

### Summer internship report

The supplied guideline asks for a cover/title page, student declaration, training-completion certificate, acknowledgement, one-page executive summary, table of contents, detailed project chapters, conclusion, references, and appendices. It states A4 portrait, Times New Roman, 1-inch top/bottom/right margins, 1.5-inch left margin, 16 pt bold chapter names, 14 pt bold headings, 12 pt normal text, and 1.15 line spacing. It also specifies centered page numbers at the bottom, lowercase Roman numerals before the table of contents, and Arabic numbering from 1 after the contents. The guideline also requires a minimum of 40 pages after the index.

### Major Project PPT

The provided PPT instructions specify:
1. Introduction — 1 slide
2. Problem Statement — 1 slide
3. Literature Review — 1–2 slides
4. Work Plan — 1 slide, aligned with weekly reports
5. Individual Contribution — 1 slide
6. Work Done So Far — 3–4 slides (different evidence for application-based vs research-based projects)
7. Work to be Done — 1 slide
8. Conclusion — 1 slide
9. References — 1 slide in IEEE format

Slides should be concise, readable, visually consistent with the university template, and use diagrams/charts/visuals where needed. The supplied instructions explicitly prohibit AI-generated images in any part of the presentation. Slow-Steady should preserve this requirement and must not generate AI images for these slides.

## Planned workflow

1. Choose document type and guideline profile.
2. Fill in a project workspace once.
3. Import or draft content section-by-section; distinguish student-supplied facts from AI suggestions.
4. Generate editable DOCX/PPTX.
5. Run guideline checks and display a compliance report.
6. Review the files in Microsoft Word/PowerPoint or compatible software before submission.

## Suggested architecture

- Python + FastAPI backend
- `python-docx` for DOCX generation
- `python-pptx` for PPTX generation
- PyMuPDF and parsers for extracting PDF guidelines
- Ollama with a suitable open-weight model for local AI-assisted section planning and summarization
- React/Next.js frontend (or begin with a simpler local UI)
- SQLite for local project settings and profiles

Formatting must be implemented deterministically from structured rules. An LLM may interpret or help draft content, but should not decide page margins, pagination, numbering, heading sizes, or other compliance-critical formatting at generation time.

## Roadmap

- [ ] Extract explicit JUET report/PPT rules into versioned JSON profiles.
- [ ] Build the project metadata form and reusable workspace.
- [ ] Generate the JUET internship report DOCX.
- [ ] Generate the Major Project PPTX from supplied project content.
- [ ] Add project report and Minor I/II profiles after verifying their complete requirements.
- [ ] Add validation checks and clear warnings for ambiguous or conflicting source examples.
- [ ] Add tests using example inputs and verify generated files by rendering them to PDF.
- [ ] Document local installation and optional Ollama setup.

## Academic integrity and privacy

- Never invent certificates, signatures, supervisor approvals, results, references, attendance, or project contributions.
- Clearly label AI-drafted text and require student review.
- Do not create AI-generated images for JUET Major Project presentations because the supplied instructions prohibit them.
- Prefer local processing for private reports. Do not upload documents to a hosted model unless the user explicitly configures that service.
- Retain original source files and explain the guideline source used for each check.

## Current status

This repository currently contains the initial project README and MIT license. The application is not yet implemented. Features in the roadmap are planned, not available functionality.

## License

MIT. See [LICENSE](LICENSE).
