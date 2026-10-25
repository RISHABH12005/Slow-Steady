from io import BytesIO
from datetime import date
from fastapi import FastAPI, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = FastAPI(title="Slow-Steady", description="JUET academic document generator")

PAGE = """<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Slow-Steady</title><style>
:root{font-family:Inter,Segoe UI,Arial,sans-serif;color:#172033;background:#f4f6fa}*{box-sizing:border-box}body{margin:0}.wrap{max-width:980px;margin:auto;padding:36px 20px}.hero{background:#14233b;color:white;padding:30px;border-radius:20px}.hero p{color:#d1d9e6;max-width:680px;line-height:1.6}.pill{display:inline-block;background:#dbeafe;color:#1d4ed8;border-radius:99px;padding:6px 10px;font-size:12px;font-weight:700}.panel{background:white;border:1px solid #e1e6ef;border-radius:16px;padding:24px;margin-top:20px}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}.full{grid-column:1/-1}label{display:block;font-weight:650;font-size:13px;margin-bottom:7px}input,textarea,select{width:100%;padding:12px;border:1px solid #cbd5e1;border-radius:9px;font:inherit;background:#fff}textarea{min-height:100px;resize:vertical}button{border:0;background:#2563eb;color:white;font-weight:700;padding:13px 18px;border-radius:10px;cursor:pointer}button:hover{background:#1d4ed8}.muted{color:#64748b;font-size:13px;line-height:1.6}.section{margin-top:22px;padding-top:18px;border-top:1px solid #e5eaf1}h1{margin:12px 0;font-size:clamp(30px,5vw,46px)}h2{margin-top:0}footer{padding:24px 0;color:#64748b;font-size:12px}@media(max-width:650px){.grid{grid-template-columns:1fr}.full{grid-column:auto}.wrap{padding:18px 12px}.panel,.hero{padding:20px}}
</style></head><body><main class="wrap"><header class="hero"><span class="pill">JUET · DOCX generator MVP</span><h1>Less formatting.<br>More project work.</h1><p>Enter your project details once and generate an editable internship report structure using the supplied JUET summer internship guidelines. This MVP creates a formatted starting document; it does not invent project content or approvals.</p></header>
<form class="panel" method="post" action="/generate/internship-report"><h2>Internship report details</h2><p class="muted">Required fields are marked by the browser. You can leave optional sections blank and complete them later in Word.</p><div class="grid">
<div><label>Project title *</label><input name="project_title" required placeholder="Specific project title"></div>
<div><label>Organization *</label><input name="organization" required placeholder="Organization name"></div>
<div><label>Student name *</label><input name="student_name" required placeholder="Full name"></div>
<div><label>Enrollment number *</label><input name="enrollment" required placeholder="e.g. 231B264"></div>
<div><label>Supervisor / guide</label><input name="guide" placeholder="Guide name"></div>
<div><label>Degree / department</label><input name="department" value="Computer Science and Engineering"></div>
<div><label>Report year</label><input name="year" value="2026" type="number"></div>
<div><label>Internship period</label><input name="period" placeholder="e.g. 1 June – 31 July 2026"></div>
<div class="full"><label>Executive summary</label><textarea name="summary" placeholder="Purpose, work performed, outcomes and learning. Use only verified information."></textarea></div>
<div class="full"><label>Project overview</label><textarea name="overview" placeholder="Problem, intended users, objectives and scope"></textarea></div>
<div class="full"><label>Work performed</label><textarea name="work" placeholder="Describe implementation activities, decisions, testing and evidence"></textarea></div>
<div class="full"><label>Results and discussion</label><textarea name="results" placeholder="Observed results and limitations. Do not claim unverified outcomes."></textarea></div>
<div class="full"><label>Conclusion</label><textarea name="conclusion" placeholder="Summary and learning outcomes"></textarea></div>
<div class="full"><label>References</label><textarea name="references" placeholder="One reference per line. Verify bibliographic details."></textarea></div>
</div><div class="section"><button type="submit">Generate editable DOCX</button><p class="muted">Applies A4 page setup, Times New Roman, guideline margins and heading styles. Page numbering and automatic table of contents still need validation in Word/LibreOffice in this early MVP.</p></div></form>
<section class="panel"><h2>Current scope</h2><p>First milestone: JUET internship report DOCX structure. Major/minor project report profiles, PPTX generation, PDF guideline import, saved workspaces and full compliance validation are planned, not yet implemented.</p><p class="muted">Academic integrity: review every section, add required evidence and certificate, and confirm the latest department rules before submission.</p></section><footer>Slow-Steady · Open-source academic document automation</footer></main></body></html>"""

def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)

def configure_doc(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
    sec.top_margin = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.right_margin = Inches(1)
    sec.left_margin = Inches(1.5)
    styles = doc.styles
    normal = styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.15
    for name, size in [('Heading 1',16),('Heading 2',14),('Heading 3',12)]:
        st = styles[name]
        st.font.name = 'Times New Roman'
        st.font.size = Pt(size)
        st.font.bold = True
        st.paragraph_format.keep_with_next = True
    # Footer placeholder; exact field numbering is inserted as a Word field.
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run()
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE')
    run._r.addnext(fld)

def add_centered(doc, text, size=12, bold=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.bold = bold
    r.font.name = 'Times New Roman'
    r.font.size = Pt(size)
    return p

def add_body(doc, text):
    if text and text.strip():
        for para in text.strip().split('\n'):
            doc.add_paragraph(para.strip())

@app.get("/", response_class=HTMLResponse)
def home():
    return HTMLResponse(PAGE)

@app.post("/generate/internship-report")
def generate_internship_report(
    project_title: str = Form(...),
    organization: str = Form(...),
    student_name: str = Form(...),
    enrollment: str = Form(...),
    guide: str = Form(""),
    department: str = Form("Computer Science and Engineering"),
    year: str = Form(str(date.today().year)),
    period: str = Form(""),
    summary: str = Form(""),
    overview: str = Form(""),
    work: str = Form(""),
    results: str = Form(""),
    conclusion: str = Form(""),
    references: str = Form("")
):
    doc = Document()
    configure_doc(doc)

    # Cover page based on the supplied JUET internship-report annexure.
    add_centered(doc, project_title.upper(), 16, True)
    doc.add_paragraph()
    add_centered(doc, "Completed At", 14)
    add_centered(doc, organization, 16)
    doc.add_paragraph()
    add_centered(doc, "A REPORT", 14, True)
    add_centered(doc, "Submitted by", 14)
    add_centered(doc, f"{student_name} ({enrollment})", 14)
    if guide.strip():
        doc.add_paragraph()
        add_centered(doc, f"Under the guidance of: {guide}", 14)
    doc.add_paragraph()
    add_centered(doc, "Submitted in partial fulfillment for the award of the degree of", 12)
    add_centered(doc, "BACHELOR OF TECHNOLOGY", 14, True)
    add_centered(doc, "IN", 12, True)
    add_centered(doc, department.upper(), 14, True)
    doc.add_paragraph()
    add_centered(doc, "JAYPEE UNIVERSITY OF ENGINEERING AND TECHNOLOGY", 14, True)
    add_centered(doc, f"Guna, Madhya Pradesh · {year}", 12)
    doc.add_page_break()

    doc.add_heading("DECLARATION", 1)
    doc.add_paragraph(
        f'I hereby declare that the work reported in the summer internship report entitled '
        f'“{project_title}”, submitted in partial fulfilment for the award of the B.Tech degree '
        f'at Jaypee University of Engineering and Technology, Guna, is presented to the best '
        f'of my knowledge and belief. I understand that I am responsible for ensuring the '
        f'originality and accuracy of this report.'
    )
    doc.add_paragraph("\n\nSignature of the Student")
    doc.add_paragraph(f"Place: ____________________\nDate: ____________________")
    doc.add_page_break()

    doc.add_heading("INTERNSHIP COMPLETION CERTIFICATE", 1)
    doc.add_paragraph("[Insert the authentic organization-issued internship completion certificate here.]")
    doc.add_page_break()

    doc.add_heading("ACKNOWLEDGEMENT", 1)
    doc.add_paragraph("[Write a truthful acknowledgement of the people and organizations that supported the work.]")
    doc.add_page_break()

    doc.add_heading("EXECUTIVE SUMMARY", 1)
    add_body(doc, summary or "[Write a one-page summary of the internship purpose, work performed and learning outcomes.]")
    doc.add_page_break()

    doc.add_heading("TABLE OF CONTENTS", 1)
    doc.add_paragraph("[In Microsoft Word, use References → Table of Contents to insert/update an automatic table of contents after completing the report.]")
    doc.add_paragraph("Declaration\nInternship Completion Certificate\nAcknowledgement\nExecutive Summary\n1. Introduction\n2. Description of Organization\n3. Description of Work\n4. System Analysis and Design\n5. Implementation and Testing\n6. Results and Discussion\n7. Conclusion\n8. References\n9. Appendices")
    doc.add_page_break()

    doc.add_heading("1. INTRODUCTION", 1)
    doc.add_heading("1.1 Project Overview", 2)
    add_body(doc, overview or "[Describe the specific project, problem, objectives, intended users and scope.]")
    doc.add_heading("1.2 Hardware Specification", 2)
    doc.add_paragraph("[List only the hardware relevant to the project and explain its role.]")
    doc.add_heading("1.3 Software Specification", 2)
    doc.add_paragraph("[List relevant software and versions briefly; focus the report on work performed.]")
    doc.add_heading("1.4 Tools and Technologies Used", 2)
    doc.add_paragraph("[Describe only tools relevant to implementation, as required by the guideline.]")

    doc.add_heading("2. DESCRIPTION OF ORGANIZATION", 1)
    doc.add_heading("2.1 About the Organization", 2)
    doc.add_paragraph(f"[Describe {organization}, using verified information.]")
    doc.add_heading("2.2 Team Structure and Work Culture", 2)
    doc.add_paragraph("[Describe the team and working practices based on your actual experience.]")

    doc.add_heading("3. DESCRIPTION OF WORK", 1)
    add_body(doc, work or "[Describe project activities in logical sections. Add labelled diagrams, tables, screenshots and evidence where relevant.]")

    doc.add_heading("4. SYSTEM ANALYSIS AND DESIGN", 1)
    doc.add_paragraph("[For application-based projects, describe requirements, architecture, design methodology and relevant DFD/UML diagrams. For research-based projects, replace this chapter with Proposed Methodology.]")
    doc.add_paragraph("[Insert authentic diagrams and refer to each figure/table in the text.]")

    doc.add_heading("5. IMPLEMENTATION AND TESTING", 1)
    doc.add_paragraph("[Explain implementation steps and test cases with actual outcomes. Do not include source code in the report.]")

    doc.add_heading("6. RESULTS AND DISCUSSION", 1)
    add_body(doc, results or "[Present verified results, screenshots, tables or measurements, and interpret what they demonstrate. For research-based projects, compare results with existing work.]")

    doc.add_heading("7. CONCLUSION", 1)
    add_body(doc, conclusion or "[Summarize completed work, outcomes, limitations and learning.]")

    doc.add_heading("8. REFERENCES", 1)
    if references.strip():
        for line in references.strip().splitlines():
            doc.add_paragraph(line.strip())
    else:
        doc.add_paragraph("[Add every source cited in the report. Verify authors, titles, publication details and citation order.]")

    doc.add_heading("9. APPENDICES", 1)
    doc.add_paragraph("[Add supplementary material as Appendix A, Appendix B, etc. Include each appendix in the table of contents.]")
    doc.add_paragraph("\n")
    doc.add_paragraph("Review checklist before submission:")
    for item in [
        "Confirm the latest department guideline and required report length (the supplied guideline states at least 40 pages after the index).",
        "Insert the authentic completion certificate and required university/organization material.",
        "Add labelled figures and tables and reference them in the text.",
        "Update the table of contents and verify Roman/Arabic page numbering in Word or LibreOffice.",
        "Verify every statement, result and reference; remove all placeholders.",
    ]:
        doc.add_paragraph(item, style='List Bullet')

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    filename = "Slow-Steady-JUET-Internship-Report.docx"
    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )
