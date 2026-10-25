from __future__ import annotations
from io import BytesIO
from datetime import date
from html import escape
from fastapi import FastAPI, Form
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from pptx.dml.color import RGBColor

app = FastAPI(title="Slow-Steady", description="JUET academic report and presentation generator", version="0.2.0")

REPORT_PROFILES = {
    "internship": {"label":"Summer Internship Report", "chapters":[
        ("1. INTRODUCTION", "Describe the specific project, problem, objectives and scope."),
        ("2. DESCRIPTION OF ORGANIZATION", "Use verified facts about the organization and team."),
        ("3. DESCRIPTION OF WORK", "Describe internship activities in logical sections."),
        ("4. SYSTEM ANALYSIS AND DESIGN", "Application-based: architecture, methodology and DFD/UML. Research-based: Proposed Methodology."),
        ("5. IMPLEMENTATION AND TESTING", "Explain implementation and tests with actual results; do not include source code."),
        ("6. RESULTS AND DISCUSSION", "Present verified results and interpretation. Research-based work should compare existing work."),
        ("7. CONCLUSION", "Summarize work, outcomes, limitations and learning."),
        ("8. REFERENCES", "List sources actually cited; verify details and citation order."),
        ("9. APPENDICES", "Supplementary material labelled Appendix A, Appendix B, etc.")
    ]},
    "minor1": {"label":"Minor Project I Report", "chapters":[
        ("1. INTRODUCTION","Problem context, motivation, objectives and scope."),
        ("2. PROBLEM STATEMENT","Define the problem and intended users."),
        ("3. LITERATURE REVIEW","Summarize relevant prior work with verified citations."),
        ("4. PROPOSED METHODOLOGY","Describe the planned approach and rationale."),
        ("5. WORK PLAN","Milestones, timeline and deliverables."),
        ("6. EXPECTED OUTCOMES","State intended outcomes, not fabricated results."),
        ("7. CONCLUSION","Summarize the proposal and next steps."),
        ("8. REFERENCES","List sources actually cited.")
    ]},
    "minor2": {"label":"Minor Project II Report", "chapters":[
        ("1. INTRODUCTION","Project context, objectives and scope."),
        ("2. LITERATURE REVIEW","Prior work and identified gap."),
        ("3. SYSTEM ANALYSIS AND DESIGN","Requirements, architecture, design and diagrams."),
        ("4. IMPLEMENTATION / METHODOLOGY","Describe actual work completed; adapt for research projects."),
        ("5. TESTING / ANALYSIS","Test cases, experiments and verified outcomes."),
        ("6. RESULTS AND DISCUSSION","Interpret results and note limitations."),
        ("7. CONCLUSION AND FUTURE WORK","Summarize progress and next steps."),
        ("8. REFERENCES","List sources actually cited.")
    ]},
    "major1": {"label":"Major Project I Report", "chapters":[
        ("1. INTRODUCTION","Context, motivation, problem and objectives."),
        ("2. PROBLEM STATEMENT","Specific problem, scope and constraints."),
        ("3. LITERATURE REVIEW","Prior work and research/engineering gap."),
        ("4. REQUIREMENTS AND METHODOLOGY","Requirements, proposed method, design and plan."),
        ("5. WORK COMPLETED","Evidence-backed progress and contribution of each member."),
        ("6. WORK PLAN","Remaining milestones and timeline."),
        ("7. CONCLUSION","Current status, risks and next steps."),
        ("8. REFERENCES","Verified sources in the department-required style.")
    ]},
    "major2": {"label":"Major Project II Report", "chapters":[
        ("1. INTRODUCTION","Context, problem, objectives and scope."),
        ("2. LITERATURE REVIEW","Prior work and gap."),
        ("3. SYSTEM ANALYSIS AND DESIGN","Architecture, methodology and relevant diagrams."),
        ("4. IMPLEMENTATION / EXPERIMENTS","Actual implementation or research experiments."),
        ("5. TESTING AND EVALUATION","Methods, test cases, measurements and results."),
        ("6. RESULTS AND DISCUSSION","Interpretation, comparisons and limitations."),
        ("7. CONCLUSION AND FUTURE WORK","Conclusions supported by evidence."),
        ("8. REFERENCES","Verified sources in the required style."),
        ("9. APPENDICES","Supplementary material.")
    ]}
}
PPT_SECTIONS = [
 ("Introduction",1,"Project context, motivation and objective."),
 ("Problem Statement",1,"Specific problem, users, scope and constraints."),
 ("Literature Review",2,"Prior work and the gap; use 1–2 slides."),
 ("Work Plan",1,"Timeline and milestones aligned with weekly reports."),
 ("Individual Contribution",1,"Clearly identify each member's actual contribution."),
 ("Work Done So Far",4,"Use 3–4 slides with evidence. Application-based: design, implementation, testing. Research-based: methodology, experiments, analysis."),
 ("Work to be Done",1,"Remaining tasks, milestones and risks."),
 ("Conclusion",1,"Current status and evidence-backed conclusion."),
 ("References",1,"References in IEEE format.")
]

CSS = """
:root{color-scheme:dark;--bg:#090d16;--panel:#111827;--panel2:#151f30;--line:#273449;--text:#eef3fb;--muted:#9baac0;--blue:#83aaff;--green:#75ddb0;font-family:Inter,ui-sans-serif,system-ui,Segoe UI,Arial,sans-serif}
*{box-sizing:border-box}body{margin:0;background:radial-gradient(ellipse at 80% -10%,#172949 0,transparent 42%),var(--bg);color:var(--text)}
.wrap{max-width:1120px;margin:auto;padding:28px 20px 60px}.top{display:flex;justify-content:space-between;align-items:center;gap:16px;margin-bottom:28px}.brand{font-weight:800;letter-spacing:-.04em;font-size:20px}.brand span{color:var(--blue)}.tag{border:1px solid var(--line);border-radius:999px;padding:7px 11px;color:var(--muted);font-size:12px}
.hero{padding:34px;border:1px solid #2a3b56;border-radius:22px;background:linear-gradient(120deg,#121e31,#111827 65%);box-shadow:0 22px 60px #0003}.eyebrow{font-size:12px;text-transform:uppercase;letter-spacing:.16em;color:var(--blue);font-weight:750}h1{font-size:clamp(34px,5vw,56px);line-height:1.04;letter-spacing:-.05em;margin:15px 0}h2{font-size:21px;letter-spacing:-.02em;margin:0 0 10px}h3{font-size:15px;margin:0 0 8px}.hero p,.muted{color:var(--muted);line-height:1.65;max-width:760px}.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-top:24px}.stat{padding:14px;background:#0b1220;border:1px solid var(--line);border-radius:12px}.stat strong{display:block;font-size:18px}.stat span{font-size:12px;color:var(--muted)}
.tabs{display:flex;gap:8px;flex-wrap:wrap;margin:22px 0}.tab{border:1px solid var(--line);border-radius:9px;padding:10px 13px;color:var(--muted);text-decoration:none}.tab.active,.tab:hover{background:#1a2c48;color:var(--text);border-color:#46638c}
.layout{display:grid;grid-template-columns:minmax(0,1.5fr) minmax(260px,.8fr);gap:18px;align-items:start}.panel{background:color-mix(in srgb,var(--panel) 95%,transparent);border:1px solid var(--line);border-radius:16px;padding:22px;margin-bottom:18px;min-width:0}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:15px}.full{grid-column:1/-1}label{display:block;font-weight:650;font-size:12px;color:#d9e3f1;margin-bottom:7px}input,textarea,select{width:100%;padding:12px 13px;border:1px solid #33435b;border-radius:9px;background:#0b1220;color:var(--text);font:inherit;outline:none}input:focus,textarea:focus,select:focus{border-color:var(--blue);box-shadow:0 0 0 3px #83aaff18}textarea{min-height:88px;resize:vertical}button{font:inherit;cursor:pointer}.primary{background:var(--blue);color:#07101f;border:0;padding:13px 17px;border-radius:10px;font-weight:800}.secondary{background:#182337;color:var(--text);border:1px solid var(--line);padding:12px 15px;border-radius:10px;font-weight:700}.primary:hover{filter:brightness(1.08)}.row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}.muted{font-size:13px}.section{padding-top:18px;margin-top:18px;border-top:1px solid var(--line)}.list{display:grid;gap:10px}.item{padding:12px;background:#0b1220;border:1px solid var(--line);border-radius:10px}.item p{font-size:13px;color:var(--muted);margin:5px 0 0;line-height:1.5}.check{color:var(--green)}footer{color:#7889a1;font-size:12px;padding:12px 0}@media(max-width:780px){.layout{grid-template-columns:1fr}.hero{padding:24px}.stats{grid-template-columns:1fr 1fr}}@media(max-width:520px){.wrap{padding:16px 12px 40px}.grid{grid-template-columns:1fr}.full{grid-column:auto}.stats{grid-template-columns:1fr}.top{align-items:flex-start}}
"""
PAGE = """<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Slow-Steady · Academic Document Studio</title><style>""" + CSS + """</style></head><body><main class="wrap">
<header class="top"><div class="brand">slow<span>·</span>steady</div><div class="tag">JUET document studio · v0.2</div></header>
<section class="hero"><div class="eyebrow">Academic document automation</div><h1>Spend less time<br>formatting documents.</h1><p>Generate editable reports and presentations from a reusable project profile. The first version includes JUET report templates and the supplied Major Project presentation structure, with explicit placeholders where your own verified content is required.</p><div class="stats"><div class="stat"><strong>5</strong><span>report profiles</span></div><div class="stat"><strong>DOCX + PPTX</strong><span>editable exports</span></div><div class="stat"><strong>Rule-based</strong><span>format validation</span></div></div></section>
<nav class="tabs"><a class="tab active" href="/">Generator</a><a class="tab" href="/health">System status</a><a class="tab" href="/docs">API docs</a></nav>
<div class="layout"><form class="panel" method="post" action="/generate"><h2>Build your document</h2><p class="muted">Fill in project details once. Select a report profile or the Major Project presentation.</p>
<div class="grid">
<div class="full"><label>Document type</label><select name="document_type" id="document_type" onchange="toggleFormat()"><option value="internship">Summer Internship Report · DOCX</option><option value="minor1">Minor Project I Report · DOCX</option><option value="minor2">Minor Project II Report · DOCX</option><option value="major1">Major Project I Report · DOCX</option><option value="major2">Major Project II Report · DOCX</option><option value="major_ppt">Major Project Presentation · PPTX</option></select></div>
<div class="full"><label>Project title *</label><input name="project_title" required placeholder="A specific, meaningful project title"></div>
<div><label>Student / presenter *</label><input name="student_name" required placeholder="Full name"></div><div><label>Enrollment number</label><input name="enrollment" placeholder="Enrollment number"></div>
<div><label>Organization / company</label><input name="organization" placeholder="Internship organization (if applicable)"></div><div><label>Supervisor / project guide</label><input name="guide" placeholder="Guide or supervisor"></div>
<div><label>Department</label><input name="department" value="Computer Science and Engineering"></div><div><label>Academic year</label><input name="year" type="number" value="2026"></div>
<div class="full"><label>Project overview / problem statement</label><textarea name="overview" placeholder="Problem, intended users, objectives and scope"></textarea></div>
<div class="full"><label>Work completed / methodology</label><textarea name="work" placeholder="Actual work completed, design, implementation or research methodology"></textarea></div>
<div class="full"><label>Results / analysis</label><textarea name="results" placeholder="Verified findings, test results, measurements and limitations"></textarea></div>
<div class="full"><label>Individual contribution / work plan</label><textarea name="contribution" placeholder="Actual contribution by each member and/or milestones aligned with weekly reports"></textarea></div>
<div class="full"><label>Conclusion / future work</label><textarea name="conclusion" placeholder="Evidence-backed conclusions and remaining work"></textarea></div>
<div class="full"><label>References (one per line)</label><textarea name="references" placeholder="Paste verified references; PPT output uses an IEEE-style references slide"></textarea></div>
</div><div class="section row"><button class="primary" type="submit">Generate document ↗</button><button class="secondary" type="reset">Clear form</button></div><p class="muted">Generated files are drafts. Replace placeholders, verify content, and confirm current department requirements before submission.</p></form>
<aside><section class="panel"><h2>Profiles included</h2><div class="list"><div class="item"><h3>Internship report</h3><p>Front matter, organization, work description, design, implementation/testing, results, references and appendices.</p></div><div class="item"><h3>Minor I / Minor II</h3><p>Proposal and progress-focused chapter structures; check against your official report format.</p></div><div class="item"><h3>Major I / Major II</h3><p>Planning/progress and final implementation/evaluation report structures.</p></div><div class="item"><h3>Major Project PPT</h3><p>Nine required sections, including work plan, individual contribution, work completed and IEEE references.</p></div></div></section>
<section class="panel"><h2>Compliance checks</h2><div class="list"><div class="item"><span class="check">✓</span> DOCX page setup and typography</div><div class="item"><span class="check">✓</span> Required section headings</div><div class="item"><span class="check">✓</span> Missing-content warnings</div><div class="item"><span class="check">✓</span> PPT slide order and count</div><div class="item"><span class="check">!</span> Page numbering and TOC need final review in Word/LibreOffice</div></div><p class="muted">JUET's supplied Major Project PPT instructions prohibit AI-generated images. This generator uses text-only slides and does not insert AI-generated images.</p></section></aside></div><footer>Slow-Steady · Open-source academic document automation · Review generated files before submission.</footer></main>
<script>function toggleFormat(){}</script></body></html>"""

def set_page_field(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = paragraph.add_run()
    fld = OxmlElement("w:fldSimple"); fld.set(qn("w:instr"), "PAGE"); r._r.addnext(fld)

def configure_doc(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
    sec.top_margin = Inches(1); sec.bottom_margin = Inches(1)
    sec.right_margin = Inches(1); sec.left_margin = Inches(1.5)
    normal = doc.styles["Normal"]; normal.font.name = "Times New Roman"; normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.15
    for name, size in [("Heading 1",16),("Heading 2",14),("Heading 3",12)]:
        st=doc.styles[name]; st.font.name="Times New Roman"; st.font.size=Pt(size); st.font.bold=True
        st.paragraph_format.keep_with_next=True
    set_page_field(sec.footer.paragraphs[0])

def centered(doc, text, size=12, bold=False):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(text); r.bold=bold; r.font.name="Times New Roman"; r.font.size=Pt(size)
    return p

def add_text(doc, text, placeholder):
    text=(text or "").strip()
    if not text:
        doc.add_paragraph("[To complete: " + placeholder + "]")
    else:
        for part in text.splitlines():
            if part.strip(): doc.add_paragraph(part.strip())

def build_report(data):
    doc=Document(); configure_doc(doc)
    typ=data["document_type"]; profile=REPORT_PROFILES[typ]
    centered(doc,data["project_title"].upper(),16,True); doc.add_paragraph()
    centered(doc,profile["label"],14,True); centered(doc,"Submitted by",14)
    centered(doc,f'{data["student_name"]} ({data["enrollment"]})'.strip(),14)
    if data["guide"]: centered(doc,"Under the guidance of: "+data["guide"],14)
    doc.add_paragraph(); centered(doc,"JAYPEE UNIVERSITY OF ENGINEERING AND TECHNOLOGY",14,True)
    centered(doc,data["department"].upper(),12); centered(doc,f'Guna, Madhya Pradesh · {data["year"]}',12)
    doc.add_page_break()
    doc.add_heading("DOCUMENT REVIEW CHECKLIST",1)
    doc.add_paragraph("This file is a draft generated from the selected profile. Replace every bracketed instruction, verify all facts and references, and check current departmental rules.")
    doc.add_page_break()
    if typ=="internship":
        doc.add_heading("DECLARATION",1)
        doc.add_paragraph(f'I hereby declare that the work reported in “{data["project_title"]}” is presented to the best of my knowledge and belief. I am responsible for ensuring the originality and accuracy of this report.')
        doc.add_paragraph("Signature of the Student\n\nPlace: ____________________\nDate: ____________________"); doc.add_page_break()
        doc.add_heading("INTERNSHIP COMPLETION CERTIFICATE",1); doc.add_paragraph("[Insert the authentic organization-issued completion certificate. Do not fabricate a certificate.]"); doc.add_page_break()
        doc.add_heading("ACKNOWLEDGEMENT",1); doc.add_paragraph("[Write a truthful acknowledgement.]"); doc.add_page_break()
        doc.add_heading("EXECUTIVE SUMMARY",1); add_text(doc,data["overview"],"one-page purpose, work done and learning outcomes"); doc.add_page_break()
        doc.add_heading("TABLE OF CONTENTS",1); doc.add_paragraph("[In Word or LibreOffice, insert/update an automatic table of contents after completing the document.]"); doc.add_page_break()
    for i,(heading,guide) in enumerate(profile["chapters"]):
        doc.add_heading(heading,1)
        if "INTRODUCTION" in heading or "PROBLEM" in heading:
            add_text(doc,data["overview"],guide)
        elif any(k in heading for k in ["WORK","IMPLEMENTATION","METHODOLOGY","SYSTEM ANALYSIS","DESCRIPTION OF WORK"]):
            add_text(doc,data["work"],guide)
        elif any(k in heading for k in ["RESULT","TESTING","ANALYSIS","EVALUATION"]):
            add_text(doc,data["results"],guide)
        elif "CONCLUSION" in heading:
            add_text(doc,data["conclusion"],guide)
        elif "REFERENCES" in heading:
            add_text(doc,data["references"],guide)
        else:
            doc.add_paragraph("[To complete: "+guide+"]")
    doc.add_heading("FINAL REVIEW CHECKLIST",1)
    checks=["Replace every placeholder with accurate content.","Insert authentic certificates and required evidence.","Verify all citations and reference details.","Check figure/table captions and references in text.","Update the table of contents and verify page numbering.","For internship reports, supplied JUET guidelines state a minimum of 40 pages after the index."]
    for c in checks: doc.add_paragraph(c,style="List Bullet")
    out=BytesIO(); doc.save(out); out.seek(0); return out

def add_slide(prs,title,body,number,total):
    slide=prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb=RGBColor(12,18,31)
    # top accent bar
    shape=slide.shapes.add_shape(1,0,0,prs.slide_width,PptInches(.12))
    shape.fill.solid(); shape.fill.fore_color.rgb=RGBColor(110,157,255); shape.line.fill.background()
    tx=slide.shapes.add_textbox(PptInches(.7),PptInches(.55),PptInches(12),PptInches(1.1)).text_frame
    tx.word_wrap=True; p=tx.paragraphs[0]; p.text=title; p.font.name="Aptos Display"; p.font.size=PptPt(28); p.font.bold=True; p.font.color.rgb=RGBColor(239,244,255)
    bodybox=slide.shapes.add_textbox(PptInches(.85),PptInches(2),PptInches(11.5),PptInches(4.4)).text_frame
    bodybox.word_wrap=True
    lines=[x.strip() for x in (body or "").splitlines() if x.strip()]
    if not lines: lines=["[Add verified project-specific content.]"]
    for i,line in enumerate(lines[:6]):
        p=bodybox.paragraphs[0] if i==0 else bodybox.add_paragraph()
        p.text=("• " if len(lines)>1 else "")+line
        p.font.name="Aptos"; p.font.size=PptPt(21 if len(lines)<=3 else 18); p.font.color.rgb=RGBColor(211,222,241); p.space_after=PptPt(18)
    footer=slide.shapes.add_textbox(PptInches(.7),PptInches(7.05),PptInches(12),PptInches(.25)).text_frame.paragraphs[0]
    footer.text=f"Slow-Steady  ·  {number:02d} / {total:02d}"; footer.font.size=PptPt(10); footer.font.color.rgb=RGBColor(139,157,185)

def build_ppt(data):
    prs=Presentation(); prs.slide_width=PptInches(13.333); prs.slide_height=PptInches(7.5)
    contents={
      "Introduction":[data["overview"] or "Project context, motivation and objectives."],
      "Problem Statement":[data["overview"] or "Define the problem, users, scope and constraints."],
      "Literature Review":["Summarize relevant prior work with verified citations.","State the gap your project addresses."],
      "Work Plan":[data["contribution"] or "Add weekly-report-aligned milestones and dates."],
      "Individual Contribution":[data["contribution"] or "State each team member’s actual contribution."],
      "Work Done So Far":[data["work"] or "Add evidence of work completed.",data["results"] or "Add verified tests, experiments or analysis.","Add implementation/design evidence; research projects should show methodology and analysis."],
      "Work to be Done":["Add remaining tasks, milestones, owners and risks."],
      "Conclusion":[data["conclusion"] or "Summarize only conclusions supported by current evidence."],
      "References":[data["references"] or "Add verified references in IEEE format."]
    }
    sections=[]
    for title,count,desc in PPT_SECTIONS:
        bodies=contents[title]
        for n in range(count):
            body=bodies[n] if n<len(bodies) else desc
            if title=="Work Done So Far" and count==4 and n>=len(bodies): body="Add a distinct evidence-backed detail: design, implementation, testing or analysis."
            sections.append((title if count==1 else f"{title} ({n+1})",body))
    total=len(sections)
    for i,(title,body) in enumerate(sections,1): add_slide(prs,data["project_title"] if i==1 else title,body,i,total)
    out=BytesIO(); prs.save(out); out.seek(0); return out

@app.get("/",response_class=HTMLResponse)
def home(): return HTMLResponse(PAGE)

@app.get("/health")
def health(): return {"status":"ok","service":"slow-steady","version":"0.2.0","profiles":list(REPORT_PROFILES),"pptx":True}

@app.get("/api/profiles")
def profiles():
    return {"reports":{k:{"label":v["label"],"chapters":[x[0] for x in v["chapters"]]} for k,v in REPORT_PROFILES.items()},
            "major_ppt":[{"section":t,"slides":n,"guidance":g} for t,n,g in PPT_SECTIONS],
            "note":"Minor/Major report headings are initial starter profiles and must be checked against the relevant official guideline."}

@app.post("/generate")
def generate(document_type:str=Form(...),project_title:str=Form(...),student_name:str=Form(...),enrollment:str=Form(""),organization:str=Form(""),guide:str=Form(""),department:str=Form("Computer Science and Engineering"),year:str=Form(str(date.today().year)),overview:str=Form(""),work:str=Form(""),results:str=Form(""),contribution:str=Form(""),conclusion:str=Form(""),references:str=Form("")):
    data={k:v.strip() if isinstance(v,str) else v for k,v in locals().copy().items() if k not in {"document_type"}}
    data["document_type"]=document_type
    if not project_title.strip() or not student_name.strip(): return JSONResponse({"error":"Project title and student/presenter are required."},status_code=422)
    if document_type=="major_ppt":
        out=build_ppt(data); filename="Slow-Steady-Major-Project-Presentation.pptx"; media="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    elif document_type in REPORT_PROFILES:
        out=build_report(data); filename=f'Slow-Steady-{document_type}-Report.docx'; media="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    else: return JSONResponse({"error":"Unknown document profile."},status_code=422)
    return StreamingResponse(out,media_type=media,headers={"Content-Disposition":f'attachment; filename="{filename}"'})

@app.post("/api/validate")
def validate(document_type:str=Form(...),project_title:str=Form(""),student_name:str=Form(""),overview:str=Form(""),work:str=Form(""),results:str=Form(""),references:str=Form("")):
    issues=[]
    for key,label in [("project_title","Project title"),("student_name","Student/presenter name")]:
        if not locals()[key].strip(): issues.append({"severity":"error","rule":f"required_{key}","message":f"{label} is missing."})
    if not overview.strip(): issues.append({"severity":"warning","rule":"overview_present","message":"Project overview/problem statement is empty."})
    if not work.strip(): issues.append({"severity":"warning","rule":"work_present","message":"Work completed/methodology is empty."})
    if not results.strip(): issues.append({"severity":"warning","rule":"results_present","message":"Results/analysis is empty."})
    if not references.strip(): issues.append({"severity":"warning","rule":"references_present","message":"No references supplied; add verified citations where used."})
    if document_type=="major_ppt":
        return {"profile":document_type,"checks":{"required_sections":len(PPT_SECTIONS),"expected_slide_count":sum(x[1] for x in PPT_SECTIONS),"ai_generated_images":False},"issues":issues,"passed":not any(i["severity"]=="error" for i in issues)}
    return {"profile":document_type,"checks":{"required_heading_profile":document_type in REPORT_PROFILES,"a4_page_setup":True,"times_new_roman_12pt_body":True,"line_spacing_1_15":True},"issues":issues,"passed":not any(i["severity"]=="error" for i in issues)}
