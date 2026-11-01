# Slow and Steady — Academic Document Studio

*Submission for the [Hacktoberfest Open-Source AI Challenge — Week 3: Slow and Steady](https://dev.to/challenges/hacktoberfest-week3-2026-10-19).*

## What I Built

**Slow-Steady** is a web-based academic document studio designed to simplify the preparation of project reports and presentations.

Students often spend too much time formatting documents, organizing sections, and repeatedly checking layouts instead of focusing on their actual project work. Slow-Steady aims to make this process more structured and convenient by providing reusable document profiles, a focused editor, a preview, and editable file exports.

The project offers eight document options:

1. Summer Internship Report
2. Minor Project I Report
3. Minor Project II Report
4. Major Project I Report
5. Major Project II Report
6. Major Project Presentation
7. Major & Minor Project PPT
8. Summer Internship Report PPT

Users select a document type from the dashboard and enter their project details in a workspace designed around an editor and document preview.

The goal is simple: **less repetitive formatting, clearer document preparation, and more time for meaningful project work.**

## Demo

**[Open the Slow-Steady application](https://slow-steady.onrender.com/)**

The application is designed to bring document selection, editing, previewing, and editable DOCX/PPTX generation into one workflow.

Useful deployment endpoints:

- **Application:** [slow-steady.onrender.com](https://slow-steady.onrender.com/)
- **Health check:** [ /health](https://slow-steady.onrender.com/health)
- **API documentation:** [ /docs](https://slow-steady.onrender.com/docs)

A live deployment can change or become temporarily unavailable. Check the health endpoint before relying on it for document generation.

## Code

The project is open source and available on GitHub:

**[RISHABH12005/Slow-Steady](https://github.com/RISHABH12005/Slow-Steady)**

You can explore the implementation, review the document-generation workflow, report issues, and contribute improvements. The repository also includes automated tests and compliance checks through GitHub Actions.

## How I Built It

I built Slow-Steady as a Python web application using FastAPI and document-generation libraries for editable Word documents and PowerPoint presentations.

### Technology stack

- **FastAPI:** Handles the web application, API endpoints, and document-generation requests.
- **python-docx:** Generates editable Word reports.
- **python-pptx:** Generates editable PowerPoint presentations.
- **HTML, CSS, and JavaScript:** Provide the dashboard, responsive editor layout, and live preview interactions.
- **pytest and GitHub Actions:** Support automated testing and continuous integration.
- **Render:** Hosts the deployed web application.

### Workflow

1. Select the required report or presentation profile.
2. Enter project information and content in the editor.
3. Review the generated structure in the preview.
4. Generate an editable DOCX or PPTX file.
5. Open the exported file and complete a final review before submission.

The application uses reusable document structures and project details to reduce repeated setup work. Generated documents remain editable so students can refine the content and formatting.

## Document Formatting and Academic Review

Slow-Steady includes document profiles and formatting rules informed by the academic guidelines available for the project. For example, the internship report profile uses A4 page setup, Times New Roman body text, defined margins, heading sizes, and line spacing.

The Major Project presentation generator creates editable, text-based slides based on the supplied presentation structure. It does not insert AI-generated images.

Generated files are starting points, not a guarantee of academic compliance. Students should replace placeholders, verify references, check the latest departmental instructions, and review page numbering, table of contents, layout, and required sections before submission. Some formatting requirements may need final adjustment in Word or PowerPoint.

## Why Open Innovation Matters

Open innovation makes it possible for developers and students to inspect how a tool works, adapt it to their own needs, and contribute improvements rather than depending entirely on a closed service.

For Slow-Steady, an open-source approach means contributors can improve report structures, add document profiles, strengthen formatting validation, and make the tool more useful for different academic requirements.

It also makes development more transparent: users can inspect the source code and automated tests instead of treating the application as a black box.

The project can grow through community contributions, shared feedback, and reusable solutions. A document profile created for one department can become a starting point for another institution, provided its specific academic rules are reviewed and incorporated.

## My Agent Session

I used AI assistance during development to refine the interface, improve the document-selection workflow, troubleshoot automated test failures, and make the preview easier to navigate.

A public agent-session or DevRelay recording link is not included because I do not have a verified shareable session URL.

## The Lazy Payoff

The biggest payoff is reducing repetitive setup work.

Instead of starting from a blank document every time, a student can select a document profile, enter project information, review the preview, and generate an editable file.

I have not measured the time saved in a formal user test, so I will not claim a specific percentage or number of hours. The next step is to test the application with students, measure how long common document-preparation tasks take, and use their feedback to improve the workflow.

**Slow-Steady is built around one idea: spend less time fighting document formatting and more time working on the project itself.**

## Prize Categories

Select partner or prize categories only when the project meets the official challenge eligibility requirements. No unverified partner integration or prize eligibility is claimed here.

---

*Thanks for checking out Slow-Steady. Feedback, bug reports, and contributions are welcome.*
