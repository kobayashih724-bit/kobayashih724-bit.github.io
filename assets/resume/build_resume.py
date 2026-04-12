"""
Generate Haruto_Kobayashi_Resume.txt, .docx, and .pdf from one source.
TXT: stdlib only. DOCX/PDF: pip install python-docx fpdf2
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

DIR = Path(__file__).resolve().parent
OUT_TXT = DIR / "Haruto_Kobayashi_Resume.txt"
OUT_DOCX = DIR / "Haruto_Kobayashi_Resume.docx"
OUT_PDF = DIR / "Haruto_Kobayashi_Resume.pdf"

SIDE_MARGIN = 14
TOP_MARGIN = 12
BOTTOM_SAFE = 20


@dataclass(frozen=True)
class Job:
    company: str
    title: str
    dates: str
    bullets: tuple[str, ...]


SUMMARY = (
    "Full-stack developer with 5 years of professional experience building web applications, "
    "APIs, and data-driven features. Strong in TypeScript, Python, and modern front-end frameworks. "
    "Passionate about reliable systems, clear UX, and shipping maintainable code for real users."
)

SKILLS: tuple[str, ...] = (
    "Languages: JavaScript, TypeScript, Python, Go, PHP",
    "Databases: MySQL, PostgreSQL, MongoDB",
    "Front-end libraries: React, Vue.js, Angular",
    "Frameworks and APIs: FastAPI, Express, Laravel, Next.js, TensorFlow, OpenAI API, LangChain",
    "DevOps and cloud: Git, Docker, Kubernetes, AWS, GCP, JIRA",
    "Soft skills: Problem solving, teamwork, time management, adaptability",
)

JOBS: tuple[Job, ...] = (
    Job(
        "NovaSoft Digital",
        "Full-Stack Developer",
        "January 2025 - October 2025 | Remote, full-time contract",
        (
            "End-to-end web apps: React front end, Node.js/Express and MongoDB back end.",
            "REST APIs and third-party integrations for cross-platform data flow.",
            "CI/CD pipelines for automated testing and deployment (reduced deployment time by 30%).",
            "Stack: React.js, Node.js, Express.js, MongoDB, TypeScript, REST APIs, Git.",
        ),
    ),
    Job(
        "CloudNova Technologies",
        "Front-End Developer",
        "July 2024 - December 2024 | Remote, full-time contract",
        (
            "React and TypeScript components for large single-page applications.",
            "Jest and Cypress testing (reduced production bugs).",
            "Led migration of legacy UI to modern React.",
            "Stack: React.js, Redux, TypeScript, Jest, Cypress, Git, REST APIs.",
        ),
    ),
    Job(
        "BrightWeb Solutions",
        "Front-End Developer",
        "January 2023 - June 2023 | Remote, part-time contract",
        (
            "Reusable React components and responsive layouts (Tailwind CSS, Flexbox).",
            "Performance tuning: code splitting and lazy loading (faster page loads).",
            "Stack: React.js, Tailwind CSS, JavaScript (ES6+), HTML5, CSS3, Git.",
        ),
    ),
)

PROJECTS: tuple[str, ...] = (
    "Coworking Cafe (React, Express) - https://www.coworkingcafe.com/",
    "Mission Control / Lunar MC (Angular, Python) - https://app.lunarmc.ai",
    "Stay AI (Vue, Laravel) - https://stay.ai/",
)

COURSEWORK: tuple[str, ...] = (
    "Data Structures and Algorithms",
    "Database Management Systems",
    "Web Application Development",
    "Machine Learning",
)


def write_txt() -> None:
    sep = "=" * 80
    sub = "-" * 80
    lines: list[str] = [
        sep,
        "HARUTO KOBAYASHI",
        "Software Developer",
        "",
        "kobayashih724@gmail.com",
        "https://kobayashih724-bit.github.io/",
        sep,
        "",
        "PROFESSIONAL SUMMARY",
        sub,
        SUMMARY,
        "",
        "TECHNICAL SKILLS",
        sub,
    ]
    for s in SKILLS:
        lines.append(f"  • {s}")
    lines.extend(["", "EXPERIENCE", sub, ""])
    for job in JOBS:
        lines.append(f"{job.company} — {job.title}")
        lines.append(job.dates)
        lines.append("")
        for b in job.bullets:
            lines.append(f"  • {b}")
        lines.append("")
    lines.extend(["SELECTED PROJECTS", sub])
    for p in PROJECTS:
        lines.append(f"  • {p}")
    lines.extend(
        [
            "",
            "EDUCATION",
            sub,
            "The University of Tokyo",
            "Tokyo, Japan",
            "",
            "Bachelor of Technology in Information and Communication Technology",
            "",
            "Relevant coursework:",
        ]
    )
    for c in COURSEWORK:
        lines.append(f"  • {c}")
    lines.extend(["", sep, ""])
    OUT_TXT.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {OUT_TXT} ({OUT_TXT.stat().st_size} bytes)")


def add_bullets_docx(doc: Any, lines: tuple[str, ...] | list[str]) -> None:
    for line in lines:
        doc.add_paragraph(line, style="List Bullet")


def write_docx() -> None:
    from docx import Document
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.shared import Pt

    doc = Document()

    t = doc.add_paragraph()
    t.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r = t.add_run("Haruto Kobayashi")
    r.bold = True
    r.font.size = Pt(20)

    st = doc.add_paragraph()
    st.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    st.add_run("Software Developer").font.size = Pt(12)

    em = doc.add_paragraph()
    em.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    em.add_run("kobayashih724@gmail.com")

    pf = doc.add_paragraph()
    pf.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    pf.add_run("https://kobayashih724-bit.github.io/")

    doc.add_paragraph()

    doc.add_heading("Professional summary", level=1)
    doc.add_paragraph(SUMMARY)

    doc.add_heading("Technical skills", level=1)
    add_bullets_docx(doc, SKILLS)

    doc.add_heading("Experience", level=1)
    for job in JOBS:
        p = doc.add_paragraph()
        p.add_run(f"{job.company} | {job.title}").bold = True
        wp = doc.add_paragraph(job.dates)
        if wp.runs:
            wp.runs[0].italic = True
        add_bullets_docx(doc, job.bullets)
        doc.add_paragraph()

    doc.add_heading("Selected projects", level=1)
    add_bullets_docx(doc, PROJECTS)

    doc.add_heading("Education", level=1)
    p = doc.add_paragraph()
    p.add_run("The University of Tokyo").bold = True
    doc.add_paragraph("Tokyo, Japan")
    doc.add_paragraph("Bachelor of Technology in Information and Communication Technology")
    rc = doc.add_paragraph()
    rc.add_run("Relevant coursework:").bold = True
    add_bullets_docx(doc, COURSEWORK)

    doc.save(str(OUT_DOCX))
    print(f"Wrote {OUT_DOCX} ({OUT_DOCX.stat().st_size} bytes)")


def write_pdf() -> None:
    from fpdf import FPDF

    class ResumePDF(FPDF):
        def footer(self) -> None:
            self.set_y(-12)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(110, 110, 110)
            self.cell(0, 8, f"Page {self.page_no()}", align="C")

    def w_full(pdf: FPDF) -> float:
        return pdf.w - pdf.l_margin - pdf.r_margin

    def pdf_heading(pdf: ResumePDF, text: str) -> None:
        pdf.ln(3)
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(0, 105, 92)
        pdf.multi_cell(w_full(pdf), 5.5, text)
        pdf.set_text_color(0, 0, 0)

    def pdf_body(pdf: ResumePDF, text: str) -> None:
        pdf.set_font("Helvetica", "", 9)
        pdf.multi_cell(w_full(pdf), 4.2, text)

    def pdf_bullets(pdf: ResumePDF, lines: tuple[str, ...]) -> None:
        pdf.set_font("Helvetica", "", 9)
        for line in lines:
            pdf.multi_cell(w_full(pdf), 4.2, f"- {line}")

    pdf = ResumePDF(orientation="P", unit="mm", format="A4")
    pdf.set_margins(SIDE_MARGIN, TOP_MARGIN, SIDE_MARGIN)
    pdf.set_auto_page_break(auto=True, margin=BOTTOM_SAFE)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 17)
    pdf.set_text_color(0, 77, 64)
    pdf.cell(w_full(pdf), 9, "Haruto Kobayashi", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w_full(pdf), 5, "Software Developer", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(w_full(pdf), 4.5, "kobayashih724@gmail.com", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(w_full(pdf), 4.5, "Portfolio: https://kobayashih724-bit.github.io/", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    pdf_heading(pdf, "Professional summary")
    pdf_body(pdf, SUMMARY)

    pdf_heading(pdf, "Technical skills")
    pdf_bullets(pdf, SKILLS)

    pdf.add_page()

    pdf_heading(pdf, "Experience")
    for job in JOBS:
        pdf.set_font("Helvetica", "B", 10)
        pdf.multi_cell(w_full(pdf), 4.5, f"{job.company} | {job.title}")
        pdf.set_font("Helvetica", "I", 8.5)
        pdf.set_text_color(75, 75, 75)
        pdf.multi_cell(w_full(pdf), 3.8, job.dates)
        pdf.set_text_color(0, 0, 0)
        pdf_bullets(pdf, job.bullets)
        pdf.ln(1.5)

    pdf_heading(pdf, "Selected projects")
    pdf_bullets(pdf, PROJECTS)

    pdf_heading(pdf, "Education")
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.multi_cell(w_full(pdf), 4.5, "The University of Tokyo")
    pdf.set_font("Helvetica", "", 9)
    pdf.multi_cell(w_full(pdf), 4.2, "Tokyo, Japan")
    pdf.multi_cell(w_full(pdf), 4.2, "Bachelor of Technology in Information and Communication Technology")
    pdf.ln(1)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(w_full(pdf), 4, "Relevant coursework:", new_x="LMARGIN", new_y="NEXT")
    pdf_bullets(pdf, COURSEWORK)

    pdf.output(str(OUT_PDF))
    print(f"Wrote {OUT_PDF} ({OUT_PDF.stat().st_size} bytes, {pdf.page_no()} page(s))")


def main() -> None:
    write_txt()
    write_docx()
    write_pdf()


if __name__ == "__main__":
    main()
