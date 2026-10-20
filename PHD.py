"""
GitHub -> Google Colab -> Telegram Word-document research pipeline.

What it does:
1. Accepts a research topic and instructions.
2. Uses the OpenAI Responses API to draft a structured academic manuscript.
3. Creates a formatted .docx Word file.
4. Sends the Word file to Telegram using the Telegram Bot API.

Integrity note:
- This creates a research draft, not a guaranteed Scopus/ABDC acceptance.
- Never invent data, statistical results, citations, DOIs, or sample sizes.
- If no dataset/source file is supplied, the model is instructed to mark empirical
  results as unavailable and provide a proposed analysis plan rather than fake results.
"""

import os
import re
import sys
import argparse
from pathlib import Path
from datetime import datetime

import requests
from openai import OpenAI
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

DEFAULT_MODEL = os.getenv("OPENAI_MODEL", "gpt-6-astra")
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "output"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(
            f"Missing {name}. Set it as a Colab secret/environment variable. "
            "Do not paste secret keys into a public GitHub file."
        )
    return value


def build_prompt(topic: str, instructions: str, pages: int) -> str:
    return f"""
Prepare a rigorous academic research-article DRAFT on this topic:
{topic}

Additional user instructions:
{instructions or "Use a standard empirical research article structure."}

Target length: approximately {pages} Word pages, depending on formatting.

Required structure:
TITLE
ABSTRACT
KEYWORDS
1. INTRODUCTION
2. RESEARCH PROBLEM AND RESEARCH GAP
3. OBJECTIVES
4. RESEARCH QUESTIONS
5. THEORETICAL FOUNDATION
6. LITERATURE REVIEW
7. CONCEPTUAL FRAMEWORK
8. HYPOTHESES (only where logically appropriate)
9. RESEARCH METHODOLOGY
10. VARIABLES AND OPERATIONAL DEFINITIONS (use a simple Markdown table)
11. DATA ANALYSIS PLAN
12. RESULTS / EXPECTED ANALYSIS
13. DISCUSSION
14. THEORETICAL IMPLICATIONS
15. MANAGERIAL / POLICY IMPLICATIONS
16. LIMITATIONS AND FUTURE RESEARCH
17. CONCLUSION
REFERENCES

Academic integrity requirements:
- Do NOT invent observations, survey responses, sample sizes, coefficients, p-values,
  statistical test results, citations, journal names, page numbers, or DOI links.
- If no dataset is provided, state clearly that actual results cannot be reported.
  Provide an analysis plan and label any proposed hypotheses as proposed.
- Cite only sources you can identify with reasonable confidence. Mark any citation
  that needs verification as [REFERENCE VERIFICATION REQUIRED].
- Do not claim that the article is accepted, indexed, or guaranteed to meet Scopus/ABDC.
- Distinguish established literature, proposed framework, and untested hypotheses.
- Use formal, clear academic English and avoid repetitive generic statements.
- Tables should be in Markdown format, with a header row and separator row.
- Include figure descriptions in a clear form such as [FIGURE: ...]. Do not pretend
  that a figure is based on real data unless actual data were supplied.
- End with a short checklist of items the author must verify before submission.

Output only the manuscript draft. Do not include explanations about your own process.
"""


def generate_manuscript(topic: str, instructions: str, pages: int) -> str:
    api_key = required_env("OPENAI_API_KEY")
    client = OpenAI(api_key=api_key)
    prompt = build_prompt(topic, instructions, pages)
    response = client.responses.create(
        model=DEFAULT_MODEL,
        instructions=(
            "You are an academic research-writing assistant. Prioritize research integrity, "
            "traceable claims, transparent methods, and explicit uncertainty. Never fabricate data."
        ),
        input=prompt,
        max_output_tokens=12000,
    )
    text = (response.output_text or "").strip()
    if not text:
        raise RuntimeError("The API returned no manuscript text. Try again or shorten the request.")
    return text


def shade_cell(cell, fill="1F4E78"):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_border(cell):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("sz"), "4")
        element.set(qn("color"), "B7C9D6")


def add_markdown_table(doc, lines):
    rows = []
    for line in lines:
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if all(re.fullmatch(r":?-{3,}:?", c.replace(" ", "")) for c in cells):
            continue
        rows.append(cells)
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=0, cols=ncols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        cells = table.add_row().cells
        for j in range(ncols):
            value = row[j] if j < len(row) else ""
            cells[j].text = value.replace("**", "")
            set_cell_border(cells[j])
            for p in cells[j].paragraphs:
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(8.5)
            if i == 0:
                shade_cell(cells[j])
                for p in cells[j].paragraphs:
                    for run in p.runs:
                        run.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
    doc.add_paragraph()


def add_text_with_bold(paragraph, text):
    parts = re.split(r"(\*\*.*?\*\*)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            paragraph.add_run(part)


def markdown_to_docx(markdown_text: str, topic: str, output_path: Path):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15

    # Simple page border.
    sectPr = section._sectPr
    pgBorders = OxmlElement("w:pgBorders")
    pgBorders.set(qn("w:offsetFrom"), "page")
    for edge in ("top", "left", "bottom", "right"):
        tag = "w:" + edge
        element = OxmlElement(tag)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "8")
        element.set(qn("w:space"), "18")
        element.set(qn("w:color"), "4472C4")
        pgBorders.append(element)
    sectPr.append(pgBorders)

    # Title page header.
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(topic.strip())
    title_run.bold = True
    title_run.font.name = "Times New Roman"
    title_run.font.size = Pt(16)
    title_run.font.color.rgb = RGBColor(31, 78, 121)

    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = note.add_run("Research Article Draft | Generated " + datetime.now().strftime("%d %B %Y"))
    r.italic = True
    r.font.size = Pt(9)

    lines = markdown_text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if line.startswith("|") and "|" in line[1:]:
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            add_markdown_table(doc, table_lines)
            continue
        if line.startswith("# "):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=1)
        elif line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=2)
        elif re.match(r"^\d+[\.\)]\s+", line):
            p = doc.add_paragraph(style="List Number")
            add_text_with_bold(p, re.sub(r"^\d+[\.\)]\s+", "", line))
        elif line.startswith(("- ", "* ")):
            p = doc.add_paragraph(style="List Bullet")
            add_text_with_bold(p, line[2:])
        elif line.startswith("[FIGURE:"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(line)
            run.italic = True
            run.font.color.rgb = RGBColor(89, 89, 89)
        else:
            p = doc.add_paragraph()
            add_text_with_bold(p, line)
        i += 1

    # Footer with page field.
    for sec in doc.sections:
        footer = sec.footer.paragraphs[0]
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.add_run("Research draft — verify references and all claims | Page ")
        fld = OxmlElement("w:fldSimple")
        fld.set(qn("w:instr"), "PAGE")
        footer._p.append(fld)

    doc.save(output_path)


def send_document_to_telegram(file_path: Path, caption: str):
    token = required_env("TELEGRAM_BOT_TOKEN")
    chat_id = required_env("TELEGRAM_CHAT_ID")
    url = f"https://api.telegram.org/bot{token}/sendDocument"
    with file_path.open("rb") as file_obj:
        response = requests.post(
            url,
            data={"chat_id": chat_id, "caption": caption[:900]},
            files={
                "document": (
                    file_path.name,
                    file_obj,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            },
            timeout=120,
        )
    if not response.ok:
        raise RuntimeError(f"Telegram sendDocument failed: {response.status_code} {response.text}")
    payload = response.json()
    if not payload.get("ok"):
        raise RuntimeError(f"Telegram API error: {payload}")
    print("Word document sent successfully to Telegram.")


def main():
    parser = argparse.ArgumentParser(description="Generate a research article Word document and send it to Telegram.")
    parser.add_argument("--topic", default=os.getenv("RESEARCH_TOPIC", ""), help="Research article topic")
    parser.add_argument("--instructions", default=os.getenv("RESEARCH_INSTRUCTIONS", ""), help="Additional instructions")
    parser.add_argument("--pages", type=int, default=int(os.getenv("TARGET_PAGES", "12")), help="Approximate target Word pages")
    parser.add_argument("--no-telegram", action="store_true", help="Create the Word document without sending to Telegram")
    args = parser.parse_args()

    topic = args.topic.strip()
    if not topic:
        topic = input("Enter your research article topic: ").strip()
    if not topic:
        raise RuntimeError("A research topic is required.")

    print("Generating manuscript draft using OpenAI API...")
    manuscript = generate_manuscript(topic, args.instructions, args.pages)

    safe_name = re.sub(r"[^A-Za-z0-9]+", "_", topic).strip("_")[:70] or "research_article"
    output_path = OUTPUT_DIR / f"{safe_name}_Research_Article.docx"
    markdown_path = OUTPUT_DIR / f"{safe_name}_manuscript.md"
    markdown_path.write_text(manuscript, encoding="utf-8")
    markdown_to_docx(manuscript, topic, output_path)

    print(f"Word document created: {output_path.resolve()}")
    print(f"Editable manuscript text saved: {markdown_path.resolve()}")

    if not args.no_telegram:
        send_document_to_telegram(
            output_path,
            f"Research article Word draft: {topic}\nPlease verify references, data, and journal requirements."
        )

    print("Pipeline completed.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        sys.exit(1)
