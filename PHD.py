# -*- coding: utf-8 -*-
"""
Research Article Generation Pipeline
Google Colab -> OpenAI API -> Word Document -> Telegram

Features:
1. Generate an academic research manuscript.
2. Create a formatted Word document.
3. Save manuscript as Markdown.
4. Send Word document to Telegram.
5. Handle UTF-8 encoding errors.
6. Avoid fabricated research data and references.
"""

import os
import re
import sys
import argparse
import getpass
import traceback

from pathlib import Path
from datetime import datetime

# =========================================================
# 1. UTF-8 ENCODING FIX
# =========================================================

os.environ["PYTHONUTF8"] = "1"
os.environ["PYTHONIOENCODING"] = "utf-8"

for stream in (sys.stdout, sys.stderr):
    if hasattr(stream, "reconfigure"):
        try:
            stream.reconfigure(
                encoding="utf-8",
                errors="backslashreplace"
            )
        except Exception:
            pass


def safe_print(message, file=None):
    """Print messages without crashing on encoding errors."""

    stream = file or sys.stdout

    try:
        print(str(message), file=stream)
    except (UnicodeEncodeError, UnicodeError):
        try:
            print(
                str(message).encode(
                    "ascii",
                    errors="backslashreplace"
                ).decode("ascii"),
                file=stream
            )
        except Exception:
            pass


# =========================================================
# 2. IMPORT REQUIRED LIBRARIES
# =========================================================

try:
    import requests
    from openai import OpenAI

    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

except ImportError as exc:
    safe_print(f"Missing library: {exc}")
    safe_print(
        "Run: !pip install -U openai requests python-docx"
    )
    raise


# =========================================================
# 3. CONFIGURATION & TELEGRAM CREDENTIALS
# =========================================================

DEFAULT_MODEL = os.getenv(
    "OPENAI_MODEL",
    "gpt-6-astra"
)

# ನಿಮ್ಮ ಟೆಲಿಗ್ರಾಮ್ ಬಾಟ್ ಟೋಕನ್ ಮತ್ತು ಚಾಟ್ ಐಡಿ ಇಲ್ಲಿ ಅಳವಡಿಸಲಾಗಿದೆ
TELEGRAM_BOT_TOKEN_CONFIG = "8837048625:AAFaq30r4xWFAKbUif-fXSjjuAZ54N6PQeU"
TELEGRAM_CHAT_ID_CONFIG = "1262412842"

OUTPUT_DIR = Path(
    os.getenv("OUTPUT_DIR", "/content/research-output")
)

OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)


# =========================================================
# 4. SECURE CREDENTIAL INPUT
# =========================================================

def get_secret(env_name, prompt_text):

    value = os.getenv(env_name, "").strip()

    if not value:
        try:
            value = getpass.getpass(
                prompt_text
            ).strip()
        except Exception:
            value = input(
                prompt_text
            ).strip()

    if not value:
        raise RuntimeError(
            f"{env_name} is missing. "
            "Enter the required credential."
        )

    os.environ[env_name] = value

    return value


# =========================================================
# 5. RESEARCH ARTICLE PROMPT
# =========================================================

def build_prompt(topic, instructions, pages):

    return f"""
Prepare a rigorous academic research article DRAFT.

RESEARCH TOPIC:
{topic}

ADDITIONAL INSTRUCTIONS:
{instructions or "Use a standard academic research article structure."}

TARGET LENGTH:
Approximately {pages} Word pages, depending on formatting.

REQUIRED STRUCTURE:

1. TITLE
2. ABSTRACT
3. KEYWORDS
4. INTRODUCTION
5. RESEARCH PROBLEM
6. RESEARCH GAP
7. RESEARCH OBJECTIVES
8. RESEARCH QUESTIONS
9. THEORETICAL FOUNDATION
10. LITERATURE REVIEW
11. CONCEPTUAL FRAMEWORK
12. HYPOTHESES, WHERE APPROPRIATE
13. RESEARCH METHODOLOGY
14. VARIABLES AND OPERATIONAL DEFINITIONS
15. DATA ANALYSIS PLAN
16. RESULTS OR EXPECTED ANALYSIS
17. DISCUSSION
18. THEORETICAL IMPLICATIONS
19. MANAGERIAL OR POLICY IMPLICATIONS
20. LIMITATIONS AND FUTURE RESEARCH
21. CONCLUSION
22. REFERENCES

ACADEMIC INTEGRITY:

- Never invent survey responses or observations.
- Never fabricate sample sizes or statistical results.
- Never invent coefficients, p-values or hypothesis-testing results.
- Never fabricate journal articles, authors, citations or DOI links.
- Clearly distinguish actual findings from proposed analysis.
- If no dataset is supplied, state that actual empirical results
  cannot be reported.
- Provide a proposed statistical analysis plan when appropriate.
- Mark uncertain references as
  [REFERENCE VERIFICATION REQUIRED].
- Do not guarantee Scopus or ABDC acceptance.
- Use formal academic English.
- Avoid repetitive generic content.
- Use Markdown tables with headings and separator rows.
- Identify proposed diagrams using [FIGURE: ...].
- Include a submission-verification checklist.

Write a comprehensive, logically structured academic draft.

Output only the manuscript.
"""


# =========================================================
# 6. OPENAI API MANUSCRIPT GENERATION
# =========================================================

def generate_manuscript(topic, instructions, pages):

    api_key = get_secret(
        "OPENAI_API_KEY",
        "Enter OpenAI API key (input hidden): "
    )

    client = OpenAI(
        api_key=api_key
    )

    prompt = build_prompt(
        topic,
        instructions,
        pages
    )

    safe_print(
        "Generating manuscript draft using OpenAI API..."
    )

    try:

        response = client.responses.create(
            model=DEFAULT_MODEL,

            instructions=(
                "You are an academic research-writing assistant. "
                "Prioritize research integrity, traceable claims, "
                "transparent methods and explicit uncertainty. "
                "Never fabricate data or references."
            ),

            input=prompt,

            max_output_tokens=12000
        )

    except Exception as exc:

        raise RuntimeError(
            "OpenAI manuscript generation failed. "
            "Check your API key, model access, billing, "
            "internet connection and API error details. "
            f"Original error: {exc}"
        ) from exc

    manuscript = (
        response.output_text or ""
    ).strip()

    if not manuscript:

        raise RuntimeError(
            "The API returned no manuscript text. "
            "Try again or shorten the request."
        )

    return manuscript


# =========================================================
# 7. WORD TABLE FORMATTING
# =========================================================

def shade_cell(cell, fill="1F4E78"):

    tc_pr = cell._tc.get_or_add_tcPr()

    shading = OxmlElement("w:shd")

    shading.set(
        qn("w:fill"),
        fill
    )

    tc_pr.append(shading)


def set_cell_border(cell):

    tc = cell._tc

    tc_pr = tc.get_or_add_tcPr()

    borders = tc_pr.first_child_found_in(
        "w:tcBorders"
    )

    if borders is None:

        borders = OxmlElement(
            "w:tcBorders"
        )

        tc_pr.append(borders)

    for edge in (
        "top",
        "left",
        "bottom",
        "right",
        "insideH",
        "insideV"
    ):

        tag = "w:" + edge

        element = borders.find(
            qn(tag)
        )

        if element is None:

            element = OxmlElement(
                tag
            )

            borders.append(element)

        element.set(
            qn("w:val"),
            "single"
        )

        element.set(
            qn("w:sz"),
            "4"
        )

        element.set(
            qn("w:color"),
            "B7C9D6"
        )


def add_markdown_table(doc, lines):

    rows = []

    for line in lines:

        cells = [
            cell.strip()
            for cell in line.strip().strip("|").split("|")
        ]

        if all(
            re.fullmatch(
                r":?-{3,}:?",
                cell.replace(" ", "")
            )
            for cell in cells
        ):
            continue

        rows.append(cells)

    if not rows:
        return

    number_of_columns = max(
        len(row) for row in rows
    )

    table = doc.add_table(
        rows=0,
        cols=number_of_columns
    )

    table.style = "Table Grid"

    table.alignment = (
        WD_TABLE_ALIGNMENT.CENTER
    )

    for row_index, row in enumerate(rows):

        cells = table.add_row().cells

        for column_index in range(
            number_of_columns
        ):

            value = (
                row[column_index]
                if column_index < len(row)
                else ""
            )

            cells[column_index].text = (
                value.replace("**", "")
            )

            set_cell_border(
                cells[column_index]
            )

            for paragraph in (
                cells[column_index].paragraphs
            ):

                paragraph.paragraph_format.space_after = Pt(2)

                for run in paragraph.runs:

                    run.font.name = "Times New Roman"
                    run.font.size = Pt(8.5)

            if row_index == 0:

                shade_cell(
                    cells[column_index]
                )

                for paragraph in (
                    cells[column_index].paragraphs
                ):

                    for run in paragraph.runs:

                        run.bold = True

                        run.font.color.rgb = RGBColor(
                            255,
                            255,
                            255
                        )

    doc.add_paragraph()


# =========================================================
# 8. BOLD TEXT SUPPORT
# =========================================================

def add_text_with_bold(paragraph, text):

    parts = re.split(
        r"(\*\*.*?\*\*)",
        text
    )

    for part in parts:

        if not part:
            continue

        if (
            part.startswith("**")
            and part.endswith("**")
        ):

            run = paragraph.add_run(
                part[2:-2]
            )

            run.bold = True

        else:

            paragraph.add_run(part)


# =========================================================
# 9. MARKDOWN TO WORD CONVERSION
# =========================================================

def markdown_to_docx(
    markdown_text,
    topic,
    output_path
):

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

    # Page border

    section_properties = section._sectPr

    page_borders = OxmlElement(
        "w:pgBorders"
    )

    page_borders.set(
        qn("w:offsetFrom"),
        "page"
    )

    for edge in (
        "top",
        "left",
        "bottom",
        "right"
    ):

        element = OxmlElement(
            "w:" + edge
        )

        element.set(
            qn("w:val"),
            "single"
        )

        element.set(
            qn("w:sz"),
            "8"
        )

        element.set(
            qn("w:space"),
            "18"
        )

        element.set(
            qn("w:color"),
            "4472C4"
        )

        page_borders.append(element)

    section_properties.append(
        page_borders
    )

    # Title

    title_paragraph = doc.add_paragraph()

    title_paragraph.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    title_run = title_paragraph.add_run(
        topic.strip()
    )

    title_run.bold = True
    title_run.font.name = "Times New Roman"
    title_run.font.size = Pt(16)

    title_run.font.color.rgb = RGBColor(
        31,
        78,
        121
    )

    # Document date

    note = doc.add_paragraph()

    note.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    note_run = note.add_run(
        "Research Article Draft | Generated "
        + datetime.now().strftime("%d %B %Y")
    )

    note_run.italic = True
    note_run.font.size = Pt(9)

    # Process manuscript lines

    lines = markdown_text.splitlines()

    index = 0

    while index < len(lines):

        line = lines[index].strip()

        if not line:

            index += 1
            continue

        # Markdown tables

        if (
            line.startswith("|")
            and "|" in line[1:]
        ):

            table_lines = []

            while (
                index < len(lines)
                and lines[index].strip().startswith("|")
            ):

                table_lines.append(
                    lines[index].strip()
                )

                index += 1

            add_markdown_table(
                doc,
                table_lines
            )

            continue

        # Headings

        if line.startswith("# "):

            doc.add_heading(
                line[2:].strip(),
                level=1
            )

        elif line.startswith("## "):

            doc.add_heading(
                line[3:].strip(),
                level=1
            )

        elif line.startswith("### "):

            doc.add_heading(
                line[4:].strip(),
                level=2
            )

        # Numbered lists

        elif re.match(
            r"^\d+[\.\)]\s+",
            line
        ):

            paragraph = doc.add_paragraph(
                style="List Number"
            )

            add_text_with_bold(
                paragraph,
                re.sub(
                    r"^\d+[\.\)]\s+",
                    "",
                    line
                )
            )

        # Bullet lists

        elif line.startswith(("- ", "* ")):

            paragraph = doc.add_paragraph(
                style="List Bullet"
            )

            add_text_with_bold(
                paragraph,
                line[2:]
            )

        # Figure placeholders

        elif line.startswith("[FIGURE:"):

            paragraph = doc.add_paragraph()

            paragraph.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER
            )

            run = paragraph.add_run(line)

            run.italic = True

            run.font.color.rgb = RGBColor(
                89,
                89,
                89
            )

        # Normal paragraphs

        else:

            paragraph = doc.add_paragraph()

            add_text_with_bold(
                paragraph,
                line
            )

        index += 1

    # Footer and page numbers

    for sec in doc.sections:

        footer = sec.footer.paragraphs[0]

        footer.alignment = (
            WD_ALIGN_PARAGRAPH.CENTER
        )

        footer.add_run(
            "Research draft - verify references and claims | Page "
        )

        field = OxmlElement(
            "w:fldSimple"
        )

        field.set(
            qn("w:instr"),
            "PAGE"
        )

        footer._p.append(field)

    # Save Word document

    doc.save(output_path)


# =========================================================
# 10. TELEGRAM DOCUMENT DELIVERY
# =========================================================

def send_document_to_telegram(
    file_path,
    caption
):

    if TELEGRAM_BOT_TOKEN_CONFIG:
        token = TELEGRAM_BOT_TOKEN_CONFIG
    else:
        token = get_secret(
            "TELEGRAM_BOT_TOKEN",
            "Enter Telegram bot token (input hidden): "
        )

    if TELEGRAM_CHAT_ID_CONFIG:
        chat_id = TELEGRAM_CHAT_ID_CONFIG
    else:
        chat_id = os.getenv(
            "TELEGRAM_CHAT_ID",
            ""
        ).strip()

    if not chat_id:

        try:
            chat_id = input(
                "Enter Telegram chat ID: "
            ).strip()
        except Exception:
            pass

    if not chat_id:

        raise RuntimeError(
            "TELEGRAM_CHAT_ID is missing."
        )

    url = (
        f"https://api.telegram.org/bot"
        f"{token}/sendDocument"
    )

    with file_path.open("rb") as file_obj:

        response = requests.post(
            url,

            data={
                "chat_id": chat_id,
                "caption": caption[:900]
            },

            files={
                "document": (
                    file_path.name,
                    file_obj,
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                )
            },

            timeout=120
        )

    if not response.ok:

        raise RuntimeError(
            f"Telegram upload failed: HTTP {response.status_code}. "
            "Check bot token, chat ID and permissions."
        )

    payload = response.json()

    if not payload.get("ok"):

        raise RuntimeError(
            "Telegram API rejected the document. "
            "Check bot token and chat ID."
        )

    safe_print(
        "Word document sent successfully to Telegram."
    )


# =========================================================
# 11. MAIN PIPELINE
# =========================================================

def main():

    parser = argparse.ArgumentParser(
        description=(
            "Generate an academic research article "
            "Word document and optionally send it to Telegram."
        )
    )

    parser.add_argument(
        "--topic",
        default=os.getenv("RESEARCH_TOPIC", ""),
        help="Research article topic"
    )

    parser.add_argument(
        "--instructions",
        default=os.getenv("RESEARCH_INSTRUCTIONS", ""),
        help="Additional research instructions"
    )

    parser.add_argument(
        "--pages",
        type=int,
        default=int(
            os.getenv("TARGET_PAGES", "12")
        ),
        help="Approximate target Word pages"
    )

    parser.add_argument(
        "--no-telegram",
        action="store_true",
        help="Create Word document without Telegram delivery"
    )

    # Compatible with Google Colab injected arguments

    args, _unknown = parser.parse_known_args()

    topic = args.topic.strip()

    if not topic:

        try:
            topic = input(
                "Enter your research article topic: "
            ).strip()
        except Exception:
            pass

    if not topic:

        raise RuntimeError(
            "Research topic is required."
        )

    # Generate manuscript

    safe_print(
        "Generating manuscript draft using OpenAI API..."
    )

    manuscript = generate_manuscript(
        topic,
        args.instructions,
        args.pages
    )

    # Create filenames

    safe_name = re.sub(
        r"[^A-Za-z0-9]+",
        "_",
        topic
    ).strip("_")[:70] or "research_article"

    output_path = (
        OUTPUT_DIR
        / f"{safe_name}_Research_Article.docx"
    )

    markdown_path = (
        OUTPUT_DIR
        / f"{safe_name}_manuscript.md"
    )

    # Save UTF-8 Markdown

    markdown_path.write_text(
        manuscript,
        encoding="utf-8"
    )

    # Create Word document

    markdown_to_docx(
        manuscript,
        topic,
        output_path
    )

    safe_print(
        f"Word document created: {output_path.resolve()}"
    )

    safe_print(
        f"Manuscript text saved: {markdown_path.resolve()}"
    )

    # Send to Telegram unless disabled

    if not args.no_telegram:

        send_document_to_telegram(
            output_path,

            (
                f"Research article Word draft: {topic}\n"
                "Please verify references, data and "
                "journal requirements."
            )
        )

    safe_print(
        "Pipeline completed successfully."
    )


# =========================================================
# 12. ERROR HANDLING
# =========================================================

if __name__ == "__main__":

    try:

        main()

    except KeyboardInterrupt:

        safe_print(
            "\nProcess cancelled by user.",
            file=sys.stderr
        )

        sys.exit(1)

    except Exception as exc:

        safe_print(
            f"\nERROR: {exc}",
            file=sys.stderr
        )

        if os.getenv(
            "SHOW_TRACEBACK",
            "1"
        ) == "1":

            traceback.print_exc()

        sys.exit(1)
