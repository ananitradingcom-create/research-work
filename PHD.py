# -*- coding: utf-8 -*-
"""
Advanced Open-Source Academic Research Article Generation Pipeline
Google Colab T4 GPU -> Hugging Face (Qwen2.5-7B-Instruct) -> Word Document -> Telegram

Features:
1. No API Key required (Runs locally on Colab Free GPU).
2. Generates rigorous PhD-level international standard academic manuscripts.
3. Includes 45+ academic references (APA 7th format).
4. Detailed analytical markdown tables and figure placeholders.
5. Sends formatted Word document directly to Telegram.
"""

import os
import re
import sys
import argparse
import traceback
import torch

from pathlib import Path
from datetime import datetime
from transformers import AutoModelForCausalLM, AutoTokenizer, pipeline

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
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

except ImportError as exc:
    safe_print(f"Missing library: {exc}")
    safe_print(
        "Run: !pip install -U transformers torch accelerate requests python-docx"
    )
    raise


# =========================================================
# 3. CONFIGURATION & TELEGRAM CREDENTIALS
# =========================================================

# ಅಂತರರಾಷ್ಟ್ರೀಯ ಪಿಎಚ್‌ಡಿ ಸ್ಟ್ಯಾಂಡರ್ಡ್‌ಗಾಗಿ ಅತ್ಯುತ್ತಮ ಓಪನ್ ಸೋರ್ಸ್ ಮಾಡೆಲ್
DEFAULT_MODEL = "Qwen/Qwen2.5-7B-Instruct"

# ನಿಮ್ಮ ಟೆಲಿಗ್ರಾಮ್ ಬಾಟ್ ಟೋಕನ್ ಮತ್ತು ಚಾಟ್ ಐಡಿ
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
# 4. ADVANCED PhD-LEVEL RESEARCH PROMPT
# =========================================================

def build_prompt(topic, instructions, pages):

    return f"""
You are a distinguished Professor of Management and Senior Research Scientist publishing in ABDC Grade-A* and Scopus Q1 journals. 

Prepare an exceptionally rigorous, exhaustive, and academically elite research article DRAFT.

RESEARCH TOPIC:
{topic}

ADDITIONAL INSTRUCTIONS:
{instructions or "Employ advanced empirical methodologies, robust theoretical framing, and comprehensive statistical frameworks."}

TARGET LENGTH & DEPTH:
Equivalent to approximately {pages} comprehensive academic pages with dense, profound scholarly discourse.

MANDATORY STRUCTURAL SECTIONS:
1. TITLE (Academic & Impactful)
2. ABSTRACT (Structured: Background, Purpose, Methodology, Findings, Implications - 300 words)
3. KEYWORDS (5-6 advanced scholarly keywords)
4. INTRODUCTION & BACKGROUND
5. STATEMENT OF THE RESEARCH PROBLEM
6. THEORETICAL GAPS & INTELLECTUAL JUSTIFICATION
7. RESEARCH OBJECTIVES & RESEARCH QUESTIONS
8. THEORETICAL FOUNDATION & PARADIGM
9. EXHAUSTIVE LITERATURE REVIEW (Synthesizing global scholarly debates)
10. CONCEPTUAL FRAMEWORK & THEORETICAL MODEL
11. HYPOTHESES DEVELOPMENT (H1 to H6 with theoretical derivations)
12. RESEARCH METHODOLOGY (Epistemology, sampling, psychometric validation, econometric models)
13. VARIABLES AND OPERATIONAL DEFINITIONS (Detailed Markdown Table)
14. COMPREHENSIVE DATA ANALYSIS PLAN & EXPECTED STATISTICAL OUTPUT
15. DETAILED EMPIRICAL RESULTS / SIMULATED FINDINGS (Including Regression & Structural Equation Modeling Tables)
16. DISCUSSION OF FINDINGS (Connecting with prior global literature)
17. THEORETICAL & EPISTEMOLOGICAL IMPLICATIONS
18. MANAGERIAL, STRATEGIC & POLICY IMPLICATIONS
19. METHODOLOGICAL LIMITATIONS & FUTURE RESEARCH AVENUES
20. CONCLUSION
21. REFERENCES (MANDATORY: Provide AT LEAST 45 to 50 rigorous, authentic scholarly journal references in APA 7th edition format spanning from 2010 to 2026).

ACADEMIC RIGOR & FORMATTING RULES:
- Use sophisticated international academic English vocabulary (e.g., ontological positioning, nomological network, empirical triangulation, hermeneutic synthesis).
- Include at least 4 detailed Markdown tables with clear column headings and borders.
- Include specific graph/diagram placeholders using format: [FIGURE 1: Conceptual Structural Model].
- Do not summarize or truncate; write the complete, full-length comprehensive manuscript.

Output the complete academic manuscript.
"""


# =========================================================
# 5. OPEN-SOURCE HUGGING FACE MANUSCRIPT GENERATION
# =========================================================

def generate_manuscript(topic, instructions, pages):

    safe_print(f"Loading open-source model ({DEFAULT_MODEL}) on GPU...")

    try:
        tokenizer = AutoTokenizer.from_pretrained(DEFAULT_MODEL)
        
        # T4 GPU ಗಾಗಿ bfloat16 ಅಥವಾ 4-bit ಲೋಡಿಂಗ್
        model = AutoModelForCausalLM.from_pretrained(
            DEFAULT_MODEL,
            torch_dtype=torch.bfloat16,
            device_map="auto"
        )
        
        pipe = pipeline(
            "text-generation",
            model=model,
            tokenizer=tokenizer,
            max_new_tokens=8192,
            temperature=0.7,
            top_p=0.9,
            repetition_penalty=1.1
        )

    except Exception as exc:
        raise RuntimeError(
            f"Failed to load Hugging Face model: {exc}"
        ) from exc

    prompt = build_prompt(topic, instructions, pages)

    safe_print("Generating comprehensive PhD-level manuscript draft...")

    try:
        messages = [
            {"role": "system", "content": "You are an elite academic research writer producing Scopus Q1 level manuscripts with 45+ references."},
            {"role": "user", "content": prompt}
        ]
        
        formatted_prompt = tokenizer.apply_chat_template(
            messages, 
            tokenize=False, 
            add_generation_prompt=True
        )

        outputs = pipe(
            formatted_prompt,
            max_new_tokens=8192,
            do_sample=True,
            temperature=0.7,
            pad_token_id=tokenizer.eos_token_id
        )

        generated_text = outputs[0]["generated_text"]
        
        # ಅಲ್ಗಾರಿದಮ್ ರೆಸ್ಪಾನ್ಸ್ ಮಾತ್ರ ಹೊರತೆಗೆಯುವುದು
        if "<|im_start|>assistant" in generated_text:
            manuscript = generated_text.split("<|im_start|>assistant")[-1].strip()
        else:
            manuscript = generated_text[len(formatted_prompt):].strip()

    except Exception as exc:
        raise RuntimeError(
            f"Model generation failed: {exc}"
        ) from exc

    if not manuscript:
        raise RuntimeError("The model returned empty text. Try again.")

    return manuscript


# =========================================================
# 6. WORD TABLE FORMATTING
# =========================================================

def shade_cell(cell, fill="1F4E78"):
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)


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
        element.set(qn("w:sz"), "4")
        element.set(qn("w:color"), "B7C9D6")


def add_markdown_table(doc, lines):
    rows = []
    for line in lines:
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        if all(re.fullmatch(r":?-{3,}:?", cell.replace(" ", "")) for cell in cells):
            continue
        rows.append(cells)

    if not rows:
        return

    number_of_columns = max(len(row) for row in rows)
    table = doc.add_table(rows=0, cols=number_of_columns)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for row_index, row in enumerate(rows):
        cells = table.add_row().cells
        for column_index in range(number_of_columns):
            value = row[column_index] if column_index < len(row) else ""
            cells[column_index].text = value.replace("**", "")
            set_cell_border(cells[column_index])

            for paragraph in cells[column_index].paragraphs:
                paragraph.paragraph_format.space_after = Pt(2)
                for run in paragraph.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(8.5)

            if row_index == 0:
                shade_cell(cells[column_index])
                for paragraph in cells[column_index].paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)

    doc.add_paragraph()


# =========================================================
# 7. BOLD TEXT SUPPORT
# =========================================================

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


# =========================================================
# 8. MARKDOWN TO WORD CONVERSION
# =========================================================

def markdown_to_docx(markdown_text, topic, output_path):
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

    # Title
    title_paragraph = doc.add_paragraph()
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_paragraph.add_run(topic.strip())
    title_run.bold = True
    title_run.font.name = "Times New Roman"
    title_run.font.size = Pt(16)
    title_run.font.color.rgb = RGBColor(31, 78, 121)

    # Document date
    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    note_run = note.add_run("International PhD Research Manuscript | Generated " + datetime.now().strftime("%d %B %Y"))
    note_run.italic = True
    note_run.font.size = Pt(9)

    lines = markdown_text.splitlines()
    index = 0

    while index < len(lines):
        line = lines[index].strip()
        if not line:
            index += 1
            continue

        if line.startswith("|") and "|" in line[1:]:
            table_lines = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                table_lines.append(lines[index].strip())
                index += 1
            add_markdown_table(doc, table_lines)
            continue

        if line.startswith("# "):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=1)
        elif line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=2)
        elif re.match(r"^\d+[\.\)]\s+", line):
            paragraph = doc.add_paragraph(style="List Number")
            add_text_with_bold(paragraph, re.sub(r"^\d+[\.\)]\s+", "", line))
        elif line.startswith(("- ", "* ")):
            paragraph = doc.add_paragraph(style="List Bullet")
            add_text_with_bold(paragraph, line[2:])
        elif "[FIGURE" in line or "[GRAPH" in line:
            paragraph = doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.add_run(line)
            run.italic = True
            run.font.color.rgb = RGBColor(89, 89, 89)
        else:
            paragraph = doc.add_paragraph()
            add_text_with_bold(paragraph, line)

        index += 1

    doc.save(output_path)


# =========================================================
# 9. TELEGRAM DOCUMENT DELIVERY
# =========================================================

def send_document_to_telegram(file_path, caption):
    token = TELEGRAM_BOT_TOKEN_CONFIG
    chat_id = TELEGRAM_CHAT_ID_CONFIG

    if not token or not chat_id:
        raise RuntimeError("Telegram credentials are missing.")

    url = f"https://api.telegram.org/bot{token}/sendDocument"

    with file_path.open("rb") as file_obj:
        response = requests.post(
            url,
            data={"chat_id": chat_id, "caption": caption[:900]},
            files={"document": (file_path.name, file_obj, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
            timeout=120
        )

    if not response.ok or not response.json().get("ok"):
        raise RuntimeError(f"Telegram upload failed: {response.text}")

    safe_print("Word document sent successfully to Telegram.")


# =========================================================
# 10. MAIN PIPELINE
# =========================================================

def main():
    parser = argparse.ArgumentParser(description="Open-Source PhD Research Pipeline")
    parser.add_argument("--topic", default=os.getenv("RESEARCH_TOPIC", ""), help="Research topic")
    parser.add_argument("--instructions", default="", help="Instructions")
    parser.add_argument("--pages", type=int, default=15, help="Target depth")

    args, _ = parser.parse_known_args()
    topic = args.topic.strip()

    if not topic:
        raise RuntimeError("Research topic is required.")

    manuscript = generate_manuscript(topic, args.instructions, args.pages)

    safe_name = re.sub(r"[^A-Za-z0-9]+", "_", topic).strip("_")[:70] or "research_article"
    output_path = OUTPUT_DIR / f"{safe_name}_PhD_Manuscript.docx"
    markdown_path = OUTPUT_DIR / f"{safe_name}_manuscript.md"

    markdown_path.write_text(manuscript, encoding="utf-8")
    markdown_to_docx(manuscript, topic, output_path)

    safe_print(f"Word document created: {output_path.resolve()}")

    send_document_to_telegram(
        output_path,
        f"PhD International Manuscript: {topic}\nReferences (45+) & Tables included."
    )

    safe_print("Pipeline completed successfully.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        safe_print(f"\nERROR: {exc}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)
