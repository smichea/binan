"""Word version of the history, for editing and review (not for print).

Joins the introduction, the history and the appendices into word/history.docx:
headings, bold and italic text, pipe tables, and the sources as numbered
references at the end. Run from the book/ folder:  python3 convert_md_to_docx.py
"""

import re
from pathlib import Path

from docx import Document

BOOK = Path(__file__).resolve().parent
SOURCES = [BOOK / "raw_files" / name for name in ("introduction.md", "history.md", "appendices.md")]
OUTPUT = BOOK / "word" / "history.docx"

LINK = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
CITATION = re.compile(r"\s*\(\[([^\]]+)\]\(([^)]+)\)\)")
INLINE = re.compile(r"(\*\*[^*]+\*\*|\*[^*]+\*)")


def add_runs(paragraph, text):
    """Write text with **bold** and *italic* runs."""
    for part in INLINE.split(text):
        if part.startswith("**") and part.endswith("**"):
            paragraph.add_run(part[2:-2]).bold = True
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            paragraph.add_run(part[1:-1]).italic = True
        elif part:
            paragraph.add_run(part)


def main():
    references = []

    def to_reference(match):
        references.append(f"{match.group(1)} – {match.group(2)}")
        return f" [{len(references)}]"

    doc = Document()
    for source in SOURCES:
        text = source.read_text(encoding="utf-8")
        text = CITATION.sub(to_reference, text)
        text = LINK.sub(lambda m: m.group(1) + to_reference(m), text)
        for block in text.split("\n\n"):
            block = block.strip()
            if not block:
                continue
            if block.startswith("#"):
                level = len(block) - len(block.lstrip("#"))
                doc.add_heading(block.lstrip("#").strip(), level=min(level, 4))
            elif block.startswith("|"):
                rows = [r for r in block.splitlines() if not re.match(r"^\|\s*:?-", r)]
                cells = [[c.strip() for c in r.strip("|").split("|")] for r in rows]
                table = doc.add_table(rows=len(cells), cols=max(len(r) for r in cells))
                table.style = "Table Grid"
                for i, row in enumerate(cells):
                    for j, value in enumerate(row):
                        cell = table.cell(i, j)
                        cell.text = ""
                        add_runs(cell.paragraphs[0], value)
            else:
                add_runs(doc.add_paragraph(), " ".join(block.splitlines()))
    if references:
        doc.add_page_break()
        doc.add_heading("References", level=1)
        for number, reference in enumerate(references, start=1):
            doc.add_paragraph(f"[{number}] {reference}")
    doc.save(OUTPUT)
    print(f"wrote {OUTPUT.relative_to(BOOK.parent)}")


if __name__ == "__main__":
    main()
