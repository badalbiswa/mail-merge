"""
Tea with Nessie — Python mail-merge scaffold

Complete the TODOs to merge starter_files/tea_with_nessie_invitation.docx
with each row from starter_files/nessie_guests.xlsx.

Suggested reading:
  https://automatetheboringstuff.com/  (Excel & Word chapters)

Usage (from the assignment folder):
  python scaffold_python/mail_merge.py
"""

from __future__ import annotations

import re
import shutil
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

ASSIGNMENT_DIR = Path(__file__).resolve().parent.parent
STARTER = ASSIGNMENT_DIR / "starter_files"
OUTPUT_DIR = Path(__file__).resolve().parent / "output"

EXCEL_PATH = STARTER / "nessie_guests.xlsx"
TEMPLATE_PATH = STARTER / "tea_with_nessie_invitation.docx"

# Matches <<field_name>> including letters, digits, and underscores
FIELD_PATTERN = re.compile(r"<<([a-zA-Z0-9_]+)>>")


def load_guests(path: Path) -> list[dict[str, str]]:
    """Return a list of guest dictionaries from the Guests sheet."""
    wb = load_workbook(path, read_only=True, data_only=True)
    # TODO: select the "Guests" worksheet
    ws = wb["Guests"]

    rows = list(ws.iter_rows(values_only=True))
    # TODO: first row is headers; remaining rows are guest records
    headers = [str(h) for h in rows[0]]
    guests: list[dict[str, str]] = []
    for row in rows[1:]:
        if row is None or all(cell is None for cell in row):
            continue
        record = {
            headers[i]: ("" if row[i] is None else str(row[i]))
            for i in range(len(headers))
        }
        guests.append(record)
    wb.close()
    return guests


def replace_merge_fields(doc: Document, data: dict[str, str]) -> None:
    """
    Replace every <<field>> in the document with values from data.

    Word stores text in runs; a field might be split across runs.
    This scaffold replaces within each paragraph by rebuilding paragraph text.
    """
    for paragraph in doc.paragraphs:
        if "<<" not in paragraph.text:
            continue
        # TODO: build new text by substituting each <<field>> with data[field]
        new_text = paragraph.text
        for match in FIELD_PATTERN.finditer(paragraph.text):
            field = match.group(1)
            value = data.get(field, f"<<MISSING:{field}>>")
            new_text = new_text.replace(f"<<{field}>>", value)

        # Simplest reliable approach for this assignment: clear runs, write one run
        if paragraph.runs:
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]:
                run.text = ""
        else:
            paragraph.add_run(new_text)

    # Also check tables if students extend the template later
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if "<<" not in paragraph.text:
                        continue
                    new_text = paragraph.text
                    for match in FIELD_PATTERN.finditer(paragraph.text):
                        field = match.group(1)
                        value = data.get(field, f"<<MISSING:{field}>>")
                        new_text = new_text.replace(f"<<{field}>>", value)
                    if paragraph.runs:
                        paragraph.runs[0].text = new_text
                        for run in paragraph.runs[1:]:
                            run.text = ""


def merge_one(template_path: Path, guest: dict[str, str], out_path: Path) -> None:
    """Copy the template, fill fields for one guest, save."""
    # TODO: open a fresh copy of the template for each guest
    doc = Document(template_path)
    replace_merge_fields(doc, guest)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)


def main() -> None:
    if not EXCEL_PATH.exists() or not TEMPLATE_PATH.exists():
        raise SystemExit(
            "Starter files missing. Run: python generate_data.py"
        )

    guests = load_guests(EXCEL_PATH)
    print(f"Loaded {len(guests)} guests from {EXCEL_PATH.name}")

    # Clear previous output for a clean run
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    venue_counts: dict[str, int] = {}

    for guest in guests:
        # TODO: choose a clear filename, e.g. invitation_NESS-1001_Fiona_MacLeod.docx
        guest_id = guest.get("guest_id", "unknown")
        first = guest.get("first_name", "guest")
        last = guest.get("last_name", "")
        filename = f"invitation_{guest_id}_{first}_{last}.docx"
        out_path = OUTPUT_DIR / filename
        merge_one(TEMPLATE_PATH, guest, out_path)
        print(f"  wrote {out_path.name}")

        venue = guest.get("venue_name", "Unknown venue")
        venue_counts[venue] = venue_counts.get(venue, 0) + 1

    print(f"Done. Letters are in {OUTPUT_DIR}")

    # Customisation: console summary of guests per venue, proving all three
    # venues are represented and letting the committee sanity-check numbers
    # before the tea urns go on.
    print("\nGuests per venue:")
    for venue, count in sorted(venue_counts.items()):
        print(f"  {venue}: {count}")


if __name__ == "__main__":
    main()
