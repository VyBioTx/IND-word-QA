#!/usr/bin/env -S pixi run python
"""Create a test docx by modifying the real report as template, embedding all 14 issue types."""

from __future__ import annotations
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from lxml import etree

REAL_REPORT = Path("/tmp/opencode/study_report.docx")
OUTPUT_DIR = Path("/tmp/opencode")
OUTPUT = OUTPUT_DIR / "study_report_modified.docx"


def copy_and_inject(path: Path = OUTPUT) -> Path:
    import shutil
    shutil.copy(str(REAL_REPORT), str(path))

    doc = Document(str(path))

    # ---- TEXT-001: Duplicate word ----
    doc.add_paragraph("The the study demonstrated significant antitumor activity.")
    dup_idx = len(doc.paragraphs) - 1

    # ---- TEXT-002: Repeated spaces ----
    doc.add_paragraph("The  efficacy  was evaluated  on day 21.")
    space_idx = len(doc.paragraphs) - 1

    # ---- NUM-001: Wrong percentage ----
    doc.add_paragraph("5/8 animals (75.0%) showed tumor regression in the high dose group.")
    pct_idx = len(doc.paragraphs) - 1  # 5/8 = 62.5%, not 75.0%

    # ---- TABLE-001: Wrong table total ----
    table = doc.add_table(rows=5, cols=2, style="Table Grid")
    table.cell(0, 0).text = "Group"
    table.cell(0, 1).text = "Count"
    table.cell(1, 0).text = "Vehicle"
    table.cell(1, 1).text = "8"
    table.cell(2, 0).text = "Low dose"
    table.cell(2, 1).text = "8"
    table.cell(3, 0).text = "High dose"
    table.cell(3, 1).text = "8"
    table.cell(4, 0).text = "Total"
    table.cell(4, 1).text = "22"

    # ---- TABLE-002: Empty cell ----
    table2 = doc.add_table(rows=3, cols=2, style="Table Grid")
    table2.cell(0, 0).text = "Parameter"
    table2.cell(0, 1).text = "Value"
    table2.cell(1, 0).text = "Cmax"
    table2.cell(1, 1).text = ""
    table2.cell(2, 0).text = "Tmax"
    table2.cell(2, 1).text = "2h"

    # ---- TABLE-003: N/A inconsistency ----
    table3 = doc.add_table(rows=3, cols=1, style="Table Grid")
    table3.cell(0, 0).text = "Not Applicable"
    table3.cell(1, 0).text = "N/A"
    table3.cell(2, 0).text = "NA"

    # ---- REF-001: Missing table/figure reference ----
    doc.add_paragraph("As shown in Figure 10, the tumor growth was inhibited.")
    doc.add_paragraph("Detailed PK parameters are listed in Table 25.")

    # ---- HEAD-001: Heading numbering gap ----
    doc.add_heading("6.1  Summary of Findings", level=2)
    doc.add_heading("6.3  Future Directions", level=2)

    # ---- WORD-001: Header/footer version mismatch ----
    first_section = doc.sections[0]
    first_section.header.paragraphs[0].text = "Protocol v2.0"
    first_section.footer.paragraphs[0].text = "Confidential - Version 2.1"

    # ---- ABBR-001: Undefined abbreviation ----
    doc.add_paragraph("MTD was determined based on the toxicity profile observed during the study.")
    doc.add_paragraph("SAE were reported and followed up according to the protocol.")

    # ---- ABBR-002: Multiple definitions for same abbreviation ----
    doc.add_paragraph("The maximum tolerated dose (MTD) was established at 30 mg/kg.")
    doc.add_paragraph("Mean tumor diameter (MTD) was also recorded at each time point.")

    # ---- TERM-001: Terminology inconsistency ----
    doc.add_paragraph("PROTOCOL-001 was followed for all procedures.")
    doc.add_paragraph("All deviations from PROTOCOL001 were documented.")

    doc.save(str(path))
    _inject_xml(path)
    return path


def _inject_xml(path: Path) -> None:
    with ZipFile(str(path), "r") as source:
        members = {name: source.read(name) for name in source.namelist()}

    # Insert a Word comment
    members["word/comments.xml"] = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Reviewer" w:date="2022-08-01T10:00:00Z">
    <w:p><w:r><w:t>Please verify the statistical analysis for Group 5.</w:t></w:r></w:p>
  </w:comment>
</w:comments>"""

    # Inject a tracked change
    document_xml = members["word/document.xml"].decode("utf-8")
    insert_marker = '<w:ins w:id="2" w:author="Author" w:date="2022-07-30T14:00:00Z"><w:r><w:t>revised value</w:t></w:r></w:ins>'
    document_xml = document_xml.replace("<w:body>", f"<w:body>{insert_marker}", 1)
    members["word/document.xml"] = document_xml.encode("utf-8")

    with ZipFile(str(path), "w", ZIP_DEFLATED) as target:
        for name, content in members.items():
            target.writestr(name, content)


if __name__ == "__main__":
    path = copy_and_inject()
    print(f"Created: {path}")
    doc = Document(str(path))
    print(f"Paragraphs: {len(doc.paragraphs)}, Tables: {len(doc.tables)}")
