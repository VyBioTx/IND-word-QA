from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document

FIXTURE_DIR = Path(__file__).resolve().parent
COMPOUND_REPORT = FIXTURE_DIR / "CPD-001_study_report.docx"


def create_compound_report(path: Path = COMPOUND_REPORT) -> Path:
    doc = Document()

    section = doc.sections[0]
    section.header.paragraphs[0].text = "CPD-001 Study Report v2.0"
    section.footer.paragraphs[0].text = "Confidential - Version 2.1"

    doc.add_heading("1  Introduction", level=1)
    doc.add_paragraph("CPD-001 is a novel compound under investigation for the treatment of breast cancer.")
    doc.add_paragraph("The aim of this study is to evaluate the  in vivo efficacy of CPD001.")
    doc.add_paragraph("Tumor growth inhibition (TGI) was assessed as a primary endpoint.")
    doc.add_paragraph("PK parameters were evaluated to characterize the pharmacokinetic profile.")

    doc.add_heading("2  Methods", level=1)

    doc.add_heading("2.1  Cell Culture", level=2)
    doc.add_paragraph("MCF-7 cells were cultured in DMEM medium supplemented with 10% FBS.")
    doc.add_paragraph("See Figure 1 for cell growth characteristics.")

    doc.add_heading("2.2  Animal Model", level=2)
    doc.add_paragraph("Female BALB/c nude mice (6-8 weeks old) were inoculated subcutaneously.")
    doc.add_paragraph("The  dose  of 5x10^6 cells was injected into the right flank.")

    doc.add_heading("2.3  Study Design", level=2)
    doc.add_paragraph("Animals were randomized into 4 groups (n=8 per group).")
    doc.add_paragraph("See Table 3 for group allocation.")

    doc.add_heading("3  Results", level=1)

    doc.add_heading("3.1  Tumor Growth Inhibition", level=2)
    doc.add_paragraph("The the tumor volume data are summarized in Table 1 below.")
    doc.add_paragraph("3/8 animals (40.0%) in Group 1 showed complete regression.")

    doc.add_paragraph("Table 1. Tumor volume (mm^3) on Day 28")
    table = doc.add_table(rows=6, cols=4)
    table.cell(0, 0).text = "Group"
    table.cell(0, 1).text = "Mean (mm^3)"
    table.cell(0, 2).text = "SD"
    table.cell(0, 3).text = "N"
    table.cell(1, 0).text = "Vehicle"
    table.cell(1, 1).text = "1850"
    table.cell(1, 2).text = "420"
    table.cell(1, 3).text = "8"
    table.cell(2, 0).text = "CPD001 10mg/kg"
    table.cell(2, 1).text = "980"
    table.cell(2, 2).text = "310"
    table.cell(2, 3).text = "8"
    table.cell(3, 0).text = "CPD001 30mg/kg"
    table.cell(3, 1).text = "520"
    table.cell(3, 2).text = ""
    table.cell(3, 3).text = "8"
    table.cell(4, 0).text = "CPD001 60mg/kg"
    table.cell(4, 1).text = "310"
    table.cell(4, 2).text = "95"
    table.cell(4, 3).text = "8"
    table.cell(5, 0).text = "Total"
    table.cell(5, 1).text = "3500"
    table.cell(5, 2).text = "-"
    table.cell(5, 3).text = "32"

    doc.add_heading("3.2  Pharmacokinetic Analysis", level=2)
    doc.add_paragraph("PK analysis was performed on plasma samples collected at 6 time points.")
    doc.add_paragraph("AUC(0-24h) was calculated for each dose group.")
    doc.add_paragraph("Maximum concentration (Cmax) was determined for each animal.")
    doc.add_paragraph("Pharmacokinetic (PK) parameters are shown in Table 2 below.")

    doc.add_paragraph("Table 2. PK parameters")
    pk_table = doc.add_table(rows=3, cols=1)
    pk_table.cell(0, 0).text = "Not Applicable"
    pk_table.cell(1, 0).text = "N/A"
    pk_table.cell(2, 0).text = "NA"

    doc.add_heading("3.3  Body Weight Changes", level=2)
    doc.add_paragraph("Body weight was monitored throughout the study.")
    doc.add_paragraph("No significant changes were observed in any group.")
    doc.add_paragraph("Protein Kinase (PK) signaling pathways may contribute to tumor growth.")
    doc.add_paragraph("See Table 2 for comprehensive analyses.")
    doc.add_paragraph("Please refer to Figure 2 for graphical representation.")

    doc.add_heading("3.4  Safety Assessment", level=2)
    doc.add_paragraph("No treatment-related mortality was observed.")
    doc.add_paragraph("TEAE were monitored and recorded daily.")
    doc.add_paragraph("dose-limiting toxicity (DLT) was not observed at any dose level.")

    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph("")

    doc.add_heading("4.1  Overall Conclusions", level=2)
    doc.add_paragraph("CPD-001 demonstrated dose-dependent antitumor activity in the MCF-7 model.")

    doc.add_heading("5  References", level=1)
    doc.add_paragraph("All procedures followed institutional guidelines.")

    doc.save(str(path))
    _inject_compound_xml(path)
    return path


def _inject_compound_xml(path: Path) -> None:
    with ZipFile(str(path), "r") as source:
        members = {name: source.read(name) for name in source.namelist()}

    members["word/comments.xml"] = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Internal Reviewer" w:date="2026-05-15T14:30:00Z">
    <w:p><w:r><w:t>Please verify the statistics for Group 2.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="1" w:author="QA Reviewer" w:date="2026-05-20T09:15:00Z">
    <w:p><w:r><w:t>Add SD value for Group 3 (CPD001 30mg/kg).</w:t></w:r></w:p>
  </w:comment>
</w:comments>"""

    document_xml = members["word/document.xml"].decode("utf-8")
    document_xml = document_xml.replace(
        "<w:body>",
        '<w:body>'
        '<w:p>'
        '<w:ins w:id="1" w:author="Author" w:date="2026-05-10T10:00:00Z">'
        '<w:r><w:t>revised PK value</w:t></w:r>'
        '</w:ins>'
        '</w:p>',
        1,
    )
    members["word/document.xml"] = document_xml.encode("utf-8")

    with ZipFile(str(path), "w", ZIP_DEFLATED) as target:
        for name, content in members.items():
            target.writestr(name, content)


if __name__ == "__main__":
    path = create_compound_report()
    print(f"Created: {path}")
