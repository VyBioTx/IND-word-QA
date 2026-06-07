from __future__ import annotations

from pathlib import Path

from app.models.document import Document
from app.rules.base import RuleContext
from app.services.docx_parser import DocxParserService
from app.services.qa_engine import QAEngine
from app.tests.fixtures.generate_compound_report import create_compound_report
from app.tests.fixtures.generate_sample_docx import create_sample_bad_report, create_sample_clean_report


def _run_fixture(path: Path) -> list:
    document = Document(
        id=f"doc_{path.stem}",
        project_id="project_comprehensive",
        filename=path.name,
        stored_path=str(path),
        file_size=path.stat().st_size,
        status="parsed",
    )
    parsed = DocxParserService().parse(document.id, str(path))
    return QAEngine().run_document_qa(
        RuleContext(
            project_id=document.project_id,
            document=document,
            parsed_document=parsed,
            all_project_documents=[parsed],
        )
    )


def test_compound_report_triggers_all_14_rules(tmp_path: Path) -> None:
    path = create_compound_report(tmp_path / "CPD-001_study_report.docx")
    issues = _run_fixture(path)

    rule_ids = {issue.rule_id for issue in issues}
    all_14_rules = {
        "TEXT-001", "TEXT-002", "NUM-001",
        "TABLE-001", "TABLE-002", "TABLE-003",
        "REF-001",
        "HEAD-001",
        "WORD-001", "WORD-002", "WORD-003",
        "ABBR-001", "ABBR-002",
        "TERM-001",
    }
    missing = all_14_rules - rule_ids

    print(f"\n{'='*70}")
    print(f"测试报告 — 审核规则触发情况 (Issues detected: {len(issues)})")
    print(f"{'='*70}")

    report_lines = []
    for rule_id in sorted(all_14_rules):
        matched = [i for i in issues if i.rule_id == rule_id]
        status = "PASS" if matched else "FAIL"
        symbol = "✓" if matched else "✗"
        count = len(matched)
        desc = matched[0].title if matched else "(未触发)"
        line = f"  {symbol} {rule_id:12s} | {status:4s} | {desc}"
        report_lines.append(line)
        print(line)
        if matched:
            for m in matched:
                print(f"             -> Source: {m.source_text[:60]}...")
                print(f"             -> Location: {m.location}")
                print(f"             -> Description: {m.description[:80]}")

    print(f"\n{'='*70}")
    print(f"Missing rules: {missing if missing else 'None — all 14 rules triggered!'}")
    print(f"{'='*70}")

    assert len(missing) == 0, f"Rules not triggered: {missing}"


def test_sample_bad_report_generates_expected_issues(tmp_path: Path) -> None:
    issues = _run_fixture(create_sample_bad_report(tmp_path / "sample_bad_report.docx"))
    rule_ids = {issue.rule_id for issue in issues}
    assert len(issues) >= 8
    assert {"TEXT-001", "NUM-001", "TABLE-001", "WORD-002", "WORD-003"}.issubset(rule_ids)


def test_sample_clean_report_generates_no_high_issues(tmp_path: Path) -> None:
    issues = _run_fixture(create_sample_clean_report(tmp_path / "sample_clean_report.docx"))
    assert not [issue for issue in issues if issue.severity == "High"]
