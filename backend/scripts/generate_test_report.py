#!/usr/bin/env -S pixi run python
"""Generate comprehensive CPD001 study report test results."""

from pathlib import Path
import json

from app.models.document import Document
from app.rules.base import RuleContext
from app.services.docx_parser import DocxParserService
from app.services.qa_engine import QAEngine
from app.tests.fixtures.generate_compound_report import create_compound_report


def run_qa(path: Path):
    document = Document(
        id="doc_compound",
        project_id="project_compound",
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


def generate_report(tmp_path: Path) -> str:
    path = create_compound_report(tmp_path / "CPD-001_study_report.docx")
    issues = run_qa(path)

    rule_ids_found = {i.rule_id for i in issues}

    all_rules = [
        ("TEXT-001", "DuplicateWord", "Low", "重复英文单词", lambda i: "the the" in i.source_text.lower() if i.source_text else False),
        ("TEXT-002", "RepeatedSpaces", "Low", "连续多个空格", lambda i: True),
        ("NUM-001", "PercentageCalculation", "High", "百分比计算不匹配", lambda i: "40.0%" in i.source_text),
        ("TABLE-001", "SimpleTableTotal", "High", "表格合计不匹配", lambda i: "3500" in i.source_text),
        ("TABLE-002", "EmptyTableCell", "Medium", "表格空单元格", lambda i: True),
        ("TABLE-003", "InconsistentNANotation", "Low", "N/A写法不一致", lambda i: True),
        ("REF-001", "MissingTableFigureRef", "Medium", "引用表/图不存在", lambda i: True),
        ("HEAD-001", "HeadingNumberingGap", "Medium", "标题编号跳号", lambda i: True),
        ("WORD-001", "HeaderFooterVersion", "High", "页眉页脚版本不一致", lambda i: True),
        ("WORD-002", "RemainingComments", "High", "遗留Word批注", lambda i: True),
        ("WORD-003", "TrackedChanges", "High", "遗留修订痕迹", lambda i: True),
        ("ABBR-001", "AbbreviationFirstDef", "Medium", "缩写首次未定义", lambda i: True),
        ("ABBR-002", "AbbreviationMultiDef", "Medium", "缩写多定义冲突", lambda i: True),
        ("TERM-001", "SimpleTerminology", "Medium", "术语拼写不一致", lambda i: True),
    ]

    lines = []
    lines.append("=" * 78)
    lines.append("测试报告 — CPD-001 化合物")
    lines.append("药效学和药代动力学研究 — 审核系统测试报告")
    lines.append("=" * 78)
    lines.append("")
    lines.append(f"测试时间: 2026-06-04")
    lines.append(f"测试文档: CPD-001_study_report.docx")
    lines.append(f"审核规则总数: 14")
    lines.append(f"触发规则数: {len(rule_ids_found)}")
    lines.append(f"未触发规则数: {14 - len(rule_ids_found)}")
    lines.append(f"总检出问题数: {len(issues)}")
    lines.append("")

    lines.append("-" * 78)
    lines.append("│ 规则ID      │ 严重等级   │ 预期触发  │ 实际触发  │ 结果")
    lines.append("-" * 78)

    pass_count = 0
    fail_count = 0
    for rule_id, _, severity, name, _ in all_rules:
        triggered = rule_id in rule_ids_found
        status = "✓ PASS" if triggered else "✗ FAIL"
        if triggered:
            pass_count += 1
        else:
            fail_count += 1
        matched = [i for i in issues if i.rule_id == rule_id]
        count = f"{len(matched)} issues" if matched else ""
        lines.append(f"│ {rule_id:10s} │ {severity:10s} │ Yes        │ {'Yes' if triggered else 'No ':9s} │ {status}  {count}")
    lines.append("-" * 78)
    lines.append(f"│ 总计        │           │ 14 rules   │ {pass_count} pass / {fail_count} fail   │ {'ALL PASS' if fail_count == 0 else f'{fail_count} FAILED':10s}")
    lines.append("-" * 78)
    lines.append("")

    lines.append("=" * 78)
    lines.append("详细信息：每条规则的测试用例、预期结果和实际结果")
    lines.append("=" * 78)
    lines.append("")

    for rule_id, class_name, severity, name, checker in all_rules:
        matched = [i for i in issues if i.rule_id == rule_id]
        triggered = len(matched) > 0

        lines.append(f"[{rule_id}] {name} (Severity: {severity})")
        lines.append(f"  规则类: {class_name}Rule")
        lines.append(f"  预期:    系统应检测出该文档中的{name}问题")
        lines.append(f"  实际:    {'已触发' if triggered else '未触发'} ({len(matched)} 个问题)")

        if matched:
            for idx, issue in enumerate(matched):
                lines.append(f"  --- Issue {idx + 1} ---")
                lines.append(f"  原文:    {issue.source_text[:80]}")
                lines.append(f"  描述:    {issue.description[:120]}")
                lines.append(f"  建议:    {issue.suggestion[:120]}")
                lines.append(f"  位置:    {issue.location}")
                if issue.evidence:
                    lines.append(f"  证据:    {json.dumps(issue.evidence, ensure_ascii=False)[:120]}")
        else:
            lines.append(f"  原因:    文档中未包含触发该规则的内容")
        lines.append("")

    lines.append("=" * 78)
    lines.append("测试文档嵌入的问题清单")
    lines.append("=" * 78)
    lines.append("")
    embedded_issues = [
        ("TEXT-001", "重复英文单词 'the the'", "第18段: 'The the tumor volume data...'"),
        ("TEXT-002", "Heading/段落中连续多个空格", "多处（14处），如 '1  Introduction', 'The  dose  of'"),
        ("NUM-001", "百分比计算错误", "第19段: '3/8 animals (40.0%)'，正确值应为37.5%"),
        ("TABLE-001", "表格合计行数值错误", "Table 1: Total列显示3500，实际应为3660"),
        ("TABLE-002", "表格空单元格", "Table 1: SD列中CPD001 30mg/kg行为空"),
        ("TABLE-003", "N/A写法不一致", "Table 2中混合使用 Not Applicable, N/A, NA"),
        ("REF-001", "引用的Table/Figure不存在", "'Figure 1', 'Table 3', 'Figure 2'无对应标题"),
        ("HEAD-001", "标题编号跳号", "缺少 4 (从 3.4 直接跳到 4.1，再跳到 5)"),
        ("WORD-001", "页眉页脚版本号不一致", "Header: v2.0, Footer: Version 2.1"),
        ("WORD-002", "遗留Word批注", "两条批注: Internal Reviewer + QA Reviewer"),
        ("WORD-003", "遗留修订痕迹", "document.xml中插入标记 '<w:ins>'"),
        ("ABBR-001", "缩写首次使用未定义", "'PK parameters' 首次使用PK未定义"),
        ("ABBR-002", "缩写多定义", "PK被定义为 'pharmacokinetic' 和 'protein kinase'"),
        ("TERM-001", "术语拼写不一致", "CPD-001 vs CPD001"),
    ]
    for rule_id, issue, location in embedded_issues:
        lines.append(f"  {rule_id}: {issue}")
        lines.append(f"          {location}")
        lines.append("")

    lines.append("=" * 78)
    lines.append("结论")
    lines.append("=" * 78)
    lines.append("")
    if fail_count == 0:
        lines.append("所有 14 条审核规则均已成功触发。")
        lines.append(f"系统从CPD-001研究报告中检出 {len(issues)} 个问题，")
        lines.append("覆盖所有严重等级（High/Medium/Low）。")
        lines.append("测试通过！")
    else:
        lines.append(f"未通过规则数: {fail_count}/14")

    lines.append("")
    return "\n".join(lines)


if __name__ == "__main__":
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        report = generate_report(Path(tmp))
    print(report)
