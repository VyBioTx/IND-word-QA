#!/usr/bin/env -S pixi run python
"""Analyze the modified report and run full QA, generating detailed test report."""

from pathlib import Path
import json

from app.models.document import Document
from app.rules.base import RuleContext
from app.services.docx_parser import DocxParserService
from app.services.qa_engine import QAEngine
from scripts.generate_from_real_template import copy_and_inject


def run_qa(path: Path):
    document = Document(
        id="doc_modified",
        project_id="project_modified",
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


def save_report(report: str, path: Path):
    path.write_text(report, encoding="utf-8")
    print(f"Report saved to: {path}")


if __name__ == "__main__":
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        modified_path = tmp_path / "modified_report.docx"
        path = copy_and_inject(modified_path)
        print(f"\nModified document: {path}")
        print(f"File size: {path.stat().st_size:,} bytes")

        issues = run_qa(path)
        rule_ids_found = {i.rule_id for i in issues}

        all_14_rules = [
            "TEXT-001", "TEXT-002", "NUM-001",
            "TABLE-001", "TABLE-002", "TABLE-003",
            "REF-001", "HEAD-001",
            "WORD-001", "WORD-002", "WORD-003",
            "ABBR-001", "ABBR-002", "TERM-001",
        ]
        missing = set(all_14_rules) - rule_ids_found

        lines = []
        lines.append("=" * 78)
        lines.append("审核系统测试报告 — 基于真实研究报告模板")
        lines.append("=" * 78)
        lines.append("")
        lines.append(f"原始文档: {path.name}")
        lines.append(f"总段落数(含嵌入): 见分析输出")
        lines.append(f"总规则数: 14")
        lines.append(f"触发规则数: {len(rule_ids_found)}")
        lines.append(f"总检出问题数: {len(issues)}")
        lines.append("")
        lines.append("-" * 78)
        lines.append(f"{'规则ID':12s} | {'严重等级':10s} | {'预期':8s} | {'实际':8s} | {'结果':10s} | {'问题数':8s}")
        lines.append("-" * 78)
        pass_count = 0
        fail_count = 0
        for rid in all_14_rules:
            matched = [i for i in issues if i.rule_id == rid]
            triggered = rid in rule_ids_found
            status = "✓ PASS" if triggered else "✗ FAIL"
            if triggered:
                pass_count += 1
            else:
                fail_count += 1
            severity = matched[0].severity if matched else "?"
            lines.append(f"{rid:12s} | {severity:10s} | {'Yes':8s} | {'Yes' if triggered else 'No':8s} | {status:10s} | {str(len(matched)):8s}")
        lines.append("-" * 78)
        lines.append(f"{'TOTAL':12s} | {'':10s} | {str(len(all_14_rules))+' rules':8s} | {f'{pass_count} pass':8s} | {f'{fail_count} fail':10s} | {str(len(issues)):8s}")
        lines.append("-" * 78)
        lines.append("")
        lines.append("=" * 78)
        lines.append("各规则详细信息")
        lines.append("=" * 78)
        lines.append("")

        for rid in all_14_rules:
            matched = [i for i in issues if i.rule_id == rid]
            triggered = rid in rule_ids_found
            severity = matched[0].severity if matched else "?"
            title = matched[0].title if matched else "(未触发)"
            lines.append(f"[{rid}] {title} (Severity: {severity})")
            lines.append(f"  预期: 应检测出该文档中的对应问题")
            lines.append(f"  实际: {'已触发' if triggered else '未触发'} ({len(matched)} 个问题)")
            if matched:
                for idx, issue in enumerate(matched[:3]):
                    lines.append(f"  -- Issue {idx+1} --")
                    lines.append(f"  原文: {issue.source_text[:100]}")
                    lines.append(f"  描述: {issue.description[:120]}")
                    lines.append(f"  位置: {issue.location}")
            lines.append("")

        lines.append("=" * 78)
        lines.append("嵌入的14类问题对照表")
        lines.append("=" * 78)
        lines.append("")
        embedded = [
            ("TEXT-001", "重复单词 'the the'", "新增段落"),
            ("TEXT-002", "连续空格 '  efficacy  was'", "新增段落"),
            ("NUM-001", "百分比错误 5/8=75.0% (应为62.5%)", "新增段落"),
            ("TABLE-001", "表格Total=22 (实际应=24)", "新增表格"),
            ("TABLE-002", "表格空单元格 (Cmax值为空)", "新增表格"),
            ("TABLE-003", "N/A写法不一致 (Not Applicable/N/A/NA)", "新增表格"),
            ("REF-001", "引用Figure 10和Table 25不存在", "新增段落"),
            ("HEAD-001", "标题编号跳号 6.1→6.3 (缺6.2)", "新增标题"),
            ("WORD-001", "页眉页脚版本 [v1.0] vs [Version 1.1]", "修改页眉页脚"),
            ("WORD-002", "遗留Word批注 XML注入", "XML注入"),
            ("WORD-003", "遗留修订痕迹 XML注入", "XML注入"),
            ("ABBR-001", "缩写MTD和SAE首次使用未定义", "新增段落"),
            ("ABBR-002", "MTD多定义 (maximum tolerated dose vs mean tumor diameter)", "新增段落"),
            ("TERM-001", "术语不一致 (CPD-001 vs CPD001)", "新增段落"),
        ]
        for rid, issue, method in embedded:
            triggered = rid in rule_ids_found
            lines.append(f"  {'✓' if triggered else '✗'} {rid}: {issue}")
            lines.append(f"     嵌入方式: {method}")
            lines.append("")

        lines.append("=" * 78)
        if fail_count == 0:
            lines.append("结论: 所有14条规则均成功触发！系统能够正确发现嵌入的所有问题类型。")
            lines.append(f"共检出 {len(issues)} 个问题。")
        else:
            lines.append(f"结论: {fail_count}/14 条规则未触发。")
        lines.append("=" * 78)

        report_text = "\n".join(lines)
        print(report_text)

        # Save report
        out_dir = Path(__file__).resolve().parent.parent
        save_report(report_text, out_dir / "modified_report_test_result.txt")
