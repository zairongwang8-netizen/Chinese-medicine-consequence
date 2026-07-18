#!/usr/bin/env python3
"""Flag overclaiming and formulaic prose in Markdown, text, or DOCX manuscripts."""

from argparse import ArgumentParser
from collections import Counter
from pathlib import Path
import re
from xml.etree import ElementTree
from zipfile import ZipFile


PATTERNS = [
    (
        "Major",
        "结论越级",
        re.compile(r"(?:充分证明|全面揭示|明确揭示|首次揭示|完全阐明|彻底阐明|必然导致|关键机制|确切机制|证明|证实)"),
        "核对研究设计和直接证据；本草考证或基础研究通常需要降调。",
    ),
    (
        "Major",
        "传统与现代强行对应",
        re.compile(r"(?:现代研究|现代药理).{0,24}(?:证明|证实).{0,36}(?:传统|性味|归经|功效|证候)"),
        "改为有限关联，并说明传统术语与现代指标不一一对应。",
    ),
    (
        "Major",
        "复方/成分归因风险",
        re.compile(r"(?:复方|方剂|提取物|单体|成分).{0,36}(?:说明|证明|证实).{0,24}(?:该药|本药|药材)"),
        "核对研究对象，避免把复方、提取物或成分结果归给单味饮片。",
    ),
    (
        "Minor",
        "空泛价值判断",
        re.compile(
            r"(?:具有(?:十分|非常|重大|重要)?(?:的)?(?:理论意义|现实意义|实践价值|应用价值)|"
            r"提供(?:了)?(?:重要)?理论依据|推动.{0,16}现代化(?:发展)?|"
            r"为未来研究提供(?:了)?(?:新的)?(?:思路|方向)|"
            r"为.{0,20}提供(?:了)?(?:全新|新的)(?:思路|方向))"
        ),
        "补充具体对象、证据和边界；没有信息增量时删除。",
    ),
    (
        "Minor",
        "模板化衔接",
        re.compile(r"(?:综上所述|值得注意的是|进一步而言|不难发现|毋庸置疑)"),
        "检查是否承担必要逻辑功能；重复或可直接衔接时删除。",
    ),
    (
        "Minor",
        "宏大目标",
        re.compile(
            r"(?:填补.{0,12}(?:国内外)?空白|推动中医药国际化|促进中医药走向世界|"
            r"助力健康中国|实现创造性转化和创新性发展|为[^，。；]{0,20}奠定坚实基础|"
            r"具有广阔(?:的)?应用前景|实现.{0,20}有机统一|可广泛推广)"
        ),
        "除非正文提供直接证据和具体路径，否则改为可检验的有限判断。",
    ),
]

STRONG_CLAIM = re.compile(r"(?:证明|证实|揭示|阐明|显著(?:提高|改善|降低)|关键机制)")
CITATION = re.compile(r"(?:\[\d+(?:[-,]\d+)*\]|\(\d+(?:[-,]\d+)*\)|［\d+(?:[-，,]\d+)*］)")
SENTENCE_SPLIT = re.compile(r"(?<=[。！？；])")
REPETITION_THRESHOLDS = {
    "综上所述": 3,
    "值得注意的是": 3,
    "进一步而言": 3,
    "此外": 4,
    "同时": 4,
}


def read_markdown_or_text(path):
    return [(f"L{number}", line.strip()) for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1)]


def read_docx(path):
    with ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("word/document.xml"))
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    records = []
    for number, paragraph in enumerate(root.iter(namespace + "p"), 1):
        text = "".join(node.text or "" for node in paragraph.iter(namespace + "t")).strip()
        if text:
            records.append((f"P{number}", text))
    return records


def read_records(path):
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return read_docx(path)
    if suffix in {".md", ".txt"}:
        return read_markdown_or_text(path)
    raise ValueError("Supported inputs: .md, .txt, .docx")


def iter_sentences(records):
    in_references = False
    for location, text in records:
        heading = text.lstrip("# ").strip()
        if heading in {"参考文献", "References", "REFERENCES"}:
            in_references = True
            continue
        if in_references:
            continue
        for sentence in SENTENCE_SPLIT.split(text):
            sentence = sentence.strip()
            if sentence:
                yield location, sentence


def audit(records):
    findings = []
    phrase_counts = Counter()
    for location, sentence in iter_sentences(records):
        for phrase in REPETITION_THRESHOLDS:
            phrase_counts[phrase] += sentence.count(phrase)
        for severity, category, pattern, advice in PATTERNS:
            for match in pattern.finditer(sentence):
                phrase = match.group(0)
                findings.append(
                    {
                        "severity": severity,
                        "category": category,
                        "location": location,
                        "phrase": phrase,
                        "excerpt": sentence,
                        "advice": advice,
                    }
                )
        if STRONG_CLAIM.search(sentence) and not CITATION.search(sentence):
            findings.append(
                {
                    "severity": "Major",
                    "category": "强结论缺少可见引文",
                    "location": location,
                    "phrase": STRONG_CLAIM.search(sentence).group(0),
                    "excerpt": sentence,
                    "advice": "检查该句及相邻句的引文是否直接支持结论；脚本不识别所有Word上标，需人工复核。",
                }
            )
    return findings, phrase_counts


def escape_cell(value):
    return str(value).replace("|", "\\|").replace("\n", " ")


def render_report(path, findings, phrase_counts):
    major = sum(item["severity"] == "Major" for item in findings)
    minor = sum(item["severity"] == "Minor" for item in findings)
    lines = [
        "# 多阶段模拟审稿语言风险初筛",
        "",
        f"- 文件：`{path}`",
        f"- Major线索：{major}",
        f"- Minor线索：{minor}",
        "- 说明：本报告只标记风险，不代表最终审稿判断，也不自动修改原文。",
        "",
        "## 逐项线索",
        "",
        "| 级别 | 位置 | 类型 | 命中词 | 语境 | 人工复核建议 |",
        "|---|---|---|---|---|---|",
    ]
    if findings:
        for item in findings:
            lines.append(
                "| {severity} | {location} | {category} | {phrase} | {excerpt} | {advice} |".format(
                    **{key: escape_cell(value) for key, value in item.items()}
                )
            )
    else:
        lines.append("| - | - | 未命中预设风险词 | - | 仍须进行人工证据与创新性审查 | - |")

    repeated = [
        (phrase, count)
        for phrase, count in phrase_counts.items()
        if count >= REPETITION_THRESHOLDS[phrase]
    ]
    lines.extend(["", "## 重复表达", ""])
    if repeated:
        for phrase, count in sorted(repeated, key=lambda item: (-item[1], item[0])):
            lines.append(f"- `{phrase}`：{count}次")
    else:
        lines.append("- 未发现达到预设重复阈值的连接表达。")
    lines.extend(
        [
            "",
            "## 人工复审提醒",
            "",
            "- 直接引文、古籍原文和法规原文不得按本报告机械改写。",
            "- 未命中不等于没有AI式模板语言，也不等于证据充分。",
            "- 修改后重新执行AMA顺序、证据匹配，并交由原三名模拟审稿人复审。",
            "",
        ]
    )
    return "\n".join(lines)


def main():
    parser = ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Markdown, text, or DOCX manuscript")
    parser.add_argument("--output", type=Path, help="Write a Markdown audit report")
    parser.add_argument(
        "--fail-on",
        choices=("none", "major", "any"),
        default="none",
        help="Optional CI exit threshold; default only reports findings",
    )
    args = parser.parse_args()

    records = read_records(args.input)
    findings, phrase_counts = audit(records)
    report = render_report(args.input, findings, phrase_counts)
    if args.output:
        args.output.write_text(report, encoding="utf-8")
        print(args.output)
    else:
        print(report)

    has_major = any(item["severity"] == "Major" for item in findings)
    if args.fail_on == "major" and has_major:
        raise SystemExit(1)
    if args.fail_on == "any" and findings:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
