#!/usr/bin/env python3
"""Validate or reorder canonical numeric citations in a Markdown manuscript.

This utility only handles the internal ``[n]`` representation used by the skill.
It does not choose a journal's citation system or format bibliography entries.
"""

from argparse import ArgumentParser
from pathlib import Path
import re


CITATION_RE = re.compile(r"\[((?:\d+(?:-\d+)?)(?:\s*,\s*\d+(?:-\d+)?)*)\]")
REFERENCE_RE = re.compile(r"^\[(\d+)\]\s+(.*)$")
REFERENCE_HEADING = "## 参考文献"


def expand(group):
    numbers = []
    for item in re.split(r"\s*,\s*", group):
        if "-" in item:
            start, end = map(int, item.split("-", 1))
            if end < start:
                raise ValueError(f"Invalid citation range: {item}")
            numbers.extend(range(start, end + 1))
        else:
            numbers.append(int(item))
    return numbers


def compact(numbers):
    values = sorted(set(numbers))
    if not values:
        return ""
    parts = []
    start = previous = values[0]
    for value in values[1:]:
        if value == previous + 1:
            previous = value
            continue
        parts.append(str(start) if start == previous else f"{start}-{previous}")
        start = previous = value
    parts.append(str(start) if start == previous else f"{start}-{previous}")
    return ",".join(parts)


def parse(path):
    text = path.read_text(encoding="utf-8")
    body, marker, reference_block = text.partition(REFERENCE_HEADING)
    if not marker:
        raise ValueError(f"Missing heading: {REFERENCE_HEADING}")

    references = {}
    for line in reference_block.strip().splitlines():
        match = REFERENCE_RE.match(line)
        if match:
            number = int(match.group(1))
            if number in references:
                raise ValueError(f"Duplicate reference number: {number}")
            references[number] = match.group(2)
    if not references:
        raise ValueError("No canonical [n] references found")
    return body, references


def first_citation_order(body):
    order = []
    seen = set()
    occurrences = []
    for match in CITATION_RE.finditer(body):
        numbers = expand(match.group(1))
        occurrences.extend(numbers)
        for number in numbers:
            if number not in seen:
                seen.add(number)
                order.append(number)
    return order, occurrences


def validate(body, references):
    order, occurrences = first_citation_order(body)
    undefined = sorted(set(occurrences) - set(references))
    uncited = sorted(set(references) - set(occurrences))
    expected = list(range(1, len(references) + 1))
    sequential = order == expected
    return order, undefined, uncited, sequential


def reorder(path, body, references):
    order, undefined, uncited, _ = validate(body, references)
    if undefined:
        raise ValueError(f"Undefined citations: {undefined}")
    if uncited:
        raise ValueError(f"Uncited references: {uncited}")

    mapping = {old: new for new, old in enumerate(order, start=1)}

    def replace(match):
        return f"[{compact(mapping[number] for number in expand(match.group(1)))}]"

    new_body = CITATION_RE.sub(replace, body)
    new_references = [f"[{mapping[old]}] {references[old]}" for old in order]
    path.write_text(
        new_body.rstrip()
        + "\n\n"
        + REFERENCE_HEADING
        + "\n"
        + "\n".join(new_references)
        + "\n",
        encoding="utf-8",
    )


def main():
    parser = ArgumentParser(description=__doc__)
    parser.add_argument("manuscript", type=Path)
    parser.add_argument(
        "--write",
        action="store_true",
        help="Rewrite canonical [n] citations and references in first-use order",
    )
    args = parser.parse_args()

    body, references = parse(args.manuscript)
    if args.write:
        reorder(args.manuscript, body, references)
        body, references = parse(args.manuscript)

    order, undefined, uncited, sequential = validate(body, references)
    print(
        "system=numeric-sequential "
        f"references={len(references)} citations={len(order)} sequential={sequential}"
    )
    if undefined:
        print(f"undefined={undefined}")
    if uncited:
        print(f"uncited={uncited}")
    if undefined or uncited or not sequential:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
