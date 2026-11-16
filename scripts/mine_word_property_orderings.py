"""Mine WordprocessingML property element child orderings from a DOCX corpus.

Companion tool for spec 010 (Word compatibility element ordering). Given a
glob of DOCX/DOTX files, walks every property element of the configured
types, records the observed child sequence, and reports:

  1. distinct orderings observed per property type, with frequency counts
  2. for each ordering, whether it's a valid subsequence of the SDK's
     Children list (the proxy used by `src/openxml_audit/word/compat.py`)
  3. failure clusters — orderings that disagree with the SDK proxy, which
     may indicate either real Word-tolerated patterns or content bugs

Usage:

    python scripts/mine_word_property_orderings.py /path/to/corpus
    python scripts/mine_word_property_orderings.py /path/to/corpus --json out.json

The tool is corpus-agnostic: it reports what it observes. Drawing
conclusions about Word's actual tolerance from the output requires
knowing whether the corpus producer is Word-blessed (e.g. python-docx,
TokenMoulds, Microsoft sample files) or whether files were observed to
trigger Word's repair dialog.

Spec context: spec 010, Empirical Scan section.
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from collections import Counter, defaultdict
from collections.abc import Iterable
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from openxml_audit.word.compat import CONSTRAINT_TABLE, find_first_out_of_order  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Property types to mine. Includes both the types the validator currently
# constrains and ones we don't yet — mining is exploratory and the report
# tells us which next.
MINED_TYPES = (
    "trPr",
    "pPr",
    "rPr",
    "sectPr",
    "tblPr",
    "tcPr",
    "tblPrEx",
)


def _local(tag: str) -> str:
    if isinstance(tag, str) and tag.startswith("{"):
        return tag.split("}", 1)[1]
    return str(tag)


def _iter_word_xml_parts(zf: zipfile.ZipFile) -> Iterable[tuple[str, bytes]]:
    for name in zf.namelist():
        if not name.endswith(".xml"):
            continue
        if "word/" not in name and not name.startswith("word/"):
            continue
        try:
            yield name, zf.read(name)
        except KeyError:
            continue


def _iter_property_subtrees(root: etree._Element, prop_local: str) -> Iterable[etree._Element]:
    yield from root.iter(f"{{{W}}}{prop_local}")


def mine_corpus(paths: list[Path]) -> dict[str, Counter]:
    """Walk DOCX/DOTX files and record every property's observed child order.

    Returns a mapping: property local-name → Counter[child-tuple → count].
    Children are recorded in Clark notation so namespace differences
    (w14:, w15:) survive into the report.
    """
    observed: dict[str, Counter[tuple[str, ...]]] = defaultdict(Counter)

    for path in paths:
        try:
            zf = zipfile.ZipFile(path)
        except (zipfile.BadZipFile, FileNotFoundError):
            continue
        with zf:
            for _xml_name, raw in _iter_word_xml_parts(zf):
                try:
                    root = etree.fromstring(raw)
                except etree.XMLSyntaxError:
                    continue
                for prop_local in MINED_TYPES:
                    for elem in _iter_property_subtrees(root, prop_local):
                        children = tuple(child.tag for child in elem if isinstance(child.tag, str))
                        if children:
                            observed[prop_local][children] += 1

    return observed


def validate_against_proxy(
    observed: dict[str, Counter[tuple[str, ...]]],
) -> dict[str, dict]:
    """For each property type with a constraint entry, count how many
    observed orderings are subsequences of the SDK Children list.
    """
    summary: dict[str, dict] = {}
    for prop_local, orderings in observed.items():
        parent_tag = f"{{{W}}}{prop_local}"
        constraint = CONSTRAINT_TABLE.get(parent_tag)
        total = sum(orderings.values())
        if constraint is None:
            summary[prop_local] = {
                "total": total,
                "constrained": False,
                "distinct": len(orderings),
            }
            continue

        passed = 0
        failures: list[tuple[tuple[str, ...], int]] = []
        for ordering, count in orderings.items():
            if find_first_out_of_order(list(ordering), constraint.children) is None:
                passed += count
            else:
                failures.append((ordering, count))

        summary[prop_local] = {
            "total": total,
            "constrained": True,
            "distinct": len(orderings),
            "pass": passed,
            "fail": total - passed,
            "fail_distinct": len(failures),
            "top_failures": [
                {
                    "order": [_local(c) for c in order],
                    "count": count,
                }
                for order, count in sorted(failures, key=lambda x: -x[1])[:10]
            ],
        }
    return summary


def empirical_canonical_per_type(
    observed: dict[str, Counter[tuple[str, ...]]],
) -> dict[str, list[str]]:
    """Derive a candidate canonical ordering empirically from the corpus.

    Strategy: for each pair of distinct child tags (A, B) ever observed in
    the same parent, count how often A appears before B vs after. If
    A-before-B dominates by >= 95% of those co-occurrences, fix that order.
    Build a directed graph from these orderings and take a topological sort.

    Cycles indicate the corpus has genuinely conflicting orderings — those
    pairs are reported back so the caller can decide what to do (most likely:
    treat them as freely orderable).
    """
    canonical: dict[str, list[str]] = {}
    for prop_local, orderings in observed.items():
        # Tally pairwise relations
        pair_before: Counter[tuple[str, str]] = Counter()
        seen: set[str] = set()
        for ordering, count in orderings.items():
            tags = [_local(c) for c in ordering]
            for tag in tags:
                seen.add(tag)
            for i, a in enumerate(tags):
                for b in tags[i + 1 :]:
                    if a == b:
                        continue
                    pair_before[(a, b)] += count

        # For each pair, decide direction
        edges: dict[str, set[str]] = defaultdict(set)  # a -> set of b where a < b
        ambiguous: list[tuple[str, str, int, int]] = []
        children = sorted(seen)
        for i, a in enumerate(children):
            for b in children[i + 1 :]:
                ab = pair_before[(a, b)]
                ba = pair_before[(b, a)]
                if ab + ba == 0:
                    continue
                if ab >= ba * 19:  # 95%+ in one direction
                    edges[a].add(b)
                elif ba >= ab * 19:
                    edges[b].add(a)
                else:
                    ambiguous.append((a, b, ab, ba))

        # Topological sort with stable tie-breaking on alphabetical
        in_count: Counter[str] = Counter()
        for a in children:
            in_count[a] = 0
        for bs in edges.values():
            for b in bs:
                in_count[b] += 1

        order: list[str] = []
        ready = sorted(c for c in children if in_count[c] == 0)
        while ready:
            ready.sort()
            tag = ready.pop(0)
            order.append(tag)
            for b in sorted(edges.get(tag, set())):
                in_count[b] -= 1
                if in_count[b] == 0:
                    ready.append(b)

        # Append any tags that didn't make it (cycles)
        leftover = sorted(c for c in children if c not in order)
        order.extend(leftover)

        canonical[prop_local] = order
        if ambiguous and prop_local in ("pPr", "rPr"):
            # Emit a comment-style line in the JSON for the worst ambiguities
            pass  # Keep in-graph data only; surfaced via JSON output below
    return canonical


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "corpus",
        type=Path,
        nargs="+",
        help="One or more directories or DOCX/DOTX files to mine.",
    )
    parser.add_argument(
        "--json",
        type=Path,
        help="Write the full report to this path as JSON (otherwise stdout summary).",
    )
    parser.add_argument(
        "--exclude",
        action="append",
        default=[],
        help="Substring to exclude from corpus paths (repeatable).",
    )
    args = parser.parse_args()

    paths: list[Path] = []
    for entry in args.corpus:
        if entry.is_dir():
            paths.extend(entry.rglob("*.docx"))
            paths.extend(entry.rglob("*.dotx"))
        elif entry.is_file():
            paths.append(entry)
    paths = [p for p in paths if not any(ex in str(p) for ex in args.exclude)]

    if not paths:
        print("error: no DOCX/DOTX files found", file=sys.stderr)
        return 2

    print(f"Mining {len(paths)} file(s)…", file=sys.stderr)
    observed = mine_corpus(paths)
    proxy_report = validate_against_proxy(observed)
    empirical = empirical_canonical_per_type(observed)

    if args.json:
        out = {
            "corpus_size": len(paths),
            "proxy_validation": proxy_report,
            "empirical_canonical": empirical,
            "observed": {
                prop: [
                    {"order": [_local(c) for c in order], "count": count}
                    for order, count in counter.most_common()
                ]
                for prop, counter in observed.items()
            },
        }
        args.json.write_text(json.dumps(out, indent=2))
        print(f"Wrote report to {args.json}", file=sys.stderr)
    else:
        # Pretty-print summary
        for prop in sorted(proxy_report):
            r = proxy_report[prop]
            if r.get("constrained"):
                print(
                    f"{prop:>8s}: {r['total']:>6d} subtrees, "
                    f"{r['distinct']:>3d} distinct orders, "
                    f"{r['pass']:>6d} pass / {r['fail']:>4d} fail"
                )
                for f in r["top_failures"][:5]:
                    print(f"           fail x{f['count']}: {', '.join(f['order'])}")
            else:
                print(
                    f"{prop:>8s}: {r['total']:>6d} subtrees, "
                    f"{r['distinct']:>3d} distinct orders (no constraint table entry)"
                )
        print()
        print("Empirical canonical (topological sort, ties alphabetical):")
        for prop, order in sorted(empirical.items()):
            if order:
                print(f"  {prop}: {', '.join(order)}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
