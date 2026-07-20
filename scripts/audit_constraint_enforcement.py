"""Audit whether bridged semantic constraints actually enforce anything.

Spec 037. The schematron bridge can produce a perfectly-formed constraint that
never matches an attribute, so it never rejects anything — a silent inert rule.
`test_schematron_coverage.py` counts such a rule as converted; nothing else
asks whether it fires.

Method (executed, the gold standard): build each constraint the way the running
validator does, synthesise an instance that violates it — with the attribute in
the form the SDK schema data *declares* (qualified iff its QName carries a
prefix) — and run `validate()`. Emits a finding -> LIVE. Silent on a genuine
violation -> DEAD.

Harness self-check: `w:` attributes are genuinely namespaced in instance
documents, so they MUST classify LIVE. If any classify DEAD the harness itself
is wrong and no number it prints can be trusted.

Usage:
    python -m scripts.audit_constraint_enforcement
    python -m scripts.audit_constraint_enforcement --list-dead
"""

from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict
from pathlib import Path

from lxml import etree

from openxml_audit.codegen import get_schematron_registry
from openxml_audit.codegen.schematron_bridge import (
    _get_namespace_map,
    create_constraint_from_schematron,
)
from openxml_audit.context import ValidationContext

SCHEMA_DIR = Path(__file__).resolve().parent.parent / "data" / "openxml" / "schemas"

# Tier 1 (executed). Only UNIQUE_ATTRIBUTE has a probe that is *known* to
# violate every constraint of its type: two elements sharing one value. This
# tier passes the harness self-check.
#
# A generic "long non-numeric string" probe was tried for the other types and
# rejected: it produces false DEADs (ATTRIBUTE_NOT_EQUAL is only violated by a
# value that *equals* the forbidden one; ATTRIBUTE_COMPARISON needs two
# attributes). The harness check caught it — 5 w: constraints wrongly went
# DEAD. Executing those types needs per-type violation generators.
EXECUTED_TYPES = {"UNIQUE_ATTRIBUTE"}

# Tier 2 (declarative). Every attribute-bearing type: does the constraint look
# for the attribute in the form real files write it? Cross-validated by
# agreeing exactly with tier 1 on UNIQUE_ATTRIBUTE.
ATTRIBUTE_BEARING = {
    "UNIQUE_ATTRIBUTE",
    "ATTRIBUTE_VALUE_RANGE",
    "ATTRIBUTE_VALUE_LENGTH",
    "ATTRIBUTE_EQUALS",
    "ATTRIBUTE_NOT_EQUAL",
    "ATTRIBUTE_VALUE_PATTERN",
    "ATTRIBUTE_COMPARISON",
    "ELEMENT_REFERENCE",
}


def build_attr_index() -> dict[str, dict[str, set[str]]]:
    """element qname -> attr local -> declared prefixes ('' means unqualified).

    Attributes are frequently declared on a base type rather than the element
    type itself (w:CT_FtnEdn/w:endnote declares none; its BaseClass
    FootnoteEndnoteType declares w:id), so the BaseClass chain is walked.

    `ClassName` is NOT unique across namespaces — 385 names collide (e.g.
    `ColorType` is both `a:CT_Color/` with no attributes and `x:CT_Color/`
    with five). A global ClassName index silently resolves a DrawingML base
    to a SpreadsheetML type and invents attributes. Resolution is therefore
    scoped to the declaring schema file first, falling back to a global lookup
    only when the name is globally unambiguous.
    """
    all_types = []
    by_file: dict[int, dict] = {}
    global_by_class_name: dict = defaultdict(list)
    for path in sorted(SCHEMA_DIR.glob("*.json")):
        file_index: dict = {}
        for type_def in json.loads(path.read_text()).get("Types", []):
            type_def["__file"] = id(file_index)
            all_types.append(type_def)
            if class_name := type_def.get("ClassName"):
                file_index[class_name] = type_def
                global_by_class_name[class_name].append(type_def)
        by_file[id(file_index)] = file_index

    def resolve_base(base: str, type_def) -> dict | None:
        # Same schema file wins; cross-namespace inheritance is only followed
        # when exactly one type in the whole dataset claims the name.
        local = by_file.get(type_def.get("__file"), {}).get(base)
        if local is not None:
            return local
        candidates = global_by_class_name.get(base, [])
        return candidates[0] if len(candidates) == 1 else None

    def own_attrs(type_def):
        pairs = []
        for attr in type_def.get("Attributes", []):
            qname = attr.get("QName", "")
            if ":" in qname:
                prefix, local = qname.split(":", 1)
                pairs.append((local, prefix))
        return pairs

    def with_inherited(type_def):
        pairs, seen, current = [], set(), type_def
        while current is not None:
            pairs.extend(own_attrs(current))
            base = current.get("BaseClass")
            if base is None or base in seen:
                break
            seen.add(base)
            current = resolve_base(base, current)
        return pairs

    index: dict = defaultdict(lambda: defaultdict(set))
    for type_def in all_types:
        name = type_def.get("Name", "")
        if "/" not in name:
            continue
        element_qname = name.split("/", 1)[1]
        if not element_qname:  # abstract type, reachable only via BaseClass
            continue
        for local, prefix in with_inherited(type_def):
            index[element_qname][local].add(prefix)
    return index


def declared_attr_name(index, context, attr_local, nsmap, rule_prefix):
    """The attribute name as a real instance writes it, or None if unknown.

    An element can declare both forms of one local name (p:sldMasterId has
    `id` and `r:id`). The schematron's own prefix disambiguates: a foreign
    prefix means the qualified attribute, the element's own prefix is the
    SDK's stand-in for unqualified.
    """
    prefixes = index.get(context, {}).get(attr_local)
    if not prefixes:
        return None
    if len(prefixes) > 1:
        if rule_prefix in prefixes:
            prefix = rule_prefix
        elif "" in prefixes:
            prefix = ""
        else:
            return None
    else:
        prefix = next(iter(prefixes))
    if prefix == "":
        return attr_local
    ns = nsmap.get(prefix)
    return f"{{{ns}}}{attr_local}" if ns else None


def fires_on_violation(constraint, attr_name, tag) -> bool:
    """Run a uniqueness constraint against a genuine duplicate.

    True if it reports. Only valid for UNIQUE_ATTRIBUTE — see EXECUTED_TYPES.
    """
    root = etree.Element("root")
    for _ in range(2):
        etree.SubElement(root, tag).set(attr_name, "SAME")

    context = ValidationContext()
    for probe in root:
        try:
            constraint.validate(probe, context)
        except Exception:
            return False
    return len(context.errors) > 0


def audit():
    index = build_attr_index()
    nsmap = _get_namespace_map()
    results = []

    for rule in get_schematron_registry().get_interpretable_rules():
        rule_type = str(rule.rule_type).split(".")[-1]
        if rule_type not in ATTRIBUTE_BEARING:
            continue
        constraint = create_constraint_from_schematron(rule)
        if constraint is None or getattr(constraint, "attribute", None) is None:
            continue
        if ":" not in rule.context:
            continue

        rule_attr = getattr(rule, "attribute", "") or ""
        prefix = rule_attr.split(":")[0] if ":" in rule_attr else "(none)"
        attr_name = declared_attr_name(index, rule.context, constraint.attribute, nsmap, prefix)
        context_prefix, local = rule.context.split(":", 1)
        ns = nsmap.get(context_prefix)
        if attr_name is None or ns is None:
            results.append((rule_type, rule.context, constraint.attribute, prefix, None, None))
            continue

        # Tier 2: does the constraint look for the attribute as files write it?
        looks_for = (
            f"{{{constraint.namespace}}}{constraint.attribute}"
            if getattr(constraint, "namespace", None)
            else constraint.attribute
        )
        matches = looks_for == attr_name

        # Tier 1: only where a correct violation probe exists.
        live = None
        if rule_type in EXECUTED_TYPES:
            live = fires_on_violation(constraint, attr_name, f"{{{ns}}}{local}")

        results.append((rule_type, rule.context, constraint.attribute, prefix, live, matches))
    return results


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--list-dead", action="store_true", help="print every inert constraint")
    args = parser.parse_args()

    results = audit()
    resolved = [r for r in results if r[5] is not None]
    unresolved = [r for r in results if r[5] is None]

    live = [r for r in results if r[4] is True]
    dead = [r for r in results if r[4] is False]

    print("=== TIER 1: executed (UNIQUE_ATTRIBUTE) ===")
    print("Constraint run against a genuine duplicate. Silent -> inert.")
    print(f"  live       {len(live)}")
    print(f"  dead       {len(dead)}")

    print("\n=== TIER 2: declarative (all attribute-bearing types) ===")
    print("Does the constraint look for the attribute as real files write it?")
    print(f"  {'type':28} {'ok':>6} {'mismatch':>9}")
    by_type: dict = defaultdict(lambda: [0, 0])
    for rule_type, _, _, _, _, matches in resolved:
        by_type[rule_type][0 if matches else 1] += 1
    total_mismatch = 0
    for rule_type in sorted(by_type, key=lambda k: -by_type[k][1]):
        n_ok, n_bad = by_type[rule_type]
        total_mismatch += n_bad
        print(f"  {rule_type:28} {n_ok:6} {n_bad:9}")
    print(f"  {'TOTAL':28} {len(resolved) - total_mismatch:6} {total_mismatch:9}")
    print(f"  unresolved (not counted): {len(unresolved)}")

    mismatch_by_prefix = Counter(r[3] for r in resolved if not r[5])
    print("\n--- mismatched by schematron attribute prefix ---")
    for prefix, count in mismatch_by_prefix.most_common(12):
        print(f"  {prefix:8} {count}")

    if args.list_dead:
        print("\n--- every inert UNIQUE_ATTRIBUTE constraint (tier 1) ---")
        for _, context, attr, _, _, _ in sorted(dead):
            print(f"  {context}/@{attr}")

    print("\n=== HARNESS CHECK (tier 1) ===")
    print("w:/r: attributes are genuinely namespaced, so they must be LIVE.")
    dead_by_prefix = Counter(r[3] for r in dead)
    live_by_prefix = Counter(r[3] for r in live)
    bad = {p: dead_by_prefix[p] for p in ("w", "r") if dead_by_prefix[p]}
    if bad:
        print(f"  FAIL — classified DEAD but must be LIVE: {bad}")
        print("  The harness is wrong; the numbers above are not trustworthy.")
        return 1
    print(f"  PASS — w={live_by_prefix['w']} live, r={live_by_prefix['r']} live, 0 dead")

    # Tier 1 and tier 2 must agree where they overlap, or one of them is wrong.
    t1_dead = len(dead)
    t2_mismatch_unique = by_type["UNIQUE_ATTRIBUTE"][1]
    print("\n=== CROSS-VALIDATION ===")
    if t1_dead != t2_mismatch_unique:
        print(f"  FAIL — tier1 dead={t1_dead} != tier2 mismatch={t2_mismatch_unique}")
        return 1
    print(f"  PASS — both tiers agree on UNIQUE_ATTRIBUTE ({t1_dead})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
