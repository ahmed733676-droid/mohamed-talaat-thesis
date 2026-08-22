#!/usr/bin/env python3
"""Stylometric analysis of the thesis for AI-detection risk.

AI-detection tools (GPTZero, Originality.ai, Turnitin's indicator) score two things:

  perplexity  - how predictable each word is given the ones before it. Machine text is
                unusually predictable.
  burstiness  - variance in sentence length and structure. Human writing alternates long
                complex sentences with short ones. Machine text is unusually uniform.

Low variance is what gets flagged, in either direction: uniformly long padded sentences and
uniformly short clipped sentences both read as machine-generated. This script measures the
signals a detector would measure, so the risk can be located per chapter instead of guessed at.

    python3 supervision/stylometry.py --md Thesis_Complete.md
"""

from __future__ import annotations

import argparse
import re
import statistics as st
from collections import Counter
from pathlib import Path

# Sentence-length coefficient of variation. Human academic prose typically lands 0.45-0.75.
CV_HEALTHY = (0.45, 0.80)
MEAN_HEALTHY = (17.0, 27.0)

BRITISH_US_PAIRS = [
    ("fulfilment", "fulfillment"), ("behaviour", "behavior"), ("colour", "color"),
    ("favourable", "favorable"), ("centre", "center"), ("litre", "liter"),
    ("polymerisation", "polymerization"), ("standardised", "standardized"),
    ("randomised", "randomized"), ("minimising", "minimizing"),
    ("characterisation", "characterization"), ("analyse", "analyze"),
    ("aesthetic", "esthetic"), ("dentine", "dentin"), ("recognised", "recognized"),
    ("utilised", "utilized"), ("summarised", "summarized"), ("emphasised", "emphasized"),
]
# practise/practice are different parts of speech in British English, so they are not a
# convention signal. "Organization" is part of ISO's registered name and must not be changed.

# Register problems: colloquial or non-technical usage in a scholarly text.
COLLOQUIAL = [
    (r"\bstands? proud\b", "stands proud"),
    (r"\bstand up or leave pits\b", "stand up"),
    (r"\bare thin\b|\bis thin\b", "'thin' for 'sparse'"),
    (r"\bsat? on (that|the) line\b", "'on the line'"),
    (r"\bsits? on\b(?! the)", "'sits on' for 'is based on'"),
    (r"\brun beside\b", "'run beside'"),
    (r"\bshed more mass\b|\bshedding\b", "'shed' for 'lose'"),
    (r"\bare sold\b|\bis sold\b|\bsold on that claim\b", "'sold' for 'marketed'"),
    (r"\bwere written to\b", "'written' applied to a product"),
    (r"\bfine enough\b", "'fine enough' for 'sufficiently sensitive'"),
    (r"\bin a straight line\b", "'in a straight line' for 'linearly'"),
    (r"\bthe roughness result is cleaner\b", "'cleaner'"),
    (r"\bevery group roughened\b", "'roughened' as intransitive verb"),
    (r"\bafter the run\b", "'the run'"),
    (r"\bdiscs? a group\b", "'a group' for 'per group'"),
    (r"\bthat is the result\b", "'that is the result'"),
    (r"\bis not a clinical lifetime\b", "vague"),
    (r"\bthe files\b", "'the files'"),
]

# Terminology errors specific to this thesis.
TERMINOLOGY = [
    (r"\bthree pastes\b", "'pastes' applied to all three materials, two of which the thesis "
                          "defines as injectables rather than pastes"),
    (r"\bcurrent commercial pastes\b", "'pastes' applied to the injectables"),
    (r"\bX-generation\b", "'X-generation' is not established terminology"),
    (r"\bthose three factors\b", "'those three factors' has no three-item antecedent"),
    (r"\bIFU\b", "'IFU' used without definition and absent from the abbreviations list"),
    (r"\bClass II box\b", "abrupt undefined clinical reference"),
]

# Sentence openers that repeat mechanically are a burstiness signal.
def opener(sentence: str) -> str:
    words = re.findall(r"[A-Za-z\u00e6\u2019'-]+", sentence)
    return " ".join(words[:2]).lower() if words else ""


def sentences_of(text: str) -> list[str]:
    prose = "\n".join(
        ln for ln in text.splitlines()
        if not ln.strip().startswith(("|", "#", ">", "-", "*"))
        and not re.match(r"^\s*\d+\.\s", ln)
    )
    prose = re.sub(r"\*\*(.+?)\*\*", r"\1", prose)
    prose = re.sub(r"\s+", " ", prose)
    # Protect decimals, initials and abbreviations from the sentence splitter.
    prose = re.sub(r"(\d)\.(\d)", r"\1<DOT>\2", prose)
    prose = re.sub(r"\b([A-Z])\.", r"\1<DOT>", prose)
    prose = re.sub(r"\b(et al|vs|cf|e\.g|i\.e|approx|Dr|Prof|No|Fig)\.", r"\1<DOT>", prose)
    parts = re.split(r"(?<=[.!?])\s+", prose)
    return [p.replace("<DOT>", ".").strip() for p in parts if len(p.strip()) > 1]


def words_of(text: str) -> list[str]:
    return re.findall(r"\b[\w\u2019'-]+\b", text.lower())


def chapters_of(md: str) -> dict[str, str]:
    lines = md.splitlines()
    heads = [(i, m.group(2).strip()) for i, ln in enumerate(lines)
             if (m := re.match(r"^(#{1,2})\s+(.*\S)\s*$", ln))]
    out: dict[str, str] = {}
    for idx, (i, title) in enumerate(heads):
        if re.fullmatch(r"Chapter \d+", title):
            continue
        end = heads[idx + 1][0] if idx + 1 < len(heads) else len(lines)
        out[title] = out.get(title, "") + "\n" + "\n".join(lines[i + 1:end])
    return out


def analyse(name: str, text: str) -> dict | None:
    sents = sentences_of(text)
    if len(sents) < 12:
        return None
    lengths = [len(words_of(s)) for s in sents]
    mean = st.mean(lengths)
    sd = st.pstdev(lengths)
    cv = sd / mean if mean else 0.0
    openers = Counter(opener(s) for s in sents if opener(s))
    repeated = [(o, c) for o, c in openers.most_common(6) if c >= 3]
    return {
        "name": name,
        "n": len(sents),
        "mean": mean,
        "sd": sd,
        "cv": cv,
        "median": st.median(lengths),
        "shortest": min(lengths),
        "longest": max(lengths),
        "pct_le6": 100 * sum(1 for n in lengths if n <= 6) / len(lengths),
        "pct_ge30": 100 * sum(1 for n in lengths if n >= 30) / len(lengths),
        "repeated_openers": repeated,
        "mattr": mattr(words_of(text)),
    }


def mattr(words: list[str], window: int = 400) -> float:
    """Moving-average type-token ratio. Plain TTR falls as text lengthens, so it cannot be
    compared against a fixed range; averaging over fixed-size windows removes that bias."""
    if len(words) <= window:
        return len(set(words)) / max(1, len(words))
    ratios = [len(set(words[i:i + window])) / window
              for i in range(0, len(words) - window + 1, max(1, window // 4))]
    return sum(ratios) / len(ratios)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--md", required=True, type=Path)
    args = ap.parse_args()
    md = args.md.read_text(encoding="utf-8")

    print("=" * 84)
    print("STYLOMETRIC / AI-DETECTION RISK ANALYSIS")
    print("=" * 84)

    chapters = chapters_of(md)
    targets = ["Abstract", "Introduction", "Review of literature", "Materials and methods",
               "Results", "Discussion", "Conclusions and recommendations"]

    print("\nBURSTINESS  (sentence-length variation -- the primary detector signal)")
    print(f"{'Section':<34}{'sents':>6}{'mean':>7}{'SD':>7}{'CV':>7}{'<=6w':>7}{'>=30w':>7}{'range':>10}")
    print("-" * 84)
    rows = []
    for t in targets:
        body = next((v for k, v in chapters.items() if k.lower() == t.lower()), None)
        if body is None:
            continue
        r = analyse(t, body)
        if not r:
            continue
        rows.append(r)
        print(f"{t[:32]:<34}{r['n']:>6}{r['mean']:>7.1f}{r['sd']:>7.1f}{r['cv']:>7.2f}"
              f"{r['pct_le6']:>6.0f}%{r['pct_ge30']:>6.0f}%"
              f"{str(r['shortest']) + '-' + str(r['longest']):>10}")

    body_all = "\n".join(next((v for k, v in chapters.items() if k.lower() == t.lower()), "")
                         for t in targets)
    overall = analyse("WHOLE THESIS", body_all)
    print("-" * 84)
    if overall:
        print(f"{'WHOLE THESIS':<34}{overall['n']:>6}{overall['mean']:>7.1f}{overall['sd']:>7.1f}"
              f"{overall['cv']:>7.2f}{overall['pct_le6']:>6.0f}%{overall['pct_ge30']:>6.0f}%"
              f"{str(overall['shortest']) + '-' + str(overall['longest']):>10}")

    print(f"\nHealthy ranges: mean {MEAN_HEALTHY[0]:.0f}-{MEAN_HEALTHY[1]:.0f} words, "
          f"CV {CV_HEALTHY[0]:.2f}-{CV_HEALTHY[1]:.2f}, <=6w under 10%, >=30w at least 8%")

    print("\nVERDICT PER SECTION")
    print("-" * 84)
    for r in rows:
        flags = []
        if r["mean"] < MEAN_HEALTHY[0]:
            flags.append(f"mean {r['mean']:.1f} below {MEAN_HEALTHY[0]:.0f}")
        if r["cv"] < CV_HEALTHY[0]:
            flags.append(f"CV {r['cv']:.2f} below {CV_HEALTHY[0]:.2f} (uniform = machine-like)")
        if r["pct_le6"] > 10:
            flags.append(f"{r['pct_le6']:.0f}% fragments")
        if r["pct_ge30"] < 8:
            flags.append(f"only {r['pct_ge30']:.0f}% complex sentences")
        verdict = "HIGH RISK" if len(flags) >= 3 else "AT RISK" if flags else "acceptable"
        print(f"  {r['name']:<32} {verdict}")
        for f in flags:
            print(f"      - {f}")
        if r["repeated_openers"]:
            print(f"      - repeated openers: "
                  + ", ".join(f"'{o}' x{c}" for o, c in r["repeated_openers"]))

    print("\nSPELLING CONVENTION  (mixed en-GB / en-US is a machine-translation artifact)")
    print("-" * 84)
    # Reference titles must be reproduced exactly as published, so the reference list is
    # excluded: an American spelling inside a cited title is correct, not an inconsistency.
    body_only = md
    if (rm := re.search(r"^#{1,2}\s*(chapter 7\s*)?#*\s*references\s*$", md, re.I | re.M)):
        body_only = md[:rm.start()]
    low = body_only.lower()
    gb, us = [], []
    for b, a in BRITISH_US_PAIRS:
        nb, na = len(re.findall(rf"\b{b}\b", low)), len(re.findall(rf"\b{a}\b", low))
        if nb:
            gb.append(f"{b} x{nb}")
        if na:
            us.append(f"{a} x{na}")
    print(f"  British forms: {', '.join(gb) or 'none'}")
    print(f"  American forms: {', '.join(us) or 'none'}")
    if gb and us:
        print("  MIXED -- pick one convention (British, to match the body) and apply it throughout.")

    print("\nREGISTER  (colloquial usage in a scholarly text)")
    print("-" * 84)
    for patterns, heading in ((COLLOQUIAL, "colloquial"), (TERMINOLOGY, "terminology")):
        seen = set()
        for pat, label in patterns:
            for m in re.finditer(pat, md, re.I):
                ln = md.count("\n", 0, m.start()) + 1
                if (label, ln) in seen:
                    continue
                seen.add((label, ln))
                frag = re.sub(r"\s+", " ", md[max(0, m.start() - 55):m.end() + 55])
                print(f"  [{heading}] line {ln}: {label}")
                print(f"      ...{frag.strip()}...")

    print("\nLEXICAL DIVERSITY")
    print("-" * 84)
    if overall:
        print(f"  MATTR (400-word window): {overall['mattr']:.3f}")
        print("  Length-independent, so comparable across drafts. Track it during the rewrite:")
        print("  it should hold or rise. A falling MATTR while the word count climbs means the")
        print("  added length is restatement rather than new content.")
    print("=" * 84)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
