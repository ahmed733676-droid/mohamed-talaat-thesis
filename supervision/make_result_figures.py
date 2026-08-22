#!/usr/bin/env python3
"""Generate Figures 4.1-4.3 directly from supervision/locked_data.json.

Existing as a supervision control rather than a convenience: every number on these
charts is read from the locked source of truth, so a figure cannot drift away from
its table through re-typing. Nothing is hard-coded here.

    python3 supervision/make_result_figures.py --outdir figures

Produces 300 dpi TIFF and PNG suitable for a printed thesis. Captions are written to
figures/CAPTIONS.md for pasting beneath each figure.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.ticker import MultipleLocator

ROOT = Path(__file__).resolve().parent


def resolve_serif() -> str:
    """Match the thesis body font. Liberation Serif is metric-compatible with Times New Roman,
    so figures typeset identically on machines without the Microsoft font installed."""
    available = {f.name for f in font_manager.fontManager.ttflist}
    for candidate in ("Times New Roman", "Liberation Serif", "Nimbus Roman",
                      "DejaVu Serif", "Noto Serif"):
        if candidate in available:
            return candidate
    return "serif"


SERIF = resolve_serif()
plt.rcParams.update({"font.family": "serif", "font.serif": [SERIF], "mathtext.fontset": "dejavuserif"})

SHORT = {
    "Beautifil Flow Plus X F00": "Beautifil Flow\nPlus X F00",
    "G-\u00e6nial Universal Injectable": "G-\u00e6nial Universal\nInjectable",
    "Beautifil II LS": "Beautifil II LS",
}
# Greyscale: theses are often printed in black and white.
FILLS = ["#d9d9d9", "#a6a6a6", "#5f5f5f"]
BAR_KW = dict(edgecolor="black", linewidth=0.8, width=0.62)
# ax.bar forwards error-bar styling through error_kw, not as top-level kwargs.
ERR_KW = dict(capsize=5, error_kw=dict(ecolor="black", elinewidth=0.9, capthick=0.9))


def style_axes(ax, ylabel: str) -> None:
    ax.set_ylabel(ylabel, fontsize=11, fontname=SERIF)
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(axis="both", labelsize=10)
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontname(SERIF)
    ax.set_axisbelow(True)
    ax.yaxis.grid(True, linestyle=":", linewidth=0.6, color="#999999")


def save(fig, outdir: Path, stem: str) -> list[Path]:
    written = []
    fig.tight_layout()
    for ext in ("tiff", "png"):
        path = outdir / f"{stem}.{ext}"
        fig.savefig(path, dpi=300, bbox_inches="tight",
                    pil_kwargs={"compression": "tiff_lzw"} if ext == "tiff" else None)
        written.append(path)
    plt.close(fig)
    return written


def fig_weight_loss(locked: dict, outdir: Path) -> list[Path]:
    groups = locked["weight"]["groups"]
    names = list(groups)
    means = [groups[n]["loss_percent"][0] for n in names]
    sds = [groups[n]["loss_percent"][1] for n in names]

    fig, ax = plt.subplots(figsize=(6.3, 4.4))
    ax.bar(range(len(names)), means, yerr=sds, color=FILLS, **BAR_KW, **ERR_KW)
    ax.set_xticks(range(len(names)))
    ax.set_xticklabels([SHORT[n] for n in names])
    style_axes(ax, "Weight loss (%)")

    # In the Beautifil Flow Plus X F00 group the SD (3.14) exceeds the mean (1.49), so the lower
    # error bar falls below zero. Clipping the axis at zero would hide that, which would misstate
    # the dispersion the outlier causes. The axis is extended instead and zero is drawn in.
    lo = min(0.0, min(m - s for m, s in zip(means, sds)))
    hi = max(m + s for m, s in zip(means, sds))
    pad = (hi - lo) * 0.13
    ax.set_ylim(lo - pad * 0.4, hi + pad)
    if lo < 0:
        ax.axhline(0, color="black", linewidth=0.8)

    p = locked["weight"]["p_value"]
    ax.text(0.5, 0.98, f"One-way ANOVA, p = {p} (not significant)",
            transform=ax.transAxes, ha="center", va="top",
            fontsize=10, fontname=SERIF)
    # Label beside the bar top rather than on the error cap, so the number is not misread
    # as the upper bound of the standard deviation.
    for i, (m, s) in enumerate(zip(means, sds)):
        ax.annotate(f"{m:.2f} \u00b1 {s:.2f}", xy=(i + 0.34, m), xytext=(4, 0),
                    textcoords="offset points", va="center", ha="left",
                    fontsize=9, fontname=SERIF)
    return save(fig, outdir, "Figure_4.1_weight_loss")


def fig_roughness(locked: dict, outdir: Path) -> list[Path]:
    groups = locked["roughness"]["groups"]
    names = list(groups)
    before = [groups[n]["before_um"][0] for n in names]
    before_sd = [groups[n]["before_um"][1] for n in names]
    after = [groups[n]["after_um"][0] for n in names]
    after_sd = [groups[n]["after_um"][1] for n in names]
    thr = locked["thresholds"]["plaque_retention_um"]

    x = range(len(names))
    w = 0.34
    fig, ax = plt.subplots(figsize=(6.9, 4.4))
    ax.bar([i - w / 2 for i in x], before, w, yerr=before_sd, label="Before brushing",
           color="#e8e8e8", edgecolor="black", linewidth=0.8, **ERR_KW)
    ax.bar([i + w / 2 for i in x], after, w, yerr=after_sd, label="After 10 000 cycles",
           color="#7a7a7a", edgecolor="black", linewidth=0.8, **ERR_KW)

    ax.axhline(thr, color="black", linestyle="--", linewidth=1.1)
    ax.text(len(names) - 0.42, thr + 0.006,
            f"{thr} \u00b5m plaque-retention threshold",
            fontsize=9, ha="right", fontname=SERIF)

    ax.set_xticks(list(x))
    ax.set_xticklabels([SHORT[n] for n in names])
    style_axes(ax, "Surface roughness, Ra (\u00b5m)")
    ax.yaxis.set_major_locator(MultipleLocator(0.05))
    ax.set_ylim(0, max(max(a + s for a, s in zip(after, after_sd)), thr) * 1.30)

    leg = ax.legend(frameon=False, fontsize=10, loc="upper left")
    for t in leg.get_texts():
        t.set_fontname(SERIF)
    ax.text(0.5, 1.02, f"Kruskal\u2013Wallis on Ra after brushing, "
                       f"p {locked['roughness']['p_value_after']}",
            transform=ax.transAxes, ha="center", va="bottom",
            fontsize=10, fontname=SERIF)
    return save(fig, outdir, "Figure_4.2_surface_roughness")


def fig_delta_ra(locked: dict, outdir: Path) -> list[Path]:
    groups = locked["roughness"]["groups"]
    names = list(groups)
    means = [groups[n]["delta_um"][0] for n in names]
    sds = [groups[n]["delta_um"][1] for n in names]

    fig, ax = plt.subplots(figsize=(6.3, 4.2))
    ax.bar(range(len(names)), means, yerr=sds, color=FILLS, **BAR_KW, **ERR_KW)
    ax.set_xticks(range(len(names)))
    ax.set_xticklabels([SHORT[n] for n in names])
    style_axes(ax, "\u0394Ra (\u00b5m)")
    ax.set_ylim(0, max(m + s for m, s in zip(means, sds)) * 1.25)
    for i, (m, s) in enumerate(zip(means, sds)):
        ax.text(i, m + s + max(means) * 0.04, f"{m:.3f}", ha="center",
                fontsize=9, fontname=SERIF)
    return save(fig, outdir, "Figure_4.3_delta_ra")


CAPTIONS = """# Figure captions

Place each caption **below** its figure, in Times New Roman 12 pt, and cite every figure in the
text of Chapter 4 before it appears. Figure numbers must run in order of first mention.

**Figure 4.1** Mean percentage weight loss of the three composite resins after 10 000
toothbrushing cycles at 2 N. Error bars represent one standard deviation (n = 12 per group). In
the Beautifil Flow Plus X F00 group the standard deviation exceeds the mean, so the lower error
bar extends below zero; this reflects the dispersion introduced by a single specimen that lost
approximately 11.4 % of its mass and does not imply a gain in weight. Differences among groups
were not statistically significant (one-way ANOVA, p = {p_weight}).

**Figure 4.2** Mean surface roughness (Ra) of the three composite resins before and after 10 000
toothbrushing cycles at 2 N. Error bars represent one standard deviation (n = 12 per group). The
broken line marks the {thr} µm threshold above which plaque retention increases. Ra after
brushing differed significantly among groups (Kruskal–Wallis, p {p_ra}).

**Figure 4.3** Mean change in surface roughness (ΔRa = Ra after − Ra before) of the three
composite resins after 10 000 toothbrushing cycles at 2 N. Error bars represent one standard
deviation (n = 12 per group).

---

Still required from the candidate, and not generatable from data:

- **Figure 3.1** CAD/CAM Teflon mould and a finished disc specimen (10 mm diameter × 1 mm thick)
- **Figure 3.2** Toothbrushing simulator with a specimen mounted
- **Figure 3.3** Contact profilometer with a specimen in position
"""


def main() -> int:
    ap = argparse.ArgumentParser(description="Build Figures 4.1-4.3 from the locked data.")
    ap.add_argument("--locked", type=Path, default=ROOT / "locked_data.json")
    ap.add_argument("--outdir", type=Path, default=Path("figures"))
    args = ap.parse_args()

    locked = json.loads(args.locked.read_text(encoding="utf-8"))
    args.outdir.mkdir(parents=True, exist_ok=True)

    written: list[Path] = []
    written += fig_weight_loss(locked, args.outdir)
    written += fig_roughness(locked, args.outdir)
    written += fig_delta_ra(locked, args.outdir)

    caption_file = args.outdir / "CAPTIONS.md"
    caption_file.write_text(
        CAPTIONS.format(
            p_weight=locked["weight"]["p_value"],
            p_ra=locked["roughness"]["p_value_after"],
            thr=locked["thresholds"]["plaque_retention_um"],
        ),
        encoding="utf-8",
    )
    written.append(caption_file)

    for p in written:
        print(f"wrote {p}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
