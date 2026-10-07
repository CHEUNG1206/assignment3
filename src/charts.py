"""PNG charts for the PDFs and the markdown cards.

Every value plotted is passed in by build_briefing.py, which reads it from
the New Listing Reports or the constants that carry a source URL. Nothing is
computed here except quarterly sums of those listings.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

BLUE = "#1F4E79"
GREY = "#8C8C8C"
PALETTE = ["#1F4E79", "#2E86AB", "#A23B72", "#F18F01", "#C73E1D", "#3B8B5A", "#6C5B7B", "#8C8C8C"]


def _quarter_start(day: dt.date) -> dt.date:
    return dt.date(day.year, 3 * ((day.month - 1) // 3) + 1, 1)


def industry_quarters(
    listings: list[dict],
    industries: list[tuple[str, callable]],
    delivered: dict[str, str],
    rule_dates: list[tuple[str, dt.date]],
    first: dt.date,
    last: dt.date,
    path: Path,
) -> None:
    """Small multiples: quarterly IPO funds per policy industry, with address dates as dashed lines."""
    quarters: list[dt.date] = []
    q = _quarter_start(first)
    while q <= last:
        quarters.append(q)
        q = dt.date(q.year + (q.month + 3 > 12), (q.month + 2) % 12 + 1, 1)
    fig, axes = plt.subplots(len(industries), 1, figsize=(7.2, 1.55 * len(industries)), sharex=True)
    for ax, (label, test) in zip(axes, industries):
        sums = {k: 0.0 for k in quarters}
        for item in listings:
            if item["listed"] and first <= item["listed"] <= last and test(item):
                sums[_quarter_start(item["listed"])] += item["funds_hkd"] / 1e9
        ax.bar(quarters, [sums[k] for k in quarters], width=70, color=BLUE, align="edge")
        for year, day in delivered.items():
            d = dt.date.fromisoformat(day)
            ax.axvline(d, color="#C73E1D", lw=0.8, ls="--")
        for _name, d in rule_dates:
            ax.axvline(d, color="#3B8B5A", lw=0.8, ls=":")
        ax.set_ylabel("HK$ bn", fontsize=7)
        ax.set_title(label, fontsize=8, loc="left")
        ax.tick_params(labelsize=7)
        ax.spines[["top", "right"]].set_visible(False)
    top = axes[0]
    for year, day in delivered.items():
        top.annotate(year, (dt.date.fromisoformat(day), top.get_ylim()[1]), fontsize=6, color="#C73E1D", ha="left", va="top")
    fig.text(
        0.01,
        0.005,
        "Red dashed: Policy Address delivery. Green dotted: Chapter 18C effective 31 Mar 2023, TECH launched 6 May 2025.",
        fontsize=6.5,
        color=GREY,
    )
    fig.tight_layout(rect=(0, 0.02, 1, 1))
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=200)
    plt.close(fig)


def window_shares(windows: list[str], series: dict[str, list[float]], path: Path) -> None:
    """Grouped bars: each industry's share of Main Board funds, window by window."""
    fig, ax = plt.subplots(figsize=(7.2, 3.0))
    n = len(series)
    width = 0.8 / n
    for k, (label, values) in enumerate(series.items()):
        xs = [i + k * width for i in range(len(windows))]
        ax.bar(xs, [v * 100 for v in values], width=width, label=label, color=PALETTE[k % len(PALETTE)])
    ax.set_xticks([i + 0.4 - width / 2 for i in range(len(windows))])
    ax.set_xticklabels(windows, fontsize=7)
    ax.set_ylabel("% of Main Board IPO funds in the window", fontsize=7)
    ax.tick_params(labelsize=7)
    ax.legend(fontsize=6.5, ncol=3, frameon=False)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=200)
    plt.close(fig)


def c18_controls(years: list[str], funds_bn: list[float], base_rate: list[float | None], southbound: list[float | None], path: Path) -> None:
    """Chapter 18C funds by year on top; the Base Rate and Southbound ADT for the same years below."""
    fig, (ax, low) = plt.subplots(2, 1, figsize=(7.2, 3.6), sharex=True, gridspec_kw={"height_ratios": [3, 2]})
    ax.bar(years, funds_bn, color=BLUE)
    for x, v in zip(years, funds_bn):
        ax.annotate(f"{v:.1f}", (x, v), ha="center", va="bottom", fontsize=6.5)
    ax.set_ylabel("18C funds, HK$ bn", fontsize=7)
    ax.axvline(0.5, color="#3B8B5A", ls=":", lw=1)
    ax.text(0.53, max(funds_bn) * 0.55, "Chapter 18C\neffective\n31 Mar 2023", fontsize=6.5, color="#3B8B5A")
    ax.set_ylim(0, max(funds_bn) * 1.15)
    ax.tick_params(labelsize=7)
    ax.spines[["top", "right"]].set_visible(False)
    pts = [(y, s) for y, s in zip(years, southbound) if s is not None]
    low.plot([p[0] for p in pts], [p[1] for p in pts], color="#F18F01", marker="o", label="Southbound ADT, HK$ bn (left)")
    low.set_ylabel("HK$ bn", fontsize=7)
    low.tick_params(labelsize=7)
    rate_ax = low.twinx()
    rate = [(y, r * 100) for y, r in zip(years, base_rate) if r is not None]
    rate_ax.plot([p[0] for p in rate], [p[1] for p in rate], color="#C73E1D", marker="s", ls="--", label="Base Rate at year end, % (right)")
    rate_ax.set_ylim(3, 6.5)
    rate_ax.tick_params(labelsize=7)
    handles = low.get_legend_handles_labels()[0] + rate_ax.get_legend_handles_labels()[0]
    labels = low.get_legend_handles_labels()[1] + rate_ax.get_legend_handles_labels()[1]
    low.legend(handles, labels, fontsize=6.5, frameon=False, loc="upper left")
    low.axvline(0.5, color="#3B8B5A", ls=":", lw=1)
    low.spines[["top"]].set_visible(False)
    fig.text(0.01, 0.005, "2026 Q1: listings to 31 Mar 2026. Base Rate and Southbound ADT for 2026 are not stated in source.", fontsize=6.5, color=GREY)
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=200)
    plt.close(fig)


def before_after(industries: list[str], before: list[float], after: list[float], labels: tuple[str, str], path: Path) -> None:
    """Paired bars for one address: the 12 months before and the window after."""
    fig, ax = plt.subplots(figsize=(7.2, 2.6))
    xs = range(len(industries))
    ax.bar([x - 0.2 for x in xs], before, width=0.4, color=GREY, label=labels[0])
    ax.bar([x + 0.2 for x in xs], after, width=0.4, color=BLUE, label=labels[1])
    ax.set_xticks(list(xs))
    ax.set_xticklabels(industries, fontsize=6.5, rotation=15)
    ax.set_ylabel("IPO funds, HK$ bn", fontsize=7)
    ax.tick_params(labelsize=7)
    ax.legend(fontsize=6.5, frameon=False)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=200)
    plt.close(fig)


def draft_counts(metrics: list[str], before: list[int], after: list[int], path: Path) -> None:
    """Before and after counts for the better-writing run."""
    fig, ax = plt.subplots(figsize=(7.2, 2.4))
    xs = range(len(metrics))
    ax.barh([x + 0.2 for x in xs], before, height=0.4, color=GREY, label="influence_before.md")
    ax.barh([x - 0.2 for x in xs], after, height=0.4, color=BLUE, label="influence_after.md")
    ax.set_yticks(list(xs))
    ax.set_yticklabels(metrics, fontsize=7)
    ax.tick_params(labelsize=7)
    for x, (b, a) in enumerate(zip(before, after)):
        ax.annotate(str(b), (b, x + 0.2), fontsize=6, va="center", xytext=(2, 0), textcoords="offset points")
        ax.annotate(str(a), (a, x - 0.2), fontsize=6, va="center", xytext=(2, 0), textcoords="offset points")
    ax.legend(fontsize=6.5, frameon=False, loc="lower center", bbox_to_anchor=(0.5, 1.0), ncol=2)
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_xlim(0, max(before + after) * 1.12)
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=200)
    plt.close(fig)
