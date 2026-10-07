"""Read the HKEX New Listing Reports (Main Board) into one list of IPOs.

Each report lists every Main Board new listing of a calendar year. An IPO
takes one numbered row for the Hong Kong offer, marked (a), and one row
marked with a ditto sign for the international offer, marked (b). Funds
raised is the sum of the two rows. Listings by introduction print "N/A".

HKEX prints no industry column. The only HKEX classification in the file is
the name suffix: B for a Chapter 18A biotech company and P for a
pre-commercial Chapter 18C company. Every other industry tag below is a
keyword screen on the company name and is labelled as such.
"""

from __future__ import annotations

import datetime as dt
import re
from pathlib import Path

from openpyxl import load_workbook

NLR_DIR = Path(__file__).resolve().parents[1] / "data" / "hkex_nlr"
NLR_URL = (
    "https://www2.hkexnews.hk/-/media/HKEXnews/Homepage/New-Listings/"
    "New-Listing-Information/New-Listing-Report/Main/NLR{year}_Eng.xlsx"
)
NLR_YEARS = range(2021, 2027)

# Policy industry -> name keywords. Case-insensitive, matched on word starts.
# A name screen misses companies whose names do not state the industry
# (MiniMax, for example), so each total is a lower bound for that industry.
NAME_SCREEN = {
    "Life and health": r"pharm|bio|therap|medic|health|hospital|genetic|gene\b|diagnos|vaccin|clinic",
    # "Intelligent" and "smart" are left out: in these names they mostly mean
    # property services, control valves, or equipment.
    "AI and robotics": r"\brobo|\bai\b|artificial intelligence|autonomous|\bdeep|unmanned",
    "Semi-conductors": r"semiconductor|chip|microelectronic|integrated circuit|silicon|wafer",
    "New energy": r"energy|batter|hydrogen energy|solar|photovolt|lithium|new power|charging|electric vehicle",
    "Film and creative": r"film|pictures|movie|animation|studio|entertainment|\bmedia\b|\bculture|\bcultural",
    "Aerospace": r"aero|aircraft|aviation|\bspace|satellite|drone",
    "Construction": r"construction|(?<!gene )engineering|building",
}

# Stock code -> (industry, reason). Names that match a keyword but whose
# business, as the name itself states it, is outside the policy industry.
NAME_SCREEN_EXCLUDE = {
    "02370": ("Life and health", "Healthy Living is property management."),
    "02530": ("Life and health", "Health Food is food."),
}


def _clean(name: object) -> str:
    return " ".join(str(name).replace("_x000D_", " ").split())


def _markers(name: str) -> set[str]:
    """Return the HKEX suffix letters, e.g. 'Foo Ltd. - SW - H Shares' -> {S, W}."""
    found: set[str] = set()
    for group in re.findall(r"-\s*([BPSW]{1,3})(?=\s*(?:-|$))", name):
        found.update(group)
    return found


def load_listings() -> list[dict]:
    """Every Main Board IPO from 2021 to the latest report, one dict per IPO."""
    listings: list[dict] = []
    for year in NLR_YEARS:
        path = NLR_DIR / f"NLR{year}.xlsx"
        sheet = load_workbook(path, data_only=True, read_only=True).active
        rows = list(sheet.iter_rows(values_only=True))
        coverage = _clean(rows[0][0])
        current: dict | None = None
        for row in rows[2:]:
            number = row[0]
            is_new = isinstance(number, int) or (isinstance(number, str) and number.strip().isdigit())
            if is_new:
                name = _clean(row[2])
                current = {
                    "report_year": year,
                    "coverage": coverage,
                    "code": str(row[1]).strip(),
                    "name": name,
                    "prospectus": row[3].date() if isinstance(row[3], dt.datetime) else None,
                    "listed": row[4].date() if isinstance(row[4], dt.datetime) else None,
                    "funds_hkd": 0.0,
                    "hk_offer_hkd": 0.0,
                    "intl_offer_hkd": 0.0,
                    "markers": _markers(name),
                    "source": NLR_URL.format(year=year),
                }
                listings.append(current)
            elif not (current and row[3] == '"'):
                # Footnotes and blank rows end the IPO above them.
                current = None
                continue
            if isinstance(row[8], (int, float)):
                current["funds_hkd"] += float(row[8])
                # The numbered row is the Hong Kong offer; ditto rows are the international offer.
                current["hk_offer_hkd" if is_new else "intl_offer_hkd"] += float(row[8])
    for item in listings:
        item["by_introduction"] = item["funds_hkd"] == 0
        item["tags"] = screen(item)
    return listings


def window_rows(listings: list[dict], windows: list[tuple[str, dt.date, dt.date]]) -> list[dict]:
    """IPO count and HK$ million per (window, industry), with the Main Board total of the same window.

    The denominator is the same file's Main Board total for that window, so
    the share never mixes this source with an HKEX annual-statistics total.
    """
    out: list[dict] = []
    for label, start, end in windows:
        inside = [i for i in listings if i["listed"] and start <= i["listed"] <= end]
        board = sum(i["funds_hkd"] for i in inside) / 1e6
        for industry in ["Chapter 18A biotech (B marker)", *NAME_SCREEN]:
            if industry.startswith("Chapter 18A"):
                hits = [i for i in inside if "B" in i["markers"]]
            else:
                hits = [i for i in inside if industry in i["tags"]]
            funds = sum(i["funds_hkd"] for i in hits) / 1e6
            out.append(
                {
                    "window": label,
                    "start": start,
                    "end": end,
                    "industry": industry,
                    "ipos": len(hits),
                    "funds_hkd_million": round(funds, 1),
                    "board_ipos": len(inside),
                    "board_hkd_million": round(board, 1),
                }
            )
    return out


# Companies the name screen misses. HKEX does not print their industry in
# these files; the industry here is the trainee's reading, so these rows are
# shown beside the screen totals and never added into them.
KNOWN_MISSES = {
    "03750": ("New energy", "Battery maker. Named among the largest 2025 IPOs in the Listing Committee Report 2025, p.2, without an industry."),
    "06082": ("Semi-conductors", "GPU designer. Chapter 18C, HKEX 18C Explained."),
    "02577": ("Semi-conductors", "Gallium-nitride chip maker."),
    "02533": ("Semi-conductors", "Automotive AI chip designer. Chapter 18C, HKEX 18C Explained."),
    "00100": ("AI and robotics", "AI model developer. Chapter 18C, P marker in the NLR."),
    "02513": ("AI and robotics", "AI model developer. Chapter 18C, HKEX 18C Explained."),
    "00800": ("AI and robotics", "Autonomous driving. Chapter 18C, HKEX 18C Explained."),
    "02525": ("AI and robotics", "Lidar maker. Named as a homecoming listing in the Listing Committee Report 2025, p.2."),
}


def screen(item: dict) -> dict[str, str]:
    """Policy industry -> how it was assigned ('HKEX B marker' or 'name screen')."""
    tags: dict[str, str] = {}
    lowered = item["name"].lower()
    for industry, pattern in NAME_SCREEN.items():
        excluded = NAME_SCREEN_EXCLUDE.get(item["code"], ("", ""))[0] == industry
        if re.search(pattern, lowered) and not excluded:
            tags[industry] = "name screen"
    if "B" in item["markers"]:
        tags["Life and health"] = "HKEX B marker"
    return tags
