#!/usr/bin/env python3
"""Build the Policy Address funding card, the audit workbook, the charts, and the PDFs.

Amounts are the ones read from the English Policy Address XML and the HKEX
pages on 2026-10-07. Industry outcomes come from every Main Board IPO in the
HKEX New Listing Reports (data/hkex_nlr), counted from each delivery date to
the day before the next address. The workbook stores the printed amounts and
the listing rows as inputs; every total, share, and conversion is a formula.
"""

from __future__ import annotations

import datetime as dt
import re
import shutil
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

import charts
from nlr import KNOWN_MISSES, NAME_SCREEN, NAME_SCREEN_EXCLUDE, NLR_URL, load_listings, window_rows

ROOT = Path(__file__).resolve().parents[1]
A1 = ROOT / "A1_CHEUNG_Yuk_Yuen"
A2 = ROOT / "A2_CHEUNG_Yuk_Yuen"
A3 = ROOT / "A3_CHEUNG_Yuk_Yuen"
FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
XML = "https://www.ceo.gov.hk/public/open-data/en/policy_address/pa{year}.xml"
HKEX_2022 = (
    "https://www.hkex.com.hk/-/media/HKEX-Market/Market-Data/Statistics/"
    "Consolidated-Reports/Annual-Market-Statistics/e_2022-Market-Statistics.pdf"
)
HKEX_2023 = (
    "https://www.hkex.com.hk/-/media/HKEX-Market/Market-Data/Statistics/"
    "Consolidated-Reports/Annual-Market-Statistics/FY_2023-Annual-Market-Stat_Eng.pdf"
)
HKEX_2024 = (
    "https://www.hkex.com.hk/-/media/HKEX-Market/Market-Data/Statistics/"
    "Consolidated-Reports/Annual-Market-Statistics/2024FY-Ann-Mkt-Stat_eng.pdf"
)
HKEX_2025 = (
    "https://www.hkex.com.hk/-/media/HKEX-Market/News/News-Release/2026/"
    "260226news/260226news_eng.pdf"
)
HKEX_HL = (
    "https://www.hkex.com.hk/Market-Data/Statistics/Consolidated-Reports/"
    "HKEX-Monthly-Market-Highlights?sc_lang=en"
)
HKEX_2024_RESULTS = (
    "https://www.hkex.com.hk/-/media/HKEX-Market/News/News-Release/2025/"
    "250227news/250227_FY-Results_e.pdf"
)
LC = (
    "https://www.hkex.com.hk/-/media/HKEX-Market/Listing/How-We-Regulate/"
    "Listing-Committee/Listing-Committee-Report/AnnualRpt_{year}dec.pdf"
)
SHEET = "https://docs.google.com/spreadsheets/d/1Nr4XOrgmN3wJ44HHLVwDYQaRkt7oi6D_qz0EIGy9lJ0/edit"

NEW = "new in this address"
EARLIER = "already in place before this address"

# year, programme, amount, unit, industries, source, note
# unit is the unit in the English Policy Address. HK$ billion and HK$ million
# both convert. "as printed" would not.
FUNDING = [
    ("2022", "Co-Investment Fund", 30, "HK$ billion", "Enterprises setting up in Hong Kong", XML.format(year=2022), "Set aside from the Future Fund."),
    ("2022", "RAISe+ Scheme", 10, "HK$ billion", "University research teams with potential to become start-ups", XML.format(year=2022), "Earmarked for launch the next year, on a matching basis, for at least 100 teams."),
    ("2022", "Strategic Tech Fund", 5, "HK$ billion", "Life and health technology, AI and data science, advanced manufacturing and new energy technology", XML.format(year=2022), "Used with OASES. The goal of attracting at least 100 I&T enterprises is a target, not a second government sum."),
    ("2023", "New Industrialisation Acceleration Scheme", 10, "HK$ billion", "Life and health technologies, AI and data science, advanced manufacturing, new energy technologies", XML.format(year=2023), "Matching of 1 (Government) to 2 (company), funding ceiling HK$200 million."),
    ("2023", "Film Development Fund and CreateSmart Initiative", 4.3, "HK$ billion", "Film, cultural and creative industries", XML.format(year=2023), "One injection shared by the two schemes."),
    ("2024", "Research Matching Grant Scheme", 1.5, "HK$ billion", "Research at institutions", XML.format(year=2024), "Paragraph 78. A new round."),
    ("2024", "I&T Industry-Oriented Fund", 10, "HK$ billion", "Life and health technology, AI and robotics, semi-conductors and smart devices, advanced materials, new energy", XML.format(year=2024), "Fund-of-funds. The 2026 address says it is expected to launch in 2026."),
    ("2024", "Innovation and Technology Venture Fund", 1.5, "HK$ billion", "Start-ups in strategic industries", XML.format(year=2024), "Redeployed money, set up jointly with the market, on a matching basis."),
    ("2024", "I&T Accelerator Pilot Scheme", 180, "HK$ million", "Start-up accelerators", XML.format(year=2024), "Paragraph 80. Matching 1 to 2, subsidy ceiling HK$30 million."),
    ("2024", "New Energy Transport Fund", 750, "HK$ million", "Electric taxis and franchised buses, hydrogen fuel-cell heavy vehicles", XML.format(year=2024), "Paragraph 85 says around HK$750 million."),
    ("2025", "Frontier Technology Research Support Scheme", 3, "HK$ billion", "AI and other frontier research", XML.format(year=2025), "Paragraph 68. Announced earlier. Applications to open soon."),
    ("2025", "Hong Kong AI Research and Development Institute", 1, "HK$ billion", "AI research, transfer, and use cases", XML.format(year=2025), "Paragraph 69. Earmarked for establishment in 2026."),
    ("2025", "Aerospace Technology special call", 100, "HK$ million", "Aerospace technology", XML.format(year=2025), "The sentence says over HK$100 million for six university projects. The cell stores 100."),
    ("2026", "Innovation and Technology Venture Fund, enhanced scheme", 1, "HK$ billion", "I&T venture funds", XML.format(year=2026), "An additional earmark. It is not the HK$1.5 billion redeployed in 2024."),
    ("2026", "Artificial Intelligence Subsidy Scheme", 1, "HK$ billion", "High-performance computing for the AI institute", XML.format(year=2026), "Paragraph 201."),
    ("2026", "AI for Welfare Lab", 300, "HK$ million", "AI in social welfare services", XML.format(year=2026), "Paragraph 113. WelfareHK.ai."),
    ("2026", "Construction Innovation and Technology Fund", 1, "HK$ billion", "I&T in the construction industry", XML.format(year=2026), "Paragraph 333. The Construction Industry Council's HK$400 million is not government funding."),
]

# (year, programme) -> (status, the verb clause as printed in the English XML)
FUNDING_STATUS = {
    ("2022", "Co-Investment Fund"): (NEW, "set aside $30 billion from the Future Fund to establish the Co-Investment Fund"),
    ("2022", "RAISe+ Scheme"): (NEW, "We will earmark $10 billion to launch"),
    ("2022", "Strategic Tech Fund"): (EARLIER, "making use of the $5 billion Strategic Tech Fund"),
    ("2023", "New Industrialisation Acceleration Scheme"): (NEW, "set up a $10 billion New Industrialisation Acceleration Scheme"),
    ("2023", "Film Development Fund and CreateSmart Initiative"): (NEW, "We will also inject a total of $4.3 billion"),
    ("2024", "Research Matching Grant Scheme"): (NEW, "The Government will launch a new round of Research Matching Grant Scheme totalling $1.5 billion"),
    ("2024", "I&T Industry-Oriented Fund"): (NEW, "setting up a $10 billion I&T Industry-Oriented Fund"),
    ("2024", "Innovation and Technology Venture Fund"): (NEW, "We will redeploy $1.5 billion"),
    ("2024", "I&T Accelerator Pilot Scheme"): (NEW, "The Government will launch the I&T Accelerator Pilot Scheme with a funding allocation of $180 million"),
    ("2024", "New Energy Transport Fund"): (NEW, "The Government will earmark around $750 million"),
    ("2025", "Frontier Technology Research Support Scheme"): (EARLIER, "The Government earlier announced the launch of the $3 billion"),
    ("2025", "Hong Kong AI Research and Development Institute"): (EARLIER, "We have also earmarked $1 billion"),
    ("2025", "Aerospace Technology special call"): (EARLIER, "has allocated over $100 million"),
    ("2026", "Innovation and Technology Venture Fund, enhanced scheme"): (EARLIER, "An additional $1 billion has been earmarked"),
    ("2026", "Artificial Intelligence Subsidy Scheme"): (NEW, "The Government will inject $1 billion"),
    ("2026", "AI for Welfare Lab"): (NEW, "A sum of $300 million will be allocated"),
    ("2026", "Construction Innovation and Technology Fund"): (EARLIER, "We injected $1 billion into the Construction Innovation and Technology Fund (CITF) this year"),
}

# Outcome-year HKEX rows. year is the market year, which is the year after the address.
# kind distinguishes IPO from ADT so the summary formulas can select each.
OUTCOMES = [
    ("2023", "IPO", 46294.9, "HK$ million", "-55.8%", "29 Dec 2023", HKEX_2023, "Market Statistics 2023, p.3. Main Board and GEM. Year after the 19 Oct 2022 address."),
    ("2023", "ADT", 105012, "HK$ million", "-15.9%", "29 Dec 2023", HKEX_2023, "Market Statistics 2023, p.3. Average daily turnover."),
    ("2024", "IPO", 87477.8, "HK$ million", "+88.9%", "31 Dec 2024", HKEX_2024, "Market Statistics 2024, p.3. Main Board and GEM. Year after the 25 Oct 2023 address."),
    ("2024", "ADT", 131825, "HK$ million", "+25.5%", "31 Dec 2024", HKEX_2024, "Market Statistics 2024, p.3. Average daily turnover."),
    ("2025", "IPO", 286.9, "HK$ billion", "+226%", "31 Dec 2025", HKEX_2025, "2025 Final Results, p.11. 119 IPOs. Page 17 prints 286.8 and is not averaged in."),
    ("2025", "ADT", 249.8, "HK$ billion", "+90%", "31 Dec 2025", HKEX_2025, "2025 Final Results, p.11. Headline ADT. Year after the 16 Oct 2024 address."),
    ("2026", "IPO", 342.4, "HK$ billion", "+153%", "31 Aug 2026", HKEX_HL, "First eight months only, versus HK$135.1 billion a year earlier. Part of the year after the 17 Sep 2025 address."),
    ("2026", "ADT", 282.5, "HK$ billion", "+14%", "31 Aug 2026", HKEX_HL, "First eight months, versus HK$248.3 billion a year earlier."),
]

DELIVERED = {
    "2022": "2022-10-19",
    "2023": "2023-10-25",
    "2024": "2024-10-16",
    "2025": "2025-09-17",
    "2026": "2026-09-16",
}

# address year -> outcome year, window label, label
WINDOWS = [
    ("2022", "2023", "2023 full year", "following year recorded, source does not attribute it to the address"),
    ("2023", "2024", "2024 full year", "following year recorded, source does not attribute it to the address"),
    ("2024", "2025", "2025 full year", "following year recorded, source does not attribute it to the address"),
    ("2025", "2026", "2026, first eight months", "following year only partly published"),
    ("2026", "", "2027 not published", "following year not yet published"),
]

# The 2026 New Listing Report runs "up to 29 September 2026". load_listings() checks the heading.
NLR_END = dt.date(2026, 9, 29)


def delivery_windows() -> list[tuple[str, dt.date, dt.date]]:
    """Baseline year before the first address, then delivery date to the day before the next one."""
    days = [dt.date.fromisoformat(DELIVERED[y]) for y in sorted(DELIVERED)]
    out = [("Before 2022", days[0].replace(year=days[0].year - 1), days[0] - dt.timedelta(days=1))]
    for k, year in enumerate(sorted(DELIVERED)):
        end = days[k + 1] - dt.timedelta(days=1) if k + 1 < len(days) else NLR_END
        out.append((f"After {year}", days[k], end))
    return out


INDUSTRIES = ["Chapter 18A biotech (B marker)", *NAME_SCREEN]

# Which funding sentences name each industry, copied from the FUNDING industries column.
NAMED_IN = {
    "Chapter 18A biotech (B marker)": "Life and health in 2022, 2023, 2024",
    "Life and health": "2022 Strategic Tech Fund, 2023 NIAS, 2024 Industry-Oriented Fund",
    "AI and robotics": "AI and data science 2022 and 2023; AI and robotics 2024; AI 2025 and 2026",
    "Semi-conductors": "2024 Industry-Oriented Fund only",
    "New energy": "2022, 2023, 2024 (and the 2024 New Energy Transport Fund)",
    "Film and creative": "2023 Film Development Fund and CreateSmart",
    "Aerospace": "2025 special call, already allocated",
    "Construction": "2026 CITF, already injected",
}

# Chapter 18C listings from HKEX "18C, Explained", as of end-March 2026.
# Funds are HK$ million, as printed. The page also states a combined HK$28.4 billion.
C18_URL = (
    "https://www.hkexgroup.com/Media-Centre/Insight/Insight/2026/"
    "HKEX-Insight/18C-Explained?sc_lang=en"
)
C18_LISTINGS = [
    ("XtalPi Holdings Ltd.", "2024-06-13", 2024, 1036),
    ("Black Sesame International Holdings Ltd.", "2024-08-08", 2024, 1036),
    ("Shenzhen Dobot Corp. Ltd.", "2024-12-23", 2024, 831),
    ("Beijing Yunji Technology Co., Ltd.", "2025-10-16", 2025, 759),
    ("Deepexi Technology Co., Ltd.", "2025-10-28", 2025, 710),
    ("WeRide Inc.", "2025-11-06", 2025, 2392),
    ("CiDi Inc.", "2025-12-19", 2025, 1422),
    ("Beijing 51World Digital Twin Technology Co., Ltd.", "2025-12-30", 2025, 731),
    ("Shanghai Biren Technology Co., Ltd.", "2026-01-02", 2026, 6420),
    ("Knowledge Atlas Technology JSC Ltd.", "2026-01-08", 2026, 5000),
    ("MiniMax Group Inc.", "2026-01-09", 2026, 5540),
    ("Guangdong Huayan Robotics Co., Ltd.", "2026-03-30", 2026, 1579),
    ("Shandong Extreme Vision Technology Co., Ltd.", "2026-03-30", 2026, 499),
    ("Shanghai FourSemi Semiconductor Co., Ltd.", "2026-03-31", 2026, 480),
]
# year, base rate at year end, southbound ADT HK$ billion, board IPO HK$ million, note
# 2026 is only the first quarter of listings. Rate, southbound and board IPO for that quarter are not on these pages.
CONTROLS = [
    (2022, 0.0475, 31.7, 104570.3, "Before Chapter 18C. Base Rate end-2022. Southbound ADT. IPO from Market Statistics 2022."),
    (2023, 0.0575, 31.1, 46294.9, "Chapter effective 31 Mar 2023. No 18C listing this year. Base Rate end-2023 was 5.75%."),
    (2024, 0.0475, 48.2, 87477.8, "Three 18C listings. Base Rate cut from 5.75% to 4.75%. Southbound ADT HK$48.2 billion."),
    (2025, 0.0400, 121.1, 285800.0, "Five 18C listings. Base Rate cut from 4.75% to 4.00%. Southbound ADT HK$121.1 billion. Board IPO is the Listing Committee figure."),
    (2026, None, None, None, "Listings through 31 Mar 2026 only. A full-year Base Rate, Southbound ADT and board IPO are not stated in source."),
]
HKMA_2022 = "https://www.hkma.gov.hk/media/eng/publication-and-research/annual-report/2022/14_Monetary_Stability.pdf"
HKMA_2024 = "https://www.hkma.gov.hk/media/eng/publication-and-research/annual-report/2024/14_Monetary_Stability.pdf"
HKMA_2025 = "https://www.hkma.gov.hk/media/eng/publication-and-research/annual-report/2025/AR_2025E.pdf"
SOUTHBOUND_URL = (
    "https://www.hkexgroup.com/Media-Centre/Insight/Insight/2026/"
    "HKEX-Insight/Stock-Connect-2025-Review?sc_lang=en"
)

ATTRIBUTION_2024 = (
    "While the global IPO market continued to be impacted by high interest rates and slow economic growth, "
    "Hong Kong's IPO market saw a significant recovery in the second half of the year. The increased confidence "
    "was driven by several sizeable deals, including, on 17 September 2024, the largest IPO in the city since 2021. "
    "It was given further momentum by the announcement of stimulus measures in Mainland China in the later part of the year."
)
ATTRIBUTION_2025 = (
    "Hong Kong led a rebound in the global IPO market in 2025, with the surge of listings being further accelerated by "
    "global demand for diversified growth opportunities, the evolution of frontier technologies in the region and local market reforms."
)

MARKET_WORDING = {
    "2022": "The address says Hong Kong has an active stock market. It states no IPO total and no turnover.",
    "2023": "Paragraph 71 accepts the recommendations of the Task Force on Enhancing Stock Market Liquidity. It states no IPO total.",
    "2024": "The financial-market paragraphs set a direction for artificial intelligence, tokenisation, and stablecoins. They state no IPO total.",
    "2025": "Paragraph 92, delivered on 17 September, says the Hang Seng Index had risen by over 20% year to date, average daily turnover had reached HK$250 billion, and IPO funds by the end of August exceeded HK$130 billion. Those figures pre-date delivery.",
    "2026": "Paragraph 36, delivered on 16 September, says IPO funds as of the end of August exceeded HK$340 billion. The HKEX highlights page dates that total to 31 August.",
}


def hkd_million_formula(row: int) -> str:
    """Convert the printed unit on this Inputs row into HK$ million."""
    return (
        f'=IF(E{row}="HK$ billion",D{row}*1000,'
        f'IF(E{row}="HK$ million",D{row},'
        f'IF(E{row}="億元",D{row}*100,"")))'
    )


def to_million(amount: float, unit: str) -> float:
    return amount * 1000 if unit == "HK$ billion" else amount


def m(value: float) -> str:
    """HK$ million with one decimal and thousands separators, as the tables print it."""
    return f"{value:,.1f}"


def pct(value: float) -> str:
    return f"{value * 100:.1f}%"


class Data:
    """Every computed figure the prose uses, built once from the inputs above."""

    def __init__(self) -> None:
        self.listings = load_listings()
        heading = next(i["coverage"] for i in self.listings if i["report_year"] == 2026)
        assert "29 September 2026" in heading, heading
        self.windows = delivery_windows()
        self.rows = window_rows(self.listings, self.windows)
        self.cell = {(r["window"], r["industry"]): r for r in self.rows}
        self.board = {r["window"]: r for r in self.rows if r["industry"] == INDUSTRIES[0]}
        last = self.windows[-1][2]
        self.cumulative = window_rows(
            self.listings, [(f"From {y} delivery", dt.date.fromisoformat(DELIVERED[y]), last) for y in sorted(DELIVERED)]
        )
        self.cum = {(r["window"], r["industry"]): r for r in self.cumulative}
        self.new_total = {y: 0.0 for y in DELIVERED}
        self.earlier_total = {y: 0.0 for y in DELIVERED}
        for year, item, amount, unit, *_rest in FUNDING:
            status = FUNDING_STATUS[(year, item)][0]
            bucket = self.new_total if status == NEW else self.earlier_total
            bucket[year] += to_million(amount, unit)
        self.by_code = {i["code"]: i for i in self.listings}
        self.misses = []
        for code, (industry, reason) in KNOWN_MISSES.items():
            item = self.by_code[code]
            window = next(w for w, s, e in self.windows if s <= item["listed"] <= e)
            self.misses.append((code, item, industry, window, reason))
        self.first_listing = {}
        start = self.windows[0][1]
        for industry in INDUSTRIES:
            hits = sorted(i["listed"] for i in self.listings if i["listed"] >= start and self.tagged(i, industry))
            self.first_listing[industry] = hits[0] if hits else None
        self.year_rows = {}
        for year in range(2021, 2027):
            inside = [i for i in self.listings if i["listed"].year == year]
            self.year_rows[year] = {
                "board": sum(i["funds_hkd"] for i in inside) / 1e6,
                "board_ipos": len(inside),
                **{ind: sum(i["funds_hkd"] for i in inside if self.tagged(i, ind)) / 1e6 for ind in INDUSTRIES},
                "b_count": sum("B" in i["markers"] for i in inside),
            }

    @staticmethod
    def tagged(item: dict, industry: str) -> bool:
        if industry.startswith("Chapter 18A"):
            return "B" in item["markers"]
        return industry in item["tags"]

    def funds(self, window: str, industry: str) -> float:
        return self.cell[(window, industry)]["funds_hkd_million"]

    def ipos(self, window: str, industry: str) -> int:
        return self.cell[(window, industry)]["ipos"]

    def share(self, window: str, industry: str) -> float:
        board = self.board[window]["board_hkd_million"]
        return self.funds(window, industry) / board if board else 0.0

    def miss(self, code: str) -> tuple:
        return next(x for x in self.misses if x[0] == code)


def influence_texts(d: Data) -> tuple[str, str]:
    """The rough draft and the better-writing rewrite. Every figure is filled from Data."""
    semis = d.funds("After 2025", "Semi-conductors")
    semi_n = d.ipos("After 2025", "Semi-conductors")
    semi_share = pct(d.share("After 2025", "Semi-conductors"))
    semi_first = d.first_listing["Semi-conductors"]
    ai = d.funds("After 2025", "AI and robotics")
    ai_shares = [d.share(w, "AI and robotics") for w in ("After 2023", "After 2024", "After 2025")]
    bio = d.funds("After 2025", INDUSTRIES[0])
    bio_before = pct(d.share("Before 2022", INDUSTRIES[0]))
    bio_after = pct(d.share("After 2025", INDUSTRIES[0]))
    board_before = m(d.board["Before 2022"]["board_hkd_million"])
    board_after = m(d.board["After 2025"]["board_hkd_million"])
    film_after = sum(d.funds(w, "Film and creative") for w in ("After 2023", "After 2024", "After 2025"))
    film_n = sum(d.ipos(w, "Film and creative") for w in ("After 2023", "After 2024", "After 2025"))
    aero_after = d.ipos("After 2025", "Aerospace")
    catl = d.miss("03750")
    ne_window = catl[3]
    ne_with = d.funds(ne_window, "New energy") + catl[1]["funds_hkd"] / 1e6
    ne_with_share = pct(ne_with / d.board[ne_window]["board_hkd_million"])
    days_2026 = (d.windows[-1][2] - d.windows[-1][1]).days + 1

    # Claims the prose makes in words, checked against the numbers.
    assert all(d.ipos(w, "Semi-conductors") == 0 for w in ("Before 2022", "After 2022", "After 2023", "After 2024"))
    assert bio == max(d.funds(w, INDUSTRIES[0]) for w, _s, _e in d.windows)
    assert film_n == 1 and aero_after == 0 and d.ipos("After 2026", "Construction") == 0

    before = f"""
It is important to note that the Policy Addresses from 2022 to 2026 have played a pivotal role in shaping Hong Kong's IPO landscape. Semi-conductors, which the 2024 address named beside the HK$10 billion I&T Industry-Oriented Fund, had no Main Board IPO by the name screen in the four windows from 19 October 2021 to 16 September 2025, and then {semi_n} semiconductor companies listed in the window after the 2025 address, starting on {semi_first:%-d %B %Y} and raising HK${m(semis)} million, {semi_share} of the Main Board, which is a testament to the power of targeted policy. However, it should be noted that the 2026 address says the fund is expected to launch in 2026, so the money had not been paid out when these companies listed. AI and robotics has also seen remarkable growth, going from no IPO before the 2022 address to HK${m(ai)} million after the 2025 address, although its share of the Main Board stayed between {pct(min(ai_shares))} and {pct(max(ai_shares))} in the three windows after the 2023 address. Life and health technology remains a key pillar, with Chapter 18A biotech companies raising HK${m(bio)} million after the 2025 address, the most in any window, even though its share fell from {bio_before} before the 2022 address to {bio_after}. Meanwhile, the Main Board as a whole rose from HK${board_before} million to HK${board_after} million over the same windows, and HKEX attributes the 2025 rebound to global demand for diversified growth opportunities, frontier technologies in the region, and local market reforms, and does not mention the Policy Address. The HK$4.3 billion for film and CreateSmart was followed by only one film or creative IPO, HK${m(film_after)} million, in the three windows after the 2023 address, the aerospace allocation of over HK$100 million by no aerospace IPO after the 2025 address, and the HK$1 billion construction injection by {days_2026} days of data with no construction IPO. It is also worth mentioning that the name screen misses Contemporary Amperex Technology, and adding its HK${m(catl[1]['funds_hkd'] / 1e6)} million would make new energy {ne_with_share} of the window after the 2024 address. Taken together, these findings underscore that policy and the market are deeply interconnected, and that the government's vision is not merely words but a driving force behind the market's transformation.
""".strip()

    after = f"""
Semi-conductors had no Main Board IPO by the name screen in the four windows from 19 October 2021 to 16 September 2025. In the window after the 2025 address, {semi_n} listed, starting on {semi_first:%-d %B %Y}, and raised HK${m(semis)} million, {semi_share} of the Main Board. The 2024 address named semi-conductors beside the HK$10 billion I&T Industry-Oriented Fund. The 2026 address says that fund is expected to launch in 2026, so the money had not been paid out when these companies listed.

AI and robotics went from no IPO before the 2022 address to HK${m(ai)} million after the 2025 address. Its share of the Main Board stayed between {pct(min(ai_shares))} and {pct(max(ai_shares))} in the three windows after the 2023 address. Chapter 18A biotech raised HK${m(bio)} million after the 2025 address, the most in any window, while its share fell from {bio_before} before the 2022 address to {bio_after}.

Over the same windows the Main Board rose from HK${board_before} million to HK${board_after} million. HKEX attributes the 2025 rebound to global demand for diversified growth opportunities, frontier technologies in the region, and local market reforms. It does not mention the Policy Address.

The HK$4.3 billion for film and CreateSmart was followed by only one film or creative IPO, HK${m(film_after)} million, in the three windows after the 2023 address. No aerospace IPO followed the 2025 address, which cited an allocation of over HK$100 million. The HK$1 billion construction injection has {days_2026} days of data after it and no construction IPO.

The name screen misses Contemporary Amperex Technology. Adding its HK${m(catl[1]['funds_hkd'] / 1e6)} million would make new energy {ne_with_share} of the window after the 2024 address.
""".strip()
    return before, after


TELL_PHRASES = [
    "it is important to note",
    "pivotal role",
    "landscape",
    "testament to",
    "it should be noted",
    "remarkable",
    "key pillar",
    "meanwhile",
    "it is also worth mentioning",
    "taken together",
    "underscore",
    "deeply interconnected",
    "not merely",
    "driving force",
    "transformation",
]


def text_metrics(text: str) -> dict[str, int]:
    words = re.findall(r"[A-Za-z0-9$&%.,'-]+", text)
    sentences = [s for s in re.split(r"(?<=[.])\s+(?=[A-Z])", text) if s.strip()]
    lowered = text.lower()
    figures = re.findall(r"\d[\d,]*(?:\.\d+)?%?", text)
    return {
        "Words": len(words),
        "Sentences": len(sentences),
        "Longest sentence, words": max(len(s.split()) for s in sentences),
        "Stock phrases": sum(lowered.count(p) for p in TELL_PHRASES),
        "Figures": len(figures),
    }


def figures_in(text: str) -> set[str]:
    return set(re.findall(r"\d[\d,]*(?:\.\d+)?%?", text))


def chart_paths() -> dict[str, Path]:
    out = A1 / "outputs" / "charts"
    return {
        "quarters": out / "industry_quarterly.png",
        "shares": out / "industry_window_shares.png",
        "c18": out / "c18_before_after.png",
        "run2": A2 / "outputs" / "run2_2022_before_after.png",
        "draft": A3 / "outputs" / "better_writing_counts.png",
    }


def write_charts(d: Data, before: str, after: str) -> None:
    paths = chart_paths()
    plotted = ["Chapter 18A biotech (B marker)", "Life and health", "AI and robotics", "Semi-conductors", "New energy"]
    charts.industry_quarters(
        d.listings,
        [(name, (lambda ind: (lambda item: Data.tagged(item, ind)))(name)) for name in plotted],
        DELIVERED,
        [("Chapter 18C", dt.date(2023, 3, 31)), ("TECH", dt.date(2025, 5, 6))],
        d.windows[0][1],
        NLR_END,
        paths["quarters"],
    )
    labels = [w for w, _s, _e in d.windows[:-1]]
    charts.window_shares(labels, {ind: [d.share(w, ind) for w in labels] for ind in plotted}, paths["shares"])
    c18_by_year = {y: sum(f for _n, _d, yy, f in C18_LISTINGS if yy == y) / 1000 for y in (2022, 2023, 2024, 2025, 2026)}
    charts.c18_controls(
        ["2022", "2023", "2024", "2025", "2026 Q1"],
        [c18_by_year[y] for y in (2022, 2023, 2024, 2025, 2026)],
        [c[1] for c in CONTROLS],
        [c[2] for c in CONTROLS],
        paths["c18"],
    )
    run2_inds = [i for i in INDUSTRIES if i not in ("Aerospace",)]
    charts.before_after(
        [i.replace(" (B marker)", "") for i in run2_inds],
        [d.funds("Before 2022", i) / 1000 for i in run2_inds],
        [d.funds("After 2022", i) / 1000 for i in run2_inds],
        ("19 Oct 2021 to 18 Oct 2022", "19 Oct 2022 to 24 Oct 2023"),
        paths["run2"],
    )
    mb, ma = text_metrics(before), text_metrics(after)
    charts.draft_counts(list(mb), list(mb.values()), [ma[k] for k in mb], paths["draft"])


def window_label(d: Data, window: str) -> str:
    _w, start, end = next(x for x in d.windows if x[0] == window)
    return f"{start:%-d %b %Y} to {end:%-d %b %Y}"


def industry_table_md(d: Data) -> list[str]:
    heads = [w for w, _s, _e in d.windows]
    lines = ["| Policy industry | Tag basis | " + " | ".join(heads) + " |", "| --- | --- | " + " | ".join("---" for _ in heads) + " |"]
    for ind in INDUSTRIES:
        basis = "HKEX B marker" if ind.startswith("Chapter 18A") else "name screen"
        cells = [f"{d.ipos(w, ind)} / {m(d.funds(w, ind))} / {pct(d.share(w, ind))}" for w in heads]
        lines.append(f"| {ind} | {basis} | " + " | ".join(cells) + " |")
    lines.append("| Main Board, all IPOs | NLR rows | " + " | ".join(f"{d.board[w]['board_ipos']} / {m(d.board[w]['board_hkd_million'])} / 100%" for w in heads) + " |")
    return lines


def write_markdown(d: Data, after: str) -> None:
    a1 = A1 / "outputs"
    a2 = A2 / "outputs"
    a1.mkdir(parents=True, exist_ok=True)
    a2.mkdir(parents=True, exist_ok=True)

    lines = [
        "# Funding and the market after each address",
        "",
        "Prepared for CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics.",
        "Period: John Lee took office on 1 July 2022. Policy Addresses from 19 October 2022 to 16 September 2026. Market sources read on 7 October 2026.",
        "",
        "The question is what the Government said it would fund, and what happened on HKEX to the industries it named after that address. The board total is context. Each industry is counted from every Main Board IPO in the HKEX New Listing Reports.",
        "",
        "A funding row is kept only when one sentence names an amount and commits, earmarks, injects, redeploys, or sets up a fund for an industry, technology, or cultural-industry programme. The verb decides whether the money is new in that address.",
        "",
        "## What the address said about the market",
        "",
        "These sentences describe the market up to delivery. They are context, not the outcome.",
        "",
    ]
    for year in sorted(DELIVERED):
        lines.append(f"- {year}, delivered {DELIVERED[year]}. {MARKET_WORDING[year]}")

    lines += ["", "## Stated government amounts", ""]
    lines.append("| Year | Programme | Amount | Status | Verb as printed | Industries |")
    lines.append("| --- | --- | --- | --- | --- | --- |")
    for year, item, amount, unit, industry, url, note in FUNDING:
        status, verb = FUNDING_STATUS[(year, item)]
        lines.append(f"| {year} | {item} | {amount} {unit} | {status} | \"{verb}\" | {industry} |")
    lines += ["", "Source for every row: the English XML for that year, " + XML.format(year="YEAR") + ". Notes per row are in the workbook, sheet Inputs.", ""]
    lines.append("| Address | New in this address, HK$ million | Already in place, HK$ million |")
    lines.append("| --- | --- | --- |")
    for year in sorted(DELIVERED):
        lines.append(f"| {year} | {m(d.new_total[year])} | {m(d.earlier_total[year])} |")
    lines += [
        "",
        "The 2025 address states no new amount for an industry. Its three amounts were announced, earmarked, or allocated before 17 September 2025.",
        "",
        "## HKEX board in the following calendar year",
        "",
        "| Address | Delivered | Outcome window | Label |",
        "| --- | --- | --- | --- |",
    ]
    for address, _outcome, window, label in WINDOWS:
        lines.append(f"| {address} | {DELIVERED[address]} | {window} | {label} |")
    lines += ["", "| Market year | Indicator | Value | Unit | Printed change | As of | Source |", "| --- | --- | --- | --- | --- | --- | --- |"]
    for year, kind, amount, unit, change, as_of, url, note in OUTCOMES:
        lines.append(f"| {year} | {kind} | {amount} | {unit} | {change} | {as_of} | {url} |")
    lines += [
        "",
        "2027 is not stated in source. These totals are context for the industry rows. A funding amount is not divided by them, because a government fund and the money investors pay for new shares are different sums.",
        "",
        "## Every IPO, by the industry the address named",
        "",
        f"Source: HKEX New Listing Reports, Main Board, {NLR_URL.format(year='2021')} through NLR2026_Eng.xlsx. Copies are in `data/hkex_nlr/`. The 2026 report runs up to 29 September 2026. Funds raised is the Hong Kong offer row plus the international offer row. Listings by introduction raise nothing and count as IPOs with zero funds.",
        "",
        "HKEX prints no industry column. Its only industry tags are the B marker for Chapter 18A biotech and the P marker for pre-commercial Chapter 18C companies. Every other row below is a keyword screen on the company name. It is called a name screen and it is a lower bound.",
        "",
        "Name-screen keywords, matched case-insensitively:",
        "",
    ]
    for ind, pattern in NAME_SCREEN.items():
        lines.append(f"- {ind}: `{pattern}`")
    lines += ["", "Removed by hand, with the reason the name gives:", ""]
    for code, (ind, reason) in NAME_SCREEN_EXCLUDE.items():
        lines.append(f"- {code} {d.by_code[code]['name']}, from {ind}. {reason}")
    lines += [
        "",
        "Windows run from a delivery date to the day before the next address. The first column is the 12 months before the first address. Each cell is IPOs / HK$ million / share of the Main Board total in the same window, from the same file.",
        "",
    ]
    lines += [f"- {w}: {window_label(d, w)}" for w, _s, _e in d.windows]
    lines += [""] + industry_table_md(d)
    lines += [
        "",
        "Life and health includes the Chapter 18A row. The After 2026 window is 14 days long.",
        "",
        "![Quarterly IPO funds by policy industry](charts/industry_quarterly.png)",
        "",
        "![Share of Main Board IPO funds by window](charts/industry_window_shares.png)",
        "",
        "## Companies the name screen misses",
        "",
        "HKEX does not print these companies' industries in the New Listing Reports. The industry is the trainee's reading, so the funds sit beside the totals above and are not added in.",
        "",
        "| Code | Company | Listed | Window | Industry, trainee's reading | HK$ million | Basis |",
        "| --- | --- | --- | --- | --- | --- | --- |",
    ]
    for code, item, ind, window, reason in d.misses:
        lines.append(f"| {code} | {item['name']} | {item['listed']} | {window} | {ind} | {m(item['funds_hkd'] / 1e6)} | {reason} |")
    catl = d.miss("03750")
    ne_w = catl[3]
    ne_with = d.funds(ne_w, "New energy") + catl[1]["funds_hkd"] / 1e6
    lines += [
        "",
        f"The largest effect is new energy. In the window {ne_w.lower().replace('after ', 'after the ')} address the screen finds HK${m(d.funds(ne_w, 'New energy'))} million, {pct(d.share(ne_w, 'New energy'))} of the Main Board. With Contemporary Amperex Technology it would be HK${m(ne_with)} million, {pct(ne_with / d.board[ne_w]['board_hkd_million'])}.",
        "",
        "## Timing",
        "",
        "The New Listing Report prints the prospectus date and the listing date, not the date the application was filed. The preparation time before a listing therefore cannot be measured from it. The table below gives each industry's share from each delivery date to 29 September 2026, so a listing that came late is still counted after the address that named its industry.",
        "",
        "| Industry | First listing in the data | " + " | ".join(f"From {y}" for y in sorted(DELIVERED)) + " |",
        "| --- | --- | " + " | ".join("---" for _ in DELIVERED) + " |",
    ]
    for ind in INDUSTRIES:
        first = d.first_listing[ind]
        cells = []
        for y in sorted(DELIVERED):
            r = d.cum[(f"From {y} delivery", ind)]
            cells.append(f"{m(r['funds_hkd_million'])} ({pct(r['funds_hkd_million'] / r['board_hkd_million'])})")
        lines.append(f"| {ind} | {first or 'none'} | " + " | ".join(cells) + " |")
    lines += [
        "",
        f"First listing in the data counts from {d.windows[0][1]}. Semi-conductors first appears on {d.first_listing['Semi-conductors']}, about 14 months after the 16 October 2024 address that named it, and inside the window after the 2025 address.",
        "",
        "## HKEX figures that disagree",
        "",
        "Each share uses one source for both numerator and denominator. The New Listing Report shares use the same report's Main Board total. Different official totals are kept and not averaged.",
        "",
        "| Topic | Source | Figure |",
        "| --- | --- | --- |",
    ]
    for topic, source, figure in hkex_statements(d):
        lines.append(f"| {topic} | {source} | {figure} |")
    lines += [
        "",
        "## From the day the rule started",
        "",
        "The measure is Chapter 18C, effective 31 March 2023. Listings under that chapter before that day were zero, because the chapter did not exist. The HK$10 billion I&T Industry-Oriented Fund is left out. The 2026 Policy Address says it is expected to launch in 2026, and that page gives no disbursement date.",
        "",
        "Source for every listing: HKEX, 18C Explained, as of end-March 2026. The page states 14 companies and HK$28.4 billion. The rows below are the same table.",
        "",
        "| Company | Listing date | Funds, HK$ million |",
        "| --- | --- | --- |",
    ]
    for name, listed, _year, funds in C18_LISTINGS:
        lines.append(f"| {name} | {listed} | {funds} |")
    c18 = {y: (sum(1 for *_x, yy, _f in C18_LISTINGS if yy == y), sum(f for *_x, yy, f in C18_LISTINGS if yy == y)) for y in range(2022, 2027)}
    lines += [
        "",
        f"Page: {C18_URL}",
        "",
        "| Year | 18C listings | 18C funds, HK$ million | Base Rate, year end | Southbound ADT, HK$ billion | Board IPO, HK$ million | NLR B-marker IPOs |",
        "| --- | --- | --- | --- | --- | --- | --- |",
    ]
    for year, rate, south, board, _note in CONTROLS:
        label = "2026 to 31 Mar" if year == 2026 else str(year)
        b = d.year_rows[year]["b_count"]
        b_text = f"{b} (to 29 Sep)" if year == 2026 else str(b)
        lines.append(
            f"| {label} | {c18[year][0]} | {c18[year][1]:,} | {f'{rate * 100:.2f}%' if rate else 'not stated in source'} | {south if south else 'not stated in source'} | {f'{board:,.1f}' if board else 'not stated in source'} | {b_text} |"
        )
    lines += [
        "",
        f"Base Rate: {HKMA_2022} ; {HKMA_2024} ; {HKMA_2025}",
        f"Southbound: {SOUTHBOUND_URL}",
        "",
        "![Chapter 18C funds with the Base Rate and Southbound ADT](charts/c18_before_after.png)",
        "",
        f"HKEX, 2024 Final Results, p.11: \"{ATTRIBUTION_2024}\"",
        "",
        f"HKEX, 2025 Final Results, p.11: \"{ATTRIBUTION_2025}\"",
        "",
        "Only the 2025 sentence names local market reforms. Neither names the Policy Address.",
        "",
        f"The count leaves zero on 13 June 2024, 14 months after the rule. Of the HK$28,435 million in the table, HK$19,518 million is in the first quarter of 2026, after the Base Rate had been cut to 4.00% and Southbound ADT had reached HK$121.1 billion. Chapter 18A has existed since 2018; its B-marker IPOs in the New Listing Reports were {', '.join(f'{d.year_rows[y]['b_count']} in {y}' for y in range(2021, 2026))}, and {d.year_rows[2026]['b_count']} in 2026 to 29 September.",
        "",
        "## What the figures show",
        "",
        after,
        "",
    ]
    (a1 / "term_funding_card.md").write_text("\n".join(lines), encoding="utf-8")

    run2 = [
        "# Second run: the 2022 Policy Address only",
        "",
        "Trigger: policy-funding-card. Input: the 2022 English XML.",
        "This run does not read the 2026 sentences.",
        "",
        f"Delivered: {DELIVERED['2022']}.",
        f"Market wording in the address: {MARKET_WORDING['2022']}",
        "",
        "| Programme | Amount | Status | Verb as printed | Industries |",
        "| --- | --- | --- | --- | --- |",
    ]
    for year, item, amount, unit, industry, url, note in FUNDING:
        if year == "2022":
            status, verb = FUNDING_STATUS[(year, item)]
            run2.append(f"| {item} | {amount} {unit} | {status} | \"{verb}\" | {industry} |")
    run2 += [
        "",
        f"New in this address: HK${m(d.new_total['2022'])} million. The HK$5 billion Strategic Tech Fund already existed and is not in that total.",
        "",
        "This card does not contain the additional HK$1 billion earmarked for the Innovation and Technology Venture Fund in 2026. That sentence is in the 2026 XML.",
        "",
        f"Outcome window: {window_label(d, 'After 2022')}, the day of delivery to the day before the 2023 address. Baseline: {window_label(d, 'Before 2022')}.",
        "",
        "| Industry named in 2022 | Tag basis | Before: IPOs / HK$ million / share | After: IPOs / HK$ million / share |",
        "| --- | --- | --- | --- |",
    ]
    for ind in ["Chapter 18A biotech (B marker)", "Life and health", "AI and robotics", "New energy"]:
        basis = "HKEX B marker" if ind.startswith("Chapter 18A") else "name screen"
        cells = [f"{d.ipos(w, ind)} / {m(d.funds(w, ind))} / {pct(d.share(w, ind))}" for w in ("Before 2022", "After 2022")]
        run2.append(f"| {ind} | {basis} | {cells[0]} | {cells[1]} |")
    run2.append(
        f"| Main Board | NLR rows | {d.board['Before 2022']['board_ipos']} / {m(d.board['Before 2022']['board_hkd_million'])} / 100% | {d.board['After 2022']['board_ipos']} / {m(d.board['After 2022']['board_hkd_million'])} / 100% |"
    )
    run2 += [
        "",
        "![2022 address: the year before and the window after](run2_2022_before_after.png)",
        "",
        "Advanced manufacturing is also named in 2022. No keyword set separates it from other manufacturers by name, so it is not stated in source.",
        "",
        f"New energy before the address is one IPO, Tianqi Lithium, HK${m(d.funds('Before 2022', 'New energy'))} million on 13 July 2022. After the address the screen finds none. Life and health raised less money after the address and took a larger share, because the Main Board fell from HK${m(d.board['Before 2022']['board_hkd_million'])} million to HK${m(d.board['After 2022']['board_hkd_million'])} million.",
        "",
        "Board context, Market Statistics 2023, p.3: IPO funds HK$46,294.9 million in calendar 2023, down 55.8%; average daily turnover HK$105,012 million, down 15.9%.",
        "",
        "Label: following year recorded, source does not attribute it to the address.",
        "",
    ]
    (a2 / "run2_2022.md").write_text("\n".join(run2), encoding="utf-8")


def hkex_statements(d: Data) -> list[tuple[str, str, str]]:
    """Official figures that disagree or need a definition, each with its page."""
    nlr = d.year_rows
    return [
        ("2022 IPO funds", "Market Statistics 2022, p.3, Main Board and GEM", "HK$104,570.3 million"),
        ("2022 IPO funds", "New Listing Report 2022, Main Board rows summed", f"{nlr[2022]['board_ipos']} IPOs, HK${m(nlr[2022]['board'])} million"),
        ("2023 IPO funds", "Market Statistics 2023, p.3, Main Board and GEM", "HK$46,294.9 million"),
        ("2023 IPO funds", "New Listing Report 2023, Main Board rows summed", f"{nlr[2023]['board_ipos']} IPOs, HK${m(nlr[2023]['board'])} million"),
        ("2024 IPO funds", "Market Statistics 2024, p.3, Main Board and GEM", "HK$87,477.8 million"),
        ("2024 IPO funds", "2024 Final Results, p.11", "71 new listings, HK$88.0 billion"),
        ("2024 IPO funds", "New Listing Report 2024, Main Board rows summed", f"{nlr[2024]['board_ipos']} IPOs, HK${m(nlr[2024]['board'])} million"),
        ("2025 IPO funds", "Listing Committee Report 2025, p.2", "HK$285.8 billion"),
        ("2025 IPO funds", "2025 Final Results, pp.7 and 11", "HK$286.9 billion, 119 IPOs"),
        ("2025 IPO funds", "2025 Final Results, p.17", "HK$286.8 billion"),
        ("2025 IPO funds", "New Listing Report 2025, Main Board rows summed", f"{nlr[2025]['board_ipos']} IPOs, HK${m(nlr[2025]['board'])} million"),
        ("2025 Chapter 18C count", "Listing Committee Report 2025, p.3", "five Specialist Technology Companies by year end"),
        ("2025 Chapter 18C count", "Listing Committee Report 2025, TECH section", "six Main Board Chapter 18C companies from the launch of TECH on 6 May 2025 to 31 December 2025"),
        ("2025 Chapter 18C count", "HKEX 18C Explained", "five listings dated 2025"),
        ("2025 Chapter 18A count", "Listing Committee Report 2025, p.3", "16 Biotech Companies"),
        ("2025 Chapter 18A count", "Listing Committee Report 2025, TECH section", "13 Main Board Chapter 18A companies from 6 May 2025 to 31 December 2025"),
        ("2025 Chapter 18A count", "New Listing Report 2025, B markers", f"{nlr[2025]['b_count']}"),
        ("2025 combined 18A and 18C funds", "Listing Committee Report 2025, p.3", "HK$19.8 billion for the 16 and the five together. Not split."),
        ("2024 new economy", "2024 Final Results, p.11", "almost 82 per cent of the capital raised. HKEX does not define new economy as the Policy Address list, so this is not used for any industry."),
    ]


def write_workbook(path: Path, d: Data) -> None:
    wb = Workbook()
    blue = Font(name="Arial", color="0000FF")
    black = Font(name="Arial", color="000000")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(name="Arial", color="FFFFFF", bold=True)

    def header(ws, names: list[str], height: int = 32) -> None:
        ws.append(names)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = height

    def widths(ws, values: list[int]) -> None:
        for col, width in enumerate(values, start=1):
            ws.column_dimensions[get_column_letter(col)].width = width

    inputs = wb.active
    inputs.title = "Inputs"
    header(inputs, ["year", "kind", "item", "amount_as_printed", "unit", "hkd_million", "source_url", "note", "status", "verb_as_printed"], 22)
    rows: list[tuple] = []
    for year, item, amount, unit, industry, url, note in FUNDING:
        status, verb = FUNDING_STATUS[(year, item)]
        rows.append((year, "funding", f"{item} | {industry}", amount, unit, url, note, status, verb))
    for year, kind, amount, unit, change, as_of, url, note in OUTCOMES:
        rows.append((year, f"outcome_{kind}", kind, amount, unit, url, f"{as_of}. Printed change {change}. {note}", "", ""))
    for offset, row in enumerate(rows, start=2):
        year, kind, item, amount, unit, url, note, status, verb = row
        for col, value in ((1, int(year)), (2, kind), (3, item), (4, amount), (5, unit), (7, url), (9, status), (10, verb)):
            inputs.cell(offset, col, value).font = blue
        cell = inputs.cell(offset, 6, hkd_million_formula(offset))
        cell.font = black
        cell.number_format = "#,##0.0"
        note_cell = inputs.cell(offset, 8, note)
        note_cell.font = blue
        note_cell.alignment = Alignment(wrap_text=True)
        inputs.cell(offset, 4).comment = Comment("Input copied from the source in column G. Do not replace with a calculated total.", "policy-funding-card")
    inputs.auto_filter.ref = f"A1:J{1 + len(rows)}"
    widths(inputs, [10, 16, 55, 18, 14, 16, 36, 55, 32, 60])

    summary = wb.create_sheet("Year_after")
    header(summary, ["address_year", "outcome_year", "board_window", "new_funding_hkd_million", "already_in_place_hkd_million", "outcome_ipo_hkd_million", "outcome_adt_hkd_million"])
    for index, (address, outcome, window, _label) in enumerate(WINDOWS, start=2):
        summary.cell(index, 1, int(address)).font = blue
        summary.cell(index, 2, int(outcome) if outcome else "not stated in source").font = blue
        summary.cell(index, 3, window).font = blue
        summary.cell(index, 4, f'=SUMIFS(Inputs!F:F,Inputs!A:A,A{index},Inputs!B:B,"funding",Inputs!I:I,"{NEW}")').font = black
        summary.cell(index, 5, f'=SUMIFS(Inputs!F:F,Inputs!A:A,A{index},Inputs!B:B,"funding",Inputs!I:I,"{EARLIER}")').font = black
        for col in (4, 5):
            summary.cell(index, col).number_format = "#,##0.0"
        if outcome:
            summary.cell(index, 6, f'=SUMIFS(Inputs!F:F,Inputs!A:A,B{index},Inputs!B:B,"outcome_IPO")').font = black
            summary.cell(index, 7, f'=SUMIFS(Inputs!F:F,Inputs!A:A,B{index},Inputs!B:B,"outcome_ADT")').font = black
            summary.cell(index, 6).number_format = summary.cell(index, 7).number_format = "#,##0.0"
        else:
            summary.cell(index, 6, "not stated in source").font = blue
            summary.cell(index, 7, "not stated in source").font = blue
    summary["A8"] = "New funding counts only rows whose verb is new in that address. The two money columns are not divided by IPO funds."
    summary["A9"] = "2026 IPO and ADT are the first eight months. 2027 is not stated in source."
    widths(summary, [14, 20, 26, 26, 28, 24, 24])

    listings = wb.create_sheet("NLR_listings")
    flag_cols = {"B_marker": "H", "P_marker": "I"}
    names = ["code", "company", "listed", "listing_year", "hk_offer_hkd_million", "intl_offer_hkd_million", "funds_hkd_million", "B_marker", "P_marker"]
    for k, ind in enumerate(NAME_SCREEN):
        flag_cols[ind] = get_column_letter(10 + k)
        names.append(ind)
    names += ["excluded_by_hand", "known_miss_industry", "source_url"]
    header(listings, names)
    for offset, item in enumerate(d.listings, start=2):
        listings.cell(offset, 1, item["code"]).font = blue
        listings.cell(offset, 2, item["name"]).font = blue
        listings.cell(offset, 3, item["listed"]).font = blue
        listings.cell(offset, 3).number_format = "yyyy-mm-dd"
        listings.cell(offset, 4, f"=YEAR(C{offset})").font = black
        listings.cell(offset, 5, round(item["hk_offer_hkd"] / 1e6, 4)).font = blue
        listings.cell(offset, 6, round(item["intl_offer_hkd"] / 1e6, 4)).font = blue
        listings.cell(offset, 7, f"=E{offset}+F{offset}").font = black
        for col in (5, 6, 7):
            listings.cell(offset, col).number_format = "#,##0.0"
        listings.cell(offset, 8, int("B" in item["markers"])).font = blue
        listings.cell(offset, 9, int("P" in item["markers"])).font = blue
        for k, ind in enumerate(NAME_SCREEN):
            listings.cell(offset, 10 + k, int(ind in item["tags"])).font = blue
        base = 10 + len(NAME_SCREEN)
        excl = NAME_SCREEN_EXCLUDE.get(item["code"])
        listings.cell(offset, base, f"{excl[0]}: {excl[1]}" if excl else "").font = blue
        miss = KNOWN_MISSES.get(item["code"])
        listings.cell(offset, base + 1, miss[0] if miss else "").font = blue
        listings.cell(offset, base + 2, item["source"]).font = blue
    last_listing = 1 + len(d.listings)
    listings.auto_filter.ref = f"A1:{get_column_letter(len(names))}{last_listing}"
    widths(listings, [9, 50, 12, 10, 14, 14, 14, 9, 9] + [11] * len(NAME_SCREEN) + [40, 18, 40])

    win = wb.create_sheet("Windows")
    header(win, ["window", "start", "end", "note"])
    for offset, (label, start, end) in enumerate(d.windows, start=2):
        win.cell(offset, 1, label).font = blue
        win.cell(offset, 2, start).font = blue
        win.cell(offset, 3, end).font = blue
        win.cell(offset, 2).number_format = win.cell(offset, 3).number_format = "yyyy-mm-dd"
        note = "12 months before the first address" if offset == 2 else ("Ends on the last day of the 2026 New Listing Report" if offset == 1 + len(d.windows) else "Delivery date to the day before the next address")
        win.cell(offset, 4, note).font = blue
    widths(win, [14, 12, 12, 52])

    ind_ws = wb.create_sheet("Industry_windows")
    header(ind_ws, ["window", "start", "end", "policy_industry", "named_in", "ipos", "funds_hkd_million", "board_ipos", "board_hkd_million", "share_of_board"])
    r = 2
    for w_index, (label, _s, _e) in enumerate(d.windows, start=2):
        for ind in INDUSTRIES:
            col = flag_cols["B_marker"] if ind.startswith("Chapter 18A") else flag_cols[ind]
            ind_ws.cell(r, 1, f"=Windows!A{w_index}").font = black
            ind_ws.cell(r, 2, f"=Windows!B{w_index}").font = black
            ind_ws.cell(r, 3, f"=Windows!C{w_index}").font = black
            ind_ws.cell(r, 2).number_format = ind_ws.cell(r, 3).number_format = "yyyy-mm-dd"
            ind_ws.cell(r, 4, ind).font = blue
            ind_ws.cell(r, 5, NAMED_IN[ind]).font = blue
            rng = f'NLR_listings!$C:$C,">="&B{r},NLR_listings!$C:$C,"<="&C{r}'
            ind_ws.cell(r, 6, f"=COUNTIFS({rng},NLR_listings!${col}:${col},1)").font = black
            ind_ws.cell(r, 7, f"=SUMIFS(NLR_listings!$G:$G,{rng},NLR_listings!${col}:${col},1)").font = black
            ind_ws.cell(r, 8, f"=COUNTIFS({rng})").font = black
            ind_ws.cell(r, 9, f"=SUMIFS(NLR_listings!$G:$G,{rng})").font = black
            ind_ws.cell(r, 10, f'=IF(I{r}=0,"",G{r}/I{r})').font = black
            ind_ws.cell(r, 7).number_format = ind_ws.cell(r, 9).number_format = "#,##0.0"
            ind_ws.cell(r, 10).number_format = "0.0%"
            r += 1
    ind_ws.cell(r + 1, 1, "Industry flags: B marker from HKEX; every other industry is a name screen on the company name and is a lower bound.").font = blue
    ind_ws.cell(r + 2, 1, "Share = this row's funds / the Main Board total of the same window, both from the New Listing Reports.").font = blue
    ind_ws.auto_filter.ref = f"A1:J{r - 1}"
    widths(ind_ws, [14, 12, 12, 30, 44, 8, 18, 10, 18, 14])

    by_year = wb.create_sheet("Industry_by_year")
    header(by_year, ["listing_year", *INDUSTRIES, "main_board_hkd_million", "addresses_delivered"])
    for offset, year in enumerate(range(2021, 2027), start=2):
        by_year.cell(offset, 1, year).font = blue
        for k, ind in enumerate(INDUSTRIES, start=2):
            col = flag_cols["B_marker"] if ind.startswith("Chapter 18A") else flag_cols[ind]
            by_year.cell(offset, k, f"=SUMIFS(NLR_listings!$G:$G,NLR_listings!$D:$D,$A{offset},NLR_listings!${col}:${col},1)").font = black
            by_year.cell(offset, k).number_format = "#,##0"
        board_col = 2 + len(INDUSTRIES)
        by_year.cell(offset, board_col, f"=SUMIFS(NLR_listings!$G:$G,NLR_listings!$D:$D,$A{offset})").font = black
        by_year.cell(offset, board_col).number_format = "#,##0"
        delivered = DELIVERED.get(str(year), "")
        by_year.cell(offset, board_col + 1, delivered or "none").font = blue
    by_year["A9"] = "2026 runs to 29 September. HK$ million. Name-screen columns are lower bounds."
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Main Board IPO funds by policy industry and listing year"
    chart.y_axis.title = "HK$ million"
    chart.x_axis.title = "Listing year (address delivered Sep or Oct of 2022 to 2026)"
    chart.y_axis.majorGridlines = None
    chart.add_data(Reference(by_year, min_col=2, max_col=1 + len(INDUSTRIES), min_row=1, max_row=7), titles_from_data=True)
    chart.set_categories(Reference(by_year, min_col=1, min_row=2, max_row=7))
    chart.height, chart.width = 9, 22
    chart.legend.position = "b"
    by_year.add_chart(chart, "A11")
    widths(by_year, [12] + [16] * len(INDUSTRIES) + [22, 20])

    misses = wb.create_sheet("Known_misses")
    header(misses, ["code", "company", "listed", "window", "industry_trainee_reading", "funds_hkd_million", "basis"])
    for offset, (code, item, ind, window, reason) in enumerate(d.misses, start=2):
        for col, value in ((1, code), (2, item["name"]), (3, item["listed"]), (4, window), (5, ind), (7, reason)):
            misses.cell(offset, col, value).font = blue
        misses.cell(offset, 3).number_format = "yyyy-mm-dd"
        misses.cell(offset, 6, f"=SUMIFS(NLR_listings!$G:$G,NLR_listings!$A:$A,A{offset})").font = black
        misses.cell(offset, 6).number_format = "#,##0.0"
    misses.cell(len(d.misses) + 3, 1, "Not added to Industry_windows. HKEX does not state these industries in the New Listing Reports.").font = blue
    widths(misses, [9, 50, 12, 14, 22, 18, 70])

    stmt = wb.create_sheet("HKEX_statements")
    header(stmt, ["topic", "source", "figure"])
    for offset, row in enumerate(hkex_statements(d), start=2):
        for col, value in enumerate(row, start=1):
            stmt.cell(offset, col, value).font = blue
            stmt.cell(offset, col).alignment = Alignment(wrap_text=True)
    widths(stmt, [28, 50, 70])

    c18_ws = wb.create_sheet("C18_listings")
    header(c18_ws, ["company", "listing_date", "year", "funds_hkd_million", "source_url"], 22)
    for offset, (name, listed, year, funds) in enumerate(C18_LISTINGS, start=2):
        for col, value in ((1, name), (2, listed), (3, year), (4, funds), (5, C18_URL)):
            c18_ws.cell(offset, col, value).font = blue
        c18_ws.cell(offset, 4).number_format = "#,##0"
    last = 1 + len(C18_LISTINGS)
    c18_ws.cell(last + 2, 1, "Page total, HK$ billion, as printed").font = blue
    c18_ws.cell(last + 2, 2, 28.4).font = blue
    c18_ws.cell(last + 3, 1, "Sum of the rows, HK$ million").font = black
    c18_ws.cell(last + 3, 2, f"=SUM(D2:D{last})").font = black
    c18_ws.cell(last + 3, 2).number_format = "#,##0"
    widths(c18_ws, [55, 16, 12, 22, 55])

    rule = wb.create_sheet("C18_before_after")
    header(rule, ["year", "c18_listings", "c18_funds_hkd_million", "base_rate_year_end", "southbound_adt_hkd_billion", "board_ipo_hkd_million", "c18_share_of_board_ipo", "nlr_b_marker_ipos", "note"])
    for offset, (year, rate, south, board, note) in enumerate(CONTROLS, start=2):
        rule.cell(offset, 1, year).font = blue
        rule.cell(offset, 2, f"=COUNTIF(C18_listings!C:C,A{offset})").font = black
        rule.cell(offset, 3, f"=SUMIF(C18_listings!C:C,A{offset},C18_listings!D:D)").font = black
        rule.cell(offset, 3).number_format = "#,##0"
        if rate is None:
            for col in (4, 5, 6, 7):
                rule.cell(offset, col, "not stated in source").font = blue
        else:
            rule.cell(offset, 4, rate).font = blue
            rule.cell(offset, 4).number_format = "0.00%"
            rule.cell(offset, 5, south).font = blue
            rule.cell(offset, 6, board).font = blue
            rule.cell(offset, 6).number_format = "#,##0.0"
            rule.cell(offset, 7, f'=IF(F{offset}="", "", C{offset}/F{offset})').font = black
            rule.cell(offset, 7).number_format = "0.0%"
        rule.cell(offset, 8, f"=SUMIFS(NLR_listings!$H:$H,NLR_listings!$D:$D,A{offset})").font = black
        rule.cell(offset, 9, note).font = blue
        rule.cell(offset, 9).alignment = Alignment(wrap_text=True)
    rule["A8"] = "Chapter 18C effective 31 March 2023. First listing 13 June 2024. 2026 B-marker IPOs run to 29 September."
    rule["A9"] = "The I&T Industry-Oriented Fund has no disbursement date in the 2026 Policy Address. It is not in this table."
    rule["A10"] = "2022 Base Rate: HKMA Annual Report 2022. 2023 and 2024: HKMA Annual Report 2024. 2025: HKMA Annual Report 2025."
    rule["A11"] = "Southbound ADT: HKEX Stock Connect 2025 Review, February 2026."
    widths(rule, [10, 14, 22, 18, 24, 22, 20, 16, 70])

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def styles() -> dict[str, ParagraphStyle]:
    if "WQY" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("WQY", FONT_PATH, subfontIndex=0))
    return {
        "h1": ParagraphStyle("h1", fontName="WQY", fontSize=14, leading=18, textColor=colors.HexColor("#1F4E79"), spaceAfter=6),
        "h2": ParagraphStyle("h2", fontName="WQY", fontSize=11, leading=14, textColor=colors.HexColor("#1F4E79"), spaceBefore=8, spaceAfter=4),
        "body": ParagraphStyle("body", fontName="WQY", fontSize=8.5, leading=12, alignment=TA_LEFT, spaceAfter=4),
        "small": ParagraphStyle("small", fontName="WQY", fontSize=7, leading=9, textColor=colors.HexColor("#333333")),
    }


def pdf_table(data: list[list[str]], col_widths: list[float]) -> Table:
    style = styles()["small"]
    # Paragraph cells ignore the TEXTCOLOR table command, so the header row carries its own white font.
    wrapped = [[Paragraph(f'<font color="white">{cell}</font>', style) for cell in data[0]]]
    wrapped += [[Paragraph(str(cell), style) for cell in row] for row in data[1:]]
    table = Table(wrapped, colWidths=col_widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, -1), "WQY"),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F7F9FB")),
            ]
        )
    )
    return table


def figure(path: Path, width_mm: float = 186) -> Image:
    """Scale a PNG to the text width, keeping its aspect ratio."""
    from reportlab.lib.utils import ImageReader

    w, h = ImageReader(str(path)).getSize()
    return Image(str(path), width=width_mm * mm, height=width_mm * mm * h / w)


def write_pdf(path: Path, story: list) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=12 * mm, rightMargin=12 * mm, topMargin=12 * mm, bottomMargin=12 * mm, title=path.stem)
    doc.build(story)


def industry_pdf_rows(d: Data) -> list[list[str]]:
    heads = [w for w, _s, _e in d.windows]
    rows = [["Industry"] + heads]
    for ind in INDUSTRIES:
        rows.append([ind] + [f"{d.ipos(w, ind)} / {m(d.funds(w, ind))} / {pct(d.share(w, ind))}" for w in heads])
    rows.append(["Main Board"] + [f"{d.board[w]['board_ipos']} / {m(d.board[w]['board_hkd_million'])}" for w in heads])
    return rows


def write_pdfs(d: Data, before: str, after: str) -> None:
    s = styles()
    paths = chart_paths()
    after_paras = [Paragraph(p, s["body"]) for p in after.split("\n\n")]

    report = [
        Paragraph("Assignment 1: Policy Address funding and the industries it named", s["h1"]),
        Paragraph("CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics. 7 October 2026.", s["body"]),
        Paragraph("4. Result", s["h2"]),
        Paragraph(f"Card: A1_CHEUNG_Yuk_Yuen/outputs/term_funding_card.md. Sheet: {SHEET}. Workbook: outputs/term_market.xlsx, sheets Industry_windows and Industry_by_year.", s["body"]),
        *after_paras,
        Paragraph("Funding stated in each address", s["h2"]),
        pdf_table(
            [["Address", "New in this address, HK$ m", "Already in place, HK$ m"]] + [[y, m(d.new_total[y]), m(d.earlier_total[y])] for y in sorted(DELIVERED)],
            [30 * mm, 78 * mm, 78 * mm],
        ),
        Spacer(1, 3 * mm),
        Paragraph("IPOs by policy industry, each window (IPOs / HK$ million / share of the Main Board)", s["h2"]),
        pdf_table(industry_pdf_rows(d), [30 * mm] + [26 * mm] * len(d.windows)),
        Paragraph("Biotech uses HKEX's B marker. The other rows are a name screen and a lower bound. The After 2026 window is 14 days.", s["small"]),
        Spacer(1, 2 * mm),
        figure(paths["quarters"], 120),
        figure(paths["shares"]),
        Paragraph("Companies the name screen misses", s["h2"]),
        pdf_table(
            [["Company", "Listed", "Window", "Industry, trainee's reading", "HK$ m"]]
            + [[item["name"], str(item["listed"]), window, ind, m(item["funds_hkd"] / 1e6)] for _c, item, ind, window, _r in d.misses],
            [70 * mm, 22 * mm, 24 * mm, 42 * mm, 28 * mm],
        ),
        Paragraph("Chapter 18C from its effective date", s["h2"]),
        pdf_table(
            [
                ["Year", "18C listings", "18C funds, HK$ m", "Base Rate", "Southbound ADT", "Board IPO, HK$ m"],
                ["2022", "0", "0", "4.75%", "31.7", "104,570.3"],
                ["2023", "0", "0", "5.75%", "31.1", "46,294.9"],
                ["2024", "3", "2,903", "4.75%", "48.2", "87,477.8"],
                ["2025", "5", "6,014", "4.00%", "121.1", "285,800"],
                ["2026 to 31 Mar", "6", "19,518", "not stated", "not stated", "not stated"],
            ],
            [28 * mm, 24 * mm, 32 * mm, 24 * mm, 32 * mm, 36 * mm],
        ),
        figure(paths["c18"]),
        Paragraph(f"HKEX 2024: \"{ATTRIBUTION_2024}\" HKEX 2025: \"{ATTRIBUTION_2025}\" Only 2025 names local market reforms. Neither names the Policy Address.", s["small"]),
        Paragraph("HKEX figures that disagree", s["h2"]),
        pdf_table([["Topic", "Source", "Figure"]] + [list(r) for r in hkex_statements(d)], [36 * mm, 66 * mm, 84 * mm]),
        Paragraph("Programmes with a stated amount", s["h2"]),
        pdf_table(
            [["Year", "Programme", "Amount", "Status", "Verb as printed"]]
            + [[y, item, f"{amount} {unit}", FUNDING_STATUS[(y, item)][0], FUNDING_STATUS[(y, item)][1]] for y, item, amount, unit, _i, _u, _n in FUNDING],
            [12 * mm, 46 * mm, 22 * mm, 36 * mm, 70 * mm],
        ),
    ]
    # The submitted report is one page (src/submission.py); the full tables and charts go to outputs/.
    result_at = next(k for k, item in enumerate(report) if isinstance(item, Paragraph) and item.getPlainText() == "4. Result")
    analysis = [
        Paragraph("Assignment 1 analysis: Policy Address funding and the industries it named", s["h1"]),
        Paragraph("CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics. 7 October 2026. Companion to report.pdf.", s["body"]),
    ] + report[result_at + 1 :]
    write_pdf(A1 / "outputs" / "analysis.pdf", analysis)



def better_writing_record(before: str, after: str) -> str:
    """The run of better-writing on the findings draft, step by step as SKILL.md orders them."""
    mb, ma = text_metrics(before), text_metrics(after)
    kept = figures_in(after) <= figures_in(before)
    assert kept, figures_in(after) - figures_in(before)
    found = [p for p in TELL_PHRASES if p in before.lower()]
    left = [p for p in TELL_PHRASES if p in after.lower()]
    assert not left, left
    sentence_lengths = [len(x.split()) for x in re.split(r"(?<=[.])\s+(?=[A-Z])", after.replace("\n\n", " ")) if x.strip()]
    return "\n".join(
        [
            "# better-writing run on the findings paragraph",
            "",
            "Skill: `.agents/skills/better-writing/SKILL.md`, version 1.3.0. Input: `influence_before.md`. Output: `influence_after.md`.",
            "",
            "The agent wrote the draft in the generic style the skill targets, then read SKILL.md and its references and followed the steps below.",
            "",
            "## 1. Brief",
            "",
            "- Genre: the findings paragraph of a coursework card that checks public documents.",
            "- Audience: a reader who will open the XML and the HKEX files.",
            "- Outcome: state what the IPO figures show after each address, at their real strength.",
            "- Constraints from `AGENTS.md`: amounts, dates, labels, and URLs are frozen. No figure may be added.",
            "- Mode: full rewrite, not light edit. The brief did not ask to keep a voice.",
            "",
            "## 2. Reading",
            "",
            "Reading this as: an analytical findings paragraph for a reader who checks sources, with a neutral tone, optimising for claims at their true strength.",
            "",
            "Dials: directness high, warmth low, personality low, density high, evidence high, polish medium.",
            "",
            "## 3. Audit",
            "",
            "No near-conclusive artefacts. The tells cluster, so the audit proceeds.",
            "",
            "| Phrase in the draft | Pattern | Reference |",
            "| --- | --- | --- |",
            '| "It is important to note that", "it should be noted", "It is also worth mentioning" | Throat-clearing | `structures-and-phrases.md`, Throat-clearing |',
            '| "played a pivotal role in shaping", "remarkable growth", "key pillar" | Significance inflation | `ai-writing-patterns.md`, Significance inflation |',
            '| "landscape" | Business jargon | `structures-and-phrases.md`, Business jargon |',
            '| "a testament to the power of targeted policy" | Significance inflation; claims a cause the figures do not show | `ai-writing-patterns.md`, Significance inflation |',
            '| "Meanwhile" | Mechanical transition | `structures-and-phrases.md`, Mechanical transitions |',
            '| "Taken together, these findings underscore" | Essay-scaffold closer | `structures-and-phrases.md`, Essay-scaffold closers |',
            '| "not merely words but a driving force" | Negative parallelism; inflation | `ai-writing-patterns.md`, Negative parallelism |',
            '| "deeply interconnected", "transformation" | Vague connection | `ai-writing-patterns.md`, Vague connection |',
            "| One paragraph, sentences up to 70 words | Uniform density, backtracking | `preflight.md`, Structure check |",
            "",
            f"Stock phrases found: {len(found)} kinds. Left after the rewrite: {len(left)}.",
            "",
            "## 4. Rewrite",
            "",
            "Moves, in the skill's order: cut throat-clearing; restate inflation at its true strength; split the long sentences so each holds one claim; put the doer first; end on the strongest fact already in the draft. "
            "The two causal clauses were cut, not reworded, because the draft's own facts (the fund had not launched, the board grew fourfold, HKEX names other causes) contradict them.",
            "",
            "## 5. Self-audit and preflight",
            "",
            f"- Preservation: every figure in the rewrite appears in the draft ({len(figures_in(after))} of {len(figures_in(before))}). Checked by `figures_in()` in `src/build_briefing.py`.",
            "- Scope words kept: \"only\", \"no\", \"most\", and \"between\" survived.",
            "- Last sentence: the new-energy sensitivity, a fact from the draft. No invented verdict.",
            "- Anti-slop: no throat-clearing, no closer, no dash hinges, no vague attribution. HKEX is named.",
            f"- Structure: sentence lengths {sentence_lengths}. Five paragraphs, each about one group of industries.",
            "- Taste: directness yes, specificity yes, restraint yes.",
            "",
            "## Change note",
            "",
            "Cut two causal claims, \"a testament to the power of targeted policy\" and \"a driving force behind the market's transformation\". The draft's own figures show timing only. Split one paragraph into five.",
            "",
            "## Counts",
            "",
            "| Measure | Before | After |",
            "| --- | --- | --- |",
            *[f"| {k} | {mb[k]} | {ma[k]} |" for k in mb],
            "",
            "![Before and after counts](better_writing_counts.png)",
            "",
        ]
    )


def write_prose_files(d: Data, before: str, after: str) -> None:
    a3 = A3 / "outputs"
    a3.mkdir(parents=True, exist_ok=True)
    (a3 / "influence_before.md").write_text(before + "\n", encoding="utf-8")
    (a3 / "influence_after.md").write_text(after + "\n", encoding="utf-8")
    (a3 / "better_writing_run.md").write_text(better_writing_record(before, after), encoding="utf-8")

    (A1 / "task_brief.md").write_text(
        "\n".join(
            [
                "# Task brief",
                "",
                "Author: CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics. University, degree, and contact details are not in the source assignments, so they stay blank.",
                "",
                "Audience: a reader who wants to check public documents.",
                "",
                "Deliverable: one card for the Policy Addresses after John Lee took office on 1 July 2022, through the address of 16 September 2026. The card records the industries with a stated government amount, whether the amount is new in that address, and the IPOs of those industries on the HKEX Main Board from each delivery date to the day before the next address.",
                "",
                "Success criteria:",
                "",
                "- Every amount has the official XML or an HKEX URL.",
                "- Each funding row carries its verb and a status. Only new money counts towards an address's total.",
                "- Industry figures come from every IPO in the HKEX New Listing Reports, not from a top-ten list.",
                "- The outcome window starts on the delivery date. A partly published window is reported with its days.",
                "- Two different official figures are both kept.",
                "- No funding amount is divided by an IPO total.",
                "",
                "Sources: the English Policy Address XML for 2022 through 2026, the HKEX New Listing Reports for 2021 through 2026, and the HKEX pages copied into the Google Sheet HKEX original market figures.",
                "",
                "Boundaries: read those public files. Write under A1_CHEUNG_Yuk_Yuen/, A2_CHEUNG_Yuk_Yuen/, A3_CHEUNG_Yuk_Yuen/, data/, src/, and .cursor/skills/. Do not trade, register, or email.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    (ROOT / "README.md").write_text(
        "\n".join(
            [
                "# Policy Address funding and the industries it named",
                "",
                "Coursework for CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics.",
                "",
                "The card lists industry and technology programmes with a stated amount in the Policy Addresses from 2022 to 2026, marks whether each amount is new, and follows every Main Board IPO in those industries from each delivery date, using the HKEX New Listing Reports.",
                "",
                "## Where the work is",
                "",
                "- Assignment 1: `A1_CHEUNG_Yuk_Yuen/`. `report.pdf` is the one-page report; `outputs/analysis.pdf` holds every table and chart.",
                "- Assignment 2: `A2_CHEUNG_Yuk_Yuen/`. The skill is `policy-funding-card`.",
                "- Assignment 3: `A3_CHEUNG_Yuk_Yuen/`. The external skill is better-writing; the run is in `outputs/better_writing_run.md`.",
                "- HKEX New Listing Reports, 2021 to 2026: `data/hkex_nlr/`.",
                f"- Google Sheet: {SHEET}",
                "",
                "`HKEX_original` holds the HKEX figures. `Policy_funding` holds the funding rows with their status. `Year_after` holds the board comparison. `Industry` holds each policy industry by window. `C18_before_after` starts on 31 March 2023, when Chapter 18C took effect.",
                "",
                "## Run it again",
                "",
                "```bash",
                "python3 -m pip install -r requirements.txt",
                "python3 src/build_briefing.py",
                "python3 .agents/skills/xlsx/scripts/recalc.py A1_CHEUNG_Yuk_Yuen/outputs/term_market.xlsx",
                "```",
                "",
                "Recalculation needs LibreOffice (`soffice`). The Policy Address PDFs have no text layer in this environment. The text used here is the Chief Executive's Office XML.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    skill_src = ROOT / ".cursor" / "skills" / "policy-funding-card"
    dest = A2 / ".cursor" / "skills" / "policy-funding-card"
    if dest.exists():
        shutil.rmtree(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(skill_src, dest)

    agents = (ROOT / "AGENTS.md").read_text(encoding="utf-8")
    (A1 / "AGENTS.md").write_text(agents, encoding="utf-8")


def main() -> None:
    d = Data()
    before, after = influence_texts(d)
    write_charts(d, before, after)
    write_markdown(d, after)
    write_prose_files(d, before, after)
    xlsx = A1 / "outputs" / "term_market.xlsx"
    write_workbook(xlsx, d)
    write_pdfs(d, before, after)
    import submission

    submission.write_all(d, before, after)
    print("wrote", xlsx)


if __name__ == "__main__":
    main()
