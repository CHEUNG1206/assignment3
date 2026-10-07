#!/usr/bin/env python3
"""Build the Policy Address funding card, the audit workbook, and the PDFs.

The outcome of an address is the HKEX year after it. Amounts are the ones
read from the English Policy Address XML and the HKEX pages on 2026-10-07.
The workbook stores those amounts as inputs and converts them to HK$ million
with a formula.
"""

from __future__ import annotations

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
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parents[1]
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
SHEET = "https://docs.google.com/spreadsheets/d/1Nr4XOrgmN3wJ44HHLVwDYQaRkt7oi6D_qz0EIGy9lJ0/edit"

# year, programme, amount, unit, industries, source, note
# unit is the unit in the English Policy Address. HK$ billion and HK$ million
# both convert. "as printed" would not.
FUNDING = [
    (
        "2022",
        "Co-Investment Fund",
        30,
        "HK$ billion",
        "Enterprises setting up in Hong Kong",
        XML.format(year=2022),
        "Set aside from the Future Fund.",
    ),
    (
        "2022",
        "RAISe+ Scheme",
        10,
        "HK$ billion",
        "University research teams with potential to become start-ups",
        XML.format(year=2022),
        "Earmarked for launch the next year, on a matching basis, for at least 100 teams.",
    ),
    (
        "2022",
        "Strategic Tech Fund",
        5,
        "HK$ billion",
        "Life and health technology, AI and data science, advanced manufacturing and new energy technology",
        XML.format(year=2022),
        "Used with OASES. The goal of attracting at least 100 I&T enterprises is a target, not a second government sum.",
    ),
    (
        "2023",
        "New Industrialisation Acceleration Scheme",
        10,
        "HK$ billion",
        "Life and health technologies, AI and data science, advanced manufacturing, new energy technologies",
        XML.format(year=2023),
        "Matching of 1 (Government) to 2 (company), funding ceiling HK$200 million.",
    ),
    (
        "2023",
        "Film Development Fund and CreateSmart Initiative",
        4.3,
        "HK$ billion",
        "Film, cultural and creative industries",
        XML.format(year=2023),
        "One injection shared by the two schemes.",
    ),
    (
        "2024",
        "Research Matching Grant Scheme",
        1.5,
        "HK$ billion",
        "Research at institutions",
        XML.format(year=2024),
        "Paragraph 78. A new round.",
    ),
    (
        "2024",
        "I&T Industry-Oriented Fund",
        10,
        "HK$ billion",
        "Life and health technology, AI and robotics, semi-conductors and smart devices, advanced materials, new energy",
        XML.format(year=2024),
        "Fund-of-funds. The 2026 address says it is expected to launch in 2026.",
    ),
    (
        "2024",
        "Innovation and Technology Venture Fund",
        1.5,
        "HK$ billion",
        "Start-ups in strategic industries",
        XML.format(year=2024),
        "Redeployed to set up funds jointly with the market, on a matching basis.",
    ),
    (
        "2024",
        "I&T Accelerator Pilot Scheme",
        180,
        "HK$ million",
        "Start-up accelerators",
        XML.format(year=2024),
        "Paragraph 80. Matching 1 to 2, subsidy ceiling HK$30 million.",
    ),
    (
        "2024",
        "New Energy Transport Fund",
        750,
        "HK$ million",
        "Electric taxis and franchised buses, hydrogen fuel-cell heavy vehicles",
        XML.format(year=2024),
        "Paragraph 85 says around HK$750 million.",
    ),
    (
        "2025",
        "Frontier Technology Research Support Scheme",
        3,
        "HK$ billion",
        "AI and other frontier research",
        XML.format(year=2025),
        "Paragraph 68. Announced earlier. Applications to open soon.",
    ),
    (
        "2025",
        "Hong Kong AI Research and Development Institute",
        1,
        "HK$ billion",
        "AI research, transfer, and use cases",
        XML.format(year=2025),
        "Paragraph 69. Earmarked for establishment in 2026.",
    ),
    (
        "2025",
        "Aerospace Technology special call",
        100,
        "HK$ million",
        "Aerospace technology",
        XML.format(year=2025),
        "The sentence says over HK$100 million for six university projects. The cell stores 100.",
    ),
    (
        "2026",
        "Innovation and Technology Venture Fund, enhanced scheme",
        1,
        "HK$ billion",
        "I&T venture funds",
        XML.format(year=2026),
        "An additional earmark. It is not the HK$1.5 billion redeployed in 2024.",
    ),
    (
        "2026",
        "Artificial Intelligence Subsidy Scheme",
        1,
        "HK$ billion",
        "High-performance computing for the AI institute",
        XML.format(year=2026),
        "Paragraph 201.",
    ),
    (
        "2026",
        "AI for Welfare Lab",
        300,
        "HK$ million",
        "AI in social welfare services",
        XML.format(year=2026),
        "Paragraph 113. WelfareHK.ai.",
    ),
    (
        "2026",
        "Construction Innovation and Technology Fund",
        1,
        "HK$ billion",
        "I&T in the construction industry",
        XML.format(year=2026),
        "Paragraph 333. The Construction Industry Council's HK$400 million is not government funding.",
    ),
]

# Outcome-year HKEX rows. year is the market year, which is the year after the address.
# kind distinguishes IPO from ADT so the share formula can select IPO only.
OUTCOMES = [
    ("2023", "IPO", 46294.9, "HK$ million", "-55.8%", "29 Dec 2023", HKEX_2023, "Market Statistics 2023, p.3. Main Board and GEM. Outcome of the 19 Oct 2022 address."),
    ("2023", "ADT", 105012, "HK$ million", "-15.9%", "29 Dec 2023", HKEX_2023, "Market Statistics 2023, p.3. Average daily turnover."),
    ("2024", "IPO", 87477.8, "HK$ million", "+88.9%", "31 Dec 2024", HKEX_2024, "Market Statistics 2024, p.3. Main Board and GEM. Outcome of the 25 Oct 2023 address."),
    ("2024", "ADT", 131825, "HK$ million", "+25.5%", "31 Dec 2024", HKEX_2024, "Market Statistics 2024, p.3. Average daily turnover."),
    ("2025", "IPO", 286.9, "HK$ billion", "+226%", "31 Dec 2025", HKEX_2025, "2025 Final Results, p.11. 119 IPOs. Page 17 prints 286.8 and is not averaged in."),
    ("2025", "ADT", 249.8, "HK$ billion", "+90%", "31 Dec 2025", HKEX_2025, "2025 Final Results, p.11. Headline ADT. Outcome of the 16 Oct 2024 address."),
    ("2026", "IPO", 342.4, "HK$ billion", "+153%", "31 Aug 2026", HKEX_HL, "First eight months only, versus HK$135.1 billion a year earlier. Partial outcome of the 17 Sep 2025 address."),
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


INFLUENCE_BEFORE = """
It is important to note that the government's market perspective and the HKEX figures collectively underscore a dynamic landscape across the term. The Policy Address serves as a testament to an active stock market, while also highlighting liquidity measures and, in later years, impressive IPO fundraising. These developments are not merely numbers. They reflect a broader commitment. HKEX statistics similarly showcase robust fundraising. Taken together, the figures underscore the relationship between policy and the market, marking a pivotal chapter for Hong Kong.
""".strip()

INFLUENCE_AFTER = """
The market after each Policy Address is the next calendar year, not the year the speech describes.

After 19 October 2022, HKEX recorded IPO funds of HK$46,294.9 million in 2023, down 55.8%, and average daily turnover of HK$105,012 million, down 15.9%. The address had set aside HK$30 billion for the Co-Investment Fund, HK$10 billion for RAISe+, and HK$5 billion for the Strategic Tech Fund. Those stated sums come to HK$45 billion, close to the size of the 2023 IPO total. The year after the first package was a smaller IPO market, not a larger one.

After 25 October 2023, IPO funds in 2024 were HK$87,477.8 million, up 88.9%, and average daily turnover was HK$131,825 million, up 25.5%. Horizon Robotics (HK$6.09 billion) and RoboSense (HK$1.06 billion) were on the list of the ten largest IPOs, as was InnoScience (HK$1.40 billion). Midea Group was the largest, at HK$35.67 billion. The New Industrialisation Acceleration Scheme was HK$10 billion, with a ceiling of HK$200 million per project.

After 16 October 2024, Hong Kong raised HK$286.9 billion from 119 IPOs in 2025, up 226%, and headline average daily turnover reached HK$249.8 billion, up 90%. New economy sectors accounted for 66% of capital raised, including 21 listings under the Chapter 18 series for biotech and specialist technology. HKEX attributes the rebound to global IPO demand, frontier technologies in the region, and local market reforms. The HK$10 billion I&T Industry-Oriented Fund was still expected to launch in 2026, so it was not deploying through that IPO year. The sums named in the 2024 address are about 5% of the 2025 IPO total.

After 17 September 2025, the first eight months of 2026 raised HK$342.4 billion through IPOs, up 153% on HK$135.1 billion a year earlier, and average daily turnover was HK$282.5 billion, up 14%. Those eight months fall after the address. They leave out September to December 2025 and the rest of 2026. The Hong Kong AI Research and Development Institute, earmarked at HK$1 billion, was to be established in 2026.

The address of 16 September 2026 adds HK$1 billion for the Innovation and Technology Venture Fund, HK$1 billion for the Artificial Intelligence Subsidy Scheme, HK$300 million for the AI for Welfare Lab, and HK$1 billion for the Construction Innovation and Technology Fund. HKEX has not published 2027. The August 2026 highlights stop on 31 August, before this address, so they are not its outcome.
""".strip()


def write_markdown() -> None:
    a1 = ROOT / "A1_CHEUNG_Yuk_Yuen" / "outputs"
    a2 = ROOT / "A2_CHEUNG_Yuk_Yuen" / "outputs"
    a1.mkdir(parents=True, exist_ok=True)
    a2.mkdir(parents=True, exist_ok=True)

    lines = [
        "# Funding and the market in the following year",
        "",
        "Prepared for CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics.",
        "Period: John Lee took office on 1 July 2022. Policy Addresses from 19 October 2022 to 16 September 2026. Market sources read on 7 October 2026.",
        "",
        "The question is what the Government said it would fund, and what HKEX recorded in the calendar year after that address. The same year's full-year total is the baseline. It is not the outcome.",
        "",
        "A funding row is kept only when one sentence names an amount and commits, earmarks, injects, redeploys, or sets up a fund for an industry, technology, or cultural-industry programme.",
        "",
        "## What the address said about the market",
        "",
        "These sentences describe the market up to delivery. They are context, not the following year.",
        "",
    ]
    for year in ("2022", "2023", "2024", "2025", "2026"):
        lines.append(f"- {year}, delivered {DELIVERED[year]}. {MARKET_WORDING[year]}")

    lines += ["", "## Stated government amounts", ""]
    lines.append("| Year | Programme | Amount | Industries | Source |")
    lines.append("| --- | --- | --- | --- | --- |")
    for year, item, amount, unit, industry, url, note in FUNDING:
        lines.append(f"| {year} | {item} | {amount} {unit} | {industry} | {url} |")
        lines.append(f"|  | Note |  | {note} |  |")

    lines += ["", "## HKEX in the following year", ""]
    lines.append("| Address | Delivered | Outcome window | Label |")
    lines.append("| --- | --- | --- | --- |")
    for address, _outcome, window, label in WINDOWS:
        lines.append(f"| {address} | {DELIVERED[address]} | {window} | {label} |")

    lines += ["", "| Market year | Indicator | Value | Unit | Printed change | As of | Source |"]
    lines.append("| --- | --- | --- | --- | --- | --- | --- |")
    for year, kind, amount, unit, change, as_of, url, note in OUTCOMES:
        lines.append(f"| {year} | {kind} | {amount} | {unit} | {change} | {as_of} | {url} |")
        lines.append(f"|  | Note |  | {note} |  |  |  |")
    lines += [
        "",
        "2027 is not stated in source. The 16 September 2026 address therefore has no following-year total.",
        "The 2025 results print HK$286.9 billion on pages 7 and 11, and 286.8 on page 17. The outcome row uses 286.9. 286.8 is kept in the Google Sheet and is not averaged.",
        "Market Statistics 2024, printed page 7, lists the ten largest IPOs. The amount column is in HK$ billion. Horizon Robotics is 6.09, RoboSense is 1.06, InnoScience is 1.40, and Midea Group is 35.67.",
        "",
        "## Scale",
        "",
        "Stated amounts in an address, converted to HK$ million, divided by IPO funds in the outcome window. The workbook calculates the ratio. It is a size comparison. HKEX does not call it a cause.",
        "",
        "- 2022 address, HK$45 billion stated, against 2023 IPO funds of HK$46,294.9 million.",
        "- 2023 address, HK$14.3 billion stated, against 2024 IPO funds of HK$87,477.8 million.",
        "- 2024 address, HK$13.93 billion stated, against 2025 IPO funds of HK$286.9 billion.",
        "- 2025 address, HK$4.1 billion stated, against HK$342.4 billion raised in the first eight months of 2026. A full 2026 year would make this share smaller.",
        "- 2026 address, HK$3.3 billion stated. The 2027 IPO total is not stated in source, so no share is calculated.",
        "",
        "## Influence",
        "",
        INFLUENCE_AFTER,
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
        "The outcome window is calendar 2023, which starts after 19 October 2022.",
        "",
        "| Programme | Amount | Industries |",
        "| --- | --- | --- |",
    ]
    for year, item, amount, unit, industry, url, note in FUNDING:
        if year == "2022":
            run2.append(f"| {item} | {amount} {unit} | {industry} |")
    run2 += [
        "",
        "This card does not contain the additional HK$1 billion earmarked for the Innovation and Technology Venture Fund in 2026. That sentence is in the 2026 XML.",
        "",
        "Outcome, Market Statistics 2023, p.3:",
        "",
        "- IPO funds raised, Main Board and GEM: HK$46,294.9 million, down 55.8% on the prior year.",
        "- Average daily turnover: HK$105,012 million, down 15.9%.",
        "",
        "Label: following year recorded, source does not attribute it to the address.",
        "The 2022 full-year IPO total of HK$104,570.3 million is the baseline in the 2023 table. It is not the outcome of the October 2022 address.",
        "",
    ]
    (a2 / "run2_2022.md").write_text("\n".join(run2), encoding="utf-8")


def write_workbook(path: Path) -> None:
    wb = Workbook()
    inputs = wb.active
    inputs.title = "Inputs"
    headers = [
        "year",
        "kind",
        "item",
        "amount_as_printed",
        "unit",
        "hkd_million",
        "source_url",
        "note",
    ]
    inputs.append(headers)
    blue = Font(name="Arial", color="0000FF")
    black = Font(name="Arial", color="000000")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(name="Arial", color="FFFFFF", bold=True)
    for cell in inputs[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    rows: list[tuple] = []
    for year, item, amount, unit, industry, url, note in FUNDING:
        rows.append((year, "funding", f"{item} | {industry}", amount, unit, url, note))
    for year, kind, amount, unit, change, as_of, url, note in OUTCOMES:
        rows.append((year, f"outcome_{kind}", kind, amount, unit, url, f"{as_of}. Printed change {change}. {note}"))

    for offset, row in enumerate(rows, start=2):
        year, kind, item, amount, unit, url, note = row
        inputs.cell(offset, 1, int(year)).font = blue
        inputs.cell(offset, 2, kind).font = blue
        inputs.cell(offset, 3, item).font = blue
        inputs.cell(offset, 4, amount).font = blue
        inputs.cell(offset, 5, unit).font = blue
        cell = inputs.cell(offset, 6, hkd_million_formula(offset))
        cell.font = black
        cell.number_format = "#,##0.0"
        inputs.cell(offset, 7, url).font = blue
        note_cell = inputs.cell(offset, 8, note)
        note_cell.font = blue
        note_cell.alignment = Alignment(wrap_text=True)
        inputs.cell(offset, 4).comment = Comment(
            "Input copied from the source in column G. Do not replace with a calculated total.",
            "policy-funding-card",
        )

    last_input = 1 + len(rows)
    inputs.auto_filter.ref = f"A1:H{last_input}"
    inputs.freeze_panes = "A2"
    inputs.row_dimensions[1].height = 22
    for col, width in enumerate([12, 16, 55, 20, 16, 18, 36, 62], start=1):
        inputs.column_dimensions[get_column_letter(col)].width = width

    summary = wb.create_sheet("Year_after")
    summary_headers = [
        "address_year",
        "outcome_year",
        "window",
        "stated_funding_hkd_million",
        "outcome_ipo_hkd_million",
        "funding_share_of_ipo",
        "outcome_adt_hkd_million",
    ]
    summary.append(summary_headers)
    for cell in summary[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    for index, (address, outcome, window, _label) in enumerate(WINDOWS, start=2):
        summary.cell(index, 1, int(address)).font = blue
        if outcome:
            summary.cell(index, 2, int(outcome)).font = blue
        else:
            summary.cell(index, 2, "not stated in source").font = blue
        summary.cell(index, 3, window).font = blue
        summary.cell(
            index,
            4,
            f'=SUMIFS(Inputs!F:F,Inputs!A:A,A{index},Inputs!B:B,"funding")',
        ).font = black
        summary.cell(index, 4).number_format = "#,##0.0"
        if outcome:
            summary.cell(
                index,
                5,
                f'=SUMIFS(Inputs!F:F,Inputs!A:A,B{index},Inputs!B:B,"outcome_IPO")',
            ).font = black
            summary.cell(index, 5).number_format = "#,##0.0"
            summary.cell(index, 6, f'=IF(E{index}="", "", D{index}/E{index})').font = black
            summary.cell(index, 6).number_format = "0.0%"
            summary.cell(
                index,
                7,
                f'=SUMIFS(Inputs!F:F,Inputs!A:A,B{index},Inputs!B:B,"outcome_ADT")',
            ).font = black
            summary.cell(index, 7).number_format = "#,##0.0"
        else:
            summary.cell(index, 5, "not stated in source").font = blue
            summary.cell(index, 6, "not stated in source").font = blue
            summary.cell(index, 7, "not stated in source").font = blue

    summary["A8"] = "Share = stated funding in the address year / IPO funds in the outcome window."
    summary["A9"] = "2026 IPO and ADT are the first eight months. Do not read that bar as a full year."
    summary["A10"] = "2027 is not stated in source, so the 2026 address has no share."
    for col, width in enumerate([16, 28, 28, 32, 28, 24, 28], start=1):
        summary.column_dimensions[get_column_letter(col)].width = width
    summary.freeze_panes = "A2"
    summary.row_dimensions[1].height = 32

    chart_sheet = wb.create_sheet("Chart_data")
    chart_sheet.append(["outcome_window", "stated_funding_hkd_billion", "ipo_funds_hkd_billion"])
    for cell in chart_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
    # Rows 2-5 of Year_after are the four windows that have an IPO figure.
    # Row 6 is 2027 and stays off the chart.
    for offset, source_row in enumerate(range(2, 6), start=2):
        chart_sheet.cell(offset, 1, f"=Year_after!C{source_row}").font = black
        chart_sheet.cell(offset, 2, f"=Year_after!D{source_row}/1000").font = black
        chart_sheet.cell(offset, 2).number_format = "0.0"
        chart_sheet.cell(offset, 3, f"=Year_after!E{source_row}/1000").font = black
        chart_sheet.cell(offset, 3).number_format = "0.0"
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Stated funding and IPO funds in the following year"
    chart.y_axis.title = "HK$ billion"
    chart.y_axis.majorGridlines = None
    data = Reference(chart_sheet, min_col=2, max_col=3, min_row=1, max_row=5)
    cats = Reference(chart_sheet, min_col=1, min_row=2, max_row=5)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.legend.position = "b"
    chart.y_axis.numFmt = "#,##0"
    chart_sheet.add_chart(chart, "E2")
    chart_sheet.column_dimensions["A"].width = 28
    chart_sheet.column_dimensions["B"].width = 32
    chart_sheet.column_dimensions["C"].width = 28

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
    wrapped = [[Paragraph(cell, style) for cell in row] for row in data]
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


def write_pdf(path: Path, story: list) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(path),
        pagesize=A4,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=path.stem,
    )
    doc.build(story)


def write_pdfs() -> None:
    s = styles()

    report = [
        Paragraph("Assignment 1: Policy Address funding and the market in the following year", s["h1"]),
        Paragraph("CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics. 7 October 2026.", s["body"]),
        Paragraph("1. The task", s["h2"]),
        Paragraph(
            "I am CHEUNG Yuk Yuen. The reader is someone who wants to check public documents. "
            "The task covers the Policy Addresses after John Lee took office on 1 July 2022, through 16 September 2026. "
            "For each address I record which industries the Government funds with a stated amount, and what HKEX recorded in the following calendar year. "
            "A figure with no source is marked not stated in source. Two official figures that differ are both kept. "
            "The share of stated funding to later IPO funds is a size comparison. It is not a claim that the address caused the total.",
            s["body"],
        ),
        Paragraph("2. How I did it", s["h2"]),
        Paragraph("On 7 October I set the outcome window before opening the files: the calendar year after the address. The year of the speech is only the baseline.", s["body"]),
        Paragraph("I read AGENTS.md and memory/PROFILE.md. The name on the work is CHEUNG Yuk Yuen. University, degree, and contact details stay blank because the source assignments do not state them.", s["body"]),
        Paragraph("The Policy Address PDFs have no text layer here. I used the Chief Executive's Office English XML, pa2022.xml through pa2026.xml.", s["body"]),
        Paragraph("I checked the HKEX PDFs that have a text layer and copied the figures into the Google Sheet HKEX original market figures. The 2024 annual statistics supply the year after the 2023 address in the same HK$ million series as 2022 and 2023.", s["body"]),
        Paragraph("policy-funding-card extracts one address at a time. I ran it again on 2022 alone. That run pairs the 2022 funds with the 2023 market, and it does not pick up the extra HK$1 billion named in 2026.", s["body"]),
        Paragraph("better-writing rewrote only the influence paragraph. Amounts, dates, and labels stayed in the draft.", s["body"]),
        Paragraph("The workbook stores printed amounts in blue input cells. Conversion to HK$ million, the yearly funding sum, and the funding-to-IPO share are formulas. LibreOffice recalculated them.", s["body"]),
        Paragraph("3. Course concepts", s["h2"]),
        pdf_table(
            [
                ["Concept", "How I used it", "What changed"],
                ["Plan Mode", "The outcome window was fixed before any card was written.", "The same year's HKEX total stayed a baseline."],
                ["AGENTS.md", "Missing facts stay not stated in source. Differing official figures are both kept.", "286.9 and 286.8 were not averaged. 2027 has no invented total."],
                ["Skills", "policy-funding-card extracts one year. better-writing edits the paragraph only.", "The 2022 rerun did not copy a 2026 amount."],
                ["Memory", "Only the identity stated in the earlier assignments.", "No university, degree, or Chinese personal name was added."],
                ["MCP", "HKEX figures sit in a Google Sheet, separate from the funding rows.", "The following-year comparison has its own tab."],
            ],
            [24 * mm, 80 * mm, 82 * mm],
        ),
        Paragraph("4. Result", s["h2"]),
        Paragraph(
            f"Card: A1_CHEUNG_Yuk_Yuen/outputs/term_funding_card.md. Sheet: {SHEET}. Workbook: outputs/term_market.xlsx, sheet Year_after and the chart on Chart_data.",
            s["body"],
        ),
        Paragraph(
            "After the 2022 address, 2023 IPO funds fell 55.8% to HK$46,294.9 million. After the 2023 address, 2024 IPO funds rose 88.9% to HK$87,477.8 million. After the 2024 address, 2025 IPO funds rose 226% to HK$286.9 billion, with new economy sectors at 66% of capital raised. After the 2025 address, the first eight months of 2026 raised HK$342.4 billion, up 153%. The 2026 address has no 2027 figure.",
            s["body"],
        ),
        Paragraph(
            "Self-check: every amount used below is on an official page. The industry overlap is real for 2024 listings and for the 2025 new-economy share. HKEX does not attribute those totals to these funds, and the HK$10 billion Industry-Oriented Fund had not launched during the 2025 IPO year.",
            s["body"],
        ),
        Paragraph("Following-year window", s["h2"]),
        pdf_table(
            [["Address", "Delivered", "Outcome window", "Label"]]
            + [[a, DELIVERED[a], w, lab] for a, _o, w, lab in WINDOWS],
            [22 * mm, 28 * mm, 42 * mm, 94 * mm],
        ),
        Spacer(1, 3 * mm),
        Paragraph("HKEX outcome figures", s["h2"]),
        pdf_table(
            [["Year", "Item", "Value", "Change", "What it measures"]]
            + [
                ["2023", "IPO, HK$ million", "46,294.9", "-55.8%", "Year after the 2022 address"],
                ["2023", "ADT, HK$ million", "105,012", "-15.9%", "Year after the 2022 address"],
                ["2024", "IPO, HK$ million", "87,477.8", "+88.9%", "Year after the 2023 address"],
                ["2024", "ADT, HK$ million", "131,825", "+25.5%", "Year after the 2023 address"],
                ["2025", "IPO, HK$ billion", "286.9", "+226%", "Year after the 2024 address"],
                ["2025", "Headline ADT, HK$ billion", "249.8", "+90%", "Year after the 2024 address"],
                ["2026", "IPO, 8 months, HK$ billion", "342.4", "+153%", "Part of the year after the 2025 address"],
                ["2026", "ADT, 8 months, HK$ billion", "282.5", "+14%", "Part of the year after the 2025 address"],
                ["2027", "IPO and ADT", "not stated in source", "", "Year after the 2026 address"],
            ],
            [18 * mm, 42 * mm, 36 * mm, 22 * mm, 68 * mm],
        ),
        Spacer(1, 3 * mm),
        Paragraph("Programmes with a stated amount", s["h2"]),
        pdf_table(
            [["Year", "Programme", "Amount", "Industries"]]
            + [[y, item, f"{amount} {unit}", industry] for y, item, amount, unit, industry, url, note in FUNDING],
            [14 * mm, 52 * mm, 28 * mm, 92 * mm],
        ),
    ]
    write_pdf(ROOT / "A1_CHEUNG_Yuk_Yuen" / "report.pdf", report)

    note = [
        Paragraph("Skill note: policy-funding-card", s["h1"]),
        Paragraph("Trigger: one Policy Address, turned into a funding card scored on the HKEX year after it.", s["body"]),
        Paragraph("Use it when a new address is published and the question is which industries received a stated sum, and what IPO funds and turnover did in the next calendar year.", s["body"]),
        Paragraph("Do not use it when there is no official text, when the question is a forecast, or when a newspaper figure would replace an HKEX figure.", s["body"]),
        Paragraph("The second run used only the 2022 XML. Its outcome is the 2023 HKEX year. It does not contain the additional HK$1 billion for the Innovation and Technology Venture Fund, which is a 2026 sentence.", s["body"]),
        Paragraph("Two runs side by side", s["h2"]),
        pdf_table(
            [
                ["", "2026 address, main card", "2022 address, second run"],
                ["Outcome window", "2027, not stated in source", "2023 full year"],
                ["IPO funds in that window", "not stated in source", "HK$46,294.9 million, down 55.8%"],
                ["Venture-fund sentence", "Additional HK$1 billion", "That sentence is absent"],
                ["Largest stated fund", "HK$1 billion, four programmes", "Co-Investment Fund, HK$30 billion"],
            ],
            [36 * mm, 75 * mm, 75 * mm],
        ),
    ]
    write_pdf(ROOT / "A2_CHEUNG_Yuk_Yuen" / "skill_note.pdf", note)

    reflection = [
        Paragraph("Assignment 3 reflection: better-writing", s["h1"]),
        Paragraph("CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics.", s["body"]),
        Paragraph("Where I found it", s["h2"]),
        Paragraph(
            "The assignment names better-writing. The page is https://skills.sh/forjd/better-writing/better-writing. I used it because the funding table has to keep the source sentences, and only the influence paragraph needed a rewrite.",
            s["body"],
        ),
        Paragraph("First use", s["h2"]),
        Paragraph(
            "The draft in outputs/influence_before.md was a generic English paragraph. The rewrite in influence_after.md keeps the following-year figures: 2023 down 55.8%, 2024 up 88.9%, 2025 up 226% to HK$286.9 billion, and the first eight months of 2026 at HK$342.4 billion. It also keeps the dates 19 October 2022, 25 October 2023, 16 October 2024, 17 September 2025, and 16 September 2026. No amount was added or dropped.",
            s["body"],
        ),
        Paragraph("Friction", s["h2"]),
        Paragraph(
            "The skill asks for the reader and for the facts that must not move. If the whole table is handed to it, a funding figure can be reworded. The trigger is therefore limited to the influence paragraph.",
            s["body"],
        ),
        Paragraph("Beside the skill from Assignment 2", s["h2"]),
        Paragraph(
            "policy-funding-card decides which sentence enters the table and which HKEX year is the outcome. better-writing decides how the paragraph reads once those facts are fixed. The rule worth keeping is that a rewrite may not add or remove a number that is in the source.",
            s["body"],
        ),
        Paragraph("How I would use it at work", s["h2"]),
        Paragraph(
            "After a technical note or a policy summary is checked against the source, I would use it on the explanation I hand to a colleague. It replaces a manual pass that deletes filler. It does not replace the check against the XML or the HKEX page.",
            s["body"],
        ),
        Paragraph("What I would change first in a fork", s["h2"]),
        Paragraph(
            "I would add one rule: tables, amounts, dates, and labels are frozen, and only the explanation may change.",
            s["body"],
        ),
    ]
    write_pdf(ROOT / "A3_CHEUNG_Yuk_Yuen" / "reflection.pdf", reflection)


def write_prose_files() -> None:
    a3 = ROOT / "A3_CHEUNG_Yuk_Yuen" / "outputs"
    a3.mkdir(parents=True, exist_ok=True)
    (a3 / "influence_before.md").write_text(INFLUENCE_BEFORE + "\n", encoding="utf-8")
    (a3 / "influence_after.md").write_text(INFLUENCE_AFTER + "\n", encoding="utf-8")

    (ROOT / "A1_CHEUNG_Yuk_Yuen" / "task_brief.md").write_text(
        "\n".join(
            [
                "# Task brief",
                "",
                "Author: CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics. University, degree, and contact details are not in the source assignments, so they stay blank.",
                "",
                "Audience: a reader who wants to check public documents.",
                "",
                "Deliverable: one card for the Policy Addresses after John Lee took office on 1 July 2022, through the address of 16 September 2026. The card records the industries with a stated government amount, and the HKEX IPO funds and average daily turnover in the calendar year after each address.",
                "",
                "Success criteria:",
                "",
                "- Every amount has the official XML or an HKEX URL.",
                "- The outcome window is the next calendar year. The year of the speech is the baseline.",
                "- A partly published year is reported with its months. An unpublished year stays not stated in source.",
                "- Two different official figures are both kept.",
                "- The funding-to-IPO share is labelled as a size comparison.",
                "",
                "Sources: the English Policy Address XML for 2022 through 2026, and the HKEX pages copied into the Google Sheet HKEX original market figures.",
                "",
                "Boundaries: read those public files. Write under A1_CHEUNG_Yuk_Yuen/, A2_CHEUNG_Yuk_Yuen/, A3_CHEUNG_Yuk_Yuen/, and .cursor/skills/. Do not trade, register, or email.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    (ROOT / "A1_CHEUNG_Yuk_Yuen" / "report.md").write_text(
        "\n".join(
            [
                "# Assignment 1 report",
                "",
                "CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics. 7 October 2026.",
                "",
                "## The task",
                "",
                "I am CHEUNG Yuk Yuen. The reader is someone who wants to check public documents. The task covers the Policy Addresses after John Lee took office on 1 July 2022, through 16 September 2026. For each address I record which industries the Government funds with a stated amount, and what HKEX recorded in the following calendar year. A missing fact stays `not stated in source`. Two official figures that differ are both kept. The funding-to-IPO share is a size comparison, not a causal claim.",
                "",
                "## How I did it",
                "",
                "On 7 October I set the outcome window before opening the files. An address delivered in September or October is scored on the next calendar year. The year of the speech is the baseline.",
                "",
                "I read AGENTS.md and memory/PROFILE.md. The name on the work is CHEUNG Yuk Yuen.",
                "",
                "The Policy Address PDFs have no text layer in this environment. I used the Chief Executive's Office English XML, pa2022.xml through pa2026.xml.",
                "",
                "HKEX figures were read from the annual market statistics and the 2025 final results, then copied into the Google Sheet. Market Statistics 2024 puts the year after the 2023 address on the same HK$ million IPO line as 2022 and 2023: HK$87,477.8 million.",
                "",
                "policy-funding-card extracts one address at a time. The second run uses only 2022 and pairs it with the 2023 market. It does not copy the extra HK$1 billion named in 2026.",
                "",
                "better-writing rewrote the influence paragraph only.",
                "",
                "The workbook stores printed amounts in blue cells. HK$ million, the funding sum, and the share of that sum in the next IPO total are formulas, recalculated in LibreOffice.",
                "",
                "## Course concepts",
                "",
                "| Concept | How I used it | What changed |",
                "| --- | --- | --- |",
                "| Plan Mode | The outcome window was fixed before the card was written. | The same year's HKEX total stayed a baseline. |",
                "| AGENTS.md | Missing facts stay not stated in source. Differing official figures are both kept. | 286.9 and 286.8 were not averaged. 2027 has no invented total. |",
                "| Skills | policy-funding-card extracts one year. better-writing edits the paragraph only. | The 2022 rerun did not copy a 2026 amount. |",
                "| Memory | Only the identity stated in the earlier assignments. | No university, degree, or Chinese personal name was added. |",
                "| MCP | HKEX figures sit in a Google Sheet. | The following-year comparison has its own tab. |",
                "",
                "## Result",
                "",
                "Card: `A1_CHEUNG_Yuk_Yuen/outputs/term_funding_card.md`.",
                "",
                f"Sheet: {SHEET}",
                "",
                "Workbook: `A1_CHEUNG_Yuk_Yuen/outputs/term_market.xlsx`.",
                "",
                "After the 2022 address, 2023 IPO funds fell 55.8% to HK$46,294.9 million. After the 2023 address, 2024 IPO funds rose 88.9% to HK$87,477.8 million. After the 2024 address, 2025 IPO funds rose 226% to HK$286.9 billion, and new economy sectors were 66% of capital raised. After the 2025 address, the first eight months of 2026 raised HK$342.4 billion, up 153%. The 2026 address has no published 2027 figure.",
                "",
                "Self-check: the amounts match the cited pages. The industry overlap is in the 2024 largest-IPO list and in the 2025 new-economy share. HKEX does not attribute those totals to these funds. The Industry-Oriented Fund was still expected to launch in 2026.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    (ROOT / "A2_CHEUNG_Yuk_Yuen" / "skill_note.md").write_text(
        "\n".join(
            [
                "# Skill note",
                "",
                "Name: policy-funding-card.",
                "",
                "Trigger: one Policy Address, turned into a funding card scored on the HKEX year after it.",
                "",
                "Use it when a new address is published and the question is which industries received a stated sum, and what IPO funds and turnover did in the next calendar year.",
                "",
                "Do not use it when there is no official text, when the question is a forecast, or when a newspaper figure would replace an HKEX figure.",
                "",
                "The second run used only the 2022 XML. Its outcome is the 2023 HKEX year: IPO funds of HK$46,294.9 million, down 55.8%, and average daily turnover of HK$105,012 million, down 15.9%. It does not contain the additional HK$1 billion for the Innovation and Technology Venture Fund. That sentence is in the 2026 address. The Co-Investment Fund in the 2022 address is HK$30 billion.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    (ROOT / "A3_CHEUNG_Yuk_Yuen" / "reflection.md").write_text(
        "\n".join(
            [
                "# Assignment 3 reflection",
                "",
                "CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics.",
                "",
                "## Where I found it",
                "",
                "The assignment names better-writing. The page is https://skills.sh/forjd/better-writing/better-writing. I used it because the funding table has to keep the source sentences, and only the influence paragraph needed a rewrite.",
                "",
                "## First use",
                "",
                "The draft is `outputs/influence_before.md`, a generic English paragraph. The rewrite is `influence_after.md`. It keeps the following-year path: 2023 down 55.8%, 2024 up 88.9%, 2025 up 226% to HK$286.9 billion, and the first eight months of 2026 at HK$342.4 billion. Dates and amounts were not changed.",
                "",
                "## Friction",
                "",
                "The skill asks for the reader and for the facts that must not move. If the whole table is handed to it, a funding figure can be reworded. The trigger is limited to the influence paragraph.",
                "",
                "## Beside the skill from Assignment 2",
                "",
                "policy-funding-card decides which sentence enters the table and which HKEX year is the outcome. better-writing decides how the paragraph reads once those facts are fixed. A rewrite may not add or remove a number that is in the source.",
                "",
                "## How I would use it at work",
                "",
                "After a technical note or a policy summary is checked against the source, I would use it on the explanation I hand to a colleague. It replaces a manual pass that deletes filler. It does not replace the check against the XML or the HKEX page.",
                "",
                "## What I would change first in a fork",
                "",
                "I would freeze tables, amounts, dates, and labels, and allow edits only in the explanation.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    (ROOT / "A3_CHEUNG_Yuk_Yuen" / "skill_intro.md").write_text(
        "\n".join(
            [
                "# better-writing",
                "",
                "## What it is",
                "",
                "better-writing rewrites prose for a named reader and genre. The trigger is a draft that should be clearer, or a request to remove generic AI phrasing. The steps are: set the reader and the facts that must stay, audit the draft, then rewrite. Facts, numbers, and quoted wording stay as they were.",
                "",
                "## How CHEUNG Yuk Yuen used it",
                "",
                "Install path: `.agents/skills/better-writing/SKILL.md`, also linked at `.cursor/skills/better-writing`.",
                "",
                "It was applied to the influence paragraph after policy-funding-card had filled the table. The input was `A3_CHEUNG_Yuk_Yuen/outputs/influence_before.md`. The output is `influence_after.md`. Amounts, dates, and labels were frozen. The paragraph now states the HKEX result in the year after each address.",
                "",
                "## Work application",
                "",
                "CHEUNG Yuk Yuen's post is AI Application Trainee in the Department of Physics. In that work, public technical or policy pages have to become a note someone else can check. better-writing is the last pass on the explanation. The funding table stays with the extraction skill. The manual step it replaces is deleting filler before the note is handed on.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    (ROOT / "README.md").write_text(
        "\n".join(
            [
                "# Policy Address funding and the following year's market",
                "",
                "Coursework for CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics.",
                "",
                "The card lists industry and technology programmes with a stated amount in the Policy Addresses from 2022 to 2026, then sets each address against HKEX IPO funds and average daily turnover in the next calendar year.",
                "",
                "## Where the work is",
                "",
                "- Assignment 1: `A1_CHEUNG_Yuk_Yuen/`",
                "- Assignment 2: `A2_CHEUNG_Yuk_Yuen/`. The skill is `policy-funding-card`.",
                "- Assignment 3: `A3_CHEUNG_Yuk_Yuen/`. The external skill is better-writing.",
                f"- Google Sheet: {SHEET}",
                "",
                "`HKEX_original` holds the HKEX figures. `Year_after` holds the comparison of each address with the following year.",
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
    dest = ROOT / "A2_CHEUNG_Yuk_Yuen" / ".cursor" / "skills" / "policy-funding-card"
    if dest.exists():
        shutil.rmtree(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(skill_src, dest)

    agents = (ROOT / "AGENTS.md").read_text(encoding="utf-8")
    (ROOT / "A1_CHEUNG_Yuk_Yuen" / "AGENTS.md").write_text(agents, encoding="utf-8")


def main() -> None:
    write_markdown()
    write_prose_files()
    xlsx = ROOT / "A1_CHEUNG_Yuk_Yuen" / "outputs" / "term_market.xlsx"
    write_workbook(xlsx)
    write_pdfs()
    print("wrote", xlsx)


if __name__ == "__main__":
    main()
