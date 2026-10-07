#!/usr/bin/env python3
"""Build the Policy Address funding card, the audit workbook, and the PDFs.

Amounts are the ones read from the official XML and the HKEX pages on
2026-10-07. The workbook stores those amounts as inputs and converts 億元
to HK$ million with a formula, so a later edit of an input changes the total.
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
XML = "https://www.ceo.gov.hk/public/open-data/tc/policy_address/pa{year}.xml"
HKEX_2022 = (
    "https://www.hkex.com.hk/-/media/HKEX-Market/Market-Data/Statistics/"
    "Consolidated-Reports/Annual-Market-Statistics/e_2022-Market-Statistics.pdf"
)
HKEX_2023 = (
    "https://www.hkex.com.hk/-/media/HKEX-Market/Market-Data/Statistics/"
    "Consolidated-Reports/Annual-Market-Statistics/FY_2023-Annual-Market-Stat_Eng.pdf"
)
HKEX_2025 = (
    "https://www.hkex.com.hk/-/media/HKEX-Market/News/News-Release/2026/"
    "260226news/260226news_eng.pdf"
)
HKEX_HL = (
    "https://www.hkex.com.hk/Market-Data/Statistics/Consolidated-Reports/"
    "HKEX-Monthly-Market-Highlights?sc_lang=en"
)

# kind, year, item, amount, unit, industry, source, note
# unit is the unit printed in the source. "億元" and "HK$ million" can be
# converted. "as printed" cannot.
FUNDING = [
    (
        "2022",
        "共同投資基金",
        300,
        "億元",
        "吸引企業落戶",
        XML.format(year=2022),
        "從未來基金撥出，成立共同投資基金。段落見 XML，並在附篇重覆一次。",
    ),
    (
        "2022",
        "產學研1+計劃",
        100,
        "億元",
        "大學科研成果商品化",
        XML.format(year=2022),
        "在創新及科技基金下設立，總資助額 100 億元，由 2023 年開始。",
    ),
    (
        "2022",
        "策略性創科基金",
        50,
        "億元",
        "生命健康科技、人工智能與數據科學、先進製造與新能源",
        XML.format(year=2022),
        "正文寫配合 50 億元策略性創科基金。另寫的「過百億元投資」是吸引企業的目標，不是另一筆政府撥款。",
    ),
    (
        "2023",
        "新型工業加速計劃",
        100,
        "億元",
        "生命健康科技、人工智能與數據科學、先進製造與新能源",
        XML.format(year=2023),
        "總資助額 100 億元。單一項目配對資助最多 2 億元，比例為政府 1 比公司 2。",
    ),
    (
        "2023",
        "電影發展基金及創意智優計劃",
        43,
        "億元",
        "文化創意",
        XML.format(year=2023),
        "第 63 段：兩個計劃共注資 43 億元。",
    ),
    (
        "2024",
        "研究配對補助金計劃",
        15,
        "億元",
        "院校研究",
        XML.format(year=2024),
        "第 78 段。",
    ),
    (
        "2024",
        "創科產業引導基金",
        100,
        "億元",
        "生命健康科技、人工智能與機械人、半導體與智能設備、先進材料、新能源",
        XML.format(year=2024),
        "設立母基金。2026 年正文寫預期年內啟動，沒有再寫一個新的金額。",
    ),
    (
        "2024",
        "創科創投基金聯合基金",
        15,
        "億元",
        "策略性產業初創",
        XML.format(year=2024),
        "調撥 15 億元與業界配對。",
    ),
    (
        "2024",
        "創科加速器先導計劃",
        1.8,
        "億元",
        "初創企業加速",
        XML.format(year=2024),
        "第 80 段。配對比例政府 1 比機構 2，資助上限 3,000 萬元。",
    ),
    (
        "2024",
        "新能源運輸基金預留",
        7.5,
        "億元",
        "電動的士及專營巴士、氫能重型車試驗",
        XML.format(year=2024),
        "第 85 段寫「約 7.5 億元」。",
    ),
    (
        "2025",
        "前沿科技研究支援計劃",
        30,
        "億元",
        "人工智能等前沿科研",
        XML.format(year=2025),
        "第 68 段：早前宣布，將於短期內接受申請。",
    ),
    (
        "2025",
        "香港人工智能研發院",
        10,
        "億元",
        "人工智能研發與應用",
        XML.format(year=2025),
        "第 69 段：已預留十億元，於 2026 年成立。",
    ),
    (
        "2025",
        "航天科技特別徵集",
        1,
        "億元",
        "航天科技",
        XML.format(year=2025),
        "第 159 段寫「逾 1 億元」，表內金額取正文的 1，並在備註保留「逾」。",
    ),
    (
        "2026",
        "創科創投基金優化計劃",
        10,
        "億元",
        "創科創投",
        XML.format(year=2026),
        "預留額外 10 億元。這是 2024 年 15 億元那一行以外的新句子。",
    ),
    (
        "2026",
        "人工智能資助計劃",
        10,
        "億元",
        "人工智能算力",
        XML.format(year=2026),
        "第 201 段：向計劃注資 10 億元。",
    ),
    (
        "2026",
        "社福 AI 實驗室",
        3,
        "億元",
        "社福服務中的人工智能應用",
        XML.format(year=2026),
        "第 113 段。",
    ),
    (
        "2026",
        "建造業創新及科技基金",
        10,
        "億元",
        "建造業創新科技",
        XML.format(year=2026),
        "第 333 段：政府注資 10 億元。建造業議會另出資 4 億元，不計入政府金額。",
    ),
]

# HKEX rows already checked against the source pages. value, unit, as_of, url, note
HKEX = [
    ("2022", "IPO funds raised, Main Board and GEM", 104570.3, "HK$ million", "30 Dec 2022", HKEX_2022, "Market Statistics 2022, p.3. Provisional. Includes January–June, before 1 July 2022."),
    ("2023", "IPO funds raised, Main Board and GEM", 46294.9, "HK$ million", "29 Dec 2023", HKEX_2023, "Market Statistics 2023, p.3. Provisional."),
    ("2024", "IPOs, Main Board, as printed", 87.8, "as printed", "2025 results, p.17", HKEX_2025, "Unit is not repeated on this row. The same document's narrative calls the 2025 figure HK$286.9 billion."),
    ("2025", "IPO fundraising, narrative", 286.9, "HK$ billion", "31 Dec 2025", HKEX_2025, "2025 Final Results, p.7. 119 new listings."),
    ("2025", "IPOs, Main Board, as printed", 286.8, "as printed", "2025 results, p.17", HKEX_2025, "Kept beside 286.9. Not averaged."),
    ("2026", "Funds raised through IPOs, first eight months", 342.4, "HK$ billion", "31 Aug 2026", HKEX_HL, "Before the 16 September 2026 Policy Address. Year not finished."),
]

DELIVERED = {
    "2022": "2022-10-19",
    "2023": "2023-10-25",
    "2024": "2024-10-16",
    "2025": "2025-09-17",
    "2026": "2026-09-16",
}

INFLUENCE = {
    "2022": "市場結果早於這份施政報告",
    "2023": "同一年同時出現，來源沒有寫因果",
    "2024": "同一年同時出現，來源沒有寫因果",
    "2025": "市場結果早於這份施政報告",
    "2026": "市場結果早於這份施政報告",
}

MARKET_WORDING = {
    "2022": "正文寫香港有交投活躍的股票市場，沒有寫新股集資額或成交額。",
    "2023": "第 71 段接納促進股票市場流動性專責小組的建議，沒有寫新股集資額。",
    "2024": "金融市場段落寫人工智能、代幣化和穩定幣的政策方向，沒有寫新股集資額。",
    "2025": "第 92 段寫恒生指數自年初上升逾兩成，每日平均成交額接近 2,500 億元，新股集資額至八月底累計逾 1,300 億元。",
    "2026": "第 36 段寫截至今年 8 月底的新股集資額超過 3,400 億港元。",
}


def hkd_million(amount: float, unit: str) -> float | None:
    """Return HK$ million only when the printed unit supports the conversion."""
    if unit == "億元":
        return amount * 100
    if unit == "HK$ million":
        return amount
    if unit == "HK$ billion":
        return amount * 1000
    return None


def write_markdown() -> None:
    a1 = ROOT / "A1_CHEUNG_Yuk_Yuen" / "outputs"
    a2 = ROOT / "A2_CHEUNG_Yuk_Yuen" / "outputs"
    a1.mkdir(parents=True, exist_ok=True)
    a2.mkdir(parents=True, exist_ok=True)

    lines = [
        "# 任期資金與市場卡片",
        "",
        "期間：2022 年 7 月 1 日李家超就任行政長官，至 2026 年 10 月 7 日。",
        "政府原文：行政長官辦公室公開 XML。港交所原文：Google 試算表 HKEX original market figures 內已核對的行列。",
        "",
        "這張卡只收產業、創科或金融計劃裡，同一句寫了金額，並且用了撥款、預留、注資、調撥或設立基金的項目。教育、福利和工程的其他金額留在 XML，不放進這張表。",
        "",
        "## 市場說法",
        "",
    ]
    for year in ("2022", "2023", "2024", "2025", "2026"):
        lines.append(f"- {year} 年 {DELIVERED[year]} 發表。{MARKET_WORDING[year]}")
    lines += ["", "## 寫明金額的政府計劃", ""]
    lines.append("| 年 | 計劃 | 金額 | 產業 | 來源 |")
    lines.append("| --- | --- | --- | --- | --- |")
    for year, item, amount, unit, industry, url, note in FUNDING:
        lines.append(f"| {year} | {item} | {amount} {unit} | {industry} | {url} |")
        lines.append(f"|  | 備註 |  | {note} |  |")
    lines += ["", "## 港交所同年數字", ""]
    lines.append("| 年 | 指標 | 數字 | 單位 | 頁面日期 | 影響標籤 |")
    lines.append("| --- | --- | --- | --- | --- | --- |")
    seen = set()
    for year, item, amount, unit, as_of, url, note in HKEX:
        label = INFLUENCE[year]
        lines.append(f"| {year} | {item} | {amount} | {unit} | {as_of} | {label} |")
        if year not in seen:
            seen.add(year)
    lines += [
        "",
        "2022 年港交所全年數包含 1 月至 6 月，當時他尚未就任，也包含 10 月 19 日施政報告之前的月份。",
        "2025 年施政報告引用的是八月底之前的數字，發表日是 9 月 17 日。2026 年引用的是 8 月底，發表日是 9 月 16 日，港交所每月摘要的 3,424 億元更新至 8 月 31 日。",
        "2025 年業績公告第 7 頁寫 286.9，第 17 頁印 286.8。兩格都保留。",
        "2024 年 87.8 是同一頁印出的數字，該行沒有再印單位，所以不換算成億元。",
        "",
        "## 影響說明",
        "",
        INFLUENCE_AFTER.strip(),
        "",
    ]
    (a1 / "term_funding_card.md").write_text("\n".join(lines), encoding="utf-8")

    run2 = [
        "# 第二輪：只跑 2022 年施政報告",
        "",
        "觸發：policy-funding-card，輸入是 2022 年施政報告 XML。",
        "這輪不讀 2026 年的句子。",
        "",
        f"發表日：{DELIVERED['2022']}。",
        f"市場說法：{MARKET_WORDING['2022']}",
        "",
        "| 計劃 | 金額 | 產業 |",
        "| --- | --- | --- |",
    ]
    for year, item, amount, unit, industry, url, note in FUNDING:
        if year == "2022":
            run2.append(f"| {item} | {amount} {unit} | {industry} |")
    run2 += [
        "",
        "2022 年這張卡沒有「創科創投基金優化計劃額外 10 億元」。那句在 2026 年 XML。",
        "港交所 2022 年主板加 GEM 新股集資額是 104,570.3 百萬港元，截至 2022 年 12 月 30 日。",
        f"影響標籤：{INFLUENCE['2022']}。全年數早於並跨過 10 月 19 日，來源沒有把集資額歸因於這份施政報告。",
        "",
    ]
    (a2 / "run2_2022.md").write_text("\n".join(run2), encoding="utf-8")


INFLUENCE_BEFORE = """
It is important to note that the government's market perspective and the HKEX figures collectively underscore a dynamic landscape across the term. The Policy Address serves as a testament to an active stock market, while also highlighting liquidity measures and, in later years, impressive IPO fundraising. These developments are not merely numbers. They reflect a broader commitment. HKEX statistics similarly showcase robust fundraising. Taken together, the figures underscore the relationship between policy and the market, marking a pivotal chapter for Hong Kong.
""".strip()

INFLUENCE_AFTER = """
從 2022 年 7 月 1 日到 2026 年 10 月 7 日，施政報告和港交所寫的是兩件不同的事。2022 年施政報告說股票市場交投活躍，沒有寫新股集資額。2023 年施政報告接納流動性專責小組的建議，也沒有寫集資額。2025 年施政報告寫恒生指數自年初上升逾兩成，並寫新股集資額至八月底逾 1,300 億元；發表日是 9 月 17 日，所以這組數字在發表之前。2026 年施政報告寫截至 8 月底集資額超過 3,400 億港元；港交所每月摘要寫首八個月 3,424 億元，更新日是 8 月 31 日，早於 9 月 16 日的發表。港交所的全年數還包含發表前的月份，不能用來量度發表之後的市場。表裡的政府金額是正文寫明要撥、預留或注資的計劃，不是從集資額倒推回來的。
""".strip()


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
        rows.append((year, "funding", f"{item}｜{industry}", amount, unit, url, note))
    for year, item, amount, unit, as_of, url, note in HKEX:
        rows.append((year, "hkex", item, amount, unit, url, f"{as_of}. {note}"))

    for offset, row in enumerate(rows, start=2):
        year, kind, item, amount, unit, url, note = row
        inputs.cell(offset, 1, int(year)).font = blue
        inputs.cell(offset, 2, kind).font = blue
        inputs.cell(offset, 3, item).font = blue
        inputs.cell(offset, 4, amount).font = blue
        inputs.cell(offset, 5, unit).font = blue
        # Convert only when the printed unit has a fixed factor.
        formula = f'=IF(E{offset}="億元",D{offset}*100,IF(E{offset}="HK$ million",D{offset},IF(E{offset}="HK$ billion",D{offset}*1000,"")))'
        cell = inputs.cell(offset, 6, formula)
        cell.font = black
        inputs.cell(offset, 7, url).font = blue
        note_cell = inputs.cell(offset, 8, note)
        note_cell.font = blue
        note_cell.alignment = Alignment(wrap_text=True)
        inputs.cell(offset, 4).comment = Comment(
            "Input copied from the source in column G. Do not replace with a calculated total.",
            "policy-funding-card",
        )

    inputs.auto_filter.ref = f"A1:H{1 + len(rows)}"
    inputs.freeze_panes = "A2"
    inputs.row_dimensions[1].height = 22
    for col, width in enumerate([12, 14, 42, 20, 16, 18, 36, 55], start=1):
        inputs.column_dimensions[get_column_letter(col)].width = width

    summary = wb.create_sheet("Funding_sum")
    summary.append(["year", "stated_funding_hkd_million"])
    for cell in summary[1]:
        cell.fill = header_fill
        cell.font = header_font
    for index, year in enumerate(("2022", "2023", "2024", "2025", "2026"), start=2):
        summary.cell(index, 1, int(year)).font = blue
        # SUMIFS over the formula column. Blank conversions stay out of the sum.
        summary.cell(
            index,
            2,
            f'=SUMIFS(Inputs!F:F,Inputs!A:A,A{index},Inputs!B:B,"funding")',
        ).font = black
    summary.cell(7, 1, "Total of stated amounts").font = black
    summary.cell(7, 2, "=SUM(B2:B6)").font = black
    summary["A9"] = "This total adds only rows whose unit is 億元. It is not total government expenditure."
    summary.column_dimensions["A"].width = 32
    summary.column_dimensions["B"].width = 36

    chart_sheet = wb.create_sheet("IPO_full_year")
    chart_sheet.append(["year", "ipo_hkd_billion", "what_this_is"])
    for cell in chart_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
    # Point at the Inputs rows. Row numbers follow FUNDING then HKEX order.
    # HKEX starts after len(FUNDING) + header.
    hkex_start = 2 + len(FUNDING)
    # 2022, 2023 are HK$ million; divide by 1000. 2025 narrative is HK$ billion.
    chart_sheet["A2"] = 2022
    chart_sheet["B2"] = f"=Inputs!F{hkex_start}/1000"
    chart_sheet["C2"] = "Main Board and GEM, HK$ million converted by /1000"
    chart_sheet["A3"] = 2023
    chart_sheet["B3"] = f"=Inputs!F{hkex_start + 1}/1000"
    chart_sheet["C3"] = "Main Board and GEM, HK$ million converted by /1000"
    chart_sheet["A4"] = 2025
    chart_sheet["B4"] = f"=Inputs!D{hkex_start + 3}"
    chart_sheet["C4"] = "Narrative figure already in HK$ billion. 286.8 on p.17 is not plotted."
    for row in range(2, 5):
        chart_sheet.cell(row, 1).font = blue
        chart_sheet.cell(row, 2).font = black
        chart_sheet.cell(row, 3).font = blue
    chart_sheet["A6"] = "2024 value 87.8 and 2026 value 342.4 are not on this chart."
    chart_sheet["A7"] = "2024 has no unit on the printed row. 2026 is eight months and ends before the Policy Address."
    chart = BarChart()
    chart.type = "col"
    chart.title = "IPO funds, full years with a stated unit"
    chart.y_axis.title = "HK$ billion"
    data = Reference(chart_sheet, min_col=2, min_row=1, max_row=4)
    cats = Reference(chart_sheet, min_col=1, min_row=2, max_row=4)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.legend = None
    chart_sheet.add_chart(chart, "E2")
    chart_sheet.column_dimensions["A"].width = 18
    chart_sheet.column_dimensions["B"].width = 22
    chart_sheet.column_dimensions["C"].width = 62

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def styles() -> dict[str, ParagraphStyle]:
    if "WQY" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("WQY", FONT_PATH, subfontIndex=0))
    return {
        "h1": ParagraphStyle("h1", fontName="WQY", fontSize=16, leading=22, textColor=colors.HexColor("#1F4E79"), spaceAfter=8),
        "h2": ParagraphStyle("h2", fontName="WQY", fontSize=12, leading=16, textColor=colors.HexColor("#1F4E79"), spaceBefore=8, spaceAfter=4),
        "body": ParagraphStyle("body", fontName="WQY", fontSize=9, leading=13, alignment=TA_LEFT, spaceAfter=4),
        "small": ParagraphStyle("small", fontName="WQY", fontSize=8, leading=11, textColor=colors.HexColor("#333333")),
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
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
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
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
        title=path.stem,
    )
    doc.build(story)


def write_pdfs() -> None:
    s = styles()
    usable = 182 * mm

    report = [
        Paragraph("作業 1 報告：施政報告的市場說法與產業資金", s["h1"]),
        Paragraph("CHEUNG Yuk Yuen　2026-10-07", s["body"]),
        Paragraph("1. 任務", s["h2"]),
        Paragraph(
            "我是物理系 AI Application Trainee。讀者是要核對公開文件的人。任務是從李家超 2022 年 7 月 1 日就任起，到 2026 年 10 月 7 日，整理每份施政報告怎樣寫市場，以及哪些產業計劃在同一句裡寫了金額。成功標準是每個金額都有官方 XML 或港交所出處，缺的資料標成 not stated in source，而且不把集資額說成施政報告造成的結果。",
            s["body"],
        ),
        Paragraph("2. 做法", s["h2"]),
        Paragraph("10 月 7 日先定欄位和四個影響標籤，再開檔。", s["body"]),
        Paragraph("讀 AGENTS.md 和 memory/PROFILE.md。姓名只用 CHEUNG Yuk Yuen。", s["body"]),
        Paragraph("施政報告 PDF 沒有文字層。改讀行政長官辦公室的公開 XML：pa2022.xml 至 pa2026.xml。", s["body"]),
        Paragraph("用 pdf 的做法核對可抽文字的港交所 PDF。數字先寫進 Google 試算表 HKEX original market figures。", s["body"]),
        Paragraph("用 policy-funding-card 逐年抽資金句。2022 年單獨再跑一次，確認沒有抄入 2026 年的額外 10 億元。", s["body"]),
        Paragraph("better-writing 只改寫影響說明。金額、日期和標籤留在底稿。", s["body"]),
        Paragraph("xlsx 把原文金額放在藍色輸入格，億元換算用公式。", s["body"]),
        Paragraph("3. 用到的概念", s["h2"]),
        pdf_table(
            [
                ["概念", "怎樣用", "改變了什麼"],
                ["Plan Mode", "先定欄位、標籤和來源，再寫卡片", "沒有先寫一篇市場評論"],
                ["AGENTS.md", "缺資料寫 not stated in source，兩份官方數字都保留", "2025 年的 286.9 和 286.8 沒有被平均"],
                ["Skills", "自建 policy-funding-card，再用 better-writing", "資金表和說明分成兩步"],
                ["Memory", "只記住作業裡寫過的身份", "沒有補上院校、學位或中文姓名"],
                ["MCP", "港交所原文數字放在 Google 試算表", "市場數字和施政報告金額分開存放"],
            ],
            [28 * mm, 78 * mm, 76 * mm],
        ),
        Paragraph("4. 結果", s["h2"]),
        Paragraph(
            "卡片在 A1_CHEUNG_Yuk_Yuen/outputs/term_funding_card.md。試算表在 https://docs.google.com/spreadsheets/d/1Nr4XOrgmN3wJ44HHLVwDYQaRkt7oi6D_qz0EIGy9lJ0/edit 。本地核對簿是 outputs/term_market.xlsx。",
            s["body"],
        ),
        Paragraph(
            "自我評估：資金句和港交所數字對得上來源；影響判斷停在時間先後，沒有做成因估計。",
            s["body"],
        ),
        Paragraph("逐年影響標籤", s["h2"]),
        pdf_table(
            [["年", "發表日", "標籤", "市場說法要點"]]
            + [
                [year, DELIVERED[year], INFLUENCE[year], MARKET_WORDING[year][:80]]
                for year in ("2022", "2023", "2024", "2025", "2026")
            ],
            [18 * mm, 28 * mm, 48 * mm, 88 * mm],
        ),
        Spacer(1, 4 * mm),
        Paragraph("寫明金額的計劃（億元）", s["h2"]),
        pdf_table(
            [["年", "計劃", "億元", "產業"]]
            + [[y, item, str(amount), industry] for y, item, amount, unit, industry, url, note in FUNDING],
            [16 * mm, 48 * mm, 18 * mm, 100 * mm],
        ),
    ]
    write_pdf(ROOT / "A1_CHEUNG_Yuk_Yuen" / "report.pdf", report)

    note = [
        Paragraph("Skill 備註：policy-funding-card", s["h1"]),
        Paragraph("觸發句：給一份施政報告，做成市場與資金卡。", s["body"]),
        Paragraph("在這時用：每年一份新的施政報告，要抽出寫明金額的產業計劃，並對上港交所已核對的數字。", s["body"]),
        Paragraph("不要用：沒有官方全文、要預測指數、或要把報紙數字換成港交所數字的時候。", s["body"]),
        Paragraph("第二輪用 2022 年 XML。結果沒有出現 2026 年「額外預留 10 億元」那一句。", s["body"]),
        Paragraph("兩輪對照", s["h2"]),
        pdf_table(
            [
                ["", "2026 年主卡", "2022 年第二輪"],
                ["市場數字", "截至 8 月底超過 3,400 億港元", "沒有寫集資額"],
                ["創科創投基金", "額外預留 10 億元", "這句不存在"],
                ["共同投資基金", "不在 2026 年新撥款表", "300 億元"],
            ],
            [40 * mm, 70 * mm, 72 * mm],
        ),
    ]
    write_pdf(ROOT / "A2_CHEUNG_Yuk_Yuen" / "skill_note.pdf", note)

    reflection = [
        Paragraph("作業 3 反思：better-writing", s["h1"]),
        Paragraph("發現", s["h2"]),
        Paragraph(
            "作業說明列出 better-writing，頁面是 https://skills.sh/forjd/better-writing/better-writing 。我選它，是因為資金表需要原句，說明段需要改寫。",
            s["body"],
        ),
        Paragraph("第一次使用", s["h2"]),
        Paragraph(
            "底稿是一段帶有空泛評價的英文。改寫後的中文說明只保留四份施政報告的市場原句、兩個日期，以及港交所數字早於發表日這件事。金額沒有改。",
            s["body"],
        ),
        Paragraph("摩擦", s["h2"]),
        Paragraph(
            "Skill 要求先定讀者、文體和不能改的事實。若一開始就把整張表交給它，撥款數字會有被改寫的風險。所以觸發範圍要寫成只改影響段。",
            s["body"],
        ),
        Paragraph("和作業 2 的比較", s["h2"]),
        Paragraph(
            "policy-funding-card 規定什麼句子可以進表。better-writing 規定改寫時保留事實。值得留下的規則是：改寫不得增刪來源裡的數字。",
            s["body"],
        ),
        Paragraph("工作上會怎樣用", s["h2"]),
        Paragraph(
            "技術資料或政策摘要寫完之後，用它把說明段改成可以交給同事看的版本。它取代的是手動刪掉空話，不取代核對原文。",
            s["body"],
        ),
        Paragraph("若分叉，先改哪裡", s["h2"]),
        Paragraph(
            "先加一條：表格、金額、日期和標籤凍結，只允許改說明段。",
            s["body"],
        ),
    ]
    write_pdf(ROOT / "A3_CHEUNG_Yuk_Yuen" / "reflection.pdf", reflection)


def write_prose_files() -> None:
    a3 = ROOT / "A3_CHEUNG_Yuk_Yuen" / "outputs"
    a3.mkdir(parents=True, exist_ok=True)
    (a3 / "influence_before.md").write_text(INFLUENCE_BEFORE + "\n", encoding="utf-8")
    (a3 / "influence_after.md").write_text(INFLUENCE_AFTER + "\n", encoding="utf-8")

    skill_src = ROOT / ".cursor" / "skills" / "policy-funding-card"
    dest = ROOT / "A2_CHEUNG_Yuk_Yuen" / ".cursor" / "skills" / "policy-funding-card"
    if dest.exists():
        shutil.rmtree(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(skill_src, dest)


def main() -> None:
    write_markdown()
    write_prose_files()
    xlsx = ROOT / "A1_CHEUNG_Yuk_Yuen" / "outputs" / "term_market.xlsx"
    write_workbook(xlsx)
    write_pdfs()
    print("wrote", xlsx)


if __name__ == "__main__":
    main()
