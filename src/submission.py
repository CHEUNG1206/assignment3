"""The short submission files the Lecture 3 handout asks for.

report (about one page), skill note (5-8 lines plus a run table), skill
intro, and reflection, each as .md and .pdf. The long tables and charts stay
in A1_CHEUNG_Yuk_Yuen/outputs/. Every figure is read from Data so the short
files cannot drift from the card.
"""

from __future__ import annotations

from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, Spacer

from build_briefing import (
    A1,
    A2,
    A3,
    DELIVERED,
    SHEET,
    Data,
    chart_paths,
    figure,
    m,
    pct,
    pdf_table,
    styles,
    text_metrics,
    window_label,
    write_pdf,
)

NAME = "CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics."

TIMELINE = [
    ["1", "Wrote the task brief and AGENTS.md rules: copy figures from the source, write not stated in source, keep conflicting official figures."],
    ["2", "The agent recalled memory/PROFILE.md, so the work names me only as CHEUNG Yuk Yuen, AI Application Trainee."],
    ["3", "The agent read the English Policy Address XML for 2022 to 2026 and listed every funding sentence that states an amount."],
    ["4", "The agent copied HKEX figures into a Google Sheet through the Google Sheets MCP."],
    ["5", "The first result compared each address with the same year's market. I pointed out that this counts months before the speech, so the outcome moved to the period after delivery."],
    ["6", "I asked for each funded industry, not the board total. The agent switched to every Main Board IPO in the HKEX New Listing Reports, counted from each delivery date."],
    ["7", "A review found money announced before an address counted as new, and a ratio of funding to IPO funds that compared unrelated sums. The agent added a status column from each sentence's verb and removed the ratio."],
    ["8", "The agent built the workbook with formulas, the charts and the PDFs, and packaged the workflow as the policy-funding-card skill for Assignment 2."],
]

CONCEPTS = [
    ["Concept", "How I used it", "What it changed"],
    ["AGENTS.md", "Rules on sources, missing facts, conflicting figures, and which verbs make money new.", "HKEX's 285.8, 286.8 and 286.9 billion for 2025 are all kept. The 2025 address has no new amount."],
    ["Skills", "policy-funding-card runs one address at a time. better-writing edits the findings only.", "The 2022 rerun did not pick up a 2026 amount."],
    ["Memory", "memory/PROFILE.md holds the identity from my earlier assignments.", "No university, degree, or Chinese name was added."],
    ["MCP", "The Google Sheets MCP holds the HKEX figures and the funding rows.", "Each comparison has its own tab that a reader can open."],
]


def key_finding(d: Data) -> str:
    return (
        f"Main finding: semi-conductors went from no IPO to {d.ipos('After 2025', 'Semi-conductors')} IPOs and HK${m(d.funds('After 2025', 'Semi-conductors'))} million after the 2025 address, "
        "but the HK$10 billion fund that named them has not launched. Over the same windows the whole Main Board grew about fourfold, "
        "and HKEX credits global demand, frontier technology and listing reforms. The figures show timing, not that funding moved an industry."
    )


SELF_ASSESSMENT = (
    "Self-assessment: every figure traces to a source file and the method is written down; "
    "industry totals other than biotech come from a name screen, so they are lower bounds."
)


def write_report(d: Data) -> None:
    s = styles()
    story = [
        Paragraph("Assignment 1: Policy Address funding and the industries it named", s["h1"]),
        Paragraph(f"{NAME} 7 October 2026.", s["body"]),
        Paragraph("1. The task", s["h2"]),
        Paragraph(
            "I asked the agent to find which industries John Lee's Policy Addresses from 2022 to 2026 fund with a stated amount, "
            "and what happened to those industries' IPOs on HKEX after each address. The reader is someone who checks public documents. "
            "Sources: the Policy Address XML, the HKEX New Listing Reports, and HKEX market statistics.",
            s["body"],
        ),
        Paragraph("2. How I did it", s["h2"]),
        pdf_table([["Step", "What happened"]] + TIMELINE, [12 * mm, 174 * mm]),
        Paragraph("3. Course concepts", s["h2"]),
        pdf_table(CONCEPTS, [22 * mm, 84 * mm, 80 * mm]),
        Paragraph("4. Result", s["h2"]),
        Paragraph(
            "Files: outputs/term_funding_card.md, outputs/analysis.pdf (all tables and charts), outputs/term_market.xlsx. "
            f"Google Sheet: {SHEET}",
            s["body"],
        ),
        figure(chart_paths()["shares"], 150),
        Paragraph(key_finding(d), s["body"]),
        Paragraph(SELF_ASSESSMENT, s["body"]),
    ]
    write_pdf(A1 / "report.pdf", story)

    lines = [
        "# Assignment 1 report",
        "",
        f"{NAME} 7 October 2026.",
        "",
        "## 1. The task",
        "",
        "I asked the agent to find which industries John Lee's Policy Addresses from 2022 to 2026 fund with a stated amount, and what happened to those industries' IPOs on HKEX after each address. The reader is someone who checks public documents. Sources: the Policy Address XML, the HKEX New Listing Reports, and HKEX market statistics.",
        "",
        "## 2. How I did it",
        "",
        *[f"{n}. {text}" for n, text in TIMELINE],
        "",
        "## 3. Course concepts",
        "",
        "| " + " | ".join(CONCEPTS[0]) + " |",
        "| --- | --- | --- |",
        *["| " + " | ".join(row) + " |" for row in CONCEPTS[1:]],
        "",
        "## 4. Result",
        "",
        f"Files: `outputs/term_funding_card.md`, `outputs/analysis.pdf` (all tables and charts), `outputs/term_market.xlsx`. Google Sheet: {SHEET}",
        "",
        "![Share of Main Board IPO funds by window](outputs/charts/industry_window_shares.png)",
        "",
        key_finding(d),
        "",
        SELF_ASSESSMENT,
        "",
    ]
    (A1 / "report.md").write_text("\n".join(lines), encoding="utf-8")


def skill_note_lines(d: Data) -> list[str]:
    return [
        "Skill: policy-funding-card. Trigger: \"policy-funding-card\", \"policy address\" or \"施政報告\" with one address.",
        "Use it when a new Policy Address is published and the question is which industries get a stated sum, whether the sum is new, and what their IPOs did afterwards.",
        "Do not use it without the official text, for a forecast, or to replace an HKEX figure with a newspaper one.",
        f"Second run: the 2022 address alone. It found HK${m(d.new_total['2022'])} million of new money and marked the HK$5 billion Strategic Tech Fund as already in place.",
        f"Its window, {window_label(d, 'After 2022')}, shows life and health at {pct(d.share('After 2022', 'Life and health'))} of Main Board IPO funds and no new-energy IPO by the name screen.",
        "Worked: the run kept the 2026 HK$1 billion venture-fund amount out of the 2022 card.",
        "Failed first: industries were matched only on HKEX's ten-largest IPO list, and money announced before an address counted as new.",
        "Changed: the skill now reads every IPO in the New Listing Reports from the delivery date, and records each funding verb as a status.",
    ]


def write_skill_note(d: Data) -> None:
    s = styles()
    lines = skill_note_lines(d)
    story = [
        Paragraph("Skill note: policy-funding-card", s["h1"]),
        Paragraph(NAME, s["body"]),
        *[Paragraph(line, s["body"]) for line in lines],
        Paragraph("Run 1 and run 2 side by side", s["h2"]),
        pdf_table(
            [
                ["", "Run 1: 2026 address", "Run 2: 2022 address"],
                ["New money in the address", f"HK${m(d.new_total['2026'])} million", f"HK${m(d.new_total['2022'])} million"],
                ["Already in place", f"HK${m(d.earlier_total['2026'])} million", f"HK${m(d.earlier_total['2022'])} million"],
                ["Window after delivery", window_label(d, "After 2026") + " (14 days)", window_label(d, "After 2022")],
                ["Main Board IPOs in the window", f"{d.board['After 2026']['board_ipos']}, HK${m(d.board['After 2026']['board_hkd_million'])} million", f"{d.board['After 2022']['board_ipos']}, HK${m(d.board['After 2022']['board_hkd_million'])} million"],
            ],
            [46 * mm, 70 * mm, 70 * mm],
        ),
        Spacer(1, 3 * mm),
        Paragraph("Run 2: the 12 months before the 2022 address and the window after, HK$ billion", s["h2"]),
        figure(chart_paths()["run2"]),
    ]
    write_pdf(A2 / "skill_note.pdf", story)
    md = ["# Skill note", "", *[f"- {line}" for line in lines], "", "![Run 2 before and after](outputs/run2_2022_before_after.png)", ""]
    (A2 / "skill_note.md").write_text("\n".join(md), encoding="utf-8")


REFLECTION = [
    ("Discovery", "The handout lists better-writing at https://skills.sh/forjd/better-writing/better-writing. I chose it because my findings paragraph had to read plainly while every figure stayed fixed."),
    (
        "First use",
        "I asked the agent to write a rough findings draft in the generic style the skill targets, then to apply better-writing to it. "
        "The agent read SKILL.md and its reference files and followed the five steps: brief, reading, audit, rewrite, preflight. The record is outputs/better_writing_run.md. "
        "It cut throat-clearing and two causal claims the figures do not support: \"a testament to the power of targeted policy\" and \"a driving force behind the market's transformation\".",
    ),
    ("Friction", "The skill edits prose and does not check facts. Seeing that a causal claim was unsupported needed the funding status and the fund's launch date from my own card."),
    ("vs Assignment 2", "policy-funding-card decides which facts enter the table. better-writing decides how they read. The rule worth borrowing is its preflight check that nothing in the rewrite is new."),
    ("Work fit", "I would use it on each technology-research note I hand to a colleague, as the last step after the source check. It replaces my manual pass that deletes filler and inflated claims."),
    ("One change", "I would add a preflight line that a script confirms every figure in the rewrite appears in the draft, with tables, amounts, dates and labels frozen."),
]


def write_a3(d: Data, before: str, after: str) -> None:
    s = styles()
    mb, ma = text_metrics(before), text_metrics(after)
    counts = [["Measure", "Before", "After"]] + [[k, str(mb[k]), str(ma[k])] for k in mb]
    story = [Paragraph("Assignment 3 reflection: better-writing", s["h1"]), Paragraph(NAME, s["body"])]
    for title, text in REFLECTION:
        story += [Paragraph(title, s["h2"]), Paragraph(text, s["body"])]
        if title == "First use":
            story += [pdf_table(counts, [70 * mm, 40 * mm, 40 * mm]), Spacer(1, 2 * mm), figure(chart_paths()["draft"], 160)]
    write_pdf(A3 / "reflection.pdf", story)

    md = ["# Assignment 3 reflection", "", NAME, ""]
    for title, text in REFLECTION:
        md += [f"## {title}", "", text, ""]
        if title == "First use":
            md += ["| Measure | Before | After |", "| --- | --- | --- |", *[f"| {k} | {mb[k]} | {ma[k]} |" for k in mb], "", "![Before and after counts](outputs/better_writing_counts.png)", ""]
    (A3 / "reflection.md").write_text("\n".join(md), encoding="utf-8")

    (A3 / "skill_intro.md").write_text(
        "\n".join(
            [
                "# better-writing",
                "",
                "## What it is",
                "",
                "better-writing rewrites prose for a named reader and genre. Trigger: a request to improve, humanise, de-AI, or copy-edit a draft. Steps: read the brief, set the reading, audit the draft against its reference lists, rewrite, then run a preflight checklist. Quality bar: no invented facts, figures kept, inflated claims cut or restated at their true strength.",
                "",
                "## How to use it",
                "",
                "Install: `npx skills add forjd/better-writing`, which puts it in `.agents/skills/better-writing/` and links it at `.cursor/skills/better-writing`.",
                "",
                "Invoke: type `/better-writing` in Cursor chat with the file, or ask the agent to apply better-writing to it.",
                "",
                "Input in this trial: `outputs/influence_before.md`, a rough findings draft. Output: `outputs/influence_after.md`. Each step is in `outputs/better_writing_run.md`.",
                "",
                "## Work application",
                "",
                "As AI Application Trainee in the Department of Physics, I write technology-research notes from public pages for colleagues. After the facts are checked, better-writing is the last pass that removes filler and claims the sources do not support.",
                "",
            ]
        ),
        encoding="utf-8",
    )


def write_all(d: Data, before: str, after: str) -> None:
    write_report(d)
    write_skill_note(d)
    write_a3(d, before, after)
