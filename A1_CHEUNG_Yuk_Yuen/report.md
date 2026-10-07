# Assignment 1 report

CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics. 7 October 2026.

## 1. The task

I asked the agent to find which industries John Lee's Policy Addresses from 2022 to 2026 fund with a stated amount, and what happened to those industries' IPOs on HKEX after each address. The reader is someone who checks public documents. Sources: the Policy Address XML, the HKEX New Listing Reports, and HKEX market statistics.

## 2. How I did it

1. Wrote the task brief and AGENTS.md rules: copy figures from the source, write not stated in source, keep conflicting official figures.
2. The agent recalled memory/PROFILE.md, so the work names me only as CHEUNG Yuk Yuen, AI Application Trainee.
3. The agent read the English Policy Address XML for 2022 to 2026 and listed every funding sentence that states an amount.
4. The agent copied HKEX figures into a Google Sheet through the Google Sheets MCP.
5. The first result compared each address with the same year's market. I pointed out that this counts months before the speech, so the outcome moved to the period after delivery.
6. I asked for each funded industry, not the board total. The agent switched to every Main Board IPO in the HKEX New Listing Reports, counted from each delivery date.
7. A review found money announced before an address counted as new, and a ratio of funding to IPO funds that compared unrelated sums. The agent added a status column from each sentence's verb and removed the ratio.
8. The agent built the workbook with formulas, the charts and the PDFs, and packaged the workflow as the policy-funding-card skill for Assignment 2.

## 3. Course concepts

| Concept | How I used it | What it changed |
| --- | --- | --- |
| AGENTS.md | Rules on sources, missing facts, conflicting figures, and which verbs make money new. | HKEX's 285.8, 286.8 and 286.9 billion for 2025 are all kept. The 2025 address has no new amount. |
| Skills | policy-funding-card runs one address at a time. better-writing edits the findings only. | The 2022 rerun did not pick up a 2026 amount. |
| Memory | memory/PROFILE.md holds the identity from my earlier assignments. | No university, degree, or Chinese name was added. |
| MCP | The Google Sheets MCP holds the HKEX figures and the funding rows. | Each comparison has its own tab that a reader can open. |

## 4. Result

Files: `outputs/term_funding_card.md`, `outputs/analysis.pdf` (all tables and charts), `outputs/term_market.xlsx`. Google Sheet: https://docs.google.com/spreadsheets/d/1Nr4XOrgmN3wJ44HHLVwDYQaRkt7oi6D_qz0EIGy9lJ0/edit

![Share of Main Board IPO funds by window](outputs/charts/industry_window_shares.png)

Main finding: semi-conductors went from no IPO to 12 IPOs and HK$39,983.4 million after the 2025 address, but the HK$10 billion fund that named them has not launched. Over the same windows the whole Main Board grew about fourfold, and HKEX credits global demand, frontier technology and listing reforms. The figures show timing, not that funding moved an industry.

Self-assessment: every figure traces to a source file and the method is written down; industry totals other than biotech come from a name screen, so they are lower bounds.
