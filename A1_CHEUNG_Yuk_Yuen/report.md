# Assignment 1 report

CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics. 7 October 2026.

## The task

I am CHEUNG Yuk Yuen. The reader is someone who wants to check public documents. The task covers the Policy Addresses after John Lee took office on 1 July 2022, through 16 September 2026. For each address I record which industries the Government funds with a stated amount, and what HKEX recorded in the following calendar year. A missing fact stays `not stated in source`. Two official figures that differ are both kept. The funding-to-IPO share is a size comparison, not a causal claim.

## How I did it

On 7 October I set the outcome window before opening the files. An address delivered in September or October is scored on the next calendar year. The year of the speech is the baseline.

I read AGENTS.md and memory/PROFILE.md. The name on the work is CHEUNG Yuk Yuen.

The Policy Address PDFs have no text layer in this environment. I used the Chief Executive's Office English XML, pa2022.xml through pa2026.xml.

HKEX figures were read from the annual market statistics and the 2025 final results, then copied into the Google Sheet. Market Statistics 2024 puts the year after the 2023 address on the same HK$ million IPO line as 2022 and 2023: HK$87,477.8 million.

policy-funding-card extracts one address at a time. The second run uses only 2022 and pairs it with the 2023 market. It does not copy the extra HK$1 billion named in 2026.

better-writing rewrote the influence paragraph only.

The workbook stores printed amounts in blue cells. HK$ million, the funding sum, and the share of that sum in the next IPO total are formulas, recalculated in LibreOffice.

## Course concepts

| Concept | How I used it | What changed |
| --- | --- | --- |
| Plan Mode | The outcome window was fixed before the card was written. | The same year's HKEX total stayed a baseline. |
| AGENTS.md | Missing facts stay not stated in source. Differing official figures are both kept. | 286.9 and 286.8 were not averaged. 2027 has no invented total. |
| Skills | policy-funding-card extracts one year. better-writing edits the paragraph only. | The 2022 rerun did not copy a 2026 amount. |
| Memory | Only the identity stated in the earlier assignments. | No university, degree, or Chinese personal name was added. |
| MCP | HKEX figures sit in a Google Sheet. | The following-year comparison has its own tab. |

## Result

Card: `A1_CHEUNG_Yuk_Yuen/outputs/term_funding_card.md`.

Sheet: https://docs.google.com/spreadsheets/d/1Nr4XOrgmN3wJ44HHLVwDYQaRkt7oi6D_qz0EIGy9lJ0/edit

Workbook: `A1_CHEUNG_Yuk_Yuen/outputs/term_market.xlsx`.

After the 2022 address, 2023 IPO funds fell 55.8% to HK$46,294.9 million. After the 2023 address, 2024 IPO funds rose 88.9% to HK$87,477.8 million. After the 2024 address, 2025 IPO funds rose 226% to HK$286.9 billion, and new economy sectors were 66% of capital raised. After the 2025 address, the first eight months of 2026 raised HK$342.4 billion, up 153%. The 2026 address has no published 2027 figure.

Self-check: the amounts match the cited pages. The industry overlap is in the 2024 largest-IPO list and in the 2025 new-economy share. HKEX does not attribute those totals to these funds. The Industry-Oriented Fund was still expected to launch in 2026.
