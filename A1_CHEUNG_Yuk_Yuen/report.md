# Assignment 1 report

CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics. 7 October 2026.

## The task

I am CHEUNG Yuk Yuen. The reader is someone who wants to check public documents. The task covers the Policy Addresses after John Lee took office on 1 July 2022, through 16 September 2026. For each address I record which industries the Government funds with a stated amount, whether that money is new in the address, and what happened to those industries' IPOs on HKEX after delivery. A missing fact stays `not stated in source`. Two official figures that differ are both kept.

## How I did it

The first version compared each address with the HKEX total for the year of the speech. That total includes the months before delivery, so I moved the outcome to the following calendar year. A review of that version found more problems. Some funding sentences describe money announced before the address. The industry evidence came only from HKEX's ten-largest IPO list. A ratio divided government funding by later IPO funds, which are unrelated sums. This version fixes each of them.

I read AGENTS.md and memory/PROFILE.md. The name on the work is CHEUNG Yuk Yuen.

The Policy Address PDFs have no text layer in this environment. I used the Chief Executive's Office English XML, pa2022.xml through pa2026.xml, and copied the verb of every funding sentence into a status column.

For industries I read every Main Board IPO in the HKEX New Listing Reports for 2021 to 29 September 2026. Biotech uses HKEX's B marker. Other industries use a keyword screen on the company name, published with its exclusions, and treated as a lower bound. Each address is scored from its delivery date to the day before the next one, against the 12 months before the first address.

Chapter 18C is measured from its effective date, 31 March 2023, beside the HKMA Base Rate and Southbound Stock Connect turnover. The HK$10 billion I&T Industry-Oriented Fund is left out because it has not launched.

better-writing rewrote the findings paragraph only. The run is recorded in `A3_CHEUNG_Yuk_Yuen/outputs/better_writing_run.md`.

The workbook stores the printed amounts and every listing row in blue cells. Conversions, window totals, and shares are formulas, recalculated in LibreOffice.

## Course concepts

| Concept | How I used it | What changed |
| --- | --- | --- |
| Plan Mode | The outcome window was fixed before the cards were rebuilt: delivery date to the day before the next address. | No month before a speech is counted as its outcome. |
| AGENTS.md | Missing facts stay not stated in source. Differing official figures are both kept. Verbs decide whether money is new. | 286.9, 286.8, 285.8 and the NLR sum are listed, not averaged. 2025 has no new amount. |
| Skills | policy-funding-card extracts one address. better-writing edits the findings paragraph only. | The 2022 rerun did not copy a 2026 amount. |
| Memory | Only the identity stated in the earlier assignments. | No university, degree, or Chinese personal name was added. |
| MCP | HKEX figures and the funding rows sit in a Google Sheet. | Each comparison has its own tab. |

## Result

Card: `A1_CHEUNG_Yuk_Yuen/outputs/term_funding_card.md`. Charts: `A1_CHEUNG_Yuk_Yuen/outputs/charts/`.

Sheet: https://docs.google.com/spreadsheets/d/1Nr4XOrgmN3wJ44HHLVwDYQaRkt7oi6D_qz0EIGy9lJ0/edit

Workbook: `A1_CHEUNG_Yuk_Yuen/outputs/term_market.xlsx`.

| Address | New in this address, HK$ million | Already in place, HK$ million |
| --- | --- | --- |
| 2022 | 40,000.0 | 5,000.0 |
| 2023 | 14,300.0 | 0.0 |
| 2024 | 13,930.0 | 0.0 |
| 2025 | 0.0 | 4,100.0 |
| 2026 | 1,300.0 | 2,000.0 |

Semi-conductors had no Main Board IPO by the name screen in the four windows from 19 October 2021 to 16 September 2025. In the window after the 2025 address, 12 listed, starting on 5 December 2025, and raised HK$39,983.4 million, 7.9% of the Main Board. The 2024 address named semi-conductors beside the HK$10 billion I&T Industry-Oriented Fund. The 2026 address says that fund is expected to launch in 2026, so the money had not been paid out when these companies listed.

AI and robotics went from no IPO before the 2022 address to HK$14,791.6 million after the 2025 address. Its share of the Main Board stayed between 2.8% and 3.7% in the three windows after the 2023 address. Chapter 18A biotech raised HK$18,604.4 million after the 2025 address, the most in any window, while its share fell from 9.4% before the 2022 address to 3.7%.

Over the same windows the Main Board rose from HK$127,696.1 million to HK$508,850.3 million. HKEX attributes the 2025 rebound to global demand for diversified growth opportunities, frontier technologies in the region, and local market reforms. It does not mention the Policy Address.

The HK$4.3 billion for film and CreateSmart was followed by only one film or creative IPO, HK$444.4 million, in the three windows after the 2023 address. No aerospace IPO followed the 2025 address, which cited an allocation of over HK$100 million. The HK$1 billion construction injection has 14 days of data after it and no construction IPO.

The name screen misses Contemporary Amperex Technology. Adding its HK$41,005.7 million would make new energy 25.0% of the window after the 2024 address.

![Quarterly IPO funds by policy industry](outputs/charts/industry_quarterly.png)
