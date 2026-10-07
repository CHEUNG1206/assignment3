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

Life and health shows up as Kelun-Biotech at HK$1.56 billion in 2023, then as 16 Biotech Companies in 2025. Those 16 listings and five Specialist Technology Companies raised HK$19.8 billion together, against the Listing Committee IPO total of HK$285.8 billion. HengRui Pharma raised HK$11.4 billion and is kept separate. AI and robotics shows up as Horizon Robotics at HK$6.09 billion and RoboSense at HK$1.06 billion in 2024. New energy shows up as REPT BATTERO Energy at HK$2.12 billion and Sino-Synergy Hydrogen Energy at HK$1.56 billion in 2023. Semi-conductors, advanced materials, film, aerospace and construction have no industry total in the outcome window.

Chapter 18C took effect on 31 March 2023. There was no listing under it until XtalPi on 13 June 2024. Three listings in 2024 raised HK$2,903 million. Five in 2025 raised HK$6,014 million. Six from 1 January to 31 March 2026 raised HK$19,518 million. HKEX states 14 companies and HK$28.4 billion by end-March 2026. Over the same years the HKMA Base Rate went from 4.75% at end-2022 to 5.75% at end-2023, then to 4.75% and 4.00%. Southbound ADT went from HK$31.7 billion to HK$31.1 billion, HK$48.2 billion, and HK$121.1 billion. The Industry-Oriented Fund has no disbursement date in the 2026 address, so it is not counted in these listings.
