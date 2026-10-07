# Policy Address funding and the industries it named

Coursework for CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics.

The card lists industry and technology programmes with a stated amount in the Policy Addresses from 2022 to 2026, marks whether each amount is new, and follows every Main Board IPO in those industries from each delivery date, using the HKEX New Listing Reports.

## Where the work is

- Assignment 1: `A1_CHEUNG_Yuk_Yuen/`. Charts are in `outputs/charts/`.
- Assignment 2: `A2_CHEUNG_Yuk_Yuen/`. The skill is `policy-funding-card`.
- Assignment 3: `A3_CHEUNG_Yuk_Yuen/`. The external skill is better-writing; the run is in `outputs/better_writing_run.md`.
- HKEX New Listing Reports, 2021 to 2026: `data/hkex_nlr/`.
- Google Sheet: https://docs.google.com/spreadsheets/d/1Nr4XOrgmN3wJ44HHLVwDYQaRkt7oi6D_qz0EIGy9lJ0/edit

`HKEX_original` holds the HKEX figures. `Policy_funding` holds the funding rows with their status. `Year_after` holds the board comparison. `Industry` holds each policy industry by window. `C18_before_after` starts on 31 March 2023, when Chapter 18C took effect.

## Run it again

```bash
python3 -m pip install -r requirements.txt
python3 src/build_briefing.py
python3 .agents/skills/xlsx/scripts/recalc.py A1_CHEUNG_Yuk_Yuen/outputs/term_market.xlsx
```

Recalculation needs LibreOffice (`soffice`). The Policy Address PDFs have no text layer in this environment. The text used here is the Chief Executive's Office XML.
