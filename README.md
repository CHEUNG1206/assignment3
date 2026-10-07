# Policy Address funding and the following year's market

Coursework for CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics.

The card lists industry and technology programmes with a stated amount in the Policy Addresses from 2022 to 2026, then sets each address against HKEX IPO funds and average daily turnover in the next calendar year.

## Where the work is

- Assignment 1: `A1_CHEUNG_Yuk_Yuen/`
- Assignment 2: `A2_CHEUNG_Yuk_Yuen/`. The skill is `policy-funding-card`.
- Assignment 3: `A3_CHEUNG_Yuk_Yuen/`. The external skill is better-writing.
- Google Sheet: https://docs.google.com/spreadsheets/d/1Nr4XOrgmN3wJ44HHLVwDYQaRkt7oi6D_qz0EIGy9lJ0/edit

`HKEX_original` holds the HKEX figures. `Year_after` holds the board comparison. `Industry` holds each funded industry against the HKEX line that names it.

## Run it again

```bash
python3 -m pip install -r requirements.txt
python3 src/build_briefing.py
python3 .agents/skills/xlsx/scripts/recalc.py A1_CHEUNG_Yuk_Yuen/outputs/term_market.xlsx
```

Recalculation needs LibreOffice (`soffice`). The Policy Address PDFs have no text layer in this environment. The text used here is the Chief Executive's Office XML.
