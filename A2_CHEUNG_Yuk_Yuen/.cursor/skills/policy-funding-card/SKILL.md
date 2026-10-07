---
name: policy-funding-card
description: >-
  Turns one Hong Kong Policy Address into a funding-and-market card.
  Use when the user gives a Policy Address (official XML or full text) and
  wants the government's market wording, programmes with a stated amount,
  and the HKEX figure for that year. Trigger phrases include "施政報告",
  "policy address", "產業資金", and "policy-funding-card".
---

# Policy funding card

Read one Policy Address. Record what it says about the market, which industry or technology programmes name an amount of money, and how that lines up with an HKEX figure for the same year.

## Trigger

Use this skill when the user supplies one Policy Address and asks for a market-and-funding card. One year per run. A second year is a new run, not a copy of the first.

## Steps

1. Read `AGENTS.md` and `memory/PROFILE.md` before writing about the person.
2. Read the official Traditional Chinese XML when it exists: `https://www.ceo.gov.hk/public/open-data/tc/policy_address/paYEAR.xml`. The printed PDF on policyaddress.gov.hk has no text layer in this environment, so do not treat a blank extract as "the address says nothing".
3. Record the delivery date from the XML or from the Government Information Services page for that address.
4. Market section: copy sentences about the stock market, IPO fundraising, or turnover. If the sentence has no figure, write that no figure is stated. Do not add a Hang Seng level the text does not contain.
5. Funding section: keep a row only when the same sentence states an amount and uses 撥款, 預留, 注資, 調撥, or 設立/成立 a named fund. Record the programme, the amount as printed, the industries named in that sentence, and the paragraph number when the paragraph has one.
6. Put sentences that name an industry but no amount in a separate list titled "方向已寫、金額未寫".
7. Do not treat a private investor's money, a Jockey Club sum, or a target for future private investment as government funding.
8. Read the HKEX figure from the Google Sheet `HKEX original market figures` or from the HKEX PDF already cited there. Copy the value, unit, as-of date, and URL. If two HKEX rows differ, keep both.
9. Set one influence label:
   - `市場報道把結果歸因於這項政策`
   - `同一年同時出現，來源沒有寫因果`
   - `市場結果早於這份施政報告`
   - `找不到該年官方數字`
   A full-year HKEX total that includes months before the address uses `市場結果早於這份施政報告` for the part dated before delivery, and does not become evidence of an effect after delivery.
10. Hand only the influence paragraph to better-writing. Leave amounts, dates, labels, and URLs unchanged.

## Quality bar

- Every amount and date has a source URL.
- Missing information stays `not stated in source`.
- Do not forecast the market and do not recommend buying or selling.
- Do not sum amounts that use different units or different universes (Main Board and GEM versus a Mainland-enterprise subset versus one row of the funds-by-type table).
- A rerun on another year must not reuse the previous year's amounts.

## Do not

- Do not register, email, or trade.
- Do not fill a blank university, degree, or contact field.
- Do not replace an HKEX number with a newspaper number.
