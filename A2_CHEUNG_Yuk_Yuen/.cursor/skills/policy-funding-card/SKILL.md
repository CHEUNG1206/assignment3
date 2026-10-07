---
name: policy-funding-card
description: >-
  Turns one Hong Kong Policy Address into a funding card and scores it
  against the HKEX market in the following calendar year. Use when the user
  gives a Policy Address and wants the industries the Government funds and
  what the market did after that address. Trigger phrases include
  "施政報告", "policy address", "following year", and "policy-funding-card".
---

# Policy funding card

Read one Policy Address. Record the industries and technology programmes that name an amount of government money. Then record what HKEX published for the next calendar year, which is the market after the address.

## Trigger

Use this skill when the user supplies one Policy Address and asks how that address lines up with the market afterwards. One address per run. A second address is a new run, not a copy of the first.

## Steps

1. Read `AGENTS.md` and `memory/PROFILE.md` before writing about the person. CHEUNG Yuk Yuen is the AI Application Trainee. The agent is not.
2. Read the official English XML when it exists: `https://www.ceo.gov.hk/public/open-data/en/policy_address/paYEAR.xml`. Use the Traditional Chinese XML only to check a sentence the English file does not contain. The printed PDF on policyaddress.gov.hk has no text layer in this environment.
3. Record the delivery date. An address delivered in September or October is scored on the next calendar year, not on the year of the speech.
4. Market wording: copy what that address says about the stock market, IPO funds, or turnover. Label those sentences as the market before or up to delivery. They are not the outcome.
5. Funding section: keep a row only when the same sentence states an amount and commits, earmarks, injects, redeploys, or sets up a named fund. Record the programme, the amount as printed, the industries named in that sentence, and the paragraph number when the paragraph has one.
6. Put sentences that name an industry but no amount under "Direction stated, amount not stated".
7. Do not treat a private investor's money, a Construction Industry Council sum, or a target for future private investment as government funding.
8. Outcome section: copy the HKEX IPO funds and average daily turnover for the next calendar year from the Google Sheet `HKEX original market figures` or from the HKEX PDF cited there. Copy the value, unit, as-of date, printed percentage change, and URL. If two official rows differ, keep both and do not average them. These board totals are context.
9. Industry section: one row per industry named in a funding sentence. Match it to an HKEX line for the outcome year that names that industry, a Chapter 18A or 18C count, or a company on HKEX's ten-largest IPO list whose printed name states the industry. Record the amount, unit, and URL. If the page gives no figure for that industry, write `not stated in source`. Do not use a board total as the industry result. Do not add a Chapter 18A figure to a Chapter 18C figure unless the source already prints the combined sum.
10. Set one outcome label:
   - `following year recorded, source does not attribute it to the address`
   - `following year only partly published`
   - `following year not yet published`
   - `source attributes the market result to this policy`
11. Compare scale with a formula: the industry amount in HK$ million, divided by IPO funds in the same outcome window. Label the result as a size comparison. Do not call it a causal share.
12. Hand only the influence paragraph to better-writing. Leave amounts, dates, labels, and URLs unchanged.

## Quality bar

- Every amount and date has a source URL.
- Missing information stays `not stated in source`.
- The same year's HKEX total may be quoted as the baseline. It is not the outcome.
- A partly published following year is still reported, with the months named.
- Do not forecast a year HKEX has not published, and do not recommend buying or selling.
- Do not add Main Board and GEM totals to a Mainland-enterprise subset or to one row of the funds-by-type table.
- A rerun on another year must not reuse the previous year's amounts.
- A board IPO total does not fill a blank industry row.

## Do not

- Do not register, email, or trade.
- Do not fill a blank university, degree, or contact field.
- Do not replace an HKEX number with a newspaper number.
