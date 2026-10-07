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
5. Funding section: keep a row only when the same sentence states an amount and commits, earmarks, injects, redeploys, or sets up a named fund. Record the programme, the amount as printed, the industries named in that sentence, and the paragraph number when the paragraph has one. Copy the verb into a status column. Future verbs ("will", "set up", "set aside") are `new in this address`. Past or present-perfect verbs ("earlier announced", "has allocated", "has been earmarked", "injected") are `already in place before this address`. Only the first status counts towards that address's new total.
6. Put sentences that name an industry but no amount under "Direction stated, amount not stated".
7. Do not treat a private investor's money, a Construction Industry Council sum, or a target for future private investment as government funding.
8. Outcome section: copy the HKEX IPO funds and average daily turnover for the next calendar year from the Google Sheet `HKEX original market figures` or from the HKEX PDF cited there. Copy the value, unit, as-of date, printed percentage change, and URL. If two official rows differ, keep both and do not average them. These board totals are context.
9. Industry section: read every Main Board IPO from the HKEX New Listing Report for each year (`https://www2.hkexnews.hk/-/media/HKEXnews/Homepage/New-Listings/New-Listing-Information/New-Listing-Report/Main/NLRYEAR_Eng.xlsx`). Funds raised is the Hong Kong offer row plus the international offer row. The window runs from the delivery date to the day before the next address, and the address before the first one gives a 12-month baseline. Tag Chapter 18A biotech by HKEX's B marker. Tag other policy industries with a keyword screen on the company name, publish the keywords and the exclusions, and call the result a lower bound. List large companies the screen misses beside the total, with their funds, and do not add them in. Show each industry's count, HK$ million, and share of the same file's Main Board total for the same window. If no IPO matches, write 0 and keep the row. Do not add a Chapter 18A figure to a Chapter 18C figure unless the source already prints the combined sum.
10. Set one outcome label:
   - `following year recorded, source does not attribute it to the address`
   - `following year only partly published`
   - `following year not yet published`
   - `source attributes the market result to this policy`
11. Timing: the New Listing Report prints the prospectus date and the listing date, not the application date, so the preparation lag cannot be measured from it. Show each industry's total from every delivery date to the end of the latest report next to the window figures, and say when the industry first leaves zero. For a fund, check whether the latest address says it has launched. A fund that has not launched cannot have paid for an IPO. Do not divide a funding amount by IPO funds.
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
