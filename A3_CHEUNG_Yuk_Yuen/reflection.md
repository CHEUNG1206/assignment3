# Assignment 3 reflection

CHEUNG Yuk Yuen, AI Application Trainee, Department of Physics.

## Discovery

The handout lists better-writing at https://skills.sh/forjd/better-writing/better-writing. I chose it because my findings paragraph had to read plainly while every figure stayed fixed.

## First use

I asked the agent to write a rough findings draft in the generic style the skill targets, then to apply better-writing to it. The agent read SKILL.md and its reference files and followed the five steps: brief, reading, audit, rewrite, preflight. The record is outputs/better_writing_run.md. It cut throat-clearing and two causal claims the figures do not support: "a testament to the power of targeted policy" and "a driving force behind the market's transformation".

| Measure | Before | After |
| --- | --- | --- |
| Words | 369 | 267 |
| Sentences | 9 | 15 |
| Longest sentence, words | 71 | 27 |
| Stock phrases | 15 | 0 |
| Figures | 41 | 39 |

![Before and after counts](outputs/better_writing_counts.png)

## Friction

The skill edits prose and does not check facts. Seeing that a causal claim was unsupported needed the funding status and the fund's launch date from my own card.

## vs Assignment 2

policy-funding-card decides which facts enter the table. better-writing decides how they read. The rule worth borrowing is its preflight check that nothing in the rewrite is new.

## Work fit

I would use it on each technology-research note I hand to a colleague, as the last step after the source check. It replaces my manual pass that deletes filler and inflated claims.

## One change

I would add a preflight line that a script confirms every figure in the rewrite appears in the draft, with tables, amounts, dates and labels frozen.
