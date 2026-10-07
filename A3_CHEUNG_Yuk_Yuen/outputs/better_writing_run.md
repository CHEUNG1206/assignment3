# better-writing run on the findings paragraph

Skill: `.agents/skills/better-writing/SKILL.md`, version 1.3.0. Input: `influence_before.md`. Output: `influence_after.md`.

The agent wrote the draft in the generic style the skill targets, then read SKILL.md and its references and followed the steps below.

## 1. Brief

- Genre: the findings paragraph of a coursework card that checks public documents.
- Audience: a reader who will open the XML and the HKEX files.
- Outcome: state what the IPO figures show after each address, at their real strength.
- Constraints from `AGENTS.md`: amounts, dates, labels, and URLs are frozen. No figure may be added.
- Mode: full rewrite, not light edit. The brief did not ask to keep a voice.

## 2. Reading

Reading this as: an analytical findings paragraph for a reader who checks sources, with a neutral tone, optimising for claims at their true strength.

Dials: directness high, warmth low, personality low, density high, evidence high, polish medium.

## 3. Audit

No near-conclusive artefacts. The tells cluster, so the audit proceeds.

| Phrase in the draft | Pattern | Reference |
| --- | --- | --- |
| "It is important to note that", "it should be noted", "It is also worth mentioning" | Throat-clearing | `structures-and-phrases.md`, Throat-clearing |
| "played a pivotal role in shaping", "remarkable growth", "key pillar" | Significance inflation | `ai-writing-patterns.md`, Significance inflation |
| "landscape" | Business jargon | `structures-and-phrases.md`, Business jargon |
| "a testament to the power of targeted policy" | Significance inflation; claims a cause the figures do not show | `ai-writing-patterns.md`, Significance inflation |
| "Meanwhile" | Mechanical transition | `structures-and-phrases.md`, Mechanical transitions |
| "Taken together, these findings underscore" | Essay-scaffold closer | `structures-and-phrases.md`, Essay-scaffold closers |
| "not merely words but a driving force" | Negative parallelism; inflation | `ai-writing-patterns.md`, Negative parallelism |
| "deeply interconnected", "transformation" | Vague connection | `ai-writing-patterns.md`, Vague connection |
| One paragraph, sentences up to 70 words | Uniform density, backtracking | `preflight.md`, Structure check |

Stock phrases found: 15 kinds. Left after the rewrite: 0.

## 4. Rewrite

Moves, in the skill's order: cut throat-clearing; restate inflation at its true strength; split the long sentences so each holds one claim; put the doer first; end on the strongest fact already in the draft. The two causal clauses were cut, not reworded, because the draft's own facts (the fund had not launched, the board grew fourfold, HKEX names other causes) contradict them.

## 5. Self-audit and preflight

- Preservation: every figure in the rewrite appears in the draft (30 of 30). Checked by `figures_in()` in `src/build_briefing.py`.
- Scope words kept: "only", "no", "most", and "between" survived.
- Last sentence: the new-energy sensitivity, a fact from the draft. No invented verdict.
- Anti-slop: no throat-clearing, no closer, no dash hinges, no vague attribution. HKEX is named.
- Structure: sentence lengths [22, 23, 12, 24, 18, 19, 27, 14, 21, 7, 26, 15, 16, 7, 16]. Five paragraphs, each about one group of industries.
- Taste: directness yes, specificity yes, restraint yes.

## Change note

Cut two causal claims, "a testament to the power of targeted policy" and "a driving force behind the market's transformation". The draft's own figures show timing only. Split one paragraph into five.

## Counts

| Measure | Before | After |
| --- | --- | --- |
| Words | 369 | 267 |
| Sentences | 9 | 15 |
| Longest sentence, words | 71 | 27 |
| Stock phrases | 15 | 0 |
| Figures | 41 | 39 |

![Before and after counts](better_writing_counts.png)
