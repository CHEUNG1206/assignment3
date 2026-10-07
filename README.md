# 施政報告與港交所對照

這份作業整理李家超就任行政長官（2022 年 7 月 1 日）之後，各份施政報告裡寫明金額的產業與創科計劃，並對上港交所公開統計。

## 結果在哪

- 作業 1：`A1_CHEUNG_Yuk_Yuen/`
- 作業 2：`A2_CHEUNG_Yuk_Yuen/`，Skill 是 `policy-funding-card`
- 作業 3：`A3_CHEUNG_Yuk_Yuen/`，使用 better-writing
- Google 試算表：https://docs.google.com/spreadsheets/d/1Nr4XOrgmN3wJ44HHLVwDYQaRkt7oi6D_qz0EIGy9lJ0/edit

`HKEX_original` 是港交所原文數字。`Policy_funding` 是施政報告 XML 裡的撥款句。

## 本地重跑

```bash
python3 -m pip install -r requirements.txt
python3 src/build_briefing.py
python3 .agents/skills/xlsx/scripts/recalc.py A1_CHEUNG_Yuk_Yuen/outputs/term_market.xlsx
```

重算需要 LibreOffice（`soffice`）。施政報告 PDF 在這台環境沒有文字層，原文用行政長官辦公室的 XML。
