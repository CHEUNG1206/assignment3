# 作業 1 報告

CHEUNG Yuk Yuen，2026-10-07。

## 任務

我是物理系 AI Application Trainee。讀者是要核對公開文件的人。任務是從李家超 2022 年 7 月 1 日就任起，到 2026 年 10 月 7 日，整理每份施政報告怎樣寫市場，以及哪些產業計劃在同一句裡寫了金額。成功標準是每個金額都有出處，缺的資料標成 not stated in source，而且不把集資額說成施政報告造成的結果。

## 做法

10 月 7 日先定欄位和四個影響標籤，再開檔。

讀 AGENTS.md 和 memory/PROFILE.md。姓名只用 CHEUNG Yuk Yuen。

施政報告 PDF 沒有文字層。改讀行政長官辦公室公開 XML，pa2022.xml 至 pa2026.xml。

港交所 PDF 用 pdfplumber 抽出原文，先寫進 Google 試算表。

policy-funding-card 逐年抽資金句。2022 年單獨再跑一次，確認沒有抄入 2026 年的額外 10 億元。

better-writing 只改寫影響說明。

xlsx 把原文金額放在藍色輸入格，億元換算用公式，並經 LibreOffice 重算。

## 概念

| 概念 | 怎樣用 | 改變了什麼 |
| --- | --- | --- |
| Plan Mode | 先定欄位、標籤和來源，再寫卡片 | 沒有先寫一篇市場評論 |
| AGENTS.md | 缺資料寫 not stated in source，兩份官方數字都保留 | 2025 年的 286.9 和 286.8 沒有被平均 |
| Skills | 自建 policy-funding-card，再用 better-writing | 資金表和說明分成兩步 |
| Memory | 只記住作業裡寫過的身份 | 沒有補上院校、學位或中文姓名 |
| MCP | 港交所原文數字放在 Google 試算表 | 市場數字和施政報告金額分開存放 |

## 結果

卡片：`A1_CHEUNG_Yuk_Yuen/outputs/term_funding_card.md`。

試算表：https://docs.google.com/spreadsheets/d/1Nr4XOrgmN3wJ44HHLVwDYQaRkt7oi6D_qz0EIGy9lJ0/edit

本地核對簿：`A1_CHEUNG_Yuk_Yuen/outputs/term_market.xlsx`。

自我評估：資金句和港交所數字對得上來源；影響判斷停在時間先後，沒有做成因估計。
