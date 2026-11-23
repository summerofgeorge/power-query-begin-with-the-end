# Begin with the end: Power Query and Copilot

A fictional branch-sales exercise from Stringfest Analytics. Use a picture of the desired result to help Copilot build a cleanup you can refresh next week.

## What's included

- `begin-with-the-end.xlsx`: the Raw_East and Raw_West source tables and a manually prepared example on Target_Example. The starter has no query.
- `desired-output.png`: the first eight target records, ready to attach to Copilot.
- `prompts.txt`: a short starting prompt, a build follow-up, and an optional handoff request.
- `ReportedSales.pq`: the Copilot-generated M query used in the desktop demonstration, if you want to practice with it.

## Try it

1. Open a copy of the workbook in Excel for Microsoft 365 on Windows. Attach the PNG to your Copilot conversation and make the workbook available as the source. If your Excel pane cannot accept images, attach both files in Copilot Chat.
2. Send the starting prompt, read the plan, and clarify anything it misunderstood. Then send the build follow-up. Whether Copilot can save the query directly depends on the experience available in your account.
3. Check that ReportedSales appears under Data > Queries & Connections. The screenshot shows the desired shape; the raw tables supply all the output values.

The pictured layout fills branch names down, removes subtotals, separates product codes from names, turns date columns into rows, and stacks East and West together. A blank sale means not reported; zero is a real sale. Keep product codes such as 001 as text.

## If Copilot gives you M code

In Excel for Windows, choose Data > Get Data > From Other Sources > Blank Query. Open Home > Advanced Editor, replace the existing code with the complete M query, and select Done. Name the query ReportedSales. Choose Close & Load > Close & Load To..., select Table and New worksheet, leave Add this data to the Data Model unchecked, and click OK. If the query already loaded elsewhere, right-click it in Queries & Connections and choose Load To....

You can use the contents of ReportedSales.pq instead. In the original test, Copilot supplied this code because its Get Data skill was not enabled in the account; the query was installed and refreshed in desktop Excel.

## Two quick checks

1. Compare a few familiar numbers: Cleveland / 001 / November 2 should be 120; its November 4 zero should remain; SKU 001 should keep its leading zeros. Select the numeric Sales cells and check Excel's status-bar Sum: the full sample totals $8,005, matching the prepared example. The supplied query produces 57 rows.
2. Change Raw_East!C2 from 120 to 145 and choose Data > Refresh All. That output sale should become 145 and the total $8,030. Restore 120 and refresh again afterward.

## Next week

Replace the data inside the two raw tables, make sure the tables include the new rows and date columns, clear out leftover old rows, and refresh. Keep the source layout consistent. This sample expects date headings like 2026-11-02; when adding one manually, format the header cell as Text before entering it so Excel does not rewrite it as a regional date. If the export's headings or layout change, review the query.

Save the workbook with its query. Leave the next person a short note about the two tables to update and Data > Refresh All, and have them try a refresh. The prepared target is a reference for the original sample, not a source for next week's query.
